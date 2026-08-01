#!/usr/bin/env python3
import copy
import html
import logging
import pathlib
import sys
import zipfile
import lxml.etree as ET

# Настройка логов, чтобы в консоль лишнее не сыпалось
logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")

# Пространства имен для ворда, без них xml не пропарсить
word_ns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}


class DocxReader:
    # Класс для чтения начинки docx (это же просто zip архив)

    def __init__(self, docx_path):
        self.docx_path = pathlib.Path(docx_path)
        if self.docx_path.exists() == False:
            raise FileNotFoundError(f"Файл '{docx_path}' не найден!")
        try:
            self.archive = zipfile.ZipFile(self.docx_path, "r")
        except zipfile.BadZipFile as e:
            raise ValueError(f"Файл '{docx_path}' сломан или это не docx!") from e

    def read_file(self, put_vnutri):
        # Читаем файлик из архива в байтах
        if put_vnutri in self.archive.namelist():
            return self.archive.read(put_vnutri)
        return None

    def extract_media(self, put_vnutri, put_snaruzhi):
        # Вытаскиваем картинки наружу в папку
        paths_to_check = [put_vnutri, f"word/{put_vnutri}"]
        for target_path in paths_to_check:
            if target_path in self.archive.namelist():
                out_file = pathlib.Path(put_snaruzhi)
                out_file.parent.mkdir(parents=True, exist_ok=True)
                with open(out_file, "wb") as f:
                    f.write(self.archive.read(target_path))
                return
        logging.warning("Картинка '%s' не найдена внутри архива.", put_vnutri)

    def close(self):
        self.archive.close()


class StyleResolver:
    # Тут переводим вордовские стили в нормальные теги заголовков h1-h6
    def __init__(self, styles_xml):
        self.styles = {}
        if not styles_xml:
            return
        try:
            root = ET.fromstring(styles_xml)
            for style in root.findall(".//w:style", word_ns):
                style_id = style.get(f'{{{word_ns["w"]}}}styleId')
                name_elem = style.find(".//w:name", word_ns)
                if style_id and name_elem is not None:
                    name_val = name_elem.get(f'{{{word_ns["w"]}}}val', "").lower()
                    if "heading 1" in name_val:
                        self.styles[style_id] = "h1"
                    elif "heading 2" in name_val:
                        self.styles[style_id] = "h2"
                    elif "heading 3" in name_val:
                        self.styles[style_id] = "h3"
                    elif "heading 4" in name_val:
                        self.styles[style_id] = "h4"
                    elif "heading 5" in name_val:
                        self.styles[style_id] = "h5"
                    elif "heading 6" in name_val:
                        self.styles[style_id] = "h6"
        except Exception as e:
            logging.warning("Не удалось распарсить стили: %s", e)

    def get_heading_level(self, style_id):
        return self.styles.get(style_id)


class NumberingResolver:
    # Эта штука нужна, чтобы разбираться со списками (маркированный или нумерованный)
    def __init__(self, numbering_xml):
        self.nomera = {}
        self.abstraktnie_nomera = {}
        if not numbering_xml:
            return
        try:
            root = ET.fromstring(numbering_xml)
            for num in root.findall(".//w:num", word_ns):
                num_id = num.get(f'{{{word_ns["w"]}}}numId')
                abs_elem = num.find(".//w:abstractNumId", word_ns)
                if num_id and abs_elem is not None:
                    self.nomera[num_id] = abs_elem.get(f'{{{word_ns["w"]}}}val', "")

            for abs_num in root.findall(".//w:abstractNum", word_ns):
                abs_id = abs_num.get(f'{{{word_ns["w"]}}}abstractNumId')
                levels = {}
                for lvl in abs_num.findall(".//w:lvl", word_ns):
                    ilvl = lvl.get(f'{{{word_ns["w"]}}}ilvl')
                    fmt_elem = lvl.find(".//w:numFmt", word_ns)
                    if ilvl and fmt_elem is not None:
                        levels[ilvl] = fmt_elem.get(f'{{{word_ns["w"]}}}val', "bullet")
                if abs_id:
                    self.abstraktnie_nomera[abs_id] = levels
        except Exception as e:
            logging.warning("Не удалось распарсить нумерацию списков: %s", e)

    def get_list_info(self, p_node):
        num_pr = p_node.find(".//w:pPr/w:numPr", word_ns)
        if num_pr is not None:
            num_id_elem = num_pr.find(".//w:numId", word_ns)
            ilvl_elem = num_pr.find(".//w:ilvl", word_ns)
            if num_id_elem is not None and ilvl_elem is not None:
                num_id = num_id_elem.get(f'{{{word_ns["w"]}}}val')
                try:
                    ilvl = int(ilvl_elem.get(f'{{{word_ns["w"]}}}val', "0"))
                    return num_id, ilvl
                except ValueError:
                    return num_id, 0
        return None, None

    def get_list_type(self, num_id, ilvl):
        abs_id = self.nomera.get(num_id)
        if abs_id and abs_id in self.abstraktnie_nomera:
            fmt = self.abstraktnie_nomera[abs_id].get(str(ilvl), "bullet")
            if fmt == "bullet":
                return "ul"
        return "ol"


class RelationshipResolver:
    # Ищем ссылки и связи для картинок
    def __init__(self, rels_xml):
        self.svyazi = {}
        if not rels_xml:
            return
        try:
            root = ET.fromstring(rels_xml)
            for rel in root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                r_id = rel.get("Id")
                target = rel.get("Target")
                if r_id and target:
                    self.svyazi[r_id] = target
        except Exception as e:
            logging.warning("Ошибка при поиске связей: %s", e)

    def get_target(self, r_id):
        return self.svyazi.get(r_id)


class MathConverter:
    # Конвертер формул через XSLT, тут магия, лучше не трогать лишний раз
    def __init__(self, xsl_path):
        xsl_file = pathlib.Path(xsl_path)
        if xsl_file.exists() == False:
            raise FileNotFoundError(f"Нет XSLT файла по пути: {xsl_path}")
        try:
            self.transformer = ET.XSLT(ET.parse(str(xsl_file)))
        except Exception as e:
            raise RuntimeError(f"Не удалось запустить XSLT процессор: {e}") from e

    def convert(self, omath_node, is_block=False):
        try:
            node_copy = copy.deepcopy(omath_node)
            dummy_root = ET.Element("root")
            dummy_root.append(node_copy)

            result_tree = self.transformer(dummy_root)
            raw_xml = ET.tostring(result_tree, encoding="utf-8").decode("utf-8")

            cleaned_mathml = raw_xml.replace("mml:", "").replace("xmlns:mml", "xmlns")

            display_mode = 'display="block"' if is_block else 'display="inline"'
            cleaned_mathml = cleaned_mathml.replace("<math ", f"<math {display_mode} ", 1)
            cleaned_mathml = cleaned_mathml.replace("<math>", f"<math {display_mode}>", 1)

            return cleaned_mathml
        except Exception as e:
            logging.warning("Сломалась формула, пропускаем: %s", e)
            return '<span class="math-error">[Ошибка обработки формулы]</span>'


class HtmlWriter:
    # Просто записываю готовый HTML файл по красивому шаблону со стилями
    @staticmethod
    def write(content, output_path):
        template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Converted Document</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            line-height: 1.7;
            color: #333;
            max-width: 820px;
            margin: 0 auto;
            padding: 2.5rem 1.5rem;
        }}
        p {{
            margin: 0 0 1.2em;
            text-align: justify;
        }}
        h1, h2, h3, h4, h5, h6 {{
            color: #111;
            font-weight: 600;
            margin-top: 1.8em;
            margin-bottom: 0.6em;
            line-height: 1.25;
        }}
        h1 {{ font-size: 2em; border-bottom: 1px solid #eee; padding-bottom: 0.3em; }}
        h2 {{ font-size: 1.5em; }}
        h3 {{ font-size: 1.25em; }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 2em 0;
            font-size: 0.95em;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 12px;
            vertical-align: top;
        }}
        th {{
            background-color: #f8f9fa;
            font-weight: 600;
        }}
        .display-equation {{
            display: flex;
            justify-content: center;
            align-items: center;
            margin: 2em 0;
            width: 100%;
            overflow-x: auto;
            overflow-y: hidden;
        }}
        math[display="block"] {{
            padding: 0.5em 0;
        }}
        math[display="inline"] {{
            vertical-align: middle;
        }}
        .responsive-img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 2em auto;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }}
        .tab-space {{
            display: inline-block;
            width: 2.5em;
        }}
        ul, ol {{
            margin-bottom: 1.2em;
            padding-left: 2em;
        }}
        li {{
            margin-bottom: 0.4em;
        }}
        .math-error {{
            color: #d9534f;
            font-weight: bold;
            border: 1px dashed #d9534f;
            padding: 2px 4px;
            font-size: 0.85em;
        }}
    </style>
    <script type="text/javascript" async
      src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/mml-chtml.js">
    </script>
</head>
<body>
    {content}
</body>
</html>
"""
        out_p = pathlib.Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        with open(out_p, "w", encoding="utf-8") as f:
            f.write(template)


class DocumentConverter:
    # Главный класс, который все собирает воедино
    def __init__(self, docx_path, xsl_path):
        print("Загружаем ворд...")
        self.reader = DocxReader(docx_path)

        print("Погнали парсить XML...")
        doc_xml = self.reader.read_file("word/document.xml")
        if not doc_xml:
            raise ValueError("Внутри docx нет файла word/document.xml, архив битый!")

        print("Вытаскиваем настройки стилей и списков...")
        styles_xml = self.reader.read_file("word/styles.xml")
        numbering_xml = self.reader.read_file("word/numbering.xml")
        rels_xml = self.reader.read_file("word/_rels/document.xml.rels")

        self.style_resolver = StyleResolver(styles_xml)
        self.numbering_resolver = NumberingResolver(numbering_xml)
        self.relationship_resolver = RelationshipResolver(rels_xml)
        self.math_converter = MathConverter(xsl_path)

        # Счётчики для красивой статистики в конце
        self.p_count = 0
        self.t_count = 0
        self.img_count = 0
        self.f_count = 0

        try:
            self.document_root = ET.fromstring(doc_xml)
        except ET.XMLSyntaxError as e:
            raise ValueError(f"XML сломался на строке {e.position[0]}, кол {e.position[1]}") from e

    def convert(self, output_html_path):
        print("Конвертируем формулы...")
        print("Вытаскиваем картинки...")
        
        body = self.document_root.find("w:body", word_ns)
        if body is None:
            raise ValueError("Не найден тег <w:body> в документе!")

        print("Пишем итоговый HTML...")
        html_output_buffer = self.render_container(body)
        HtmlWriter.write(html_output_buffer, output_html_path)
        self.reader.close()

        print("Всё готово!")
        print(f"\nКоличество формул: {self.f_count}")
        print(f"Количество таблиц: {self.t_count}")
        print(f"Количество картинок: {self.img_count}")
        print(f"Количество абзацев: {self.p_count}")

    def render_container(self, parent_node):
        # Рекурсивный обход элементов, тут аккуратно со списками
        rezultat = []
        stek_spiskov = []

        for child in parent_node:
            is_list = False
            num_id, ilvl, list_type = None, None, "ul"

            if child.tag == f'{{{word_ns["w"]}}}p':
                num_id, ilvl = self.numbering_resolver.get_list_info(child)
                if num_id is not None and ilvl is not None:
                    is_list = True
                    list_type = self.numbering_resolver.get_list_type(num_id, ilvl)

            if is_list:
                while stek_spiskov and stek_spiskov[-1]["ilvl"] > ilvl:
                    closed = stek_spiskov.pop()
                    if closed["li_open"]:
                        rezultat.append("</li>")
                    rezultat.append(f"</{closed['tag']}>")

                if stek_spiskov and stek_spiskov[-1]["ilvl"] == ilvl and (stek_spiskov[-1]["num_id"] != num_id or stek_spiskov[-1]["tag"] != list_type):
                    closed = stek_spiskov.pop()
                    if closed["li_open"]:
                        rezultat.append("</li>")
                    rezultat.append(f"</{closed['tag']}>")

                if not stek_spiskov or stek_spiskov[-1]["ilvl"] < ilvl:
                    rezultat.append(f"<{list_type}>")
                    stek_spiskov.append({"num_id": num_id, "ilvl": ilvl, "tag": list_type, "li_open": False})

                current_list = stek_spiskov[-1]
                if current_list["li_open"]:
                    rezultat.append("</li>")

                rezultat.append("<li>")
                current_list["li_open"] = True

                self.p_count += 1
                rezultat.append(self.render_inline_elements(child))
            else:
                while stek_spiskov:
                    closed = stek_spiskov.pop()
                    if closed["li_open"]:
                        rezultat.append("</li>")
                    rezultat.append(f"</{closed['tag']}>")

                rezultat.append(self.render_block_node(child))

        while stek_spiskov:
            closed = stek_spiskov.pop()
            if closed["li_open"]:
                rezultat.append("</li>")
            rezultat.append(f"</{closed['tag']}>")

        return "".join(rezultat)

    def render_block_node(self, node):
        # Обработка блочных элементов типа таблиц или параграфов
        tag = node.tag

        if tag == f'{{{word_ns["w"]}}}p':
            self.p_count += 1
            style_id = None
            p_style = node.find(".//w:pPr/w:pStyle", word_ns)
            if p_style is not None:
                style_id = p_style.get(f'{{{word_ns["w"]}}}val')

            heading_tag = self.style_resolver.get_heading_level(style_id) if style_id else None
            inner_content = self.render_inline_elements(node)

            if not inner_content.strip():
                return ""

            if heading_tag:
                return f"<{heading_tag}>{inner_content}</{heading_tag}>"
            return f"<p>{inner_content}</p>"

        elif tag == f'{{{word_ns["w"]}}}tbl':
            self.t_count += 1
            rows_buffer = []
            for child in node:
                if child.tag == f'{{{word_ns["w"]}}}tr':
                    rows_buffer.append(self.render_block_node(child))
            return f"<table>{''.join(rows_buffer)}</table>"

        elif tag == f'{{{word_ns["w"]}}}tr':
            cells_buffer = []
            for child in node:
                if child.tag == f'{{{word_ns["w"]}}}tc':
                    cells_buffer.append(self.render_block_node(child))
            return f"<tr>{''.join(cells_buffer)}</tr>"

        elif tag == f'{{{word_ns["w"]}}}tc':
            colspan_attr = ""
            grid_span = node.find(".//w:tcPr/w:gridSpan", word_ns)
            if grid_span is not None:
                span_val = grid_span.get(f'{{{word_ns["w"]}}}val')
                if span_val:
                    colspan_attr = f' colspan="{span_val}"'

            cell_html = self.render_container(node)
            return f"<td{colspan_attr}>{cell_html}</td>"

        elif tag == f'{{{word_ns["m"]}}}oMathPara':
            self.f_count += 1
            formulas_buffer = []
            for child in node:
                if child.tag == f'{{{word_ns["m"]}}}oMath':
                    formulas_buffer.append(self.math_converter.convert(child, is_block=True))
            return f'<div class="display-equation">{"".join(formulas_buffer)}</div>'

        elif tag == f'{{{word_ns["m"]}}}oMath':
            self.f_count += 1
            return self.math_converter.convert(node, is_block=False)

        else:
            # Игнорируем технические свойства ворда, чтобы не спамить
            ignored_metadata = {
                f'{{{word_ns["w"]}}}pPr', f'{{{word_ns["w"]}}}rPr', f'{{{word_ns["w"]}}}tblPr',
                f'{{{word_ns["w"]}}}trPr', f'{{{word_ns["w"]}}}tcPr', f'{{{word_ns["w"]}}}sectPr'
            }
            if tag in ignored_metadata:
                return ""
            
            fallback_buffer = []
            for child in node:
                fallback_buffer.append(self.render_block_node(child))
            return "".join(fallback_buffer)

    def render_inline_elements(self, parent_node):
        # Собираем строчные элементы (текст, ссылки, инлайн формулы)
        inline_buffer = []
        for child in parent_node:
            tag = child.tag
            if tag == f'{{{word_ns["w"]}}}r':
                inline_buffer.append(self.render_run(child))
            elif tag == f'{{{word_ns["w"]}}}hyperlink':
                inline_buffer.append(self.render_hyperlink(child))
            elif tag == f'{{{word_ns["m"]}}}oMath':
                self.f_count += 1
                inline_buffer.append(self.math_converter.convert(child, is_block=False))
            elif tag == f'{{{word_ns["m"]}}}oMathPara':
                inline_buffer.append(self.render_block_node(child))
            elif tag == f'{{{word_ns["w"]}}}br':
                inline_buffer.append("<br/>")
            elif tag == f'{{{word_ns["w"]}}}tab':
                inline_buffer.append('<span class="tab-space"></span>')
            elif tag != f'{{{word_ns["w"]}}}pPr':
                inline_buffer.append(self.render_block_node(child))
        return "".join(inline_buffer)

    def render_run(self, run_node):
        # Парсим конкретный кусочек текста и смотрим его стили (жирный, курсив и т.д.)
        r_pr = run_node.find("w:rPr", word_ns)
        
        is_bold = r_pr is not None and (r_pr.find("w:b", word_ns) is not None or r_pr.find("w:bCs", word_ns) is not None)
        is_italic = r_pr is not None and (r_pr.find("w:i", word_ns) is not None or r_pr.find("w:iCs", word_ns) is not None)
        is_underline = r_pr is not None and r_pr.find("w:u", word_ns) is not None
        is_strike = r_pr is not None and r_pr.find("w:strike", word_ns) is not None

        is_super, is_sub = False, False
        if r_pr is not None:
            vert_align = r_pr.find("w:vertAlign", word_ns)
            if vert_align is not None:
                align_val = vert_align.get(f'{{{word_ns["w"]}}}val', "")
                if align_val == "superscript":
                    is_super = True
                elif align_val == "subscript":
                    is_sub = True

        run_buffer = []
        for child in run_node:
            c_tag = child.tag
            if c_tag == f'{{{word_ns["w"]}}}t':
                run_buffer.append(html.escape(child.text or ""))
            elif c_tag == f'{{{word_ns["w"]}}}br':
                run_buffer.append("<br/>")
            elif c_tag == f'{{{word_ns["w"]}}}tab':
                run_buffer.append('<span class="tab-space"></span>')
            elif c_tag == f'{{{word_ns["w"]}}}drawing':
                run_buffer.append(self.render_drawing(child))
            elif c_tag == f'{{{word_ns["m"]}}}oMath':
                self.f_count += 1
                run_buffer.append(self.math_converter.convert(child, is_block=False))

        inner_html = "".join(run_buffer)
        if is_bold: inner_html = f"<strong>{inner_html}</strong>"
        if is_italic: inner_html = f"<em>{inner_html}</em>"
        if is_underline: inner_html = f"<u>{inner_html}</u>"
        if is_strike: inner_html = f"<del>{inner_html}</del>"
        if is_super: inner_html = f"<sup>{inner_html}</sup>"
        if is_sub: inner_html = f"<sub>{inner_html}</sub>"
        return inner_html

    def render_hyperlink(self, link_node):
        r_id = link_node.get(f'{{{word_ns["r"]}}}id')
        target_url = self.relationship_resolver.get_target(r_id) if r_id else "#"
        inner_content = self.render_inline_elements(link_node)
        return f'<a href="{html.escape(target_url or "#")}">{inner_content}</a>'

    def render_drawing(self, drawing_node):
        blip = drawing_node.find(".//a:blip", word_ns)
        if blip is not None:
            r_id = blip.get(f'{{{word_ns["r"]}}}embed') or blip.get(f'{{{word_ns["r"]}}}link')
            if r_id:
                target_path = self.relationship_resolver.get_target(r_id)
                if target_path:
                    filename = pathlib.Path(target_path).name
                    dest_file_path = f"images/{filename}"
                    self.reader.extract_media(target_path, dest_file_path)
                    self.img_count += 1
                    return f'<img src="{dest_file_path}" alt="Embedded Image" class="responsive-img"/>'
        return ""


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="Простой конвертер из DOCX в HTML.")
    parser.add_argument("input_file", nargs="?", default="x.docx", help="Путь к файлу DOCX.")
    parser.add_argument("-o", "--output", default="index.html", help="Куда сохранить HTML.")
    parser.add_argument("-x", "--xsl", default="omml2mml.xsl", help="Путь к файлу omml2mml.xsl.")
    
    args = parser.parse_args()
    input_path = pathlib.Path(args.input_file)
    xsl_file = args.xsl
    output_html = args.output

    if input_path.exists() == False:
        print(f"\nОшибка: Файл '{input_path}' не найден.", file=sys.stderr)
        sys.exit(1)

    file_extension = input_path.suffix.lower()
    target_docx = str(input_path)
    
    if file_extension != ".docx":
        print(f"\nОшибка: Расширение '{file_extension}' не поддерживается. Нужен только .docx файл.", file=sys.stderr)
        sys.exit(1)

    try:
        print("DOCX:")
        converter = DocumentConverter(docx_path=target_docx, xsl_path=xsl_file)
        converter.convert(output_html_path=output_html)
                
    except FileNotFoundError as e:
        print(f"\nОшибка конфигурации: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\nЧто-то упало при конвертации: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
