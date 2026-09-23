"""Local experimental HWP-to-DOCX conversion.

본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
"""

import argparse
import base64
import hashlib
from io import BytesIO
import json
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import olefile
from hwp_reader import RECOVERY_PREFIX
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


NOTICE = "본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다."
WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def word_property(parent, name, **attributes):
    element = OxmlElement("w:" + name)
    for key, value in attributes.items():
        element.set(qn("w:" + key), str(value))
    parent.append(element)
    return element


def read_hwp(source, recover_text=False):
    if source.suffix.lower() != ".hwp":
        raise ValueError("This prototype supports binary .hwp files only, not HWPX.")
    if source.stat().st_size > 50 * 1024 * 1024:
        raise ValueError("The prototype accepts files up to 50 MB.")
    with olefile.OleFileIO(str(source)) as container:
        header = container.openstream("FileHeader").read()
    if not header.startswith(b"HWP Document File") or len(header) < 40:
        raise ValueError("Not a supported HWP document.")
    if header[35] != 5:
        raise ValueError("Only HWP version 5 files are supported.")
    flags = struct.unpack_from("<I", header, 36)[0]
    if flags & ((1 << 1) | (1 << 2) | (1 << 4) | (1 << 8) | (1 << 10)):
        raise ValueError("Password, distribution, or DRM protected files are not converted.")
    reader = Path(sys.executable).parent / "hwp5proc"
    command = [str(reader), "xml", "--embedbin", str(source)]
    if recover_text:
        command = [sys.executable, str(Path(__file__).with_name("hwp_reader.py")), str(source)]
    result = subprocess.run(
        command, capture_output=True, timeout=60
    )
    if result.returncode:
        raise ValueError("The HWP reader failed; no output was created.")
    replacements = []
    diagnostics = []
    for line in result.stderr.decode(errors="replace").splitlines():
        if line.startswith(RECOVERY_PREFIX):
            replacements = json.loads(line[len(RECOVERY_PREFIX):])
        elif line.strip():
            diagnostics.append(line)
    warnings = ["HWP reader reported a warning; fidelity needs review."] if diagnostics else []
    if replacements:
        warnings.append(f"Replaced {len(replacements)} invalid UTF-16 sequence(s) with U+FFFD; review text_replacements in the report.")
    root = ET.fromstring(result.stdout)
    root.set("text-replacements", json.dumps(replacements))
    return root, warnings


class Converter:
    def __init__(self, root, warnings, font_map=None, preserve_source_pages=False):
        self.root = root
        self.preserve_source_pages = preserve_source_pages
        self.font_map = font_map or {}
        self.document = Document()
        self.warnings = set(warnings)
        self.tables = 0
        self.images = 0
        self.assets = {node.get("storage-id"): node for node in root.iter("BinDataEmbedding")}
        self.char_shapes = list(root.iter("CharShape"))
        self.para_shapes = list(root.iter("ParaShape"))
        self.tab_defs = list(root.iter("TabDef"))
        self.borders = list(root.iter("BorderFill"))
        mappings = root.find("DocInfo/IdMappings")
        faces = list(root.iter("FaceName"))
        self.fonts = {}
        offset = 0
        for language in ("ko", "en", "cn", "jp", "other", "symbol", "user"):
            count = int(mappings.get(language + "-fonts", "0")) if mappings is not None else 0
            self.fonts[language] = [face.get("name") for face in faces[offset:offset + count]]
            offset += count

    def style_run(self, run, shape_id, language="ko"):
        if not self.char_shapes:
            return
        shape = self.char_shapes[int(shape_id)]
        run.bold = shape.get("bold") == "1"
        run.italic = shape.get("italic") == "1"
        run.underline = shape.get("underline", "none") != "none"
        run.font.size = Pt(int(shape.get("basesize", "1000")) / 100)
        run.font.color.rgb = RGBColor.from_string(shape.get("text-color", "#000000").lstrip("#"))
        spacing = shape.find("LetterSpacing")
        if spacing is not None:
            percentage = int(spacing.get(language, spacing.get("ko", "0")))
            word_property(run._element.get_or_add_rPr(), "spacing", val=round(run.font.size.pt * percentage / 100 * 20))
        expansion = shape.find("LetterWidthExpansion")
        if expansion is not None:
            word_property(run._element.get_or_add_rPr(), "w", val=expansion.get(language, expansion.get("ko", "100")))
        face = shape.find("FontFace")
        if face is not None and self.fonts["ko"] and self.fonts["en"]:
            korean = self.fonts["ko"][int(face.get("ko", "0"))]
            latin = self.fonts["en"][int(face.get("en", "0"))]
            for name in (korean, latin):
                if name in self.font_map:
                    self.warnings.add("Font substitution: " + name + " -> " + self.font_map[name])
            korean = self.font_map.get(korean, korean)
            latin = self.font_map.get(latin, latin)
            run.font.name = latin
            run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), korean)

    def style_paragraph(self, source, paragraph):
        settings = paragraph.paragraph_format
        settings.space_before = Pt(0)
        settings.space_after = Pt(0)
        if not self.para_shapes:
            return
        shape = self.para_shapes[int(source.get("parashape-id", "0"))]
        paragraph.alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT, "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "center": WD_ALIGN_PARAGRAPH.CENTER, "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
        }.get(shape.get("align"), WD_ALIGN_PARAGRAPH.LEFT)
        indent = int(shape.get("indent", "0")) / 100
        settings.left_indent = Pt(int(shape.get("doubled-margin-left", "0")) / 200 + max(0, -indent))
        settings.right_indent = Pt(int(shape.get("doubled-margin-right", "0")) / 200)
        settings.first_line_indent = Pt(indent)
        settings.space_before = Pt(int(shape.get("doubled-margin-top", "0")) / 200)
        settings.space_after = Pt(int(shape.get("doubled-margin-bottom", "0")) / 200)
        if shape.get("linespacing-type") == "ratio":
            lines = source.findall("LineSeg")
            if lines:
                settings.line_spacing = Pt(max(int(line.get("height", "1000")) + int(line.get("space-below", "0")) for line in lines) / 100)
            else:
                settings.line_spacing = int(shape.get("linespacing", "100")) / 100
        else:
            self.warnings.add("Non-proportional paragraph line spacing is approximated.")
        settings.keep_with_next = shape.get("with-next-paragraph") == "1"
        settings.keep_together = shape.get("protect") == "1"
        settings.widow_control = shape.get("protect-single-line") == "1"
        settings.page_break_before = source.get("new-page") == "1" or shape.get("start-new-page") == "1"
        tab_id = int(shape.get("tabdef-id", "0"))
        if tab_id < len(self.tab_defs):
            alignments = {"left": WD_TAB_ALIGNMENT.LEFT, "right": WD_TAB_ALIGNMENT.RIGHT,
                          "center": WD_TAB_ALIGNMENT.CENTER, "decimal": WD_TAB_ALIGNMENT.DECIMAL}
            for tab in self.tab_defs[tab_id].findall("Array/Tab"):
                settings.tab_stops.add_tab_stop(Pt(int(tab.get("pos", "0")) / 100),
                    alignments.get(tab.get("kind"), WD_TAB_ALIGNMENT.LEFT),
                    WD_TAB_LEADER.DOTS if tab.get("fill-type") == "3" else WD_TAB_LEADER.SPACES)
        if shape.get("head-shape", "none") != "none":
            self.warnings.add("Automatic numbering/bullets are not yet mapped.")

    def style_cell(self, source, cell, body=None):
        cell.vertical_alignment = {
            "top": WD_CELL_VERTICAL_ALIGNMENT.TOP, "middle": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
            "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
        }.get(source.get("valign"), WD_CELL_VERTICAL_ALIGNMENT.TOP)
        properties = cell._tc.get_or_add_tcPr()
        margins = word_property(properties, "tcMar")
        padding = source
        lines = source.findall("Paragraph/LineSeg")
        if body is not None and lines:
            expected_width = int(source.get("width", "0")) - int(body.get("padding-left", "0")) - int(body.get("padding-right", "0"))
            if all(abs(int(line.get("width", "0")) - expected_width) <= 2 for line in lines):
                padding = body
        for side in ("top", "left", "bottom", "right"):
            word_property(margins, side, w=round(int(padding.get("padding-" + side, "0")) / 5), type="dxa")
        border_id = int(source.get("borderfill-id", "0"))
        if not border_id or border_id > len(self.borders):
            return
        fill = self.borders[border_id - 1]
        borders = word_property(properties, "tcBorders")
        strokes = {"none": "nil", "solid": "single", "double-2": "double", "dashed": "dashed", "dotted": "dotted"}
        for border in fill.findall("Border"):
            side = border.get("attribute-name")
            if side not in ("top", "left", "bottom", "right"):
                continue
            width = float(border.get("width", "0.1mm").removesuffix("mm"))
            stroke = border.get("stroke-type", "none")
            if stroke not in strokes:
                self.warnings.add("Some border styles are approximated.")
            word_property(borders, side, val=strokes.get(stroke, "single"),
                          sz=max(2, round(width * 72 / 25.4 * 8)), color=border.get("color", "#000000").lstrip("#"))
        color = fill.find("FillColorPattern")
        if color is not None:
            word_property(properties, "shd", val="clear", fill=color.get("background-color", "#ffffff").lstrip("#"))

    def append_text(self, item, paragraph):
        run = paragraph.add_run(item.text or "")
        self.style_run(run, item.get("charshape-id", "0"), item.get("lang", "ko"))
        for child in item:
            if child.tag == "Tab":
                run.add_tab()
            if child.tail:
                run.add_text(child.tail)

    def drawing(self, item, paragraph):
        for node in item.iter():
            if node.tag == "Text":
                self.append_text(node, paragraph)
            elif node.tag == "ShapePicture":
                picture = node.find("PictureInfo")
                identifier = int(picture.get("bindata-id"))
                asset = self.assets.get(f"BIN{identifier:04X}")
                if asset is None or asset.get("inline") != "true":
                    raise ValueError(f"Embedded image {identifier} is unavailable; no output published.")
                data = base64.b64decode("".join((asset.text or "").split()), validate=True)
                width = int(item.get("width", "0"))
                height = int(item.get("height", "0"))
                if width <= 0 or height <= 0:
                    raise ValueError("Image dimensions must be positive.")
                try:
                    paragraph.add_run().add_picture(BytesIO(data), width=Inches(width / 7200), height=Inches(height / 7200))
                except Exception as error:
                    raise ValueError(f"Unable to insert embedded image {identifier} ({asset.get('ext')}).") from error
                self.images += 1
                paragraph.paragraph_format.line_spacing = 1.0
        if list(item.iter("TextboxParagraphList")):
            self.warnings.add("Text boxes are flattened to editable inline text; box layout is not preserved.")
        self.warnings.add("Drawings are flattened: floating positions, crops, rotations, and effects are not preserved.")

    def paragraph(self, source, target):
        paragraph = target.add_paragraph()
        self.style_paragraph(source, paragraph)
        follows_table = False
        for line in source.findall("LineSeg"):
            for item in line:
                if item.tag == "Text":
                    follows_table = False
                    self.append_text(item, paragraph)
                elif item.tag == "GShapeObjectControl":
                    follows_table = False
                    self.drawing(item, paragraph)
                elif item.tag == "FieldHyperLink":
                    for text in item.iter("Text"):
                        self.append_text(text, paragraph)
                    self.warnings.add("Hyperlink labels are preserved as editable text, without active links.")
                elif item.tag == "TableControl":
                    if not paragraph.text:
                        paragraph._element.getparent().remove(paragraph._element)
                    self.table(item, target, paragraph.alignment)
                    paragraph = target.add_paragraph()
                    self.style_paragraph(source, paragraph)
                    follows_table = True
                elif item.tag == "ControlChar":
                    name = item.get("name")
                    if name == "LINE_BREAK":
                        paragraph.add_run().add_break()
                    elif name == "TAB":
                        paragraph.add_run().add_tab()
                    elif name == "FIXWIDTH_SPACE":
                        paragraph.add_run("\u2007")
                        self.warnings.add("Fixed-width spaces are approximated with Unicode figure spaces.")
                    elif name != "PARAGRAPH_BREAK":
                        self.warnings.add("Unsupported text control: " + str(name))
                elif item.tag == "ColumnsDef":
                    if item.get("count") != "1":
                        self.warnings.add("Multiple columns are flattened.")
                elif item.tag in ("Header", "Footer") and not list(item.iter("Text")):
                    continue
                else:
                    self.warnings.add("Unsupported paragraph element: " + item.tag)
        if follows_table and not paragraph.runs:
            settings = paragraph.paragraph_format
            settings.line_spacing = Pt(1)
            settings.space_before = Pt(0)
            settings.space_after = Pt(0)
            paragraph.add_run().font.size = Pt(1)
        elif not paragraph.runs:
            last = paragraph._element.getparent()
            if last is not None:
                line = source.find("LineSeg")
                if line is not None:
                    paragraph.add_run().font.size = Pt(int(line.get("height-text", "1000")) / 100)

    def table(self, source, target, paragraph_alignment=None):
        body = source.find("TableBody")
        rows, cols = int(body.get("rows")), int(body.get("cols"))
        if rows < 1 or cols < 1 or rows * cols > 10000:
            raise ValueError("Table dimensions exceed prototype limits.")
        width = Inches(int(source.get("width", "43200")) / 7200)
        table = target.add_table(rows, cols)
        table.autofit = False
        if source.get("inline", "1") == "1":
            table.alignment = {
                WD_ALIGN_PARAGRAPH.CENTER: WD_TABLE_ALIGNMENT.CENTER,
                WD_ALIGN_PARAGRAPH.RIGHT: WD_TABLE_ALIGNMENT.RIGHT,
            }.get(paragraph_alignment, WD_TABLE_ALIGNMENT.LEFT)
        preferred_width = table._tbl.tblPr.find(qn("w:tblW"))
        preferred_width.set(qn("w:type"), "dxa")
        preferred_width.set(qn("w:w"), str(round(width.twips)))
        if source.get("inline", "1") != "1":
            self.warnings.add("Floating tables are converted to inline tables.")
        self.tables += 1
        cells = body.findall("TableRow/TableCell")
        for item in cells:
            row, col = int(item.get("row")), int(item.get("col"))
            rowspan, colspan = int(item.get("rowspan", "1")), int(item.get("colspan", "1"))
            cell = table.cell(row, col)
            if rowspan > 1 or colspan > 1:
                cell = cell.merge(table.cell(row + rowspan - 1, col + colspan - 1))
            cell.width = Inches(int(item.get("width", "7200")) / 7200)
            if colspan == 1:
                table.columns[col].width = cell.width
            self.style_cell(item, cell, body)
            if rowspan == 1:
                table.rows[row].height = Inches(int(item.get("height", "0")) / 7200)
                table.rows[row].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            for paragraph in list(cell.paragraphs):
                cell._tc.remove(paragraph._element)
            for paragraph in item.findall("Paragraph"):
                self.paragraph(paragraph, cell)
                lines = paragraph.findall("LineSeg")
                if lines and len(lines) == 1 and not paragraph.findall(".//TableControl"):
                    cell.paragraphs[-1].paragraph_format.line_spacing = Pt(int(lines[0].get("height", "1000")) / 100)
            if not cell.paragraphs:
                cell.add_paragraph()
        if source.find("TableCaption") is not None:
            self.warnings.add("Table captions are not yet supported.")

    def convert(self):
        sections = self.root.findall("BodyText/SectionDef")
        if len(sections) != 1:
            raise ValueError("This first prototype requires a single document section.")
        if self.preserve_source_pages and any(node.get("count", "1") != "1" for node in sections[0].iter("ColumnsDef")):
            raise ValueError("Source page inference is only supported for single-column documents.")
        page = sections[0].find("PageDef")
        if page is not None:
            section = self.document.sections[0]
            for target, attribute in (("page_width", "width"), ("page_height", "height"),
                                      ("top_margin", "top-offset"), ("bottom_margin", "bottom-offset"),
                                      ("left_margin", "left-offset"), ("right_margin", "right-offset"),
                                      ("header_distance", "header-offset"), ("footer_distance", "footer-offset")):
                setattr(section, target, Inches(int(page.get(attribute, "0")) / 7200))
        self.warnings.add("Font availability, pagination, and exact layout require review in Word.")
        previous_y = None
        for paragraph in sections[0].findall("ColumnSet/Paragraph"):
            lines = paragraph.findall("LineSeg")
            first_y = int(lines[0].get("y", "0")) if lines else None
            source_page_break = (self.preserve_source_pages and first_y is not None
                                 and previous_y is not None and first_y < previous_y)
            before = len(self.document.paragraphs)
            self.paragraph(paragraph, self.document)
            if source_page_break and len(self.document.paragraphs) > before:
                self.document.paragraphs[before].paragraph_format.page_break_before = True
            if lines:
                previous_y = int(lines[-1].get("y", "0"))
        if self.preserve_source_pages:
            self.warnings.add("Page breaks inferred from stored line positions; not a visual fidelity guarantee.")
        return self.document


def validate_output(root, output):
    with ZipFile(output) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
        if archive.testzip() is not None:
            raise ValueError("Generated DOCX archive is damaged.")
    expected = "".join((node.text or "") + "".join(child.tail or "" for child in node) for node in root.findall("BodyText//Text"))
    actual = "".join(node.text or "" for node in document.iter(WORD_NS + "t"))
    fixed_spaces = sum(node.get("name") == "FIXWIDTH_SPACE" for node in root.findall("BodyText//ControlChar"))
    if fixed_spaces:
        expected = expected.replace("\u2007", "")
        actual = actual.replace("\u2007", "")
    if expected != actual:
        raise ValueError("Text preservation check failed; output is not verified.")
    source_tables = list(root.iter("TableBody"))
    output_tables = list(document.iter(WORD_NS + "tbl"))
    if len(source_tables) != len(output_tables):
        raise ValueError("Table preservation check failed.")
    for source_table, output_table in zip(source_tables, output_tables):
        if len(output_table.findall(WORD_NS + "tr")) != int(source_table.get("rows")):
            raise ValueError("Table row count changed.")
        if len(output_table.find(WORD_NS + "tblGrid")) != int(source_table.get("cols")):
            raise ValueError("Table column count changed.")
    Document(output)
    pictures = list(root.findall("BodyText//ShapePicture"))
    drawings = list(document.iter("{http://schemas.openxmlformats.org/drawingml/2006/picture}pic"))
    if len(pictures) != len(drawings):
        raise ValueError("Image occurrence count changed.")
    return {"text_characters": len(actual), "text_sha256": hashlib.sha256(actual.encode()).hexdigest(),
            "tables": len(output_tables), "text_preserved": True, "table_dimensions_preserved": True,
            "images": len(drawings), "image_occurrences_preserved": True,
            "fixed_width_spaces_approximated": fixed_spaces}


def main():
    parser = argparse.ArgumentParser(description=NOTICE)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--font-map", type=Path, help="Optional JSON mapping from original font families to installed alternatives.")
    parser.add_argument("--recover-text", action="store_true", help="Replace malformed UTF-16 with U+FFFD and report each replacement; original file is unchanged.")
    parser.add_argument("--preserve-source-pages", action="store_true", help="Infer top-level page breaks from stored HWP line positions; intended for single-column documents.")
    args = parser.parse_args()
    output = args.output
    if output.suffix.lower() != ".docx":
        parser.error("Output must have a .docx extension.")
    report_path = output.with_suffix(".report.json")
    if output.exists() or report_path.exists():
        parser.error("Output or report already exists; choose a new name.")
    try:
        root, warnings = read_hwp(args.source, recover_text=args.recover_text)
        font_map = json.loads(args.font_map.read_text(encoding="utf-8")) if args.font_map else {}
        if not isinstance(font_map, dict) or not all(isinstance(key, str) and isinstance(value, str) and value for key, value in font_map.items()):
            raise ValueError("Font map must be a JSON object mapping font names to nonempty strings.")
        converter = Converter(root, warnings, font_map, preserve_source_pages=args.preserve_source_pages)
        document = converter.convert()
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=output.parent) as temporary:
            candidate = Path(temporary) / "candidate.docx"
            document.save(candidate)
            report = validate_output(root, candidate)
            candidate.replace(output)
        report["warnings"] = sorted(converter.warnings)
        report["text_replacements"] = json.loads(root.get("text-replacements", "[]"))
        report["text_comparison_basis"] = "Recovered parser text" if report["text_replacements"] else "Parser text"
        report["visual_layout_verified"] = False
        report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    except (ValueError, OSError, subprocess.TimeoutExpired, ET.ParseError) as error:
        parser.exit(1, str(error) + "\n")
    print(json.dumps(report, indent=2))
    print("Created:", output)


if __name__ == "__main__":
    main()