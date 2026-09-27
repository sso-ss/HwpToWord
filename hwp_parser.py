"""HWP 5 record decoding from Hancom's published file format specification."""

from dataclasses import dataclass
from bisect import bisect_right
import base64
import json
from pathlib import Path
import re
import struct
import xml.etree.ElementTree as ET
import zlib

import olefile

from hwp_reader import TextRecovery
from hwp_limits import HwpLimits, MIB


MAX_STREAM_BYTES = 64 * 1024 * 1024
MAX_RECORDS = 200000


@dataclass(frozen=True)
class Record:
    tag: int
    level: int
    data: bytes


def decompress_stream(data, limit=MAX_STREAM_BYTES):
    decoder = zlib.decompressobj(-15)
    try:
        decoded = decoder.decompress(data, limit + 1)
    except zlib.error as error:
        raise ValueError("Invalid compressed HWP stream.") from error
    if len(decoded) > limit or decoder.unconsumed_tail:
        raise ValueError("Decompressed HWP stream exceeds the size limit.")
    if not decoder.eof:
        raise ValueError("Incomplete compressed HWP stream data.")
    trailer = decoder.unused_data
    if trailer and trailer != struct.pack("<II", zlib.crc32(decoded), len(decoded)):
        raise ValueError("Invalid compressed HWP stream checksum or trailing data.")
    return decoded


def records(data, max_bytes=MAX_STREAM_BYTES, max_records=MAX_RECORDS):
    if len(data) > max_bytes:
        raise ValueError("HWP stream exceeds the size limit.")
    offset = 0
    count = 0
    previous_level = 0
    while offset < len(data):
        if len(data) - offset < 4:
            raise ValueError("Truncated HWP record header.")
        header = struct.unpack_from("<I", data, offset)[0]
        offset += 4
        tag, level, size = header & 1023, (header >> 10) & 1023, header >> 20
        if size == 4095:
            if len(data) - offset < 4:
                raise ValueError("Truncated HWP extended record size.")
            size = struct.unpack_from("<I", data, offset)[0]
            offset += 4
        if size > len(data) - offset:
            raise ValueError("Truncated HWP record payload.")
        if (count == 0 and level != 0) or level > previous_level + 1 or level > 64:
            raise ValueError("Invalid or excessive HWP record nesting.")
        count += 1
        if count > max_records:
            raise ValueError("HWP record count exceeds the limit.")
        yield Record(tag, level, data[offset:offset + size])
        previous_level = level
        offset += size


LANGUAGES = ("ko", "en", "cn", "jp", "other", "symbol", "user")
SIDES = ("left", "right", "top", "bottom")


def unpack(data, pattern, offset=0):
    if offset < 0 or offset + struct.calcsize("<" + pattern) > len(data):
        raise ValueError("Truncated HWP record fields.")
    return struct.unpack_from("<" + pattern, data, offset)


def element(tag, **attributes):
    return ET.Element(tag, {key.replace("_", "-"): str(value) for key, value in attributes.items()})


def color(value):
    return "#" + value.to_bytes(4, "little")[:3].hex()


@dataclass
class RecordNode:
    record: Record
    children: list


def record_tree(data, max_bytes=MAX_STREAM_BYTES, max_records=MAX_RECORDS):
    root = RecordNode(Record(0, -1, b""), [])
    stack = [root]
    for item in records(data, max_bytes, max_records):
        while stack[-1].record.level >= item.level:
            stack.pop()
        node = RecordNode(item, [])
        stack[-1].children.append(node)
        stack.append(node)
    return root.children


class HwpParser:
    def __init__(self, recover_text=False, limits=None, asset_directory=None):
        self.limits = limits if limits is not None else HwpLimits()
        self.asset_directory = Path(asset_directory) if asset_directory is not None else None
        self.recovery = TextRecovery() if recover_text else None
        self.warnings = set()
        self.root = element("HwpDoc", parser="native-hwp5")
        self.section = None
        self.assets = []
        self.version = (5, 0, 0, 0)
        self.stream_bytes = 0
        self.section_count = None

    def decode(self, data):
        try:
            text = self.recovery.decode(data) if self.recovery else data.decode("utf-16le")
        except UnicodeDecodeError as error:
            raise ValueError("Invalid UTF-16 text; use --recover-text only if character replacement is acceptable.") from error
        if any(0xE000 <= ord(char) <= 0xF8FF for char in text):
            self.warnings.add("Private-use characters are retained without legacy Hanyang font-code conversion.")
        if any(ord(char) < 32 and char not in "\t\n\r" or ord(char) in (0xFFFE, 0xFFFF) for char in text):
            raise ValueError("Text contains characters that cannot be represented in DOCX XML.")
        return text

    def doc_info(self, data):
        info = ET.SubElement(self.root, "DocInfo")
        for item in records(data, self.limits.max_stream_mb * MIB, self.limits.max_records):
            payload = item.data
            if item.tag == 16:
                self.section_count = unpack(payload, "H")[0]
            elif item.tag == 17:
                counts = unpack(payload, "8I")
                info.append(element("IdMappings", **{language + "-fonts": counts[index + 1]
                                                     for index, language in enumerate(LANGUAGES)}))
            elif item.tag == 18:
                flags = unpack(payload, "H")[0]
                if flags & 15 == 1:
                    identifier, length = unpack(payload, "HH", 2)
                    unpack(payload, str(length * 2) + "s", 6)
                    extension = self.decode(payload[6:6 + length * 2])
                    if not extension.isalnum():
                        raise ValueError("Invalid embedded asset extension.")
                    self.assets.append((identifier, extension, (flags >> 4) & 3))
                else:
                    self.warnings.add("Linked images and embedded OLE objects are unsupported.")
            elif item.tag == 19:
                length = unpack(payload, "H", 1)[0]
                name = unpack(payload, str(length * 2) + "s", 3)[0]
                info.append(element("FaceName", name=self.decode(name)))
            elif item.tag == 20:
                info.append(self.border_fill(payload))
            elif item.tag == 21:
                fonts = unpack(payload, "7H")
                widths = unpack(payload, "7B", 14)
                spacing = unpack(payload, "7b", 21)
                size, flags = unpack(payload, "iI", 42)
                if not 0 < size <= 409600:
                    raise ValueError("Invalid character font size.")
                shape = element("CharShape", basesize=size, italic=flags & 1,
                                bold=(flags >> 1) & 1, underline="single" if flags & 12 else "none",
                                text_color=color(unpack(payload, "I", 52)[0]))
                for tag, values in (("FontFace", fonts), ("LetterWidthExpansion", widths), ("LetterSpacing", spacing)):
                    shape.append(element(tag, **dict(zip(LANGUAGES, values))))
                if flags & ((1 << 15) | (1 << 16) | (7 << 18)):
                    self.warnings.add("Superscript, subscript, and strikeout styling are not preserved.")
                info.append(shape)
            elif item.tag == 22:
                count = unpack(payload, "I", 4)[0]
                if count > 10000 or len(payload) < 8 + count * 8:
                    raise ValueError("Invalid tab definition count.")
                tabs = ET.SubElement(info, "TabDef")
                array = ET.SubElement(tabs, "Array")
                for index in range(count):
                    position, kind, fill = unpack(payload, "IBB", 8 + index * 8)
                    array.append(element("Tab", pos=position,
                                         kind=("left", "right", "center", "decimal")[min(kind, 3)],
                                         fill_type=fill))
            elif item.tag == 25:
                flags, left, right, indent, before, after, spacing = unpack(payload, "I6i")
                spacing_type = flags & 3
                if self.version >= (5, 0, 2, 5):
                    spacing_flags, spacing = unpack(payload, "II", 46)
                    spacing_type = spacing_flags & 31
                shape = element("ParaShape", align=("both", "left", "right", "center", "distribute", "distribute")
                                [min((flags >> 2) & 7, 5)], doubled_margin_left=left,
                                doubled_margin_right=right, indent=indent, doubled_margin_top=before,
                                doubled_margin_bottom=after, linespacing=spacing,
                                linespacing_type="ratio" if spacing_type == 0 else "other",
                                tabdef_id=unpack(payload, "H", 28)[0],
                                protect_single_line=(flags >> 16) & 1, with_next_paragraph=(flags >> 17) & 1,
                                protect=(flags >> 18) & 1, start_new_page=(flags >> 19) & 1,
                                head_shape="none" if not flags & (3 << 23) else "numbering")
                info.append(shape)
        mappings = info.find("IdMappings")
        if mappings is None:
            raise ValueError("HWP font mappings are missing.")
        if sum(int(mappings.get(language + "-fonts")) for language in LANGUAGES) != len(info.findall("FaceName")):
            raise ValueError("HWP font mapping count mismatch.")
        for shape in info.findall("CharShape"):
            for language in LANGUAGES:
                if int(shape.find("FontFace").get(language)) >= int(mappings.get(language + "-fonts")):
                    raise ValueError("Invalid HWP font reference.")

    def border_fill(self, data):
        fill = element("BorderFill")
        widths = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1, 1.5, 2, 3, 4, 5)
        strokes = {0: "none", 1: "solid", 2: "dashed", 3: "dotted", 8: "double-2"}
        for index, side in enumerate(SIDES):
            stroke, width, rgb = unpack(data, "BBI", 2 + index * 6)
            fill.append(element("Border", attribute_name=side, stroke_type=strokes.get(stroke, "other"),
                                width=str(widths[min(width, 15)]) + "mm", color=color(rgb)))
        flags = unpack(data, "I", 32)[0]
        if flags & 1:
            background = unpack(data, "I", 36)[0]
            if background != 0xFFFFFFFF:
                fill.append(element("FillColorPattern", background_color=color(background)))
        if flags & 6:
            self.warnings.add("Gradient and image background fills are not preserved.")
        return fill

    def paragraph(self, node):
        data = node.record.data
        character_count, control_mask, shape_id, style_id, split = unpack(data, "IIHBB")
        if shape_id >= len(self.root.findall("DocInfo/ParaShape")):
            raise ValueError("Invalid paragraph style reference.")
        paragraph = element("Paragraph", parashape_id=shape_id, new_page=int(bool(split & 4)))
        payloads = {}
        controls = []
        for child in node.children:
            tag = child.record.tag
            if tag == 71:
                controls.append(child)
            elif tag in (67, 68, 69):
                if tag in payloads:
                    raise ValueError("Duplicate paragraph data record.")
                payloads[tag] = child.record.data
            elif tag != 70:
                raise ValueError(f"Unsupported paragraph record {tag}.")
            else:
                self.warnings.add("Paragraph range annotations are not preserved.")
        text = payloads.get(67, b"")
        if len(text) % 2:
            raise ValueError("Truncated UTF-16 paragraph record.")
        expected = character_count & 0x7FFFFFFF
        if len(text) // 2 != expected and not (not text and expected == 1):
            raise ValueError("Paragraph character count mismatch.")
        shapes_data = payloads.get(68, b"")
        if not shapes_data or len(shapes_data) % 8:
            raise ValueError("Invalid paragraph character-style positions.")
        shapes = list(struct.iter_unpack("<II", shapes_data))
        positions = [position for position, identifier in shapes]
        if positions[0] != 0 or any(left >= right for left, right in zip(positions, positions[1:])):
            raise ValueError("Unordered paragraph character-style positions.")
        shape_count = len(self.root.findall("DocInfo/CharShape"))
        if any(identifier >= shape_count or position > expected for position, identifier in shapes):
            raise ValueError("Invalid paragraph character-style reference.")
        line_data = payloads.get(69, b"")
        if len(line_data) % 36:
            raise ValueError("Truncated paragraph line geometry.")
        lines = []
        line_positions = []
        names = ("chpos", "y", "height", "height-text", "baseline", "space-below", "x", "width", "flags")
        for values in struct.iter_unpack("<I7iI", line_data):
            line = element("LineSeg", **dict(zip(names, values)))
            lines.append(line)
            line_positions.append(values[0])
            paragraph.append(line)
        if not lines:
            lines = [ET.SubElement(paragraph, "LineSeg")]
            line_positions = [0]
        if line_positions[0] != 0 or any(left > right for left, right in zip(line_positions, line_positions[1:])):
            raise ValueError("Unordered paragraph line positions.")
        extended = {1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23}
        wide = extended | {4, 5, 6, 7, 8, 9, 19, 20}
        position = 0
        while position < len(text) // 2:
            code = unpack(text, "H", position * 2)[0]
            shape_index = bisect_right(positions, position) - 1
            line_index = bisect_right(line_positions, position) - 1
            line = lines[line_index]
            if code < 32:
                size = 8 if code in wide else 1
                if position + size > len(text) // 2:
                    raise ValueError("Truncated paragraph control.")
                if size == 8 and unpack(text, "H", (position + 7) * 2)[0] != code:
                    raise ValueError("Mismatched paragraph control terminator.")
                if code in extended:
                    identifier = text[position * 2 + 2:position * 2 + 6]
                    match = next((child for child in controls if child.record.data[:4] == identifier), None)
                    if match is None:
                        raise ValueError("Paragraph control has no matching record.")
                    controls.remove(match)
                    rendered = self.control(match)
                    if rendered is not None:
                        line.append(rendered)
                elif code in (9, 10, 13, 31):
                    line.append(element("ControlChar", name={9: "TAB", 10: "LINE_BREAK", 13: "PARAGRAPH_BREAK",
                                                             31: "FIXWIDTH_SPACE"}[code]))
                elif code in (24, 30):
                    item = element("Text", charshape_id=shapes[shape_index][1])
                    item.text = "\u00ad" if code == 24 else "\u00a0"
                    line.append(item)
                elif code == 4:
                    line.append(element("FieldEnd"))
                else:
                    raise ValueError(f"Unsupported text control {code}.")
                position += size
                continue
            end = position + 1
            boundary = min(positions[shape_index + 1] if shape_index + 1 < len(positions) else len(text) // 2,
                           line_positions[line_index + 1] if line_index + 1 < len(lines) else len(text) // 2)
            while end < boundary and unpack(text, "H", end * 2)[0] >= 32:
                end += 1
            decoded = self.decode(text[position * 2:end * 2])
            groups = []
            for char in decoded:
                language = "en" if ord(char) < 128 else "ko"
                if groups and groups[-1][0] == language:
                    groups[-1][1].append(char)
                else:
                    groups.append((language, [char]))
            for language, characters in groups:
                item = element("Text", charshape_id=shapes[shape_index][1], lang=language)
                item.text = "".join(characters)
                line.append(item)
            position = end
        if controls:
            raise ValueError("Unreferenced paragraph control records.")
        self.resolve_fields(paragraph)
        return paragraph

    def resolve_fields(self, paragraph):
        """Keep hyperlink labels and character styles together across stored lines."""
        fields = []
        for line in paragraph.findall("LineSeg"):
            items = list(line)
            line[:] = []
            wrapper = None
            for item in items:
                if item.tag == "FieldStart":
                    if item.get("kind") == "hyperlink" and any(f.get("kind") == "hyperlink" for f in fields):
                        raise ValueError("Nested hyperlinks are unsupported.")
                    fields.append(item)
                    wrapper = None
                elif item.tag == "FieldEnd":
                    if fields:
                        fields.pop()
                    else:
                        self.warnings.add("A field ends outside this paragraph; its displayed text is retained.")
                    wrapper = None
                else:
                    link = next((field for field in fields if field.get("kind") == "hyperlink"), None)
                    if link is None:
                        line.append(item)
                    else:
                        if item.tag not in ("Text", "ControlChar", "GShapeObjectControl"):
                            raise ValueError("Unsupported content inside a hyperlink.")
                        if wrapper is None:
                            wrapper = element("FieldHyperLink", target=link.get("target"))
                            line.append(wrapper)
                        wrapper.append(item)
        if any(field.get("kind") == "hyperlink" for field in fields):
            raise ValueError("Hyperlinks spanning paragraphs or missing an end marker are unsupported.")

    def control(self, node):
        data = node.record.data
        identifier = unpack(data, "4s")[0][::-1]
        if identifier == b"secd":
            for child in node.children:
                if child.record.tag == 73:
                    values = unpack(child.record.data, "10I")
                    names = ("width", "height", "left-offset", "right-offset", "top-offset", "bottom-offset",
                             "header-offset", "footer-offset", "gutter", "flags")
                    page = element("PageDef", **dict(zip(names, values)))
                    if values[9] & 1:
                        page.set("width", str(values[1]))
                        page.set("height", str(values[0]))
                    if values[8]:
                        self.warnings.add("Binding gutter margins are not preserved.")
                    if self.section.find("PageDef") is not None:
                        raise ValueError("Multiple page definitions are unsupported.")
                    self.section.append(page)
                elif child.record.tag == 87:
                    self.warnings.add("Section control metadata is not preserved.")
                elif child.record.tag not in (74, 75):
                    raise ValueError(f"Unsupported section record {child.record.tag}.")
            return None
        if identifier == b"cold":
            return element("ColumnsDef", count=(unpack(data, "H", 4)[0] >> 2) & 255)
        if identifier == b"tbl ":
            return self.table(node)
        if identifier == b"gso ":
            return self.drawing(node)
        if identifier in (b"head", b"foot"):
            flags = unpack(data, "I", 4)[0]
            if flags & 3 == 3:
                raise ValueError("Invalid header/footer page selection.")
            header = element("Header" if identifier == b"head" else "Footer",
                             apply=("both", "even", "odd")[flags & 3])
            expected_count = None
            for child in node.children:
                if child.record.tag == 66:
                    header.append(self.paragraph(child))
                elif child.record.tag == 72:
                    if expected_count is not None:
                        raise ValueError("Multiple header/footer paragraph lists are unsupported.")
                    expected_count = unpack(child.record.data, "I")[0] & 0x7FFFFFFF
                    for nested in child.children:
                        if nested.record.tag != 66:
                            raise ValueError("Unsupported nested header/footer content.")
                        header.append(self.paragraph(nested))
                elif child.record.tag == 87:
                    self.warnings.add("Header/footer control metadata is not preserved.")
                else:
                    raise ValueError(f"Unsupported header/footer record {child.record.tag}.")
            if expected_count is not None and expected_count != len(header):
                raise ValueError("Header/footer paragraph count mismatch.")
            if any(item.tag in ("Header", "Footer") for p in header for item in p.iter()):
                raise ValueError("Nested headers and footers are unsupported.")
            return header
        if identifier == b"%hlk":
            if any(child.record.tag != 87 for child in node.children):
                raise ValueError("Unsupported nested hyperlink content.")
            length = unpack(data, "H", 9)[0]
            command = self.decode(unpack(data, f"{length * 2}s", 11)[0])
            unpack(data, "I", 11 + length * 2)  # Field instance ID.
            # HWP separates command components with semicolons; escaped semicolons
            # and backslashes belong to the address itself.
            target = []
            index = 0
            while index < len(command):
                char = command[index]
                if char == ";":
                    break
                if char == "\\" and index + 1 < len(command) and command[index + 1] in "\\;:":
                    index += 1
                    char = command[index]
                target.append(char)
                index += 1
            return element("FieldStart", kind="hyperlink", target="".join(target))
        if identifier == b"bokm" or identifier.startswith(b"%"):
            if any(child.record.tag != 87 for child in node.children):
                raise ValueError("Unsupported nested field or bookmark content.")
            self.warnings.add("Bookmarks and field behavior are not preserved; displayed labels remain.")
            return element("FieldStart", kind="display") if identifier.startswith(b"%") else None
        raise ValueError(f"Unsupported HWP control {identifier!r}; no output published.")

    def table(self, node):
        flags = unpack(node.record.data, "I", 4)[0]
        width, height = unpack(node.record.data, "II", 16)
        if not width or not height:
            raise ValueError("Invalid table dimensions.")
        table = element("TableControl", width=width, height=height, inline=flags & 1)
        definitions = [child for child in node.children if child.record.tag == 77]
        if len(definitions) != 1:
            raise ValueError("Missing or duplicate table definition.")
        data = definitions[0].record.data
        rows, columns = unpack(data, "HH", 4)
        if not rows or not columns or rows * columns > 10000:
            raise ValueError("Table dimensions exceed prototype limits.")
        body = element("TableBody", rows=rows, cols=columns,
                       **dict(zip(("padding-" + side for side in SIDES), unpack(data, "4H", 10))))
        table.append(body)
        row_nodes = [ET.SubElement(body, "TableRow") for index in range(rows)]
        occupied = set()
        current = None
        remaining = 0
        previous_address = (-1, -1)
        for child in node.children:
            if child.record.tag == 77:
                continue
            if child.record.tag == 72:
                if remaining:
                    raise ValueError("Table cell paragraph count mismatch.")
                data = child.record.data
                if len(data) < 34:
                    raise ValueError("Table captions are unsupported.")
                remaining, flags = unpack(data, "II")
                remaining &= 0x7FFFFFFF
                col, row, colspan, rowspan, width, height = unpack(data, "4H2I", 8)
                if not colspan or not rowspan or row + rowspan > rows or col + colspan > columns:
                    raise ValueError("Invalid table cell span.")
                if (row, col) <= previous_address:
                    raise ValueError("Non-row-major table cells are unsupported.")
                previous_address = (row, col)
                for row_index in range(row, row + rowspan):
                    for col_index in range(col, col + colspan):
                        address = (row_index, col_index)
                        if address in occupied:
                            raise ValueError("Overlapping table cells.")
                        occupied.add(address)
                border = unpack(data, "H", 32)[0]
                if border > len(self.root.findall("DocInfo/BorderFill")):
                    raise ValueError("Invalid table border reference.")
                current = element("TableCell", row=row, col=col, colspan=colspan, rowspan=rowspan,
                                  width=width, height=height, borderfill_id=border,
                                  valign=("top", "middle", "bottom", "top")[(flags >> 5) & 3],
                                  **dict(zip(("padding-" + side for side in SIDES), unpack(data, "4H", 24))))
                row_nodes[row].append(current)
            elif child.record.tag == 66:
                if current is None or remaining < 1:
                    raise ValueError("Table paragraph outside a declared cell.")
                current.append(self.paragraph(child))
                remaining -= 1
            else:
                raise ValueError(f"Unsupported table record {child.record.tag}.")
        if remaining or len(occupied) != rows * columns:
            raise ValueError("Incomplete table cells or paragraphs.")
        return table

    def drawing(self, node):
        width, height = unpack(node.record.data, "II", 16)
        drawing = element("GShapeObjectControl", width=width, height=height)
        pending = list(reversed(node.children))
        while pending:
            child = pending.pop()
            tag = child.record.tag
            if tag == 66:
                textbox = ET.SubElement(drawing, "TextboxParagraphList")
                textbox.append(self.paragraph(child))
            elif tag == 85:
                identifier = unpack(child.record.data, "H", 71)[0]
                picture = ET.SubElement(drawing, "ShapePicture")
                picture.append(element("PictureInfo", bindata_id=identifier))
            elif tag in (76, 72, 79):
                pending.extend(reversed(child.children))
            else:
                raise ValueError(f"Unsupported drawing record {tag}.")
        if not len(drawing):
            raise ValueError("Vector-only drawings are unsupported.")
        if len(list(drawing.iter("ShapePicture"))) > 1:
            raise ValueError("Grouped pictures are unsupported.")
        return drawing

    def stream(self, document, name, compressed):
        size = document.get_size(name)
        stream_limit = self.limits.max_stream_mb * MIB
        remaining = self.limits.max_expanded_mb * MIB - self.stream_bytes
        if size > stream_limit:
            raise ValueError("HWP stream exceeds the size limit (--max-stream-mb).")
        if remaining < 0 or (not compressed and size > remaining):
            raise ValueError("Total expanded HWP data exceeds the size limit (--max-expanded-mb).")
        with document.openstream(name) as stream:
            raw = stream.read(size + 1)
        if len(raw) != size:
            raise ValueError("HWP stream length does not match its declared size.")
        try:
            data = decompress_stream(raw, min(stream_limit, remaining)) if compressed else raw
        except ValueError as error:
            if "size limit" in str(error):
                budget = "--max-expanded-mb" if remaining < stream_limit else "--max-stream-mb"
                raise ValueError(f"Expanded HWP stream exceeds the size limit ({budget}).") from error
            raise
        self.stream_bytes += len(data)
        return data

    def parse(self, source):
        source = Path(source)
        if source.suffix.lower() != ".hwp":
            raise ValueError("This prototype supports binary .hwp files only, not HWPX.")
        self.limits.check_input(source)
        with olefile.OleFileIO(source) as document:
            header = self.stream(document, "FileHeader", False)
            if len(header) < 40 or header[:32] != b"HWP Document File".ljust(32, b"\0"):
                raise ValueError("Not a supported HWP document.")
            self.version = tuple(reversed(header[32:36]))
            if self.version[0] != 5:
                raise ValueError("Only HWP version 5 files are supported.")
            flags = unpack(header, "I", 36)[0]
            if flags & ((1 << 1) | (1 << 2) | (1 << 4) | (1 << 8) | (1 << 10)):
                raise ValueError("Password, distribution, or DRM protected files are not converted.")
            compressed = bool(flags & 1)
            sections = [name for name in document.listdir() if name[0] == "BodyText"]
            if not sections or any(len(name) != 2 or not re.fullmatch(r"Section(0|[1-9][0-9]*)", name[1])
                                   for name in sections):
                raise ValueError("Missing or invalid HWP section stream names.")
            sections.sort(key=lambda name: int(name[1][7:]))
            if [name[1] for name in sections] != [f"Section{index}" for index in range(len(sections))]:
                raise ValueError("HWP section streams must be consecutive, starting at Section0.")
            self.doc_info(self.stream(document, "DocInfo", compressed))
            if self.section_count is not None and self.section_count != len(sections):
                raise ValueError("HWP section count does not match the document properties.")
            body = ET.SubElement(self.root, "BodyText")
            for name in sections:
                self.section = ET.SubElement(body, "SectionDef", {"source-stream": name[1]})
                columns = ET.SubElement(self.section, "ColumnSet")
                for node in record_tree(self.stream(document, name, compressed),
                                        self.limits.max_stream_mb * MIB, self.limits.max_records):
                    if node.record.tag != 66:
                        raise ValueError(f"Unsupported top-level body record {node.record.tag}.")
                    columns.append(self.paragraph(node))
            used_assets = {int(picture.get("bindata-id")) for picture in self.root.iter("PictureInfo")}
            for identifier, extension, mode in self.assets:
                if identifier not in used_assets:
                    continue
                if mode == 3:
                    raise ValueError("Invalid embedded asset compression mode.")
                data = self.stream(document, ["BinData", f"BIN{identifier:04X}.{extension}"],
                                   compressed if mode == 0 else mode == 1)
                asset = element("BinDataEmbedding", storage_id=f"BIN{identifier:04X}", ext=extension)
                if self.asset_directory is None:
                    asset.set("inline", "true")
                    asset.text = base64.b64encode(data).decode("ascii")
                else:
                    filename = f"BIN{identifier:04X}.{extension}"
                    with (self.asset_directory / filename).open("xb") as extracted:
                        extracted.write(data)
                    asset.set("file", filename)
                self.root.append(asset)
                used_assets.remove(identifier)
            if used_assets:
                raise ValueError("Referenced embedded images are unavailable.")
        replacements = self.recovery.replacements if self.recovery else []
        if replacements:
            self.warnings.add(f"Replaced {len(replacements)} invalid UTF-16 sequence(s) with U+FFFD; review text_replacements in the report.")
        self.root.set("text-replacements", json.dumps(replacements))
        self.root.set("parser-warnings", json.dumps(sorted(self.warnings)))
        return self.root, sorted(self.warnings)
