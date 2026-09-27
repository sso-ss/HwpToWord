"""Experimental HWP 3.x reader based on Hancom revision 1.2, Part I.

Uses only the Python standard library. hchar is NOT Unicode or general Johab:
only ASCII and the modern Hangul syllable bit layout are decoded here.
본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
"""

import json
from pathlib import Path
import struct
import xml.etree.ElementTree as ET
import zlib


SIGNATURE = b"HWP Document File V3.00 \x1a\x01\x02\x03\x04\x05"
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_STREAM_BYTES = 64 * 1024 * 1024
MAX_PARAGRAPHS = 50000
MAX_CHARACTERS = 2000000
MAX_DEPTH = 32
LANGUAGES = ("ko", "en", "cn", "jp", "other", "symbol", "user")
SIDES = ("left", "right", "top", "bottom")
PALETTE = ("000000", "0000ff", "00ff00", "00ffff", "ff0000", "ff00ff", "ffff00", "ffffff")
VOWELS = tuple(range(3, 8)) + tuple(range(10, 16)) + tuple(range(18, 24)) + tuple(range(26, 30))
FINALS = tuple(range(1, 18)) + tuple(range(19, 30))


def element(tag, **attributes):
    return ET.Element(tag, {key.replace("_", "-"): str(value) for key, value in attributes.items()})


def number(data, offset=0, kind="H"):
    if offset + struct.calcsize("<" + kind) > len(data):
        raise ValueError("Truncated HWP 3 field.")
    return struct.unpack_from("<" + kind, data, offset)[0]


class Reader:
    def __init__(self, data):
        self.data = data
        self.offset = 0

    @property
    def remaining(self):
        return len(self.data) - self.offset

    def take(self, size):
        if size < 0 or size > self.remaining:
            raise ValueError(f"Truncated HWP 3 data at stream byte {self.offset}.")
        start = self.offset
        self.offset += size
        return self.data[start:self.offset]

    def word(self):
        return number(self.take(2))


def validate_header(header):
    if len(header) < 158 or header[:30] != SIGNATURE:
        raise ValueError("Invalid or truncated HWP 3 header.")
    info = header[30:158]
    if number(info, 24, "I") or number(info, 96):
        raise ValueError("Password or editing-protected HWP 3 files are not converted.")
    if info[125] not in (0, 1):
        raise ValueError("Unsupported HWP 3 subrevision.")
    return info


def decompress_stream(data, limit=MAX_STREAM_BYTES):
    """Return the bounded body and the separate, uncompressed preview area.

    Accept gzip framing or raw DEFLATE, optionally followed by a gzip CRC/size
    footer. Unlike HWP 5 streams, a valid HWP 3 file can have preview blocks
    after the end of compression. Those bytes must still be parsed/validated.
    """
    gzip = data.startswith(b"\x1f\x8b")
    decoder = zlib.decompressobj(31 if gzip else -15)
    try:
        body = decoder.decompress(data, limit + 1)
    except zlib.error as error:
        raise ValueError("Invalid compressed HWP 3 stream or checksum.") from error
    if len(body) > limit or decoder.unconsumed_tail:
        raise ValueError("Expanded HWP 3 stream exceeds the size limit.")
    if not decoder.eof:
        raise ValueError("Incomplete compressed HWP 3 stream.")
    tail = decoder.unused_data
    if not gzip and tail:
        footer = struct.pack("<II", zlib.crc32(body), len(body))
        if tail.startswith(footer):
            tail = tail[8:]
        elif len(tail) < 4 or number(tail, 0, "I") not in (0x80000000, 0x80000001, 0x80000002):
            raise ValueError("Invalid HWP 3 compression checksum or trailing data.")
    return body, tail


def decode_hchar(code):
    if 0x20 <= code <= 0x7e:
        return chr(code)
    initial, vowel, final = (code >> 10) & 31, (code >> 5) & 31, code & 31
    if code & 0x8000 and 2 <= initial <= 20 and vowel in VOWELS and final in FINALS:
        return chr(0xac00 + ((initial - 2) * 21 + VOWELS.index(vowel)) * 28 + FINALS.index(final))
    raise ValueError(f"Unsupported HWP 3 hchar 0x{code:04X}; legacy Hanja, symbols, "
                     "and old Hangul mappings are not implemented. --recover-text does not replace these codes.")


class Hwp3Parser:
    def __init__(self):
        self.root = element("HwpDoc", parser="native-hwp3")
        self.info = ET.SubElement(self.root, "DocInfo")
        self.warnings = {"HWP 3 support is experimental; real HWP 3 document compatibility and visual layout are unverified."}
        self.font_counts = []
        self.char_shapes = {}
        self.para_shapes = {}
        self.paragraph_count = 0
        self.character_count = 0
        self.cell_count = 0
        self.border_count = 0

    def char_shape(self, data):
        if data in self.char_shapes:
            return self.char_shapes[data]
        if len(self.char_shapes) >= 50000:
            raise ValueError("HWP 3 character style count exceeds the limit.")
        size = number(data)
        if not size or data[24] > 7:
            raise ValueError("Invalid HWP 3 character size or color.")
        if any(index >= count for index, count in zip(data[2:9], self.font_counts)):
            raise ValueError("Invalid HWP 3 font reference.")
        flags = data[26]
        shape = element("CharShape", basesize=size * 4, bold=(flags >> 1) & 1,
                        italic=flags & 1, underline="single" if flags & 4 else "none",
                        text_color="#" + PALETTE[data[24]], superscript=(flags >> 5) & 1,
                        subscript=(flags >> 6) & 1)
        if flags & 0x60 == 0x60:
            raise ValueError("Conflicting HWP 3 superscript and subscript flags.")
        for tag, values in (("FontFace", data[2:9]), ("LetterWidthExpansion", data[9:16]),
                            ("LetterSpacing", struct.unpack("7b", data[16:23]))):
            shape.append(element(tag, **dict(zip(LANGUAGES, values))))
        if flags & 0x18 or data[25]:
            self.warnings.add("HWP 3 character outlines, shadows, and shading are not preserved.")
        index = len(self.char_shapes)
        self.info.append(shape)
        self.char_shapes[data] = index
        return index

    def para_shape(self, data, flags):
        key = (data, flags & 0x78)
        if key in self.para_shapes:
            return self.para_shapes[key]
        if data[172] > 1:
            raise ValueError("HWP 3 multiple columns are unsupported.")
        if data[11] > 7:
            raise ValueError("Invalid HWP 3 paragraph alignment.")
        # Part I specifies DOS menu indices but does not give their meanings.
        # Do not silently invent a Unicode-era alignment enum for these values.
        self.warnings.add("HWP 3 paragraph alignment is approximated as left-aligned; DOS alignment mapping is unverified.")
        if any(data[180:183]):
            self.warnings.add("HWP 3 paragraph borders and shading are not preserved.")
        spacing = number(data, 6)
        index = len(self.para_shapes)
        shape = element("ParaShape", align="left", doubled_margin_left=number(data) * 8,
                        doubled_margin_right=number(data, 2) * 8, indent=number(data, 4, "h") * 4,
                        doubled_margin_top=number(data, 183) * 8, doubled_margin_bottom=number(data, 8) * 8,
                        linespacing=(spacing & 0x7fff) * 4 if spacing & 0x8000 else spacing,
                        linespacing_type="fixed" if spacing & 0x8000 else "ratio",
                        tabdef_id=index, with_next_paragraph=int(bool(flags & 0x18)),
                        protect_single_line=int(bool(flags & 0x40)))
        tabs = element("TabDef")
        array = ET.SubElement(tabs, "Array")
        for offset in range(12, 172, 4):
            kind, leader, position = struct.unpack_from("<BBH", data, offset)
            if not position:
                continue
            if kind > 3:
                raise ValueError("Invalid HWP 3 tab type.")
            array.append(element("Tab", pos=position * 4, kind=("left", "right", "center", "decimal")[kind],
                                 fill_type=3 if leader else 0))
        self.info.extend((shape, tabs))
        self.para_shapes[key] = index
        return index

    def paragraph_list(self, reader, depth=0):
        if depth > MAX_DEPTH:
            raise ValueError("HWP 3 paragraph nesting exceeds the limit.")
        paragraphs = []
        previous_shape = None
        while True:
            header = reader.take(43)
            count, lines = number(header, 1), number(header, 3)
            # The zero-length paragraph is the list sentinel, not an empty CR
            # paragraph. Its representative character shape is not rendered.
            if not count:
                if lines or header[5]:
                    raise ValueError("Invalid HWP 3 paragraph-list terminator.")
                return paragraphs
            self.paragraph_count += 1
            self.character_count += count
            if self.paragraph_count > MAX_PARAGRAPHS or self.character_count > MAX_CHARACTERS:
                raise ValueError("HWP 3 paragraph or character count exceeds the limit.")
            flags = header[6]
            if flags & 5:
                raise ValueError("HWP 3 column or section breaks are unsupported.")
            if not header[0]:
                previous_shape = reader.take(187)
            if previous_shape is None:
                raise ValueError("Missing first HWP 3 paragraph shape.")
            shape_id = self.para_shape(previous_shape, flags)
            paragraph = element("Paragraph", parashape_id=shape_id, new_page=int(bool(flags & 2)))
            if lines > count:
                raise ValueError("Invalid HWP 3 line count.")
            descriptors = [reader.take(14) for _ in range(lines)]
            positions = [number(item) for item in descriptors]
            if positions and (positions[0] != 0 or positions[-1] >= count
                              or any(a >= b for a, b in zip(positions, positions[1:]))):
                raise ValueError("Invalid HWP 3 line positions.")
            if header[5]:
                shapes = []
                for _ in range(count):
                    if reader.take(1) != b"\x01":
                        shape = self.char_shape(reader.take(31))
                    elif not shapes:
                        raise ValueError("Missing first HWP 3 character shape.")
                    shapes.append(shape)
            else:
                shapes = [self.char_shape(header[12:43])] * count
            line = ET.SubElement(paragraph, "LineSeg")
            # Do not give the shared writer invented HWP 5 y-coordinates/heights.
            # HWP 3 has page-boundary flags but no equivalent y-position array.
            if any(number(item, 12) & 0x8001 == 0x8001 for item in descriptors):
                self.warnings.add("Stored HWP 3 line page boundaries are not reconstructed; explicit paragraph page breaks are retained.")
            position = 0
            pending = []
            pending_shape = None
            pending_language = None

            def flush():
                if pending:
                    text = element("Text", charshape_id=pending_shape, lang=pending_language)
                    text.text = "".join(pending)
                    line.append(text)
                    pending.clear()

            while position < count:
                code = reader.word()
                width = 1
                if code >= 32:
                    language = "en" if code < 128 else "ko"
                    if pending_shape != shapes[position] or pending_language != language:
                        flush()
                        pending_shape = shapes[position]
                        pending_language = language
                    pending.append(decode_hchar(code))
                else:
                    flush()
                    if code == 13:
                        if position != count - 1:
                            raise ValueError("Early HWP 3 paragraph terminator.")
                        line.append(element("ControlChar", name="PARAGRAPH_BREAK"))
                    elif code == 9:
                        marker = reader.take(6)
                        if number(marker, 4) != 9:
                            raise ValueError("Invalid HWP 3 tab marker.")
                        line.append(element("ControlChar", name="TAB"))
                        width = 4
                    elif code in (30, 31):
                        if reader.word() != code:
                            raise ValueError("Invalid HWP 3 space marker.")
                        if code == 30:
                            text = element("Text", charshape_id=shapes[position])
                            text.text = "\u00a0"
                            line.append(text)
                        else:
                            line.append(element("ControlChar", name="FIXWIDTH_SPACE"))
                        width = 2
                    elif code == 10:
                        marker = reader.take(6)
                        if number(marker, 4) != 10:
                            raise ValueError("Invalid HWP 3 table marker.")
                        line.append(self.table(reader, depth + 1))
                        width = 4
                    else:
                        raise ValueError(f"Unsupported HWP 3 control {code}; no output was created.")
                position += width
                if position > count or (position == count and code != 13):
                    raise ValueError("HWP 3 paragraph length or terminator mismatch.")
            paragraphs.append(paragraph)

    def table(self, reader, depth):
        data = reader.take(84)
        if number(data, 16) != 10:
            raise ValueError("Invalid HWP 3 table information marker.")
        kind, count = number(data, 78), number(data, 80)
        if kind not in (0, 1) or number(data, 14) & 16:
            raise ValueError("HWP 3 equations, buttons, and hypertext boxes are unsupported.")
        self.cell_count += count
        if not 0 < count <= 10000 or self.cell_count > 50000 or (kind == 1 and count != 1):
            raise ValueError("Invalid or excessive HWP 3 cell count.")
        if depth > MAX_DEPTH:
            raise ValueError("HWP 3 paragraph nesting exceeds the limit.")
        cells = [reader.take(27) for _ in range(count)]
        if any(cell[25] & 0x33 for cell in cells):
            raise ValueError("HWP 3 diagonal borders and diagonal cell merges are unsupported.")
        contents = [self.paragraph_list(reader, depth) for _ in cells]
        caption = self.paragraph_list(reader, depth)
        if caption:
            raise ValueError("HWP 3 table/textbox captions are unsupported.")
        if kind == 1:
            if any(p.findall(".//TableControl") for p in contents[0]):
                raise ValueError("Tables inside HWP 3 text boxes are unsupported.")
            drawing = element("GShapeObjectControl")
            body = ET.SubElement(drawing, "TextboxParagraphList")
            body.extend(contents[0])
            return drawing
        rectangles = [tuple(number(cell, offset) for offset in (4, 6, 8, 10)) for cell in cells]
        if any(not width or not height for x, y, width, height in rectangles):
            raise ValueError("Invalid HWP 3 cell size.")
        xs = sorted({edge for x, y, w, h in rectangles for edge in (x, x + w)})
        ys = sorted({edge for x, y, w, h in rectangles for edge in (y, y + h)})
        rows, cols = len(ys) - 1, len(xs) - 1
        if rows * cols > 10000:
            raise ValueError("HWP 3 table grid exceeds the size limit.")
        x_index = {value: index for index, value in enumerate(xs)}
        y_index = {value: index for index, value in enumerate(ys)}
        occupied = set()
        table = element("TableControl", width=(xs[-1] - xs[0]) * 4, inline=int(data[8] == 0))
        body = element("TableBody", rows=rows, cols=cols)
        table.append(body)
        row_nodes = [ET.SubElement(body, "TableRow") for _ in range(rows)]
        for cell, paragraphs, (x, y, width, height) in zip(cells, contents, rectangles):
            row, col = y_index[y], x_index[x]
            rowspan, colspan = y_index[y + height] - row, x_index[x + width] - col
            covered = {(r, c) for r in range(row, row + rowspan) for c in range(col, col + colspan)}
            if covered & occupied:
                raise ValueError("Overlapping HWP 3 table cells.")
            occupied.update(covered)
            node = element("TableCell", row=row, col=col, rowspan=rowspan, colspan=colspan,
                           width=width * 4, height=max(height, number(cell, 12), number(cell, 14)) * 4,
                           valign="middle" if cell[19] else "top", borderfill_id=self.border(cell))
            for index, side in enumerate(SIDES):
                node.set("padding-" + side, str(number(data, 34 + index * 2) * 4))
            node.extend(paragraphs)
            row_nodes[row].append(node)
        if len(occupied) != rows * cols:
            raise ValueError("Incomplete HWP 3 table grid.")
        for row in row_nodes:
            row[:] = sorted(row, key=lambda item: int(item.get("col")))
        return table

    def border(self, data):
        fill = element("BorderFill")
        for side, stroke in zip(SIDES, data[20:24]):
            if stroke > 4:
                raise ValueError("Unsupported HWP 3 cell border.")
            fill.append(element("Border", attribute_name=side,
                                stroke_type=("none", "solid", "solid", "dotted", "double-2")[stroke],
                                width="0.4mm" if stroke == 2 else "0.12mm", color="#000000"))
        if data[24]:
            self.warnings.add("HWP 3 cell shading is not preserved; its color encoding is unverified.")
        self.info.append(fill)
        self.border_count += 1
        return self.border_count

    def extra_blocks(self, reader, preview=False):
        if not reader.remaining:
            return
        while reader.remaining:
            identifier = number(reader.take(4), kind="I")
            size = number(reader.take(4), kind="I")
            if identifier == (0x80000000 if preview else 0):
                if not preview and size:
                    raise ValueError("Invalid HWP 3 extra-block terminator.")
                # Preview terminator size describes the uncompressed area; it
                # is not a payload size (Part I, section 3.9).
                return
            if not preview and identifier >= 0x80000000:
                reader.offset -= 8
                return
            allowed = (0x80000001, 0x80000002) if preview else (1, 2, 3, 4, 6, 0x100, 0x101)
            if identifier not in allowed:
                raise ValueError(f"Unsupported HWP 3 additional block 0x{identifier:X}.")
            reader.take(size)
            if not preview:
                if identifier == 6:
                    raise ValueError("HWP 3 background images are unsupported.")
                self.warnings.add("HWP 3 additional assets, links, and presentation metadata are not preserved.")
        raise ValueError("Missing HWP 3 additional-block terminator.")

    def parse(self, source):
        source = Path(source)
        if source.suffix.lower() != ".hwp":
            raise ValueError("This prototype supports binary .hwp files only, not HWPX.")
        with source.open("rb") as stream:
            raw = stream.read(MAX_FILE_BYTES + 1)
        if len(raw) > MAX_FILE_BYTES:
            raise ValueError("The prototype accepts files up to 50 MB.")
        info = validate_header(raw[:158])
        reader = Reader(raw)
        reader.take(158)
        summary = reader.take(1008)
        if any(summary):
            self.warnings.add("HWP 3 document summary metadata is not copied.")
        extra = Reader(reader.take(number(info, 126)))
        while extra.remaining:
            identifier, size = extra.word(), extra.word()
            if identifier not in (1, 2):
                raise ValueError("Unsupported HWP 3 information block.")
            extra.take(size)
            self.warnings.add("HWP 3 bookmark and cross-reference metadata is not preserved.")
        data = reader.take(reader.remaining)
        compressed = bool(info[124])
        data, tail = decompress_stream(data) if compressed else (data, b"")
        reader = Reader(data)
        mappings = ET.SubElement(self.info, "IdMappings")
        for language in LANGUAGES:
            count = reader.word()
            if count > 256:
                raise ValueError("HWP 3 font count exceeds the index range.")
            self.font_counts.append(count)
            mappings.set(language + "-fonts", str(count))
            for _ in range(count):
                try:
                    name = reader.take(40).split(b"\0", 1)[0].decode("johab")
                except UnicodeDecodeError as error:
                    raise ValueError("Invalid HWP 3 Johab font name.") from error
                if any(ord(char) < 32 for char in name):
                    raise ValueError("Invalid control character in HWP 3 font name.")
                self.info.append(element("FaceName", name=name))
        styles = reader.word()
        reader.take(styles * 238)
        if styles:
            self.warnings.add("HWP 3 named styles are flattened to direct formatting.")
        section = ET.SubElement(ET.SubElement(self.root, "BodyText"), "SectionDef")
        width, height = number(info, 8), number(info, 6)
        if info[5] not in (0, 1) or not width or not height:
            raise ValueError("Invalid HWP 3 page dimensions.")
        if info[5]:
            width, height = height, width
        page = element("PageDef", width=width * 4, height=height * 4)
        for name, offset in (("top", 10), ("bottom", 12), ("left", 14), ("right", 16),
                             ("header", 18), ("footer", 20)):
            page.set(name + "-offset", str(number(info, offset) * 4))
        if number(info, 22) or number(info, 120) or info[122]:
            self.warnings.add("HWP 3 binding margins, page borders, and hidden empty lines are not preserved.")
        section.append(page)
        columns = ET.SubElement(section, "ColumnSet")
        columns.extend(self.paragraph_list(reader))
        self.extra_blocks(reader)
        if compressed:
            if reader.remaining:
                raise ValueError("Unexpected data inside the compressed HWP 3 body.")
            reader = Reader(tail)
        self.extra_blocks(reader, preview=True)
        if reader.remaining:
            raise ValueError("Unexpected trailing HWP 3 data.")
        self.root.set("text-replacements", "[]")
        self.root.set("parser-warnings", json.dumps(sorted(self.warnings)))
        return self.root, sorted(self.warnings)
