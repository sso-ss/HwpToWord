# Native HWP Parsers

The HWP-specific parser is implemented in this repository from Hancom's public
HWP 5.0 revision 1.3 and HWP 3.0/HWPML revision 1.2 specifications. It does not import, invoke, or fall back to
`pyhwp`, and no pyhwp implementation was used to author the replacement.

Specification: https://www.hancom.com/support/downloadCenter/hwpOwpml

HWP 3 source: *한글 문서 파일 구조 3.0 / HWPML*, revision 1.2:20141105,
Part I (binary HWP 3.x), not Part II (HWPML). In particular, tables 2–16 cover
framing and styles; tables 39–42 cover tabs/tables/cells; tables 64–65 cover spaces.

## Boundaries

- `hwp3_parser.py`: standard-library-only HWP 3 binary reader, bounded
  decompression, font tables, paragraph/style decoding, and rectangular tables.
  It neither imports `olefile` nor reuses the HWP 5 reader.
- `hwp_parser.py`: bounded DEFLATE decoding, record framing and nesting,
  document styles, paragraph text/control positions, tables, page settings,
  and limited drawing/image records. It emits the XML structure consumed by
  the existing Word converter.
- `hwp_reader.py`: signature-based HWP 3/5 dispatch and opt-in HWP 5 UTF-16 recovery. It serializes
  native parser output; it no longer adapts an external HWP parser.
- `hwp_limits.py`: shared positive resource budgets and CLI options for HWP 5.
- `convert_hwp.py`: launches the worker with a configurable timeout (60 seconds
  by default), spools its XML to a temporary file, writes editable DOCX, and
  validates text, table dimensions, image counts, section count and page settings,
  hyperlink destinations, and header/footer content and inheritance. The CLI keeps
  extracted images in a temporary directory for the full conversion lifetime.
- `olefile`: third-party generic OLE compound-file container reader, not an HWP
  parser. `python-docx` remains the third-party Word writer.

The original HWP file is never modified. Reports identify the reader as
`native-hwp3` or `native-hwp5`. No third-party HWP parser is installed by the
requirements file. Adding HWP 3 introduces no new package dependency; the
complete DOCX converter still uses the existing Word writer and HWP 5 OLE reader.

## HWP 3 Subset and Open Compatibility Questions

The reader accepts unprotected HWP 3.x with a 30-byte V3.00 signature and
subrevision 0 or 1. It handles uncompressed bodies, gzip framing, and raw DEFLATE
with or without a validated CRC32/size footer, followed by separate uncompressed
preview blocks. Compression variants are synthetic-test coverage, not claims
that every variant has been observed in a Hancom-generated file.

Supported content is ASCII and modern Hangul syllables, Johab font names, basic
character formatting (including super/subscript), paragraph spacing/indents,
tab stops, tabs/nonbreaking/fixed-width spaces, explicit paragraph page breaks,
page dimensions/margins, rectangular merged/nested tables and cell borders,
and flattened text boxes. HWP 3 hunit values are converted from 1/1800 inch to
the writer's 1/7200-inch units. Inherited paragraph formatting is local to each
paragraph list. Cell geometry determines row/column spans; overlapping or
incomplete grids are rejected.

Important limits and assumptions:

- `hchar` is Hancom's internal code, not UTF-16 or general Johab. Only ASCII
  and the modern syllable bit layout are mapped. Hanja, other symbols, standalone
  Jamo, and old Hangul are rejected. The published PDF refers to a separate
  manual for the full character table. `--recover-text` applies only to HWP 5.
- The PDF gives DOS alignment menu indices without their meanings. Paragraphs
  become left-aligned and the report warns instead of guessing an enum mapping.
- The implementation uses a 43-byte null-paragraph sentinel and counts the
  fixed control marker in hchar positions, excluding nested control payloads.
  These framing interpretations still need verification against genuine HWP 3
  files; no real HWP 3 fixture is currently included or locally validated.
- Stored line records do not contain HWP 5 y-coordinates. Their page-boundary
  flags are not used to reconstruct pages. `--preserve-source-pages` reports
  that limitation; explicit paragraph page breaks are always retained.
- Images/drawings, equations/buttons, fields, bookmarks as controls,
  headers/footers, footnotes, captions, diagonal cells, and column/section breaks
  are rejected. Text boxes containing tables are rejected.
- Character outlines/shadows/shading, cell shading, paragraph/page borders,
  binding margins, document metadata, and exact floating placement are not
  preserved; reports warn about encountered omissions/approximations.

Limits include 50 MB input, 64 MiB expanded body, 50,000 paragraphs, 2,000,000
logical hchars, 50,000 character styles, 32 nested paragraph lists, 10,000 grid
cells per table, and 50,000 table/textbox cells total. Unsupported body content
and malformed input fail before output publication.

## HWP 5 Supported Subset

Multi-section, unprotected HWP 5 documents, compressed or uncompressed; font
names and basic character/paragraph styles; tabs and stored line geometry;
inline and flattened floating tables with merged cells and nested tables;
basic borders and solid backgrounds; page dimensions and margins; basic
embedded pictures and rectangle text boxes; paragraph-local hyperlinks; and
section-initial headers/footers with text, tables, supported images, and links.

The parser rejects unsupported controls and malformed structures explicitly.
Unsupported header/footer structures, captions, grouped pictures, vector-only drawings,
and equations are not supported. Layout, image effects,
field behavior, and legacy private-use font mappings are not fully preserved.
Refer to the user README for conversion limitations.

Default limits are 50 MiB input, 64 MiB per stored/expanded stream, 128 MiB total
expanded data, and 200,000 records per stream. These are configurable with
`--max-input-mb`, `--max-stream-mb`, `--max-expanded-mb`, and `--max-records`.
`--reader-timeout` controls worker time in seconds. The total budget is checked
during decompression, before allocating a stream larger than the remaining
budget. Nesting remains limited to 64 levels and tables to 10,000 grid cells.
DEFLATE streams may end immediately or carry a validated CRC32/length trailer.
These checks and the worker timeout are not a security sandbox. The document
model and final image blobs still occupy memory; this is not a fully streaming
writer. CLI image payloads are extracted to temporary files instead of base64
XML. Direct parser/reader calls retain base64 output unless the caller supplies
an asset directory and keeps it alive until writing has finished. Generated file
names use the validated storage ID and extension, and the writer verifies that
image paths stay inside the supplied directory.

`BodyText/SectionN` streams must be consecutively numbered from zero, are read
in numeric order, and must match the DocInfo section count when it is present.
Each stream gets its own SectionDef while styles and image references remain
document-wide. Word sections retain page size, orientation, margins, and
header/footer distances. Later sections start on new pages; odd/even starts
and section numbering are not yet preserved and produce a warning. Missing
page settings inherit Word's preceding section settings, or defaults for the
first. Source-page inference resets at each section and rejects multiple columns
anywhere in the document. Saved section counts and explicit page settings are
checked before output publication; visual layout remains unverified.

`%hlk` command records and field-end markers become styled hyperlink spans.
The first command component is the address, with escaped semicolons and
backslashes decoded. HTTP(S), mailto, and FTP links use relationships belonging
to their actual Word part, including header/footer parts. Unsupported targets
retain labels with a warning; nested or cross-paragraph hyperlinks fail explicitly.
Other fields keep displayed text without active field behavior.

Headers/footers select both, even, or odd pages. Section-initial definitions
become separate Word stories, with explicit definitions unlinked from the prior
section and missing definitions inheriting it. Explicit empty definitions clear
inherited content. Duplicate selections and populated mid-section definitions
are rejected. Entirely empty legacy controls later in a document with no populated
headers/footers remain ignored. First-page hiding and dynamic page-number fields
are not implemented. Validation checks each emitted story, its tables/images,
hyperlink targets, and inheritance. Report body text hashes remain backward
compatible; header/footer text has its own count.

Windows is included in package metadata and uses the standard venv
`Scripts/python.exe` location. Cache setup and reuse are tested with simulated
Windows paths on macOS; actual Windows conversion is still unverified.

## Verification

```sh
.venv/bin/python -m unittest -v test_hwp_parser test_convert_hwp test_hwp3_parser test_hwp_limits_sections test_hwp_links_headers
npm test
```

Synthetic tests exercise record framing, compression/trailers, styles, control
ordering, Unicode recovery, nested tables, picture insertion, protection checks,
and malformed/unsupported data. Two optional private fixtures in `output/`
exercise existing formatting regressions and text hashes saved from earlier
conversions. Fixtures and generated documents remain untracked. Missing private
fixtures cause those tests to skip.

HWP 5 section tests build complete compound files and run them through the real
reader subprocess and Word writer, including compressed/uncompressed input,
11 sections (numeric ordering), alternating orientation, distinct margins,
invalid section names/counts, budget failures, and report checks. A sparse file
over 50 MiB verifies input-budget opt-in; it is not a large-content performance
benchmark. Stream expansion and cumulative budgets have separate boundary tests.

Link/header tests exercise complete compound files through the worker, styled
labels and controls, header-part relationships, tables, extracted header images,
odd/even selection, inheritance, empty overrides, validation failures, and
temporary-file cleanup. New link and populated-header binary structures are
covered by synthetic fixtures; broader real-document compatibility remains to
be established.

HWP 3 tests construct complete binary files from the published offsets and
cover the supported text/table subset, protection, size/depth budgets,
malformed input, subprocess dispatch, DOCX reports, and overwrite refusal.
The reader also runs with Python `-S` (no site-packages), establishing that
HWP 3 parsing needs no external Python packages. All 11,172 modern Hangul
syllables round-trip against Python's Johab codec; this is not independent
verification of Hancom's full hchar encoding. The local `.hwp` copy of the HWP 3
specification is itself an HWP 5 container and is not an HWP 3 test fixture.

Matching parser text and DOCX structure does not establish complete HWP feature
coverage or visual equivalence in Microsoft Word. Real-document image coverage
and broader HWP version coverage remain limited.

## Licensing

The new HWP implementation removes the pyhwp runtime dependency. It does not
change the repository's existing AGPL-3.0-or-later license or determine whether
the entire project can be relicensed. The generic container reader and Word
writer remain open-source dependencies. A separate ownership/license review is
required before distributing this project under other terms.

The Hancom attribution remains in the converter, CLI help, and user manuals.
