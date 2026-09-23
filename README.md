# Local HWP to Word Prototype

Converts a single-section, unprotected binary HWP 5 document into an editable
DOCX. This is a conversion experiment, not an Office Add-in or a general-purpose
production converter. HWPX is not implemented yet.

## Run

Requires Python 3.9 or newer. Package installation needs internet access;
conversion runs locally and does not upload documents.

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python convert_hwp.py input.hwp output/converted.docx
.venv/bin/python -m unittest -v test_convert_hwp
```

For a Mac copy using installed alternative Korean fonts:

```sh
.venv/bin/python convert_hwp.py input.hwp output/refined.docx --font-map fonts-macos.json
```

The optional map changes font families, not content, and each substitution is
listed in the JSON report. Ensure the mapped fonts are installed. Without this
option, the original font names are retained. Missing fonts can change layout.

Existing output files are never intentionally overwritten. Choose a new output
name when rerunning. The converter creates a DOCX and a matching JSON report.

For single-column documents, `--preserve-source-pages` can retain page starts
inferred from stored HWP line positions. This uses cached source geometry,
not a rendering of the original, and is rejected for multicolumn documents.
It does not force individual line wraps or guarantee identical pagination after
editing. Source tab stops and inline-table alignment are preserved. Cell padding
uses table defaults when the stored line widths confirm those defaults were used.

The agreement's revised output is `output/partner-agreement-refined.docx`, with
a local LibreOffice preview at `output/partner-agreement-refined.pdf`. It retains
the inferred two-page structure and starts Article 9 on page 2. Signature-cell
wrapping is corrected. Exact HWP/Word rendering and original fonts remain
unverified; the output report intentionally does not claim a visual match.

## Malformed Text Recovery

Strict UTF-16 decoding is the default. For a document that fails on invalid
UTF-16, explicitly enable recovery:

```sh
.venv/bin/python convert_hwp.py input.hwp output/recovered.docx --recover-text
```

Recovery runs in an isolated reader process without changing the source or
installed parser. Invalid sequences become U+FFFD replacement characters.
The output report lists invalid bytes, reader chunk indexes, and byte offsets
within each chunk (not absolute file offsets). Text preservation then means
matching the recovered parser text, not reproducing the invalid original bytes.
Valid surrogate pairs and Korean text are unchanged.

The functional-ingredient guide passes parsing with one replacement, but its
WMF image cannot yet be inserted by python-docx. No DOCX was published for it.
Nested tables, inline picture insertion for library-supported formats, and
flattened editable text-box/hyperlink labels have been added. WMF conversion
and visual validation of this guide remain outstanding.

## Implemented

- Native Word text runs, font names, sizes, bold, italic, underline, and color.
- Paragraph alignment, source line heights, hanging indents, and page breaks.
- Character spacing and horizontal scaling from the HWP character styles.
- Native Word tables, rectangular merged cells, widths, borders, and solid fills.
- Page dimensions and margins.
- Rejection of password, distribution, and DRM protected inputs.
- Exact source-text comparison, table count/dimension checks, ZIP integrity,
  and reopening with python-docx before publishing the output file.

The supplied mentoring sample passed the automated checks: 740 text characters,
two tables, and a 10-row by 3-column schedule table. These checks do not prove
pixel-perfect layout or full compatibility with all HWP documents.

## Limitations

Font availability and pagination need visual review in Word. HWP and Word have
different layout engines. Automatic lists, vector pictures, drawing geometry, equations,
nonempty headers/footers, field behavior, captions, and multiple sections are
not implemented. Supported embedded pictures are inserted inline; their floating
positions, crops, rotations, and effects are not preserved. Text boxes become
inline editable text and hyperlinks lose their active destinations. Unsupported
paragraph objects generate warnings, and lost
text causes validation to fail rather than silently publishing the document.
Other formatting differences may not be detected yet. Image support is preliminary;
the guide conversion is still blocked on a WMF asset. Do not use this prototype on
untrusted files: size/time limits are not a parser sandbox.

The revised sample is `output/mentoring-refined.docx`, with a locally rendered
PDF preview at `output/mentoring-refined.pdf`. It renders on one page in
LibreOffice with the installed font alternatives. Word automation did not
produce a usable export, so Word rendering remains unverified and the JSON
report continues to record `visual_layout_verified: false`.

The reference PDF helped identify and fix duplicated table-height blank
paragraphs, hanging indents, table preferred widths, and omitted character
spacing. The reference and the revised output are not pixel-identical:
substitute fonts and table row spacing still differ. Line heights are explicit
to preserve the source layout; increasing font sizes may require adjusting
paragraph line spacing in Word. Quick Look is not a reliable fidelity check.

## Licensing And Attribution

The reader, pyhwp 0.1b15, is AGPL-3.0-or-later. Its use here is experimental;
review licensing obligations or select another parser before distributing or
hosting a product. python-docx uses the MIT license. This repository does not
copy either library's implementation and does not establish a product license.

Hancom's public format documentation is the specification reference:
https://www.hancom.com/support/downloadCenter/hwpOwpml

본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

Private sample files and generated output are excluded by .gitignore.