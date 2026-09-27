# HWP to Word

English | [한국어](README.ko.md)

Convert a local `.hwp` file into an editable Word document (`.docx`) from your
terminal. Text and tables remain editable, and your original file is unchanged.

This repository uses its own HWP 3 and HWP 5 parsers, implemented from Hancom's public
format specification, with no `pyhwp` dependency or fallback. It still uses
`olefile` for the HWP 5 file container and `python-docx` for Word output.
The new HWP 3 reader uses only Python's standard library; no new dependency was
added. HWP 3 support is an experimental subset, tested with synthetic fixtures.

**Experimental:** formatting and page breaks may differ from the original.
Review the result before sharing or submitting it. HWPX is not supported yet.
Windows runtime paths are implemented and covered by simulated tests; conversion
on an actual Windows machine still needs verification. This is a command-line
tool, not a Word add-in.

## Before You Start

- Use macOS or Linux, or try the experimental Windows support.
- Install [Node.js](https://nodejs.org/en/download) 18+ (includes `npx`) and
    [Python](https://www.python.org/downloads/) 3.9+.
- Prepare a trusted, unprotected HWP 3 or HWP 5 file. The default input limit is
    50 MiB; HWP 5 allows larger files with the options below and supports multiple
    document sections. HWP 3 retains its 50 MiB and single-section limits.
- Connect to the internet for the first run to download the tool and dependencies.

Check your installed versions in Terminal:

```sh
node --version
python3 --version
```

## Convert a File

Open Terminal in the folder containing your HWP file and run:

```sh
npx hwp-to-word "input.hwp" "output.docx"
```

Replace `input.hwp` with your file name and `output.docx` with your preferred
output name. Keep quotes around paths with spaces. You can also use full paths:

```sh
npx hwp-to-word "$HOME/Downloads/My document.hwp" "$HOME/Downloads/My document.docx"
```

If `npx` asks to install `hwp-to-word`, approve the installation to continue.
The first conversion takes longer while it prepares a Python environment.
You do not need to install Python dependencies yourself.

### Run This Checkout

To use the native parser in this repository, run the local CLI from the
repository root:

```sh
node bin/hwp-to-word.mjs "input.hwp" "output.docx"
```

It prepares the Python dependencies automatically and accepts the same options
shown below. `npx hwp-to-word` uses the published npm package, which may not yet
include the changes in this checkout.

## Open and Check the Result

Successful conversion creates two files:

| File | What to do with it |
| --- | --- |
| `output.docx` | Open in Microsoft Word or another DOCX-compatible editor. |
| `output.report.json` | Open in a text editor to review conversion warnings. |

Compare the text, tables, page breaks, and Korean fonts with the original.
Missing fonts can change wrapping and spacing. If you increase font sizes,
you may also need to adjust paragraph line spacing.

The report checks text against the parser's output, table dimensions, section
counts/page settings, and image counts. It does **not** prove that every source
feature or the visual layout was preserved. `visual_layout_verified: false` means layout was not automatically
verified; it does not by itself mean conversion failed.

`"parser": "native-hwp3"` or `"parser": "native-hwp5"` identifies the reader
used. The actual file signature selects the reader; both versions use `.hwp`.

Existing output files and reports are not overwritten. To convert again, choose
a new name, such as `output-v2.docx`.

## Optional Settings

### Larger HWP 5 Files

Use the local checkout to opt into larger resource budgets, for example:

```sh
node bin/hwp-to-word.mjs "large.hwp" "large.docx" --max-input-mb 200 --max-stream-mb 128 --max-expanded-mb 512 --max-records 500000 --reader-timeout 300
```

| Option | Default | Controls |
| --- | --- | --- |
| `--max-input-mb` | 50 | Input file size |
| `--max-stream-mb` | 64 | Each stored or expanded HWP 5 stream |
| `--max-expanded-mb` | 128 | Total expanded HWP 5 streams, including images |
| `--max-records` | 200000 | Records in each HWP 5 stream |
| `--reader-timeout` | 60 | Reader time in seconds |

All values must be positive integers. Size options use MiB (1,048,576 bytes).
Raising the input budget does not automatically raise the other budgets; errors
identify the exhausted budget. HWP 3 keeps its existing internal limits.
The report records the selected HWP 5 limits and reader timeout.

Large documents can require substantially more memory than their input size.
The CLI writes intermediate XML and extracted images to temporary files, avoiding
base64 image copies in the intermediate document. Temporary images are removed
after success or failure. The Word model and packaged image bytes still occupy
memory. These options do not promise
conversion of every file under the selected size or exact layout fidelity.

### Multiple HWP 5 Sections

Sections are read in numeric order and become separate Word sections with their
own page size, orientation, margins, and header/footer distances. Each subsequent
section starts a new Word page. Odd/even section starts and section numbering
remain unsupported. A section without page settings
inherits the preceding Word section's settings (or Word defaults for the first).
Page inference, when enabled, resets at each section and requires every section
to be single-column. The report verifies section count and explicit page settings;
it does not verify pagination.

### Links, Headers, and Footers

HWP 5 hyperlinks to web, email, and FTP addresses remain clickable and keep their
label formatting. Links must begin and end within one paragraph; nested links
and links spanning paragraphs are rejected. File links, internal bookmark links,
and unsupported addresses retain their displayed labels with a warning.

Headers and footers defined in the first paragraph of each section support text,
formatting, tables, supported images, and links. They can apply to both pages or
separately to odd and even pages. Later sections inherit prior definitions unless
they provide a replacement; an explicit empty definition clears inherited content.
Repeated definitions for the same page selection, populated definitions later in
a section, first-page hiding, and automatic page-number fields are not supported.
Previously accepted documents containing only empty header/footer controls still
convert. The report checks header/footer content, inheritance, and link targets.

The existing text count/hash describe body text. Header/footer text is counted
separately; table, image, and hyperlink counts include explicitly written
header/footer stories. A definition written to both odd and even stories counts
once in each story.

### Keep Original Page Starts

For a single-column HWP 5 document, try:

```sh
npx hwp-to-word "input.hwp" "output-pages.docx" --preserve-source-pages
```

This estimates page starts from stored HWP line positions. It cannot guarantee
identical wrapping or pagination and cannot be used for multicolumn documents.
For HWP 3 this option reports a warning and adds no inferred page breaks;
explicit paragraph page breaks are preserved automatically.

### Substitute Missing Fonts

Create a `font-map.json` file mapping original font names to fonts installed on
your computer. For example, if Nanum Gothic is installed:

```json
{
    "맑은 고딕": "Nanum Gothic"
}
```

```sh
npx hwp-to-word "input.hwp" "output-fonts.docx" --font-map "font-map.json"
```

This does not install fonts. Substitutions are recorded in the report and may
change layout. Without this option, original font names are retained.

### Recover Invalid Text Encoding

If an HWP 5 file fails to parse because of invalid UTF-16, you can explicitly try:

```sh
npx hwp-to-word "input.hwp" "output-recovered.docx" --recover-text
```

**This can change text.** Invalid sequences become replacement characters
(U+FFFD). Review `text_replacements` in the report and correct the affected text
in Word. Recovery is not a general fix for every conversion error or a lossless
repair of the original file.
HWP 3 uses a different character encoding. This option does not replace or
recover unsupported HWP 3 character codes.

## Troubleshooting

| Problem | What to try |
| --- | --- |
| `npx: command not found` | Install Node.js, then reopen Terminal. |
| `Python 3.9+ is required` | Install Python 3.9+. For a nonstandard location, set `HWP_TO_WORD_PYTHON` to the executable path. |
| Virtual environment creation fails | Check that Python includes `venv`; on some Linux distributions it is a separate package. |
| Dependency installation fails | Check internet access and proxy restrictions, then retry. |
| File not found | Check the path and extension. Quote paths containing spaces. |
| Output or report already exists | Choose a different output name. |
| HWPX or protected-file error | These inputs are unsupported. Renaming the extension will not convert the format. Ask the owner for an unprotected HWP 5 or DOCX copy. |
| Reader or validation failure | The file may contain unsupported or malformed content. Keep the original; do not assume a usable DOCX was created. Use `--recover-text` only for encoding-related failures. |
| WMF image error | WMF insertion is unsupported. Export a DOCX from the original authoring application instead. |
| Runtime setup is locked | Let other setup processes finish. Remove the lock directory named in the error only after confirming no setup is running. |
| Layout differs | Check installed fonts, try the page-start option for single-column files, and adjust the result in Word. Exact matching is not supported. |

For command help:

```sh
npx hwp-to-word --help
```

## Support and Limitations

### HWP 3 (experimental)

- Recognizes binary HWP 3.x by its file signature; supports uncompressed input
  and bounded gzip/raw DEFLATE decoding.
- Converts ASCII and modern Hangul syllables, basic font styling, tabs, spaces,
  paragraph spacing, page size/margins, and explicit paragraph page breaks.
- Converts rectangular tables, ordinary merged cells, nested tables, and basic
  cell borders. Text boxes become editable inline text.
- Rejects legacy Hanja/symbol/old-Hangul codes, images, equations, fields,
  headers/footers, footnotes, captions, diagonal cells, and column/section breaks.
- Paragraph alignment becomes left alignment with a warning. Cell shading,
  paragraph/page borders, character shadows/outlines, metadata, and exact
  positioning are not preserved. Stored line boundaries do not reconstruct pages.

**Real HWP 3 documents have not yet been verified.** Tests use files built from
Hancom's published structures, including compression and failure cases. The
specification refers to a separate manual for the full internal character map
and DOS alignment ordering. Compatibility and visual fidelity need validation
with genuine HWP 3 files; synthetic tests cannot establish either.

### HWP 5

Supports editable text and tables, basic text formatting, merged cells, table
borders and fills, and multiple sections with individual page sizes, orientations,
and margins. Image support is limited.

Automatic lists, equations, vector drawings, captions, and fields other than
the hyperlink subset above are not implemented. Text boxes become editable body
text. Headers, footers, and active links follow the limits described above.
Supported pictures are placed inline; floating positions, cropping, rotation,
and effects are not preserved. A PDF preview is not generated.

Unsupported controls, unsupported header/footer structures, captions, grouped
pictures, vector-only drawings, and malformed structures are rejected rather
than silently omitted. Legacy Hanyang private-use characters are retained as
Unicode private-use characters, with a warning; glyph compatibility is not
guaranteed. See [Parser Architecture](PARSER.md) for implementation and test scope.

## Privacy and Local Files

The converter runs locally and does not upload documents. npm and Python packages
are downloaded during setup. The Python environment is reused from
`~/.cache/hwp-to-word`, or `$XDG_CACHE_HOME/hwp-to-word` when configured.
Set `HWP_TO_WORD_CACHE` to use another cache directory.

Only convert documents you trust. The parser is not sandboxed; size and time
limits do not make malicious files safe. Reports may contain document-related
information, so review them before sharing.

## License and Third-Party Software

This project is licensed under **AGPL-3.0-or-later**. See the [LICENSE](LICENSE)
file for the full terms.

The HWP 3 and HWP 5 parsers are implemented in this repository using
[Hancom's public file-format specification](https://www.hancom.com/support/downloadCenter/hwpOwpml).
The project does not depend on `pyhwp`.

The converter uses these open-source libraries:

| Library | Purpose | License |
| --- | --- | --- |
| `olefile` | Reads the OLE container used by HWP 5 files | BSD-2-Clause, with additional retained PIL notices |
| `python-docx` | Creates editable Word documents | MIT |

These libraries and their dependencies retain their own licenses. When
redistributing them, preserve the required copyright notices, license terms,
and disclaimers.

Removing `pyhwp` does not change this project's AGPL-3.0-or-later license.

본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
