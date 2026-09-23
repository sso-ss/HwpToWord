# HWP to Word

English | [한국어](README.ko.md)

Convert a local `.hwp` file into an editable Word document (`.docx`) from your
terminal. Text and tables remain editable, and your original file is unchanged.

**Experimental:** formatting and page breaks may differ from the original.
Review the result before sharing or submitting it. HWPX and Windows are not
supported yet. This is a command-line tool, not a Word add-in.

## Before You Start

- Use macOS or Linux.
- Install [Node.js](https://nodejs.org/en/download) 18+ (includes `npx`) and
    [Python](https://www.python.org/downloads/) 3.9+.
- Prepare a trusted, unprotected HWP 5 file, no larger than 50 MB, with one
    document section. A section is a layout division, not a page: multi-page files can work.
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

## Open and Check the Result

Successful conversion creates two files:

| File | What to do with it |
| --- | --- |
| `output.docx` | Open in Microsoft Word or another DOCX-compatible editor. |
| `output.report.json` | Open in a text editor to review conversion warnings. |

Compare the text, tables, page breaks, and Korean fonts with the original.
Missing fonts can change wrapping and spacing. If you increase font sizes,
you may also need to adjust paragraph line spacing.

The report checks text against the parser's output, table dimensions, and image
counts. It does **not** prove that every source feature or the visual layout was
preserved. `visual_layout_verified: false` means layout was not automatically
verified; it does not by itself mean conversion failed.

Existing output files and reports are not overwritten. To convert again, choose
a new name, such as `output-v2.docx`.

## Optional Settings

### Keep Original Page Starts

For a single-column document, try:

```sh
npx hwp-to-word "input.hwp" "output-pages.docx" --preserve-source-pages
```

This estimates page starts from stored HWP line positions. It cannot guarantee
identical wrapping or pagination and cannot be used for multicolumn documents.

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

If a file fails to parse because of invalid UTF-16, you can explicitly try:

```sh
npx hwp-to-word "input.hwp" "output-recovered.docx" --recover-text
```

**This can change text.** Invalid sequences become replacement characters
(U+FFFD). Review `text_replacements` in the report and correct the affected text
in Word. Recovery is not a general fix for every conversion error or a lossless
repair of the original file.

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

Supports editable text and tables, basic text formatting, merged cells, table
borders and fills, page size, and margins. Image support is limited.

Automatic lists, equations, vector drawings, populated headers and footers,
captions, field behavior, and multiple sections are not implemented. Text boxes
become editable body text. Hyperlink labels remain, but links are not active.
Supported pictures are placed inline; floating positions, cropping, rotation,
and effects are not preserved. A PDF preview is not generated.

## Privacy and Local Files

The converter runs locally and does not upload documents. npm and Python packages
are downloaded during setup. The Python environment is reused from
`~/.cache/hwp-to-word`, or `$XDG_CACHE_HOME/hwp-to-word` when configured.
Set `HWP_TO_WORD_CACHE` to use another cache directory.

Only convert documents you trust. The parser is not sandboxed; size and time
limits do not make malicious files safe. Reports may contain document-related
information, so review them before sharing.

## License

[AGPL-3.0-or-later](LICENSE). You may redistribute and modify this software under
the license terms. It is provided without warranty. Review source-sharing and
network-use obligations before redistributing it or offering it as a service.

Uses [pyhwp](https://github.com/mete0r/pyhwp) (AGPL-3.0-or-later) and python-docx
(MIT). Based on [Hancom's public HWP specification](https://www.hancom.com/support/downloadCenter/hwpOwpml).

본 제품은 한컴의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.