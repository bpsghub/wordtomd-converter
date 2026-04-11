# mdmaker

A Python CLI tool that converts Word documents (`.docx`) and PDF files (`.pdf`) to GitHub-flavoured Markdown (`.md`).

File type is detected automatically from the input extension.

Preserves headings, tables, nested lists, inline formatting, hyperlinks, and embedded images.

## Features

### DOCX
- **Headings** — `Heading 1–6` styles map to `#`–`######`
- **Tables** — rendered as GFM pipe tables with header separator
- **Lists** — ordered and unordered, nested, with correct counters
- **Inline formatting** — bold, italic, inline code, hyperlinks
- **Images** — extracted to a sibling directory; optional EMF/WMF → PNG conversion via Pillow
- **Code blocks** — consecutive `Code` style paragraphs become a single fenced block

### PDF
- **Headings** — detected from font size relative to body text (h1 ≥ 2×, h2 ≥ 1.6×, h3 ≥ 1.3×, h4 ≥ 1.1×)
- **Tables** — detected via PyMuPDF's table finder, rendered as GFM pipe tables
- **Lists** — bullet characters and ordered patterns (`1.`, `1)`) with indentation preserved
- **Inline formatting** — bold and italic from font flags; hyperlinks from PDF link annotations
- **Images** — extracted and saved as PNG to a sibling directory

## Install

```bash
# DOCX support only
pip install -e ".[images]"

# PDF support only
pip install -e ".[pdf]"

# Everything
pip install -e ".[images,pdf]"
```

| Extra | Package added | Purpose |
| --- | --- | --- |
| `images` | Pillow | EMF/WMF → PNG conversion for DOCX images |
| `pdf` | PyMuPDF | PDF parsing and rendering |
| `dev` | pytest | Running tests |

Requires **Python 3.9+**.

## Usage

```bash
# DOCX → Markdown (output written alongside the input file)
mdmaker input.docx

# PDF → Markdown
mdmaker input.pdf

# Explicit output path
mdmaker input.docx output.md
mdmaker input.pdf output.md

# Custom image subdirectory name
mdmaker input.docx --image-dir assets
mdmaker input.pdf --image-dir assets

# Skip image extraction
mdmaker input.docx --no-images
mdmaker input.pdf --no-images

# Verbose progress to stderr
mdmaker input.docx -v
mdmaker input.pdf -v

# Show version
mdmaker --version
```

**Exit codes:** `1` = file not found, `2` = unsupported extension, `3` = required optional dependency not installed.

## Output example

Given a document with a table and a list, `mdmaker` produces:

```markdown
# My Document

## Introduction

This is a **bold** statement with an _italic_ aside.

- Item one
- Item two
  - Nested item

| Name  | Role     | Score |
| ----- | -------- | ----- |
| Alice | Engineer | 95    |
| Bob   | Designer | 88    |
```

## Architecture

```
wordtomd/
├── cli.py            # argparse entry point; detects extension -> DocxConverter or PdfConverter
├── converter.py      # DocxConverter: walks body elements, dispatches to renderers
├── pdf_converter.py  # PdfConverter: PyMuPDF-based PDF -> Markdown conversion
├── relationships.py  # parses hyperlinks + images from word/_rels/
├── numbering.py      # resolves list format and counters from numbering.xml
├── postprocess.py    # collapses blank lines, strips trailing spaces (shared)
└── renderers/
    ├── inline.py     # bold, italic, code, hyperlinks
    ├── paragraph.py  # dispatches by style name
    ├── list_item.py  # indent + marker, numPr detection
    ├── table.py      # GFM pipe tables, merged cell handling
    └── image.py      # writes image file, returns ![](path)
```

## Dependencies

| Package | Purpose |
| --- | --- |
| `python-docx` | OOXML DOM access |
| `lxml` | XML operations |
| `Pillow` | EMF/WMF → PNG conversion (optional, `images` extra) |
| `pymupdf` | PDF parsing and rendering (optional, `pdf` extra) |

## Running tests

```bash
pytest tests/ -v
```

Test fixtures live in `tests/fixtures/`.

## License

MIT
