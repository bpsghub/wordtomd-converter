# wordtomd

A Python CLI tool that converts Word documents (`.docx`) to GitHub-flavoured Markdown (`.md`).

Preserves headings, tables, nested lists, inline formatting, hyperlinks, and embedded images.

## Features

- **Headings** — `Heading 1–6` styles map to `#`–`######`
- **Tables** — rendered as GFM pipe tables with header separator
- **Lists** — ordered and unordered, nested, with correct counters
- **Inline formatting** — bold, italic, inline code, hyperlinks
- **Images** — extracted to a sibling directory; optional EMF/WMF → PNG conversion via Pillow
- **Code blocks** — consecutive `Code` style paragraphs become a single fenced block

## Install

```bash
pip install -e ".[images]"
```

- `images` extra adds **Pillow** for EMF/WMF conversion (optional)
- `dev` extra adds **pytest**

Requires **Python 3.9+**.

## Usage

```bash
# Output written alongside the input file (input.md)
wordtomd input.docx

# Explicit output path
wordtomd input.docx output.md

# Custom image subdirectory name
wordtomd input.docx --image-dir assets

# Skip image extraction
wordtomd input.docx --no-images

# Verbose progress to stderr
wordtomd input.docx -v

# Show version
wordtomd --version
```

## Output example

Given a Word document with a table and a list, `wordtomd` produces:

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
├── cli.py            # argparse entry point -> DocxConverter
├── converter.py      # walks body elements, dispatches to renderers
├── relationships.py  # parses hyperlinks + images from word/_rels/
├── numbering.py      # resolves list format and counters from numbering.xml
├── postprocess.py    # collapses blank lines, strips trailing spaces
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
| `Pillow` | EMF/WMF → PNG conversion (optional) |

## Running tests

```bash
pytest tests/ -v
```

Test fixtures live in `tests/fixtures/`.

## License

MIT
