# wordtomd — Claude Code Context

## What this project is
A Python CLI tool that converts Word documents (`.docx`) to Markdown (`.md`), preserving headings, tables, lists, images, bold/italic, and hyperlinks.

## Install
```bash
pip install -e ".[images,dev]"
```
Requires Python 3.9+. The `images` extra adds Pillow for EMF/WMF conversion. The `dev` extra adds pytest.

## Run
```bash
wordtomd input.docx              # outputs input.md alongside the input file
wordtomd input.docx output.md    # explicit output path
wordtomd input.docx --no-images  # skip image extraction
wordtomd input.docx -v           # verbose progress to stderr
```

## Run tests
```bash
pytest tests/ -v
```
Test fixtures (`.docx` files) live in `tests/fixtures/`.

## Architecture

```
wordtomd/
├── cli.py            # argparse entry point → DocxConverter
├── converter.py      # DocxConverter: walks body elements, dispatches to renderers
├── relationships.py  # RelationshipMap: parses word/_rels/document.xml.rels
│                     #   → hyperlinks dict + images dict (rId → bytes)
├── numbering.py      # NumberingMap: parses word/numbering.xml
│                     #   → list format (bullet/ordered) + ordered counters
├── postprocess.py    # clean_output(): collapses blank lines, strips trailing spaces
└── renderers/
    ├── inline.py     # render_runs(): bold, italic, code, hyperlinks
    ├── paragraph.py  # render_paragraph(): dispatches by style name
    ├── list_item.py  # render_list_item(): indent + marker; has_num_pr() detection
    ├── table.py      # render_table(): GFM pipe tables, merged cell handling
    └── image.py      # ImageExtractor.extract(): writes image file, returns ![](path)
```

### Conversion flow
1. `RelationshipMap.from_docx()` + `NumberingMap.from_docx()` parse the ZIP directly
2. `docx.Document()` opens the same file for the high-level object model
3. `converter.py` walks `doc.element.body` children in document order
4. Each `w:p` → `render_paragraph()` → one of: heading, list item, code (buffered), image, plain paragraph
5. Each `w:tbl` → `render_table()`
6. `postprocess.clean_output()` normalizes whitespace before writing

### Key design decisions
- Uses `python-docx` (not `mammoth`) for direct OOXML access — needed for nested lists, GFM tables, and structural control
- List detection checks both the paragraph's direct `numPr` XML **and** the inherited style hierarchy (`_get_num_pr` in `list_item.py`)
- Consecutive code-style paragraphs are buffered and emitted as a single fenced block
- Blank lines between list groups use `numId` tracking in `converter.py`
- `RelationshipMap` and `NumberingMap` open the file as a raw `zipfile.ZipFile` before python-docx opens it — both reads are independent

## Dependencies
| Package | Purpose |
|---|---|
| `python-docx` | OOXML DOM access |
| `lxml` | XML operations (used by python-docx, explicit dep for xpath) |
| `Pillow` | EMF/WMF → PNG conversion (optional) |

## Windows notes
- The CLI prints ASCII `->` not `→` (Windows console cp1252 limitation)
- Use `pip install -e .` not `pip install -e ".[extras]"` if your shell doesn't support bracket quoting — use double quotes as shown above
