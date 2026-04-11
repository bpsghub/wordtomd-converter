# wordtomd — Claude Code Context

## What this project is
A Python CLI tool that converts Word documents (`.docx`) and PDF files (`.pdf`) to Markdown (`.md`), preserving headings, tables, lists, images, bold/italic, and hyperlinks. File type is detected automatically from the input extension.

## Install
```bash
pip install -e ".[images,dev]"        # DOCX support + dev tools
pip install -e ".[pdf,dev]"           # PDF support + dev tools
pip install -e ".[images,pdf,dev]"    # everything
```
Requires Python 3.9+. The `images` extra adds Pillow for EMF/WMF conversion. The `pdf` extra adds PyMuPDF for PDF conversion. The `dev` extra adds pytest.

## Run
```bash
wordtomd input.docx              # outputs input.md alongside the input file
wordtomd input.docx output.md    # explicit output path
wordtomd input.docx --no-images  # skip image extraction
wordtomd input.docx -v           # verbose progress to stderr

wordtomd input.pdf               # PDF → Markdown (requires [pdf] extra)
wordtomd input.pdf output.md     # explicit output path
wordtomd input.pdf --no-images   # skip image extraction
wordtomd input.pdf -v            # verbose progress to stderr
```

Exit codes: `1` = file not found, `2` = unsupported extension, `3` = missing optional dependency (`pymupdf` not installed).

## Run tests
```bash
pytest tests/ -v
```
Test fixtures (`.docx` files) live in `tests/fixtures/`.

## Architecture

```
wordtomd/
├── cli.py            # argparse entry point; detects extension → DocxConverter or PdfConverter
├── converter.py      # DocxConverter: walks body elements, dispatches to renderers
├── pdf_converter.py  # PdfConverter: PyMuPDF-based PDF → Markdown conversion
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

### DOCX conversion flow
1. `RelationshipMap.from_docx()` + `NumberingMap.from_docx()` parse the ZIP directly
2. `docx.Document()` opens the same file for the high-level object model
3. `converter.py` walks `doc.element.body` children in document order
4. Each `w:p` → `render_paragraph()` → one of: heading, list item, code (buffered), image, plain paragraph
5. Each `w:tbl` → `render_table()`
6. `postprocess.clean_output()` normalizes whitespace before writing

### PDF conversion flow (`pdf_converter.py`)
1. `fitz.open()` opens the PDF via PyMuPDF
2. `_collect_blocks()` iterates all pages; per page:
   - `page.find_tables()` identifies table bounding boxes to exclude from text extraction
   - `page.get_links()` builds a hyperlink map (bbox → URL)
   - `page.get_text("dict")` yields text blocks with per-span font size, flags, and origin
   - Image blocks (`type=1`) and table sentinel blocks are appended separately
   - All blocks are sorted by `(page_num, y0)` for reading order
3. `_estimate_body_size()` finds the modal font size (weighted by character count, bucketed to 0.5pt)
4. `_render_blocks()` dispatches each block:
   - Text → heading (font size ratio), list (bullet chars / ordered pattern), or paragraph
   - Table → GFM pipe table via `table.extract()`
   - Image → `fitz.Pixmap` saved as PNG
5. `postprocess.clean_output()` normalizes whitespace before writing

### Key design decisions
- **DOCX**: Uses `python-docx` (not `mammoth`) for direct OOXML access — needed for nested lists, GFM tables, and structural control
- **DOCX**: List detection checks both the paragraph's direct `numPr` XML **and** the inherited style hierarchy (`_get_num_pr` in `list_item.py`)
- **DOCX**: Consecutive code-style paragraphs are buffered and emitted as a single fenced block
- **DOCX**: Blank lines between list groups use `numId` tracking in `converter.py`
- **PDF**: Heading level is inferred from font size relative to the document body size (h1 ≥ 2×, h2 ≥ 1.6×, h3 ≥ 1.3×, h4 ≥ 1.1×)
- **PDF**: Bold/italic detected via PyMuPDF font flags (bit 4 = bold, bit 0 = italic)
- **PDF**: Table regions are detected first; overlapping text blocks are suppressed to avoid duplication
- **PDF**: List indentation level inferred from block `x0` relative to left margin (`_LEFT_MARGIN_PT`, `_INDENT_STEP_PT` constants)
- **PDF**: `PdfConverter` mirrors `DocxConverter`'s constructor signature — CLI dispatches via duck typing on `.convert()`
- Both converters share `postprocess.clean_output()` for final whitespace normalization

## Dependencies
| Package | Purpose |
|---|---|
| `python-docx` | OOXML DOM access (DOCX) |
| `lxml` | XML operations (used by python-docx, explicit dep for xpath) |
| `Pillow` | EMF/WMF → PNG conversion (optional, `images` extra) |
| `pymupdf` | PDF parsing and rendering (optional, `pdf` extra) |

## Windows notes
- The CLI prints ASCII `->` not `→` (Windows console cp1252 limitation)
- Use `pip install -e .` not `pip install -e ".[extras]"` if your shell doesn't support bracket quoting — use double quotes as shown above
