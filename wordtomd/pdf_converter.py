"""PDF to Markdown converter using PyMuPDF (fitz)."""

from __future__ import annotations

import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # pymupdf

from wordtomd.postprocess import clean_output

# ---------------------------------------------------------------------------
# Font flag bits (PyMuPDF)
# ---------------------------------------------------------------------------
_FLAG_ITALIC = 1 << 0
_FLAG_BOLD = 1 << 4

# ---------------------------------------------------------------------------
# List detection: characters that unambiguously mark a bullet item
# ---------------------------------------------------------------------------
_BULLET_CHARS = frozenset(
    "•○▪▸▶·◦►✓✗✦★"
)

# ---------------------------------------------------------------------------
# Heading detection: (minimum size ratio vs body, markdown prefix)
# Ratios are compared as: span_size / body_size >= min_ratio
# ---------------------------------------------------------------------------
_HEADING_SIZE_RATIOS: List[Tuple[float, str]] = [
    (2.0, "# "),
    (1.6, "## "),
    (1.3, "### "),
    (1.1, "#### "),
]

# ---------------------------------------------------------------------------
# List indentation heuristics
# ---------------------------------------------------------------------------
_LEFT_MARGIN_PT = 50.0   # typical PDF left margin in points
_INDENT_STEP_PT = 18.0   # ~one tab stop

# ---------------------------------------------------------------------------
# Markdown metacharacter escaping
# ---------------------------------------------------------------------------
_MD_ESCAPE_RE = re.compile(r"([\\\`\*\_\{\}\[\]\(\)#\+\-\.!|])")

# ---------------------------------------------------------------------------
# Ordered list pattern: "1.", "2)", "(3)", "a.", "b)" …
# ---------------------------------------------------------------------------
_ORDERED_RE = re.compile(r"^\(?[0-9]+[\.\)]\s|^[a-zA-Z][\.\)]\s")


def _escape_md(text: str) -> str:
    """Escape Markdown metacharacters in plain text."""
    return _MD_ESCAPE_RE.sub(r"\\\1", text)


# ---------------------------------------------------------------------------
# Internal data structures
# ---------------------------------------------------------------------------

@dataclass
class _Span:
    text: str
    size: float
    flags: int
    origin_x: float
    origin_y: float
    url: str = ""   # non-empty when span falls inside a hyperlink bbox


@dataclass
class _Block:
    page_num: int
    y0: float           # top of bbox — reading-order sort key
    x0: float           # left edge — list indentation inference
    spans: List[_Span] = field(default_factory=list)
    block_type: str = "text"   # "text" | "image" | "table"
    image_index: int = -1      # index into page.get_images() (image blocks)
    bbox: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)


# ---------------------------------------------------------------------------
# PdfConverter
# ---------------------------------------------------------------------------

class PdfConverter:
    """Convert a PDF file to GitHub-flavored Markdown.

    Constructor signature mirrors DocxConverter for interchangeable use in
    the CLI dispatcher.
    """

    def __init__(
        self,
        input_path: Path,
        output_path: Path,
        image_dir: Optional[str] = None,
        extract_images: bool = True,
        verbose: bool = False,
    ) -> None:
        self.input_path = input_path
        self.output_path = output_path
        self.extract_images = extract_images
        self.verbose = verbose

        images_dir_name = image_dir or (output_path.stem + "_images")
        self.images_dir = output_path.parent / images_dir_name
        self._image_counter = 0
        self._images_dir_created = False

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------

    def convert(self) -> None:
        self._log(f"Opening {self.input_path}")
        doc = fitz.open(str(self.input_path))

        blocks = self._collect_blocks(doc)
        body_size = self._estimate_body_size(blocks)
        self._log(f"Estimated body font size: {body_size:.1f}pt")

        output_lines = self._render_blocks(blocks, body_size, doc)
        doc.close()

        result = clean_output(output_lines)
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        self.output_path.write_text(result, encoding="utf-8")
        self._log(f"Written to {self.output_path}")

    # ------------------------------------------------------------------
    # Phase 1: collect all blocks across all pages
    # ------------------------------------------------------------------

    def _collect_blocks(self, doc: fitz.Document) -> List[_Block]:
        all_blocks: List[_Block] = []

        for page_num in range(len(doc)):
            page = doc[page_num]
            table_bboxes = self._get_table_bboxes(page)
            link_map = self._get_link_map(page)

            page_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

            for raw in page_dict.get("blocks", []):
                btype = raw.get("type", 0)
                bbox = tuple(raw["bbox"])  # (x0, y0, x1, y1)

                if btype == 1:  # image block
                    if not self.extract_images:
                        continue
                    if self._overlaps_any_table(bbox, table_bboxes):
                        continue
                    all_blocks.append(_Block(
                        page_num=page_num,
                        y0=bbox[1],
                        x0=bbox[0],
                        block_type="image",
                        image_index=raw.get("number", -1),
                        bbox=bbox,
                    ))
                    continue

                if btype != 0:
                    continue

                if self._overlaps_any_table(bbox, table_bboxes):
                    continue

                spans: List[_Span] = []
                for line in raw.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "")
                        if not text.strip():
                            continue
                        origin = span.get("origin", (0.0, 0.0))
                        url = self._resolve_link(origin[0], origin[1], link_map)
                        spans.append(_Span(
                            text=text,
                            size=span.get("size", 0.0),
                            flags=span.get("flags", 0),
                            origin_x=origin[0],
                            origin_y=origin[1],
                            url=url,
                        ))

                if not spans:
                    continue

                all_blocks.append(_Block(
                    page_num=page_num,
                    y0=bbox[1],
                    x0=bbox[0],
                    spans=spans,
                    block_type="text",
                    bbox=bbox,
                ))

            # Append a sentinel block for each table so the renderer can find it
            for tb in table_bboxes:
                all_blocks.append(_Block(
                    page_num=page_num,
                    y0=tb[1],
                    x0=tb[0],
                    block_type="table",
                    bbox=tb,
                ))

        all_blocks.sort(key=lambda b: (b.page_num, b.y0))
        return all_blocks

    def _get_table_bboxes(
        self, page: fitz.Page
    ) -> List[Tuple[float, float, float, float]]:
        try:
            tabs = page.find_tables()
            return [tuple(t.bbox) for t in tabs.tables]
        except Exception:
            return []

    def _overlaps_any_table(
        self,
        bbox: Tuple,
        table_bboxes: List[Tuple],
        threshold: float = 0.5,
    ) -> bool:
        if not table_bboxes:
            return False
        bx0, by0, bx1, by1 = bbox[0], bbox[1], bbox[2], bbox[3]
        b_area = max(0.0, bx1 - bx0) * max(0.0, by1 - by0)
        if b_area == 0:
            return False
        for tb in table_bboxes:
            tx0, ty0, tx1, ty1 = tb[0], tb[1], tb[2], tb[3]
            ix0 = max(bx0, tx0)
            iy0 = max(by0, ty0)
            ix1 = min(bx1, tx1)
            iy1 = min(by1, ty1)
            if ix1 > ix0 and iy1 > iy0:
                overlap = (ix1 - ix0) * (iy1 - iy0)
                if overlap / b_area >= threshold:
                    return True
        return False

    def _get_link_map(
        self, page: fitz.Page
    ) -> List[Tuple[Tuple[float, float, float, float], str]]:
        result = []
        for link in page.get_links():
            if link.get("kind") == fitz.LINK_URI:
                uri = link.get("uri", "")
                frm = link.get("from")
                if uri and frm:
                    result.append((tuple(frm), uri))
        return result

    def _resolve_link(
        self,
        x: float,
        y: float,
        link_map: List[Tuple[Tuple, str]],
    ) -> str:
        for bbox, uri in link_map:
            lx0, ly0, lx1, ly1 = bbox[0], bbox[1], bbox[2], bbox[3]
            if lx0 <= x <= lx1 and ly0 <= y <= ly1:
                return uri
        return ""

    # ------------------------------------------------------------------
    # Phase 2: estimate body font size
    # ------------------------------------------------------------------

    def _estimate_body_size(self, blocks: List[_Block]) -> float:
        size_counts: Counter = Counter()
        for block in blocks:
            if block.block_type != "text":
                continue
            for span in block.spans:
                # Bucket to 0.5pt to merge near-identical sizes
                rounded = round(span.size * 2) / 2
                size_counts[rounded] += len(span.text)
        if not size_counts:
            return 11.0
        return size_counts.most_common(1)[0][0]

    # ------------------------------------------------------------------
    # Phase 3: render blocks to Markdown lines
    # ------------------------------------------------------------------

    def _render_blocks(
        self,
        blocks: List[_Block],
        body_size: float,
        doc: fitz.Document,
    ) -> List[str]:
        output_lines: List[str] = []
        last_block_type = "empty"
        # Cache find_tables() results per page to avoid repeated calls
        table_cache: Dict[int, object] = {}

        def emit(block_type: str, lines: List[str]) -> None:
            nonlocal last_block_type
            if not lines:
                return
            needs_blank = (
                bool(output_lines)
                and last_block_type != "empty"
                and not (last_block_type == "list" and block_type == "list")
            )
            if needs_blank:
                output_lines.append("")
            output_lines.extend(lines)
            last_block_type = block_type

        for block in blocks:
            if block.block_type == "image":
                emit("image", self._render_image_block(block, doc))

            elif block.block_type == "table":
                if block.page_num not in table_cache:
                    try:
                        table_cache[block.page_num] = doc[block.page_num].find_tables()
                    except Exception:
                        table_cache[block.page_num] = None
                emit("table", self._render_table_block(block, table_cache[block.page_num]))

            elif block.block_type == "text":
                lines, btype = self._render_text_block(block, body_size)
                emit(btype, lines)

        return output_lines

    # ------------------------------------------------------------------
    # Text block rendering
    # ------------------------------------------------------------------

    def _render_text_block(
        self,
        block: _Block,
        body_size: float,
    ) -> Tuple[List[str], str]:
        if not block.spans:
            return [], "empty"

        # Dominant size: span with the most characters
        dominant_size = max(block.spans, key=lambda s: len(s.text)).size
        heading_prefix = self._size_to_heading(dominant_size, body_size)

        inline = self._render_inline_spans(block.spans)
        if not inline.strip():
            return [], "empty"

        if heading_prefix:
            return [f"{heading_prefix}{inline.strip()}"], "heading"

        first_text = block.spans[0].text.lstrip()
        if self._is_bullet_span(first_text):
            return [self._render_list_item(block, inline)], "list"

        return [inline.strip()], "paragraph"

    def _size_to_heading(self, size: float, body_size: float) -> str:
        if body_size <= 0:
            return ""
        ratio = size / body_size
        for min_ratio, prefix in _HEADING_SIZE_RATIOS:
            if ratio >= min_ratio:
                return prefix
        return ""

    def _is_bullet_span(self, text: str) -> bool:
        if not text:
            return False
        if text[0] in _BULLET_CHARS:
            return True
        if _ORDERED_RE.match(text):
            return True
        return False

    def _render_list_item(self, block: _Block, inline: str) -> str:
        indent_level = max(0, int((block.x0 - _LEFT_MARGIN_PT) / _INDENT_STEP_PT))
        indent = "  " * indent_level
        text = inline.strip()

        # Strip leading bullet character
        if text and text[0] in _BULLET_CHARS:
            text = text[1:].lstrip()
            return f"{indent}- {text}"

        # Strip ordered marker and preserve order (emit as "1." for simplicity)
        cleaned = _ORDERED_RE.sub("", text)
        if cleaned != text:
            return f"{indent}1. {cleaned.lstrip()}"

        return f"{indent}- {text}"

    def _render_inline_spans(self, spans: List[_Span]) -> str:
        parts = []
        for span in spans:
            text = span.text
            if not text:
                continue

            is_bold = bool(span.flags & _FLAG_BOLD)
            is_italic = bool(span.flags & _FLAG_ITALIC)

            # Also check font name for bold/italic when flags aren't set
            # (some PDFs embed bold as a separate font file)
            font_name = ""  # not available on _Span; flag-based detection is sufficient

            escaped = _escape_md(text)

            if is_bold and is_italic:
                formatted = f"***{escaped}***"
            elif is_bold:
                formatted = f"**{escaped}**"
            elif is_italic:
                formatted = f"_{escaped}_"
            else:
                formatted = escaped

            if span.url:
                formatted = f"[{formatted}]({span.url})"

            parts.append(formatted)

        return "".join(parts)

    # ------------------------------------------------------------------
    # Table block rendering
    # ------------------------------------------------------------------

    def _render_table_block(self, block: _Block, table_finder: object) -> List[str]:
        if table_finder is None:
            return []

        # Find the table whose bbox matches block.bbox within 1pt tolerance
        target = None
        for t in table_finder.tables:
            tb = tuple(t.bbox)
            if all(abs(tb[i] - block.bbox[i]) < 1.0 for i in range(4)):
                target = t
                break

        if target is None:
            return []

        try:
            rows = target.extract()
        except Exception:
            return []

        if not rows:
            return []

        col_count = max(len(row) for row in rows)

        def normalize_row(row: list) -> List[str]:
            cells = []
            for cell in row:
                text = (cell or "").replace("|", r"\|").replace("\n", "<br>").strip()
                cells.append(text)
            while len(cells) < col_count:
                cells.append("")
            return cells[:col_count]

        def pipe_row(cells: List[str]) -> str:
            return "| " + " | ".join(cells) + " |"

        header = normalize_row(rows[0])
        separator = ["---"] * col_count
        lines = [pipe_row(header), pipe_row(separator)]
        for row in rows[1:]:
            lines.append(pipe_row(normalize_row(row)))

        return lines

    # ------------------------------------------------------------------
    # Image block rendering
    # ------------------------------------------------------------------

    def _render_image_block(self, block: _Block, doc: fitz.Document) -> List[str]:
        self._image_counter += 1
        label = f"image-{self._image_counter}"

        if not self.extract_images:
            return [f"![{label}]()"]

        page = doc[block.page_num]
        img_list = page.get_images(full=True)
        idx = block.image_index

        if idx < 0 or idx >= len(img_list):
            return [f"<!-- image index {idx} out of range -->"]

        xref = img_list[idx][0]
        filename = f"{label}.png"

        try:
            pix = fitz.Pixmap(doc, xref)
            # Convert CMYK/other colorspaces to RGB for PNG compatibility
            if pix.colorspace and pix.colorspace.n > 3:
                pix = fitz.Pixmap(fitz.csRGB, pix)

            if not self._images_dir_created:
                self.images_dir.mkdir(parents=True, exist_ok=True)
                self._images_dir_created = True

            out_path = self.images_dir / filename
            pix.save(str(out_path))
            pix = None  # release

            rel_path = f"{self.images_dir.name}/{filename}"
            return [f"![{label}]({rel_path})"]
        except Exception as exc:
            self._log(f"Warning: could not extract image xref={xref}: {exc}")
            return [f"<!-- image xref={xref} extraction failed -->"]

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _log(self, msg: str) -> None:
        if self.verbose:
            print(f"[wordtomd] {msg}", file=sys.stderr)
