"""Dispatch block-level paragraph rendering."""

from __future__ import annotations

from typing import List, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from wordtomd.numbering import NumberingMap
    from wordtomd.relationships import RelationshipMap
    from wordtomd.renderers.image import ImageExtractor

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_HEADING_STYLES = {
    "heading 1": "#",
    "heading 2": "##",
    "heading 3": "###",
    "heading 4": "####",
    "heading 5": "#####",
    "heading 6": "######",
}

_CODE_STYLES = {"code", "preformatted", "html preformatted", "source code"}

BlockType = str  # "heading" | "list" | "code" | "paragraph" | "image" | "empty"


def _get_style_name(para) -> str:
    try:
        return (para.style.name or "").lower()
    except Exception:
        return ""


def _extract_images(para, image_extractor: "ImageExtractor") -> List[str]:
    """Return Markdown image strings for any drawings in this paragraph."""
    results = []
    p_el = para._p

    # Find all blip elements (a:blip) which reference images
    blip_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for blip in p_el.iter(f"{{{blip_ns}}}blip"):
        r_embed = blip.get(f"{{{_R_NS}}}embed", "")
        # Try to get alt text from the parent drawing
        alt = ""
        results.append(image_extractor.extract(r_embed, alt))

    return results


def render_paragraph(
    para,
    rel_map: "RelationshipMap",
    num_map: "NumberingMap",
    image_extractor: "ImageExtractor",
) -> Tuple[BlockType, List[str]]:
    """Return (block_type, lines) for a single paragraph."""
    from wordtomd.renderers.inline import render_runs
    from wordtomd.renderers.list_item import render_list_item, has_num_pr

    style_name = _get_style_name(para)

    # --- Images (drawings inside paragraph) ---
    image_lines = _extract_images(para, image_extractor)
    if image_lines:
        return ("image", image_lines)

    # --- Headings ---
    if style_name in _HEADING_STYLES:
        prefix = _HEADING_STYLES[style_name]
        text = render_runs(para, rel_map).strip()
        if not text:
            return ("empty", [])
        return ("heading", [f"{prefix} {text}"])

    # --- List items ---
    if has_num_pr(para):
        line = render_list_item(para, rel_map, num_map)
        return ("list", [line])

    # --- Code blocks (handled as buffered fences in converter) ---
    if style_name in _CODE_STYLES:
        text = render_runs(para, rel_map)
        return ("code", [text])

    # --- Normal text ---
    text = render_runs(para, rel_map).strip()
    if not text:
        return ("empty", [])
    return ("paragraph", [text])
