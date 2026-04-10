"""Render list items with correct indentation and ordered/bullet markers."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from wordtomd.numbering import NumberingMap
    from wordtomd.relationships import RelationshipMap

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _get_num_pr_from_xml(ppr_el):
    """Extract (numId, ilvl) directly from a pPr XML element, or (None, 0)."""
    if ppr_el is None:
        return None, 0
    num_pr = ppr_el.find(f"{{{_W_NS}}}numPr")
    if num_pr is None:
        return None, 0
    ilvl_el = num_pr.find(f"{{{_W_NS}}}ilvl")
    num_id_el = num_pr.find(f"{{{_W_NS}}}numId")
    ilvl = int(ilvl_el.get(f"{{{_W_NS}}}val", "0")) if ilvl_el is not None else 0
    num_id = num_id_el.get(f"{{{_W_NS}}}val", "") if num_id_el is not None else ""
    # numId="0" means numbering is explicitly disabled
    if num_id == "0":
        return None, 0
    return num_id or None, ilvl


def _get_num_pr(para):
    """Return (numId, ilvl) from a paragraph's numPr (direct or style-inherited)."""
    p_el = para._p
    ppr = p_el.find(f"{{{_W_NS}}}pPr")

    # 1. Check direct paragraph numPr
    num_id, ilvl = _get_num_pr_from_xml(ppr)
    if num_id:
        return num_id, ilvl

    # 2. Check inherited from paragraph style
    try:
        style = para.style
        while style is not None:
            if style.element is not None:
                style_ppr = style.element.find(f"{{{_W_NS}}}pPr")
                num_id, ilvl = _get_num_pr_from_xml(style_ppr)
                if num_id:
                    return num_id, ilvl
            # Walk up the style hierarchy
            try:
                style = style.base_style
            except Exception:
                break
    except Exception:
        pass

    return None, 0


def render_list_item(para, rel_map: "RelationshipMap", num_map: "NumberingMap") -> str:
    """Return a single Markdown list item line."""
    from wordtomd.renderers.inline import render_runs

    num_id, ilvl = _get_num_pr(para)
    indent = "  " * ilvl

    if num_id:
        fmt = num_map.get_format(num_id, ilvl)
        if fmt == "ordered":
            count = num_map.next_count(num_id, ilvl)
            marker = f"{count}."
        else:
            marker = "-"
    else:
        marker = "-"

    text = render_runs(para, rel_map)
    return f"{indent}{marker} {text}"


def has_num_pr(para) -> bool:
    """Return True if this paragraph has list numbering properties."""
    num_id, _ = _get_num_pr(para)
    return bool(num_id)
