"""Render inline (run-level) formatting: bold, italic, code, hyperlinks."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph
    from wordtomd.relationships import RelationshipMap

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_HYPERLINK = f"{{{_W_NS}}}hyperlink"
_W_R = f"{{{_W_NS}}}r"
_W_T = f"{{{_W_NS}}}t"
_W_RPR = f"{{{_W_NS}}}rPr"
_W_B = f"{{{_W_NS}}}b"
_W_I = f"{{{_W_NS}}}i"
_W_RFONTS = f"{{{_W_NS}}}rFonts"
_W_RSTYLE = f"{{{_W_NS}}}rStyle"

# Fonts that indicate inline code
_CODE_FONTS = {"courier new", "consolas", "lucida console", "monaco", "source code pro", "inconsolata"}

# Characters that must be escaped in regular Markdown text (not inside code spans)
_MD_ESCAPE_RE = re.compile(r"([\\`*_{}\[\]()#+\-.!|])")


def _escape_md(text: str) -> str:
    return _MD_ESCAPE_RE.sub(r"\\\1", text)


def _is_code_run(rpr_el) -> bool:
    """Return True if this run's formatting indicates inline code."""
    if rpr_el is None:
        return False
    # Check character style name
    style_el = rpr_el.find(_W_RSTYLE)
    if style_el is not None:
        val = style_el.get(f"{{{_W_NS}}}val", "")
        if "code" in val.lower():
            return True
    # Check font name
    fonts_el = rpr_el.find(_W_RFONTS)
    if fonts_el is not None:
        for attr in fonts_el.attrib.values():
            if attr.lower() in _CODE_FONTS:
                return True
    return False


def _render_run_element(r_el) -> str:
    """Render a single w:r element to a Markdown string."""
    rpr = r_el.find(_W_RPR)
    is_bold = rpr is not None and rpr.find(_W_B) is not None
    is_italic = rpr is not None and rpr.find(_W_I) is not None
    is_code = _is_code_run(rpr)

    # Collect text, preserving xml:space="preserve"
    text_parts = []
    for t_el in r_el.findall(_W_T):
        text_parts.append(t_el.text or "")
    text = "".join(text_parts)

    if not text:
        return ""

    if is_code:
        # Use backtick escaping: if text contains backticks, use double backticks
        if "`" in text:
            return f"`` {text} ``"
        return f"`{text}`"

    escaped = _escape_md(text)

    if is_bold and is_italic:
        return f"***{escaped}***"
    if is_bold:
        return f"**{escaped}**"
    if is_italic:
        return f"_{escaped}_"
    return escaped


def render_runs(paragraph, rel_map: "RelationshipMap") -> str:
    """Render all inline content of a paragraph to a Markdown string.

    Iterates the paragraph XML directly to capture w:hyperlink context
    that is lost when iterating paragraph.runs.
    """
    parts = []
    p_el = paragraph._p

    for child in p_el:
        tag = child.tag

        if tag == _W_HYPERLINK:
            # Collect text from inner runs
            inner_parts = []
            for r_el in child.findall(_W_R):
                inner_parts.append(_render_run_element(r_el))
            link_text = "".join(inner_parts)

            r_id = child.get(f"{{{_W_NS}}}id", "")
            url = rel_map.hyperlinks.get(r_id, "")
            if url and link_text:
                parts.append(f"[{link_text}]({url})")
            elif link_text:
                parts.append(link_text)

        elif tag == _W_R:
            parts.append(_render_run_element(child))

    return "".join(parts)
