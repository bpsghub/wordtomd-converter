"""Render Word tables as GitHub-Flavored Markdown pipe tables."""

from __future__ import annotations

from typing import List, TYPE_CHECKING

if TYPE_CHECKING:
    from wordtomd.relationships import RelationshipMap

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_VMERGE = f"{{{_W_NS}}}vMerge"
_W_GRIDSPAN = f"{{{_W_NS}}}gridSpan"
_W_TCPR = f"{{{_W_NS}}}tcPr"
_W_VAL = f"{{{_W_NS}}}val"


def _cell_text(cell, rel_map: "RelationshipMap") -> str:
    """Render all paragraphs in a cell to a single string."""
    from wordtomd.renderers.inline import render_runs

    parts = []
    for para in cell.paragraphs:
        text = render_runs(para, rel_map).strip()
        if text:
            parts.append(text)
    return " ".join(parts)


def _get_grid_span(cell) -> int:
    tc = cell._tc
    tcpr = tc.find(_W_TCPR)
    if tcpr is None:
        return 1
    gs_el = tcpr.find(_W_GRIDSPAN)
    if gs_el is None:
        return 1
    try:
        return int(gs_el.get(_W_VAL, "1"))
    except ValueError:
        return 1


def _is_vmerge_continuation(cell) -> bool:
    """Return True if this cell is a continuation of a vertical merge (not the top cell)."""
    tc = cell._tc
    tcpr = tc.find(_W_TCPR)
    if tcpr is None:
        return False
    vm = tcpr.find(_W_VMERGE)
    if vm is None:
        return False
    # If w:vMerge has no val or val != "restart", it's a continuation
    val = vm.get(_W_VAL, "")
    return val != "restart"


def render_table(table, rel_map: "RelationshipMap") -> List[str]:
    """Return a list of Markdown lines representing a GFM pipe table."""
    if not table.rows:
        return []

    # Build grid: expand horizontal spans, mark vertical merge continuations
    # First pass: collect raw cells per row
    raw_rows = []
    for row in table.rows:
        raw_rows.append(list(row.cells))

    if not raw_rows:
        return []

    # Determine column count
    col_count = max(len(r) for r in raw_rows)

    # Build cell text grid (col_count columns per row)
    grid: List[List[str]] = []
    for r_idx, row_cells in enumerate(raw_rows):
        row_data: List[str] = []
        c_idx = 0
        for cell in row_cells:
            span = _get_grid_span(cell)
            if _is_vmerge_continuation(cell):
                # Copy text from the row above if available
                text = grid[r_idx - 1][c_idx] if r_idx > 0 and c_idx < len(grid[r_idx - 1]) else ""
            else:
                text = _cell_text(cell, rel_map)
            # Escape pipe characters and replace newlines
            text = text.replace("|", r"\|").replace("\n", "<br>")
            row_data.append(text)
            # Fill spanned columns with empty (already occupied visually)
            for _ in range(span - 1):
                row_data.append("")
            c_idx += span
        # Pad row to col_count
        while len(row_data) < col_count:
            row_data.append("")
        grid.append(row_data[:col_count])

    if not grid:
        return []

    # Build Markdown table lines
    def make_row(cells: List[str]) -> str:
        return "| " + " | ".join(cells) + " |"

    header = make_row(grid[0])
    separator = "| " + " | ".join(["---"] * col_count) + " |"
    lines = [header, separator]
    for row_data in grid[1:]:
        lines.append(make_row(row_data))

    return lines
