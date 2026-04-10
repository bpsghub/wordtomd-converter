"""Post-process the generated Markdown lines for clean output."""

from __future__ import annotations

from typing import List


def clean_output(lines: List[str]) -> str:
    """Normalize whitespace and return the final Markdown string.

    - Strip trailing whitespace from every line
    - Collapse runs of 2+ consecutive blank lines to a single blank line
    - Remove leading and trailing blank lines
    - Append a single trailing newline
    """
    stripped = [line.rstrip() for line in lines]

    # Collapse consecutive blank lines
    result: List[str] = []
    prev_blank = False
    for line in stripped:
        is_blank = line == ""
        if is_blank and prev_blank:
            continue  # skip extra blank
        result.append(line)
        prev_blank = is_blank

    # Remove leading blank lines
    while result and result[0] == "":
        result.pop(0)

    # Remove trailing blank lines
    while result and result[-1] == "":
        result.pop()

    return "\n".join(result) + "\n"
