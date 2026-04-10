"""Parse word/numbering.xml to resolve list styles and track ordered-list counters."""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from typing import Dict, Optional, Tuple

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NUMBERING_PATH = "word/numbering.xml"

# numFmt values that represent ordered (numbered) lists
_ORDERED_FORMATS = {
    "decimal",
    "decimalZero",
    "upperRoman",
    "lowerRoman",
    "upperLetter",
    "lowerLetter",
    "ordinal",
    "cardinalText",
    "ordinalText",
    "hex",
    "chicago",
    "ideographDigital",
    "japaneseCounting",
    "aiueo",
    "iroha",
    "decimalFullWidth",
    "decimalHalfWidth",
    "japaneseLegal",
    "japaneseDigitalTenThousand",
    "decimalEnclosedCircle",
    "decimalFullWidth2",
    "aiueoFullWidth",
    "irohaFullWidth",
    "decimalZero",
    "ganada",
    "chosung",
    "decimalEnclosedFullstop",
    "decimalEnclosedParen",
    "decimalEnclosedCircleChinese",
    "ideographEnclosedCircle",
    "ideographTraditional",
    "ideographZodiac",
    "ideographZodiacTraditional",
    "taiwaneseCounting",
    "ideographLegalTraditional",
    "taiwaneseCountingThousand",
    "taiwaneseDigital",
    "chineseCounting",
    "chineseLegalSimplified",
    "chineseCountingThousand",
    "koreanDigital",
    "koreanCounting",
    "koreanLegal",
    "koreanDigital2",
    "vietnameseCounting",
    "numberInDash",
    "russianLower",
    "russianUpper",
    "none",
}

# These are definitely bullet-style
_BULLET_FORMATS = {"bullet"}


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


class NumberingMap:
    def __init__(self) -> None:
        # format_map[(numId, ilvl)] -> "bullet" | "ordered"
        self._format_map: Dict[Tuple[str, int], str] = {}
        # counters[(numId, ilvl)] -> current count (1-based)
        self._counters: Dict[Tuple[str, int], int] = {}
        # Track last numId to detect list context switches
        self._last_num_id: Optional[str] = None

    @classmethod
    def from_docx(cls, docx_path: str) -> "NumberingMap":
        obj = cls()
        try:
            with zipfile.ZipFile(docx_path, "r") as zf:
                if _NUMBERING_PATH not in zf.namelist():
                    return obj
                xml_bytes = zf.read(_NUMBERING_PATH)
                root = ET.fromstring(xml_bytes)
        except (zipfile.BadZipFile, ET.ParseError, KeyError):
            return obj

        # Step 1: build abstractNumId -> {ilvl -> format}
        abstract_map: Dict[str, Dict[int, str]] = {}
        for abs_num in root.findall(_w("abstractNum")):
            abs_id = abs_num.get(_w("abstractNumId"), abs_num.get(f"{{{_W_NS}}}abstractNumId", ""))
            # Fallback: the attribute may be without namespace
            if not abs_id:
                abs_id = abs_num.get("w:abstractNumId", "")
            level_map: Dict[int, str] = {}
            for lvl in abs_num.findall(_w("lvl")):
                ilvl_str = lvl.get(_w("ilvl"), lvl.get(f"{{{_W_NS}}}ilvl", "0"))
                try:
                    ilvl = int(ilvl_str)
                except ValueError:
                    ilvl = 0
                num_fmt_el = lvl.find(_w("numFmt"))
                if num_fmt_el is not None:
                    fmt_val = num_fmt_el.get(_w("val"), num_fmt_el.get(f"{{{_W_NS}}}val", "bullet"))
                    level_map[ilvl] = "ordered" if fmt_val in _ORDERED_FORMATS and fmt_val != "none" and fmt_val != "bullet" else "bullet"
                else:
                    level_map[ilvl] = "bullet"
            # abstractNumId attribute can be stored directly on the element
            abs_id_attr = abs_num.get(f"{{{_W_NS}}}abstractNumId")
            if abs_id_attr:
                abstract_map[abs_id_attr] = level_map
            else:
                # Try without namespace
                for attr_name, attr_val in abs_num.attrib.items():
                    if "abstractNumId" in attr_name:
                        abstract_map[attr_val] = level_map
                        break

        # Step 2: resolve numId -> abstractNumId
        for num_el in root.findall(_w("num")):
            num_id = None
            abs_ref = None
            for attr_name, attr_val in num_el.attrib.items():
                if "numId" in attr_name:
                    num_id = attr_val
            abs_num_id_el = num_el.find(_w("abstractNumId"))
            if abs_num_id_el is not None:
                for attr_name, attr_val in abs_num_id_el.attrib.items():
                    if "val" in attr_name:
                        abs_ref = attr_val
            if num_id and abs_ref and abs_ref in abstract_map:
                for ilvl, fmt in abstract_map[abs_ref].items():
                    obj._format_map[(num_id, ilvl)] = fmt

        return obj

    def get_format(self, num_id: str, ilvl: int) -> str:
        """Return 'bullet' or 'ordered', defaulting to 'bullet' if unknown."""
        return self._format_map.get((num_id, ilvl), "bullet")

    def next_count(self, num_id: str, ilvl: int) -> int:
        """Increment and return the counter for this (numId, ilvl).
        Resets deeper levels when a shallower level is incremented."""
        # Reset deeper levels when this level is used
        keys_to_reset = [k for k in self._counters if k[0] == num_id and k[1] > ilvl]
        for k in keys_to_reset:
            del self._counters[k]

        key = (num_id, ilvl)
        self._counters[key] = self._counters.get(key, 0) + 1
        return self._counters[key]

    def reset(self, num_id: str) -> None:
        """Reset all counters for a given numId."""
        keys_to_reset = [k for k in self._counters if k[0] == num_id]
        for k in keys_to_reset:
            del self._counters[k]
