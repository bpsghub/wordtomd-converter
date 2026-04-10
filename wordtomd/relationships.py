"""Parse word/_rels/document.xml.rels to extract hyperlinks and embedded images."""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from typing import Dict, Tuple

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
_REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

_RELS_PATH = "word/_rels/document.xml.rels"


@dataclass
class RelationshipMap:
    hyperlinks: Dict[str, str] = field(default_factory=dict)
    # rId -> (original_filename, raw_bytes)
    images: Dict[str, Tuple[str, bytes]] = field(default_factory=dict)

    @classmethod
    def from_docx(cls, docx_path: str) -> "RelationshipMap":
        obj = cls()
        try:
            with zipfile.ZipFile(docx_path, "r") as zf:
                if _RELS_PATH not in zf.namelist():
                    return obj
                rels_xml = zf.read(_RELS_PATH)
                root = ET.fromstring(rels_xml)
                for rel in root.findall(f"{{{_RELS_NS}}}Relationship"):
                    rId = rel.get("Id", "")
                    rtype = rel.get("Type", "")
                    target = rel.get("Target", "")

                    if rtype == _REL_HYPERLINK:
                        obj.hyperlinks[rId] = target

                    elif rtype == _REL_IMAGE:
                        # target is relative to word/, e.g. "media/image1.png"
                        zip_path = f"word/{target}" if not target.startswith("/") else target.lstrip("/")
                        try:
                            image_bytes = zf.read(zip_path)
                            filename = target.split("/")[-1]
                            obj.images[rId] = (filename, image_bytes)
                        except KeyError:
                            pass  # image missing from ZIP — skip silently
        except (zipfile.BadZipFile, ET.ParseError):
            pass
        return obj
