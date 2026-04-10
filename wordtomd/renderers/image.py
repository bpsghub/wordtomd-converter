"""Extract embedded images from a docx and produce Markdown image references."""

from __future__ import annotations

import io
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from wordtomd.relationships import RelationshipMap


class ImageExtractor:
    def __init__(self, rel_map: "RelationshipMap", images_dir: Path, enabled: bool = True) -> None:
        self._rel_map = rel_map
        self._images_dir = images_dir
        self._enabled = enabled
        self._counter = 0
        self._dir_created = False

    def _ensure_dir(self) -> None:
        if not self._dir_created:
            self._images_dir.mkdir(parents=True, exist_ok=True)
            self._dir_created = True

    def extract(self, r_id: str, alt_text: str = "") -> str:
        """Return a Markdown image reference, extracting the file to disk."""
        self._counter += 1
        fallback_alt = alt_text.strip() or f"image-{self._counter}"

        if not self._enabled:
            return f"![{fallback_alt}]()"

        entry = self._rel_map.images.get(r_id)
        if entry is None:
            return f"<!-- image rId={r_id} not found -->"

        filename, raw_bytes = entry
        stem = Path(filename).stem
        suffix = Path(filename).suffix.lower()

        if suffix in (".emf", ".wmf"):
            # Attempt conversion via Pillow
            try:
                from PIL import Image  # type: ignore
                img = Image.open(io.BytesIO(raw_bytes))
                out_name = f"{stem}.png"
                self._ensure_dir()
                out_path = self._images_dir / out_name
                img.save(out_path, "PNG")
                rel_path = f"{self._images_dir.name}/{out_name}"
                return f"![{fallback_alt}]({rel_path})"
            except Exception:
                return f"<!-- image: {filename} (unsupported format, could not convert) -->"

        # Direct write for supported formats
        out_name = filename
        self._ensure_dir()
        out_path = self._images_dir / out_name
        out_path.write_bytes(raw_bytes)
        rel_path = f"{self._images_dir.name}/{out_name}"
        return f"![{fallback_alt}]({rel_path})"
