"""Header overlay for PDF items (SPEC §5.3).

A PDF item gets the binder header drawn on every page and `Page n of N` in the
footer. The original is kept in `.binder/pdf_sources/<id>.source.pdf`, and the
binder's copy is always regenerated from it, so overlays never stack. If the
user drops a new PDF in place, it becomes the new original.

Invariant: same page count, and each stamped page still contains all of its
source page's text.
"""

from __future__ import annotations

import hashlib
import io
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

from .codes import HeaderSpec
from .manifest import BINDER_DIR

SOURCES = "pdf_sources"
MARGIN = 28  # points
FONT = "Helvetica"
SIZE = 8


def source_path(folder: Path, entity_id: str) -> Path:
    return folder / BINDER_DIR / SOURCES / f"{entity_id}.source.pdf"


def stamp_key(source_hash: str, spec: HeaderSpec) -> str:
    """Identifies one rendering: changes when the original or the header text changes."""
    return hashlib.sha256(f"{source_hash}|{spec!r}".encode("utf-8")).hexdigest()


def _overlay(width: float, height: float, spec: HeaderSpec, page: int, pages: int):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height), invariant=1)
    c.setFont(FONT, SIZE)
    y = height - MARGIN
    lines = list(spec.lines) or [""]
    for i, line in enumerate(lines):
        c.drawString(MARGIN, y, line)
        if i == 0 and spec.code is not None:
            c.drawRightString(width - MARGIN, y, f"{spec.ref_label} {spec.code}")
        y -= SIZE + 2
    left, center, right = spec.footer
    fy = MARGIN
    if left or center or right:
        c.drawString(MARGIN, fy + SIZE + 2, left)
        c.drawCentredString(width / 2, fy + SIZE + 2, center)
        c.drawRightString(width - MARGIN, fy + SIZE + 2, right)
    if spec.page_numbers:
        c.drawRightString(width - MARGIN, fy, f"Page {page} of {pages}")
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def stamp(source: Path, dest: Path, spec: HeaderSpec) -> None:
    reader = PdfReader(str(source))
    writer = PdfWriter()
    n = len(reader.pages)
    for i, page in enumerate(reader.pages, start=1):
        out = writer.add_page(page)
        box = out.mediabox
        out.merge_page(_overlay(float(box.width), float(box.height), spec, i, n))
    with open(dest, "wb") as f:
        writer.write(f)


def _texts(path: Path) -> list[str]:
    return [" ".join((p.extract_text() or "").split()) for p in PdfReader(str(path)).pages]


def invariant_ok(source: Path, stamped: Path) -> bool:
    src, out = _texts(source), _texts(stamped)
    return len(src) == len(out) and all(s in o for s, o in zip(src, out))
