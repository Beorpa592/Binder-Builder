"""Per-document PDFs, the merged binder PDF, and the print manifest (SPEC §11).

    with OfficeExporter() as ex:
        result = export_binder(folder, OutputOptions(merged=True), ex)

Per-document PDFs go to `.binder/pdf/<same relative path>.pdf` and are re-exported only
when their source changed. The merged PDF runs in TOC order with a bookmark for every
section and document. The binder cover comes first and is not numbered. In continuous
mode the TOC is page 1, the running number is stamped bottom center on every page from
the TOC onward, and the TOC's Page column is filled from the real page counts (the TOC
is regenerated through a normal sync, then re-exported).
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Protocol

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

from ..engine import sync
from ..engine.codes import BINDER_COVER, CKL, ITEM, SECTION_COVER, TABS, TOC, Entity, Layout
from ..engine.manifest import BINDER_DIR, Manifest
from ..engine.tree import Node
from . import print_manifest

PDF_DIR = "pdf"


class Exporter(Protocol):
    def export(self, src: Path, dest: Path) -> None: ...


class NeedsSync(Exception):
    pass


@dataclass
class OutputOptions:
    per_document: bool = True
    merged: bool = False
    manifest: bool = True
    scope: str = "all"  # all | changed | selected
    selected: list[str] | None = None


@dataclass
class ExportResult:
    pdfs: dict[str, Path] = field(default_factory=dict)
    pages: dict[str, int] = field(default_factory=dict)
    start_pages: dict[str, int] = field(default_factory=dict)  # entity id -> running page (continuous)
    merged: Path | None = None
    manifest: Path | None = None
    documents: list[str] = field(default_factory=list)  # entity ids in the outputs, in order
    warnings: list[str] = field(default_factory=list)


def pdf_path(folder: Path, rel: str) -> Path:
    return folder / BINDER_DIR / PDF_DIR / Path(rel).with_suffix(".pdf")


def printable(layout: Layout) -> list[Entity]:
    """Documents that go into the physical binder, in order (tab labels print on tab stock)."""
    return [e for e in layout.entities if e.kind != TABS and e.included]


def _current_path(folder: Path, state: dict, e: Entity) -> Path:
    rel = state["files"].get(e.entity_id, {}).get("path") or e.rel_path
    return folder / rel


def _scope(folder: Path, docs: list[Entity], state: dict, options: OutputOptions) -> list[Entity]:
    if options.scope == "selected":
        wanted = set(options.selected or [])
        return [e for e in docs if e.entity_id in wanted or (e.node is not None and e.node.id in wanted)]
    if options.scope == "changed":
        return [e for e in docs
                if print_manifest.changed_since_last_print(folder, state, e.entity_id, _current_path(folder, state, e))]
    return docs


def _export_all(folder: Path, docs: list[Entity], state: dict, exporter: Exporter, result: ExportResult,
                progress: Callable[[str], None] | None) -> None:
    exported = state.setdefault("exported", {})
    for e in docs:
        src = _current_path(folder, state, e)
        if not src.is_file():
            result.warnings.append(f"{e.code or e.title}: no file to export ({src.name})")
            continue
        dest = pdf_path(folder, src.relative_to(folder).as_posix())
        h = sync.file_hash(src)
        if not (dest.is_file() and exported.get(e.entity_id) == h):
            if progress:
                progress(f"Exporting {src.name}...")
            try:
                exporter.export(src, dest)
            except Exception as ex:  # noqa: BLE001 - one bad file must not stop the binder
                result.warnings.append(f"{src.name}: export failed ({ex})")
                continue
            exported[e.entity_id] = h
        result.pdfs[e.entity_id] = dest
        result.pages[e.entity_id] = len(PdfReader(str(dest)).pages)


def _running_starts(docs: list[Entity], pages: dict[str, int]) -> tuple[dict[str, int], dict[str, int]]:
    """Running start page per entity (TOC = 1; the binder cover is unnumbered) and per code."""
    starts: dict[str, int] = {}
    n = None
    for e in docs:
        if e.entity_id not in pages:
            continue
        if e.kind == BINDER_COVER and n is None:
            continue
        if n is None:
            n = 1
        starts[e.entity_id] = n
        n += pages[e.entity_id]
    by_code = {}
    for e in docs:
        if e.entity_id in starts and e.code:
            by_code[e.code] = starts[e.entity_id]
            if e.kind == SECTION_COVER:
                by_code[e.code[:-3]] = starts[e.entity_id]  # section row "A" -> its cover page
    return starts, by_code


def _number_overlay(width: float, height: float, number: int):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height), invariant=1)
    c.setFont("Helvetica", 10)
    c.drawCentredString(width / 2, 14, str(number))
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def _outline_title(e: Entity) -> str:
    if e.kind == BINDER_COVER:
        return "Cover"
    if e.kind == TOC:
        return "Table of Contents"
    if e.kind == CKL:
        return "Master Checklist"
    if e.kind == SECTION_COVER:
        return f"{e.code[:-3]}) {e.title}"
    return f"{e.code} {e.title}"


def merge(docs: list[Entity], result: ExportResult, out: Path, continuous: bool) -> None:
    writer = PdfWriter()
    first_numbered = None
    outline: dict[str, object] = {}  # node id -> outline item
    for e in docs:
        pdf = result.pdfs.get(e.entity_id)
        if pdf is None:
            continue
        start = len(writer.pages)
        for page in PdfReader(str(pdf)).pages:
            writer.add_page(page)
        if first_numbered is None and e.kind != BINDER_COVER:
            first_numbered = start

        parent = None
        node: Node | None = e.node
        anc = node.parent if (node is not None and e.kind == ITEM) else (node.parent if node is not None else None)
        while anc is not None and parent is None:
            parent = outline.get(anc.id)
            if parent is None and anc.is_header:
                parent = outline[anc.id] = writer.add_outline_item(anc.text, start, parent=_outline_parent(anc, outline))
            anc = anc.parent
        item = writer.add_outline_item(_outline_title(e), start, parent=parent)
        if node is not None:
            outline[node.id] = item

    total = len(writer.pages)
    if first_numbered is None:
        first_numbered = total
    if continuous:
        for i in range(first_numbered, total):
            page = writer.pages[i]
            box = page.mediabox
            page.merge_page(_number_overlay(float(box.width), float(box.height), i - first_numbered + 1))
    if first_numbered > 0:
        writer.set_page_label(0, first_numbered - 1, prefix="Cover")
    if total > first_numbered:
        writer.set_page_label(first_numbered, total - 1, style="/D", start=1)
    writer.page_mode = "/UseOutlines"
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "wb") as f:
        writer.write(f)


def _outline_parent(header: Node, outline: dict):
    p = header.parent
    while p is not None:
        if p.id in outline:
            return outline[p.id]
        p = p.parent
    return None


def merged_path(folder: Path, binder_name: str) -> Path:
    safe = "".join(ch for ch in binder_name if ch not in '<>:"/\\|?*').strip() or "Binder"
    return folder / BINDER_DIR / PDF_DIR / f"{safe} (merged).pdf"


def export_binder(folder: Path | str, options: OutputOptions, exporter: Exporter,
                  progress: Callable[[str], None] | None = None) -> ExportResult:
    folder = Path(folder)
    result = ExportResult()
    plan = sync.plan(folder)
    if plan.changes:
        if options.merged:
            raise NeedsSync("Sync the binder first: the merged PDF and its page numbers must match the TOC.")
        result.warnings.append("The binder has unsynced changes; PDFs show the files as they are now.")
    manifest = Manifest.load(folder)
    continuous = manifest.binder.page_numbers == "continuous"
    state = sync.load_state(folder)
    docs = printable(plan.layout)
    target = _scope(folder, docs, state, options)
    result.documents = [e.entity_id for e in target]
    if not target:
        result.warnings.append("No documents in this selection.")
        return result
    changed_flags = {e.entity_id: print_manifest.changed_since_last_print(
        folder, state, e.entity_id, _current_path(folder, state, e)) for e in target}

    _export_all(folder, target, state, exporter, result, progress)
    sync.save_state(folder, state)

    if options.merged:
        if continuous:
            for _ in range(3):  # filling the Page column can change the TOC's own length; settle it
                starts, by_code = _running_starts(target, result.pages)
                state = sync.load_state(folder)
                if state.get("pages") == by_code:
                    break
                state["pages"] = by_code
                sync.save_state(folder, state)
                p = sync.plan(folder)  # regenerates the TOC with the Page column, nothing else
                if p.changes:
                    r = sync.apply(p)
                    result.warnings.extend(r.failures)
                state = sync.load_state(folder)
                _export_all(folder, target, state, exporter, result, progress)
                sync.save_state(folder, state)
            result.start_pages = _running_starts(target, result.pages)[0]
        out = merged_path(folder, manifest.binder.name)
        if progress:
            progress("Merging...")
        merge(target, result, out, continuous)
        result.merged = out

    if options.manifest:
        entries = []
        for e in target:
            if e.entity_id not in result.pdfs:
                continue
            entries.append({
                "path": str(_current_path(folder, state, e)),
                "pdf_path": str(result.pdfs[e.entity_id]),
                "code": e.code or "",
                "title": e.title,
                "pages": result.pages[e.entity_id],
                "changed_since_last_print": changed_flags[e.entity_id],
            })
        result.manifest = print_manifest.write(folder, entries)
    if not options.per_document:
        pass  # per-document PDFs are the merge's inputs; they stay in .binder/pdf either way
    return result
