"""Restore from every kind of backup (SPEC §1.1 rule 3, phase 5 hardening)."""

from docx import Document
from reportlab.pdfgen import canvas

from binder_builder.engine import pdfdoc, sync
from binder_builder.engine.session import BinderSession

from .helpers import do_sync

ADMIN = "A_Administration"


def test_every_backup_kind_is_restorable(binder, tmp_path):
    r = do_sync(binder)
    assert sync.backup_info(binder, r.backup)["kind"] == "sync"
    s = BinderSession(binder)
    src = tmp_path / "other.docx"
    d = Document()
    d.add_paragraph("replacement content")
    d.save(str(src))
    target = binder / ADMIN / "02_Organizational_Chart.docx"
    before = sync.file_hash(target)
    s.fill_from_file("org-chart", src, replace=True)
    fill = next(b for b in s.backups() if b.endswith("_fill"))
    info = s.backup_info(fill)
    assert info["kind"] == "fill" and info["files"] == 1 and info["restorable"]
    s.restore(fill)
    assert sync.file_hash(target) == before
    safety = next(b for b in s.backups() if b.endswith("_before-restore"))
    assert s.backup_info(safety)["kind"] == "before-restore" and s.backup_info(safety)["files"] == 1
    s.close()


def test_restore_first_sync_returns_to_unsynced_state(binder):
    from .test_sync import _all_files

    before = {rel: sync.file_hash(binder / rel) for rel in _all_files(binder)}
    r = do_sync(binder)
    sync.restore(binder, r.backup)
    assert {rel: sync.file_hash(binder / rel) for rel in _all_files(binder)} == before
    assert sync.load_state(binder)["files"] == {}  # the pre-sync (empty) state is back too


def test_restore_brings_back_pdf_original(binder):
    do_sync(binder)
    stash = pdfdoc.source_path(binder, "signed-acknowledgement")
    first_original = stash.read_bytes()
    pdf = binder / "B_Policies_and_Procedures" / "02_Signed_Policy_Acknowledgement.pdf"
    c = canvas.Canvas(str(pdf), invariant=1)
    c.drawString(72, 700, "a newer scan")
    c.save()
    dropped_in = pdf.read_bytes()
    r = do_sync(binder)  # the new scan becomes the original; the old one goes into the backup
    assert stash.read_bytes() != first_original
    sync.restore(binder, r.backup)
    # back to just before that sync: the old original is kept again and the new scan is unprocessed
    assert stash.read_bytes() == first_original
    assert pdf.read_bytes() == dropped_in
    p = sync.plan(binder)
    assert [a.detail for a in p.changes] == ["header overlay -> B.02 (original kept in .binder/pdf_sources)"]


def test_claims_backup_is_restorable(tmp_path):
    import shutil
    from pathlib import Path

    from binder_builder.engine import importer, refs

    legacy = tmp_path / "legacy"
    shutil.copytree(Path(__file__).parent / "fixtures" / "legacy_binder", legacy)
    importer.accept(importer.analyze(legacy))
    mrp = legacy / ADMIN / "01_North_Star_MRP_Summary.docx"
    before = sync.file_hash(mrp)
    importer.apply_claims(legacy, importer.claim_scan(legacy))
    assert sync.file_hash(mrp) != before
    claims = next(b for b in sync.list_backups(legacy) if b.endswith("_claims"))
    assert sync.backup_info(legacy, claims)["kind"] == "claims"
    sync.restore(legacy, claims)
    assert sync.file_hash(mrp) == before
    assert not refs.scan_docx(mrp, "x", {}) == []  # the typed codes are plain text again
