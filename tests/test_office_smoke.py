"""Open synced files in real Word and Excel through COM (skipped where Office or pywin32 is missing).

This is the "opens the files in Word" step of the phase 2 acceptance test, plus a
check that Excel evaluates the =bb_ references.
"""

import pytest

from .helpers import do_sync

win32 = pytest.importorskip("win32com.client")


def _office(app_name):
    try:
        app = win32.DispatchEx(app_name)
    except Exception as e:  # noqa: BLE001
        pytest.skip(f"{app_name} not available: {e}")
    app.Visible = False
    app.DisplayAlerts = 0 if app_name == "Word.Application" else False
    return app


def test_synced_files_open_in_word_and_excel(binder):
    do_sync(binder)
    word = _office("Word.Application")
    try:
        for rel in ("A_Administration/03_Top_10_List.docx", "A_Administration/00_Cover.docx", "Tabs.docx",
                    "00_Cover.docx", "A_Administration/06_Meeting_Minutes.docx"):
            d = word.Documents.Open(str(binder / rel), ReadOnly=True, AddToRecentFiles=False)
            try:
                hdr = d.Sections(1).Headers(1).Range.Text
                if rel.endswith("03_Top_10_List.docx"):
                    assert "Packet Ref # A.03" in hdr
                    ftr = d.Sections(1).Footers(1).Range
                    ftr.Fields.Update()
                    assert "Page 1 of 1" in ftr.Text
                assert d.ComputeStatistics(2) >= 1
            finally:
                d.Close(False)
    finally:
        word.Quit()

    excel = _office("Excel.Application")
    try:
        wb = excel.Workbooks.Open(str(binder / "0_Control_Section/01_Open_Points.xlsx"), ReadOnly=True)
        try:
            excel.CalculateFull()
            assert wb.Worksheets(1).Range("B2").Value == "A.03"
            assert wb.Worksheets("_binder").Visible == 0
        finally:
            wb.Close(False)
        wb = excel.Workbooks.Open(str(binder / "02_CKL.xlsx"), ReadOnly=True)
        try:
            excel.CalculateFull()
            assert wb.Worksheets(1).Range("A2").Value == "TOC"
        finally:
            wb.Close(False)
    finally:
        excel.Quit()


def test_claimed_references_open_in_word(tmp_path):
    import shutil
    from pathlib import Path

    from binder_builder.engine import importer, sync

    legacy = tmp_path / "legacy"
    shutil.copytree(Path(__file__).parent / "fixtures" / "legacy_binder", legacy)
    importer.accept(importer.analyze(legacy))
    hits = importer.claim_scan(legacy)
    importer.apply_claims(legacy, hits)
    assert sync.apply(sync.plan(legacy)).ok

    word = _office("Word.Application")
    try:
        d = word.Documents.Open(str(legacy / "A_Administration" / "01_North_Star_MRP_Summary.docx"),
                                ReadOnly=True, AddToRecentFiles=False)
        try:
            tags = sorted(d.ContentControls(i + 1).Tag for i in range(d.ContentControls.Count))
            assert tags == ["bb:ref:social-media-outreach-efforts", "bb:ref:top-10-list"]
            assert "Complete the Top 10 List (A.03) first" in d.Content.Text
        finally:
            d.Close(False)
    finally:
        word.Quit()


def test_real_export_merged_continuous(tmp_path):
    """Word/Excel COM export of a small binder into a merged, continuously numbered PDF."""
    pytest.importorskip("pythoncom")
    from pypdf import PdfReader

    from binder_builder.engine.session import BinderSession
    from binder_builder.office.export_pdf import OfficeExporter
    from binder_builder.output import merge_pdf

    s = BinderSession.create(tmp_path / "Small", "Small Binder")
    s.update_binder(page_numbers="continuous")
    a = s.add_section("Admin")
    s.add_document("Letter", "docx", a.id)
    s.add_document("Budget", "xlsx", "letter")
    assert s.apply(s.plan()).ok
    s.close()
    from binder_builder.engine import sync as _sync

    sources = {p: _sync.file_hash(p) for p in (tmp_path / "Small" / "A_Admin").iterdir()}
    try:
        with OfficeExporter() as ex:
            r = merge_pdf.export_binder(tmp_path / "Small", merge_pdf.OutputOptions(merged=True), ex)
    except Exception as e:  # noqa: BLE001
        if "Invalid class string" in str(e):
            pytest.skip("Office not installed")
        raise
    assert not r.warnings, r.warnings
    assert {p: _sync.file_hash(p) for p in sources} == sources  # exporting never modifies a document
    reader = PdfReader(str(r.merged))
    assert reader.page_labels[0] == "Cover" and reader.page_labels[1] == "1"
    texts = [p.extract_text() for p in reader.pages]
    assert "Packet Ref # TOC" in texts[1]
    assert texts[1].strip().endswith("1")
    assert any("Packet Ref # A.01" in t for t in texts)
    titles = [o.title for o in reader.outline if not isinstance(o, list)]
    assert titles[:4] == ["Cover", "Table of Contents", "Master Checklist", "A) Admin"]


def test_insert_reference_into_word_and_excel(binder):
    do_sync(binder)
    from docx import Document

    from binder_builder.engine import excel as xl
    from binder_builder.engine import sync, word as wd
    from binder_builder.office.insert_reference import insert_excel, insert_word

    target = binder / "A_Administration" / "02_Organizational_Chart.docx"
    app = _office("Word.Application")
    try:
        d = app.Documents.Open(str(target), AddToRecentFiles=False)
        try:
            app.Selection.EndKey(6)  # wdStory: cursor at the end of the document
            app.Selection.TypeText("See ")
            insert_word(app.Selection, "top-10-list", "A.03")
            d.Save()
        finally:
            d.Close(False)
    finally:
        app.Quit()
    refs = {h.target_id: h.current for h in wd.scan_refs(Document(str(target)), {})}
    assert refs == {"top-10-list": "A.03"}

    book = binder / "C_Campaign_Management" / "01_Campaign_Calendar.xlsx"
    app = _office("Excel.Application")
    try:
        wb = app.Workbooks.Open(str(book))
        try:
            wb.Worksheets(1).Range("D5").Select()
            insert_excel(app.ActiveCell, "top-10-list")
            app.CalculateFull()
            assert wb.Worksheets(1).Range("D5").Value == "A.03"
            wb.Save()
        finally:
            wb.Close(False)
    finally:
        app.Quit()

    # the inserted references follow a renumber like any other
    from .helpers import edit_tree

    m, tree = edit_tree(binder)
    tree.insert_after("org-chart", {"kind": "item", "title": "Press Kit", "ext": "docx"})
    m.save()
    do_sync(binder)
    refs = {h.target_id: h.current for h in wd.scan_refs(Document(str(target)), {})}
    assert refs == {"top-10-list": "A.04"}
    assert xl.resolve_name(book, "bb_top_10_list") == "A.04"
    assert sync.plan(binder).is_empty
