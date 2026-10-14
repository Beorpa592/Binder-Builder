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
