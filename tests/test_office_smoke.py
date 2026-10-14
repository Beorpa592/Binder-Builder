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
