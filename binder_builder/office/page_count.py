"""Page counts (SPEC §2, §11).

The authoritative count is taken from the PDF that Word/Excel exported (`pdf_pages`),
which is exactly what prints. `office_pages` asks Office directly without
exporting, for quick estimates in the GUI.
"""

from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader


def pdf_pages(path: Path) -> int:
    return len(PdfReader(str(path)).pages)


def office_pages(src: Path) -> int:
    import win32com.client

    ext = src.suffix.lower()
    if ext == ".pdf":
        return pdf_pages(src)
    if ext == ".docx":
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False, Visible=False)
            try:
                return int(doc.ComputeStatistics(2))  # wdStatisticPages
            finally:
                doc.Close(False)
        finally:
            word.Quit()
    if ext == ".xlsx":
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(src), ReadOnly=True)
            try:
                total = 0
                for ws in wb.Worksheets:
                    if ws.Visible == -1:  # xlSheetVisible
                        total += int(ws.PageSetup.Pages.Count)
                return total
            finally:
                wb.Close(False)
        finally:
            excel.Quit()
    raise ValueError(src.name)
