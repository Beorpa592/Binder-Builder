"""PDF export through Word and Excel COM (SPEC §11), so the PDF matches what Office prints.

Documents are opened read-only (a file someone else has open still exports) and closed
without saving. One Word and one Excel instance serve a whole batch.

Two hard-won details:
  * Word: `SaveAs2(FileFormat=wdFormatPDF)`, not `ExportAsFixedFormat`, which crashes some
    Word builds on the following `Close`.
  * A watchdog ends the app's own automation process (never a user's Word/Excel) if one
    export takes longer than `timeout` seconds. A hidden Office dialog would otherwise
    hang the app forever. The file is reported as failed and the batch continues with a
    fresh instance.
"""

from __future__ import annotations

import os
import shutil
import threading
from pathlib import Path

WD_FORMAT_PDF = 17
XL_TYPE_PDF = 0
DEFAULT_TIMEOUT = 180  # seconds per document

RPC_DISCONNECTED = (-2147417848, -2147023174, -2147023170)  # object disconnected, RPC server unavailable/failed


class ExportTimeout(Exception):
    pass


def _automation_pids(image: str) -> set[int]:
    """PIDs of Office processes started by automation (`/Automation -Embedding`), never the user's."""
    import win32com.client

    wmi = win32com.client.GetObject("winmgmts:")
    out = set()
    for p in wmi.ExecQuery(f"SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = '{image}'"):
        if p.CommandLine and "/automation" in p.CommandLine.lower():
            out.add(int(p.ProcessId))
    return out


def _kill(pid: int) -> None:
    try:
        os.kill(pid, 9)
    except OSError:
        pass


class OfficeExporter:
    """Context manager implementing the exporter interface used by `output.merge_pdf`."""

    def __init__(self, timeout: float = DEFAULT_TIMEOUT):
        self.timeout = timeout
        self._apps: dict[str, tuple[object, set[int]]] = {}

    def __enter__(self) -> "OfficeExporter":
        return self

    def __exit__(self, *exc) -> None:
        for kind in list(self._apps):
            self._close(kind)

    def _close(self, kind: str) -> None:
        app, pids = self._apps.pop(kind, (None, set()))
        if app is not None:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001 - already gone
                for pid in pids:
                    _kill(pid)

    def _app(self, kind: str):
        import pythoncom
        import win32com.client

        if kind in self._apps:
            return self._apps[kind][0]
        pythoncom.CoInitialize()
        image, prog = ("WINWORD.EXE", "Word.Application") if kind == "word" else ("EXCEL.EXE", "Excel.Application")
        before = _automation_pids(image)
        app = win32com.client.DispatchEx(prog)
        pids = _automation_pids(image) - before  # the process this exporter owns
        app.Visible = False
        if kind == "word":
            app.DisplayAlerts = 0
        else:
            app.DisplayAlerts = False
            app.ScreenUpdating = False
        self._apps[kind] = (app, pids)
        return app

    def export(self, src: Path, dest: Path) -> None:
        dest.parent.mkdir(parents=True, exist_ok=True)
        ext = src.suffix.lower()
        if ext == ".pdf":
            shutil.copy2(src, dest)
            return
        if ext not in (".docx", ".xlsx"):
            raise ValueError(f"cannot export {src.name} to PDF")
        kind = "word" if ext == ".docx" else "excel"
        for attempt in (1, 2):
            app = self._app(kind)
            timed_out = threading.Event()

            def fire():
                timed_out.set()
                for pid in self._apps.get(kind, (None, set()))[1]:
                    _kill(pid)

            timer = threading.Timer(self.timeout, fire)
            timer.daemon = True
            timer.start()
            try:
                (self._export_word if kind == "word" else self._export_excel)(app, src, dest)
                return
            except Exception as e:  # noqa: BLE001
                self._apps.pop(kind, None) if (timed_out.is_set() or _is_disconnect(e)) else None
                if timed_out.is_set():
                    raise ExportTimeout(f"{src.name}: Office did not finish within {self.timeout:.0f}s "
                                        "(a hidden Office dialog?); skipped") from e
                if not _is_disconnect(e) or attempt == 2:
                    raise
                # the Office process went away: try once more with a fresh one
            finally:
                timer.cancel()

    @staticmethod
    def _export_word(word, src: Path, dest: Path) -> None:
        doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        try:
            doc.SaveAs2(str(dest), FileFormat=WD_FORMAT_PDF)
        finally:
            doc.Close(0)  # wdDoNotSaveChanges

    @staticmethod
    def _export_excel(excel, src: Path, dest: Path) -> None:
        wb = excel.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0, AddToMru=False)
        try:
            excel.CalculateFull()  # resolve =bb_ references before printing
            try:
                wb.ExportAsFixedFormat(XL_TYPE_PDF, str(dest))
            except Exception as e:  # noqa: BLE001
                if not _is_nothing_to_print(e):
                    raise
                # A blank workbook still prints one page with its header. Give Excel something to
                # print in this read-only copy; it is closed without saving.
                for ws in wb.Worksheets:
                    if ws.Visible == -1 and excel.WorksheetFunction.CountA(ws.UsedRange) == 0:
                        ws.Range("A1").Value = " "
                wb.ExportAsFixedFormat(XL_TYPE_PDF, str(dest))
        finally:
            wb.Close(False)


def _is_disconnect(e: Exception) -> bool:
    return bool(getattr(e, "args", None)) and e.args[0] in RPC_DISCONNECTED


def _is_nothing_to_print(e: Exception) -> bool:
    return "find anything to print" in str(e) or "-2146827284" in str(e)
