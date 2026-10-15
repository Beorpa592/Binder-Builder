"""Insert Reference (SPEC §6.1, §12): write a live reference at the cursor of the open Word
document, or into the active Excel cell, through COM."""

from __future__ import annotations

from pathlib import Path

from ..engine import excel as xl
from ..engine.word import TAG_REF_PREFIX

WD_CONTENT_CONTROL_TEXT = 1


class NoActiveDocument(Exception):
    pass


def insert_word(selection, node_id: str, code: str):
    """Insert a locked plain-text content control tagged bb:ref:<id> showing the code."""
    rng = selection.Range
    cc = rng.ContentControls.Add(WD_CONTENT_CONTROL_TEXT, rng)
    cc.Tag = f"{TAG_REF_PREFIX}{node_id}"
    cc.Title = "Binder reference"
    cc.Range.Text = code
    cc.LockContents = True
    return cc


def insert_excel(cell, node_id: str) -> None:
    cell.Formula = f"={xl.name_for(node_id)}"


def _active(prog_id: str):
    import win32com.client

    try:
        return win32com.client.GetActiveObject(prog_id)
    except Exception:  # noqa: BLE001 - not running
        return None


def insert_into_active(binder_folder: Path, node_id: str, code: str) -> str:
    """Insert into whichever Office document is active. Returns a message for the user."""
    folder = str(Path(binder_folder).resolve()).lower()
    word = _active("Word.Application")
    if word is not None and word.Documents.Count:
        doc = word.ActiveDocument
        insert_word(word.Selection, node_id, code)
        note = "" if str(doc.FullName).lower().startswith(folder) else \
            " (that document is not in this binder, so the reference won't be kept up to date)"
        return f"Inserted reference {code} into {doc.Name}{note}. Save the document in Word."
    excel = _active("Excel.Application")
    if excel is not None and excel.Workbooks.Count:
        wb = excel.ActiveWorkbook
        insert_excel(excel.ActiveCell, node_id)
        return (f"Inserted ={xl.name_for(node_id)} into {wb.Name}. It shows {code} once the workbook has been "
                "synced (the name is defined by sync). Save the workbook in Excel.")
    raise NoActiveDocument("Open the Word document or Excel workbook and put the cursor where the reference goes.")
