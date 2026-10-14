"""Import wizard (SPEC §9.1, §12): pick folder -> pairing table -> TOC disagreements -> claim
references -> first sync plan. Built in phase 3; until then the Start screen explains that."""

from __future__ import annotations

from tkinter import messagebox


def run_import_wizard(app) -> None:
    messagebox.showinfo(
        "Import existing folder",
        "Importing an existing binder (reading its TOC, pairing files and claiming references) "
        "arrives in the next version.\n\nFor now, use 'New binder...' or 'Open binder folder...'.",
        parent=app,
    )
