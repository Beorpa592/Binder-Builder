"""Start screen: recent binders, open, new, import; version and update notice (SPEC §12)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from .. import __version__


class StartView(ttk.Frame):
    def __init__(self, app):
        super().__init__(app, padding=32)
        self.app = app

        ttk.Label(self, text="Binder Builder", style="Title.TLabel").pack(anchor="w")
        ttk.Label(self, text="Keep every document in a binder in step with its Table of Contents.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 24))

        actions = ttk.Frame(self)
        actions.pack(anchor="w", pady=(0, 24))
        ttk.Button(actions, text="Open binder folder...", style="Accent.TButton",
                   command=app.choose_and_open).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="New binder...", command=app.new_binder).pack(side="left", padx=(0, 8))
        imp = ttk.Button(actions, text="Import existing folder...", command=self._import)
        imp.pack(side="left")

        ttk.Label(self, text="Recent binders", style="Heading.TLabel").pack(anchor="w")
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, pady=(4, 0))
        self.recent = tk.Listbox(frame, height=10, activestyle="none", font=("Segoe UI", 10))
        self.recent.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.recent.yview)
        sb.pack(side="right", fill="y")
        self.recent.config(yscrollcommand=sb.set)
        self._paths = app.recent.load()
        for p in self._paths:
            self.recent.insert("end", f"{Path(p).name}    ({p})")
        if not self._paths:
            self.recent.insert("end", "No recent binders yet. Open a folder or create a new binder.")
        self.recent.bind("<Double-Button-1>", self._open_selected)
        self.recent.bind("<Return>", self._open_selected)
        ttk.Label(self, text="Double-click a binder to open it.", style="Muted.TLabel").pack(anchor="w", pady=4)

        footer = ttk.Frame(self)
        footer.pack(fill="x", pady=(16, 0))
        ttk.Label(footer, text=f"Version {__version__}", style="Muted.TLabel").pack(side="left")
        self.update_frame = ttk.Frame(footer)
        self.update_frame.pack(side="right")
        if app.update_info:
            self.show_update(app.update_info)

    def _open_selected(self, _event=None) -> None:
        sel = self.recent.curselection()
        if sel and self._paths:
            self.app.open_binder(self._paths[sel[0]])

    def _import(self) -> None:
        from .import_wizard import run_import_wizard

        run_import_wizard(self.app)

    def show_update(self, info) -> None:
        for w in self.update_frame.winfo_children():
            w.destroy()
        ttk.Label(self.update_frame, text=f"Version {info.version} is available.").pack(side="left", padx=8)
        ttk.Button(self.update_frame, text="See what's new / update...", command=self.app.offer_update).pack(side="left")

    def set_busy(self, text: str) -> None:
        pass
