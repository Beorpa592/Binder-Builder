"""Print and export screen (SPEC §11, §12). The layout exists now; the outputs arrive in phase 4,
when PDF export through Word/Excel COM is built."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from ..engine.session import BinderSession


class PrintDialog(tk.Toplevel):
    def __init__(self, master, session: BinderSession):
        super().__init__(master)
        self.s = session
        self.title("Print and export")
        self.transient(master)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Outputs", style="Heading.TLabel").pack(anchor="w")
        self.opts = {}
        for key, text in (("pdf_each", "PDF of each document"), ("merged", "One merged PDF of the whole binder"),
                          ("manifest", "Print manifest for Packet Printer")):
            var = tk.BooleanVar(value=key == "manifest")
            ttk.Checkbutton(f, text=text, variable=var, state="disabled").pack(anchor="w")
            self.opts[key] = var
        ttk.Label(f, text="Which documents", style="Heading.TLabel").pack(anchor="w", pady=(12, 0))
        self.scope = tk.StringVar(value="all")
        for value, text in (("all", "All included documents"), ("changed", "Changed since last print"),
                            ("selected", "Selected in the tree")):
            ttk.Radiobutton(f, text=text, value=value, variable=self.scope, state="disabled").pack(anchor="w")
        ttk.Label(f, text="PDF export and printing arrive in the next version. Office files are already kept "
                          "in sync in place by Sync.", style="Muted.TLabel", wraplength=360).pack(anchor="w", pady=12)
        bf = ttk.Frame(f)
        bf.pack(fill="x")
        ttk.Button(bf, text="Close", command=self.destroy).pack(side="right")
        ttk.Button(bf, text="Send to Packet Printer", state="disabled").pack(side="right", padx=8)
