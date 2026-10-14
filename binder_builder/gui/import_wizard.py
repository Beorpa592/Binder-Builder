"""Import wizard (SPEC §9.1, §9.2, §12): pick folder -> pairing table -> TOC disagreements ->
claim references -> first sync plan. Nothing is written until "Accept and import"."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..engine import importer
from ..engine.session import BinderSession

CHECK_ON, CHECK_OFF = "☑", "☐"


def run_import_wizard(app) -> None:
    folder = filedialog.askdirectory(parent=app, title="Pick the existing binder folder to import", mustexist=True)
    if not folder:
        return
    if BinderSession.is_binder(folder):
        if messagebox.askyesno("Import", "This folder is already a Binder Builder binder. Open it instead?",
                               parent=app):
            app.open_binder(folder)
        return
    ImportWizard(app, Path(folder))


class ImportWizard(tk.Toplevel):
    def __init__(self, app, folder: Path):
        super().__init__(app)
        self.app = app
        self.folder = folder
        self.plan: importer.ImportPlan | None = None
        self.hits = []
        self.approved: set[str] = set()
        self.title(f"Import {folder.name}")
        self.geometry("1100x640")
        self.transient(app)
        self.body = ttk.Frame(self, padding=12)
        self.body.pack(fill="both", expand=True)
        self.nav = ttk.Frame(self, padding=(12, 0, 12, 12))
        self.nav.pack(fill="x")
        self._step_label("Reading the Table of Contents and every file...")
        folder_ = folder  # the worker must not hold a reference to this window
        app.run_task(lambda: importer.analyze(folder_), self._show_pairing, busy=None, on_error=self._failed)
        self.grab_set()

    # ---- scaffolding --------------------------------------------------------------
    def _clear(self) -> None:
        for frame in (self.body, self.nav):
            for w in frame.winfo_children():
                w.destroy()

    def _step_label(self, text: str, hint: str = "") -> None:
        self._clear()
        ttk.Label(self.body, text=text, style="Heading.TLabel").pack(anchor="w")
        if hint:
            ttk.Label(self.body, text=hint, style="Muted.TLabel", wraplength=1000).pack(anchor="w", pady=(0, 8))

    def _failed(self, e: Exception) -> None:
        if self.winfo_exists():
            messagebox.showerror("Import", f"Could not read the folder:\n\n{e}", parent=self)
            self.destroy()

    def _table(self, columns, widths, height=16) -> ttk.Treeview:
        frame = ttk.Frame(self.body)
        frame.pack(fill="both", expand=True)
        tv = ttk.Treeview(frame, columns=columns, show="headings", height=height)
        for c, w in zip(columns, widths):
            tv.heading(c, text=c)
            tv.column(c, width=w, stretch=c == columns[-1])
        sb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return tv

    # ---- step 2: pairing table ----------------------------------------------------------
    def _show_pairing(self, plan: importer.ImportPlan) -> None:
        if not self.winfo_exists():
            return
        self.plan = plan
        self._step_label(
            "Step 1 of 3: check the pairing",
            "Each Table of Contents entry is paired with the file that holds it. Matched by header code, then by "
            "file number and title, then by title alone. Fix a pair by selecting the row and choosing a file. "
            "Nothing is written yet.")
        self.pair_tv = self._table(("Code", "Title", "Matched by", "File"), (80, 300, 100, 500))
        self._fill_pairs()
        if plan.warnings:
            ttk.Label(self.body, text="\n".join(plan.warnings), foreground="#9C5700").pack(anchor="w", pady=4)
        ttk.Button(self.nav, text="Pair with file...", command=self._repair).pack(side="left")
        ttk.Button(self.nav, text="Unpair", command=self._unpair).pack(side="left", padx=6)
        self.unmatched_lbl = ttk.Label(self.nav, text="", style="Muted.TLabel")
        self.unmatched_lbl.pack(side="left", padx=12)
        ttk.Button(self.nav, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(self.nav, text="Next", style="Accent.TButton", command=self._show_disagreements).pack(
            side="right", padx=8)
        self._update_unmatched()

    def _fill_pairs(self) -> None:
        tv = self.pair_tv
        tv.delete(*tv.get_children())
        for p in self.plan.pairs.values():
            tv.insert("", "end", iid=p.entity_id,
                      values=(p.code or "", p.title, p.method or "-", p.file or "(no file: a blank one will be made)"))
        for f in self.plan.unmatched_files:
            tv.insert("", "end", values=("", "", "unmatched", f), tags=("unmatched",))
        tv.tag_configure("unmatched", foreground="#9C5700")

    def _update_unmatched(self) -> None:
        n = len(self.plan.unmatched_files)
        self.unmatched_lbl.config(text=f"{n} file(s) not paired: they stay where they are and show as orphans"
                                  if n else "Every file is paired.")

    def _repair(self) -> None:
        sel = self.pair_tv.selection()
        if not sel or sel[0] not in self.plan.pairs:
            return
        choices = sorted(set(self.plan.unmatched_files) | {p.file for p in self.plan.pairs.values() if p.file})
        dlg = _Choose(self, f"File for {self.plan.pairs[sel[0]].title}", choices)
        self.wait_window(dlg)
        if dlg.choice:
            self.plan.set_pair(sel[0], dlg.choice)
            self._fill_pairs()
            self._update_unmatched()

    def _unpair(self) -> None:
        sel = self.pair_tv.selection()
        if sel and sel[0] in self.plan.pairs:
            self.plan.set_pair(sel[0], None)
            self._fill_pairs()
            self._update_unmatched()

    # ---- step 3: disagreements ----------------------------------------------------------
    def _show_disagreements(self) -> None:
        dis = self.plan.disagreements()
        self._step_label(
            "Step 2 of 3: where the TOC and the files disagree",
            "The Table of Contents wins: on the first sync these files get the TOC's code in their header and "
            "file name. After importing you can still change the tree before you sync. Nothing is written until "
            "you click Accept and import.")
        tv = self._table(("File", "What", "File says", "TOC says"), (420, 150, 120, 300), height=10)
        for d in dis:
            tv.insert("", "end", values=(d.file or d.entity_id, d.what, d.file_says, d.toc_says))
        if not dis:
            tv.insert("", "end", values=("No disagreements.", "", "", ""))
        ttk.Button(self.nav, text="Back", command=lambda: self._show_pairing(self.plan)).pack(side="left")
        ttk.Button(self.nav, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(self.nav, text="Accept and import", style="Accent.TButton", command=self._accept).pack(
            side="right", padx=8)

    def _accept(self) -> None:
        try:
            importer.accept(self.plan)
        except Exception as e:  # noqa: BLE001 - shown to the user
            messagebox.showerror("Import", str(e), parent=self)
            return
        self._step_label("Looking for references to claim...")
        folder = self.folder
        self.app.run_task(lambda: importer.claim_scan(folder), self._show_claims, busy=None, on_error=self._failed)

    # ---- step 4: claim references ----------------------------------------------------------
    def _show_claims(self, hits) -> None:
        if not self.winfo_exists():
            return
        self.hits = [h for h in hits if h.target_id]
        self.approved = {h.key for h in self.hits}
        self._step_label(
            "Step 3 of 3: claim references",
            "These look like references to other documents. Approved ones become live references that update "
            "themselves when documents move; the text itself does not change. Unapproved ones stay as typed "
            "text and are listed in the import log. Click a row to approve or skip it.")
        self.claim_tv = self._table(("OK", "File", "Code", "Refers to", "Context"), (40, 320, 80, 120, 520))
        for h in self.hits:
            self.claim_tv.insert("", "end", iid=h.key, values=(CHECK_ON, h.file, h.code, h.target_id, h.context[:120]))
        self.claim_tv.bind("<ButtonRelease-1>", self._toggle_claim)
        ttk.Button(self.nav, text="Approve all", command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(self.nav, text="Approve none", command=lambda: self._set_all(False)).pack(side="left", padx=6)
        ttk.Button(self.nav, text="Approve all in this file", command=self._approve_file).pack(side="left")
        ttk.Button(self.nav, text="Claim approved and continue", style="Accent.TButton",
                   command=self._claim).pack(side="right")

    def _toggle_claim(self, event) -> None:
        row = self.claim_tv.identify_row(event.y)
        if not row:
            return
        on = row not in self.approved
        (self.approved.add if on else self.approved.discard)(row)
        self.claim_tv.set(row, "OK", CHECK_ON if on else CHECK_OFF)

    def _set_all(self, on: bool) -> None:
        for h in self.hits:
            (self.approved.add if on else self.approved.discard)(h.key)
            self.claim_tv.set(h.key, "OK", CHECK_ON if on else CHECK_OFF)

    def _approve_file(self) -> None:
        sel = self.claim_tv.selection()
        if not sel:
            return
        file = self.claim_tv.set(sel[0], "File")
        for h in self.hits:
            if h.file == file:
                self.approved.add(h.key)
                self.claim_tv.set(h.key, "OK", CHECK_ON)

    def _claim(self) -> None:
        approved = [h for h in self.hits if h.key in self.approved]
        rest = [h for h in self.hits if h.key not in self.approved]
        folder = self.folder
        self._step_label("Claiming references...")
        self.app.run_task(lambda: importer.apply_claims(folder, approved, rest), self._done, busy=None,
                          on_error=self._failed)

    def _done(self, result) -> None:
        if not self.winfo_exists():
            return
        skipped = "".join(f"\n  {h.file}: {h.code} ({why})" for h, why in result.skipped[:10])
        messagebox.showinfo(
            "Import complete",
            f"Claimed {len(result.claimed)} reference(s)." + (f"\nNot claimed:{skipped}" if skipped else "") +
            "\n\nNext: review the first sync. It takes over every header and applies the TOC's codes.",
            parent=self)
        self.destroy()
        self.app.open_binder(self.folder)
        view = self.app.view
        if view is not None and hasattr(view, "_sync"):
            view.after(200, view._sync)


class _Choose(tk.Toplevel):
    def __init__(self, master, title: str, choices: list[str]):
        super().__init__(master)
        self.title(title)
        self.transient(master)
        self.choice: str | None = None
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        self.lb = tk.Listbox(f, width=90, height=18)
        for c in choices:
            self.lb.insert("end", c)
        self.lb.pack(fill="both", expand=True)
        self.lb.bind("<Double-Button-1>", lambda e: self._ok())
        self.choices = choices
        bf = ttk.Frame(f)
        bf.pack(fill="x", pady=(8, 0))
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bf, text="Use this file", command=self._ok).pack(side="right", padx=8)
        self.grab_set()

    def _ok(self) -> None:
        sel = self.lb.curselection()
        if sel:
            self.choice = self.choices[sel[0]]
        self.destroy()
