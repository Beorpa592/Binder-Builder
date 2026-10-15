"""Sync screen (SPEC §7.2, §12): plan table with filters, warnings, Apply, Restore from backup."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from ..engine import sync
from ..engine.session import BinderSession

LABELS = {
    sync.MOVE: "Rename / move",
    sync.CREATE: "New file",
    sync.STAMP: "Header",
    sync.REFS: "References",
    sync.REGENERATE: "Regenerate",
    sync.CHECKLIST: "Checklist rows",
    sync.ORPHAN: "Orphan (left in place)",
    sync.WARN: "Warning",
}


class SyncDialog(tk.Toplevel):
    def __init__(self, master, app, session: BinderSession):
        super().__init__(master)
        self.app = app
        self.s = session
        self.plan: sync.Plan | None = None
        self.title("Sync")
        self.transient(master)
        self.geometry("1050x620")

        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        self.summary = ttk.Label(outer, text="Checking the binder...", style="Heading.TLabel")
        self.summary.pack(anchor="w")
        ttk.Label(outer, text="Nothing changes until you click Apply. Every file is backed up first, and "
                              "author-written text is verified unchanged after every write.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 8))

        self.filter_frame = ttk.Frame(outer)
        self.filter_frame.pack(fill="x")
        self.filters: dict[str, tk.BooleanVar] = {}

        frame = ttk.Frame(outer)
        frame.pack(fill="both", expand=True, pady=8)
        self.table = ttk.Treeview(frame, columns=("action", "file", "detail"), show="headings")
        for col, text, width in (("action", "Action", 150), ("file", "File", 380), ("detail", "Detail", 460)):
            self.table.heading(col, text=text)
            self.table.column(col, width=width, stretch=col != "action")
        self.table.tag_configure("warn", foreground="#9C5700")
        self.table.tag_configure("orphan", foreground="#777777")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=sb.set)
        self.table.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Restore from backup...", command=self._restore,
                   state="disabled" if session.read_only else "normal").pack(side="left")
        ttk.Button(buttons, text="Close", command=self.destroy).pack(side="right")
        self.apply_btn = ttk.Button(buttons, text="Apply", style="Accent.TButton", command=self._apply,
                                    state="disabled")
        self.apply_btn.pack(side="right", padx=8)
        ttk.Button(buttons, text="Check again", command=self._replan).pack(side="right")
        if session.read_only:
            ttk.Label(buttons, text="Read-only: someone else holds the lock.", style="Muted.TLabel").pack(side="right", padx=12)
        self.grab_set()
        self._replan()

    def _replan(self) -> None:
        self.summary.config(text="Checking the binder...")
        self.apply_btn.config(state="disabled")
        self.app.run_task(self.s.plan, self._show, busy=None, on_error=self._error)

    def _error(self, e: Exception) -> None:
        if self.winfo_exists():
            self.summary.config(text=f"Could not plan the sync: {e}")

    def _show(self, plan: sync.Plan) -> None:
        if not self.winfo_exists():
            return
        self.plan = plan
        counts: dict[str, int] = {}
        for a in plan.actions:
            counts[a.kind] = counts.get(a.kind, 0) + 1
        for w in self.filter_frame.winfo_children():
            w.destroy()
        ttk.Label(self.filter_frame, text="Show:").pack(side="left", padx=(0, 6))
        for kind in sync.CHANGE_KINDS + (sync.ORPHAN, sync.WARN):
            if kind not in counts:
                continue
            var = self.filters.setdefault(kind, tk.BooleanVar(value=True))
            ttk.Checkbutton(self.filter_frame, text=f"{LABELS[kind]} ({counts[kind]})", variable=var,
                            command=self._fill).pack(side="left", padx=4)
        if plan.is_empty:
            self.summary.config(text="Everything is in sync. Nothing to apply.")
        else:
            self.summary.config(text=f"{len(plan.changes)} change(s) to apply, {len(plan.orphans)} orphan file(s), "
                                     f"{len(plan.warnings)} warning(s)")
        self.apply_btn.config(state="normal" if plan.changes and not self.s.read_only else "disabled")
        self._fill()

    def _fill(self) -> None:
        self.table.delete(*self.table.get_children())
        if not self.plan:
            return
        for a in self.plan.actions:
            var = self.filters.get(a.kind)
            if var is not None and not var.get():
                continue
            tag = "warn" if a.kind == sync.WARN else "orphan" if a.kind == sync.ORPHAN else ""
            self.table.insert("", "end", values=(LABELS[a.kind], a.path, a.detail), tags=(tag,))

    def _apply(self) -> None:
        if not self.plan or not self.plan.changes:
            return
        if not messagebox.askyesno("Apply sync", f"Apply {len(self.plan.changes)} change(s)?\n\n"
                                                  "Every file touched is backed up first.", parent=self):
            return
        self.apply_btn.config(state="disabled")
        self.summary.config(text="Applying...")
        plan, session = self.plan, self.s  # the worker must not hold a reference to this dialog
        self.app.run_task(lambda: session.apply(plan), self._applied, busy=None, on_error=self._apply_error)

    def _applied(self, result: sync.ApplyResult) -> None:
        if not self.winfo_exists():
            return
        if result.failures:
            messagebox.showwarning(
                "Sync finished with problems",
                "Some files were not changed:\n\n" + "\n".join(result.failures[:15]) +
                ("\n..." if len(result.failures) > 15 else "") +
                "\n\nEverything else was applied. Close any open files and sync again.", parent=self)
        else:
            messagebox.showinfo("Sync complete", f"Applied {len(result.applied)} update(s).\n"
                                                 f"Backup: {result.backup or 'none needed'}", parent=self)
        self._replan()

    def _apply_error(self, e: Exception) -> None:
        messagebox.showerror("Sync failed", str(e), parent=self)
        self._replan()

    def _restore(self) -> None:
        backups = self.s.backups()
        if not backups:
            messagebox.showinfo("Restore", "There are no backups yet.", parent=self)
            return
        dlg = RestoreDialog(self, [self.s.backup_info(b) for b in backups])
        self.wait_window(dlg)
        if not dlg.choice:
            return
        if not messagebox.askyesno(
            "Restore from backup",
            f"Put every file from the {dlg.choice} sync back the way it was before that sync?\n\n"
            "The current versions are backed up first. Files that sync created are left in place.",
            parent=self,
        ):
            return
        restored = self.s.restore(dlg.choice)
        messagebox.showinfo("Restore", f"Restored {len(restored)} file(s).", parent=self)
        self._replan()


class RestoreDialog(tk.Toplevel):
    def __init__(self, master, backups: list[dict]):
        super().__init__(master)
        self.title("Restore from backup")
        self.transient(master)
        self.choice: str | None = None
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Pick the change to undo (newest first):").pack(anchor="w")
        self.lb = tk.Listbox(f, height=12, width=72)
        kinds = {"sync": "Sync", "claims": "Claimed references", "fill": "Filled from file",
                 "before-restore": "Before an earlier restore", "unknown": "Unknown (cannot restore)"}
        for b in backups:
            when = (b["time"] or b["name"])[:16].replace("T", " ")
            self.lb.insert("end", f"{when}   {kinds.get(b['kind'], b['kind'])}   {b['files']} file(s)")
        self.lb.pack(fill="both", expand=True, pady=6)
        self.lb.selection_set(0)
        self.backups = backups
        bf = ttk.Frame(f)
        bf.pack(fill="x")
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bf, text="Restore", command=self._ok).pack(side="right", padx=8)
        self.grab_set()

    def _ok(self) -> None:
        sel = self.lb.curselection()
        if sel and self.backups[sel[0]]["restorable"]:
            self.choice = self.backups[sel[0]]["name"]
        self.destroy()
