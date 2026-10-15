"""Print and export screen (SPEC §11, §12): PDF per document, merged PDF, print manifest;
scope all / changed since last print / selected; Send to Packet Printer."""

from __future__ import annotations

import json
import os
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..engine.manifest import BINDER_DIR
from ..engine.session import BinderSession
from ..output import merge_pdf, print_manifest


def _settings_path() -> Path:
    from .app import app_data_dir

    return app_data_dir() / "settings.json"


def load_app_settings() -> dict:
    try:
        return json.loads(_settings_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_app_settings(data: dict) -> None:
    _settings_path().write_text(json.dumps(data, indent=2), encoding="utf-8")


def _export(folder: Path, options: merge_pdf.OutputOptions):
    """Runs on a worker thread: COM is initialised and used entirely inside it."""
    from ..office.export_pdf import OfficeExporter

    with OfficeExporter() as ex:
        return merge_pdf.export_binder(folder, options, ex)


class PrintDialog(tk.Toplevel):
    def __init__(self, master, session: BinderSession, app=None, selected: str | None = None):
        super().__init__(master)
        self.s = session
        self.app = app or master.app
        self.selected = selected
        self.result: merge_pdf.ExportResult | None = None
        self.title("Print and export")
        self.transient(master)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text="Outputs", style="Heading.TLabel").pack(anchor="w")
        self.pdf_each = tk.BooleanVar(value=True)
        self.merged = tk.BooleanVar(value=False)
        self.manifest = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="PDF of each document (in .binder\\pdf)", variable=self.pdf_each).pack(anchor="w")
        ttk.Checkbutton(f, text="One merged PDF of the whole selection, with bookmarks", variable=self.merged).pack(anchor="w")
        ttk.Checkbutton(f, text="Print manifest for Packet Printer", variable=self.manifest).pack(anchor="w")
        mode = "continuous: the TOC is page 1 and every page gets a running number" \
            if session.binder.page_numbers == "continuous" else "per document: each document keeps its Page X of Y"
        ttk.Label(f, text=f"Page numbering is {mode} (change it in Settings).", style="Muted.TLabel",
                  wraplength=420).pack(anchor="w", pady=(2, 0))

        ttk.Label(f, text="Which documents", style="Heading.TLabel").pack(anchor="w", pady=(12, 0))
        self.scope = tk.StringVar(value="all")
        ttk.Radiobutton(f, text="All included documents", value="all", variable=self.scope).pack(anchor="w")
        ttk.Radiobutton(f, text="Changed since they were last printed", value="changed", variable=self.scope).pack(anchor="w")
        sel_text = f"Selected in the tree: {session.tree.find(selected).text}" if selected else "Selected in the tree (nothing selected)"
        ttk.Radiobutton(f, text=sel_text, value="selected", variable=self.scope,
                        state="normal" if selected else "disabled").pack(anchor="w")

        self.status = ttk.Label(f, text="", wraplength=420)
        self.status.pack(anchor="w", pady=12)
        bf = ttk.Frame(f)
        bf.pack(fill="x")
        ttk.Button(bf, text="Close", command=self.destroy).pack(side="right")
        self.send_btn = ttk.Button(bf, text="Send to Packet Printer", command=self._send)
        self.send_btn.pack(side="right", padx=8)
        self.export_btn = ttk.Button(bf, text="Export", style="Accent.TButton", command=self._run)
        self.export_btn.pack(side="right")
        self.open_btn = ttk.Button(bf, text="Open PDF folder", command=self._open_folder)
        self.open_btn.pack(side="left")
        self.grab_set()

    # ---- helpers ------------------------------------------------------------------------
    def _selected_ids(self) -> list[str]:
        if not self.selected:
            return []
        node = self.s.tree.find(self.selected)
        ids = [n.id for n in node.walk()]
        return ids + [f"{i}#cover" for i in ids]

    def _output_options(self, manifest: bool | None = None) -> merge_pdf.OutputOptions:
        return merge_pdf.OutputOptions(per_document=self.pdf_each.get(), merged=self.merged.get(),
                                       manifest=self.manifest.get() if manifest is None else manifest,
                                       scope=self.scope.get(), selected=self._selected_ids())

    def _busy(self, text: str) -> None:
        self.status.config(text=text)
        self.export_btn.config(state="disabled")
        self.send_btn.config(state="disabled")

    def _idle(self) -> None:
        if self.winfo_exists():
            self.export_btn.config(state="normal")
            self.send_btn.config(state="normal")

    # ---- actions ------------------------------------------------------------------------
    def _run(self, then=None) -> None:
        folder, options = self.s.folder, self._output_options(manifest=True if then else None)
        self._busy("Exporting through Word and Excel... this can take a minute for a big binder.")
        self.app.run_task(lambda: _export(folder, options), lambda r: self._done(r, then), busy=None,
                          on_error=self._failed)

    def _failed(self, e: Exception) -> None:
        self._idle()
        if isinstance(e, merge_pdf.NeedsSync):
            messagebox.showinfo("Print and export", str(e), parent=self)
        elif isinstance(e, ImportError):
            messagebox.showerror("Print and export", "PDF export needs Microsoft Word and Excel (pywin32).", parent=self)
        else:
            messagebox.showerror("Print and export", f"Export failed:\n\n{e}", parent=self)
        if self.winfo_exists():
            self.status.config(text="")

    def _done(self, result: merge_pdf.ExportResult, then=None) -> None:
        self.result = result
        self._idle()
        if not self.winfo_exists():
            return
        parts = [f"{len(result.pdfs)} document(s) exported"]
        if result.merged:
            parts.append(f"merged PDF: {result.merged.name}")
        if result.manifest:
            parts.append("print manifest written")
        text = "; ".join(parts) + "."
        if result.warnings:
            text += "\n\n" + "\n".join(result.warnings[:8])
        self.status.config(text=text)
        if then:
            then(result)
        elif result.merged and messagebox.askyesno("Export", "Open the merged PDF?", parent=self):
            os.startfile(str(result.merged))  # type: ignore[attr-defined]

    def _open_folder(self) -> None:
        d = self.s.folder / BINDER_DIR / merge_pdf.PDF_DIR
        d.mkdir(parents=True, exist_ok=True)
        os.startfile(str(d))  # type: ignore[attr-defined]

    def _send(self) -> None:
        self._run(then=self._after_send)

    def _after_send(self, result: merge_pdf.ExportResult) -> None:
        if not result.documents:
            messagebox.showinfo("Send to Packet Printer", "Nothing to print in this selection.", parent=self)
            return
        settings = load_app_settings()
        exe = settings.get("packet_printer")
        if not exe or not Path(exe).is_file():
            if messagebox.askyesno("Send to Packet Printer",
                                   f"The print manifest is ready:\n{result.manifest}\n\n"
                                   "Locate Packet Printer so Binder Builder can open it directly next time?",
                                   parent=self):
                exe = filedialog.askopenfilename(parent=self, title="Locate Packet Printer",
                                                 filetypes=[("Programs", "*.exe")])
                if exe:
                    settings["packet_printer"] = exe
                    save_app_settings(settings)
        if exe and Path(exe).is_file():
            subprocess.Popen([exe, str(result.manifest)])
        if messagebox.askyesno("Mark as printed",
                               f"Mark these {len(result.documents)} document(s) as printed?\n\n"
                               "'Changed since last print' then starts counting from now.", parent=self):
            print_manifest.mark_printed(self.s.folder, result.documents)
            self.status.config(text=self.status.cget("text") + "\nMarked as printed.")
