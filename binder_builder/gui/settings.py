"""Binder settings (SPEC §12): name, header lines with live sample, ref label, vars, footer with
presets, page number mode, tabs, cover lines, template folder."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from ..engine.codes import render_text
from ..engine.manifest import ManifestError
from ..engine.session import BinderSession, ReadOnlyError

SAMPLE_LAYOUT = ["{{org_short}}", "North Star Media", "Management Binder", "{{doc_title}}"]
FOOTER_PRESETS = {
    "(none)": ("", "", ""),
    "BMC revision block": ("Revision {{revision}}", "", "Revised {{revision_date}}"),
}
PRESET_VARS = {"BMC revision block": {"revision": "1", "revision_date": ""}}


class SettingsDialog(tk.Toplevel):
    def __init__(self, master, session: BinderSession):
        super().__init__(master)
        self.s = session
        b = session.binder
        self.title("Binder settings")
        self.transient(master)
        self.geometry("760x720")
        ro = "disabled" if session.read_only else "normal"

        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)
        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)

        # -- general / header
        g = ttk.Frame(nb, padding=12)
        nb.add(g, text="Header")
        g.columnconfigure(1, weight=1)
        self.name = tk.StringVar(value=b.name)
        ttk.Label(g, text="Binder name").grid(row=0, column=0, sticky="w")
        ttk.Entry(g, textvariable=self.name, state=ro).grid(row=0, column=1, sticky="we", pady=4)

        ttk.Label(g, text="Header lines").grid(row=1, column=0, sticky="nw", pady=4)
        hl = ttk.Frame(g)
        hl.grid(row=1, column=1, sticky="we")
        self.lines = tk.Text(hl, height=5, width=50, font=("Consolas", 10))
        self.lines.insert("1.0", "\n".join(b.header_lines))
        self.lines.pack(side="left", fill="x", expand=True)
        ttk.Button(hl, text="Use North Star\nsample", command=self._sample, state=ro).pack(side="left", padx=6)
        ttk.Label(g, text="One line per row. Placeholders: {{org_short}} {{binder_name}} {{doc_title}} {{doc_code}} "
                          "{{section_label}} and any variable below. The reference code always goes at the right "
                          "of line 1.", style="Muted.TLabel", wraplength=520).grid(row=2, column=1, sticky="w")

        self.ref_label = tk.StringVar(value=b.ref_label)
        ttk.Label(g, text="Reference label").grid(row=3, column=0, sticky="w")
        ttk.Entry(g, textvariable=self.ref_label, width=20, state=ro).grid(row=3, column=1, sticky="w", pady=4)

        ttk.Label(g, text="Variables").grid(row=4, column=0, sticky="nw", pady=4)
        self.vars_text = tk.Text(g, height=4, width=50, font=("Consolas", 10))
        self.vars_text.insert("1.0", "\n".join(f"{k} = {v}" for k, v in b.vars.items()))
        self.vars_text.grid(row=4, column=1, sticky="we")
        ttk.Label(g, text="name = value, one per line", style="Muted.TLabel").grid(row=5, column=1, sticky="w")

        ttk.Label(g, text="Sample header").grid(row=6, column=0, sticky="nw", pady=(12, 0))
        self.preview = tk.Label(g, justify="left", anchor="nw", font=("Consolas", 10), bg="white", relief="solid",
                                bd=1, padx=8, pady=6)
        self.preview.grid(row=6, column=1, sticky="we", pady=(12, 0))
        for w in (self.lines, self.vars_text):
            w.bind("<KeyRelease>", lambda e: self._update_preview())
        self.ref_label.trace_add("write", lambda *a: self._update_preview())
        self.name.trace_add("write", lambda *a: self._update_preview())

        # -- footer and pages
        f = ttk.Frame(nb, padding=12)
        nb.add(f, text="Footer && pages")
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Preset").grid(row=0, column=0, sticky="w")
        self.preset = tk.StringVar(value="(custom)")
        cb = ttk.Combobox(f, textvariable=self.preset, values=list(FOOTER_PRESETS), state="readonly", width=24)
        cb.grid(row=0, column=1, sticky="w", pady=4)
        cb.bind("<<ComboboxSelected>>", lambda e: self._apply_preset())
        self.footer = {k: tk.StringVar(value=getattr(b.footer, k)) for k in ("left", "center", "right")}
        for i, k in enumerate(("left", "center", "right"), start=1):
            ttk.Label(f, text=f"Footer {k}").grid(row=i, column=0, sticky="w")
            ttk.Entry(f, textvariable=self.footer[k], state=ro).grid(row=i, column=1, sticky="we", pady=2)
        ttk.Label(f, text="Blank means no footer text. 'Page X of Y' is always added.", style="Muted.TLabel").grid(
            row=4, column=1, sticky="w")
        ttk.Label(f, text="Page numbers").grid(row=5, column=0, sticky="nw", pady=(12, 0))
        self.pages = tk.StringVar(value=b.page_numbers)
        pf = ttk.Frame(f)
        pf.grid(row=5, column=1, sticky="w", pady=(12, 0))
        ttk.Radiobutton(pf, text="Per document (each document shows Page X of Y)", value="per_document",
                        variable=self.pages, state=ro).pack(anchor="w")
        ttk.Radiobutton(pf, text="Continuous (merged PDF numbered from the TOC onward)", value="continuous",
                        variable=self.pages, state=ro).pack(anchor="w")

        # -- tabs and cover
        t = ttk.Frame(nb, padding=12)
        nb.add(t, text="Tabs && cover")
        t.columnconfigure(1, weight=1)
        self.product = tk.StringVar(value=b.tabs.product)
        self.fit = tk.IntVar(value=b.tabs.fit_limit)
        self.style_ = tk.StringVar(value=b.tabs.label_style)
        ttk.Label(t, text="Tab product").grid(row=0, column=0, sticky="w")
        ttk.Entry(t, textvariable=self.product, width=24, state=ro).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Label(t, text="Tab fit limit (characters)").grid(row=1, column=0, sticky="w")
        ttk.Spinbox(t, from_=8, to=80, textvariable=self.fit, width=6, state=ro).grid(row=1, column=1, sticky="w", pady=2)
        ttk.Label(t, text="Tab labels").grid(row=2, column=0, sticky="w")
        sf = ttk.Frame(t)
        sf.grid(row=2, column=1, sticky="w")
        ttk.Radiobutton(sf, text="Letter and label: C) Campaign Management", value="letter_and_label",
                        variable=self.style_, state=ro).pack(anchor="w")
        ttk.Radiobutton(sf, text="Label only", value="label_only", variable=self.style_, state=ro).pack(anchor="w")
        ttk.Label(t, text="Cover lines").grid(row=3, column=0, sticky="nw", pady=(12, 0))
        self.cover = tk.Text(t, height=4, width=50, font=("Consolas", 10))
        self.cover.insert("1.0", "\n".join(b.cover.lines))
        self.cover.grid(row=3, column=1, sticky="we", pady=(12, 0))
        ttk.Label(t, text="Template folder").grid(row=4, column=0, sticky="w", pady=(12, 0))
        self.template_folder = tk.StringVar(value=b.template_folder or "")
        ttk.Entry(t, textvariable=self.template_folder, state=ro).grid(row=4, column=1, sticky="we", pady=(12, 0))
        ttk.Label(t, text="Relative to the binder folder. Blank = .binder/templates", style="Muted.TLabel").grid(
            row=5, column=1, sticky="w")
        ttk.Button(t, text="Copy the default templates there to restyle them...", command=self._export,
                   state=ro).grid(row=6, column=1, sticky="w", pady=4)

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Save", style="Accent.TButton", command=self._save, state=ro).pack(side="right", padx=8)
        self._update_preview()
        self.grab_set()

    # ---- helpers ----------------------------------------------------------------------
    def _header_lines(self) -> list[str]:
        return [ln for ln in self.lines.get("1.0", "end").splitlines() if ln.strip()]

    def _vars(self) -> dict[str, str]:
        out = {}
        for ln in self.vars_text.get("1.0", "end").splitlines():
            if "=" in ln:
                k, v = ln.split("=", 1)
                if k.strip():
                    out[k.strip()] = v.strip()
        return out

    def _sample(self) -> None:
        self.lines.delete("1.0", "end")
        self.lines.insert("1.0", "\n".join(SAMPLE_LAYOUT))
        self._update_preview()

    def _update_preview(self) -> None:
        ctx = dict(self._vars())
        ctx.update(binder_name=self.name.get(), doc_title="Sample Document", doc_code="A.01",
                   section_label="Administration")
        lines = [render_text(x, ctx) for x in self._header_lines()] or [""]
        right = f"{self.ref_label.get()} A.01"
        width = max(56, len(lines[0]) + len(right) + 4)
        first = lines[0] + " " * (width - len(lines[0]) - len(right)) + right
        self.preview.config(text="\n".join([first] + lines[1:]))

    def _apply_preset(self) -> None:
        values = FOOTER_PRESETS.get(self.preset.get())
        if values is None:
            return
        for k, v in zip(("left", "center", "right"), values):
            self.footer[k].set(v)
        current = self._vars()
        for k, v in PRESET_VARS.get(self.preset.get(), {}).items():
            current.setdefault(k, v)
        self.vars_text.delete("1.0", "end")
        self.vars_text.insert("1.0", "\n".join(f"{k} = {v}" for k, v in current.items()))

    def _export(self) -> None:
        if not self._save(close=False):
            return
        written = self.s.export_templates()
        folder = self.s.folder / (self.s.binder.template_folder or ".binder/templates")
        msg = (f"Copied {len(written)} template(s) to\n{folder}\n\nOpen them in Word or Excel to restyle. "
               "The next sync regenerates every document built from a changed template.") if written else \
            f"The templates are already in\n{folder}\n(existing templates are never overwritten)."
        messagebox.showinfo("Templates", msg, parent=self)

    def _save(self, close: bool = True) -> bool:
        try:
            self.s.update_binder(
                name=self.name.get().strip(),
                header_lines=self._header_lines(),
                ref_label=self.ref_label.get().strip() or "Packet Ref #",
                vars=self._vars(),
                footer={k: v.get() for k, v in self.footer.items()},
                page_numbers=self.pages.get(),
                tabs={"product": self.product.get(), "fit_limit": int(self.fit.get()), "label_style": self.style_.get()},
                cover={"lines": [ln for ln in self.cover.get("1.0", "end").splitlines() if ln.strip()]},
                template_folder=self.template_folder.get().strip() or None,
            )
        except (ManifestError, ReadOnlyError, ValueError, tk.TclError) as e:
            messagebox.showwarning("Can't save settings", str(e), parent=self)
            return False
        if close:
            self.destroy()
        return True
