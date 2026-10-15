"""Binder screen (SPEC §12): tree with drag-and-drop and live codes, node details, lock banner, status."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..engine.codes import ITEM
from ..engine.session import BinderSession, ReadOnlyError

KIND_NAMES = {"docx": "Word document", "xlsx": "Excel workbook", "pdf": "PDF file"}
FILL_TYPES = [("Office documents and PDFs", "*.docx *.xlsx *.pdf"), ("Word", "*.docx"), ("Excel", "*.xlsx"),
              ("PDF", "*.pdf")]
CHECK_ON, CHECK_OFF = "☑", "☐"


class BinderView(ttk.Frame):
    def __init__(self, app, session: BinderSession):
        super().__init__(app)
        self.app = app
        self.s = session
        self._drag: dict | None = None
        self._last_status = None

        self._build_banner()
        self._build_toolbar()
        body = ttk.PanedWindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=8, pady=(4, 0))
        left = ttk.Frame(body)
        right = ttk.Frame(body, padding=(12, 0, 0, 0))
        body.add(left, weight=3)
        body.add(right, weight=2)
        self._build_tree(left)
        self.details = DetailsPane(right, self)
        self.details.pack(fill="both", expand=True)
        self._build_statusbar()

        self.refresh()
        self.refresh_status()

    # ---- layout ---------------------------------------------------------------------
    def _build_banner(self) -> None:
        self.banner = ttk.Frame(self, style="Banner.TFrame")
        if not self.s.read_only:
            return
        info = self.s.lock_info
        since = f"{info.since:%d %b %H:%M}" if info else "?"
        text = (f"Locked by {info.user} on {info.machine} since {since}. You can read the tree and open "
                f"documents, but structure changes and sync are disabled." if info else "Locked by someone else.")
        ttk.Label(self.banner, text=text, style="Banner.TLabel").pack(side="left", fill="x", expand=True)
        if self.s.lock_is_stale():
            ttk.Button(self.banner, text="Break lock...", command=self._break_lock).pack(side="right", padx=8, pady=4)
        self.banner.pack(fill="x")

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self, padding=(8, 8, 8, 4))
        bar.pack(fill="x")
        ro = "disabled" if self.s.read_only else "normal"

        def btn(text, cmd, state=ro, **kw):
            b = ttk.Button(bar, text=text, command=cmd, state=state, **kw)
            b.pack(side="left", padx=(0, 4))
            return b

        btn("Add document...", lambda: self._add("item"))
        btn("Add section...", lambda: self._add("section"))
        btn("Add heading...", lambda: self._add("header"))
        btn("Remove from TOC", self._remove)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6)
        btn("↑", lambda: self._edit(self.s.move_up), width=3)
        btn("↓", lambda: self._edit(self.s.move_down), width=3)
        btn("Indent →", lambda: self._edit(self.s.indent))
        btn("← Outdent", lambda: self._edit(self.s.outdent))
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6)
        btn("Open file", self._open_file, state="normal")
        btn("Insert reference...", self._insert_reference, state="normal")
        ttk.Button(bar, text="Close binder", command=self.app.close_binder).pack(side="right")
        ttk.Button(bar, text="Print & export...", command=self._print).pack(side="right", padx=4)
        ttk.Button(bar, text="Settings...", command=self._settings).pack(side="right", padx=4)
        ttk.Button(bar, text="Sync...", style="Accent.TButton", command=self._sync).pack(side="right", padx=4)

    def _build_tree(self, parent) -> None:
        ttk.Label(parent, text="Table of Contents", style="Heading.TLabel").pack(anchor="w")
        ttk.Label(parent, text="Drag to reorder (drop on the middle of a row to put it inside). "
                               "Click the Include box to include or exclude.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 4))
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        self.tv = ttk.Treeview(frame, columns=("code", "incl"), selectmode="browse")
        self.tv.heading("#0", text="Title")
        self.tv.heading("code", text="Code")
        self.tv.heading("incl", text="Include")
        self.tv.column("#0", width=380, stretch=True)
        self.tv.column("code", width=90, stretch=False, anchor="w")
        self.tv.column("incl", width=60, stretch=False, anchor="center")
        self.tv.tag_configure("section", font=("Segoe UI", 10, "bold"))
        self.tv.tag_configure("header", font=("Segoe UI", 10, "italic"), foreground="#555555")
        self.tv.tag_configure("excluded", foreground="#999999")
        self.tv.tag_configure("drop", background="#CCE4F7")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tv.bind("<<TreeviewSelect>>", lambda e: self.details.show(self.selected()))
        self.tv.bind("<ButtonPress-1>", self._on_press)
        self.tv.bind("<B1-Motion>", self._on_motion)
        self.tv.bind("<ButtonRelease-1>", self._on_release)
        self.tv.bind("<Double-Button-1>", self._on_double)

    def _build_statusbar(self) -> None:
        bar = ttk.Frame(self, padding=(8, 4))
        bar.pack(fill="x")
        self.status = ttk.Label(bar, text="")
        self.status.pack(side="left")
        self.busy = ttk.Label(bar, text="", style="Muted.TLabel")
        self.busy.pack(side="right")
        ttk.Button(bar, text="Refresh status", command=self.refresh_status).pack(side="right", padx=8)
        ttk.Button(bar, text="Report a problem...",
                   command=lambda: self.app.show_problem_report(self.s.folder)).pack(side="right")

    # ---- tree -------------------------------------------------------------------------
    def selected(self) -> str | None:
        sel = self.tv.selection()
        return sel[0] if sel else None

    def refresh(self, select: str | None = None) -> None:
        select = select or self.selected()
        opened = {i for i in self._all_items() if self.tv.item(i, "open")}
        first = not self.tv.get_children()
        self.tv.delete(*self.tv.get_children())
        layout = self.s.layout()
        included = {e.entity_id: e.included for e in layout.entities}

        def add(node, parent_iid):
            inc = included.get(node.id, node.included)
            tags = [node.kind] + ([] if inc else ["excluded"])
            title = node.text
            if node.kind == ITEM:
                title = f"{title}   ({node.get('ext')})"
            self.tv.insert(parent_iid, "end", iid=node.id, text=title, tags=tags,
                           values=(layout.codes.get(node.id, ""), CHECK_ON if node.included else CHECK_OFF),
                           open=first or node.id in opened)
            for c in node.children:
                add(c, node.id)

        for r in self.s.tree.roots:
            add(r, "")
        if select and self.tv.exists(select):
            self.tv.selection_set(select)
            self.tv.see(select)
        self.details.show(self.selected())

    def _all_items(self, parent=""):
        for i in self.tv.get_children(parent):
            yield i
            yield from self._all_items(i)

    # ---- drag and drop with live codes -----------------------------------------------
    def _drop_target(self, y: int):
        row = self.tv.identify_row(y)
        if not row:
            return None
        x, top, w, h = self.tv.bbox(row) or (0, 0, 0, 1)
        rel = (y - top) / max(h, 1)
        node = self.s.tree.find(row)
        parent = node.parent
        if rel < 0.3:
            return parent.id if parent else None, self.s.tree.index_of(node), row, "before"
        if rel > 0.7 and not (node.children and self.tv.item(row, "open")):
            return parent.id if parent else None, self.s.tree.index_of(node) + 1, row, "after"
        return node.id, len(node.children), row, "into"

    def _on_press(self, event) -> None:
        region = self.tv.identify_region(event.x, event.y)
        row = self.tv.identify_row(event.y)
        if region == "cell" and self.tv.identify_column(event.x) == "#2" and row:
            self._toggle_include(row)
            return "break"
        self._drag = {"id": row, "y": event.y, "active": False, "target": None} if row else None

    def _on_motion(self, event) -> None:
        d = self._drag
        if not d or self.s.read_only:
            return
        if not d["active"] and abs(event.y - d["y"]) < 6:
            return
        d["active"] = True
        for i in self._all_items():
            self.tv.item(i, tags=[t for t in self.tv.item(i, "tags") if t != "drop"])
        target = self._drop_target(event.y)
        d["target"] = target
        if not target:
            self.set_busy("")
            return
        parent_id, index, row, where = target
        if row != d["id"]:
            self.tv.item(row, tags=list(self.tv.item(row, "tags")) + ["drop"])
        code = self.s.preview_move(d["id"], parent_id, index)
        name = self.s.tree.find(d["id"]).text
        self.set_busy(f"Drop {where} '{self.s.tree.find(row).text}': {name} becomes {code}" if code
                      else "Can't drop here")

    def _on_release(self, event) -> None:
        d, self._drag = self._drag, None
        for i in self._all_items():
            self.tv.item(i, tags=[t for t in self.tv.item(i, "tags") if t != "drop"])
        if not d or not d["active"] or not d["target"]:
            return
        self.set_busy("")
        parent_id, index, row, _ = d["target"]
        if row == d["id"]:
            return
        try:
            self.s.move(d["id"], parent_id, index)
        except (ValueError, ReadOnlyError) as e:
            messagebox.showwarning("Can't move there", str(e), parent=self)
            return
        self.refresh(select=d["id"])

    def _on_double(self, event) -> None:
        row = self.tv.identify_row(event.y)
        if row and self.s.tree.find(row).kind == ITEM and self.tv.identify_column(event.x) != "#2":
            self._open_file()

    # ---- actions -----------------------------------------------------------------------
    def _edit(self, fn) -> None:
        node_id = self.selected()
        if not node_id:
            return
        try:
            fn(node_id)
        except (ValueError, ReadOnlyError) as e:
            messagebox.showwarning("Binder Builder", str(e), parent=self)
            return
        self.refresh(select=node_id)

    def _toggle_include(self, node_id: str) -> None:
        if self.s.read_only:
            return
        node = self.s.tree.find(node_id)
        self.s.update_node(node_id, included=not node.included)
        self.refresh(select=node_id)

    def _add(self, kind: str) -> None:
        dlg = AddNodeDialog(self, kind, self.s, self.selected())
        self.wait_window(dlg)
        if dlg.result:
            self.refresh(select=dlg.result)

    def _remove(self) -> None:
        node_id = self.selected()
        if not node_id:
            return
        node = self.s.tree.find(node_id)
        count = sum(1 for _ in node.walk())
        extra = f" and the {count - 1} entries inside it" if count > 1 else ""
        if not messagebox.askyesno(
            "Remove from TOC",
            f"Remove '{node.text}'{extra} from the Table of Contents?\n\n"
            "No file is deleted. The file stays where it is and is listed as an orphan after the next sync; "
            "delete it yourself in Explorer if you no longer need it.",
            parent=self,
        ):
            return
        self.s.remove(node_id)
        self.refresh()

    def _open_file(self) -> None:
        node_id = self.selected()
        if not node_id:
            return
        try:
            self.s.open_in_office(node_id)
        except FileNotFoundError as e:
            messagebox.showinfo("Open file", str(e), parent=self)

    def _insert_reference(self) -> None:
        dlg = PickDocumentDialog(self, self.s, self.selected())
        self.wait_window(dlg)
        if not dlg.choice:
            return
        node_id, code = dlg.choice
        try:
            from ..office.insert_reference import NoActiveDocument, insert_into_active
        except ImportError:
            messagebox.showerror("Insert reference", "Inserting references needs Microsoft Word or Excel.", parent=self)
            return
        try:
            msg = insert_into_active(self.s.folder, node_id, code)
        except NoActiveDocument as e:
            messagebox.showinfo("Insert reference", str(e), parent=self)
            return
        messagebox.showinfo("Insert reference", msg, parent=self)

    def _settings(self) -> None:
        from .settings import SettingsDialog

        dlg = SettingsDialog(self, self.s)
        self.wait_window(dlg)
        self.refresh()

    def _sync(self) -> None:
        from .sync_view import SyncDialog

        dlg = SyncDialog(self, self.app, self.s)
        self.wait_window(dlg)
        self.refresh()
        self.refresh_status()

    def _print(self) -> None:
        from .print_view import PrintDialog

        PrintDialog(self, self.s, self.app, self.selected())

    def _break_lock(self) -> None:
        info = self.s.lock_info
        if not messagebox.askyesno(
            "Break lock",
            f"{info.user} on {info.machine} has held this binder since {info.since:%d %b %H:%M}.\n\n"
            "Break their lock? Only do this if they are not editing the structure right now.",
            parent=self,
        ):
            return
        self.s.break_lock()
        self.app.open_binder(self.s.folder)

    # ---- status ------------------------------------------------------------------------
    def set_busy(self, text: str) -> None:
        self.busy.config(text=text)

    def refresh_status(self) -> None:
        self.app.run_task(self.s.plan, self._show_status, busy="Checking the binder...",
                          on_error=lambda e: self.status.config(text=f"Could not check the binder: {e}"))

    def _show_status(self, plan) -> None:
        st = self.s.status(plan)
        self._last_status = st
        pending = f"{st.changes} change(s) waiting for sync" if st.changes else "In sync"
        self.status.config(text=f"{pending}   ·   {st.orphans} orphan file(s)   ·   "
                                f"{st.broken_refs} broken reference(s)   ·   {st.warnings} warning(s)")


class PickDocumentDialog(tk.Toplevel):
    """Pick the document a reference points at (SPEC §12 'Insert reference')."""

    def __init__(self, master, session: BinderSession, selected: str | None):
        super().__init__(master)
        self.title("Insert reference")
        self.transient(master)
        self.choice: tuple[str, str] | None = None
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Put the cursor where the reference goes in Word (or select the cell in Excel), "
                          "then pick the document it refers to. The reference shows the document's code and "
                          "updates itself on every sync.", style="Muted.TLabel", wraplength=480).pack(anchor="w")
        self.tv = ttk.Treeview(f, columns=("code",), height=16, selectmode="browse")
        self.tv.heading("#0", text="Document")
        self.tv.heading("code", text="Code")
        self.tv.column("code", width=90, stretch=False)
        self.tv.pack(fill="both", expand=True, pady=8)
        codes = session.layout().codes

        def add(node, parent):
            self.tv.insert(parent, "end", iid=node.id, text=node.text, values=(codes.get(node.id, ""),), open=True)
            for c in node.children:
                add(c, node.id)

        for r in session.tree.roots:
            add(r, "")
        if selected and self.tv.exists(selected):
            self.tv.selection_set(selected)
            self.tv.see(selected)
        self.codes = codes
        self.tv.bind("<Double-Button-1>", lambda e: self._ok())
        bf = ttk.Frame(f)
        bf.pack(fill="x")
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(bf, text="Insert at cursor", style="Accent.TButton", command=self._ok).pack(side="right", padx=8)
        self.grab_set()

    def _ok(self) -> None:
        sel = self.tv.selection()
        if sel and self.codes.get(sel[0]):
            self.choice = (sel[0], self.codes[sel[0]])
            self.destroy()


class AddNodeDialog(tk.Toplevel):
    """Add document / section / heading: title, type, position (SPEC §12 'Add document')."""

    def __init__(self, master, kind: str, session: BinderSession, selected: str | None):
        super().__init__(master)
        self.s = session
        self.kind = kind
        self.selected = selected
        self.result: str | None = None
        noun = {"item": "document", "section": "section", "header": "heading"}[kind]
        self.title(f"Add {noun}")
        self.transient(master)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text="Title:" if kind == "item" else "Label:").grid(row=0, column=0, sticky="w")
        self.text = tk.StringVar()
        e = ttk.Entry(f, textvariable=self.text, width=48)
        e.grid(row=0, column=1, sticky="we", pady=4)
        e.focus_set()
        row = 1
        if kind == "item":
            ttk.Label(f, text="Type:").grid(row=row, column=0, sticky="nw", pady=4)
            self.ext = tk.StringVar(value="docx")
            types = ttk.Frame(f)
            types.grid(row=row, column=1, sticky="w")
            for ext in ("docx", "xlsx", "pdf"):
                ttk.Radiobutton(types, text=KIND_NAMES[ext], value=ext, variable=self.ext).pack(side="left", padx=(0, 8))
            row += 1
        ttk.Label(f, text="Position:").grid(row=row, column=0, sticky="nw", pady=4)
        self.mode = tk.StringVar(value=self._default_mode())
        pos = ttk.Frame(f)
        pos.grid(row=row, column=1, sticky="w")
        sel_name = session.tree.find(selected).text if selected else None
        if sel_name:
            ttk.Radiobutton(pos, text=f"After '{sel_name}'", value="after", variable=self.mode).pack(anchor="w")
            ttk.Radiobutton(pos, text=f"Inside '{sel_name}'", value="child", variable=self.mode).pack(anchor="w")
        else:
            ttk.Label(pos, text="At the end of the Table of Contents").pack(anchor="w")
        row += 1
        self.fill_source: str | None = None
        if kind == "item":
            fill = ttk.Frame(f)
            fill.grid(row=row, column=1, sticky="w", pady=4)
            ttk.Button(fill, text="Fill from file...", command=self._pick_fill).pack(side="left")
            self.fill_lbl = ttk.Label(fill, text="(optional: copy an existing document into this position)",
                                      style="Muted.TLabel")
            self.fill_lbl.pack(side="left", padx=8)
            row += 1
        buttons = ttk.Frame(f)
        buttons.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Add", style="Accent.TButton", command=self._ok).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="left")
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()

    def _default_mode(self) -> str:
        if not self.selected:
            return "after"
        sel = self.s.tree.find(self.selected)
        if self.kind == "item" and sel.kind == "section":
            return "child"
        return "after"

    def _pick_fill(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Copy which file into this position?",
                                          filetypes=FILL_TYPES)
        if not path:
            return
        self.fill_source = path
        p = Path(path)
        self.fill_lbl.config(text=p.name)
        self.ext.set(p.suffix.lower().lstrip("."))
        if not self.text.get().strip():
            self.text.set(p.stem.split("_", 1)[-1].replace("_", " ") if p.stem[:2].isdigit() else p.stem.replace("_", " "))

    def _ok(self) -> None:
        try:
            rel = self.selected
            if self.kind == "item":
                if rel is None:
                    raise ValueError("Select a section first: documents go inside sections.")
                node = self.s.add_document(self.text.get(), self.ext.get(), rel, self.mode.get())
                if self.fill_source:
                    self.s.fill_from_file(node.id, self.fill_source)
            elif self.kind == "section":
                node = self.s.add_section(self.text.get(), rel, self.mode.get())
            else:
                node = self.s.add_header(self.text.get(), rel, self.mode.get())
        except (ValueError, ReadOnlyError) as e:
            messagebox.showwarning("Binder Builder", str(e), parent=self)
            return
        self.result = node.id
        self.destroy()


class DetailsPane(ttk.Frame):
    """Right-hand node details (SPEC §12): every option for the selected node."""

    def __init__(self, master, view: BinderView):
        super().__init__(master)
        self.view = view
        self.s = view.s
        self.node_id: str | None = None
        self.vars: dict[str, tk.Variable] = {}
        self.show(None)

    def _clear(self) -> None:
        for w in self.winfo_children():
            w.destroy()
        self.vars = {}

    def show(self, node_id: str | None) -> None:
        self._clear()
        self.node_id = node_id
        if not node_id or self.s.tree.get(node_id) is None:
            ttk.Label(self, text="Select an entry to see its details.", style="Muted.TLabel").pack(anchor="nw")
            return
        node = self.s.tree.find(node_id)
        layout = self.s.layout()
        entity = layout.entity(node_id) or layout.entity(f"{node_id}#cover")
        ro = "disabled" if self.s.read_only else "normal"
        kind_text = {"item": KIND_NAMES.get(node.get("ext"), "Document"), "section": "Section",
                     "header": "Clarifying heading (no letter, folder or tab)"}[node.kind]
        ttk.Label(self, text=node.text, style="Heading.TLabel", wraplength=380).pack(anchor="w")
        ttk.Label(self, text=f"{kind_text}   ·   code {layout.codes.get(node_id, '-')}   ·   id {node_id}",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 8))

        form = ttk.Frame(self)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        r = 0

        def field(label, key, value, widget="entry", hint=None, choices=None):
            nonlocal r
            ttk.Label(form, text=label).grid(row=r, column=0, sticky="nw", pady=3, padx=(0, 8))
            if widget == "check":
                var = tk.BooleanVar(value=bool(value))
                ttk.Checkbutton(form, variable=var, state=ro).grid(row=r, column=1, sticky="w")
            elif widget == "combo":
                var = tk.StringVar(value=value or "")
                ttk.Combobox(form, textvariable=var, values=choices, state="readonly" if ro == "normal" else "disabled",
                             width=10).grid(row=r, column=1, sticky="w")
            else:
                var = tk.StringVar(value=value or "")
                ttk.Entry(form, textvariable=var, state=ro).grid(row=r, column=1, sticky="we")
            self.vars[key] = var
            r += 1
            if hint:
                ttk.Label(form, text=hint, style="Muted.TLabel", wraplength=320).grid(row=r, column=1, sticky="w")
                r += 1

        if node.kind == "item":
            field("Title", "title", node.text)
            field("Type", "ext", node.get("ext"), "combo", choices=["docx", "xlsx", "pdf"])
            field("Included in binder", "included", node.included, "check")
            field("Suffix", "suffix", node.get("suffix"), hint="e.g. PB gives A.03.PB; suffixed entries don't take a number")
            field("Filename override", "filename", node.get("filename"),
                  hint="Title part only. Leave blank to use the title. The number prefix is always computed.")
            field("Locked header", "locked_header", node.get("locked_header"), "check",
                  hint="Never touch this document's header or footer (licensed forms, signed pages, scans)")
            f = node.get("footer") or {}
            field("Footer left", "footer_left", f.get("left"))
            field("Footer center", "footer_center", f.get("center"))
            field("Footer right", "footer_right", f.get("right"),
                  hint="Leave all three blank to use the binder's footer")
            if node.get("ext") == "xlsx":
                field("Checklist-style rows", "hybrid_rows", node.get("hybrid_rows"), "check",
                      hint="Keep one row per document in this workbook (Reference/Description columns)")
            cf = node.get("copied_from")
            ttk.Label(form, text="Copied from").grid(row=r, column=0, sticky="nw", pady=3)
            src = cf.get("source_path") or cf.get("binder_path") if cf else None
            ttk.Label(form, text=f"{src} on {cf['synced_at'][:16].replace('T', ' ')}" if cf else "-",
                      wraplength=320).grid(row=r, column=1, sticky="w")
            r += 1
        elif node.kind == "section":
            field("Label", "label", node.text)
            field("Description", "description", node.get("description"),
                  hint="One line printed under the section title on its cover sheet")
            field("Included in binder", "included", node.included, "check")
            field("Control section (0, 0.01...)", "numeric", node.get("numeric"), "check")
            field("Cover filename override", "filename", node.get("filename"),
                  hint="Title part of the cover sheet's file name (default: Cover)")
            field("Folder override", "folder", node.get("folder"),
                  hint="Label part of the folder name (default: the label)")
            field("Locked header (cover)", "locked_header", node.get("locked_header"), "check")
        else:
            field("Label", "label", node.text)

        if entity is not None:
            ttk.Label(self, text=f"File: {entity.rel_path}", style="Muted.TLabel", wraplength=380).pack(
                anchor="w", pady=(10, 0))
        buttons = ttk.Frame(self)
        buttons.pack(anchor="w", pady=12)
        ttk.Button(buttons, text="Save changes", style="Accent.TButton", command=self._save, state=ro).pack(side="left")
        if node.kind == "item":
            ttk.Button(buttons, text="Open file", command=self.view._open_file).pack(side="left", padx=8)
            ttk.Button(buttons, text="Fill from file...", command=self._fill, state=ro).pack(side="left")
            if node.get("copied_from"):
                ttk.Button(buttons, text="Re-copy from source", command=self._recopy, state=ro).pack(
                    side="left", padx=8)

    def _fill(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Copy which file into this position?",
                                          filetypes=FILL_TYPES)
        if not path:
            return
        try:
            self.s.fill_from_file(self.node_id, path)
        except FileExistsError as e:
            if not messagebox.askyesno("Replace file?", f"{e}.\n\nReplace it? The current file is backed up first.",
                                       parent=self):
                return
            self.s.fill_from_file(self.node_id, path, replace=True)
        messagebox.showinfo("Filled", "Copied. The original was not changed. The next sync takes over its header.",
                            parent=self)
        self.view.refresh(select=self.node_id)
        self.view.refresh_status()

    def _recopy(self) -> None:
        if not messagebox.askyesno("Re-copy from source",
                                   "Replace this binder's copy with the current version of the source file?\n\n"
                                   "This binder's copy is backed up first; any edits made to it here are replaced.",
                                   parent=self):
            return
        self.s.recopy(self.node_id)
        self.view.refresh(select=self.node_id)
        self.view.refresh_status()

    def _save(self) -> None:
        node = self.s.tree.find(self.node_id)
        values = {k: v.get() for k, v in self.vars.items()}
        fields = {}
        for key in ("title", "label", "ext", "suffix", "filename", "folder", "description"):
            if key in values:
                fields[key] = values[key].strip() if isinstance(values[key], str) else values[key]
        for key in ("included", "locked_header", "numeric", "hybrid_rows"):
            if key in values:
                fields[key] = bool(values[key])
        if "hybrid_rows" in fields and not fields["hybrid_rows"]:
            fields["hybrid_rows"] = None
        if "footer_left" in values:
            footer = {"left": values["footer_left"], "center": values["footer_center"], "right": values["footer_right"]}
            fields["footer"] = footer if any(v.strip() for v in footer.values()) else None
        try:
            self.s.update_node(node.id, **fields)
        except (ValueError, ReadOnlyError, Exception) as e:  # noqa: BLE001 - shown to the user
            messagebox.showwarning("Can't save", str(e), parent=self)
            return
        self.view.refresh(select=node.id)
