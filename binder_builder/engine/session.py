"""An open binder: the GUI-free controller behind every screen (SPEC §10, §12).

The GUI (tkinter today, maybe PySide later) only calls this. A session holds the
structure lock while it is open (SPEC §10). If someone else holds it, the session
is read-only and every edit raises `ReadOnlyError`. Every edit is validated and
saved to the manifest immediately, so there is nothing to forget to save.
"""

from __future__ import annotations

import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from . import sync
from .codes import Layout, compute_layout
from .generate import CKL_TEMPLATE, DEFAULT_TEMPLATES, GenContext, SECTION_COVER_TEMPLATE, TOC_TEMPLATE
from .manifest import BINDER_DIR, Manifest, ManifestError, manifest_path
from .storage import LocalFolderStorage, LockedError, LockInfo
from .tree import Node, Tree, can_place


class ReadOnlyError(Exception):
    pass


@dataclass
class Status:
    orphans: int = 0
    broken_refs: int = 0
    warnings: int = 0
    changes: int = 0


class BinderSession:
    def __init__(self, folder: Path | str, lock: bool = True):
        self.folder = Path(folder)
        self.storage = LocalFolderStorage(self.folder)
        self.manifest = self.storage.open()
        self.tree = Tree(self.manifest)
        self.read_only = False
        self.lock_info: LockInfo | None = None
        self._owns_lock = False
        if lock:
            self._take_lock()

    # ---- lifecycle --------------------------------------------------------
    @classmethod
    def create(cls, folder: Path | str, name: str) -> "BinderSession":
        folder = Path(folder)
        folder.mkdir(parents=True, exist_ok=True)
        Manifest.create(folder, name)
        return cls(folder)

    @classmethod
    def create_from_structure(cls, source: Path | str, dest: Path | str, name: str,
                              copy_documents: bool = False) -> "BinderSession":
        """A new binder with another binder's structure and settings (SPEC §14 phase 5).

        Node ids are kept, so references written against the old binder resolve in the new one.
        Legacy filename overrides and copied_from are dropped (the new binder uses computed
        names). The binder's own templates are copied. With `copy_documents`, every document
        is copied across as a starting point (recorded in copied_from, so the plan warns when
        the original changes); otherwise sync creates blank documents with the right headers.
        """
        import copy as _copy

        source, dest = Path(source), Path(dest)
        if manifest_path(dest).exists():
            raise ManifestError(f"{dest} is already a binder")
        src = Manifest.load(source)
        raw = _copy.deepcopy(src.raw)
        raw["binder"]["name"] = name
        raw["binder"].pop("filenames", None)

        def clean(nodes):
            for n in nodes or []:
                for key in ("filename", "folder", "copied_from"):
                    n.pop(key, None)
                clean(n.get("children"))

        clean(raw.get("tree"))
        dest.mkdir(parents=True, exist_ok=True)
        Manifest(raw, manifest_path(dest)).save()
        src_ctx = GenContext(source, src, Tree(src), compute_layout(Tree(src), src.binder))
        if src_ctx.template_folder().is_dir():
            dst_templates = dest / (src.binder.template_folder or ".binder/templates")
            shutil.copytree(src_ctx.template_folder(), dst_templates, dirs_exist_ok=True)
        session = cls(dest)
        if copy_documents:
            state = sync.load_state(source)["files"]
            for e in src_ctx.layout.items():
                rel = state.get(e.entity_id, {}).get("path") or e.rel_path
                if (source / rel).is_file():
                    session.fill_from_file(e.entity_id, source / rel)
        return session

    @staticmethod
    def is_binder(folder: Path | str) -> bool:
        return manifest_path(Path(folder)).is_file()

    def _take_lock(self, force: bool = False) -> None:
        try:
            self._owns_lock = self.storage.lock(force=force, max_age_hours=self.manifest.binder.lock_max_age_hours)
            self.read_only = False
            self.lock_info = self.storage.lock_info()
        except LockedError as e:
            self.read_only = True
            self.lock_info = e.info

    def lock_is_stale(self) -> bool:
        return bool(self.read_only and self.lock_info
                    and self.lock_info.age_hours() >= self.manifest.binder.lock_max_age_hours)

    def break_lock(self) -> None:
        """Break someone else's stale lock (the GUI confirms first)."""
        self._take_lock(force=True)
        if self.read_only:
            raise LockedError(self.lock_info)
        self.reload()

    def close(self) -> None:
        if self._owns_lock:
            self.storage.unlock()
            self._owns_lock = False

    def reload(self) -> None:
        self.manifest = self.storage.open()
        self.tree = Tree(self.manifest)

    # ---- reading ----------------------------------------------------------
    @property
    def binder(self):
        return self.manifest.binder

    def layout(self) -> Layout:
        return compute_layout(self.tree, self.manifest.binder)

    def preview_move(self, node_id: str, parent_id: str | None, index: int) -> str | None:
        """The code the node would get if dropped there (for live codes while dragging); None if invalid."""
        raw = Manifest(_deepcopy(self.manifest.raw))
        tree = Tree(raw)
        try:
            tree.move(node_id, parent_id, index)
        except (ValueError, ManifestError):
            return None
        return compute_layout(tree, raw.binder).codes.get(node_id)

    def path_of(self, node_id: str) -> Path | None:
        e = self.layout().entity(node_id)
        if e is None:
            return None
        state = sync.load_state(self.folder)["files"].get(node_id)
        if state and (self.folder / state["path"]).is_file():
            return self.folder / state["path"]
        return self.folder / e.rel_path

    # ---- editing (each edit saves) -------------------------------------------
    def _check(self) -> None:
        if self.read_only:
            info = self.lock_info
            who = f"{info.user} on {info.machine}" if info else "someone else"
            raise ReadOnlyError(f"read-only: the binder is locked by {who}")

    def _save(self) -> None:
        self.storage.save(self.manifest)

    def _parent_for(self, kind: str, relative_to: str | None, mode: str) -> tuple[str | None, int | None]:
        """Resolve 'after <node>' / 'child of <node>' to (parent id, index)."""
        if relative_to is None:
            return None, None
        ref = self.tree.find(relative_to)
        if mode == "child":
            return ref.id, None
        parent = ref.parent
        if not can_place(kind, parent) and can_place(kind, ref):
            return ref.id, None  # e.g. a document added "after" a section goes inside it
        return (parent.id if parent else None), self.tree.index_of(ref) + 1

    def _insert(self, data: dict, relative_to: str | None, mode: str) -> Node:
        self._check()
        parent_id, index = self._parent_for(data["kind"], relative_to, mode)
        parent = self.tree.find(parent_id) if parent_id else None
        if not can_place(data["kind"], parent):
            where = f"inside a {parent.kind}" if parent else "at the top level"
            raise ValueError(f"a {data['kind'] if data['kind'] != 'item' else 'document'} cannot go {where}")
        node = self.tree.insert(data, parent_id, index)
        self._save()
        return node

    def add_document(self, title: str, ext: str, relative_to: str | None = None, mode: str = "after") -> Node:
        title = title.strip()
        if not title:
            raise ValueError("a document needs a title")
        if ext not in ("docx", "xlsx", "pdf"):
            raise ValueError(f"unknown document type {ext!r}")
        return self._insert({"kind": "item", "title": title, "ext": ext}, relative_to, mode)

    def add_section(self, label: str, relative_to: str | None = None, mode: str = "after") -> Node:
        label = label.strip()
        if not label:
            raise ValueError("a section needs a label")
        return self._insert({"kind": "section", "label": label}, relative_to, mode)

    def add_header(self, label: str, relative_to: str | None = None, mode: str = "after") -> Node:
        label = label.strip()
        if not label:
            raise ValueError("a heading needs a label")
        return self._insert({"kind": "header", "label": label}, relative_to, mode)

    def remove(self, node_id: str) -> Node:
        """Remove from the TOC. Files are never deleted; they show up as orphans."""
        self._check()
        node = self.tree.remove(node_id)
        self._save()
        return node

    def move(self, node_id: str, parent_id: str | None, index: int) -> Node:
        self._check()
        node = self.tree.move(node_id, parent_id, index)
        self._save()
        return node

    def move_up(self, node_id: str) -> Node:
        self._check()
        node = self.tree.move_by(node_id, -1)
        self._save()
        return node

    def move_down(self, node_id: str) -> Node:
        self._check()
        node = self.tree.move_by(node_id, +1)
        self._save()
        return node

    def indent(self, node_id: str) -> Node:
        self._check()
        node = self.tree.indent(node_id)
        self._save()
        return node

    def outdent(self, node_id: str) -> Node:
        self._check()
        node = self.tree.outdent(node_id)
        self._save()
        return node

    def update_node(self, node_id: str, **fields: Any) -> Node:
        """Set node fields; None or "" clears an optional field. The id never changes."""
        self._check()
        node = self.tree.find(node_id)
        old = {k: node.raw.get(k) for k in fields}
        for key, value in fields.items():
            if key in ("id", "kind", "children"):
                raise ValueError(f"{key} cannot be changed")
            if value in ("", None) and key not in ("title", "label"):
                node.set(key, None)
            elif key in ("included", "locked_header", "numeric") and value in (True, False):
                node.set(key, None if value == (key == "included") else value)  # keep defaults implicit
            else:
                node.set(key, value)
        try:
            self.manifest.validate()
        except ManifestError:
            for k, v in old.items():
                node.set(k, v)
            raise
        self._save()
        return node

    def update_binder(self, **fields: Any) -> None:
        """Set binder settings (name, header_lines, ref_label, vars, footer, page_numbers, tabs, cover, ...)."""
        self._check()
        import copy

        b = self.manifest.raw["binder"]
        old = {k: copy.deepcopy(b.get(k)) for k in fields}
        for key, value in fields.items():
            if key in ("tabs", "cover", "footer", "filenames") and isinstance(value, dict) and isinstance(b.get(key), dict):
                for k, v in value.items():
                    b[key][k] = v
            elif value is None:
                b.pop(key, None)
            else:
                b[key] = value
        try:
            self.manifest.validate()
        except ManifestError:
            for k, v in old.items():
                if v is None:
                    b.pop(k, None)
                else:
                    b[k] = v
            raise
        self._save()

    # ---- fill from file (SPEC §9.3) -------------------------------------------
    def fill_from_file(self, node_id: str, source: Path | str, replace: bool = False) -> Path:
        """Copy any file into this document's position. The source is never modified.

        If the position already has a file, `replace` must be True (the GUI confirms) and the
        current file is backed up first. The header is taken over on the next sync.
        """
        self._check()
        source = Path(source)
        node = self.tree.find(node_id)
        if node.kind != "item":
            raise ValueError("only documents can be filled from a file")
        ext = source.suffix.lower().lstrip(".")
        if ext not in ("docx", "xlsx", "pdf"):
            raise ValueError(f"{source.name}: only .docx, .xlsx and .pdf files can fill a position")
        if not source.is_file():
            raise FileNotFoundError(str(source))
        if node.get("ext") != ext:
            node.set("ext", ext)
            self.manifest.validate()
        state = sync.load_state(self.folder)
        entity = self.layout().entity(node_id)
        current = state["files"].get(node_id, {}).get("path")
        dest = self.folder / (current if current and (self.folder / current).is_file() and current.endswith(ext)
                              else entity.rel_path)
        if dest.exists():
            if not replace:
                raise FileExistsError(f"{entity.code} already has a file ({dest.name}); confirm to replace it")
            backup = sync._backup_dir(self.folder)
            backup = backup.with_name(backup.name + "_fill")
            (backup / dest.relative_to(self.folder)).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(dest, backup / dest.relative_to(self.folder))
            sync.write_backup_index(backup, [{"path": dest.relative_to(self.folder).as_posix(), "moved_to": None}],
                                    kind="fill")
        if source.resolve() == dest.resolve():
            raise ValueError("that file already is this document")
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, dest)

        binder_path, src_id = str(source.parent), ""
        for parent in source.parents:
            if manifest_path(parent).is_file():
                binder_path = str(parent)
                src_state = sync.load_state(parent)["files"]
                rel = source.relative_to(parent).as_posix()
                src_id = next((eid for eid, f in src_state.items() if f.get("path") == rel), "")
                break
        from datetime import datetime

        node.set("copied_from", {"binder_path": binder_path, "id": src_id,
                                 "synced_at": datetime.now().isoformat(timespec="seconds"),
                                 "source_hash": sync.file_hash(source), "source_path": str(source)})
        self._save()
        state["files"][node_id] = {"path": dest.relative_to(self.folder).as_posix()}
        sync.save_state(self.folder, state)
        return dest

    def recopy(self, node_id: str) -> Path:
        """Replace this binder's copy with the current source (backed up first)."""
        cf = self.tree.find(node_id).get("copied_from")
        if not cf or not cf.get("source_path"):
            raise ValueError("this document was not copied from another file")
        return self.fill_from_file(node_id, cf["source_path"], replace=True)

    # ---- templates --------------------------------------------------------
    def export_templates(self) -> list[Path]:
        """Copy the default templates into the binder's template folder so staff can restyle them.
        Existing templates there are kept, never overwritten."""
        ctx = GenContext(self.folder, self.manifest, self.tree, self.layout())
        dest = ctx.template_folder()
        dest.mkdir(parents=True, exist_ok=True)
        written = []
        names = {self.binder.cover.template: "00_Cover.docx", self.binder.tabs.template: "Tabs.docx",
                 SECTION_COVER_TEMPLATE: SECTION_COVER_TEMPLATE, TOC_TEMPLATE: TOC_TEMPLATE,
                 CKL_TEMPLATE: CKL_TEMPLATE}
        for name, default in names.items():
            target = dest / name
            if not target.exists():
                shutil.copy2(DEFAULT_TEMPLATES / default, target)
                written.append(target)
        return written

    # ---- sync -------------------------------------------------------------
    def plan(self) -> sync.Plan:
        return sync.plan(self.folder)

    def apply(self, plan: sync.Plan) -> sync.ApplyResult:
        self._check()
        return sync.apply(plan)

    def status(self, plan: sync.Plan) -> Status:
        return Status(
            orphans=len(plan.orphans),
            broken_refs=sum(1 for a in plan.warnings if "broken reference" in a.detail),
            warnings=len(plan.warnings),
            changes=len(plan.changes),
        )

    def backups(self) -> list[str]:
        return sync.list_backups(self.folder)

    def backup_info(self, name: str) -> dict:
        return sync.backup_info(self.folder, name)

    def restore(self, backup: str) -> list[str]:
        self._check()
        return sync.restore(self.folder, backup)

    def open_in_office(self, node_id: str) -> None:
        """Open a document in Word/Excel (Windows file association)."""
        p = self.path_of(node_id)
        if p is None or not p.exists():
            raise FileNotFoundError("this document has no file yet; sync first")
        os.startfile(str(p))  # type: ignore[attr-defined]  # Windows only


def _deepcopy(raw):
    import copy

    return copy.deepcopy(raw)


def binder_dir(folder: Path | str) -> Path:
    return Path(folder) / BINDER_DIR
