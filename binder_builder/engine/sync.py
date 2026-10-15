"""Sync: plan, apply with backup and invariant check, state, log, restore (SPEC §7).

Safety rules (SPEC §1.1) enforced here:
  * no user file is ever deleted; removed TOC rows leave files in place as orphans
  * every change is listed in the plan before apply
  * every file about to be modified or moved is copied into .binder/backups/<ts>/ first
  * every write to an author's Word/Excel file goes to a temp file, is body-hashed
    against the original, and only replaces the original if the hashes match.
    Hybrid workbooks (the checklist) use the user-column hash instead. Generated
    documents (TOC, covers, tabs) are app-owned and replaced wholesale.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from . import excel, pdfdoc, word
from .codes import BINDER_COVER, CKL, ITEM, Entity, Layout, compute_layout, header_spec
from .generate import CKL_TEMPLATE, GENERATED, GenContext, content_signature, render, template_hash
from .generate import checklist
from .manifest import BINDER_DIR, Manifest
from .storage import LocalFolderStorage
from .storage.local_folder import current_user
from .tree import Tree

STATE_NAME = "state.json"
BACKUPS = "backups"
WORK = "tmp"

CREATE = "create_file"
MOVE = "move_file"
STAMP = "stamp_header"
REFS = "update_refs"
REGENERATE = "regenerate"
CHECKLIST = "checklist_rows"
ORPHAN = "orphan"
WARN = "warn"

CHANGE_KINDS = (MOVE, CREATE, STAMP, REFS, REGENERATE, CHECKLIST)
ACTION_ORDER = {k: i for i, k in enumerate(CHANGE_KINDS + (ORPHAN, WARN))}

CREATABLE = (ITEM, CKL)  # created when missing; generated documents are rendered instead
IGNORED_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}


@dataclass
class Action:
    kind: str
    path: str  # relative path the action concerns (current location, or target for creates)
    detail: str = ""
    entity_id: str | None = None
    target: str | None = None

    def row(self) -> tuple[str, str, str]:
        return self.kind, self.path, self.detail


@dataclass
class FileWork:
    """Everything the apply step will do to one entity's file."""

    entity: Entity
    current: str | None  # relative path now (None: to be created)
    target: str
    move: bool = False
    create: bool = False
    write: bool = False  # header stamp, ref update and/or checklist rows needed
    regenerate: bool = False  # generated document to be re-rendered
    hybrid: tuple[str, str] | None = None  # (reference heading, description heading)
    pdf_restash: bool = False  # PDF item: the file on disk is a new original to keep in pdf_sources


@dataclass
class Plan:
    folder: Path
    actions: list[Action]
    layout: Layout
    work: dict[str, FileWork]
    rows: list[tuple[str, str, str, str]]

    @property
    def changes(self) -> list[Action]:
        return [a for a in self.actions if a.kind in CHANGE_KINDS]

    @property
    def warnings(self) -> list[Action]:
        return [a for a in self.actions if a.kind == WARN]

    @property
    def orphans(self) -> list[Action]:
        return [a for a in self.actions if a.kind == ORPHAN]

    @property
    def is_empty(self) -> bool:
        return not self.changes

    def render(self) -> str:
        if not self.actions:
            return "Nothing to do: the binder matches its manifest."
        w1 = max(len(a.kind) for a in self.actions)
        w2 = max(len(a.path) for a in self.actions)
        lines = [f"{'ACTION'.ljust(w1)}  {'FILE'.ljust(w2)}  DETAIL"]
        for a in self.actions:
            lines.append(f"{a.kind.ljust(w1)}  {a.path.ljust(w2)}  {a.detail}")
        lines.append(f"\n{len(self.changes)} change(s), {len(self.orphans)} orphan(s), {len(self.warnings)} warning(s)")
        return "\n".join(lines)


@dataclass
class ApplyResult:
    backup: str | None
    applied: list[Action] = field(default_factory=list)
    failures: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.failures


# ---- state ------------------------------------------------------------------

def state_path(folder: Path) -> Path:
    return folder / BINDER_DIR / STATE_NAME


def load_state(folder: Path) -> dict:
    p = state_path(folder)
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {"format": 1, "files": {}, "printed": {}}


def save_state(folder: Path, state: dict) -> None:
    p = state_path(folder)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(state, indent=2, sort_keys=True), encoding="utf-8")
    tmp.replace(p)


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def body_hash(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".docx":
        return word.body_hash(path)
    if ext == ".xlsx":
        return excel.body_hash(path)
    return file_hash(path)


# ---- helpers ------------------------------------------------------------------

def _norm(rel: str) -> str:
    return os.path.normcase(rel.replace("/", os.sep))


def _in_use(path: Path) -> bool:
    """True if another process (Word, Excel) holds the file open for writing."""
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def _disk_files(folder: Path, skip_dirs: set[str]) -> list[str]:
    out = []
    for root, dirs, files in os.walk(folder):
        rel_root = Path(root).relative_to(folder)
        dirs[:] = [d for d in dirs if (rel_root / d).as_posix() not in skip_dirs]
        for f in files:
            if f.startswith("~$") or f.lower() in IGNORED_NAMES:
                continue
            out.append((rel_root / f).as_posix())
    return sorted(out)


def binder_rows(tree: Tree, layout: Layout) -> list[tuple[str, str, str, str]]:
    """Rows of the `_binder` sheet: every coded node in tree order, then coded generated documents."""
    rows = []
    folders = {s.node.id: s.folder.rsplit("/", 1)[-1] for s in layout.sections}
    for n in tree.walk():
        code = layout.codes.get(n.id)
        if code is None:
            continue
        e = layout.entity(n.id)
        rows.append((n.id, code, n.text, e.filename if e else folders.get(n.id, "")))
    for e in layout.entities:
        if e.kind != ITEM and e.code:
            rows.append((e.entity_id, e.code, e.title, e.filename))
    return rows


def checklist_specs(layout: Layout, items_only: bool = False) -> list[checklist.RowSpec]:
    """Documents that get a checklist row, in TOC order (SPEC §8.1: TOC, CKL, covers and tabs too)."""
    out = []
    for e in layout.entities:
        if e.kind == BINDER_COVER or (items_only and e.kind != ITEM):
            continue
        out.append(checklist.RowSpec(e.entity_id, e.code, e.title, e.included))
    return out


def _hybrid_columns(e: Entity) -> tuple[str, str] | None:
    if e.kind == CKL:
        return checklist.REFERENCE, checklist.DESCRIPTION
    if e.kind == ITEM and e.ext == "xlsx" and e.node is not None and e.node.get("hybrid_rows"):
        cols = list(e.node.get("hybrid_columns") or [checklist.REFERENCE, checklist.DESCRIPTION])
        return cols[0], cols[1]
    return None


def _spec_for(e: Entity, manifest: Manifest):
    if not e.stamp or e.locked_header:
        return None
    return header_spec(e, manifest.binder)


def _render_generated(e: Entity, ctx: GenContext, out: Path, rows) -> None:
    """Render a generated document and stamp its header: the full desired file."""
    render(e, ctx, out)
    spec = _spec_for(e, ctx.manifest)
    if e.ext == "docx":
        word.update(out, out, spec, ctx.layout.codes)
    else:
        excel.update(out, out, spec, rows)


# ---- plan -------------------------------------------------------------------

def plan(folder: Path | str) -> Plan:
    folder = Path(folder)
    manifest = LocalFolderStorage(folder).open()
    tree = Tree(manifest)
    layout = compute_layout(tree, manifest.binder)
    rows = binder_rows(tree, layout)
    state = load_state(folder)
    ctx = GenContext(folder, manifest, tree, layout, pages=_pages(manifest, state))
    actions: list[Action] = [Action(WARN, "", w) for w in layout.warnings]

    def exists(rel: str) -> bool:
        return (folder / rel).is_file()

    # 1. locate each entity's file: state.json path first, then the computed path.
    located: dict[str, str] = {}
    claimed: dict[str, str] = {}
    for e in layout.entities:
        rel = state["files"].get(e.entity_id, {}).get("path")
        if rel and exists(rel) and _norm(rel) not in claimed:
            located[e.entity_id] = rel
            claimed[_norm(rel)] = e.entity_id
    for e in layout.entities:
        if e.entity_id in located:
            continue
        if exists(e.rel_path) and _norm(e.rel_path) not in claimed:
            located[e.entity_id] = e.rel_path
            claimed[_norm(e.rel_path)] = e.entity_id
        elif exists(e.rel_path):
            other = claimed[_norm(e.rel_path)]
            actions.append(Action(WARN, e.rel_path, f"{e.entity_id} and {other} both claim this file", e.entity_id))

    # 2. moves
    work: dict[str, FileWork] = {}
    sources = {_norm(rel) for rel in located.values()}
    for e in layout.entities:
        cur = located.get(e.entity_id)
        if cur is None:
            continue
        fw = FileWork(e, cur, cur, hybrid=_hybrid_columns(e))
        if cur != e.rel_path:
            tgt = _norm(e.rel_path)
            if exists(e.rel_path) and tgt != _norm(cur) and tgt not in sources:
                actions.append(Action(WARN, cur, f"cannot move to {e.rel_path}: a file is already there", e.entity_id))
            else:
                fw.move, fw.target = True, e.rel_path
                sources.discard(_norm(cur))
        work[e.entity_id] = fw

    # 3. missing files: create blank items/checklist, render generated documents
    for e in layout.entities:
        if e.entity_id in located:
            continue
        if e.kind in GENERATED:
            work[e.entity_id] = FileWork(e, None, e.rel_path, regenerate=True)
        elif e.kind in CREATABLE:
            if e.ext == "pdf":
                actions.append(Action(WARN, e.rel_path, f"{e.code} {e.title}: PDF item needs a file", e.entity_id))
                continue
            work[e.entity_id] = FileWork(e, None, e.rel_path, create=True, write=True, hybrid=_hybrid_columns(e))

    # 4. in-use check, then content inspection
    (folder / BINDER_DIR / WORK).mkdir(parents=True, exist_ok=True)
    scratch = Path(tempfile.mkdtemp(prefix="plan-", dir=folder / BINDER_DIR / WORK))  # one per plan: plans may overlap
    try:
        for eid, fw in list(work.items()):
            e = fw.entity
            if fw.current is None:
                if fw.regenerate:
                    actions.append(Action(REGENERATE, fw.target, f"create {e.title}", eid))
                continue
            path = folder / fw.current
            if (fw.move or e.ext in ("docx", "xlsx", "pdf")) and _in_use(path):
                actions.append(Action(WARN, fw.current, "open in Word/Excel by someone; skipped, re-sync later", eid))
                work[eid] = FileWork(e, fw.current, fw.current)
                continue
            if e.kind in GENERATED:
                _inspect_generated(fw, ctx, rows, state, scratch, actions)
            elif e.ext == "docx":
                _inspect_docx(fw, path, manifest, layout, actions)
            elif e.ext == "xlsx":
                _inspect_xlsx(fw, path, manifest, layout, rows, actions)
            elif e.ext == "pdf":
                _inspect_pdf(fw, folder, path, manifest, state, actions)
            if fw.move:
                actions.append(Action(MOVE, fw.current, f"-> {fw.target}", eid, fw.target))
    finally:
        shutil.rmtree(scratch, ignore_errors=True)  # the app's own scratch renders

    for eid, fw in work.items():
        if fw.create:
            actions.append(Action(CREATE, fw.target, f"{fw.entity.code or ''} {fw.entity.title}".strip(), eid))
            if fw.hybrid:
                actions.append(Action(CHECKLIST, fw.target, "rows for every document", eid))

    # copied documents whose source changed since the copy (SPEC §9.3)
    for e in layout.items():
        cf = e.node.get("copied_from") if e.node is not None else None
        if not cf or not cf.get("source_path"):
            continue
        src = Path(cf["source_path"])
        if not src.is_file():
            actions.append(Action(WARN, e.rel_path, f"copied from {src}, which can no longer be found", e.entity_id))
        elif cf.get("source_hash") and file_hash(src) != cf["source_hash"]:
            actions.append(Action(WARN, e.rel_path, f"source changed since copy ({src}); re-copy from the "
                                                    "document's details if you want the new version", e.entity_id))

    # 5. orphans: files on disk nobody claims (the template folder is the app's)
    claimed_now = {_norm(fw.current) for fw in work.values() if fw.current}
    claimed_now |= {_norm(rel) for rel in located.values()}
    skip = {BINDER_DIR, ctx.template_folder().relative_to(folder).as_posix()}
    for rel in _disk_files(folder, skip):
        if _norm(rel) not in claimed_now:
            actions.append(Action(ORPHAN, rel, "not in the TOC; left in place"))

    actions.sort(key=lambda a: ACTION_ORDER[a.kind])
    return Plan(folder, actions, layout, work, rows)


def _inspect_generated(fw: FileWork, ctx: GenContext, rows, state, scratch: Path, actions: list[Action]) -> None:
    e = fw.entity
    tpl = ctx.template_for(e)
    reasons = []
    if state["files"].get(e.entity_id, {}).get("template") not in (None, template_hash(tpl)):
        reasons.append("template changed")
    out = scratch / f"{e.entity_id.replace('#', '_')}.{e.ext}"
    _render_generated(e, ctx, out, rows)
    if content_signature(out) != content_signature(ctx.folder / fw.current):
        reasons.append("content differs from the TOC")
    if reasons:
        fw.regenerate = True
        actions.append(Action(REGENERATE, fw.target, f"{e.title}: {', '.join(reasons)}", e.entity_id))


def _pages(manifest: Manifest, state: dict) -> dict[str, int] | None:
    """Running page numbers from the last merged export, for the TOC's Page column (continuous mode)."""
    return state.get("pages") if manifest.binder.page_numbers == "continuous" else None


def _pdf_key(folder: Path, e: Entity, manifest: Manifest) -> str | None:
    spec = _spec_for(e, manifest)
    src = pdfdoc.source_path(folder, e.entity_id)
    if spec is None or not src.is_file():
        return None
    return pdfdoc.stamp_key(file_hash(src), spec)


def _inspect_pdf(fw: FileWork, folder: Path, path: Path, manifest: Manifest, state: dict,
                 actions: list[Action]) -> None:
    e, eid = fw.entity, fw.entity.entity_id
    spec = _spec_for(e, manifest)
    if spec is None:
        return  # locked header: the PDF is left exactly as it is
    src = pdfdoc.source_path(folder, eid)
    recorded = state["files"].get(eid, {})
    if not src.is_file() or (recorded.get("hash") and file_hash(path) != recorded["hash"]):
        fw.write = fw.pdf_restash = True
        actions.append(Action(STAMP, fw.target, f"header overlay -> {e.code} (original kept in .binder/pdf_sources)",
                              eid))
    elif recorded.get("pdf_key") != pdfdoc.stamp_key(file_hash(src), spec):
        fw.write = True
        actions.append(Action(STAMP, fw.target, f"header overlay -> {e.code}", eid))


def _inspect_docx(fw: FileWork, path: Path, manifest: Manifest, layout: Layout, actions: list[Action]) -> None:
    e, eid = fw.entity, fw.entity.entity_id
    st = word.inspect(path, _spec_for(e, manifest), layout.codes)
    if not st.header_ok:
        fw.write = True
        actions.append(Action(STAMP, fw.target, f"header -> {e.code}", eid))
    if st.stale_refs:
        fw.write = True
        detail = ", ".join(f"{r.target_id}: {r.current or '(empty)'} -> {r.desired}" for r in st.stale_refs)
        actions.append(Action(REFS, fw.target, detail, eid))
    for r in st.broken_refs:
        actions.append(Action(WARN, fw.target, f"broken reference to {r.target_id!r} in: {r.context!r}", eid))


def _inspect_xlsx(fw: FileWork, path: Path, manifest: Manifest, layout: Layout, rows, actions: list[Action]) -> None:
    e, eid = fw.entity, fw.entity.entity_id
    bad = excel.unsafe_parts(path)
    if bad:
        actions.append(Action(WARN, fw.current,
                              f"contains {', '.join(bad)} which cannot be saved safely; header and refs not updated", eid))
        return
    wb = excel.load(path)
    if fw.hybrid:
        # Rows first: a removed row's =bb_ formula becomes [removed] and must not count as a reference.
        ch = checklist.sync(wb, checklist_specs(layout, items_only=e.kind == ITEM), *fw.hybrid)
        if ch.error:
            actions.append(Action(WARN, fw.target, ch.error, eid))
        elif ch.any:
            fw.write = True
            actions.append(Action(CHECKLIST, fw.target, ch.describe(), eid))
    st = excel.inspect_wb(wb, _spec_for(e, manifest), rows)
    if not st.header_ok:
        fw.write = True
        actions.append(Action(STAMP, fw.target, f"header -> {e.code}", eid))
    if not st.refs_ok:
        fw.write = True
        actions.append(Action(REFS, fw.target, "_binder sheet and bb_ names", eid))
    for name, cells in st.broken.items():
        actions.append(Action(WARN, fw.target, f"broken reference {name} in {', '.join(cells)}", eid))


# ---- apply ------------------------------------------------------------------

def _timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%dT%H-%M-%S")


def _backup_dir(folder: Path) -> Path:
    base = folder / BINDER_DIR / BACKUPS
    ts = _timestamp()
    d = base / ts
    i = 1
    while d.exists():
        i += 1
        d = base / f"{ts}_{i}"
    return d


def write_backup_index(backup: Path, files: list[dict], created: list[str] | tuple = (), kind: str = "sync") -> None:
    """Every backup folder describes itself, so any of them can be restored."""
    backup.mkdir(parents=True, exist_ok=True)
    (backup / "index.json").write_text(json.dumps({
        "kind": kind, "time": datetime.now().isoformat(timespec="seconds"),
        "files": files, "created": list(created)}, indent=2), encoding="utf-8")


def backup_info(folder: Path | str, name: str) -> dict:
    """Name, kind, time and file counts of one backup, for the restore screen."""
    d = Path(folder) / BINDER_DIR / BACKUPS / name
    try:
        index = json.loads((d / "index.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"name": name, "kind": "unknown", "time": "", "files": 0, "created": 0, "restorable": False}
    return {"name": name, "kind": index.get("kind", "sync"), "time": index.get("time", ""),
            "files": len(index.get("files", [])), "created": len(index.get("created", [])), "restorable": True}


def _remove_empty_dirs(folder: Path, rels: set[str]) -> None:
    """Remove folders left empty by moves. Only empty directories; never files."""
    for rel in sorted(rels, key=lambda r: -r.count("/")):
        d = folder / rel
        while d != folder and d.is_dir() and not any(d.iterdir()):
            d.rmdir()
            d = d.parent


def _write_file(fw: FileWork, dst: Path, tmp: Path, p: Plan, manifest: Manifest) -> str | None:
    """Stamp/refs/checklist into `tmp`, verify the invariant, then replace `dst`. Returns a failure or None."""
    e = fw.entity
    spec = _spec_for(e, manifest)
    if e.ext == "pdf":
        src = pdfdoc.source_path(p.folder, e.entity_id)
        if fw.pdf_restash:
            src.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(dst, src)  # the previous original, if any, is in this sync's backup
        pdfdoc.stamp(src, tmp, spec)
        if not pdfdoc.invariant_ok(src, tmp):
            return f"{fw.target}: overlay would have changed page content; original kept (invariant check)"
        os.replace(tmp, dst)
        return None
    if e.ext == "docx":
        before = word.body_hash(dst)
        word.update(dst, tmp, spec, p.layout.codes)
        after = word.body_hash(tmp)
    elif fw.hybrid:
        wb = excel.load(dst)
        before = checklist.user_hash(wb, *fw.hybrid)
        ch = checklist.sync(wb, checklist_specs(p.layout, items_only=e.kind == ITEM), *fw.hybrid)
        excel.apply_to_wb(wb, spec, p.rows)
        wb.save(str(tmp))
        after = checklist.user_hash(excel.load(tmp), *fw.hybrid)
        if ch.error:
            return f"{fw.target}: {ch.error}"
    else:
        before = excel.body_hash(dst)
        excel.update(dst, tmp, spec, p.rows)
        after = excel.body_hash(tmp)
    if after != before:
        return f"{fw.target}: body changed during write; original kept (invariant check)"
    os.replace(tmp, dst)
    return None


def apply(p: Plan, force_lock: bool = False) -> ApplyResult:
    folder = p.folder
    storage = LocalFolderStorage(folder)
    manifest = storage.open()
    tree = Tree(manifest)
    ctx = GenContext(folder, manifest, tree, p.layout, pages=_pages(manifest, load_state(folder)))
    acquired = storage.lock(force=force_lock, max_age_hours=manifest.binder.lock_max_age_hours)
    result = ApplyResult(backup=None)
    work_dir = folder / BINDER_DIR / WORK
    try:
        # 1. snapshot every existing file that will be moved or modified
        touched = [fw for fw in p.work.values() if fw.current and (fw.move or fw.write or fw.regenerate)]
        backup = _backup_dir(folder)
        index = []
        for fw in touched:
            dest = backup / fw.current
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(folder / fw.current, dest)
            index.append({"path": fw.current, "moved_to": fw.target if fw.move else None,
                          "entity_id": fw.entity.entity_id})
        for fw in touched:
            stash = pdfdoc.source_path(folder, fw.entity.entity_id)
            if fw.pdf_restash and stash.is_file():
                dest = backup / BINDER_DIR / pdfdoc.SOURCES / stash.name
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(stash, dest)
        created = [fw for fw in p.work.values() if fw.current is None]
        if index or created:
            write_backup_index(backup, index, [fw.target for fw in created])
            # the state before this sync (the default empty state on a first sync), so restore puts it back
            (backup / "_state.json").write_text(json.dumps(load_state(folder), indent=2), encoding="utf-8")
            result.backup = backup.name

        # 2. moves: two-phase through temporary names so chains and swaps never collide
        moves = [fw for fw in touched if fw.move]
        staged = []
        for i, fw in enumerate(moves):
            src = folder / fw.current
            before = file_hash(src)
            tmp = src.with_name(f"{src.name}.bbmove{i}")
            try:
                os.rename(src, tmp)
                staged.append((fw, tmp, before))
            except OSError as e:
                result.failures.append(f"move {fw.current}: {e}")
                fw.target, fw.move = fw.current, False
        emptied = set()
        for fw, tmp, before in staged:
            dst = folder / fw.target
            dst.parent.mkdir(parents=True, exist_ok=True)
            try:
                if dst.exists():
                    raise FileExistsError(f"{fw.target} already exists")
                os.rename(tmp, dst)
                if file_hash(dst) != before:
                    raise OSError("content changed during move")
                if "/" in fw.current:
                    emptied.add(fw.current.rsplit("/", 1)[0])
                result.applied.append(Action(MOVE, fw.current, f"-> {fw.target}", fw.entity.entity_id, fw.target))
            except OSError as e:
                os.rename(tmp, folder / fw.current)
                result.failures.append(f"move {fw.current} -> {fw.target}: {e}")
                fw.target = fw.current
        _remove_empty_dirs(folder, emptied)

        # 3. creates
        for fw in created:
            if fw.regenerate:
                continue
            dst = folder / fw.target
            dst.parent.mkdir(parents=True, exist_ok=True)
            if dst.exists():
                result.failures.append(f"create {fw.target}: a file is already there")
                fw.write = False
                continue
            if fw.entity.kind == CKL:
                shutil.copy2(ctx.template(CKL_TEMPLATE), dst)
            elif fw.entity.ext == "docx":
                word.create_blank(dst)
            else:
                excel.create_blank(dst, fw.entity.title)
            result.applied.append(Action(CREATE, fw.target, fw.entity.title, fw.entity.entity_id))

        # 4. header stamps and refs, each through the body invariant check (hybrids last)
        work_dir.mkdir(parents=True, exist_ok=True)
        writes = sorted((fw for fw in p.work.values() if fw.write), key=lambda fw: fw.hybrid is not None)
        regens = [fw for fw in p.work.values() if fw.regenerate]
        for fw in [w for w in writes if not w.hybrid]:
            _apply_write(fw, p, manifest, work_dir, result)

        # 5. regenerated documents (app-owned: replaced wholesale, snapshot already taken)
        for fw in regens:
            e = fw.entity
            dst = folder / fw.target
            tmp = work_dir / f"gen_{e.entity_id.replace('#', '_')}.{e.ext}"
            try:
                _render_generated(e, ctx, tmp, p.rows)
                dst.parent.mkdir(parents=True, exist_ok=True)
                os.replace(tmp, dst)
                result.applied.append(Action(REGENERATE, fw.target, e.title, e.entity_id))
            except Exception as ex:  # noqa: BLE001 - report and keep going
                result.failures.append(f"regenerate {fw.target}: {ex}")
            finally:
                if tmp.exists():
                    tmp.unlink()  # our own scratch render

        # 6. checklist rows (hybrid workbooks)
        for fw in [w for w in writes if w.hybrid]:
            _apply_write(fw, p, manifest, work_dir, result)

        # 7. state
        state = load_state(folder)
        files = {}
        for e in p.layout.entities:
            fw = p.work.get(e.entity_id)
            rel = fw.target if fw else e.rel_path
            if rel and (folder / rel).is_file():
                entry = {"path": rel, "hash": file_hash(folder / rel)}
                if e.kind in GENERATED:
                    entry["template"] = template_hash(ctx.template_for(e))
                if e.ext == "pdf" and (key := _pdf_key(folder, e, manifest)):
                    entry["pdf_key"] = key
                files[e.entity_id] = entry
        state["files"] = files
        save_state(folder, state)

        user, machine = current_user()
        storage.append_log({
            "time": datetime.now().isoformat(timespec="seconds"),
            "user": user, "machine": machine,
            "actions": len(result.applied),
            "warnings": [a.detail for a in p.warnings],
            "failures": result.failures,
            "backup": result.backup,
        })
    finally:
        if acquired:
            storage.unlock()
    return result


def _apply_write(fw: FileWork, p: Plan, manifest: Manifest, work_dir: Path, result: ApplyResult) -> None:
    dst = p.folder / fw.target
    if not dst.exists():
        return
    e = fw.entity
    tmp = work_dir / f"{e.entity_id.replace('#', '_')}.{e.ext}"
    try:
        failure = _write_file(fw, dst, tmp, p, manifest)
        if failure:
            result.failures.append(failure)
        else:
            kind = CHECKLIST if fw.hybrid else STAMP
            result.applied.append(Action(kind, fw.target, "header/refs updated", e.entity_id))
    except Exception as ex:  # noqa: BLE001 - report and keep going; original untouched
        result.failures.append(f"{fw.target}: {ex}")
    finally:
        if tmp.exists():
            tmp.unlink()  # our own scratch copy


# ---- restore ------------------------------------------------------------------

def list_backups(folder: Path | str) -> list[str]:
    d = Path(folder) / BINDER_DIR / BACKUPS
    return sorted((x.name for x in d.iterdir() if x.is_dir()), reverse=True) if d.exists() else []


def restore(folder: Path | str, backup: str) -> list[str]:
    """Put every file in a backup back where it was before that sync.

    Files the sync moved are moved back; current contents are snapshotted into a
    new backup first. Files the sync created are left in place (never deleted).
    """
    folder = Path(folder)
    src = folder / BINDER_DIR / BACKUPS / backup
    index = json.loads((src / "index.json").read_text(encoding="utf-8"))
    storage = LocalFolderStorage(folder)
    acquired = storage.lock()
    restored = []
    try:
        safety = _backup_dir(folder)
        safety = safety.with_name(safety.name + "_before-restore")
        saved = []
        for entry in index["files"]:
            for rel in filter(None, (entry["path"], entry.get("moved_to"))):
                if (folder / rel).is_file():
                    d = safety / rel
                    d.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(folder / rel, d)
                    saved.append({"path": rel, "moved_to": None})
        write_backup_index(safety, saved, kind="before-restore")
        if state_path(folder).exists():
            shutil.copy2(state_path(folder), safety / "_state.json")
        for entry in index["files"]:
            orig, moved = entry["path"], entry.get("moved_to")
            if moved and (folder / moved).is_file() and not (folder / orig).exists():
                (folder / orig).parent.mkdir(parents=True, exist_ok=True)
                os.rename(folder / moved, folder / orig)
            (folder / orig).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src / orig, folder / orig)
            restored.append(orig)
        # kept PDF originals (pdf_sources) that the sync replaced
        stash_dir = src / BINDER_DIR / pdfdoc.SOURCES
        if stash_dir.is_dir():
            for f in stash_dir.iterdir():
                dst = folder / BINDER_DIR / pdfdoc.SOURCES / f.name
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(f, dst)
        if (src / "_state.json").exists():
            shutil.copy2(src / "_state.json", state_path(folder))
        user, machine = current_user()
        storage.append_log({"time": datetime.now().isoformat(timespec="seconds"), "user": user,
                            "machine": machine, "restore": backup, "files": restored})
    finally:
        if acquired:
            storage.unlock()
    return restored
