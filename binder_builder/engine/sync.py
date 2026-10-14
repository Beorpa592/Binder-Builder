"""Sync: plan, apply with backup and invariant check, state, log, restore (SPEC §7).

Safety rules (SPEC §1.1) enforced here:
  * no user file is ever deleted; removed TOC rows leave files in place as orphans
  * every change is listed in the plan before apply
  * every file about to be modified or moved is copied into .binder/backups/<ts>/ first
  * every Word/Excel write goes to a temp file, is body-hashed against the original,
    and only replaces the original if the hashes match
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from . import excel, word
from .codes import CKL, ITEM, SECTION_COVER, TOC, Entity, Layout, compute_layout, header_spec
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

# Entities created blank when missing in this phase (covers and tabs are generated in phase 1).
CREATABLE = (ITEM, SECTION_COVER, TOC, CKL)
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
    write: bool = False  # header stamp and/or ref update needed


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
    except PermissionError:
        return True
    except OSError:
        return True


def _disk_files(folder: Path) -> list[str]:
    out = []
    for root, dirs, files in os.walk(folder):
        rel_root = Path(root).relative_to(folder)
        dirs[:] = [d for d in dirs if not (rel_root == Path(".") and d == BINDER_DIR)]
        for f in files:
            if f.startswith("~$") or f.lower() in IGNORED_NAMES:
                continue
            out.append((rel_root / f).as_posix())
    return sorted(out)


def binder_rows(tree: Tree, layout: Layout) -> list[tuple[str, str, str, str]]:
    """Rows of the `_binder` sheet: one per coded node, in tree order."""
    rows = []
    folders = {s.node.id: s.folder.rsplit("/", 1)[-1] for s in layout.sections}
    for n in tree.walk():
        code = layout.codes.get(n.id)
        if code is None:
            continue
        e = layout.entity(n.id)
        fname = e.filename if e else folders.get(n.id, "")
        rows.append((n.id, code, n.text, fname))
    return rows


def _spec_for(e: Entity, manifest: Manifest):
    if not e.stamp or e.locked_header:
        return None
    return header_spec(e, manifest.binder)


# ---- plan -------------------------------------------------------------------

def plan(folder: Path | str) -> Plan:
    folder = Path(folder)
    manifest = LocalFolderStorage(folder).open()
    tree = Tree(manifest)
    layout = compute_layout(tree, manifest.binder)
    rows = binder_rows(tree, layout)
    state = load_state(folder)
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
        fw = FileWork(e, cur, cur)
        if cur != e.rel_path:
            tgt = _norm(e.rel_path)
            if exists(e.rel_path) and tgt != _norm(cur) and tgt not in sources:
                actions.append(Action(WARN, cur, f"cannot move to {e.rel_path}: a file is already there", e.entity_id))
            else:
                fw.move, fw.target = True, e.rel_path
                sources.discard(_norm(cur))
        work[e.entity_id] = fw

    # 3. creates
    for e in layout.entities:
        if e.entity_id in located:
            continue
        if e.kind not in CREATABLE:
            continue
        if e.ext == "pdf":
            actions.append(Action(WARN, e.rel_path, f"{e.code} {e.title}: PDF item needs a file", e.entity_id))
            continue
        work[e.entity_id] = FileWork(e, None, e.rel_path, create=True, write=True)

    # 4. in-use check, then content inspection
    for eid, fw in list(work.items()):
        if fw.create:
            continue
        path = folder / fw.current
        if (fw.move or fw.entity.ext in ("docx", "xlsx")) and _in_use(path):
            actions.append(Action(WARN, fw.current, "open in Word/Excel by someone; skipped, re-sync later", eid))
            work[eid] = FileWork(fw.entity, fw.current, fw.current)
            continue
        e = fw.entity
        spec = _spec_for(e, manifest)
        if e.ext == "docx":
            st = word.inspect(path, spec, layout.codes)
            if not st.header_ok:
                fw.write = True
                actions.append(Action(STAMP, fw.target, f"header -> {e.code}", eid))
            if st.stale_refs:
                fw.write = True
                detail = ", ".join(f"{r.target_id}: {r.current or '(empty)'} -> {r.desired}" for r in st.stale_refs)
                actions.append(Action(REFS, fw.target, detail, eid))
            for r in st.broken_refs:
                actions.append(Action(WARN, fw.target, f"broken reference to {r.target_id!r} in: {r.context!r}", eid))
        elif e.ext == "xlsx":
            bad = excel.unsafe_parts(path)
            if bad:
                actions.append(Action(WARN, fw.current,
                                      f"contains {', '.join(bad)} which cannot be saved safely; header and refs not updated",
                                      eid))
            else:
                st = excel.inspect(path, spec, rows)
                if not st.header_ok:
                    fw.write = True
                    actions.append(Action(STAMP, fw.target, f"header -> {e.code}", eid))
                if not st.refs_ok:
                    fw.write = True
                    actions.append(Action(REFS, fw.target, "_binder sheet and bb_ names", eid))
                for name, cells in st.broken.items():
                    actions.append(Action(WARN, fw.target, f"broken reference {name} in {', '.join(cells)}", eid))
        if fw.move:
            actions.append(Action(MOVE, fw.current, f"-> {fw.target}", eid, fw.target))

    for eid, fw in work.items():
        if fw.create:
            actions.append(Action(CREATE, fw.target, f"{fw.entity.code or ''} {fw.entity.title}".strip(), eid))

    # 5. orphans: files on disk nobody claims
    claimed_now = {_norm(fw.current) for fw in work.values() if fw.current}
    claimed_now |= {_norm(rel) for rel in located.values()}
    for rel in _disk_files(folder):
        if _norm(rel) not in claimed_now:
            actions.append(Action(ORPHAN, rel, "not in the TOC; left in place"))

    actions.sort(key=lambda a: ACTION_ORDER[a.kind])
    return Plan(folder, actions, layout, work, rows)


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


def _remove_empty_dirs(folder: Path, rels: set[str]) -> None:
    """Remove folders left empty by moves. Only empty directories; never files."""
    for rel in sorted(rels, key=lambda r: -r.count("/")):
        d = (folder / rel)
        while d != folder and d.is_dir() and not any(d.iterdir()):
            d.rmdir()
            d = d.parent


def apply(p: Plan, force_lock: bool = False) -> ApplyResult:
    folder = p.folder
    storage = LocalFolderStorage(folder)
    manifest = storage.open()
    acquired = storage.lock(force=force_lock, max_age_hours=manifest.binder.lock_max_age_hours)
    result = ApplyResult(backup=None)
    try:
        touched = [fw for fw in p.work.values() if not fw.create and (fw.move or fw.write)]
        backup = _backup_dir(folder)
        index = []
        if touched:
            for fw in touched:
                dest = backup / fw.current
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(folder / fw.current, dest)
                index.append({"path": fw.current, "moved_to": fw.target if fw.move else None,
                              "entity_id": fw.entity.entity_id})
        created = [fw for fw in p.work.values() if fw.create]
        if index or created:
            backup.mkdir(parents=True, exist_ok=True)
            (backup / "index.json").write_text(json.dumps({
                "files": index, "created": [fw.target for fw in created]}, indent=2), encoding="utf-8")
            result.backup = backup.name
            if state_path(folder).exists():
                shutil.copy2(state_path(folder), backup / "_state.json")

        # moves: two-phase through temporary names so chains and swaps never collide
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
                emptied.add(fw.current.rsplit("/", 1)[0] if "/" in fw.current else "")
                result.applied.append(Action(MOVE, fw.current, f"-> {fw.target}", fw.entity.entity_id, fw.target))
            except OSError as e:
                os.rename(tmp, folder / fw.current)
                result.failures.append(f"move {fw.current} -> {fw.target}: {e}")
                fw.target = fw.current
        _remove_empty_dirs(folder, {r for r in emptied if r})

        # creates
        for fw in created:
            dst = folder / fw.target
            dst.parent.mkdir(parents=True, exist_ok=True)
            if dst.exists():
                result.failures.append(f"create {fw.target}: a file is already there")
                continue
            if fw.entity.ext == "docx":
                word.create_blank(dst)
            else:
                excel.create_blank(dst, fw.entity.title)
            result.applied.append(Action(CREATE, fw.target, fw.entity.title, fw.entity.entity_id))

        # header stamps and refs, each through the body invariant check
        work_dir = folder / BINDER_DIR / WORK
        work_dir.mkdir(parents=True, exist_ok=True)
        for fw in p.work.values():
            if not fw.write:
                continue
            dst = folder / fw.target
            if not dst.exists():
                continue
            e = fw.entity
            spec = _spec_for(e, manifest)
            tmp = work_dir / f"{e.entity_id.replace('#', '_')}.{e.ext}"
            try:
                before = body_hash(dst)
                if e.ext == "docx":
                    word.update(dst, tmp, spec, p.layout.codes)
                else:
                    excel.update(dst, tmp, spec, p.rows)
                after = body_hash(tmp)
                if after != before:
                    result.failures.append(f"{fw.target}: body changed during write; original kept (invariant check)")
                    continue
                os.replace(tmp, dst)
                result.applied.append(Action(STAMP, fw.target, "header/refs updated", e.entity_id))
            except Exception as ex:  # noqa: BLE001 - report and keep going; original untouched
                result.failures.append(f"{fw.target}: {ex}")
            finally:
                if tmp.exists():
                    tmp.unlink()  # our own scratch copy

        # state
        state = load_state(folder)
        files = {}
        for e in p.layout.entities:
            fw = p.work.get(e.entity_id)
            rel = fw.target if fw else None
            if rel is None and (folder / e.rel_path).is_file():
                rel = e.rel_path
            if rel and (folder / rel).is_file():
                files[e.entity_id] = {"path": rel, "hash": file_hash(folder / rel)}
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
        for entry in index["files"]:
            for rel in filter(None, (entry["path"], entry.get("moved_to"))):
                if (folder / rel).is_file():
                    d = safety / rel
                    d.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(folder / rel, d)
        for entry in index["files"]:
            orig, moved = entry["path"], entry.get("moved_to")
            if moved and (folder / moved).is_file() and not (folder / orig).exists():
                (folder / orig).parent.mkdir(parents=True, exist_ok=True)
                os.rename(folder / moved, folder / orig)
            (folder / orig).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src / orig, folder / orig)
            restored.append(orig)
        if (src / "_state.json").exists():
            shutil.copy2(src / "_state.json", state_path(folder))
        else:
            save_state(folder, {**load_state(folder), "files": {}})
        user, machine = current_user()
        storage.append_log({"time": datetime.now().isoformat(timespec="seconds"), "user": user,
                            "machine": machine, "restore": backup, "files": restored})
    finally:
        if acquired:
            storage.unlock()
    return restored
