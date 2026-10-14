"""Import an existing binder folder (SPEC §9.1).

    plan = analyze(folder)        # reads only; changes nothing on disk
    plan.set_pair(eid, rel)       # the user fixes pairs in the pairing table
    accept(plan)                  # writes .binder/ (manifest, state, import log) and nothing else
    hits = claim_scan(folder)     # SPEC §9.2: code-shaped text that could become live references
    apply_claims(folder, approved)

The first sync then goes through the normal plan/review/apply flow. Where the TOC
and a file disagree, the TOC wins and the disagreement is listed.
"""

from __future__ import annotations

import difflib
import json
import re
import shutil
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from ruamel.yaml.comments import CommentedMap, CommentedSeq

from . import refs, sync
from .codes import ITEM, SECTION_COVER, compute_layout, filename_part
from .generate import DEFAULT_TEMPLATE_FOLDER
from .manifest import BINDER_DIR, Manifest, _yaml
from .storage import LocalFolderStorage
from .tree import Tree, slugify

IMPORT_LOG = "import.json"
SECTION_RE = re.compile(r"^(0|[A-Z]{1,2})$")
ITEM_RE = re.compile(r"^(0|[A-Z]{1,2})(\.\d{2,3})+(\.[A-Z]{1,3})?$")
HEADER_CODE_RE = re.compile(r"([A-Za-z][A-Za-z ]*#:?)\s*((?:0|[A-Z]{1,2})(?:\.(?:\d{2,3}|[A-Z]{1,3}))*)\b")
OFFICE_EXT = (".docx", ".xlsx", ".pdf")
MIN_SIM_WITH_CODE = 0.5
MIN_SIM_TITLE_ONLY = 0.85


# ---- TOC parsing ---------------------------------------------------------------------

@dataclass
class TocRow:
    row: int
    kind: str  # section | header | item
    code: str | None
    title: str
    included: bool = True


def _code_text(v) -> str | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        s = str(v)
    elif isinstance(v, float):
        s = f"{v:.2f}"  # legacy TOCs store 0.01 as a float
    elif isinstance(v, str):
        s = v.strip()
    else:
        return None
    return s if SECTION_RE.match(s) or ITEM_RE.match(s) else None


def find_toc(folder: Path) -> Path | None:
    cands = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx" and "toc" in p.stem.lower()
             and not p.name.startswith("~$")]
    return sorted(cands, key=lambda p: (not p.name.startswith("01_"), p.name))[0] if cands else None


def read_toc(path: Path) -> list[TocRow]:
    ws = load_workbook(path, data_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    marks = Counter(i for r in rows for i, v in enumerate(r) if isinstance(v, str) and v.strip().lower() == "x")
    mark_col = marks.most_common(1)[0][0] if marks and marks.most_common(1)[0][1] >= 2 else None
    out: list[TocRow] = []
    started = False
    for n, r in enumerate(rows, start=1):
        cells = [(i, v) for i, v in enumerate(r) if v not in (None, "") and i != mark_col]
        if not cells:
            continue
        marked = mark_col is not None and mark_col < len(r) and isinstance(r[mark_col], str) \
            and r[mark_col].strip().lower() == "x"
        code = _code_text(cells[0][1])
        texts = [str(v).strip() for _, v in cells[1:] if isinstance(v, str) and v.strip()]
        if code and SECTION_RE.match(code) and texts:
            out.append(TocRow(n, "section", code, texts[0]))
            started = True
        elif code and ITEM_RE.match(code) and texts:
            out.append(TocRow(n, "item", code, texts[0], included=marked))
            started = True
        elif started and isinstance(cells[0][1], str) and len(cells) == 1:
            out.append(TocRow(n, "header", None, cells[0][1].strip()))
    return out


def build_tree(rows: list[TocRow]) -> tuple[CommentedSeq, dict[str, str], list[str]]:
    """TOC rows -> raw manifest tree. Returns (tree, {node id: TOC code}, warnings).

    A clarifying header groups every section after it until the next header.
    """
    top = CommentedSeq()
    toc_codes: dict[str, str] = {}
    warnings: list[str] = []
    ids: set[str] = set()
    sections: dict[str, CommentedMap] = {}
    items: dict[str, CommentedMap] = {}
    header: CommentedMap | None = None
    last_section: CommentedMap | None = None

    def new_id(text: str) -> str:
        base = slugify(text)
        i, nid = 1, base
        while nid in ids:
            i += 1
            nid = f"{base}-{i}"
        ids.add(nid)
        return nid

    def children(m: CommentedMap) -> CommentedSeq:
        if "children" not in m:
            m["children"] = CommentedSeq()
        return m["children"]

    for r in rows:
        if r.kind == "header":
            header = CommentedMap(id=new_id(r.title), kind="header", label=r.title)
            top.append(header)
        elif r.kind == "section":
            node = CommentedMap(id=new_id(r.title), kind="section", label=r.title)
            if r.code == "0":
                node["numeric"] = True
            (children(header) if header is not None else top).append(node)
            sections[r.code] = node
            last_section = node
            toc_codes[node["id"]] = r.code
        else:
            parts = r.code.split(".")
            suffix = parts[-1] if parts[-1].isalpha() else None
            base_parts = parts[:-1] if suffix else parts
            parent_code = ".".join(base_parts[:-1]) if not suffix else ".".join(base_parts[:-1])
            parent = sections.get(parent_code) or items.get(parent_code)
            if parent is None:
                warnings.append(f"TOC row {r.row}: no parent for {r.code}; placed in the last section")
                parent = last_section
            if parent is None:
                warnings.append(f"TOC row {r.row}: {r.code} appears before any section; skipped")
                continue
            node = CommentedMap(id=new_id(r.title), kind="item", title=r.title, ext="docx")
            if suffix:
                node["suffix"] = suffix
            if not r.included:
                node["included"] = False
            children(parent).append(node)
            items[r.code] = node
            toc_codes[node["id"]] = r.code
    return top, toc_codes, warnings


# ---- file inspection ---------------------------------------------------------------------

def _title_key(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", s.lower().replace("&", " and ")).strip()


def similarity(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, _title_key(a), _title_key(b)).ratio()


_NAME_RE = re.compile(r"^(?P<ord>\d{2,3}(?:_\d{2,3})*)_(?:(?P<suf>[A-Z]{1,3})_)?(?P<rest>.+)$")


def split_filename(stem: str) -> tuple[str | None, str | None, str]:
    m = _NAME_RE.match(stem)
    if not m:
        return None, None, stem
    return m.group("ord"), m.group("suf"), m.group("rest")


def legacy_header(path: Path) -> list[str]:
    """Header lines as text (Word: section 1 primary header paragraphs; Excel: header blocks)."""
    try:
        if path.suffix.lower() == ".docx":
            sec = Document(str(path)).sections[0]
            if sec.header.is_linked_to_previous:
                return []
            return [p.text for p in sec.header.paragraphs]
        if path.suffix.lower() == ".xlsx":
            ws = load_workbook(path).worksheets[0]
            h = ws.oddHeader
            return [x for x in ((h.left.text or ""), (h.center.text or ""), (h.right.text or "")) if x]
    except Exception:  # noqa: BLE001 - unreadable header: no code from it
        return []
    return []


def header_code(lines: list[str]) -> tuple[str | None, str | None]:
    for line in lines:
        m = HEADER_CODE_RE.search(line.replace("&&", "&"))
        if m:
            return m.group(2), m.group(1).strip()
    return None, None


# ---- the import plan ---------------------------------------------------------------------

@dataclass
class Pair:
    entity_id: str
    kind: str
    code: str | None  # code from the TOC
    title: str
    file: str | None = None
    method: str | None = None  # "header code" | "filename" | "title" | "root" | "manual"


@dataclass
class Disagreement:
    entity_id: str
    file: str | None
    what: str
    file_says: str
    toc_says: str

    def describe(self) -> str:
        return f"{self.what}: file says {self.file_says}, TOC says {self.toc_says} (TOC wins)"


@dataclass
class ImportPlan:
    folder: Path
    toc_file: str | None
    manifest: Manifest
    toc_codes: dict[str, str]
    pairs: dict[str, Pair]
    files: list[str]
    headers: dict[str, list[str]] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    @property
    def unmatched_files(self) -> list[str]:
        used = {p.file for p in self.pairs.values() if p.file}
        return [f for f in self.files if f not in used]

    @property
    def unmatched_nodes(self) -> list[Pair]:
        return [p for p in self.pairs.values() if p.file is None]

    def set_pair(self, entity_id: str, rel: str | None) -> None:
        for p in self.pairs.values():
            if rel and p.file == rel and p.entity_id != entity_id:
                p.file, p.method = None, None
        pair = self.pairs[entity_id]
        pair.file, pair.method = rel, ("manual" if rel else None)

    def disagreements(self) -> list[Disagreement]:
        out = []
        layout = compute_layout(Tree(self.manifest), self.manifest.binder)
        for nid, toc_code in self.toc_codes.items():
            computed = layout.codes.get(nid)
            if computed and computed != toc_code:
                out.append(Disagreement(nid, None, "numbering", toc_code, computed))
        for p in self.pairs.values():
            if not p.file or p.kind != ITEM or not p.code:
                continue
            hc, _ = header_code(self.headers.get(p.file, []))
            if hc and hc != p.code:
                out.append(Disagreement(p.entity_id, p.file, "header code", hc, p.code))
            ordinal, _, _ = split_filename(Path(p.file).stem)
            want = "_".join(x for x in p.code.split(".")[1:] if x.isdigit())
            if ordinal and ordinal != want:
                out.append(Disagreement(p.entity_id, p.file, "filename number", ordinal, want))
            e = layout.entity(p.entity_id)
            if e and Path(p.file).parent.as_posix() != Path(e.rel_path).parent.as_posix():
                out.append(Disagreement(p.entity_id, p.file, "folder", Path(p.file).parent.as_posix(),
                                        Path(e.rel_path).parent.as_posix()))
        return out


def _binder_files(folder: Path) -> list[str]:
    out = []
    for p in sorted(folder.rglob("*")):
        rel = p.relative_to(folder)
        if p.is_file() and p.suffix.lower() in OFFICE_EXT and not p.name.startswith("~$") \
                and BINDER_DIR not in rel.parts:
            out.append(rel.as_posix())
    return out


def analyze(folder: Path | str) -> ImportPlan:
    """Read the TOC and every file, propose pairs. Writes nothing."""
    folder = Path(folder)
    warnings: list[str] = []
    toc = find_toc(folder)
    if toc is None:
        warnings.append("no TOC workbook found in the binder folder (a root .xlsx with 'TOC' in its name)")
        rows: list[TocRow] = []
    else:
        rows = read_toc(toc)
    tree_raw, toc_codes, w = build_tree(rows)
    warnings.extend(w)

    raw = _yaml().load(_new_manifest_text())
    raw["binder"]["name"] = folder.name
    raw["tree"] = tree_raw
    manifest = Manifest(raw)
    files = _binder_files(folder)
    headers = {rel: legacy_header(folder / rel) for rel in files if not rel.lower().endswith(".pdf")}

    layout = compute_layout(Tree(manifest), manifest.binder)
    pairs: dict[str, Pair] = {}
    for e in layout.entities:
        code = toc_codes.get(e.entity_id, e.code) if e.kind == ITEM else e.code
        pairs[e.entity_id] = Pair(e.entity_id, e.kind, code, e.title)

    taken: set[str] = set()
    root = [f for f in files if "/" not in f]

    def take(eid: str, rel: str | None, method: str) -> None:
        if rel and rel not in taken:
            pairs[eid].file, pairs[eid].method = rel, method
            taken.add(rel)

    # root documents
    if toc is not None:
        take("#toc", toc.name, "root")
    take("#ckl", next((f for f in root if f.lower().endswith(".xlsx") and "ckl" in f.lower()), None), "root")
    take("#binder-cover", next((f for f in root if f.lower().endswith(".docx") and f.startswith("00_")), None), "root")
    take("#tabs", next((f for f in root if f.lower().endswith(".docx") and "tab" in f.lower()), None), "root")

    # section folders and covers
    dirs = sorted({f.split("/")[0] for f in files if "/" in f})
    section_dir: dict[str, str] = {}
    for s in layout.sections:
        code = s.code
        by_prefix = [d for d in dirs if d.split("_", 1)[0] == code]
        best = by_prefix[0] if by_prefix else max(dirs, key=lambda d: similarity(d, s.node.text), default=None)
        if best and (by_prefix or similarity(best, s.node.text) >= MIN_SIM_TITLE_ONLY):
            section_dir[s.node.id] = best
            cover = next((f for f in files if f.startswith(best + "/") and f.count("/") == 1
                          and Path(f).name.startswith("00_") and f.lower().endswith(".docx")), None)
            take(f"{s.node.id}#cover", cover, "filename")

    items = [e for e in layout.entities if e.kind == ITEM]
    candidates = [f for f in files if f not in taken]

    def sim_to(e, rel: str) -> float:
        ordinal, suf, rest = split_filename(Path(rel).stem)
        return max(similarity(e.title, rest), similarity(e.title, f"{suf or ''} {rest}"))

    # pass 1: header code (with a sanity check on the title, so a stale code can't pair the wrong file)
    for e in items:
        code = pairs[e.entity_id].code
        best = [(sim_to(e, f), f) for f in candidates if f not in taken
                and header_code(headers.get(f, []))[0] == code and sim_to(e, f) >= MIN_SIM_WITH_CODE]
        if best:
            take(e.entity_id, max(best)[1], "header code")
    # pass 2: filename ordinal inside the section's folder, plus title similarity
    for e in items:
        if pairs[e.entity_id].file:
            continue
        code = pairs[e.entity_id].code or ""
        want = "_".join(x for x in code.split(".")[1:] if x.isdigit())
        sec = e.node
        while sec is not None and not sec.is_section:
            sec = sec.parent
        d = section_dir.get(sec.id) if sec else None
        best = [(sim_to(e, f), f) for f in candidates if f not in taken and d and f.startswith(d + "/")
                and split_filename(Path(f).stem)[0] == want and sim_to(e, f) >= MIN_SIM_WITH_CODE]
        if best:
            take(e.entity_id, max(best)[1], "filename")
    # pass 3: title similarity alone
    for e in items:
        if pairs[e.entity_id].file:
            continue
        best = [(sim_to(e, f), f) for f in candidates if f not in taken and sim_to(e, f) >= MIN_SIM_TITLE_ONLY]
        if best:
            take(e.entity_id, max(best)[1], "title")

    return ImportPlan(folder, toc.name if toc else None, manifest, toc_codes, pairs, files, headers, warnings)


def _new_manifest_text() -> str:
    from .manifest import NEW_BINDER_TEMPLATE

    return NEW_BINDER_TEMPLATE


# ---- accepting -----------------------------------------------------------------------------

def _infer_header(plan: ImportPlan) -> tuple[list[str], dict[str, str], str]:
    """Header lines, vars and ref label from the legacy Word headers (most common pattern)."""
    patterns: Counter = Counter()
    labels: Counter = Counter()
    org = Counter()
    for p in plan.pairs.values():
        if p.kind != ITEM or not p.file or not p.file.lower().endswith(".docx"):
            continue
        lines = [ln for ln in plan.headers.get(p.file, []) if ln.strip()]
        if not lines:
            continue
        _, label = header_code(lines)
        if label:
            labels[label] += 1
        first = HEADER_CODE_RE.sub("", lines[0]).strip(" \t")
        out = []
        for i, ln in enumerate([first] + lines[1:]):
            if i == 0 and first:
                org[first] += 1
                out.append("{{org_short}}")
            elif _title_key(ln) == _title_key(p.title):
                out.append("{{doc_title}}")
            else:
                out.append(ln.strip())
        patterns[tuple(out)] += 1
    if not patterns:
        return ["{{org_short}}", "{{binder_name}}", "{{doc_title}}"], {"org_short": plan.folder.name}, "Packet Ref #"
    lines = list(patterns.most_common(1)[0][0])
    vars_ = {"org_short": org.most_common(1)[0][0] if org else plan.folder.name}
    label = labels.most_common(1)[0][0] if labels else "Packet Ref #"
    return lines, vars_, label


def _title_part(rel: str, suffix: str | None = None) -> str:
    """Filename minus its computed prefix: `03_Top_10_List` -> `Top_10_List`, `03_PB_X` -> `X` for suffix PB.
    A leading capitalised word is only treated as a suffix when the node has that suffix."""
    stem = Path(rel).stem
    m = re.match(r"^\d{2,3}(?:_\d{2,3})*_", stem)
    rest = stem[m.end():] if m else stem
    if suffix and rest.startswith(f"{suffix}_"):
        rest = rest[len(suffix) + 1:]
    return rest


def accept(plan: ImportPlan) -> Path:
    """Write .binder/ for the confirmed pairing. Touches nothing outside .binder/."""
    folder = plan.folder
    if (folder / BINDER_DIR / "binder.yaml").exists():
        raise FileExistsError(f"{folder} already has a binder manifest")
    m = plan.manifest
    tree = Tree(m)
    lines, vars_, label = _infer_header(plan)
    b = m.raw["binder"]
    b["header_lines"] = lines
    b["vars"] = vars_
    b["ref_label"] = label

    # filename overrides so nothing is renamed on first sync unless the TOC says so
    names = CommentedMap()
    for eid, key, default in (("#toc", "toc", "TOC"), ("#ckl", "ckl", "CKL"), ("#binder-cover", "cover", "Cover")):
        f = plan.pairs[eid].file
        if f:
            stem = Path(f).stem
            part = stem.split("_", 1)[1] if re.match(r"^\d{2}_", stem) else stem
            if part != default:
                names[key] = part
    tabs = plan.pairs["#tabs"].file
    if tabs and Path(tabs).stem != "Tabs":
        names["tabs"] = Path(tabs).stem
    if names:
        b["filenames"] = names
    cover_file = plan.pairs["#binder-cover"].file
    if cover_file:
        try:
            paras = [p.text.strip() for p in Document(str(folder / cover_file)).paragraphs if p.text.strip()]
            if paras:
                b["cover"]["lines"] = paras[:6]
        except Exception:  # noqa: BLE001 - keep the default cover lines
            pass

    for p in plan.pairs.values():
        if not p.file:
            continue
        if p.kind == ITEM:
            node = tree.find(p.entity_id)
            node.set("ext", Path(p.file).suffix.lower().lstrip("."))
            part = _title_part(p.file, node.get("suffix"))
            if part != filename_part(node.text):
                node.set("filename", part)
        elif p.kind == SECTION_COVER:
            node = tree.find(p.entity_id.split("#")[0])
            part = _title_part(p.file)
            if part != "Cover":
                node.set("filename", part)
            folder_part = p.file.split("/")[0].split("_", 1)[-1]
            if folder_part != filename_part(node.text):
                node.set("folder", folder_part)
    m.validate()

    bdir = folder / BINDER_DIR
    bdir.mkdir(parents=True, exist_ok=True)
    m.save(bdir / "binder.yaml")
    sync.save_state(folder, {"format": 1, "printed": {},
                             "files": {eid: {"path": p.file} for eid, p in plan.pairs.items() if p.file}})
    if tabs:
        tdir = folder / DEFAULT_TEMPLATE_FOLDER
        tdir.mkdir(parents=True, exist_ok=True)
        if not (tdir / "Tabs.docx").exists():
            shutil.copy2(folder / tabs, tdir / "Tabs.docx")  # the legacy table layout becomes the template
    log = {
        "imported_at": datetime.now().isoformat(timespec="seconds"),
        "toc_file": plan.toc_file,
        "pairs": [asdict(p) for p in plan.pairs.values()],
        "unmatched_files": plan.unmatched_files,
        "disagreements": [asdict(d) for d in plan.disagreements()],
        "previous_headers": {p.file: plan.headers.get(p.file, []) for p in plan.pairs.values() if p.file},
        "warnings": plan.warnings,
        "unclaimed": [],
    }
    (bdir / IMPORT_LOG).write_text(json.dumps(log, indent=2), encoding="utf-8")
    return bdir / "binder.yaml"


# ---- claiming references after import (SPEC §9.2) ----------------------------------------

def claim_scan(folder: Path | str) -> list[refs.Hit]:
    folder = Path(folder)
    manifest = Manifest.load(folder)
    tree = Tree(manifest)
    layout = compute_layout(tree, manifest.binder)
    code_to_id = {code: nid for nid, code in layout.codes.items()}
    state = sync.load_state(folder)["files"]
    files = []
    for e in layout.entities:
        if e.kind != ITEM or e.ext == "pdf":
            continue  # generated docs are regenerated; the CKL adopts its rows itself
        rel = state.get(e.entity_id, {}).get("path") or e.rel_path
        if (folder / rel).is_file():
            files.append(rel)
    return refs.scan(folder, files, code_to_id)


def apply_claims(folder: Path | str, approved: list[refs.Hit], unapproved: list[refs.Hit] = ()) -> refs.ClaimResult:
    """Turn approved hits into live references, with backup and the claim invariant per file."""
    folder = Path(folder)
    storage = LocalFolderStorage(folder)
    manifest = storage.open()
    tree = Tree(manifest)
    layout = compute_layout(tree, manifest.binder)
    rows = sync.binder_rows(tree, layout)
    acquired = storage.lock()
    result = refs.ClaimResult()
    try:
        by_file: dict[str, list[refs.Hit]] = {}
        for h in approved:
            by_file.setdefault(h.file, []).append(h)
        backup = sync._backup_dir(folder)
        backup = backup.with_name(backup.name + "_claims")
        work = folder / BINDER_DIR / sync.WORK
        work.mkdir(parents=True, exist_ok=True)
        for rel, hits in by_file.items():
            src = folder / rel
            dest = backup / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            tmp = work / f"claim{src.suffix}"
            try:
                if src.suffix.lower() == ".docx":
                    before = refs.full_text_docx(src)
                    r = refs.claim_docx(src, tmp, hits)
                    ok = refs.full_text_docx(tmp) == before
                else:
                    before = refs.xlsx_cells(src)
                    r = refs.claim_xlsx(src, tmp, hits, rows)
                    ok = refs.claim_invariant_xlsx(before, refs.xlsx_cells(tmp), r.claimed)
                if not ok:
                    result.skipped.extend((h, "claim would have changed the text; file left alone") for h in hits)
                    continue
                shutil.move(str(tmp), str(src))
                result.claimed.extend(r.claimed)
                result.skipped.extend(r.skipped)
            finally:
                if tmp.exists():
                    tmp.unlink()  # our own scratch copy
        log_path = folder / BINDER_DIR / IMPORT_LOG
        if log_path.exists():
            log = json.loads(log_path.read_text(encoding="utf-8"))
            log["unclaimed"] = [asdict(h) for h in unapproved] + [dict(asdict(h), reason=why) for h, why in result.skipped]
            log["claimed"] = log.get("claimed", []) + [asdict(h) for h in result.claimed]
            log_path.write_text(json.dumps(log, indent=2), encoding="utf-8")
        user, machine = sync.current_user()
        storage.append_log({"time": datetime.now().isoformat(timespec="seconds"), "user": user, "machine": machine,
                            "claims": len(result.claimed), "skipped": len(result.skipped), "backup": backup.name})
    finally:
        if acquired:
            storage.unlock()
    return result

