from pathlib import Path

from binder_builder.engine import sync
from binder_builder.engine.manifest import Manifest
from binder_builder.engine.tree import Tree


def do_sync(folder: Path) -> sync.ApplyResult:
    p = sync.plan(folder)
    r = sync.apply(p)
    assert r.ok, r.failures
    return r


def edit_tree(folder: Path):
    """Context-free helper: returns (manifest, tree); call manifest.save() when done."""
    m = Manifest.load(folder)
    return m, Tree(m)


def office_files(folder: Path) -> list[Path]:
    return sorted(p for p in folder.rglob("*") if p.suffix in (".docx", ".xlsx") and ".binder" not in p.parts)


def author_files(folder: Path) -> list[str]:
    """Relative paths of the author-owned documents (tree items), per the current layout."""
    from binder_builder.engine.codes import compute_layout
    m, tree = edit_tree(folder)
    return [e.rel_path for e in compute_layout(tree, m.binder).items() if e.ext in ("docx", "xlsx")]


def body_hashes_by_entity(folder: Path) -> dict[str, str]:
    """Body hash of every author-owned Office file (tree items), keyed by entity id via state.json.

    Generated documents (TOC, covers, tabs) are app-owned and excluded; the CKL is hybrid."""
    state = sync.load_state(folder)
    out = {}
    for eid, info in state["files"].items():
        if "#" in eid:
            continue
        p = folder / info["path"]
        if p.suffix in (".docx", ".xlsx"):
            out[eid] = sync.body_hash(p)
    return out
