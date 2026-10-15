"""Print manifest for Packet Printer, and changed-only printing (SPEC §11).

`.binder/print_manifest.json` is an ordered JSON list of
`{path, pdf_path, code, title, pages, changed_since_last_print}` with absolute
paths. `state.json["printed"]` records each file's hash when it was last
printed. Printing marks the selected files as printed.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

from ..engine import excel, sync, word
from ..engine.manifest import BINDER_DIR

MANIFEST_NAME = "print_manifest.json"


def manifest_path(folder: Path) -> Path:
    return Path(folder) / BINDER_DIR / MANIFEST_NAME


def print_signature(path: Path) -> str:
    """Hash of what prints, not of the file's bytes.

    A workbook's hidden `_binder` sheet changes whenever any code in the binder changes; that
    must not make every workbook "changed since last print". Word: body (including reference
    text) plus header/footer. Excel: every visible sheet's header/footer and cells, with
    `=bb_` references replaced by the code they show. PDF: the file itself.
    """
    ext = path.suffix.lower()
    if ext == ".docx":
        text = word.full_signature(path)
    elif ext == ".xlsx":
        wb = excel.load(path)
        codes = {excel.name_for(r[0]).lower(): r[1] for r in (excel.read_binder_rows(wb) or [])}
        out = []
        for ws in excel._user_sheets(wb):
            if ws.sheet_state != "visible":
                continue
            out.append(f"## {ws.title}")
            for hdr, ftr in excel._header_targets(ws):
                out.append(f"{excel._parts(hdr)}{excel._parts(ftr)}")
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if isinstance(v, str) and v.startswith("="):
                        v = excel._NAME_REF.sub(lambda m: repr(codes.get(m.group(1).lower(), m.group(1))), v)
                    if v is not None:
                        out.append(f"{c.coordinate}={excel._norm(v)}")
        text = "\n".join(out)
    else:
        return sync.file_hash(path)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def changed_since_last_print(folder: Path, state: dict, entity_id: str, path: Path) -> bool:
    printed = state.get("printed", {}).get(entity_id)
    return printed is None or not path.is_file() or print_signature(path) != printed


def write(folder: Path, entries: list[dict]) -> Path:
    p = manifest_path(folder)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(entries, indent=2), encoding="utf-8")
    return p


def read(folder: Path) -> list[dict]:
    return json.loads(manifest_path(folder).read_text(encoding="utf-8"))


def mark_printed(folder: Path, entity_ids: list[str]) -> None:
    """Record the current version of each file as printed (after the user sends it to print)."""
    folder = Path(folder)
    state = sync.load_state(folder)
    printed = state.setdefault("printed", {})
    for eid in entity_ids:
        info = state["files"].get(eid)
        if info and (folder / info["path"]).is_file():
            printed[eid] = print_signature(folder / info["path"])
    sync.save_state(folder, state)
