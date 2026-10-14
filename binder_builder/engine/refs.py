"""Find, resolve and claim references (SPEC §6, §9.2).

Claiming turns a typed code ("see the Top 10 List (A.03)") into a live reference:
a `bb:ref:<id>` content control in Word, or a `=bb_<id>` formula in Excel. The user
approves every hit. A claim may only wrap existing text, never change it, so it
has its own invariant (`claim_invariant_*`): the full text of the document,
including the text inside controls, must be identical before and after.
"""

from __future__ import annotations

import copy
import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from . import excel, word

CODE_RE = re.compile(r"(?<![A-Za-z0-9.])(?:[A-Z]{1,2}\.\d{2}(?:\.\d{2})*(?:\.[A-Z]{1,3})?|0\.\d{2})(?![A-Za-z0-9]|\.\d)")

_W_P = qn("w:p")
_W_R = qn("w:r")
_W_T = qn("w:t")
_W_SDT = qn("w:sdt")
_W_RPR = qn("w:rPr")


@dataclass
class Hit:
    file: str  # binder-relative path
    location: str  # "p<n>" (Word paragraph index) or "Sheet!A1"
    start: int  # character offset in the paragraph/cell text
    code: str
    target_id: str | None  # None: no document has this code
    context: str

    @property
    def key(self) -> str:
        return f"{self.file}|{self.location}|{self.start}|{self.code}"


@dataclass
class ClaimResult:
    claimed: list[Hit] = field(default_factory=list)
    skipped: list[tuple[Hit, str]] = field(default_factory=list)


# ---- scanning -------------------------------------------------------------------

def _in_bb_sdt(el) -> bool:
    p = el.getparent()
    while p is not None:
        if p.tag == _W_SDT and (word.sdt_tag(p) or "").startswith("bb:"):
            return True
        p = p.getparent()
    return False


def _paragraphs(doc) -> list:
    return [p for p in doc.element.body.iter(_W_P) if not _in_bb_sdt(p)]


def _run_segments(p) -> list[tuple[object, str, bool]]:
    """Direct child runs of a paragraph: (run, text, plain). Plain runs hold exactly one w:t."""
    segs = []
    for child in p:
        if child.tag != _W_R:
            continue
        content = [c for c in child if c.tag != _W_RPR]
        text = "".join((c.text or "") if c.tag == _W_T else "\t" if c.tag == qn("w:tab") else "" for c in content)
        plain = len(content) == 1 and content[0].tag == _W_T
        segs.append((child, text, plain))
    return segs


def scan_docx(path: Path, rel: str, code_to_id: dict[str, str]) -> list[Hit]:
    doc = Document(str(path))
    hits = []
    for i, p in enumerate(_paragraphs(doc)):
        text = "".join(t for _, t, _ in _run_segments(p))
        for m in CODE_RE.finditer(text):
            hits.append(Hit(rel, f"p{i}", m.start(), m.group(0), code_to_id.get(m.group(0)), text.strip()))
    return hits


def scan_xlsx(path: Path, rel: str, code_to_id: dict[str, str]) -> list[Hit]:
    wb = excel.load(path)
    hits = []
    for ws in excel._user_sheets(wb):
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or v.startswith("="):
                    continue
                for m in CODE_RE.finditer(v):
                    hits.append(Hit(rel, f"{ws.title}!{c.coordinate}", m.start(), m.group(0),
                                    code_to_id.get(m.group(0)), v.strip()))
    return hits


def scan(folder: Path, files: list[str], code_to_id: dict[str, str]) -> list[Hit]:
    hits: list[Hit] = []
    for rel in files:
        p = folder / rel
        if p.suffix.lower() == ".docx":
            hits.extend(scan_docx(p, rel, code_to_id))
        elif p.suffix.lower() == ".xlsx" and not excel.unsafe_parts(p):
            hits.extend(scan_xlsx(p, rel, code_to_id))
    return hits


# ---- claiming: Word ---------------------------------------------------------------

def _split_run(run, at: int):
    """Split a plain run at character `at`; returns (left, right) with the right part in place."""
    t = run.find(_W_T)
    text = t.text or ""
    left = copy.deepcopy(run)
    lt = left.find(_W_T)
    lt.text = text[:at]
    lt.set(qn("xml:space"), "preserve")
    t.text = text[at:]
    t.set(qn("xml:space"), "preserve")
    run.addprevious(left)
    return left, run


def claim_paragraph(p, start: int, end: int, target_id: str) -> str | None:
    """Wrap text [start, end) of a paragraph in a bb:ref control. Returns a reason if impossible."""
    pos = 0
    inside = []
    for run, text, plain in _run_segments(p):
        r0, r1 = pos, pos + len(text)
        pos = r1
        if r1 <= start or r0 >= end:
            continue
        if not plain:
            return "the code is split across a tab, field or other non-text content"
        if r0 < start:
            _, run = _split_run(run, start - r0)
            r0 = start
            text = run.find(_W_T).text
        if r1 > end:
            left, _ = _split_run(run, end - r0)
            run = left
        inside.append(run)
    if not inside:
        return "text not found"
    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    tag = OxmlElement("w:tag")
    tag.set(qn("w:val"), f"{word.TAG_REF_PREFIX}{target_id}")
    pr.append(tag)
    lock = OxmlElement("w:lock")
    lock.set(qn("w:val"), "sdtContentLocked")
    pr.append(lock)
    sdt.append(pr)
    content = OxmlElement("w:sdtContent")
    sdt.append(content)
    inside[0].addprevious(sdt)
    for run in inside:
        content.append(run)
    return None


def claim_docx(src: Path, dest: Path, hits: list[Hit]) -> ClaimResult:
    doc = Document(str(src))
    paras = _paragraphs(doc)
    res = ClaimResult()
    # right to left within each paragraph, so earlier offsets stay valid
    for h in sorted(hits, key=lambda h: (int(h.location[1:]), -h.start)):
        if h.target_id is None:
            res.skipped.append((h, "no document has this code"))
            continue
        reason = claim_paragraph(paras[int(h.location[1:])], h.start, h.start + len(h.code), h.target_id)
        if reason:
            res.skipped.append((h, reason))
        else:
            res.claimed.append(h)
    doc.save(str(dest))
    return res


def full_text_docx(path: Path) -> str:
    """All body text including the text inside every control (the claim invariant)."""
    doc = Document(str(path))
    parts = []
    for p in doc.element.body.iter(_W_P):
        parts.append("".join(t.text or "" for t in p.iter(_W_T)))
    return hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()


# ---- claiming: Excel --------------------------------------------------------------

def _xl_str(s: str) -> str:
    return '"' + s.replace('"', '""') + '"'


def claim_cell_formula(value: str, hits: list[Hit]) -> str:
    """`A.03` -> `=bb_x`; `see A.03 now` -> `="see "&bb_x&" now"`."""
    pieces = []
    pos = 0
    for h in sorted(hits, key=lambda h: h.start):
        if h.start > pos:
            pieces.append(_xl_str(value[pos:h.start]))
        pieces.append(excel.name_for(h.target_id))
        pos = h.start + len(h.code)
    if pos < len(value):
        pieces.append(_xl_str(value[pos:]))
    return "=" + "&".join(pieces)


def claim_xlsx(src: Path, dest: Path, hits: list[Hit], rows) -> ClaimResult:
    wb = excel.load(src)
    res = ClaimResult()
    by_cell: dict[str, list[Hit]] = {}
    for h in hits:
        if h.target_id is None:
            res.skipped.append((h, "no document has this code"))
        else:
            by_cell.setdefault(h.location, []).append(h)
    for loc, cell_hits in by_cell.items():
        sheet, coord = loc.rsplit("!", 1)
        cell = wb[sheet][coord]
        if not isinstance(cell.value, str):
            res.skipped.extend((h, "cell changed since the scan") for h in cell_hits)
            continue
        cell.value = claim_cell_formula(cell.value, cell_hits)
        res.claimed.extend(cell_hits)
    # the names the new formulas use must exist
    excel.apply_to_wb(wb, None, rows)
    wb.save(str(dest))
    return res


def xlsx_cells(path: Path) -> dict[str, object]:
    wb = excel.load(path)
    return {f"{ws.title}!{c.coordinate}": excel._norm(c.value)
            for ws in excel._user_sheets(wb) for row in ws.iter_rows() for c in row if c.value is not None}


def claim_invariant_xlsx(before: dict, after: dict, claimed: list[Hit]) -> bool:
    """Only the claimed cells may differ, and each must now be a formula."""
    changed = {k for k in before.keys() | after.keys() if before.get(k) != after.get(k)}
    allowed = {h.location for h in claimed}
    return changed <= allowed and all(str(after.get(k, "")).startswith("str:'=") for k in changed)
