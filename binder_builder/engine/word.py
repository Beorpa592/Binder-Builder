"""Word (.docx) support: content controls, header/footer writer, refs, body hash (SPEC §5.1, §6.1, §7.4).

The app owns section 1's primary (and, if present, first-page and even-page)
header and footer. Everything it writes there lives in content controls tagged
`bb:*`, so it can be found and compared later without guessing at runs.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .codes import HeaderSpec

TAG_CODE = "bb:code"
TAG_REF_PREFIX = "bb:ref:"
BROKEN = "[?]"

_W_SDT = qn("w:sdt")
_W_T = qn("w:t")
_W_TAB = qn("w:tab")
_W_R = qn("w:r")
_W_P = qn("w:p")
_W_TC = qn("w:tc")
_W_BR = qn("w:br")
_W_CR = qn("w:cr")
_W_DELTEXT = qn("w:delText")
_W_INSTR = qn("w:instrText")
_W_FLDSIMPLE = qn("w:fldSimple")
_W_FLDCHAR = qn("w:fldChar")
_W_VAL = qn("w:val")

DEFAULT_TEXT_WIDTH_TWIPS = 9360  # 6.5" between margins


# ---- generic content-control helpers ----------------------------------------

def sdt_tag(sdt) -> str | None:
    pr = sdt.find(qn("w:sdtPr"))
    if pr is None:
        return None
    tag = pr.find(qn("w:tag"))
    return tag.get(_W_VAL) if tag is not None else None


def sdt_text(sdt) -> str:
    content = sdt.find(qn("w:sdtContent"))
    if content is None:
        return ""
    return "".join(t.text or "" for t in content.iter(_W_T))


def set_sdt_text(sdt, text: str) -> None:
    """Rewrite a content control's text, keeping the first run's formatting."""
    content = sdt.find(qn("w:sdtContent"))
    if content is None:
        content = OxmlElement("w:sdtContent")
        sdt.append(content)
    ts = list(content.iter(_W_T))
    if not ts:
        content.append(_run(text))
        return
    ts[0].text = text
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""


def make_sdt(tag: str, text: str, locked: bool = True):
    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    t = OxmlElement("w:tag")
    t.set(_W_VAL, tag)
    pr.append(t)
    if locked:
        lock = OxmlElement("w:lock")
        lock.set(_W_VAL, "sdtContentLocked")
        pr.append(lock)
    sdt.append(pr)
    content = OxmlElement("w:sdtContent")
    content.append(_run(text))
    sdt.append(content)
    return sdt


def _run(text: str = "", tab: bool = False):
    r = OxmlElement("w:r")
    if tab:
        r.append(OxmlElement("w:tab"))
    if text:
        t = OxmlElement("w:t")
        t.text = text
        t.set(qn("xml:space"), "preserve")
        r.append(t)
    return r


def _field(instr: str, placeholder: str = "1"):
    f = OxmlElement("w:fldSimple")
    f.set(qn("w:instr"), f" {instr} ")
    f.append(_run(placeholder))
    return f


def _is_bb(tag: str | None) -> bool:
    return bool(tag) and tag.startswith("bb:")


# ---- body hash (SPEC §7.4) --------------------------------------------------

def _body_text(body) -> str:
    parts: list[str] = []

    def rec(e) -> None:
        tag = e.tag
        if tag == _W_SDT and _is_bb(sdt_tag(e)):
            parts.append("\x1e")  # position of an app-owned control, not its text
            return
        if tag in (_W_T, _W_DELTEXT, _W_INSTR):
            parts.append(e.text or "")
        elif tag == _W_TAB and e.getparent() is not None and e.getparent().tag == _W_R:
            parts.append("\t")
        elif tag in (_W_BR, _W_CR):
            parts.append("\n")
        for c in e:
            rec(c)
        if tag == _W_P:
            parts.append("\n")
        elif tag == _W_TC:
            parts.append("\x1f")

    rec(body)
    return "".join(parts)


def body_hash(path: Path | str) -> str:
    doc = Document(str(path))
    return hashlib.sha256(_body_text(doc.element.body).encode("utf-8")).hexdigest()


# ---- header/footer model ----------------------------------------------------

def _part_signature(el) -> str:
    """Canonical text of a header/footer part, robust to Word re-saving it.

    App controls render as <tag>text</>; fields render as {INSTR} whether
    simple or complex, with their cached results ignored.
    """
    parts: list[str] = []
    state = {"field_depth": 0, "in_result": [], "instr": []}

    def rec(e) -> None:
        tag = e.tag
        if tag == _W_SDT and _is_bb(sdt_tag(e)):
            parts.append(f"<{sdt_tag(e)}>{sdt_text(e)}</>")
            return
        if tag == _W_FLDSIMPLE:
            parts.append("{" + (e.get(qn("w:instr")) or "").strip().split(" ")[0].upper() + "}")
            return
        if tag == _W_FLDCHAR:
            kind = e.get(qn("w:fldCharType"))
            if kind == "begin":
                state["field_depth"] += 1
                state["instr"].append("")
                state["in_result"].append(False)
            elif kind == "separate" and state["in_result"]:
                state["in_result"][-1] = True
            elif kind == "end" and state["field_depth"]:
                state["field_depth"] -= 1
                instr = state["instr"].pop().strip()
                state["in_result"].pop()
                parts.append("{" + (instr.split(" ")[0].upper() if instr else "") + "}")
            return
        in_field = state["field_depth"] > 0
        if tag == _W_INSTR:
            if in_field:
                state["instr"][-1] += e.text or ""
            return
        if in_field:
            # skip field results and anything else inside the field
            for c in e:
                rec(c)
            return
        if tag == _W_T:
            parts.append(e.text or "")
        elif tag == _W_TAB and e.getparent() is not None and e.getparent().tag == _W_R:
            parts.append("\t")
        for c in e:
            rec(c)
        if tag == _W_P:
            parts.append("\n")

    rec(el)
    return "".join(parts)


def _style_exists(doc, style_id: str) -> bool:
    return any(s.get(qn("w:styleId")) == style_id for s in doc.styles.element.findall(qn("w:style")))


def _para(style: str | None, tabs: list[tuple[str, int]] = (), jc: str | None = None):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    if style:
        ps = OxmlElement("w:pStyle")
        ps.set(_W_VAL, style)
        ppr.append(ps)
    if tabs:
        tabs_el = OxmlElement("w:tabs")
        for kind, pos in tabs:
            t = OxmlElement("w:tab")
            t.set(_W_VAL, kind)
            t.set(qn("w:pos"), str(pos))
            tabs_el.append(t)
        ppr.append(tabs_el)
    if jc:
        j = OxmlElement("w:jc")
        j.set(_W_VAL, jc)
        ppr.append(j)
    p.append(ppr)
    return p


def _text_width(doc) -> int:
    sec = doc.sections[0]
    try:
        w = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)
        return w if w > 0 else DEFAULT_TEXT_WIDTH_TWIPS
    except TypeError:
        return DEFAULT_TEXT_WIDTH_TWIPS


def build_header(spec: HeaderSpec, style: str | None, width: int) -> list:
    paras = []
    lines = list(spec.lines) or [""]
    for i, line in enumerate(lines, start=1):
        if i == 1 and spec.code is not None:
            p = _para(style, [("right", width)])
            p.append(make_sdt("bb:header:1", line))
            p.append(_run(tab=True))
            p.append(_run(f"{spec.ref_label} "))
            p.append(make_sdt(TAG_CODE, spec.code))
        else:
            p = _para(style)
            p.append(make_sdt(f"bb:header:{i}", line))
        paras.append(p)
    return paras


def build_footer(spec: HeaderSpec, style: str | None, width: int) -> list:
    paras = []
    left, center, right = spec.footer
    if left or center or right:
        p = _para(style, [("center", width // 2), ("right", width)])
        p.append(make_sdt("bb:footer:left", left))
        p.append(_run(tab=True))
        p.append(make_sdt("bb:footer:center", center))
        p.append(_run(tab=True))
        p.append(make_sdt("bb:footer:right", right))
        paras.append(p)
    if spec.page_numbers:
        p = _para(style, jc="right")
        p.append(_run("Page "))
        p.append(_field("PAGE"))
        p.append(_run(" of "))
        p.append(_field("NUMPAGES"))
        paras.append(p)
    if not paras:
        paras.append(_para(style))
    return paras


def _targets(doc):
    """(header, footer) pairs the app owns: section 1 primary, plus first/even if enabled."""
    sec = doc.sections[0]
    pairs = [(sec.header, sec.footer)]
    if sec.different_first_page_header_footer:
        pairs.append((sec.first_page_header, sec.first_page_footer))
    if doc.settings.odd_and_even_pages_header_footer:
        pairs.append((sec.even_page_header, sec.even_page_footer))
    return pairs


def _current_signature(hf) -> str | None:
    if hf.is_linked_to_previous:
        return None
    return _part_signature(hf._element)


def _desired_signatures(doc, spec: HeaderSpec) -> tuple[str, str]:
    width = _text_width(doc)
    h = OxmlElement("w:hdr")
    for p in build_header(spec, None, width):
        h.append(p)
    f = OxmlElement("w:ftr")
    for p in build_footer(spec, None, width):
        f.append(p)
    return _part_signature(h), _part_signature(f)


def header_matches(doc, spec: HeaderSpec) -> bool:
    want_h, want_f = _desired_signatures(doc, spec)
    for hdr, ftr in _targets(doc):
        if _current_signature(hdr) != want_h or _current_signature(ftr) != want_f:
            return False
    return True


def _replace_part(hf, paras) -> None:
    hf.is_linked_to_previous = False
    el = hf._element
    for child in list(el):
        el.remove(child)
    for p in paras:
        el.append(p)


def write_header(doc, spec: HeaderSpec) -> None:
    width = _text_width(doc)
    h_style = "Header" if _style_exists(doc, "Header") else None
    f_style = "Footer" if _style_exists(doc, "Footer") else None
    for hdr, ftr in _targets(doc):
        _replace_part(hdr, build_header(spec, h_style, width))
        _replace_part(ftr, build_footer(spec, f_style, width))


def read_code(path: Path | str) -> str | None:
    doc = Document(str(path))
    hdr = doc.sections[0].header
    if hdr.is_linked_to_previous:
        return None
    for sdt in hdr._element.iter(_W_SDT):
        if sdt_tag(sdt) == TAG_CODE:
            return sdt_text(sdt)
    return None


def header_text(path: Path | str) -> str:
    """Plain text of section 1's primary header and footer (for tests and import logs)."""
    doc = Document(str(path))
    sec = doc.sections[0]
    out = []
    for hf in (sec.header, sec.footer):
        if not hf.is_linked_to_previous:
            out.append(_part_signature(hf._element))
    return "\n---\n".join(out)


# ---- references (SPEC §6.1) -------------------------------------------------

@dataclass
class RefHit:
    target_id: str
    current: str
    desired: str
    context: str


def _paragraph_text(el) -> str:
    p = el
    while p is not None and p.tag != _W_P:
        p = p.getparent()
    if p is None:
        return ""
    return "".join(t.text or "" for t in p.iter(_W_T)).strip()


def ref_controls(doc) -> list:
    return [s for s in doc.element.body.iter(_W_SDT) if (sdt_tag(s) or "").startswith(TAG_REF_PREFIX)]


def scan_refs(doc, codes: dict[str, str]) -> list[RefHit]:
    hits = []
    for sdt in ref_controls(doc):
        target = sdt_tag(sdt)[len(TAG_REF_PREFIX):]
        desired = codes.get(target, BROKEN)
        hits.append(RefHit(target, sdt_text(sdt), desired, _paragraph_text(sdt)))
    return hits


def update_refs(doc, codes: dict[str, str]) -> int:
    n = 0
    for sdt in ref_controls(doc):
        target = sdt_tag(sdt)[len(TAG_REF_PREFIX):]
        desired = codes.get(target, BROKEN)
        if sdt_text(sdt) != desired:
            set_sdt_text(sdt, desired)
            n += 1
    return n


# ---- file-level operations used by sync -------------------------------------

@dataclass
class DocState:
    header_ok: bool
    refs: list[RefHit] = field(default_factory=list)

    @property
    def stale_refs(self) -> list[RefHit]:
        return [r for r in self.refs if r.current != r.desired]

    @property
    def broken_refs(self) -> list[RefHit]:
        return [r for r in self.refs if r.desired == BROKEN]


def inspect(path: Path | str, spec: HeaderSpec | None, codes: dict[str, str]) -> DocState:
    doc = Document(str(path))
    ok = True if spec is None else header_matches(doc, spec)
    return DocState(ok, scan_refs(doc, codes))


def update(src: Path | str, dest: Path | str, spec: HeaderSpec | None, codes: dict[str, str]) -> None:
    """Read `src`, stamp header (unless spec is None) and refs, save to `dest`."""
    doc = Document(str(src))
    if spec is not None and not header_matches(doc, spec):
        write_header(doc, spec)
    update_refs(doc, codes)
    doc.save(str(dest))


def create_blank(path: Path | str) -> None:
    Document().save(str(path))
