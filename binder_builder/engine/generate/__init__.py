"""Generated documents (SPEC §8): binder cover, TOC, section covers, tabs, and the hybrid checklist.

Generated documents are app-owned and re-rendered from templates. Templates are
looked up in the binder's template folder first (so staff can restyle them in
Office), then in the default set shipped in `binder_builder/templates/`.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

from ..codes import BINDER_COVER, SECTION_COVER, TABS, TOC, Entity, Layout
from ..manifest import BINDER_DIR, Manifest
from ..tree import Tree

GENERATED = (BINDER_COVER, TOC, SECTION_COVER, TABS)
DEFAULT_TEMPLATES = Path(__file__).resolve().parent.parent.parent / "templates"
DEFAULT_TEMPLATE_FOLDER = f"{BINDER_DIR}/templates"

SECTION_COVER_TEMPLATE = "Section_Cover.docx"
TOC_TEMPLATE = "01_TOC.xlsx"
CKL_TEMPLATE = "02_CKL.xlsx"


@dataclass
class GenContext:
    folder: Path
    manifest: Manifest
    tree: Tree
    layout: Layout
    pages: dict[str, int] | None = None  # code -> running page number, continuous mode (phase 4)

    @property
    def binder(self):
        return self.manifest.binder

    def template_folder(self) -> Path:
        return self.folder / (self.binder.template_folder or DEFAULT_TEMPLATE_FOLDER)

    def template(self, name: str) -> Path:
        local = self.template_folder() / name
        return local if local.is_file() else DEFAULT_TEMPLATES / name

    def template_for(self, entity: Entity) -> Path:
        return self.template(template_name(entity, self.binder))


def template_name(entity: Entity, binder) -> str:
    if entity.kind == BINDER_COVER:
        return binder.cover.template
    if entity.kind == TOC:
        return TOC_TEMPLATE
    if entity.kind == SECTION_COVER:
        return SECTION_COVER_TEMPLATE
    if entity.kind == TABS:
        return binder.tabs.template
    raise ValueError(entity.kind)


def template_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def render(entity: Entity, ctx: GenContext, out: Path) -> list[str]:
    """Render one generated document's content to `out` (headers are stamped by the caller)."""
    from . import cover, section_cover, tabs, toc

    tpl = ctx.template_for(entity)
    if entity.kind == BINDER_COVER:
        return cover.render(ctx, tpl, out)
    if entity.kind == TOC:
        return toc.render(ctx, tpl, out)
    if entity.kind == SECTION_COVER:
        return section_cover.render(ctx, entity, tpl, out)
    if entity.kind == TABS:
        return tabs.render(ctx, tpl, out)
    raise ValueError(entity.kind)


def content_signature(path: Path) -> str:
    """Hash of everything a generated document shows: body, header/footer, cell values."""
    from .. import excel, word

    text = word.full_signature(path) if path.suffix.lower() == ".docx" else excel.full_signature(path)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()
