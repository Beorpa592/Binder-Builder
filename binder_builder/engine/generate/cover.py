"""Binder cover (SPEC §8): `binder.cover.lines`, no code in the header."""

from __future__ import annotations

from pathlib import Path

from docxtpl import DocxTemplate


def render(ctx, template: Path, out: Path) -> list[str]:
    tpl = DocxTemplate(str(template))
    lines = list(ctx.binder.cover.lines) or [ctx.binder.name]
    tpl.render({"lines": lines, "binder_name": ctx.binder.name}, autoescape=True)
    tpl.save(str(out))
    return []
