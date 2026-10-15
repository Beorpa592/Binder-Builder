"""Manifest loading, validation and the format_version upgrade path (SPEC §4, phase 5)."""

import pytest

from binder_builder.engine import manifest as mf
from binder_builder.engine.manifest import Manifest, ManifestError


def test_upgrade_path_runs_steps_backs_up_and_keeps_comments(binder, monkeypatch):
    path = binder / ".binder" / "binder.yaml"
    original = path.read_text()

    def v1_to_v2(raw):
        raw["binder"]["vars"]["upgraded_by"] = "step 1"

    monkeypatch.setattr(mf, "FORMAT_VERSION", 2)
    monkeypatch.setitem(mf.UPGRADES, 1, v1_to_v2)
    m = Manifest.load(binder)
    assert m.raw["format_version"] == 2
    assert m.binder.vars["upgraded_by"] == "step 1"
    text = path.read_text()
    assert "format_version: 2" in text
    assert "# clarifying header: no letter, no folder, no tab" in text  # comments survive
    assert (binder / ".binder" / "binder.v1.yaml.bak").read_text() == original
    # loading again changes nothing further
    Manifest.load(binder)
    assert path.read_text() == text


def test_missing_upgrade_step_is_reported(binder, monkeypatch):
    monkeypatch.setattr(mf, "FORMAT_VERSION", 3)
    monkeypatch.setitem(mf.UPGRADES, 1, lambda raw: None)
    with pytest.raises(ManifestError, match="no upgrade path from format_version 2"):
        Manifest.load(binder)


def test_newer_manifest_refused_with_clear_message(binder):
    path = binder / ".binder" / "binder.yaml"
    path.write_text(path.read_text().replace("format_version: 1", "format_version: 99"))
    with pytest.raises(ManifestError, match="newer Binder Builder"):
        Manifest.load(binder)


@pytest.mark.parametrize("bad,message", [
    ("    kind: section\n    label: Control Section", "duplicate"),
    ("format_version: 1", "format_version"),
])
def test_invalid_manifests_rejected(binder, bad, message):
    path = binder / ".binder" / "binder.yaml"
    text = path.read_text()
    if message == "duplicate":
        text = text.replace("{id: notes-to-reviewer,", "{id: open-points,")
    else:
        text = text.replace(bad, "format_version: one")
    path.write_text(text)
    with pytest.raises(ManifestError, match=message):
        Manifest.load(binder)


def test_unknown_field_rejected(binder):
    path = binder / ".binder" / "binder.yaml"
    path.write_text(path.read_text().replace("ext: xlsx, filename: Open_Points}", "ext: xlsx, filenme: Open_Points}"))
    with pytest.raises(ManifestError, match="filenme"):
        Manifest.load(binder)
