# Binder Builder

A Windows desktop app that keeps the Word and Excel documents in a physical binder
in sync with its Table of Contents. Every document has a permanent internal ID, and
its reference code (`A.07`) is recomputed from its position in the TOC on every sync.
Move a document, sync, and its filename, folder, header and every reference to it
follow.

- **Specification:** [SPEC.md](SPEC.md) (authoritative)
- **Where the build departs from it:** [DEVIATIONS.md](DEVIATIONS.md)
- **Progress per phase:** [PHASES.md](PHASES.md)

## Safety rules

The app never deletes a file, and never changes body content an author wrote. It
shows every change before applying it, and snapshots every file into
`.binder/backups/<timestamp>/` before touching it. Each Word and Excel write is
checked against a body hash and discarded if anything outside the app-owned
header, footer and reference fields changed.

## Development

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -e .[dev]
.venv/Scripts/python -m pytest
```

Command line (used by the tests and the GUI):

```bash
python -m binder_builder sync "<binder folder>" --plan
python -m binder_builder sync "<binder folder>" --apply
python -m binder_builder backups "<binder folder>"
python -m binder_builder restore "<binder folder>" <backup>
```

Test fixtures are synthetic placeholder binders. See
[tests/fixtures/README.md](tests/fixtures/README.md). The repo is public, so real
binder content must never be committed.
