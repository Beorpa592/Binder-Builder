# Test fixtures

`synthetic_binder/` is a made-up binder with the same structure as the North Star
samples: the same folder layout, file naming (including the legacy names
`01_TOC_m.xlsx`, `01_Open_Points.xlsx` and `00_Cover_Sheet_CS.docx`), header layout
and TOC layout. **Every word in it is placeholder text.** The repo is public, so
real binder content must never be committed here.

The fixture represents a binder whose manifest has already been written (as if
imported) but which has never been synced:

- Headers are legacy plain text, not app-owned content controls, and one uses the
  inconsistent label `Ref #:`. The first sync takes them all over.
- Some references are already claimed: `bb:ref:*` content controls in Word
  bodies, and `=bb_<id>` formulas in Excel cells.
- The legacy TOC stores Control codes as floats (`0.01`), as the samples do.
- It covers the edge cases: a suffixed node (`A.03.PB`), a sub-item (`D.01.01`),
  a clarifying header (`Content Mediums`), a `locked_header` document, an
  `included: false` document, a PDF item, and a document with a different
  first-page header.

## Regenerating

```bash
.venv/Scripts/python tests/fixtures/make_fixture.py
```

The script deletes `synthetic_binder/` and rebuilds it. Tests never sync the fixture
in place. They copy it to a temp folder first (`tests/conftest.py`).

Golden files for generated documents (phase 1) live in `golden/` and are rebuilt by
`make_golden.py`.
