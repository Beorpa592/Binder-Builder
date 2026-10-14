# Phase log

## Phase 0: Core engine proof (tag `phase-0`)

**Built:** `engine/manifest.py` (round-trip YAML, pydantic validation, version gate),
`tree.py` (node edits that keep YAML comments), `codes.py` (codes, folders, filenames,
tab labels, header text), `word.py` (`bb:*` content controls, header/footer writer,
`bb:ref` updates, body hash), `excel.py` (header writer, hidden `_binder` sheet, `bb_`
names, body hash), `storage/` (interface plus local-folder backend with lock file),
`sync.py` (plan, apply, backups, invariant check, `state.json`, `sync.log`, restore),
and the CLI. There is also a synthetic fixture binder with its generator.

**Acceptance (all passing):**
- Load the fixture, sync, and the second plan is empty.
- Insert an item between A.02 and A.03, then sync. A.03 onward (including `A.03.PB`)
  is renamed, every header code shifts, every `bb:ref` and every `bb_` name in every
  workbook resolves to the new code, and every body hash is unchanged.
- Delete a node. Its file still exists and is reported as an orphan.
- Plus: locked headers stay untouched, swapping sections loses no files, backups
  hold the original bytes, a tampered write is rejected by the invariant check,
  restore works, someone else's lock blocks apply, the CLI works, and the engine
  contains no GUI or COM imports.
- Manually verified in real Word and Excel through COM. Stamped headers render,
  `Page X of Y` computes, and `=bb_top_10_list` evaluates to the code.

**Deviations:** DEVIATIONS.md items 1–27. The ones with real consequences:
- Workbooks containing charts, images, pivots or macros are never written (#25).
- Files renamed by hand in Explorer show up as orphans (#17).
