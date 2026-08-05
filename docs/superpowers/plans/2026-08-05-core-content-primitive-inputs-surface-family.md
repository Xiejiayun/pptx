# Core Content and Primitive Inputs Surface Family Implementation Plan

> **For agentic workers:** Evidence agents write only to `/tmp`; the main
> agent owns repository edits, review, matrix generation, commit, and push.

**Goal:** Close 14 PptxGenJS core content, color, coordinate, and margin atoms
by validating and packaging the existing strict native capabilities.

**Architecture:** Add one aggregate upstream control and one shared lifecycle
probe. Reuse the existing text/table/color/coordinate/margin implementations
and run expensive package/browser validation once for the complete family.

**Tech Stack:** TypeScript, Vitest, lossless OOXML, PptxGenJS 4.0.1, npm pack,
persistent Chrome, and the packaged `pptx-inspect` CLI.

## Global Constraints

- Close exactly 14 atoms with the expected 5 supported / 9 deliberate split.
- Do not add a second text, table, color, coordinate, or margin codec.
- Evidence agents do not modify repository files.
- Run focused tests once and expensive npm/Chrome/full-suite gates once.
- Commit and push the complete capability family as one unit.

---

### Task 1: Lock the upstream 14-atom control

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

- [ ] Add the aggregate test title `closes PptxGenJS core content and
  primitive inputs through strict native state`.
- [ ] Assert the exact 14 distinct IDs and expected 5/9/0 classification.
- [ ] Exercise plain/rich text, plain/structured/rich table cells, sRGB/theme
  color, numeric/percentage coordinates, and scalar/tuple margins.
- [ ] Lock the PptxGenJS asymmetric tuple-margin order with direct OOXML and
  compare the native documented order.
- [ ] Run the aggregate adapter control once.

### Task 2: Add one shared lifecycle probe

**Files:**
- Create: `scripts/core-content-primitive-inputs-14-lifecycle-probe.mjs`

- [ ] Create one two-slide plain/rich text and mixed-table fixture.
- [ ] Cover exact no-op, invalid isolation, rollback, live edits, duplicate or
  source isolation, and write/reopen.
- [ ] Bind text, color, coordinate, margin, table content/options, and OOXML
  assertions to exact owners.
- [ ] Require stable relationships and zero error/warning diagnostics.

### Task 3: Reuse one installed-package and browser artifact batch

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

- [ ] Run the shared probe once through the actual npm tarball.
- [ ] Run the same probe once through the tarball browser export in persistent
  Chrome.
- [ ] Retain both evidence decks and validate them with the actual tarball CLI.
- [ ] Compare decompressed package parts and record the PowerPoint 2010 warning
  contract.

### Task 4: Backfill the matrix and documentation

**Files:**
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.json`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`

- [ ] Register exactly 14 entries with code/test/package/OOXML/client/control
  evidence and the final runtime-confirmed status map.
- [ ] Lock 1,734 classified, 40 unverified, 97.75%, and zero stale,
  unsupported, or diagnostic entries.
- [ ] Document strict native representations without claiming PptxGenJS
  recursive/flat input shapes as native API shapes.
- [ ] Regenerate the audit twice and require byte-identical artifacts.

### Task 5: Batch review, commit, and synchronize

- [ ] Run TypeScript, focused tests, full Vitest, npm pack, Chrome,
  PowerPoint 2010 validation, exact OOXML checks, and surface audit once.
- [ ] Complete Agent A/B/C and main-agent review with zero findings.
- [ ] Stage only this capability family, commit, push `main`, fetch, and
  require local/remote divergence `0 0`.
