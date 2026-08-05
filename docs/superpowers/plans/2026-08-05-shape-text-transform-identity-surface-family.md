# Shape and Text Transform Identity Surface Family Implementation Plan

> **For agentic workers:** Evidence agents write only to `/tmp`; the main
> agent owns repository edits, review, matrix generation, commit, and push.

**Goal:** Close 13 PptxGenJS shape/text transform and identity atoms by reusing
the existing native state and permanently validating the shared name editor.

**Architecture:** Keep the single direct `cNvPr@name` codec, rename its
normalizer to an owner-neutral contract, and add one cross-owner lifecycle
probe. Reuse existing transform/geometry tests and batch package/browser
evidence across all 13 atoms.

**Tech Stack:** TypeScript, Vitest, lossless OOXML, OPC transactions,
PptxGenJS 4.0.1, npm pack, persistent Chrome, `pptx-inspect`.

## Global Constraints

- Close exactly 13 atoms with the exact 2 supported / 10 deliberate / 1
  defect classification.
- Do not add a second name, transform, radius, preset, or text-box codec.
- Evidence agents do not modify repository files.
- Run focused tests once and expensive npm/Chrome/full-suite gates once.
- Commit and push the complete capability family as one unit.

---

### Task 1: Generalize and lock shared identity editing

**Files:**
- Modify: `packages/model/src/image-appearance.internal.ts`
- Modify: `packages/model/src/image-appearance.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

- [x] Rename the shared name normalizer to `normalizeShapeName` and use the
  owner-neutral `Shape name` validation context.
- [x] Add an ordinary preset-shape/text lifecycle test for no-op, invalid
  input, rollback, duplicate isolation, escaped/empty names, and reopen.
- [x] Add malformed and wrong-namespace ordinary shape owner rejection with
  exact package and mutation isolation.
- [x] Run the focused model and codec gate once.

### Task 2: Lock the upstream 13-atom control

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

- [x] Add the exact aggregate test title `closes PptxGenJS shape and text
  transform identity through strict native state`.
- [x] Assert the 13 distinct IDs and exact 2/10/1 classification.
- [x] Lock the declared-but-ignored `ShapeProps.shapeName` defect and reuse
  existing legal transform/geometry fixtures.
- [x] Run the aggregate PptxGenJS control once.

### Task 3: Add one shared installed/client probe

**Files:**
- Create: `scripts/shape-text-transform-identity-13-lifecycle-probe.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

- [x] Use one native lifecycle probe for ordinary shape and text name edits.
- [x] Require unchanged transform, geometry, text, relationships, and zero
  diagnostics after edit, duplicate, rollback, and reopen.
- [x] Reuse the probe in the actual tarball and persistent Chrome smoke.
- [x] Retain npm/browser evidence PPTX files for exact OOXML validation.

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

- [x] Register exactly 13 entries with code/test/package/OOXML/client/control
  evidence and the 2/10/1 status map.
- [x] Lock 1,720 classified, 54 unverified, 96.96%, and zero stale,
  unsupported, or diagnostic entries.
- [x] Replace stale shape/text transform and identity gap statements.
- [x] Regenerate the audit twice and require byte-identical artifacts.

### Task 5: Batch review, commit, and synchronize

- [x] Run TypeScript, focused tests, full Vitest, npm pack, Chrome,
  PowerPoint 2010 validation, exact OOXML checks, and surface audit once.
- [x] Complete Agent A/B/C and main-agent review with zero findings.
- [x] Stage only this capability family, commit, push `main`, fetch, and
  require local/remote divergence `0 0`.
