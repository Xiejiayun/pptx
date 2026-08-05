# Image Identity and Visual Effects Surface Family Implementation Plan

> **For agentic workers:** Implement task-by-task with evidence agents writing only to `/tmp`; the main agent owns repository edits, review, matrix updates, commit, and push.

**Goal:** Support PptxGenJS-equivalent image identity, rounding, shadow, and transparency through zero-input creation and lossless editing.

**Architecture:** Extend the shared image creation appearance, add one focused lossless image-appearance codec for metadata/geometry/blip transparency, and reuse the existing strict shadow codec for pictures. Evidence is batched across the five atoms.

**Tech Stack:** TypeScript, Vitest, lossless OOXML, OPC package transactions, npm pack, Playwright/Chrome.

## Global Constraints

- PptxGenJS control version remains exactly 4.0.1.
- Evidence agents do not modify repository files.
- Editors preserve unrelated XML, relationships, and media bytes.
- The family receives one focused gate, one full gate, one commit, and one push.

---

### Task 1: Lock creation behavior

**Files:**
- Modify: `packages/model/src/image.ts`
- Modify: `packages/model/src/image-create.internal.ts`
- Modify: `packages/model/src/svg-image-create.internal.ts`
- Test: `packages/model/src/image-create.internal.test.ts`
- Test: `packages/model/src/svg-image-create.internal.test.ts`

- [x] Add failing normalization tests for rounding, transparency, and a detached frozen shadow.
- [x] Add failing raster/SVG rendering assertions for ellipse geometry, `alphaModFix` ordering, and `effectLst`.
- [x] Extend the public options and shared normalized appearance minimally.
- [x] Run both focused creation test files and verify they pass.

### Task 2: Add safe lossless editors

**Files:**
- Create: `packages/model/src/image-appearance.internal.ts`
- Create: `packages/model/src/image-appearance.internal.test.ts`
- Modify: `packages/model/src/shape-shadow.internal.ts`
- Test: `packages/model/src/shape-shadow.internal.test.ts`

- [x] Write failing tests for identity, rounding, transparency, no-op, namespace variants, and malformed ownership.
- [x] Implement direct-fragment readers and replacers with strict ambiguity checks.
- [x] Allow the existing shadow codec to operate on `p:pic/p:spPr` without changing shape semantics.
- [x] Run the image-appearance and shadow focused tests.

### Task 3: Expose the live model lifecycle

**Files:**
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`

- [x] Write one failing lifecycle test covering create, edit, no-op, duplicate isolation, rollback, relationship stability, and reopen.
- [x] Add the shared name setter plus image alt-text, rounding, transparency, and shadow accessors.
- [x] Route mutations through OPC transactions and write only when a codec reports a change.
- [x] Run the focused model lifecycle gate.

### Task 4: Carry appearance through the SDK and adapter control

**Files:**
- Modify: `packages/sdk/src/raster-image-source.ts`
- Test: `packages/sdk/src/raster-image-source.test.ts`
- Test: `packages/sdk/src/index.test.ts`
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`

- [x] Add the three appearance keys to strict SDK source-option forwarding.
- [x] Add one SDK create/edit/reopen assertion group.
- [x] Add one PptxGenJS fixture/control comparison for all five atoms.
- [x] Run the family-focused SDK and adapter tests once.

### Task 5: Batch package, browser, documentation, and matrix evidence

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/implementation-progress.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/api/README.md`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.json`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.md`

- [x] Reuse one lifecycle state in native npm and persistent Chrome probes.
- [x] Validate exact image relationship counts and direct OOXML fragments.
- [x] Replace stale unsupported documentation with the implemented boundary.
- [x] Register three supported and two deliberate-difference manifest entries.
- [x] Run focused tests, TypeScript, full Vitest, npm pack, Chrome, PowerPoint 2010 validation, and the audit gate once.
- [x] Review the staged diff, commit the family, push `main`, and confirm zero divergence.
