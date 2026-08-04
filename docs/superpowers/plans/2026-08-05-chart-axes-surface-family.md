# Chart Axes Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> all shared-file edits, review, commit, and push. Steps use checkbox
> (`- [ ]`) syntax for tracking.

**Goal:** Close the complete 155-atom PptxGenJS chart-axis family with strict
native category/date, value, and editable series axes plus runtime, package,
browser, and OOXML evidence.

**Architecture:** Split the public axis model and normalizers by legal OOXML
axis kind, extend the renderer/state reader symmetrically, and bind the entire
declaration slice to one aggregate evidence pipeline. Existing chart editing
and package transactions remain the lifecycle boundary.

**Tech Stack:** Node.js 22 ESM, TypeScript 5.9, Vitest 3.2, PptxGenJS 4.0.1,
Playwright/Chrome, OOXML, stable audit JSON and Markdown.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-05-chart-axes-surface-family-design.md`
  exactly.
- Close all 155 chart-axis atoms in one batch; do not leave the advanced-axis
  subset for a later family.
- Keep category/date, value, and series axis properties type-safe and reject
  illegal cross-kind field combinations before package mutation.
- Evidence agents never modify shared files.
- Run focused tests once for the family. Run npm, Chrome, OOXML, PowerPoint
  2010, and full gates once at the batch boundary.
- Never stage `.pnpm-store/` or temporary PPTX/package/browser artifacts.
- Finish with review, one family commit, push, fetch, and divergence `0 0`.

---

### Task 1: Split and normalize the strict axis model

**Files:**

- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/chart-options.internal.ts`
- Modify: `packages/model/src/chart-definition.internal.test.ts`

**Interfaces:**

- Produces: `ChartCategoryAxisOptions`, `ChartValueAxisOptions`,
  `ChartSeriesAxisOptions`, `ChartTimeUnit`, and `ChartDisplayUnit`.
- Produces: primary/secondary category/value fields and the 3-D series field on
  `ChartOptions`.

- [ ] **Step 1: Add failing type and normalization tests**

Require date inference, all time/display tokens, numeric zero crossings,
positive frequencies, display-label dependency, per-axis frozen objects, and
rejection of invalid kinds, tokens, ranges, and unknown cross-kind fields.

- [ ] **Step 2: Implement the three public interfaces**

Move only legal fields into each interface and keep common axis properties in
an internal/public base interface without restoring permissive union fields.

- [ ] **Step 3: Implement axis-specific normalizers**

Use exact allowlists and normalize all nested values before freezing. Infer
`kind: 'date'` from time units and validate all cross-field invariants.

- [ ] **Step 4: Run focused definition tests**

Run:
`pnpm vitest run packages/model/src/chart-definition.internal.test.ts`

Expected: all tests pass.

### Task 2: Render and read every axis kind symmetrically

**Files:**

- Modify: `packages/model/src/chart-render.internal.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/chart-render.internal.test.ts`
- Modify: `packages/model/src/chart-state.internal.test.ts`

**Interfaces:**

- Consumes: the three strict option interfaces from Task 1.
- Produces: schema-ordered `catAx/dateAx/valAx/serAx` XML and canonical
  detached readback.

- [ ] **Step 1: Add one failing renderer aggregate**

Render category, date, value, and series axes containing numeric-zero and
`autoZero` crossings, three time-unit positions, frequency, multi-level
labels, every display-unit token, optional display label, and series title,
font, line, grid, units, orientation, visibility, and label position.

- [ ] **Step 2: Implement strict axis rendering**

Add shared crossing/common fragments plus kind-specific fragments. Pass the
single editable series option object into `renderSeriesAxis` for a 3-D bar axis
set.

- [ ] **Step 3: Add one failing state-reader aggregate**

Read the renderer output and a representative PptxGenJS date/value/series
fixture. Require date kind, non-default values, numeric zero, series role
resolution, frozen nested lines/titles/colors, and canonical omission of
defaults.

- [ ] **Step 4: Implement strict state reading**

Resolve the third `serAx` member, read only fields legal for the element kind,
and populate the 3-D chart's series options without changing two-axis or
axis-free charts.

- [ ] **Step 5: Run focused render/state tests**

Run:
`pnpm vitest run packages/model/src/chart-render.internal.test.ts packages/model/src/chart-state.internal.test.ts`

Expected: all tests pass.

### Task 3: Prove public lifecycle and runtime parity

**Files:**

- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

- [ ] **Step 1: Add one six-format SDK aggregate**

Create date/value/series axis charts in every format, duplicate, edit the
source, clear optional fields, reject one malformed replacement with byte and
journal rollback, write, reopen, and require the duplicate to retain its
snapshot with zero package errors.

- [ ] **Step 2: Extend the single PptxGenJS control**

Cover flat and nested category/value axes, all nine display units, date units,
frequency, multi-level labels, secondary axes, valid series fields, and the
eleven locked defects in one runtime/control family.

- [ ] **Step 3: Run focused public tests once**

Run the chart-axis model tests, the single adapter family test, the single SDK
aggregate, and package typecheck once. Expected: all pass.

### Task 4: Extend actual-package and persistent-browser aggregates

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

- [ ] **Step 1: Extend the installed-package aggregate**

Import only the packed public package, typecheck all three axis interfaces,
then create/edit/write/reopen representative date/value/series axes and assert
their exact public values and XML.

- [ ] **Step 2: Extend one persistent Chrome state**

Reuse the existing browser document/session. Cover immediate state, exact
OOXML, edit/clear, malformed rollback, duplicate isolation, write/reopen, MIME,
and diagnostics for all advanced axis semantics.

- [ ] **Step 3: Build/install and run each aggregate once**

Run the actual npm pack smoke once and the real-Chrome smoke once. Persist
compact JSON and artifact hashes under `/tmp`.

### Task 5: Bind the exact matrix family

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

- [ ] **Step 1: Generate the exact 155 entries**

Generate 79 supported, 65 deliberate-difference, and 11 defect-excluded
entries from owner/property/token arrays. Bind implementation, focused tests,
PptxGenJS control, package, OOXML, and browser evidence.

- [ ] **Step 2: Lock family and global totals**

Assert the exact family IDs/statuses and global totals: 716 supported, 431
deliberate, 94 deprecated, 370 defect, 163 unverified, and 1,611 classified.

- [ ] **Step 3: Regenerate deterministically**

Run `node scripts/pptxgenjs-surface-audit.mjs --write` twice and require the
second generation to leave tracked artifacts unchanged.

### Task 6: Review, validate, commit, and push

- [ ] **Step 1: Run the batch gates once**

Run focused tests, package typecheck, actual tarball, persistent Chrome,
OOXML/package validation, PowerPoint 2010 validation, and the full suite.

- [ ] **Step 2: Review the complete diff**

Check type legality, schema ordering, reader/renderer symmetry, default
canonicalization, rollback/isolation, exact 155-ID scope, evidence paths, and
absence of unrelated tracked files.

- [ ] **Step 3: Commit and push the family**

Stage only intended tracked files, commit with
`feat: complete chart axes surface family`, push `main`, fetch, and require
`main...origin/main` divergence `0 0`.
