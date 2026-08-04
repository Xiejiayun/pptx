# Chart Presentation Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> all shared-file edits, review, commit, and push. Steps use checkbox
> (`- [ ]`) syntax for tracking.

**Goal:** Close the exact 91-atom PptxGenJS chart-presentation family while
repairing hierarchical pie/doughnut data-label readback and preserving unsafe
point/custom label XML.

**Architecture:** Reuse the existing strict presentation model for 88 atoms.
Add a safe series/point label promotion reader and make the editor synchronize
label XML only when the public group data-label object changes. Bind the exact
matrix slice to one runtime/package/browser/OOXML evidence batch.

**Tech Stack:** Node.js 22 ESM, TypeScript 5.9, Vitest 3.2, PptxGenJS 4.0.1,
Playwright/Chrome, lossless OOXML, stable audit JSON and Markdown.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-05-chart-presentation-surface-family-design.md`
  exactly.
- Close exactly 91 atoms: 26 supported and 65 deliberate differences.
- Do not include the eight point-style atoms or the residual 23-atom family.
- Evidence agents never modify shared files.
- Preserve custom/partial/unknown series and point label XML on unrelated
  edits.
- Run focused, npm, Chrome, OOXML, TypeScript, and full gates once per family.
- Never stage `.pnpm-store/` or `/tmp` artifacts.
- Finish with review, one family commit, push, and divergence `0 0`.

---

### Task 1: Read safe effective hierarchical data labels

**Files:**

- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/chart-state.internal.test.ts`

**Interfaces:**

- Produces: a shared internal predicate/result for safely promotable direct
  series `c:dLbls`.
- Produces: effective `ChartDataLabelOptions` with explicit false point flags
  overriding series defaults.

- [x] **Step 1: Add failing reader aggregates**

Cover a complete three-point pie series with `bestFit`, category-name, and
percent overrides; a doughnut false override; multiple equal series; partial
indices; duplicate/out-of-range indices; custom `c:tx`; non-empty `c:spPr`;
and unknown extension payload.

- [x] **Step 2: Implement dense safe point promotion**

Read the value-cache point count, require exact `0..count-1` coverage, merge
supported point options over series defaults, apply explicit false flags, and
return a result only when all points are equal.

- [x] **Step 3: Implement series-to-group promotion**

Use direct group labels as fallback. Promote only when every series has an
equal safe result. Export the same safe predicate for the editor.

- [x] **Step 4: Run the reader tests**

Run:
`./node_modules/.bin/vitest run packages/model/src/chart-state.internal.test.ts`

Expected: all tests pass.

### Task 2: Canonicalize only explicitly edited safe labels

**Files:**

- Modify: `packages/model/src/chart-edit.internal.ts`
- Modify: `packages/model/src/chart-edit.internal.test.ts`

**Interfaces:**

- Consumes: the safe-promotion predicate from Task 1.
- Produces: group-local `dataLabels` change detection and narrow `dLbls`
  synchronization.

- [x] **Step 1: Add failing edit-isolation tests**

Require an unrelated title edit to preserve imported series/point label XML
exactly. Require an explicit promoted-label edit and clear to remove only safe
series layers, write canonical group labels, retain extensions outside those
layers, and reopen with the requested definition.

- [x] **Step 2: Add per-group label ownership**

Exclude `dLbls` from group owned names when current and next label options are
equal. When changed, synchronize the direct group layer and remove only series
layers accepted by the shared safe predicate.

- [x] **Step 3: Run editor and state tests**

Run:
`./node_modules/.bin/vitest run packages/model/src/chart-edit.internal.test.ts packages/model/src/chart-state.internal.test.ts`

Expected: all tests pass.

### Task 3: Lock actual PptxGenJS runtime and public lifecycle

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

- [x] **Step 1: Extend one PptxGenJS presentation control**

Generate actual pie/doughnut cases for all `showLabel`/`showPercent`
combinations and `bestFit`, plus a custom scatter label control. Require exact
effective readback, explicit edit/clear, and unrelated custom-label
preservation.

- [x] **Step 2: Extend one six-format SDK aggregate**

Create/edit/clear/duplicate/reopen representative data labels, legend,
data-table, title, marker, grouping, pie/doughnut, radar, blank, and 3-D
presentation state across all six formats with rollback and package validation.

- [x] **Step 3: Run focused public tests once**

Run only the presentation adapter control and SDK aggregate together with the
model/editor tests. Expected: all pass.

### Task 4: Extend actual-package and persistent-browser aggregates

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

- [x] **Step 1: Extend the installed-package aggregate**

Use only the packed public package. Create, edit, clear, write, and reopen the
representative presentation catalog and the hierarchical label cases. Assert
workbook synchronization, rollback, duplicate isolation, and exact OOXML.

- [x] **Step 2: Extend the persistent Chrome aggregate**

Reuse one browser state and require the same public values plus MIME,
diagnostics, console/page/network error counts, and explicit false readback.

- [x] **Step 3: Run each expensive gate once**

Build/install one tarball and run one real-Chrome session. Persist evidence and
hashes under `/tmp` only.

### Task 5: Bind the exact 91-atom matrix

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

- [x] **Step 1: Generate exact entries from the reviewed map**

Create 26 supported and 65 deliberate-difference entries using the exact IDs,
native anchors, runtime control, package, OOXML, and browser evidence.

- [x] **Step 2: Lock family and global totals**

Assert 91 unique family IDs and totals of 742 supported, 496 deliberate, 94
deprecated, 370 defect, 72 unverified, and 1,702 classified.

- [x] **Step 3: Generate twice deterministically**

Run `node scripts/pptxgenjs-surface-audit.mjs --write` twice and require
identical JSON/Markdown hashes, zero diagnostics, and zero stale entries.

### Task 6: Review, validate, commit, and push

- [x] **Step 1: Run the family gates**

Run focused tests, TypeScript, actual tarball, persistent Chrome,
OOXML/PowerPoint 2010, and the full suite.

- [x] **Step 2: Review the complete diff**

Check dense-index safety, explicit-false merging, safe promotion boundaries,
unrelated-edit preservation, exact 91-ID scope, evidence paths, deterministic
artifacts, and absence of unrelated tracked files.

- [x] **Step 3: Commit and push**

Stage only intended files, commit with
`feat: complete chart presentation surface family`, push `main`, and require
`main...origin/main` divergence `0 0`.
