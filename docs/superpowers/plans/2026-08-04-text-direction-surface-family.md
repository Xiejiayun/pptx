# Text Direction Surface Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close all 38 remaining PptxGenJS 4.0.1 text-direction atoms with one reusable runtime, native, package, browser, OOXML, and matrix evidence batch.

**Architecture:** Three read-only evidence agents classify declarations, lock PptxGenJS runtime behavior, and inventory native/client evidence. The main agent exclusively edits repository files, reuses existing deep lifecycle coverage, adds only missing aggregate probes, regenerates the matrix, reviews the batch, and performs one family commit and push.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, persistent Chromium, `pptx-inspect`, OOXML package inspection, npm pack.

## Global Constraints

- Implement the approved design in `docs/superpowers/specs/2026-08-04-text-direction-surface-family-design.md`.
- Close exactly 38 currently unverified atoms: 18 supported and 20 defect-excluded.
- Do not duplicate or change the five completed `IChartOpts.textDirection` atoms.
- Keep native table and text-box direction token sets strict and separate.
- Reuse existing deep lifecycle tests; new probes cover only missing aggregate scope.
- Run npm pack, persistent Chrome, `pptx-inspect`, and full tests once for the batch.
- Evidence agents do not edit repository files.
- Do not stage `.pnpm-store/` or temporary PPTX/package/browser artifacts.
- Finish with review, one family commit, push, fetch, and zero local/remote divergence.

---

### Task 1: Lock all remaining PptxGenJS owners

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS table-cell `textDirection`, text-body `vert`, slide-number, placeholder, and `tableToSlides` writers.
- Produces: the stable test title `locks text direction behavior across every declared owner`.

- [x] **Step 1: Add one aggregate owner fixture**

Create all four table-direction tokens, all seven text-box `vert` tokens, and all four ignored tokens for each inert owner. Include positive controls so missing owner output cannot look like ignored behavior.

- [x] **Step 2: Assert exact active output**

Require table `tcPr@vert` token semantics, table-default inheritance, cell overrides, all seven `bodyPr@vert` tokens, canonical horizontal behavior, and outer `vert` precedence.

- [x] **Step 3: Assert exact inert output**

Require no owner-attributable direction change for placeholder, slide-number, `tableToSlides`, or `TextPropsOptions.textDirection` variants.

- [x] **Step 4: Run the focused adapter test**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks text direction behavior across every declared owner"
```

Expected: one matching aggregate test passes.

---

### Task 2: Add compact six-format native evidence

**Files:**

- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: `TableCellTextDirection`, `TextBoxTextDirection`, `PRESENTATION_FORMAT_PROFILES`, table/text models, and exact OOXML.
- Produces: the stable test title `creates and reopens text direction owners in all six formats`.

- [x] **Step 1: Create all active tokens in every format**

For each presentation format, create one text shape per seven-value text-box catalog and a table containing the four table-direction states, including a table default and explicit cell override.

- [x] **Step 2: Assert model and exact OOXML state**

Require exact detached model snapshots, `a:bodyPr@vert`, `a:tcPr@vert`, correct horizontal semantics, and no cross-use of text-box-only tokens on table cells.

- [x] **Step 3: Write and reopen**

Write, reopen, re-read both owner types, and require zero error diagnostics and a valid package.

- [x] **Step 4: Run the focused SDK test**

```bash
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts \
  -t "creates and reopens text direction owners in all six formats"
```

Expected: one matching aggregate test passes.

---

### Task 3: Add packed-package and persistent-browser family states

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Consumes: existing `tableTextDirectionState`, text shape creation/edit/reopen APIs, exact XML helpers, installed tarball exports, and the browser bundle.
- Produces: stable `textDirectionFamily` and `textDirectionFamilyState` report markers.

- [x] **Step 1: Add the packed-package aggregate**

Reuse the table state and add one seven-token text-box fixture covering create, exact OOXML, edit, clear, duplicate, reopen, invalid rollback, and diagnostics.

- [x] **Step 2: Mirror the aggregate in Chrome**

Run the same seven-token text-box state beside the existing table state and require the family marker in the persistent-browser expected report.

- [x] **Step 3: Validate syntax and the packed tarball**

```bash
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
cd packages/pptx
npm pack --ignore-scripts --pack-destination /tmp/pptx-text-direction-pack
cd ../..
node scripts/smoke-npm-package.mjs /tmp/pptx-text-direction-pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: `textDirectionFamily=true` and every family-state field is true.

---

### Task 4: Close and regenerate exactly 38 matrix atoms

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: stable focused test titles, package/browser markers, native anchors, and runtime classification.
- Produces: exactly 38 closed entries with totals 18 supported and 20 defect-excluded.

- [x] **Step 1: Generate entries from frozen lists**

Define active/inert owners and both token catalogs once. Generate IDs deterministically and attach status-specific evidence without hand-writing 38 records.

- [x] **Step 2: Lock exact family and global counts**

Require 1,125 closed entries. Require global totals: supported 552, deliberate-difference 178, deprecated-alias 83, defect-excluded 312, unverified 649.

- [x] **Step 3: Regenerate and verify the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node scripts/pptxgenjs-surface-audit.mjs --check
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: check reports only the 649 remaining unverified atoms; no unsupported or stale entries.

---

### Task 5: Review, commit, push, and synchronize

**Files:**

- Review every file changed by Tasks 1-4 plus this plan and its design.

**Interfaces:**

- Consumes: all family evidence and regenerated audit totals.
- Produces: one reviewed commit synchronized with `origin/main`.

- [x] **Step 1: Run all family gates**

Run focused tests, TypeScript, syntax, surface audit, npm tarball, persistent Chrome, `pptx-inspect`, and the full Vitest suite once at the batch boundary.

- [x] **Step 2: Review the generated diff**

Programmatically compare HEAD and generated JSON. Require exactly 38 status changes, all from `unverified`, with exact 18/20 classification and no unrelated staged path.

- [x] **Step 3: Commit and push**

```bash
git add <only-text-direction-family-files>
git commit -m "docs: close text direction surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`; `.pnpm-store/` remains untracked and unstaged.
