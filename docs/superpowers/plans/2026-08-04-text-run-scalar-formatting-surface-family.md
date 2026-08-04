# Text Run Scalar Formatting Surface Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close exactly 54 inherited PptxGenJS scalar text-formatting atoms with direct runtime, six-format native, package, browser, OOXML, and generated-matrix evidence.

**Architecture:** Three read-only evidence agents classify the declaration matrix, lock PptxGenJS runtime behavior, and inventory native/client evidence. The main agent owns all repository edits, adds one aggregate control per validation layer, reuses existing deep lifecycle coverage, and performs one reviewed capability-family commit and push.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, persistent Chrome, `pptx-inspect`, OOXML package inspection, npm pack.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-04-text-run-scalar-formatting-surface-family-design.md` exactly.
- Close only the six-owner by nine-property target set: 54 currently unverified atoms.
- Require exact classification totals: 19 supported, 14 deliberate-difference, and 21 defect-excluded.
- Do not reopen `IChartOpts` or include adjacent paragraph, geometry, transparency, or effect fields.
- Reuse deep lifecycle coverage; new tests are compact owner/family aggregates.
- Run focused tests once per layer and npm pack, persistent Chrome, `pptx-inspect`, and full tests once at the batch boundary.
- Evidence agents remain read-only; only the main agent edits repository files.
- Do not stage `.pnpm-store/` or temporary package, browser, or PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero local/remote divergence.

---

### Task 1: Lock the six-owner PptxGenJS runtime matrix

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS text, placeholder, slide-number, table, table-cell, and `tableToSlides` writers.
- Produces: the stable test title `locks scalar text formatting behavior across every declared owner`.

- [ ] **Step 1: Add one aggregate owner fixture**

Create a rich text shape, placeholder, slide number, table defaults, table-cell rich text, and `tableToSlides` variants using distinct values for all nine properties. Include owner-attributable names and computed-CSS positive controls.

- [ ] **Step 2: Assert active output exactly**

Read isolated owner fragments and require exact `lang`, `sz`, `b`, `i`, solid fill, highlight, typeface, paragraph count, and `a:br` state. Prove table inheritance, cell override behavior, and slide-number `rPr` versus `defRPr` placement.

- [ ] **Step 3: Assert inert and deliberate behavior exactly**

Require no owner-attributable output for every `X` matrix cell, byte-identical `tableToSlides` variants, the CSS positive control, the cell `bold: false` fallback defect, and permissive color/font/highlight differences.

- [ ] **Step 4: Run the focused adapter test**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks scalar text formatting behavior across every declared owner"
```

Expected: one aggregate test passes.

---

### Task 2: Add compact six-format native evidence

**Files:**

- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: `RichTextRunStyle`, `RichTextRun`, `AddTableOptions`, `AddTableCellOptions`, `SlideNumberOptions`, placeholder APIs, and `PRESENTATION_FORMAT_PROFILES`.
- Produces: the stable title `creates and reopens scalar text formatting owners in all six formats`.

- [ ] **Step 1: Create all effective owner states**

For every format, create rich text and placeholder runs with every scalar style, a table with table and cell defaults plus explicit rich runs, and a styled slide number.

- [ ] **Step 2: Assert detached model and exact OOXML**

Require semantic sRGB/scheme colors, `fontFamily`, exact point sizes, false preservation, language, highlight, bold/italic, canonical paragraph breaks, `a:br`, and slide-number style placement.

- [ ] **Step 3: Write and reopen**

Write each format, reopen it, re-read every owner, validate the package, and require zero error/warning diagnostics.

- [ ] **Step 4: Run the focused SDK test**

```bash
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts \
  -t "creates and reopens scalar text formatting owners in all six formats"
```

Expected: one aggregate test passes.

---

### Task 3: Add packed-package and persistent-browser family gates

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Consumes: existing `slideNumbers`, `tableTextDefaults`, and `richTextBreakLine` states plus one compact rich-run style fixture.
- Produces: stable `textRunScalarFamily` and `textRunScalarFamilyState` report fields.

- [ ] **Step 1: Add the packed-package aggregate**

Reuse existing family prerequisites and add only missing run-style create/edit/reopen/OOXML/diagnostic checks. Fail with the complete family state if any field is false.

- [ ] **Step 2: Mirror the aggregate in the browser**

Use the same field order and values in the real-browser report and expected object. Preserve console, page, and network error gates.

- [ ] **Step 3: Validate syntax and the packed tarball**

```bash
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
mkdir -p /tmp/pptx-text-run-scalar-pack
cd packages/pptx && npm pack --ignore-scripts --pack-destination /tmp/pptx-text-run-scalar-pack
cd ../.. && node scripts/smoke-npm-package.mjs \
  /tmp/pptx-text-run-scalar-pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: `textRunScalarFamily=true` and every family-state field is true.

---

### Task 4: Generate exactly 54 matrix entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: stable aggregate titles, native anchors, client markers, and the frozen six-by-nine classification table.
- Produces: 54 closed entries and new global totals 571/192/83/333/595.

- [ ] **Step 1: Generate entries from frozen tables**

Define owner/property status tables once, generate deterministic property IDs, and attach status-specific native, test, package, OOXML, client, control, and note evidence.

- [ ] **Step 2: Lock exact family and global counts**

Require 1,179 manifest entries and exact family totals 19/14/21. Require global totals: supported 571, deliberate-difference 192, deprecated-alias 83, defect-excluded 333, unverified 595.

- [ ] **Step 3: Regenerate and verify artifacts**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: generated files are deterministic, unsupported/stale/diagnostics are zero, and exactly 595 atoms remain unverified.

---

### Task 5: Validate OOXML artifacts and the full repository

**Files:**

- Review every file changed by Tasks 1-4 plus this plan and design.

**Interfaces:**

- Consumes: created/edited family PPTX artifacts and all focused gates.
- Produces: complete batch evidence ready for commit.

- [ ] **Step 1: Validate created and edited PPTX files**

Use `pptx-inspect --json package validate`, `package diff`, and exact `part read` on explicit `/tmp` files. Require zero errors/warnings and only the intended slide XML change.

- [ ] **Step 2: Run persistent Chrome once**

Run the existing browser smoke in one real Chrome persistent session. Require `textRunScalarFamily=true`, every state true, and console/page/network error counts all zero.

- [ ] **Step 3: Run full gates once**

```bash
./node_modules/.bin/tsc -b --pretty false
./node_modules/.bin/vitest run
git diff --check
```

Expected: typecheck and all tests pass with no whitespace errors.

---

### Task 6: Review, commit, push, and synchronize

**Files:**

- Review and stage only the ten capability-family files from Tasks 1-5.

**Interfaces:**

- Consumes: verified family diff and exact matrix comparison against `HEAD`.
- Produces: one commit synchronized with `origin/main`.

- [ ] **Step 1: Review exact matrix and repository diff**

Programmatically require exactly 54 status changes, all from `unverified`, with exact 19/14/21 classification and no unrelated staged path.

- [ ] **Step 2: Commit and push**

```bash
git commit -m "docs: close scalar text formatting surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`; temporary and `.pnpm-store/` paths remain unstaged.
