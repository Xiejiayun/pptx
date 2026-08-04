# Tab Stops Surface Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the PptxGenJS 4.0.1 `tabStops` capability family in one reviewed batch using reusable runtime, native, package, browser, and OOXML evidence.

**Architecture:** Three read-only evidence workers independently resolve the surface matrix, PptxGenJS runtime/control behavior, and native/package/browser/OOXML coverage. The main worker owns every repository edit, adds one aggregate fixture per validation layer, updates the manifest and generated matrices, reviews the combined evidence, and performs the single family commit and push.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, JSZip, persistent Chromium, `pptx-inspect`, OOXML package inspection, npm pack.

## Global Constraints

- Treat the user's four-lane proposal and the approved paragraph-tab-stops design as final; do not request another design decision.
- Evaluate exactly six declared owners: `PlaceholderProps`, `SlideNumberProps`, `TableCellProps`, `TableProps`, `TableToSlidesProps`, and `TextPropsOptions`.
- Evaluate seven atoms per owner: `property:tabStops`, inline `alignment`, inline `position`, and tokens `ctr`, `dec`, `l`, and `r`.
- Do not include `IChartOpts.tabStops`; it was already closed in the chart family.
- Classify from runtime and OOXML evidence, not TypeScript declarations alone. Preserve `defect-excluded`, `deliberate-difference`, and `deprecated-alias` semantics.
- Run focused tests once for the family and npm pack, persistent Chromium, and the full repository suite once for the batch.
- Evidence workers do not edit repository files. Only the main worker updates tests, scripts, generated matrices, and Git state.
- Do not stage `.pnpm-store/` or temporary PPTX/package artifacts.
- Finish with review, one family commit, push, fetch, and zero local/remote divergence.

---

### Task 1: Lock PptxGenJS owner and token behavior

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 owner options and generated slide OOXML.
- Produces: the stable test title `locks tab stop behavior across every declared owner` and an exact owner/token classification table.

- [x] **Step 1: Add one aggregate runtime/control test**

Create one fixture for each of the six owners. For every owner, exercise `tabStops: [{ position: 1.25, alignment: 'l' }, { position: 2.5, alignment: 'ctr' }, { position: 3.75, alignment: 'r' }, { position: 4.5, alignment: 'dec' }]`, write the deck, and inspect the owning slide XML.

- [x] **Step 2: Assert effective and inert behavior precisely**

For effective owners, require the exact `a:tab` sequence and EMU positions `1143000`, `2286000`, `3429000`, and `4114800`. For ignored owners, require that the control text remains present while no owner-attributable `a:tabLst` is emitted. Record runtime failures as defects rather than treating them as unsupported native behavior.

- [x] **Step 3: Run the focused adapter test once**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks tab stop behavior across every declared owner"
```

Expected: one matching aggregate test passes.

---

### Task 2: Lock native six-format lifecycle and OOXML behavior

**Files:**

- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: `PptxDocument`, text, placeholder, table-cell, table, and slide-number creation/editing APIs.
- Produces: the stable test title `creates edits clears and reopens tab stops in all six formats` plus exact OOXML assertions.

- [x] **Step 1: Add the six-format lifecycle fixture**

For `16:9`, `16:10`, `4:3`, `A4`, `letter`, and `screen16x10`, create representative text, placeholder, table-cell, table-default, and slide-number content with all four semantic alignments (`left`, `center`, `right`, `decimal`) and `\t` text.

- [x] **Step 2: Exercise one complete mutation lifecycle**

Require `create -> write -> reopen -> edit -> clear -> write -> reopen`. Assert public snapshots, isolation of returned arrays, preservation of unrelated paragraph properties, explicit empty `tabLst`, and removal when cleared.

- [x] **Step 3: Assert exact native OOXML and diagnostics**

Require `l`, `ctr`, `r`, and `dec` in source order with exact EMU positions, valid package relationships, and zero error diagnostics. Validate one produced PPTX with `pptx-inspect --json package validate` and read the exact slide part for the tab-list evidence.

- [x] **Step 4: Run the focused SDK test once**

```bash
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts \
  -t "creates edits clears and reopens tab stops in all six formats"
```

Expected: one matching aggregate test passes.

---

### Task 3: Add packed-package and persistent-browser aggregate probes

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Consumes: installed tarball ESM/CJS/browser exports, packed declarations, `write()`/`writeBlob()`, and existing ZIP/XML helpers.
- Produces: identical `tabStopsState` markers in package and browser reports.

- [x] **Step 1: Add one stable package state**

Define `tabStopsState` with boolean fields covering typed declarations, create snapshot, exact OOXML, edit snapshot, explicit-empty clear, reopen, detached getter/no-op isolation, invalid-input rollback, and zero diagnostics. Include effective native owners without rebuilding per owner.

- [x] **Step 2: Exercise the installed tarball once**

Require equivalent ESM, CJS, and browser-condition results and inspect the packed declaration block for `ParagraphTabStop`, `ParagraphTabStopAlignment`, and `tabStops` ownership.

- [x] **Step 3: Mirror the lifecycle in persistent Chromium**

Add the same stable state to `scripts/playwright-browser-smoke.js`, using `writeBlob()` and reopen. Require the exact state, package diagnostics `0`, and global console/page/network error counts `0`.

- [x] **Step 4: Run one package and browser batch**

```bash
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
cd packages/pptx
npm pack --ignore-scripts --pack-destination /tmp/pptx-tab-stops-pack
cd ../..
node scripts/smoke-npm-package.mjs /tmp/pptx-tab-stops-pack/jiayunxie-pptx-0.1.0.tgz
```

Then execute the browser smoke through the persistent Chromium harness. Expected: both reports expose a truthy `tabStops` marker with every `tabStopsState` field true.

---

### Task 4: Close the 42 capability atoms and regenerate the matrix

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: generated compatibility/progress files selected by `scripts/pptxgenjs-surface-audit.mjs --write`

**Interfaces:**

- Consumes: stable focused test titles, package/browser markers, native implementation anchors, and the runtime/control classification.
- Produces: exactly 42 newly closed owner/property/token atoms and updated exact category totals.

- [x] **Step 1: Generate only the declared entries**

Use explicit frozen owner and path lists rather than hand-written duplicates. Attach shared code, test, package, browser, and design evidence. Assert that `IChartOpts.tabStops` is absent from this family.

- [x] **Step 2: Derive and lock exact totals**

```bash
node scripts/pptxgenjs-surface-audit.mjs --check
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Update count assertions from audit output. Expected closed count is the previous `925` plus the exact number of accepted family entries; never replace an observed count with the estimate.

- [x] **Step 3: Regenerate and inspect the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
git diff --check
git diff --stat
```

Confirm that every changed matrix row belongs to this family and all generated progress percentages agree with the audit JSON.

---

### Task 5: Batch verification, review, commit, and push

**Files:**

- Review every file changed by Tasks 1-4.

**Interfaces:**

- Consumes: all family evidence and generated audit totals.
- Produces: one reviewed commit synchronized with `origin/main`.

- [x] **Step 1: Run the family and repository gates**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks tab stop behavior across every declared owner"
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts \
  -t "creates edits clears and reopens tab stops in all six formats"
./node_modules/.bin/tsc -b --pretty false
pnpm test
```

Run npm tarball, `pptx-inspect`, and persistent Chromium once each after focused tests.

- [x] **Step 2: Review the final diff and evidence graph**

Check owner classifications against runtime OOXML, native alignment names against PptxGenJS tokens, package/browser markers against actual lifecycle assertions, generated counts against the manifest, and staged paths against the family scope.

- [x] **Step 3: Commit and push the family**

```bash
git add <only-tab-stops-family-files>
git commit -m "docs: close tab stops surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`; `.pnpm-store/` remains untracked and unstaged.
