# Underline Surface Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close all 120 PptxGenJS 4.0.1 underline atoms with one reusable runtime, native, package, browser, OOXML, and matrix evidence batch.

**Architecture:** Three read-only evidence agents lock the declaration matrix, upstream runtime/control behavior, and native/package/client coverage. The main agent exclusively edits repository files, adds one aggregate fixture per validation layer, regenerates the matrix, reviews the combined evidence, and performs one family commit and push.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, JSZip, persistent Chromium, `pptx-inspect`, OOXML package inspection, npm pack.

## Global Constraints

- Implement the approved design in `docs/superpowers/specs/2026-08-04-underline-surface-family-design.md` without another user decision.
- Evaluate exactly six owners and twenty atoms per owner.
- Active owners are `PlaceholderProps`, `TableCellProps`, `TableProps`, and `TextPropsOptions`.
- Inert owners are `SlideNumberProps` and `TableToSlidesProps`.
- Preserve native `dotDashHeavy` and reject upstream invalid `dotDashHeave`.
- Express PptxGenJS `none` semantically as native `underline: false`.
- Run focused tests once for the family and npm pack, persistent Chromium, `pptx-inspect`, and full tests once for the batch.
- Evidence agents do not edit repository files. Only the main agent updates tests, scripts, plans, specs, generated matrices, and Git state.
- Do not stage `.pnpm-store/` or temporary PPTX/package/browser harness artifacts.
- Finish with review, one family commit, push, fetch, and zero local/remote divergence.

---

### Task 1: Lock the six-owner PptxGenJS control matrix

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 underline declarations and generated slide/layout OOXML.
- Produces: the stable test title `locks underline behavior across every declared owner` and an exact owner/token classification table.

- [x] **Step 1: Add one aggregate runtime fixture**

Create all seventeen declared style tokens for each active owner in a bounded set of slides/layouts. Create equivalent inert-owner variants and positive controls in the same presentation or deterministic paired presentations.

- [x] **Step 2: Assert exact active output**

Extract owner-attributable XML and require the complete style sequence. Lock `dbl` plus color, color-only, empty object, `none`, invalid `dotDashHeave`, and undeclared-correct `dotDashHeavy` behavior.

- [x] **Step 3: Assert exact inert output**

Require positive-control content for `SlideNumberProps` and `TableToSlidesProps`, but no owner-attributable underline attribute or fill for any declared token.

- [x] **Step 4: Run the focused adapter test once**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks underline behavior across every declared owner"
```

Expected: one matching aggregate test passes.

---

### Task 2: Lock the native six-format underline lifecycle

**Files:**

- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: `PptxDocument`, `RichTextRunStyle.underline`, `RichTextUnderline`, strict colors, text/placeholder/table models, and six presentation formats.
- Produces: the stable test title `creates edits clears and reopens underline owners in all six formats` with exact DrawingML assertions.

- [x] **Step 1: Add one six-format aggregate fixture**

For every key of `PRESENTATION_FORMAT_PROFILES`, create text, placeholder, table-cell, and table-wide equivalent content. Exercise the fifteen common styles, native `false`, native-only `words`/`dotDashHeavy`, sRGB color, and scheme color.

- [x] **Step 2: Exercise the full mutation lifecycle**

Require create, detached snapshots, duplicate, edit, explicit disable, clear, failed invalid-token rollback, write, reopen, and continued editing without package residue.

- [x] **Step 3: Assert exact OOXML and diagnostics**

Require exact `a:rPr@u` values, `u="none"`, exact `a:uFill` color structure, no `dotDashHeave`, valid relationships, and zero error diagnostics.

- [x] **Step 4: Run the focused SDK test once**

```bash
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts \
  -t "creates edits clears and reopens underline owners in all six formats"
```

Expected: one matching aggregate test passes.

---

### Task 3: Add packed-package and persistent-browser aggregate states

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Consumes: installed tarball ESM/CJS/browser exports, packed declarations, `write()`/`writeBlob()`, strict runtime validation, and existing ZIP/XML helpers.
- Produces: stable `underlineFamily` and `underlineFamilyState` markers in package and browser reports.

- [x] **Step 1: Add the packed-package state**

Cover typed declarations, common/native style catalogs, immediate snapshots, exact OOXML/color, disable/clear, detached getter isolation, invalid-input rollback, reopen, and diagnostics in one installed-tarball run.

- [x] **Step 2: Mirror the state in the browser smoke**

Use the browser bundle and `writeBlob()` to produce the same semantic state. Require the state to be truthy and retain the global console/page/network error gates.

- [x] **Step 3: Run syntax and package validation once**

```bash
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
cd packages/pptx
npm pack --ignore-scripts --pack-destination /tmp/pptx-underline-pack
cd ../..
node scripts/smoke-npm-package.mjs /tmp/pptx-underline-pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: the package report exposes a truthy marker with every state field true. Execute the matching probe once in persistent Chromium and require empty browser diagnostic logs.

---

### Task 4: Close and regenerate exactly 120 matrix atoms

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: stable focused test titles, package/browser markers, native anchors, and runtime/control classification.
- Produces: exactly 120 closed entries with totals 68 supported, 8 deliberate-difference, and 44 defect-excluded.

- [x] **Step 1: Generate entries from frozen lists**

Define the six owners, two inline fields, and seventeen tokens once. Generate IDs deterministically and attach status-specific evidence without copying 120 hand-written records.

- [x] **Step 2: Lock exact family and global counts**

Update immutable-manifest assertions to require 1,087 closed entries and family totals 68/8/44. Assert every owner contributes twenty entries and no non-underline ID enters the family.

- [x] **Step 3: Regenerate the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node scripts/pptxgenjs-surface-audit.mjs --check
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected global totals: supported 534, deliberate-difference 178, deprecated-alias 83, defect-excluded 292, unverified 687, unsupported 0, stale 0.

---

### Task 5: Review, commit, push, and synchronize

**Files:**

- Review every file changed by Tasks 1-4 plus this plan and its design.

**Interfaces:**

- Consumes: all family evidence and regenerated audit totals.
- Produces: one reviewed commit synchronized with `origin/main`.

- [x] **Step 1: Run all family gates**

Run both focused tests, TypeScript, syntax, surface audit, npm tarball, persistent Chromium, `pptx-inspect`, and the full Vitest repository suite once at the appropriate batch boundary.

- [x] **Step 2: Review the evidence graph and generated diff**

Programmatically compare the prior and generated JSON. Require exactly 120 changed statuses, no non-underline ID, exact 68/8/44 category totals, and no unrelated staged path.

- [x] **Step 3: Commit and push**

```bash
git add <only-underline-family-files>
git commit -m "docs: close underline surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`; `.pnpm-store/` remains untracked and unstaged.
