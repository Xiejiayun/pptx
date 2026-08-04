# tableToSlides Surface Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close exactly 42 PptxGenJS `tableToSlides` declaration atoms with one runtime control and reused native, package, browser, OOXML, and matrix evidence.

**Architecture:** Three read-only agents classify declarations, inspect the real PptxGenJS writer, and inventory release artifacts. The main agent owns all repository edits, adds one aggregate upstream control, generates the family manifest from frozen tables, runs each expensive gate once, and performs one reviewed commit and push.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, persistent Chrome, npm pack, `pptx-inspect`, OOXML package inspection.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-04-table-to-slides-surface-family-design.md` exactly.
- Close only the 42 target atoms and require the exact 4/20/2/16 classification.
- Reuse existing deep native, package, browser, and OOXML evidence; add only the missing aggregate PptxGenJS control and matrix assertions.
- Run focused tests once per layer and npm pack, Chrome, `pptx-inspect`, TypeScript, and full Vitest once at the batch boundary.
- Evidence agents remain read-only; only the main agent edits repository files.
- Do not stage `.pnpm-store/` or temporary package, browser, or PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero local/remote divergence.

---

### Task 1: Lock all 42 PptxGenJS runtime atoms

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: the public PptxGenJS 4.0.1 `tableToSlides()` method, DOM fixture, and public-output importer.
- Produces: the stable title `locks every declared tableToSlides option and nested addition against PptxGenJS 4.0.1`.

- [x] **Step 1: Add active and nested-addition controls**

Generate small and overflowing DOM tables. Prove weights, repeated headers,
master selection, start positions, geometry, slide margins, addition order, all
eight nested fields, synchronous return, caller mutation, and warning behavior.

- [x] **Step 2: Add deprecated and defect controls**

Prove both deprecated aliases match their canonical targets. Compare baseline
output with each inherited inert field and each scalar/vector union branch.
Require `autoPage:false` still paginates, requested header count is ignored,
DOM widths replace `colW`, `rowH` is not forwarded, and flat styling/name fields
do not appear. Lock `margin` as pagination-only and `verbose` as console-only.

- [x] **Step 3: Run the focused adapter control**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks every declared tableToSlides option and nested addition against PptxGenJS 4.0.1"
```

Expected: the one aggregate test passes without leaking console stubs or global
DOM state.

---

### Task 2: Generate and verify the 42 manifest entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the aggregate control title plus existing native, packed-package,
  browser, and OOXML anchors.
- Produces: exactly 42 new entries and global counts 575/212/85/349/553.

- [x] **Step 1: Generate entries from frozen tables**

Define one property-status table, one nested-field table, and one union table.
Generate deterministic IDs and attach status-specific native mappings, controls,
notes, canonical aliases, serialization flags, and direct evidence.

- [x] **Step 2: Lock exact family and global counts**

Assert 42 distinct IDs, 1,221 manifest entries, exact 4/20/2/16 family counts,
valid alias targets, and global totals 575 supported, 212 deliberate-difference,
85 deprecated-alias, 349 defect-excluded, and 553 unverified.

- [x] **Step 3: Regenerate and inspect the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: deterministic artifacts, zero unsupported/stale/diagnostics, and 553
unverified atoms.

---

### Task 3: Run the batched artifact gates

**Files:**

- Review only; retained outputs use explicit temporary paths.

**Interfaces:**

- Consumes: the existing `tableToSlidesState` package/browser fixtures.
- Produces: one family-level release evidence bundle.

- [x] **Step 1: Run focused native tests once**

```bash
./node_modules/.bin/vitest run \
  packages/sdk/src/index.test.ts \
  packages/sdk/src/table-to-slides.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "tableToSlides|HTML table"
```

Expected: all selected native and control tests pass.

- [x] **Step 2: Pack and smoke once**

Pack `packages/pptx` to an explicit temporary directory and run
`scripts/smoke-npm-package.mjs` with an explicit retained table-to-slides output.
Require every `tableToSlidesState` value and CLI package inspection field true.

- [x] **Step 3: Run persistent Chrome once**

Run the existing browser smoke in one persistent real-Chrome session. Require
every table-to-slides lifecycle field true and zero validation, console, page,
and network errors.

- [x] **Step 4: Inspect the retained PPTX once**

Run `pptx-inspect --json package inspect`, PowerPoint 2010 `package validate`,
`slides list`, and exact slide/relationship `part read` on the retained file.
Require editable table parts, additions, owned relationships, and zero errors.

- [x] **Step 5: Run repository-wide gates once**

```bash
./node_modules/.bin/tsc -b --pretty false
./node_modules/.bin/vitest run
git diff --check
```

Expected: all gates pass.

---

### Task 4: Review, commit, push, and synchronize

**Files:**

- Review and stage only the capability-family files from Tasks 1-3.

**Interfaces:**

- Consumes: verified diff and a programmatic matrix comparison against `HEAD`.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Review exact status migration**

Require exactly 42 transitions from `unverified`, exact 4/20/2/16 counts, no
other atom changes, no unrelated staged paths, and no `.pnpm-store/` files.

- [x] **Step 2: Commit and push**

```bash
git commit -m "docs: close tableToSlides surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
