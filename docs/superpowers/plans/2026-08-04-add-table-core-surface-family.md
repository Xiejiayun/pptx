# addTable Core Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns all
> repository edits, review, commit, and push. Steps use checkbox (`- [ ]`) syntax
> for tracking.

**Goal:** Close exactly 45 PptxGenJS `Slide.addTable`, `TableProps`,
`TableCellProps`, and `BorderProps` declaration atoms by reusing the existing
editable-table lifecycle and adding one missing aggregate upstream control.

**Architecture:** Frozen status tables generate the exact manifest family. One
adapter test locks the inventory and remaining upstream boundaries; existing
native, npm, browser, and OOXML probes provide positive lifecycle evidence. All
expensive gates run once at the batch boundary.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, npm pack,
persistent Chrome, `pptx-inspect`, OOXML package inspection.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-04-add-table-core-surface-family-design.md` exactly.
- Close only the 45 target atoms and require the exact 13/28/1/3 classification.
- Reuse existing table lifecycle evidence; add only the missing aggregate
  PptxGenJS control and matrix assertions.
- Run focused tests once and npm pack, Chrome, OOXML inspection, TypeScript, and
  full Vitest once at the batch boundary.
- Do not stage `.pnpm-store/` or temporary package, browser, or PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the remaining upstream boundaries

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 `Slide.addTable`, the existing importer, table XML
  helpers, and native `SlideModel.addTable`.
- Produces: the stable title `locks the remaining addTable core declarations
  against PptxGenJS 4.0.1`.

- [x] **Step 1: Lock the exact 45-ID inventory**

Build four explicit ID groups and assert 13 supported, 28
deliberate-difference, one deprecated alias, three defect exclusions, 45 unique
IDs, and no inherited already-closed atom.

- [x] **Step 2: Lock method, naming, logging, and alias behavior**

Require PptxGenJS to return its `Slide`, mutate caller table inputs, encode
`objectName`, log only when `verbose` participates in auto-page measurement,
and place continuation tables at the same Y for `newSlideStartY` and
`autoPageSlideStartY`. Require native to return an editable `TableModel`, retain
caller input, use `name`, and avoid global logging.

- [x] **Step 3: Lock all three defect exclusions with positive controls**

Compare table-level transparency to a byte-equivalent baseline and prove a
cell-level transparency value emits the expected alpha. Compare cell-local
character and line weight variants under one table-level setting, require
identical pagination/output, and prove table-level weights still change the
positive-control boundary. Assert caller cell character weight is overwritten.

- [x] **Step 4: Run the aggregate control once**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks the remaining addTable core declarations against PptxGenJS 4.0.1"
```

Expected: one aggregate test passes without leaked console stubs or mutated
global state.

---

### Task 2: Generate and verify the 45 manifest entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the aggregate control plus existing property-specific adapter,
  native, package, browser, and OOXML anchors.
- Produces: exactly 45 new entries and global counts 588/240/86/352/508.

- [x] **Step 1: Generate entries from frozen tables**

Define direct-property status tables for all three owners, direct union tables,
native-name mappings, status-specific notes, and the one canonical alias.
Generate deterministic IDs and attach only evidence that materially covers each
capability.

- [x] **Step 2: Lock exact family and global counts**

Assert 45 distinct IDs, 1,266 manifest entries, exact 13/28/1/3 family counts,
the alias target, three defect IDs, and global totals 588 supported, 240
deliberate-difference, 86 deprecated-alias, 352 defect-excluded, and 508
unverified.

- [x] **Step 3: Regenerate and inspect the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: deterministic artifacts, zero stale/unsupported/diagnostic entries,
and 508 unverified atoms.

---

### Task 3: Run the batched table evidence gates

**Files:**

- Review only; retain artifacts under explicit fresh `/tmp` directories.

**Interfaces:**

- Consumes: existing table creation, paging, merge, hyperlink, border, margin,
  alignment, and reopen probes.
- Produces: one family-level release evidence bundle.

- [x] **Step 1: Run focused tests once**

Run the aggregate adapter control together with existing SDK table tests through
one table-focused Vitest invocation. Require every selected test to pass.

- [x] **Step 2: Pack and smoke once**

Build `packages/pptx`, pack to a fresh explicit temporary directory, run
`scripts/smoke-npm-package.mjs`, record tarball file count and SHA-256, and
require every relevant table state true.

- [x] **Step 3: Run persistent Chrome once**

Execute the packed browser smoke in one persistent real-Chrome session. Require
all relevant table lifecycle fields true and validation/console/page/network
errors all zero.

- [x] **Step 4: Inspect retained OOXML once**

Run `pptx-inspect` package inspection, PowerPoint 2010 validation, slide list,
and exact slide/relationship reads on a retained table deck. Require editable
table grid, row heights, merges, borders, margins, hyperlinks, zero errors, and
byte-stable reopen.

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

- Review and stage only the seven capability-family files from Tasks 1-2.

**Interfaces:**

- Consumes: the verified diff and programmatic matrix comparison against
  `HEAD`.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Review exact migration**

Require exactly 45 changed atoms, all from `unverified`, exact 13/28/1/3
transitions, no other atom content changes, no unrelated staged paths, and no
`.pnpm-store/` files.

- [x] **Step 2: Commit and push**

```bash
git commit -m "docs: close addTable core surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
