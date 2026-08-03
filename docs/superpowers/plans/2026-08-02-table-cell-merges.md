# Table-Cell Merge, Colspan, and Rowspan Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. This repository task executes inline without subagents and performs review, commit, and push after every independently testable deliverable.

**Goal:** Add strict PptxGenJS-compatible `colspan` / `rowspan` table creation plus safe existing-table merge topology snapshots, lossless merge, and lossless unmerge.

**Architecture:** Normalize logical input rows into a complete physical cell matrix before any package mutation, then render canonical DrawingML anchor/continuation attributes while allocating relationships only for anchors. Add one exact internal topology boundary that recognizes whole-table merge regions and rewrites only affected `a:tc` start tags; expose its frozen state and transactional editors through the existing physical-coordinate `TableModel` API.

**Tech Stack:** TypeScript 5.9, Vitest 3, lossless DrawingML/OPC transactions, pnpm workspaces, PptxGenJS 4.0.1 conformance fixtures, packed ESM/CJS/browser exports, CLI/Inspector, Google Chrome, and PowerPoint 2010 validation.

## Global Constraints

- Public creation fields are exactly `AddTableCellOptions.colspan?: number` and `rowspan?: number`; no alias family is added.
- Span values are positive safe integers; omitted and `1` mean one physical cell and emit no span attribute.
- Outer rows and the first logical row are non-empty dense arrays; later logical rows may be empty only when active row spans completely cover the physical row.
- First-row colspan sum defines physical column count; every physical row must be completely and exactly covered without overlap or out-of-bounds rectangles.
- Expanded tables may contain at most 1,000,000 physical cells so hostile single-cell spans cannot cause unbounded allocation.
- Public snapshots and indexed editors use zero-based physical row/column coordinates.
- Valid existing merge state is recognized all-or-nothing against the direct grid, rows, cells, and semantic `gridSpan` / `rowSpan` / `hMerge` / `vMerge` pattern.
- `mergeRegions` returns frozen row-major regions, `[]` for recognized unmerged tables, and `undefined` for unsupported topology.
- Merge and unmerge preserve every physical cell's content, style, opaque XML, and relationship; only the four merge attributes of affected `a:tc` start tags change.
- Exact existing merge and unmerged-cell unmerge are package/journal/date/identity no-ops; overlapping non-identical merges reject.
- All input/topology/relationship validation occurs before observable mutation, and any transaction failure rolls back completely.
- No-span creation remains byte-identical. Existing physical cell editors retain their current behavior, including continuation-cell editing without anchor redirection.
- Do not add row/column CRUD, cell insert/delete, auto-page, repeated headers, content measurement, layout recomputation, or `tableToSlides`.
- Every task is reviewed, committed, pushed, and followed by local/remote divergence verification of `0 0`.

---

## File Structure

- Create `packages/model/src/table-cell-merge.internal.ts`: exact topology parser, frozen state, strict coordinate/span normalization, and lossless start-tag merge/unmerge edits.
- Create `packages/model/src/table-cell-merge.internal.test.ts`: recognized topology, unsupported topology, normalization, no-op, preservation, and mutation tests.
- Modify `packages/model/src/table-create.internal.ts`: logical-row normalization, physical matrix expansion, anchor/continuation union, span attributes, and physical relationship matrices.
- Modify `packages/model/src/table-create.internal.test.ts`: layout, rendering, validation, detachment, relationship, sizing, and legacy parity tests.
- Modify `packages/model/src/slide.ts`: public creation fields and anchor-only hyperlink preparation across the physical matrix.
- Modify `packages/model/src/shapes.ts`: public merge snapshot types/getters and transactional `TableModel` editors.
- Modify `packages/model/src/model.test.ts`: public snapshot/editor, rollback, duplicate, format, and reopen lifecycle coverage.
- Modify `packages/sdk/src/index.test.ts`: aggregate SDK/root runtime and compile-time API contracts.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: PptxGenJS 4.0.1 legal final-state conformance and strict-difference evidence.
- Modify `scripts/smoke-npm-package.mjs` and `scripts/playwright-browser-smoke.js`: actual tarball Node/NodeNext/browser/CLI/Inspector and real-Chrome merge proof.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: API, support matrix, proof, progress, and remaining work.

### Task 1: Exact Merge Topology Boundary

**Files:**
- Create: `packages/model/src/table-cell-merge.internal.ts`
- Create: `packages/model/src/table-cell-merge.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, and the direct table/grid/row/cell ownership conventions in `table-physical-cells.internal.ts`.
- Produces: `TableMergeRegionState`, `TableCellMergeState`, `TableMergeState`, `readTableMergeState()`, `normalizeTableMergeRegionInput()`, `replaceTableMergeRegion()`, and `clearTableMergeRegionAt()`.

- [ ] **Step 1: Write failing recognized/unsupported topology tests**

Use one direct `3 × 4` table with a horizontal merge, vertical merge, offset `2 × 2` merge, and unrelated opaque attributes/children. Assert the exact internal contract:

```ts
const state = readTableMergeState(frame)!;
expect(state.regions).toEqual([
  { rowIndex: 0, columnIndex: 0, rowspan: 1, colspan: 2 },
  { rowIndex: 0, columnIndex: 3, rowspan: 2, colspan: 1 },
  { rowIndex: 1, columnIndex: 1, rowspan: 2, colspan: 2 },
]);
expect(state.cells[2]![2]).toEqual({
  rowIndex: 1,
  columnIndex: 1,
  rowspan: 2,
  colspan: 2,
  isAnchor: false,
});
expect(Object.isFrozen(state.regions)).toBe(true);
expect(Object.isFrozen(state.cells[2]![2])).toBe(true);
```

Add `[]` for a recognized unmerged table. Add `undefined` cases for missing/repeated direct table/grid, zero rows/columns, jagged rows, grid mismatch, repeated unqualified merge attributes, invalid integers/booleans, out-of-bounds anchors, orphan continuations, wrong continuation span, overlap, and continuation-owned spans. Add namespaced `x:gridSpan` / `x:hMerge` impostors and prove they are ignored and preserved rather than treated as direct merge tokens.

- [ ] **Step 2: Run the focused test and verify the red state**

Run:

```bash
pnpm exec vitest run packages/model/src/table-cell-merge.internal.test.ts
```

Expected: module-not-found/test failure because the topology boundary does not exist.

- [ ] **Step 3: Implement exact semantic recognition and frozen snapshots**

Define these internal shapes exactly:

```ts
export interface TableMergeRegionState {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly rowspan: number;
  readonly colspan: number;
}

export interface TableCellMergeState extends TableMergeRegionState {
  readonly isAnchor: boolean;
}

export interface TableMergeState {
  readonly regions: readonly Readonly<TableMergeRegionState>[];
  readonly cells: readonly (readonly (Readonly<TableCellMergeState> | undefined)[])[];
}

export function readTableMergeState(
  frame: XmlElement,
): Readonly<TableMergeState> | undefined;
```

Require exactly one direct `graphic/graphicData/tbl`, one direct `tblGrid`, at least one direct `gridCol`, at least one direct `tr`, and equal direct `tc` counts matching the grid. Parse only unique unqualified exact-name attributes. Accept `/^[0-9]+$/` positive safe integers with leading zeros and XML Schema booleans `1|true|0|false`. Discover non-continuation span anchors, fill an expected region matrix, reject overlap/bounds failures, and then compare every cell's actual semantic tokens with its expected anchor/top/left/interior role. Freeze every region, cell merge state, row, and outer array; sort anchors row-major.

- [ ] **Step 4: Add failing normalizer and lossless mutation tests**

Assert strict runtime normalization:

```ts
expect(normalizeTableMergeRegionInput(1, 2, 3, 4)).toEqual({
  rowIndex: 1,
  columnIndex: 2,
  rowspan: 3,
  colspan: 4,
});
```

Reject fractions, negatives, unsafe integers, non-numbers, and `1 × 1`. For a recognized table, assert `replaceTableMergeRegion()` adds one canonical horizontal/vertical/2D region; a second identical call returns `false`; non-overlapping merge succeeds; partial/full non-identical overlap throws; bounds errors throw. Assert `clearTableMergeRegionAt()` accepts anchor or continuation, removes the whole region, and returns `false` on an unmerged cell. Capture exact cell content, style, hyperlinks, quote styles of unrelated attributes, opaque children, other merges, and start-tag-external whitespace before edits and prove all remain byte-equal.

- [ ] **Step 5: Implement one-start-tag-per-cell edits**

Add these exact functions:

```ts
export function normalizeTableMergeRegionInput(
  rowIndex: unknown,
  columnIndex: unknown,
  rowspan: unknown,
  colspan: unknown,
): Readonly<TableMergeRegionState>;

export function replaceTableMergeRegion(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  region: Readonly<TableMergeRegionState>,
  partUri: string,
): boolean;

export function clearTableMergeRegionAt(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  rowIndex: unknown,
  columnIndex: unknown,
  partUri: string,
): boolean;
```

Both editors must call `readTableMergeState()` first and throw `ModelParseError('Table merge state is not safely editable', partUri)` when unsupported. Build the expected four merge tokens for each affected physical coordinate, remove only exact unqualified merge attributes from that cell's original start-tag slice, insert canonical non-default tokens before `>`/`/>`, and call one `xml.replace(cell.start, cell.startTagEnd, nextStartTag)` per changed cell. Do not patch cell bodies or relationships.

- [ ] **Step 6: Run tests, typecheck, and review the boundary**

Run:

```bash
pnpm exec vitest run packages/model/src/table-cell-merge.internal.test.ts
pnpm exec vitest run packages/model/src/table-column-widths.internal.test.ts \
  packages/model/src/table-row-heights.internal.test.ts
pnpm exec tsc -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
```

Review direct-child and unqualified-attribute ownership, semantic lexical acceptance, all-or-nothing recognition, row-major freezing, overlap rules, no-op behavior, patch isolation, and exact part URI errors.

- [ ] **Step 7: Commit, push, and verify synchronization**

```bash
git add packages/model/src/table-cell-merge.internal.ts \
  packages/model/src/table-cell-merge.internal.test.ts
git commit -m "feat: recognize table cell merges"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 2: Logical-to-Physical Span Creation

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`
- Modify: `packages/model/src/slide.ts`

**Interfaces:**
- Consumes: existing strict cell/text/style normalization and table-cell hyperlink preparation.
- Produces: public `colspan` / `rowspan` creation fields, an internal anchor/continuation physical-cell union, full physical rows, and canonical span OOXML.

- [ ] **Step 1: Add failing layout, rendering, relationship, and parity tests**

Cover horizontal, vertical, offset `2 × 2`, full-row vertical coverage with `[]`, multiple non-overlapping spans, and this mixed input:

```ts
const rows = [
  ['A', { text: rich, options: { colspan: 2, rowspan: 2, hyperlink } }, 'B'],
  ['C', 'D'],
  ['E', 'F', 'G', 'H'],
] as const;
```

Assert a `3 × 4` physical definition, correct anchor/continuation roles, physical column widths, anchor-only rich/default hyperlink relationships, no continuation relationship, canonical attribute order, and PptxGenJS-equivalent final span semantics. Mutate every source object/array after normalization and prove the definition is detached. Store a no-span legacy XML baseline and require exact equality.

Add pre-mutation rejection for empty outer/first row, uncovered empty row, underfill, overfill, overlap, non-contiguous placement, bottom/right overflow, and span values that are non-number, non-finite, fractional, zero, negative, or unsafe. Assert invalid later rich text/hyperlink/span leaves package bytes, relationships, shape ID allocation, dates, and journal unchanged.

- [ ] **Step 2: Run focused creation tests and verify the red state**

Run:

```bash
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  -t "colspan|rowspan|logical table spans"
```

Expected: unsupported-option or rectangular-row failures because span fields and physical expansion are absent.

- [ ] **Step 3: Normalize logical cells and expand a strict physical matrix**

Add `colspan` and `rowspan` to the descriptor-safe cell option key list and `AddTableCellOptions`. Split the internal cell type:

```ts
interface NormalizedTableAnchorCell extends NormalizedTableTextDefaults {
  readonly kind: 'anchor';
  readonly rowspan: number;
  readonly colspan: number;
  // existing text, rich text, style, and hyperlink fields
}

interface NormalizedTableContinuationCell {
  readonly kind: 'continuation';
  readonly rowSpan: number;
  readonly gridSpan: number;
  readonly verticalContinuation: boolean;
  readonly horizontalContinuation: boolean;
}

type NormalizedTableCell =
  | NormalizedTableAnchorCell
  | NormalizedTableContinuationCell;
```

Use a strict positive-safe-integer helper with default `1`. Keep outer/first-row non-empty while allowing later dense empty rows. Sum first-row colspans with overflow protection, allocate `rowCount × columnCount`, place each anchor at the row's leftmost free coordinate, reserve its complete rectangle, create role-specific continuations, and require every coordinate filled after each row. Resolve table defaults only for anchors.

- [ ] **Step 4: Render canonical continuations and physical relationship matrices**

Render anchors through the unchanged rich-text/style path plus ordered span attributes. Render continuations exactly as:

```xml
<a:tc rowSpan="2" hMerge="1"><a:tcPr/></a:tc>
```

with only attributes appropriate to their role. Make default and rich-run hyperlink ID matrices match physical dimensions; continuation entries are always `undefined` / empty. In `SlideModel.prepareTableCellHyperlinks()` and relationship creation, skip `kind === 'continuation'`; keep all target resolution before observable mutation.

- [ ] **Step 5: Run focused/regression gates and review creation**

Run:

```bash
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-cell-rich-text.internal.test.ts \
  packages/model/src/table-cell-hyperlink.internal.test.ts
pnpm exec vitest run packages/model/src/model.test.ts -t "table creation"
pnpm exec tsc -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
```

Review first-row column semantics, full physical coverage, row-span reservation, anchor-only defaults/relationships, continuation XML, no input mutation, strict errors before mutation, and no-span byte parity.

- [ ] **Step 6: Commit, push, and verify synchronization**

```bash
git add packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts packages/model/src/slide.ts
git commit -m "feat: create table cell spans"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 3: Public Merge Snapshots

**Files:**
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 `readTableMergeState()` and existing physical `TableModel.rows` projection.
- Produces: exported `TableMergeRegion`, `TableCellMerge`, `TableCell.merge`, and `TableModel.mergeRegions`.

- [ ] **Step 1: Add failing public snapshot tests**

Open native, PptxGenJS, and hand-authored valid merged tables and assert:

```ts
expect(table.mergeRegions).toEqual([
  { rowIndex: 0, columnIndex: 1, rowspan: 2, colspan: 2 },
]);
expect(table.rows[1]!.cells[2]!.merge).toEqual({
  rowIndex: 0,
  columnIndex: 1,
  rowspan: 2,
  colspan: 2,
  isAnchor: false,
});
```

Require physical row/cell counts, anchor/continuation content visibility, frozen/detached snapshots, `[]` for unmerged, and `undefined`/omitted merge state for malformed topology without losing ordinary cell text/style snapshots.

- [ ] **Step 2: Run model tests and verify the red state**

Run:

```bash
pnpm exec vitest run packages/model/src/model.test.ts -t "table merge snapshots"
```

Expected: type/runtime failures because public merge fields do not exist.

- [ ] **Step 3: Expose structural types and project one merge state per rows read**

Add the exact public types from the design to `shapes.ts`. In `TableModel.rows`, call `readTableMergeState(element)` once, retain row/column indices during mapping, and append `merge` only when the state matrix has an entry. Add `mergeRegions` as a direct getter returning the internal already-frozen structural snapshot without caching across XML mutations.

- [ ] **Step 4: Run tests, typecheck, and review snapshots**

Run:

```bash
pnpm exec vitest run packages/model/src/model.test.ts -t "table merge snapshots|table rows"
pnpm exec tsc -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
```

Review one parser call per getter, physical coordinates, row-major order, deep freezing, unsupported-state distinction, no stale cache, and unchanged existing cell snapshots.

- [ ] **Step 5: Commit, push, and verify synchronization**

```bash
git add packages/model/src/shapes.ts packages/model/src/model.test.ts
git commit -m "feat: expose table merge snapshots"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 4: Lossless Public Merge and Unmerge

**Files:**
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 normalizer/editors and Task 3 public snapshot state.
- Produces: `TableModel.mergeCells()` and `TableModel.unmergeCell()` transactional physical-coordinate editors.

- [ ] **Step 1: Add failing lifecycle and rollback tests**

Create a `4 × 4` table with unique text, rich styles, cell hyperlinks, borders/fills, opaque cell attributes/children, and lexical merge-token variants. Exercise horizontal, vertical, 2D, offset, multiple non-overlapping, exact repeat, continuation-addressed unmerge, and unmerged no-op. Before/after every no-op capture part bytes, relationships, ZIP entry dates, table identity, all cell snapshots, diagnostics, and mutation journal.

Prove merge hides no physical state from the model; edit a continuation with existing setters, unmerge, and confirm its content/style/link remains. Add overlap, bounds, invalid numeric input, unsupported topology, deleted handle, injected `setPart` failure, duplicate-slide isolation, move, six formats, and write/reopen cases.

- [ ] **Step 2: Run focused lifecycle tests and verify the red state**

Run:

```bash
pnpm exec vitest run packages/model/src/model.test.ts \
  -t "mergeCells|unmergeCell|table merge lifecycle"
```

Expected: type/runtime failures because the public editors do not exist.

- [ ] **Step 3: Add thin transactional TableModel methods**

Implement exactly:

```ts
mergeCells(rowIndex: number, columnIndex: number, rowspan: number, colspan: number): void {
  const region = normalizeTableMergeRegionInput(
    rowIndex,
    columnIndex,
    rowspan,
    colspan,
  );
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableMergeRegion(xml, element, region, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

Implement `unmergeCell()` in the same shape with `clearTableMergeRegionAt()`. Do not add relationship mutation, model cache invalidation, hidden content transfer, anchor redirection, or style synthesis.

- [ ] **Step 4: Run lifecycle/regression gates and review editors**

Run:

```bash
pnpm exec vitest run packages/model/src/table-cell-merge.internal.test.ts \
  packages/model/src/model.test.ts -t "table merge|table cell"
pnpm exec vitest run packages/model/src/table-column-widths.internal.test.ts \
  packages/model/src/table-row-heights.internal.test.ts \
  packages/model/src/table-cell-rich-text.internal.test.ts
pnpm typecheck
git diff --check
```

Review transaction boundaries, normalization timing, exact no-ops, rollback, relationship immutability, physical continuation behavior, duplicate/reopen isolation, and no unrelated refactor.

- [ ] **Step 5: Commit, push, and verify synchronization**

```bash
git add packages/model/src/shapes.ts packages/model/src/model.test.ts
git commit -m "feat: edit table cell merges"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 5: SDK Contract and PptxGenJS 4.0.1 Conformance

**Files:**
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: Tasks 2–4 exported API through `@pptx/sdk` and aggregate `@jiayunxie/pptx` surfaces.
- Produces: permanent runtime/type contract and legal PptxGenJS final-state evidence.

- [ ] **Step 1: Add SDK/root positive and negative contract tests**

Create mixed spans through `PptxDocument.create()`, inspect physical rows and merge snapshots, merge/unmerge an ordinary table, write/reopen, and assert stable identity. Add compile-only positive uses for both option fields, snapshot interfaces, and both methods. Add `@ts-expect-error` cases for wrong option names/casing, readonly mutation, strings/fractions in typed positions, wrong arity, and logical-coordinate-only fields.

- [ ] **Step 2: Add PptxGenJS legal output and strict-difference tests**

Generate public PptxGenJS 4.0.1 horizontal, vertical, rectangular, offset, and full-row-span cases, open them natively, and compare physical dimensions, region snapshots, cell text, and semantic four-token matrices. Native-unmerge and re-merge each legal fixture, then write/reopen. Separately lock PptxGenJS caller-array mutation and its lopsided-row, negative/fractional, and out-of-bounds output defects; assert native creation rejects equivalent invalid inputs before mutation rather than copying them.

- [ ] **Step 3: Run aggregate tests and declarations**

Run:

```bash
pnpm exec vitest run packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts -t "table merge|colspan|rowspan"
pnpm typecheck
pnpm build
pnpm --filter @jiayunxie/pptx build
git diff --check
```

Review aggregate exports, exact public naming, Node/browser-safe code paths, type-negative usefulness, legal final-state parity, and explicit strict differences.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "test: verify table cell merge parity"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 6: Public Documentation and Compatibility Matrix

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: completed runtime/type/conformance behavior from Tasks 1–5.
- Produces: user-facing creation/edit examples, physical-coordinate contract, strict compatibility notes, and updated remaining-work ordering.

- [ ] **Step 1: Document exact creation, snapshot, and editing examples**

Add one common example to the three API guides:

```ts
const table = slide.addTable([
  [{ text: 'Summary', options: { colspan: 2, rowspan: 2 } }, 'Total'],
  ['42'],
]);

console.log(table.mergeRegions);
table.unmergeCell(1, 1); // any physical member resolves the region
table.mergeCells(0, 0, 2, 2);
```

Explain logical creation rows, first-row physical width, fully covered `[]` rows, zero-based physical snapshot/editor coordinates, continuation hidden-state preservation, exact no-ops, and malformed topology refusal.

- [ ] **Step 2: Update parity/progress records without claiming proof not yet run**

Move merge/colspan/rowspan from the table row's missing column to supported creation/read/edit. Record PptxGenJS legal output parity and strict invalid-input differences. Set remaining advanced-table order to row/column CRUD, auto-page/repeated headers, measurement/layout recomputation, and `tableToSlides`; update the formerly stale approximate overall percentage consistently, but leave package/browser/PowerPoint result placeholders out until Task 8 has real values.

- [ ] **Step 3: Review docs and run focused gates**

Run:

```bash
rg -n "mergeCells|unmergeCell|colspan|rowspan|row/column CRUD|tableToSlides" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
pnpm exec vitest run packages/model/src/table-cell-merge.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts -t "table merge|colspan|rowspan"
git diff --check
```

Review code/doc naming, logical versus physical semantics, no stale “merge unsupported” claim, no premature proof numbers, and consistent remaining scope.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: document table cell merges"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 7: Actual Packed and Browser Verification

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball and browser conditional export.
- Produces: stable JSON proof for Node, NodeNext types, browser, CLI, Inspector, and real Chrome create/snapshot/merge/unmerge/reopen.

- [ ] **Step 1: Extend actual-package smoke assertions**

In the installed-package fixture, create a rectangular span, assert `mergeRegions` and all physical `TableCell.merge` values, unmerge through a continuation coordinate, edit the restored continuation, merge again, write/reopen, and return `tableCellMerges: true`. Extend the NodeNext consumer with compile-time uses of both option fields, both snapshot types, and both editors. Extend CLI/Inspector checks to prove the resulting slide remains one table with expected physical text.

- [ ] **Step 2: Extend real-Chrome smoke assertions**

Use the browser export to perform the same create/snapshot/unmerge/edit/remerge/reopen lifecycle and emit independent booleans for all five stages. Fail the smoke on any validation, console, page, or network error. Keep the generated evidence deck available to the final proof step.

- [ ] **Step 3: Build, pack twice, and run smoke/proof gates**

Run the existing deterministic clean-build/pack workflow twice, compare sorted dist hash manifests and tarball bytes, then execute:

```bash
pnpm build
pnpm --filter @jiayunxie/pptx build
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
git diff --check
```

Record actual tarball file count, SHA-256, Node/type/browser/CLI/Inspector booleans, Chrome version/stage booleans, validation/console/page/network counts, evidence directory, part/relationship counts, and table/region dimensions for Task 8.

- [ ] **Step 4: Review proof code, commit, push, and verify synchronization**

Review that every check loads the packed artifact rather than workspace sources, NodeNext resolves declarations, browser import uses the conditional export, Chrome reopens produced bytes, and output JSON is deterministic/stable.

```bash
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git commit -m "test: verify packed table cell merges"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 8: Full Gates, Native Validation, and Recorded Proof

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: all completed code/tests/docs and Task 7's actual artifacts.
- Produces: final reproducible verification record and the authoritative remaining-work/progress statement.

- [ ] **Step 1: Run clean type/build/focused/full/performance gates**

Run:

```bash
pnpm typecheck
pnpm build
pnpm --filter @jiayunxie/pptx build
pnpm exec vitest run packages/model/src/table-cell-merge.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
pnpm exec vitest run --maxWorkers=2
pnpm test:performance
```

Capture exact file/test counts, skipped counts, duration, and 1000-part performance time. Re-run any flaky external step until its final state is unambiguous; do not weaken assertions.

- [ ] **Step 2: Validate and inspect the evidence deck**

Run PowerPoint 2010 compatibility validation through the existing validator/CLI flow, then use CLI/Inspector part and relationship inspection on the exact packed-browser evidence deck. Require zero errors/warnings, rectangular physical rows, expected region tokens, no orphan relationships, and exact reopen state. Run the final real-Chrome smoke and require all stage booleans true plus zero validation/console/page/network errors.

- [ ] **Step 3: Record only actual final values and update progress**

Add exact full/focused/performance counts, tarball file count/SHA-256, packed surface results, Chrome version/stage results, PowerPoint 2010 counts, evidence directory, deck part/relationship/table/merge dimensions, and all implementation/proof commit IDs. Mark merge/colspan/rowspan complete and report the new overall percentage plus the remaining advanced-table items.

- [ ] **Step 4: Final review, commit, push, and confirm clean synchronization**

Run:

```bash
git diff --check
git status --short
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: record table cell merge proof"
git push origin main
git rev-list --left-right --count origin/main...HEAD
git status --short --branch
```

Expected divergence: `0 0`; only the pre-existing untracked `.pnpm-store/` may remain.
