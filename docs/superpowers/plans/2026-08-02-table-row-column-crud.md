# Table Row and Column CRUD Implementation Plan

> **For agentic workers:** Execute this plan inline without subagents. Every task ends with focused verification, self-review, commit, push, and local/remote divergence 0 0 before the next task begins.

**Goal:** Add strict, merge-aware, relationship-safe physical row and column insertion/deletion for existing PowerPoint tables while preserving every surviving OOXML state losslessly.

**Architecture:** Add one exact direct-table structure boundary that owns only grid columns, physical rows/cells, direct sizes, transforms, and recognized merge topology. Row/column operations splice source ranges and canonical empty fragments, rewrite only changed merge/dimension tokens, and return removed relationship IDs for final-reference garbage collection inside the existing OPC transaction.

**Tech Stack:** TypeScript 5.9, Vitest 3, lossless DrawingML/OPC transactions, pnpm workspaces, PptxGenJS 4.0.1 import fixtures, generated declarations, packed ESM/browser exports, CLI/Inspector, Google Chrome, and PowerPoint 2010 validation.

## Global Constraints

- Public indexes are zero-based physical coordinates; insert and delete counts default to one.
- Insert indexes are in [0, itemCount]; delete ranges must be wholly in bounds and leave at least one physical row and column.
- Counts are positive safe integers; inserted total physical cells may not exceed 1,000,000.
- Row heights are non-negative safe EMU after rounding; zero means automatic. Column widths are positive safe EMU after rounding.
- Omitted insert dimensions copy the item at the insertion index, or the last item when appending.
- Inserted cells are canonical empty, editable plain single-run cells with no relationship or inferred neighbor style.
- Recognized merge regions shift, expand, shrink, promote anchors, or dissolve deterministically; unsupported topology rejects before mutation.
- Only changed grid/row/cell/dimension/merge source ranges may be rewritten. Surviving content, style, opaque XML, relationship references, lexical forms, and table/shape identity remain unchanged.
- Deletion removes a slide relationship only when its ID appeared in deleted cells and no source-part reference remains after the structural write.
- Column edits always synchronize transform width to the direct grid-width sum. Row edits synchronize transform height only when every final direct row height is positive; otherwise existing valid height is preserved.
- Input, structure, merge, range, dimension, relationship, and overflow validation occurs before observable mutation; injected or outer transaction failure rolls back completely.
- Do not add logical cell-content insert overloads, retained creation defaults, style-template inference, auto-page, repeated headers, content measurement, layout recomputation, or tableToSlides.

---

### Task 1: Exact Structure Boundary and Inputs

**Files:**
- Create: packages/model/src/table-structure-edit.internal.ts
- Create: packages/model/src/table-structure-edit.internal.test.ts
- Modify: packages/model/src/table-create.internal.ts
- Modify: packages/model/src/table-cell-merge.internal.ts

**Interfaces:**

    export interface NormalizedTableRowInsert {
      readonly rowIndex: number;
      readonly count: number;
      readonly rowHeights?: readonly number[];
    }

    export interface NormalizedTableColumnInsert {
      readonly columnIndex: number;
      readonly count: number;
      readonly columnWidths?: readonly number[];
    }

    export interface NormalizedTableDelete {
      readonly index: number;
      readonly count: number;
    }

    export interface EditableTableStructure {
      readonly table: XmlElement;
      readonly grid: XmlElement;
      readonly gridColumns: readonly XmlElement[];
      readonly rows: readonly XmlElement[];
      readonly cells: readonly (readonly XmlElement[])[];
      readonly columnWidths: readonly number[];
      readonly rowHeights: readonly number[];
      readonly widthAttribute: XmlAttribute;
      readonly heightAttribute: XmlAttribute;
      readonly width: number;
      readonly height: number;
      readonly mergeState: Readonly<TableMergeState>;
    }

    export function normalizeTableRowInsertInput(
      rowIndex: unknown,
      options: unknown,
    ): Readonly<NormalizedTableRowInsert>;

    export function normalizeTableColumnInsertInput(
      columnIndex: unknown,
      options: unknown,
    ): Readonly<NormalizedTableColumnInsert>;

    export function normalizeTableDeleteInput(
      index: unknown,
      count: unknown,
      axis: 'row' | 'column',
    ): Readonly<NormalizedTableDelete>;

    export function requireEditableTableStructure(
      frame: XmlElement,
      partUri: string,
    ): Readonly<EditableTableStructure>;

    export interface TableCellMergeTokens {
      readonly rowSpan?: number;
      readonly gridSpan?: number;
      readonly vertical?: true;
      readonly horizontal?: true;
    }

    export function renderEmptyTableCellFragment(
      tokens?: Readonly<TableCellMergeTokens>,
    ): string;

    export function replaceTableCellMergeAttributes(
      xml: LosslessXmlDocument,
      cell: XmlElement,
      tokens: Readonly<TableCellMergeTokens>,
    ): boolean;

- [ ] **Step 1: Write failing strict-input and structure tests**

Cover prepend/middle/append indexes, omitted/default count, scalar/exact-array sizes, null-prototype options, zero row height, and rounded EMU values. Reject non-number/non-finite/fractional/negative/unsafe indexes and counts, zero count, sparse/accessor/symbol/inherited/unknown options, non-positive column width, negative row height, and dimension arrays whose length differs from count.

Use one direct 2 x 3 table with a recognized 2 x 2 merge and assert exact elements, [100, 200, 300] widths, [40, 60] heights, transform attributes, and merge state. Add rejection fixtures for repeated/missing direct table/grid/xfrm/ext, empty grid/rows, jagged cells, grid mismatch, invalid/repeated owned size tokens, unsafe transform, malformed merge topology, and foreign-namespace impostors.

- [ ] **Step 2: Run the focused test and verify the red state**

    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/model/src/table-structure-edit.internal.test.ts

Expected: module-not-found or missing-export failure.

- [ ] **Step 3: Implement descriptor-safe normalization and exact parsing**

Use ordinary/null-prototype data readers and dense arrays. Normalize count omitted to one, freeze returned vectors/objects, and parse only unique exact unqualified w, h, cx, and cy. Require one direct table/grid/transform path, rectangular non-empty cells, and recognized merge state; throw ModelParseError with message Table structure is not safely editable on structural failure.

- [ ] **Step 4: Add canonical empty-cell and merge-token helper tests**

Assert the empty fragment contains one direct txBody/bodyPr/lstStyle/p/r/t, one direct tcPr, no hyperlink, and no merge token. Assert all four merge roles render in canonical order. Prove semantic no-op for canonical/equivalent existing tokens and exact preservation of unrelated attributes, quote styles, content, properties, and whitespace.

- [ ] **Step 5: Export the minimal renderer and attribute helper**

Keep NormalizedTableCell private. Implement renderEmptyTableCellFragment inside table-create.internal.ts by rendering the existing normal empty-text cell, then adding the requested merge tokens to that cell's start tag. Do not route continuation tokens through the creation-only continuation branch because that branch omits txBody; every structurally inserted physical cell must retain one editable empty txBody. Export the current merge start-tag helper as replaceTableCellMergeAttributes and compare semantic tokens before scheduling a replacement.

- [ ] **Step 6: Run focused/regression/type gates and review**

    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/model/src/table-structure-edit.internal.test.ts \
      packages/model/src/table-create.internal.test.ts \
      packages/model/src/table-cell-merge.internal.test.ts
    pnpm --config.verify-deps-before-run=false exec tsc \
      -p packages/model/tsconfig.json --noEmit --pretty false
    git diff --check

Review direct ownership, descriptor safety, safe integer math, whole-table merge recognition, freezing, error messages, and absence of speculative public surface.

- [ ] **Step 7: Commit, push, and verify synchronization**

    git add packages/model/src/table-structure-edit.internal.ts \
      packages/model/src/table-structure-edit.internal.test.ts \
      packages/model/src/table-create.internal.ts \
      packages/model/src/table-cell-merge.internal.ts
    git commit -m "feat: recognize editable table structure"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD

Expected divergence: 0 0.

### Task 2: Merge-Aware Row Insertion and Deletion

**Files:**
- Modify: packages/model/src/table-structure-edit.internal.ts
- Modify: packages/model/src/table-structure-edit.internal.test.ts

**Interfaces:**

    export function insertTableRows(
      xml: LosslessXmlDocument,
      frame: XmlElement,
      input: Readonly<NormalizedTableRowInsert>,
      partUri: string,
    ): void;

    export function deleteTableRows(
      xml: LosslessXmlDocument,
      frame: XmlElement,
      input: Readonly<NormalizedTableDelete>,
      partUri: string,
    ): ReadonlySet<string>;

- [ ] **Step 1: Write failing unmerged-row splice tests**

Cover prepend/middle/append, one/multiple rows, copied/default/scalar/vector heights, editable empty cells, exact row order, immediate height vector, all-positive transform sum, auto-height transform preservation, and 1,000,000-cell guard. Delete prepend/middle/tail/range; reject out-of-bounds and final-row deletion. Compare raw source slices for every surviving cell, grid, table property, extension, and unknown child.

- [ ] **Step 2: Implement lossless unmerged row splice**

Materialize omitted heights from state.rowHeights[rowIndex] or the last height, render complete direct rows, insert before the target row or table row-extension boundary, and update only cy when every final height is positive. Delete exact row source ranges and return all namespace-correct non-empty relationship IDs found under removed rows.

- [ ] **Step 3: Write failing merge-aware row tests**

Cover horizontal, vertical, rectangular, offset, and multiple regions. Assert insertion before anchor shifts without token rewrite, insertion inside expands rowspan, insertion at region end does not expand, and new members have canonical continuation roles. Assert partial deletion shrinks 2D/vertical spans, anchor deletion promotes correctly, 1 x 1 dissolves, a fully deleted region disappears, hidden survivor content/style/opaque XML remains, and reopen matches.

- [ ] **Step 4: Implement row-region projection**

Compute target regions before mutation. Render merge tokens into inserted row cells. Exclude removed elements from attribute edits, promote the first surviving row when needed, and call replaceTableCellMergeAttributes only for surviving members whose semantic role changes. Revalidate target bounds and overlap before scheduling replacements.

- [ ] **Step 5: Add patch-isolation and overlapping-source tests**

Prove invalid final geometry is rejected before any XML patch is scheduled, no replacement target overlaps a deleted row range, and a thrown internal edit never reaches package state because serialization/setPart has not run. Reserve slide-part, relationship, date, journal, identity, and outer-transaction rollback injection for Task 4, where those state changes occur.

- [ ] **Step 6: Run gates, review, commit, and push**

    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/model/src/table-structure-edit.internal.test.ts \
      packages/model/src/table-cell-merge.internal.test.ts \
      packages/model/src/table-row-heights.internal.test.ts
    pnpm --config.verify-deps-before-run=false exec tsc \
      -p packages/model/tsconfig.json --noEmit --pretty false
    git diff --check
    git add packages/model/src/table-structure-edit.internal.ts \
      packages/model/src/table-structure-edit.internal.test.ts
    git commit -m "feat: splice table rows"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD

Review row ordering, merge promotion, automatic-height behavior, relationship-ID collection, source isolation, and rollback. Expected divergence: 0 0.

### Task 3: Merge-Aware Column Insertion and Deletion

**Files:**
- Modify: packages/model/src/table-structure-edit.internal.ts
- Modify: packages/model/src/table-structure-edit.internal.test.ts

**Interfaces:**

    export function insertTableColumns(
      xml: LosslessXmlDocument,
      frame: XmlElement,
      input: Readonly<NormalizedTableColumnInsert>,
      partUri: string,
    ): void;

    export function deleteTableColumns(
      xml: LosslessXmlDocument,
      frame: XmlElement,
      input: Readonly<NormalizedTableDelete>,
      partUri: string,
    ): ReadonlySet<string>;

- [ ] **Step 1: Write failing unmerged-column splice tests**

Cover prepend/middle/append, one/multiple columns, copied/default/scalar/vector widths, exact gridCol position, one new cell per row, exact cx sum, and expansion overflow. Delete prepend/middle/tail/ranges; reject out-of-bounds and final-column deletion. Freeze raw source slices for every surviving cell/row/extension/table property.

- [ ] **Step 2: Implement lossless unmerged column splice**

Materialize omitted widths from state.columnWidths[columnIndex] or the last width, splice matching direct grid columns and one empty cell into every direct row, and set cx to the final safe width sum. Delete matching grid/cell ranges and return relationship IDs under removed cells.

- [ ] **Step 3: Write failing merge-aware column tests**

Mirror row coverage for insertion before/inside/end/after horizontal and 2D regions, multiple regions, partial shrink, horizontal-to-vertical/1 x 1 degradation, anchor-column promotion, hidden survivor state, and canonical new continuation roles. A vertical-only one-column region cannot be expanded by strict-inside column insertion.

- [ ] **Step 4: Implement column-region projection**

Reuse the axis-independent region mapper. Insert correct tokens per row, update only geometrically changed survivors, validate final rectangles, and keep row start/end tags plus non-cell children byte-identical.

- [ ] **Step 5: Add patch-isolation, sum-overflow, and relationship tests**

Reject final grid-width overflow and invalid geometry before applying the patch plan. Prove no grid/cell/transform replacement ranges overlap, no failed internal edit reaches package state, and relationship-ID collection is exact across unique/shared links. Reserve package rollback injection for Task 4.

- [ ] **Step 6: Run gates, review, commit, and push**

    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/model/src/table-structure-edit.internal.test.ts \
      packages/model/src/table-cell-merge.internal.test.ts \
      packages/model/src/table-column-widths.internal.test.ts
    pnpm --config.verify-deps-before-run=false exec tsc \
      -p packages/model/tsconfig.json --noEmit --pretty false
    git diff --check
    git add packages/model/src/table-structure-edit.internal.ts \
      packages/model/src/table-structure-edit.internal.test.ts
    git commit -m "feat: splice table columns"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD

Review grid/cell parity, exact width math, merge-axis behavior, source isolation, ID collection, and rollback. Expected divergence: 0 0.

### Task 4: Public TableModel CRUD and Relationship Lifecycle

**Files:**
- Modify: packages/model/src/shapes.ts
- Modify: packages/model/src/model.test.ts

**Interfaces:**
- Produces InsertTableRowsOptions, InsertTableColumnsOptions, insertRows, deleteRows, insertColumns, and deleteColumns exactly as defined by the design.

- [ ] **Step 1: Add failing public lifecycle tests**

Create/import a table containing plain/rich text, URL/internal hyperlinks, mixed styles, opaque XML, explicit/auto row sizes, non-uniform widths, and multiple merges. Exercise all four methods at beginning/middle/end, edit new cells, and assert live rows/widths/heights/merges, stable TableModel identity, survivor state, six-format reopen, duplicate isolation, and unaffected siblings.

- [ ] **Step 2: Add failing relationship GC tests**

Delete unique/shared URL/internal hyperlink cells. Require unique final references removed, shared IDs retained until the final member is deleted, links outside the table preserved, and no orphan introduced. Inject removeRelationship failure and prove rollback.

- [ ] **Step 3: Implement public types and transaction wrappers**

Normalize inputs, resolve the frame, run the internal editor, serialize/set the slide part, parse the updated slide, and remove each returned relationship only when it exists and relationshipReferenceCount(updatedXml, id) equals zero. Keep the table wrapper live.

- [ ] **Step 4: Add exact error and isolation coverage**

Assert errors for invalid indexes/counts/options/dimensions, last-row/last-column deletion, unsafe structure/merge, overflow, and stale wrappers. Inject slide-part write, relationship removal, ZIP-date, and outer transaction failures; capture part bytes, relationship part, dates, journal, identities, and sizes before failures and require exact restoration.

- [ ] **Step 5: Run gates, review, commit, and push**

    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/model/src/table-structure-edit.internal.test.ts \
      packages/model/src/model.test.ts -t "table row|table column|table structure"
    pnpm --config.verify-deps-before-run=false exec tsc \
      -p packages/model/tsconfig.json --noEmit --pretty false
    git diff --check
    git add packages/model/src/shapes.ts packages/model/src/model.test.ts
    git commit -m "feat: edit table rows and columns"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD

Expected divergence: 0 0.

### Task 5: SDK/Root Contracts and PptxGenJS Import Editing

**Files:**
- Modify: packages/sdk/src/index.test.ts
- Modify: packages/pptx/src/index.test.ts
- Modify: packages/pptxgenjs-adapter/src/index.test.ts

- [ ] **Step 1: Add SDK/root runtime and type tests**

Through aggregate exports, create a merged linked table, insert/delete rows/columns, fill new cells, duplicate, write all six formats, reopen, and assert dimensions/topology/content/relationships. Add positive uses and ts-expect-error cases for wrong names/types, missing indexes, extra arguments, and non-numeric count.

- [ ] **Step 2: Add PptxGenJS import/edit tests**

Generate legal plain, rich, linked, sized, and merged tables only through public PptxGenJS APIs. Import, apply all four native editors, and compare surviving text/style/click/merge/grid/row state plus reopen. Record that PptxGenJS has no existing-deck CRUD API and keep malformed output preservation-only.

- [ ] **Step 3: Run gates and review**

    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts \
      packages/pptxgenjs-adapter/src/index.test.ts \
      -t "table row|table column|table structure"
    pnpm --config.verify-deps-before-run=false typecheck
    pnpm --config.verify-deps-before-run=false build
    git diff --check

Review exports, negative contracts, six formats, imported state, strict divergence, and dependency boundaries.

- [ ] **Step 4: Commit, push, and verify synchronization**

    git add packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts \
      packages/pptxgenjs-adapter/src/index.test.ts
    git commit -m "test: verify table structure editing contracts"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD

Expected divergence: 0 0.

### Task 6: Documentation and Compatibility Matrix

**Files:**
- Modify: README.md
- Modify: packages/pptx/README.md
- Modify: docs/api/README.md
- Modify: docs/compatibility/pptxgenjs-baseline.md
- Modify: docs/implementation-progress.md
- Modify: CHANGELOG.md

- [ ] **Step 1: Document exact API and structural behavior**

Add examples for physical insert/delete, default/explicit sizes, filling canonical cells, and merge expansion/promotion. Document last-item guards, safety limit, survivor preservation, relationship GC, automatic-row transform behavior, and PptxGenJS boundary.

- [ ] **Step 2: Update support and remaining matrices**

Move row/column CRUD to supported editing. Keep logical creation distinct. Set remaining order to auto-page/repeated headers, content measurement/layout recomputation, tableToSlides, then final audit. Recalculate progress without claiming 100% before proof.

- [ ] **Step 3: Run consistency review**

    rg -n "insertRows|deleteRows|insertColumns|deleteColumns|row/column CRUD|tableToSlides" \
      README.md packages/pptx/README.md docs/api/README.md \
      docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/model/src/table-structure-edit.internal.test.ts \
      packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
      packages/pptxgenjs-adapter/src/index.test.ts \
      -t "table row|table column|table structure"
    git diff --check

Review spelling, physical/logical distinction, merge semantics, completed scope, and absence of premature proof numbers.

- [ ] **Step 4: Commit, push, and verify synchronization**

    git add README.md packages/pptx/README.md docs/api/README.md \
      docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
    git commit -m "docs: document table row column CRUD"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD

Expected divergence: 0 0.

### Task 7: Actual Packed and Real-Browser Verification

**Files:**
- Modify: scripts/smoke-npm-package.mjs
- Modify: scripts/playwright-browser-smoke.js

- [ ] **Step 1: Extend actual-package declarations and Node smoke**

Assert packed options/methods and NodeNext positive/negative contracts. In installed Node, create/import a linked merged table, perform all four operations, edit new cells, verify relationship GC/survivors, reopen, and emit stable tableStructureEditing booleans.

- [ ] **Step 2: Extend CLI/Inspector proof**

Write a final evidence deck and run package inspect, PowerPoint 2010 validate, slides list, exact slide part, and relationship part reads. Require rectangular rows, expected sums/merge tokens/text, no orphan relationships, and no unexpected diagnostics.

- [ ] **Step 3: Extend real-Chrome lifecycle**

Run the same lifecycle from the packed browser export, download an evidence deck, and require every stage true plus zero document validation, console, page, and network errors.

- [ ] **Step 4: Build/pack twice and execute proof gates**

Compare sorted dist SHA-256 manifests and tarball bytes, then run:

    node --check scripts/smoke-npm-package.mjs
    node --check scripts/playwright-browser-smoke.js
    node scripts/smoke-npm-package.mjs \
      /tmp/pptx-table-structure-editing-pack/jiayunxie-pptx-0.1.0.tgz
    git diff --check

Serve the extracted tarball browser module over loopback, execute the checked-in callback in installed Chrome, and retain all JSON/PPTX evidence outside the repository.

- [ ] **Step 5: Review, commit, push, and verify synchronization**

    git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
    git commit -m "test: verify packed table structure editing"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD

Expected divergence: 0 0.

### Task 8: Full Gates and Recorded Proof

**Files:**
- Modify: README.md
- Modify: packages/pptx/README.md
- Modify: docs/api/README.md
- Modify: docs/compatibility/pptxgenjs-baseline.md
- Modify: docs/implementation-progress.md
- Modify: CHANGELOG.md

- [ ] **Step 1: Run final gates**

    pnpm --config.verify-deps-before-run=false typecheck
    pnpm --config.verify-deps-before-run=false build
    pnpm --config.verify-deps-before-run=false --filter @jiayunxie/pptx build
    pnpm --config.verify-deps-before-run=false exec vitest run \
      packages/model/src/table-structure-edit.internal.test.ts \
      packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
      packages/pptxgenjs-adapter/src/index.test.ts packages/pptx/src/index.test.ts
    pnpm --config.verify-deps-before-run=false exec vitest run --maxWorkers=2
    pnpm --config.verify-deps-before-run=false test:performance

Record exact counts/durations and rerun existing time-boundary flakes without weakening assertions.

- [ ] **Step 2: Revalidate exact evidence**

Run packed Node/browser smoke, real Chrome, PowerPoint 2010, package inspect, slides list, exact slide part, and exact relationship part. Require all states true, rectangular structure, expected merge promotion, no orphan relationships, and no unexpected diagnostics.

- [ ] **Step 3: Record actual values and recalculate progress**

Add exact gates, dist/tarball counts and hashes, installed surfaces, Chrome version/states, validation counts, evidence directory, structure/merge/relationship dimensions, commit IDs, new percentage, and remaining order.

- [ ] **Step 4: Final review, commit, push, and confirm synchronization**

    git diff --check
    git status --short
    git add README.md packages/pptx/README.md docs/api/README.md \
      docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
    git commit -m "docs: record table structure editing proof"
    git push origin main
    git rev-list --left-right --count origin/main...HEAD
    git status --short --branch

Expected divergence: 0 0; only the pre-existing untracked .pnpm-store/ may remain.
