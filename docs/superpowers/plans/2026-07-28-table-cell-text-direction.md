# Table Cell Text Direction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, lossless table-cell text-direction snapshots and indexed editing aligned with PptxGenJS 4.0.1's four-value table API.

**Architecture:** Introduce a table-cell-specific direction codec for direct `tcPr@vert`, then expose its read state through existing immutable `TableCell` snapshots and mutate it through `TableModel.setCellTextDirection()`. Keep the four-value table surface separate from seven-value text-box direction, and patch only one cell's direct property inside an OPC transaction.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 real-output fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public values are exactly `horz | vert | vert270 | wordArtVert` in a dedicated `TableCellTextDirection` type.
- Public read surface is optional `TableCell.textDirection`; public write surface is `TableModel.setCellTextDirection(rowIndex, columnIndex, value)`.
- Storage is direct `a:tc/a:tcPr@vert`, not cell or shape `bodyPr@vert`.
- Getter reads one unique direct `tcPr` and one unique unqualified direct `vert`; it does not resolve defaults or inheritance.
- Explicit `horz` writes `vert="horz"`; `undefined` removes the direct attribute.
- PptxGenJS omitted/explicit-horz output imports as `undefined` because both omit the wire attribute.
- Table-level PptxGenJS direction is consumed only through the per-cell attributes its public writer materializes.
- Do not add table creation, table-level defaults, cell fit, vertical alignment, margins, fill, border, merge mutation, rich text, or row/column mutation.
- Preserve `tcPr` formatting/metadata, `txBody`, merge state, neighboring cells, table relationships, and `TableModel` identity.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, empty native/baseline package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Public snapshot, strict codec, and model-level mutation

**Files:**
- Create: `packages/model/src/table-cell-text-direction.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `TableModel.rows`, `SlideModel.setXml()`, and nested OPC transactions.
- Produces: `TableCellTextDirection`, `TableCell.textDirection`, `normalizeTableCellTextDirection()`, `readTableCellTextDirection()`, `replaceTableCellTextDirection()`, and `TableModel.setCellTextDirection()`.

- [ ] **Step 1: Add failing strict-read tests**

Replace the fixture table with one physical row containing cells for direct `horz`, `vert`, `vert270`, and `wordArtVert`, followed by absent, empty, case variant, whitespace, namespaced `x:vert`, duplicate direct vert, unknown `eaVert`, descendant `bodyPr@vert`, repeated `tcPr`, and missing `tcPr`. Keep an unrelated valid text property on the last malformed cells. Assert:

```ts
const directions = table.rows[0]!.cells.map(({ textDirection }) => textDirection);
expect(directions.slice(0, 4)).toEqual(['horz', 'vert', 'vert270', 'wordArtVert']);
expect(directions.slice(4)).toEqual(Array(10).fill(undefined));
expect(pkg.mutations).toEqual(journal);
```

Mutate a returned snapshot through a test-only cast and prove a fresh `table.rows` read still returns the source direction.

- [ ] **Step 2: Run the model suite and verify red**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new snapshot assertions fail because `TableCell.textDirection` does not exist.

- [ ] **Step 3: Add the dedicated public type and strict reader**

In `shapes.ts`, add:

```ts
export type TableCellTextDirection = 'horz' | 'vert' | 'vert270' | 'wordArtVert';

export interface TableCell {
  readonly text: string;
  readonly textDirection?: TableCellTextDirection;
}
```

In the new internal codec, use an exact token set and direct-child helper:

```ts
const DIRECTIONS = new Set<TableCellTextDirection>([
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
]);

export function readTableCellTextDirection(
  xml: LosslessXmlDocument,
  cell: XmlElement,
): TableCellTextDirection | undefined {
  const properties = directChildren(cell, 'tcPr');
  if (properties.length !== 1) return undefined;
  const attributes = properties[0]!.attributes.filter(({ name }) => name === 'vert');
  if (attributes.length !== 1) return undefined;
  const value = attributes[0]!.value;
  return DIRECTIONS.has(value as TableCellTextDirection)
    ? value as TableCellTextDirection
    : undefined;
}
```

In `TableModel.rows`, add `textDirection` only when this reader returns a value. Do not read `bodyPr`, descendants, or effective defaults.

- [ ] **Step 4: Add failing lossless-edit and rollback tests**

Build cells with unknown single-quoted `vert`, expanded `tcPr` containing margins/fill/unknown XML, a merged placeholder with self-closing `tcPr`, an adjacent valid direction, repeated/missing `tcPr`, and unchanged text bodies. Cover:

```ts
table.setCellTextDirection(0, 0, 'vert270');
table.setCellTextDirection(0, 1, 'wordArtVert');
table.setCellTextDirection(0, 2, 'horz');
table.setCellTextDirection(0, 0, undefined);
```

Assert only targeted attributes change, existing quote style survives replacement, add/clear preserve cell-property children, the adjacent cell is byte-identical, `setCellText()` and transform edits preserve direction, same-value assignment produces no new mutation, invalid indices/structures change no bytes, and an outer transaction rollback restores bytes/journal/table identity and all direction snapshots.

- [ ] **Step 5: Implement validation and one-attribute patching**

Add exact validation:

```ts
export function normalizeTableCellTextDirection(
  value: unknown,
  context: string,
): TableCellTextDirection {
  if (typeof value !== 'string' || !DIRECTIONS.has(value as TableCellTextDirection)) {
    throw new TypeError(`${context} must be horz, vert, vert270, or wordArtVert`);
  }
  return value as TableCellTextDirection;
}
```

Implement `replaceTableCellTextDirection(xml, cell, value, partUri): boolean` by parsing only `xml.original(tcPr)`. Require exactly one direct `tcPr` and at most one unqualified `vert`; throw `ModelParseError` before source mutation when ambiguous. Return false for same-value or absent-clear no-ops. Otherwise preserve the existing attribute's quote style through `replaceAttribute()`, insert a new escaped attribute immediately before `/>` or `>`, or remove the attribute plus its leading horizontal whitespace, then replace only the original `tcPr` span and return true.

Expose the indexed method:

```ts
setCellTextDirection(
  rowIndex: number,
  columnIndex: number,
  value: TableCellTextDirection | undefined,
): void {
  const direction = value === undefined
    ? undefined
    : normalizeTableCellTextDirection(value, 'Table cell text direction');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    if (replaceTableCellTextDirection(xml, cell, direction, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

- [ ] **Step 6: Re-run the model suite**

Expected: strict reads, lossless edits, no-ops, invalid structure, and rollback tests pass without changing unrelated snapshots.

### Task 2: SDK lifecycle, duplication, reopen, and invalid-input isolation

**Files:**
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public type/method, `PptxDocument.open()`, duplicate, write/reopen, transactions, and `TableModel.setCellText()`.
- Produces: public lifecycle and package-isolation evidence without adding table creation APIs.

- [ ] **Step 1: Add a minimal editable table fixture**

Create a dedicated JSZip fixture with one table, two physical rows, normal and merged cells, self-closing/expanded `tcPr`, direct directions, text bodies, and unknown metadata. Keep it independent from `titleFixture()` so existing shape-index tests remain unchanged.

- [ ] **Step 2: Cover public edit lifecycle**

Open the fixture, retain slide/table identity, duplicate before editing, and set the four public values across different cells. Clear one direct value, edit text in another cell, mutate a detached rows snapshot, and edit the table transform. Assert directions, text, neighboring cells, merged placeholder, unknown XML, and source duplicate remain isolated.

Roll back a direction edit inside `document.transaction()` and assert exact bytes, mutation journal, slide/table identity, and live rows are restored. Write/reopen and compare edited versus duplicated directions.

- [ ] **Step 3: Cover invalid input and coordinates before mutation**

For `null`, booleans, numbers, empty/case/whitespace strings, `eaVert`, `mongolianVert`, `wordArtVertRtl`, objects, arrays, and symbols, assert `TypeError`. For negative, fractional, NaN, infinite, or out-of-range row/column locations, assert `RangeError`. Confirm bytes, journal, shape count, table identity, text, and directions remain unchanged.

- [ ] **Step 4: Run typecheck and focused suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck and both suites pass.

### Task 3: PptxGenJS conformance, compatibility docs, and packed surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public table-cell surface and PptxGenJS public `addTable()` output.
- Produces: 4.0.1 table/default conformance, split compatibility status, and packed Node/browser/declaration evidence.

- [ ] **Step 1: Add real PptxGenJS 4.0.1 conformance**

Extend the test-only PptxGenJS interface with public `addTable()`. Generate a table with outer `textDirection: 'vert270'` and cells for inherited, explicit horz, vert, vert270, wordArtVert, and runtime-passthrough `eaVert`; generate a second table with omitted and explicit horz under no default. Import through `importPptxGenJS()` and assert:

```ts
expect(first.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
  'vert270',
  undefined,
  'vert',
  'vert270',
  'wordArtVert',
  undefined,
]);
expect(second.rows[0]!.cells.map(({ textDirection }) => textDirection))
  .toEqual([undefined, undefined]);
```

Assert raw `tcPr@vert` tokens, unchanged cell text, write/reopen values, and no private-field access.

- [ ] **Step 2: Update API, compatibility, and changelog**

Add a `TableModel` example for snapshot reads, four-value edits, explicit horz, and clear. Document direct `tcPr` scope, physical indices, strict getter, PptxGen table-default materialization, omitted/horz wire collapse, invalid runtime passthrough, and separation from seven-value text-box direction.

Split the remaining compatibility row into:

```md
| table-cell `textDirection` | `TableCell.textDirection` / `TableModel.setCellTextDirection()` | 已支持 |
| table-cell fit | 尚无完整公开 API | 部分支持，后续逐项补齐 |
```

Add one precise changelog bullet and npm README capability note. Do not claim table creation or table-level defaults.

- [ ] **Step 3: Extend actual-package Node/browser/declaration smoke**

Inject a minimal valid table graphic frame into an existing created slide through the public package part, then read/edit/clear cell direction through the packed `TableModel` in Node and browser conditions. Verify explicit horz, vert270, wordArtVert, omission, cell text preservation, and adjacent-cell isolation.

Compile this declaration fixture:

```ts
const cellDirection: TableCellTextDirection = 'vert270';
const table = createdDocument.slides[0].shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
const snapshotDirection: TableCellTextDirection | undefined =
  table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, cellDirection);
table?.setCellTextDirection(0, 0, undefined);
```

- [ ] **Step 4: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: all functional tests pass, only default performance is skipped in the full run, and isolated performance passes.

### Task 4: Tarball, native/baseline validation, visual QA, and delivery

**Files:**
- Review all Task 1–3 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and all prior evidence.
- Produces: reviewed `feat: support table cell text direction` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_direction_package_dir=$(mktemp -d /tmp/pptx-table-cell-direction-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_direction_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_cell_direction_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke all report true.

- [ ] **Step 2: Generate same-source native and hand-patched table files**

Use public PptxGenJS 4.0.1 to create one wide table with large horizontal cells and no direction attributes. Open one copy through `PptxDocument` and apply `horz`, `vert`, `vert270`, and `wordArtVert` with `setCellTextDirection()`; patch the same four direct `tcPr@vert` attributes by hand in another copy. Save explicit outputs under `/tmp/pptx-table-cell-direction-native/native.pptx` and `/tmp/pptx-table-cell-direction-baseline/baseline.pptx`.

Validate and compare:

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-direction-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-direction-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-cell-direction-native/native.pptx /tmp/pptx-table-cell-direction-baseline/baseline.pptx
```

Expected: both packages report zero errors/warnings and added/removed/changed diff arrays are empty.

- [ ] **Step 3: Render, inspect, and run overflow checks**

Export both through `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf`, rasterize every page with `pdftoppm -png`, and inspect each image at full size. Confirm horizontal, 90-degree, 270-degree, and stacked text match labels and PptxGenJS behavior; native/baseline renders are pixel-identical; no repair, clipping, overlap, or unexpected neighboring-cell change occurs.

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-direction-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-direction-baseline/baseline.pptx
```

Expected: both overflow checks pass.

- [ ] **Step 4: Final review, commit, push, and verify remote state**

Run `git diff --check`, inspect the complete diff, and verify the four-token value boundary, direct `tcPr` ownership, explicit horz/clear behavior, strict reads, physical indices, cell/neighbor isolation, no table-creation/fit expansion, and all validation evidence. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, run `git diff --cached --check`, commit, and push through the verified SSH-over-443 channel:

```sh
git commit -m "feat: support table cell text direction"
git rev-list --left-right --count origin/main...HEAD
```

Expected: remote count is `0 0`.
