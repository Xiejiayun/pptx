# Table Text Direction Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user requires inline execution without subagents.

**Goal:** Add strict table-level `textDirection` creation that materializes a four-value default onto uncovered physical cells with PptxGenJS 4.0.1 final-state parity.

**Architecture:** Extend the existing descriptor-safe table option normalizer with one optional `TableCellTextDirection`, resolve it onto normalized cells that lack a cell marker, and reuse the current creation renderer for direct `tcPr@vert`. Preserve explicit cell `horz` through precedence resolution, but let the renderer collapse it to no direct attribute; retain no table metadata.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span model, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice headless, Poppler.

## Global Constraints

- Public creation property is `AddTableOptions.textDirection?: TableCellTextDirection` with exact values `horz`, `vert`, `vert270`, and `wordArtVert`.
- A valid cell-level `AddTableCellOptions.textDirection` overrides the table value; omitted or runtime-undefined cell values inherit.
- Explicit cell `horz` blocks a non-horizontal table value but creates no direct attribute.
- Resolved non-horizontal state is materialized on each uncovered physical cell as direct `a:tcPr@vert`; never write `a:bodyPr@vert` or table metadata.
- Omitted/runtime-undefined/explicit-horizontal table values produce identical direction bytes for the same rows and cell overrides.
- Serialization remains margins, optional anchor, optional non-horizontal vert, L/R/T/B borders, then fill.
- Reuse `normalizeTableCellTextDirection()` and the existing cell renderer; do not accept aliases, coercion, case folding, whitespace trimming, or text-box-only direction tokens.
- Clearing a materialized cell after creation removes its direct state and does not reapply the original table value.
- Existing `setCellTextDirection(..., 'horz')` direct editing semantics remain unchanged and may explicitly write `vert="horz"`.
- No table-level getter/editor, fit creation, rich/multi-paragraph cells, merge, hyperlink, auto-page, or layout recomputation in this slice.
- Every successful task is reviewed, committed, pushed, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only gates do not create empty commits.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and materialize the table default

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing `normalizeTableCellTextDirection()`, normalized cell `textDirection`, and `renderTableCellTextDirectionAttribute()`.
- Produces: `NormalizedTableDefinition.rows` whose uncovered cells carry the resolved table value.

- [ ] **Step 1: Add failing precedence and exact-output tests**

Add adjacent to the cell-level direction test:

```ts
it('materializes strict table text direction onto uncovered cells', () => {
  const rows = [[
    'String',
    { text: 'Object' },
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { textDirection: undefined } },
    { text: 'Horizontal', options: { textDirection: 'horz' } },
    { text: 'Vertical', options: { textDirection: 'vert' } },
    { text: 'Rotate 270', options: { textDirection: 'vert270' } },
    { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
  ]];
  const definition = normalizeTableDefinition(rows, {
    textDirection: 'vert270',
  });
  expect(definition.rows[0]!.map(({ textDirection }) => textDirection)).toEqual([
    'vert270',
    'vert270',
    'vert270',
    'vert270',
    'horz',
    'vert',
    'vert270',
    'wordArtVert',
  ]);

  const xml = renderTableGraphicFrame(60, definition);
  expect([...xml.matchAll(/<a:tcPr([^>]*)>/g)].map((match) =>
    match[1]!.match(/\svert="([^"]+)"/)?.[1])).toEqual([
    'vert270',
    'vert270',
    'vert270',
    'vert270',
    undefined,
    'vert',
    'vert270',
    'wordArtVert',
  ]);
  expect(xml).not.toMatch(/<a:bodyPr[^>]*\svert=/);

  const omitted = renderTableGraphicFrame(
    61,
    normalizeTableDefinition([['Same']], {}),
  );
  const runtimeUndefined = renderTableGraphicFrame(
    61,
    normalizeTableDefinition([['Same']], { textDirection: undefined }),
  );
  const horizontal = renderTableGraphicFrame(
    61,
    normalizeTableDefinition([['Same']], { textDirection: 'horz' }),
  );
  expect(runtimeUndefined).toBe(omitted);
  expect(horizontal).toBe(omitted);
  expect(omitted).not.toMatch(/<a:tcPr[^>]*\svert=/);
});
```

Extend malformed table options with a getter-free accessor and the strict invalid set:

```ts
let tableTextDirectionAccessorCalls = 0;
const accessorTableTextDirectionOptions = {};
Object.defineProperty(accessorTableTextDirectionOptions, 'textDirection', {
  get() {
    tableTextDirectionAccessorCalls += 1;
    return 'vert';
  },
  enumerable: true,
  configurable: true,
});
const invalidTableTextDirections = [
  null,
  false,
  true,
  0,
  '',
  'Vert',
  ' vert ',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
  [],
  {},
  Symbol('vert'),
];
```

Add the accessor object and every `{ textDirection }` to `invalidOptions`; require `tableTextDirectionAccessorCalls` to remain zero.

- [ ] **Step 2: Run focused tests and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "table text direction|malformed matrices"
```

Expected: table `textDirection` is rejected as an unsupported option.

- [ ] **Step 3: Implement the minimal normalization overlay**

Add `textDirection` to `OPTION_KEYS`. Resolve it after margins and before the final return without changing the renderer:

```ts
const tableTextDirection = normalizedOptions.textDirection === undefined
  ? undefined
  : normalizeTableCellTextDirection(
    normalizedOptions.textDirection,
    'Table textDirection',
  );
const directionResolvedRows = tableTextDirection === undefined
  ? marginResolvedRows
  : marginResolvedRows.map((row) => row.map((cell) =>
    cell.textDirection === undefined
      ? { ...cell, textDirection: tableTextDirection }
      : cell));
```

Feed `directionResolvedRows` into table `valign` propagation and keep its output as the returned `rows`. Explicit `horz` is a defined marker and therefore remains untouched.

- [ ] **Step 4: Run focused regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Require exact precedence, omitted/horizontal byte equality, getter-free rejection, no body direction, and all existing table-default tests to pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Review only the new supported key, strict normalizer call, immutable overlay, explicit-horz marker, and existing renderer ownership/order. Stage only the two files and commit:

```text
feat: propagate table text direction during creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`; `.pnpm-store/` remains untracked.

---

### Task 2: Expose the public option and prove the model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1's normalization overlay.
- Produces: public `AddTableOptions.textDirection?: TableCellTextDirection` plus create/read/edit/write/reopen evidence.

- [ ] **Step 1: Add the typed option and lifecycle assertions first**

Add a model lifecycle test:

```ts
it('materializes table text direction through edit, duplicate, write, and reopen', async () => {
  const pkg = await OpcPackage.open(await modelFixture());
  const model = new PresentationModel(pkg);
  const slide = model.addSlide();
  const options: AddTableOptions = {
    name: 'Table text direction lifecycle',
    textDirection: 'vert270',
    columnWidths: inches(2),
    rowHeights: inches(1),
  };
  const table = slide.addTable([[
    'Inherited string',
    { text: 'Inherited object' },
    { text: 'Horizontal override', options: { textDirection: 'horz' } },
    { text: 'Vertical override', options: { textDirection: 'vert' } },
    { text: 'Stacked override', options: { textDirection: 'wordArtVert' } },
  ]], options);
  expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
    'vert270',
    'vert270',
    undefined,
    'vert',
    'wordArtVert',
  ]);

  const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
  const duplicateTable = duplicate.shapes.find(
    (shape): shape is TableModel => shape instanceof TableModel,
  );
  expect(duplicateTable).toBeInstanceOf(TableModel);

  table.setCellTextDirection(0, 0, undefined);
  table.setCellTextDirection(0, 1, 'wordArtVert');
  expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
    undefined,
    'wordArtVert',
    undefined,
    'vert',
    'wordArtVert',
  ]);
  expect(duplicateTable!.rows[0]!.cells.map(
    ({ textDirection }) => textDirection)).toEqual([
    'vert270',
    'vert270',
    undefined,
    'vert',
    'wordArtVert',
  ]);

  const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
  const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri);
  const reopenedTable = reopenedSlide?.shapes.find(
    (shape): shape is TableModel => shape instanceof TableModel
      && shape.name === 'Table text direction lifecycle',
  );
  expect(reopenedTable).toBeInstanceOf(TableModel);
  expect(reopenedTable!.rows[0]!.cells.map(
    ({ textDirection }) => textDirection)).toEqual([
    undefined,
    'wordArtVert',
    undefined,
    'vert',
    'wordArtVert',
  ]);
});
```

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: excess-property error for `AddTableOptions.textDirection`.

- [ ] **Step 3: Widen only the public table option**

In `packages/model/src/slide.ts` add:

```ts
readonly textDirection?: TableCellTextDirection;
```

Place it with the existing table style options. Do not add a `TableModel` property, definition-level metadata, or table editor.

- [ ] **Step 4: Run model gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only the two files and commit:

```text
feat: expose table text direction creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 3: Cover the SDK lifecycle and invalid public inputs

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: public `AddTableOptions.textDirection` and existing physical-cell editor.
- Produces: package-facing propagation, override, rollback, duplicate, and zero-mutation evidence.

- [ ] **Step 1: Add table-default coverage to the created-table lifecycle**

Create a table with `textDirection: 'vert270'`, two inherited cells, explicit `horz`, `vert`, and `wordArtVert` overrides. Require immediate snapshots:

```ts
[
  ['vert270', 'vert270', undefined],
  ['vert', 'wordArtVert', 'vert270'],
]
```

Capture the direct XML tokens and all unrelated text, alignment, margin, border, fill, column-width, row-height, and identity state. Duplicate before editing; clear one inherited cell, change another, and throw inside an outer transaction after a third edit. Require rollback to restore the pre-transaction final direct matrix and duplicate isolation.

- [ ] **Step 2: Prove write/reopen and non-reinheritance**

Write and reopen the source and duplicate. Require the cleared inherited cell to stay `undefined`, not regain `vert270`; require the explicit cell `horz` to remain no direct attribute; and require the duplicate to preserve the original materialized values.

- [ ] **Step 3: Expand invalid table-option coverage**

Extend the existing invalid table-cell direction test with table inputs:

```ts
expect(() => slide.addTable(
  [['Accessor table']],
  accessorOptions as AddTableOptions,
)).toThrow(TypeError);
for (const textDirection of invalidDirections) {
  expect(() => slide.addTable(
    [['Invalid table']],
    { textDirection } as unknown as AddTableOptions,
  )).toThrow(TypeError);
}
```

Require zero getter calls, unchanged part URI/byte map, unchanged mutation journal, stable slide/table identity, and unchanged existing direct directions.

- [ ] **Step 4: Run SDK and model suites**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only the SDK test and commit:

```text
test: cover table text direction creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 final-state parity

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: PptxGenJS 4.0.1 public runtime and native table creation.
- Produces: supported-value precedence equality and strict-difference evidence.

- [ ] **Step 1: Create equivalent native and PptxGenJS tables**

Use table `textDirection: 'vert270'` and cells:

```ts
[
  { text: 'Inherited', options: {} },
  { text: 'Horizontal', options: { textDirection: 'horz' } },
  { text: 'Vertical', options: { textDirection: 'vert' } },
  { text: 'Rotate 270', options: { textDirection: 'vert270' } },
  { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
]
```

Use equal geometry, column widths, row heights, margins, vertical alignment, borders, and fill in the supported native/PptxGenJS representation.

- [ ] **Step 2: Assert final direct state and invalid differences**

Require imported/native direct tokens and snapshots:

```ts
['vert270', undefined, 'vert', 'vert270', 'wordArtVert']
```

Require equal text, transform, grid, rows, margins, vertical alignment, borders, and fill; require no `bodyPr@vert`; write/reopen both and require the same state.

Add a PptxGenJS table-level runtime-invalid `textDirection: 'eaVert'` baseline. Require its XML to remain preserved and strict snapshot to be `undefined`; require native table creation with the same token to throw before mutation.

- [ ] **Step 3: Run adapter and model suites**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the adapter test and commit:

```text
test: compare table text direction creation with pptxgenjs
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 5: Smoke the actual Node/browser/types package

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: packed `@jiayunxie/pptx` Node/browser/declaration surfaces.
- Produces: `tableTextDirectionCreation: true` in smoke JSON without weakening `tableCellTextDirectionCreation`.

- [ ] **Step 1: Extend Node smoke**

Pass `textDirection: 'vert270'` in the main created table options. Remove the first cell direction so it inherits, use explicit `vert` on the second cell, retain explicit `wordArtVert` and `horz` on the remaining cells, and require immediate/direct/reopened state:

```js
[
  ['vert270', 'vert'],
  ['wordArtVert', undefined],
]
```

Add `tableTextDirectionCreation` through real assertions for materialization, cell override, explicit-horizontal blocking, direct token order, and reopen. Keep `tableCellTextDirectionCreation` true as the cell-level gate.

- [ ] **Step 2: Extend browser and declaration smoke**

Mirror Node behavior in the browser bundle fixture. In the declaration fixture add:

```ts
const tableOptions: AddTableOptions = {
  textDirection: cellDirection,
};
slide?.addTable([['Typed table direction']], tableOptions);
```

The existing `cellDirection: TableCellTextDirection = 'vert270'` proves exact type reuse.

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_direction_package_dir=$(mktemp -d /tmp/pptx-table-direction-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_direction_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_direction_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require every smoke flag true, including `tableTextDirectionCreation`, `tableCellTextDirectionCreation`, `types`, and CLI `0.1.0`.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the smoke script and commit:

```text
test: smoke packed table text direction creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 6: Document the public contract and remaining boundary

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: completed runtime/package evidence.
- Produces: accurate table-level direction creation and unsupported documentation.

- [ ] **Step 1: Update examples and contract**

Add `textDirection: 'vert270'` to a table option example while keeping cell overrides. Document:

- exact four-value table input;
- omitted/undefined cell inheritance and valid cell precedence;
- explicit cell `horz` blocks a non-horizontal table default but emits no direct token;
- direct per-cell materialization and no retained table metadata;
- clear does not re-inherit and the editor may explicitly write horizontal;
- omitted/undefined/horizontal table values preserve equivalent bytes;
- margins → anchor → vert → borders/fill serialization order.

- [ ] **Step 2: Update compatibility and unsupported lists**

Mark table-level `textDirection` creation supported and PptxGenJS-final-state compatible. Remove only that creation gap. Keep table-level direction getter/editor, fit creation, merge, hyperlink, rich cells, advanced borders/fills, auto-page, repeated headers, row insertion/deletion, and layout recomputation unsupported.

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'table-level (direction|`textDirection`).*(default.*unsupported|尚未支持)|table-level direction default' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
```

Review hits in context: creation support is valid, while table-level read/edit must remain pending.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the four documents and commit:

```text
docs: document table text direction creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 7: Run full gates and real-deck QA

**Files:**
- Review every Task 1–6 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: implementation, tests, docs, actual tarball, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces: a verified set of pushed commits; any defect gets its own focused fix/review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-direction-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-direction-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_direction_qa_package_dir=$(mktemp -d /tmp/pptx-table-direction-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_direction_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_direction_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
command -v pptx-inspect
pptx-inspect --json doctor
```

Require every smoke flag true, offline doctor success, and CLI `0.1.0`.

- [ ] **Step 3: Generate real native and PptxGenJS decks**

Create `/tmp/pptx-table-direction-qa.mjs` with `apply_patch`:

```js
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import {
  PptxDocument,
  TableModel,
  inches,
} from 'file:///Users/jeremy/workspace/pptx/packages/pptx/dist/index.js';

const require = createRequire(
  '/Users/jeremy/workspace/pptx/packages/pptxgenjs-adapter/package.json',
);
const PptxGenJS = require('pptxgenjs');
const output = '/tmp/pptx-table-direction-create-qa';
await mkdir(output, { recursive: true });

const baseRows = [
  ['Inherited A', 'Inherited B', 'Inherited C', 'Inherited D'],
  ['Inherited E', 'Inherited F', 'Inherited G', 'Inherited H'],
];
const tableOptions = {
  name: 'Table direction creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(2), inches(2), inches(2), inches(2)],
  rowHeights: [inches(1.5), inches(1.5)],
  margin: 7.2,
  valign: 'middle',
};

function build(rows, options = {}) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const table = document.addSlide().addTable(rows, { ...tableOptions, ...options });
  return { document, table };
}

const omitted = build(baseRows);
const runtimeUndefined = build(baseRows, { textDirection: undefined });
const horizontal = build(baseRows, { textDirection: 'horz' });
await writeFile(output + '/omitted.pptx', await omitted.document.write());
await writeFile(output + '/undefined.pptx', await runtimeUndefined.document.write());
await writeFile(output + '/horizontal.pptx', await horizontal.document.write());

for (const textDirection of ['vert', 'vert270', 'wordArtVert']) {
  const built = build(baseRows, { textDirection });
  assert.deepEqual(
    built.table.rows.map(({ cells }) => cells.map((cell) => cell.textDirection)),
    Array.from({ length: 2 }, () => Array(4).fill(textDirection)),
  );
  await writeFile(
    output + `/table-${textDirection}.pptx`,
    await built.document.write(),
  );
}

const mixedRows = [
  [
    'Inherited string',
    { text: 'Horizontal override', options: { textDirection: 'horz' } },
    { text: 'Vertical override', options: {
      textDirection: 'vert',
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    } },
    { text: 'Stacked override', options: {
      textDirection: 'wordArtVert',
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
    } },
  ],
  [
    { text: 'Inherited object', options: { margin: { top: 4, left: 18 } } },
    { text: 'Rotate override', options: { textDirection: 'vert270' } },
    { text: 'Inherited fill', options: {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
    } },
    { text: 'Horizontal blocks', options: { textDirection: 'horz' } },
  ],
];
const source = build(mixedRows, { textDirection: 'vert270' });
const sourceDirections = source.table.rows.map(({ cells }) =>
  cells.map(({ textDirection }) => textDirection));
assert.deepEqual(sourceDirections, [
  ['vert270', undefined, 'vert', 'wordArtVert'],
  ['vert270', 'vert270', 'vert270', undefined],
]);
await writeFile(output + '/mixed-source.pptx', await source.document.write());

source.table.setCellTextDirection(0, 0, undefined);
source.table.setCellTextDirection(0, 1, 'wordArtVert');
source.table.setCellText(0, 1, 'Edited to stacked');
const editedBytes = await source.document.write();
await writeFile(output + '/mixed-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[0].cells[0].textDirection, undefined);
assert.equal(reopenedTable.rows[0].cells[1].textDirection, 'wordArtVert');
assert.equal(reopenedTable.rows[0].cells[1].text, 'Edited to stacked');
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/mixed-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      { text: 'Inherited string', options: {} },
      { text: 'Horizontal override', options: { textDirection: 'horz' } },
      { text: 'Vertical override', options: {
        textDirection: 'vert',
        fill: { color: 'D9EAF7' },
      } },
      { text: 'Stacked override', options: {
        textDirection: 'wordArtVert',
        border: { type: 'solid', color: 'C00000', pt: 2 },
      } },
    ],
    [
      { text: 'Inherited object', options: { margin: [4, 7.2, 3.6, 18] } },
      { text: 'Rotate override', options: { textDirection: 'vert270' } },
      { text: 'Inherited fill', options: { fill: { color: 'FFF2CC' } } },
      { text: 'Horizontal blocks', options: { textDirection: 'horz' } },
    ],
  ],
  {
    x: 0.75,
    y: 0.75,
    w: 8,
    h: 3,
    colW: [2, 2, 2, 2],
    rowH: [1.5, 1.5],
    margin: 0.1,
    valign: 'middle',
    textDirection: 'vert270',
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceDirections,
  reopenedDirections: reopenedTable.rows.map(({ cells }) =>
    cells.map(({ textDirection }) => textDirection)),
  editedText: reopenedTable.rows[0].cells[1].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run:

```sh
node /tmp/pptx-table-direction-qa.mjs
```

Require source directions `[["vert270",null,"vert","wordArtVert"],["vert270","vert270","vert270",null]]`, reopened first row `[null,"wordArtVert","vert","wordArtVert"]`, widths `[1828800,1828800,1828800,1828800]`, and heights `[1371600,1371600]`.

- [ ] **Step 4: Validate and diff packages**

```sh
for deck in /tmp/pptx-table-direction-create-qa/*.pptx; do
  pptx-inspect --json package validate "$deck" --profile powerpoint-2010
done
pptx-inspect --json package diff \
  /tmp/pptx-table-direction-create-qa/omitted.pptx \
  /tmp/pptx-table-direction-create-qa/undefined.pptx
pptx-inspect --json package diff \
  /tmp/pptx-table-direction-create-qa/omitted.pptx \
  /tmp/pptx-table-direction-create-qa/horizontal.pptx
pptx-inspect --json package diff \
  /tmp/pptx-table-direction-create-qa/mixed-source.pptx \
  /tmp/pptx-table-direction-create-qa/mixed-edited.pptx
pptx-inspect --json package diff \
  /tmp/pptx-table-direction-create-qa/mixed-edited.pptx \
  /tmp/pptx-table-direction-create-qa/mixed-reopened.pptx
```

Require zero errors/warnings. Omitted/undefined/horizontal diffs must have zero changed parts. Source→edited must change only `/ppt/slides/slide1.xml`; edited→reopened must have zero changed parts.

- [ ] **Step 5: Render, inspect, and check overflow**

```sh
table_direction_render_dir=$(mktemp -d /tmp/pptx-table-direction-rendered.XXXXXX)
mkdir -p "$table_direction_render_dir/source" \
  "$table_direction_render_dir/edited" \
  "$table_direction_render_dir/baseline"
table_direction_source_profile=$(mktemp -d /tmp/pptx-table-direction-source-profile.XXXXXX)
table_direction_edited_profile=$(mktemp -d /tmp/pptx-table-direction-edited-profile.XXXXXX)
table_direction_baseline_profile=$(mktemp -d /tmp/pptx-table-direction-baseline-profile.XXXXXX)
soffice -env:UserInstallation="file://$table_direction_source_profile" \
  --headless --convert-to pdf --outdir "$table_direction_render_dir/source" \
  /tmp/pptx-table-direction-create-qa/mixed-source.pptx
soffice -env:UserInstallation="file://$table_direction_edited_profile" \
  --headless --convert-to pdf --outdir "$table_direction_render_dir/edited" \
  /tmp/pptx-table-direction-create-qa/mixed-edited.pptx
soffice -env:UserInstallation="file://$table_direction_baseline_profile" \
  --headless --convert-to pdf --outdir "$table_direction_render_dir/baseline" \
  /tmp/pptx-table-direction-create-qa/pptxgenjs-baseline.pptx
pdftoppm -png -r 180 "$table_direction_render_dir/source/mixed-source.pdf" \
  "$table_direction_render_dir/source/slide"
pdftoppm -png -r 180 "$table_direction_render_dir/edited/mixed-edited.pdf" \
  "$table_direction_render_dir/edited/slide"
pdftoppm -png -r 180 "$table_direction_render_dir/baseline/pptxgenjs-baseline.pdf" \
  "$table_direction_render_dir/baseline/slide"
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-direction-create-qa/mixed-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-direction-create-qa/pptxgenjs-baseline.pptx
```

Inspect every PNG at original detail. Require visible vertical/rotated/stacked positions, explicit-horizontal blocking, preserved margins/borders/fills, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content. LibreOffice may display `wordArtVert` horizontally; accept it only when the native and PptxGenJS files behave identically and the OOXML tokens are exact.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict normalization, descriptor safety, precedence, direct ownership, ordering, omission/horizontal equivalence, declarations, lifecycle isolation, PptxGenJS differences, smoke flags, docs boundaries, and invalid-input no-mutation assertions.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `?? .pnpm-store/`. QA-only success creates no empty commit.
