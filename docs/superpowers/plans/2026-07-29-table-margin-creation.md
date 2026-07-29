# Table Margin Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to execute this plan inline task-by-task. The user requires inline execution without subagents. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict point-based table-level `margin` creation that resolves canonical → table → cell sides onto physical cells with PptxGenJS 4.0.1 final-state parity.

**Architecture:** Extend the descriptor-safe table option normalizer with one optional `TextBoxMarginInput`, merge normalized table sides under normalized cell sides, and reuse the current direct cell margin renderer. Keep only resolved physical-cell state; do not add table metadata, a table getter, or a table editor.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler.

## Global Constraints

- Public creation property is `AddTableOptions.margin?: TextBoxMarginInput`; native units remain points.
- Resolution is per side: canonical top/bottom `3.6pt` and left/right `7.2pt`, then table sides, then cell sides.
- Cell scalar/TRBL replaces all table sides; cell partial named input overrides only supplied sides and inherits the rest.
- Omitted/runtime-undefined/empty cell margin inherits a table value; omitted/runtime-undefined/empty table margin preserves existing bytes.
- Resolved state is materialized only as direct `a:tcPr@marL/marR/marT/marB`; never write body insets or table metadata.
- Creation after resolution does not retain the source layer; clearing a cell later does not reapply the table value.
- Serialization remains margins, optional anchor, L/R/T/B borders, then fill.
- Reuse `normalizeTextBoxMargins()` and `renderTableCellMarginAttributes()`; do not copy PptxGenJS's first-value `<1` inches / `>=1` points rule.
- No `TableModel.margins`, table-level margin editor, horizontal alignment, direction/fit creation, table-level border/fill, merge, hyperlink, rich text, auto-page, or layout recomputation enters this slice.
- Every successful small item is reviewed, committed, pushed, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only success creates no empty commit.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and materialize table margin sides

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: `normalizeTextBoxMargins()`, normalized cell `margins`, and `renderTableCellMarginAttributes()`.
- Produces: `NormalizedTableDefinition.rows` whose cell margin sides contain the table/cell overlay.

- [ ] **Step 1: Add failing precedence and exact-output tests**

Add a focused test after cell-level margin creation coverage:

```ts
it('materializes strict table margins under cell margin sides', () => {
  const definition = normalizeTableDefinition([[
    'String',
    { text: 'Object' },
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { margin: undefined } },
    { text: 'Partial', options: { margin: { bottom: 12 } } },
    { text: 'Zero', options: { margin: 0 } },
    { text: 'Tuple', options: { margin: [1, 2, 3, 4] } },
  ]], { margin: { top: 9, left: 18 }, valign: 'middle' });

  expect(definition.rows[0]!.map(({ margins }) => margins)).toEqual([
    { top: 9, left: 18 },
    { top: 9, left: 18 },
    { top: 9, left: 18 },
    { top: 9, left: 18 },
    { top: 9, bottom: 12, left: 18 },
    { top: 0, right: 0, bottom: 0, left: 0 },
    { top: 1, right: 2, bottom: 3, left: 4 },
  ]);

  const xml = renderTableGraphicFrame(41, definition);
  const margins = [...xml.matchAll(
    /<a:tcPr marL="(-?\d+)" marR="(-?\d+)" marT="(-?\d+)" marB="(-?\d+)"/g,
  )].map((match) => match.slice(1).map(Number));
  expect(margins).toEqual([
    [228600, 91440, 114300, 45720],
    [228600, 91440, 114300, 45720],
    [228600, 91440, 114300, 45720],
    [228600, 91440, 114300, 45720],
    [228600, 91440, 114300, 152400],
    [0, 0, 0, 0],
    [50800, 25400, 12700, 38100],
  ]);
  expect(xml).toMatch(/marB="45720" anchor="ctr"><a:lnL/);
});
```

Add omission equality:

```ts
const omitted = renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], {}),
);
const runtimeUndefined = renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], { margin: undefined }),
);
const empty = renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], { margin: {} }),
);
expect(runtimeUndefined).toBe(omitted);
expect(empty).toBe(omitted);
```

Extend malformed table options with:

```ts
let tableMarginAccessorCalls = 0;
const accessorTableMarginOptions = {};
Object.defineProperty(accessorTableMarginOptions, 'margin', {
  get() {
    tableMarginAccessorCalls += 1;
    return 1;
  },
  enumerable: true,
});
const invalidTableMargins: unknown[] = [
  null,
  false,
  '1',
  Number.NaN,
  Number.POSITIVE_INFINITY,
  Number.NEGATIVE_INFINITY,
  [1, 2, 3],
  [1, 2, 3, 4, 5],
  Object.create({ top: 1 }),
  { top: 1, extra: 2 },
  { top: 1, [Symbol('margin')]: 2 },
];
expect(() => normalizeTableDefinition(
  [['Accessor']],
  accessorTableMarginOptions,
)).toThrow();
for (const margin of invalidTableMargins) {
  expect(() => normalizeTableDefinition([['Invalid']], { margin })).toThrow();
}
expect(tableMarginAccessorCalls).toBe(0);
```

Reuse the existing accessor tuple/object/class/sparse cases for deeper shapes.

- [ ] **Step 2: Run focused tests and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "table margins|malformed matrices"
```

Expected: table `margin` is rejected as an unsupported option.

- [ ] **Step 3: Implement the minimal normalization overlay**

Add `margin` to `OPTION_KEYS`. Immediately after `readOptions(options)`, normalize it:

```ts
const tableMargins = normalizeTextBoxMargins(
  normalizedOptions.margin as TextBoxMarginInput | undefined,
  'Table margin',
);
const marginResolvedRows = tableMargins === undefined
  ? normalizedRows
  : normalizedRows.map((row) => row.map((cell) => ({
    ...cell,
    margins: { ...tableMargins, ...(cell.margins ?? {}) },
  })));
```

Apply the existing table `valign` overlay to `marginResolvedRows`, not `normalizedRows`. Do not add a definition-level margin field or change `renderTableCell()`.

- [ ] **Step 4: Run focused regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Require exact per-side precedence, empty/omitted byte equality, getter-free rejection, canonical output order, and all existing margin/valign tests to pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only the two files and commit:

```text
feat: propagate table margin during creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`; `.pnpm-store/` remains untracked.

---

### Task 2: Expose the public option and prove the model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1's per-side overlay.
- Produces: public `AddTableOptions.margin?: TextBoxMarginInput` and create/read/edit/write/reopen evidence.

- [ ] **Step 1: Add the typed option and lifecycle assertions first**

Add a model test using:

```ts
const options: AddTableOptions = {
  name: 'Table margin lifecycle',
  margin: { top: 9, left: 18 },
  columnWidths: inches(2),
  rowHeights: inches(1),
};
const table = slide.addTable([[
  'Inherited string',
  { text: 'Partial override', options: { margin: { bottom: 12 } } },
  { text: 'Zero override', options: { margin: 0 } },
  { text: 'Tuple override', options: { margin: [1, 2, 3, 4] } },
]], options);
expect(table.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
  { top: 9, right: 7.2, bottom: 3.6, left: 18 },
  { top: 9, right: 7.2, bottom: 12, left: 18 },
  { top: 0, right: 0, bottom: 0, left: 0 },
  { top: 1, right: 2, bottom: 3, left: 4 },
]);
```

Duplicate the slide, then clear the first source cell and replace the second with `{ right: 5 }`. Require the source to expose `[undefined, { right: 5 }, ...]`, the duplicate to retain all original materialized values, and write/reopen to preserve both matrices.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: excess-property error for `AddTableOptions.margin`.

- [ ] **Step 3: Widen only the public table option**

In `packages/model/src/slide.ts` add:

```ts
readonly margin?: TextBoxMarginInput;
```

Do not add a `TableModel` property or editor.

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
feat: expose table margin creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 3: Cover the SDK lifecycle and invalid public inputs

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: public `AddTableOptions.margin` and existing `TableModel` cell editor.
- Produces: package-facing propagation, partial override, duplicate, rollback, reopen, and no-mutation evidence.

- [ ] **Step 1: Add a focused public lifecycle test**

Add:

```ts
it('materializes public table margins through duplicate, rollback, and reopen', async () => {
  const document = PptxDocument.create();
  const slide = document.addSlide();
  const table = slide.addTable([[
    'Inherited string',
    { text: 'Partial override', options: { margin: { bottom: 12 } } },
    { text: 'Zero override', options: { margin: 0 } },
    { text: 'Tuple override', options: { margin: [1, 2, 3, 4] } },
  ]], {
    name: 'SDK table margin lifecycle',
    margin: { top: 9, left: 18 },
    columnWidths: inches(2),
    rowHeights: inches(1),
  });
  const original = table.rows[0]!.cells.map(({ margins }) => margins);
  expect(original).toEqual([
    { top: 9, right: 7.2, bottom: 3.6, left: 18 },
    { top: 9, right: 7.2, bottom: 12, left: 18 },
    { top: 0, right: 0, bottom: 0, left: 0 },
    { top: 1, right: 2, bottom: 3, left: 4 },
  ]);

  const duplicate = document.duplicateSlide(0);
  const duplicateTable = duplicate.shapes[0] as TableModel;
  expect(duplicateTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(original);

  table.setCellMargins(0, 0, undefined);
  table.setCellMargins(0, 1, { right: 5 });
  const edited = table.rows[0]!.cells.map(({ margins }) => margins);
  expect(edited).toEqual([
    undefined,
    { right: 5 },
    { top: 0, right: 0, bottom: 0, left: 0 },
    { top: 1, right: 2, bottom: 3, left: 4 },
  ]);

  const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
  const rollbackJournal = [...document.opcPackage.mutations];
  expect(() => document.transaction(() => {
    table.setCellMargins(0, 2, { left: 11 });
    slide.addTable([['Temporary']], { margin: 6 });
    throw new Error('restore table margin defaults');
  })).toThrow('restore table margin defaults');
  expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
  expect(document.opcPackage.mutations).toEqual(rollbackJournal);
  expect(slide.shapes).toHaveLength(1);
  expect(slide.shapes[0]).toBe(table);
  expect(table.rows[0]!.cells.map(({ margins }) => margins)).toEqual(edited);
  expect(duplicateTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(original);

  const reopened = await PptxDocument.open(await document.write());
  const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
  const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
  expect(reopenedTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(edited);
  expect(reopenedDuplicate.rows[0]!.cells.map(({ margins }) => margins)).toEqual(original);
});
```

- [ ] **Step 2: Extend reopen and isolation assertions**

Write/reopen and require:

```ts
expect(reopenedTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
  undefined,
  { right: 5 },
  { top: 0, right: 0, bottom: 0, left: 0 },
  { top: 1, right: 2, bottom: 3, left: 4 },
]);
```

Require the duplicate to retain `original`, the source to retain `edited`, and both tables to keep their exact width/height vectors after reopen.

- [ ] **Step 3: Expand invalid table-option coverage**

In the existing invalid table margin test, add an accessor-backed `AddTableOptions` and run every `invalidMargins` value through:

```ts
expect(() => slide.addTable(
  [['Accessor table']],
  accessorOptions as AddTableOptions,
)).toThrow();
for (const margin of invalidMargins) {
  expect(() => slide.addTable(
    [['Invalid table']],
    { margin } as unknown as AddTableOptions,
  )).toThrow();
}
```

Require zero getter calls, unchanged bytes/journal/slide identity, and the existing table unchanged.

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
test: cover table margin defaults
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 final-state parity

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: PptxGenJS 4.0.1 runtime and native table creation.
- Produces: paired-unit final direct-state equality and strict-difference evidence.

- [ ] **Step 1: Create equivalent native and PptxGenJS tables**

Use PptxGenJS table `margin: [0.05, 0.1, 0.15, 0.2]` and native table `margin: [3.6, 7.2, 10.8, 14.4]`. Give both tables equal geometry, row/column sizes, `valign: 'middle'`, border/fill choices, and these cells:

```ts
[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Zero', options: { margin: 0 } },
  { text: 'Points', options: { margin: [1, 2, 3, 4] } },
]
```

- [ ] **Step 2: Assert exact final state and differences**

Require both tables to expose:

```ts
[
  { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
  { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
  { top: 0, right: 0, bottom: 0, left: 0 },
  { top: 1, right: 2, bottom: 3, left: 4 },
]
```

Read slide XML and require equal `marL/marR/marT/marB` vectors, anchors, text, geometry, border, and fill snapshots. Write/reopen both and require the same matrices.

Retain PptxGenJS's `0.1` inch versus native `0.1` point difference as an explicit assertion. Require malformed native table margins to throw before mutation.

- [ ] **Step 3: Run adapter and model suites**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the adapter test and commit:

```text
test: compare table margin defaults with pptxgenjs
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 5: Smoke the actual Node/browser/types package

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: packed `@jiayunxie/pptx` Node/browser/type surfaces.
- Produces: `tableMarginCreation: true` in smoke JSON.

- [ ] **Step 1: Extend Node smoke**

Pass `margin: { top: 9, left: 18 }` in the created table options. Existing full cell margin overrides remain unchanged; the cell whose creation margin is `{}` must initially expose:

```js
{ top: 9, right: 7.2, bottom: 3.6, left: 18 }
```

Clear that inherited cell and require `undefined` immediately and after reopen. Add `tableMarginCreation` only through real snapshot/XML/reopen assertions.

The relevant changes are:

```js
], {
  name: 'Created smoke table',
  columnWidths: [inches(1), inches(3)],
  rowHeights: [inches(0.5), inches(1.5)],
  margin: { top: 9, left: 18 },
  valign: 'middle',
});
// Before edit, the empty-options cell inherits the table sides.
// After edit, the direct state is cleared and does not re-inherit.
createdTable.setCellMargins(1, 1, undefined);
```

- [ ] **Step 2: Extend browser and declaration smoke**

Mirror table default, inheritance, clear, and reopen in the browser bundle fixture. In the declaration fixture add:

```ts
const tableOptions: AddTableOptions = {
  margin: cellMargins,
  valign: cellAlignment,
};
```

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_margin_package_dir=$(mktemp -d /tmp/pptx-table-margin-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_margin_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_margin_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require every smoke flag true, including `tableMarginCreation`, `tableVerticalAlignmentCreation`, `types`, and CLI `0.1.0`.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the smoke script and commit:

```text
test: smoke packed table margin defaults
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
- Produces: accurate table-level creation and unsupported documentation.

- [ ] **Step 1: Update examples and contract**

Add a table `margin` default to examples while retaining a cell partial override. Document:

- exact point scalar/TRBL/named value shapes;
- canonical → table → cell per-side resolution;
- direct physical-cell materialization and fixed attribute order;
- no retained table metadata, so later clear does not reapply;
- omitted/runtime-undefined/empty table bytes;
- immediate `TableCell.margins` access and existing whole-replacement editor.

- [ ] **Step 2: Update compatibility and unsupported lists**

Mark table-level margin creation as supported and paired-unit PptxGenJS-final-state compatible. Keep a table-level getter/editor, dual-unit interpretation, horizontal alignment, direction/fit creation, table-level border/fill, merge, hyperlink, rich text, auto-page, and layout recomputation unsupported.

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'table-level `margin`.*(unsupported|尚未支持)|AddTableOptions.*margin.*unsupported' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
```

Review hits in context: creation support is valid, while table-level read/edit remains unsupported.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the four documents and commit:

```text
docs: document table margin creation
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
  --outputFile=/tmp/pptx-table-margin-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-margin-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_margin_qa_package_dir=$(mktemp -d /tmp/pptx-table-margin-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_margin_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_margin_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
```

- [ ] **Step 3: Generate nine real native and PptxGenJS decks**

Create `/tmp/pptx-table-margin-qa.mjs` with `apply_patch` using this complete source:

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
const output = '/tmp/pptx-table-margin';
await mkdir(output, { recursive: true });

const baseRows = [
  ['Inherited A', 'Inherited B', 'Inherited C', 'Inherited D'],
  ['Inherited E', 'Inherited F', 'Inherited G', 'Inherited H'],
];
const tableOptions = {
  name: 'Table margin creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(1.5), inches(2.5), inches(2), inches(2)],
  rowHeights: [inches(0.8), inches(1.6)],
  valign: 'middle',
};

function build(rows, options = {}) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const table = document.addSlide().addTable(rows, { ...tableOptions, ...options });
  return { document, table };
}

const omitted = build(baseRows);
const runtimeUndefined = build(baseRows, { margin: undefined });
const empty = build(baseRows, { margin: {} });
await writeFile(output + '/omitted.pptx', await omitted.document.write());
await writeFile(output + '/undefined.pptx', await runtimeUndefined.document.write());
await writeFile(output + '/empty.pptx', await empty.document.write());

const zero = build(baseRows, { margin: 0 });
assert.deepEqual(
  zero.table.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
  Array.from({ length: 2 }, () => Array.from({ length: 4 }, () => ({
    top: 0, right: 0, bottom: 0, left: 0,
  }))),
);
await writeFile(output + '/table-zero.pptx', await zero.document.write());

const tupleMargin = { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 };
const tuple = build(baseRows, { margin: [3.6, 7.2, 10.8, 14.4] });
assert.deepEqual(
  tuple.table.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
  Array.from({ length: 2 }, () => Array.from({ length: 4 }, () => tupleMargin)),
);
await writeFile(output + '/table-tuple.pptx', await tuple.document.write());

const mixedRows = [
  [
    'Inherited string',
    { text: 'Inherited object', options: {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    } },
    { text: 'Zero override', options: {
      margin: 0,
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
    } },
    { text: 'Point tuple', options: { margin: [1, 2, 3, 4] } },
  ],
  [
    { text: 'Partial bottom', options: { margin: { bottom: 12 } } },
    { text: 'Inherited fill', options: {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
    } },
    { text: 'Inherited border', options: {
      border: {
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 3,
          style: 'dash',
        },
      },
    } },
    { text: 'Inherited D' },
  ],
];
const source = build(mixedRows, { margin: [3.6, 7.2, 10.8, 14.4] });
const sourceMargins = source.table.rows.map(({ cells }) =>
  cells.map(({ margins }) => margins));
assert.deepEqual(sourceMargins, [
  [tupleMargin, tupleMargin, { top: 0, right: 0, bottom: 0, left: 0 },
    { top: 1, right: 2, bottom: 3, left: 4 }],
  [{ top: 3.6, right: 7.2, bottom: 12, left: 14.4 },
    tupleMargin, tupleMargin, tupleMargin],
]);
await writeFile(output + '/mixed-source.pptx', await source.document.write());

source.table.setCellMargins(0, 0, undefined);
source.table.setCellMargins(0, 1, { right: 5 });
source.table.setCellText(0, 1, 'Edited inherited margin');
const editedBytes = await source.document.write();
await writeFile(output + '/mixed-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[0].cells[0].margins, undefined);
assert.deepEqual(reopenedTable.rows[0].cells[1].margins, { right: 5 });
assert.equal(reopenedTable.rows[0].cells[1].text, 'Edited inherited margin');
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/mixed-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      'Inherited string',
      { text: 'Inherited object', options: { fill: { color: 'D9EAF7' } } },
      { text: 'Zero override', options: {
        margin: 0,
        border: { type: 'solid', color: 'C00000', pt: 2 },
      } },
      { text: 'Point tuple', options: { margin: [1, 2, 3, 4] } },
    ],
    [
      { text: 'Partial bottom', options: { margin: [3.6, 7.2, 12, 14.4] } },
      { text: 'Inherited fill', options: { fill: { color: 'FFF2CC' } } },
      { text: 'Inherited border', options: {
        border: [
          { type: 'none' },
          { type: 'none' },
          { type: 'dash', color: '70AD47', pt: 3 },
          { type: 'none' },
        ],
      } },
      { text: 'Inherited D', options: {} },
    ],
  ],
  {
    x: 0.75,
    y: 0.75,
    w: 8,
    h: 2.4,
    colW: [1.5, 2.5, 2, 2],
    rowH: [0.8, 1.6],
    margin: [0.05, 0.1, 0.15, 0.2],
    valign: 'middle',
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceMargins,
  reopenedMargins: reopenedTable.rows.map(({ cells }) =>
    cells.map(({ margins }) => margins)),
  editedText: reopenedTable.rows[0].cells[1].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run `node /tmp/pptx-table-margin-qa.mjs` and require all assertions to pass.

- [ ] **Step 4: Validate and diff packages**

```sh
for deck in /tmp/pptx-table-margin/*.pptx; do
  node packages/pptx/dist/cli.js --json package validate \
    "$deck" --profile powerpoint-2010
done
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-margin/omitted.pptx \
  /tmp/pptx-table-margin/undefined.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-margin/omitted.pptx \
  /tmp/pptx-table-margin/empty.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-margin/mixed-source.pptx \
  /tmp/pptx-table-margin/mixed-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-margin/mixed-edited.pptx \
  /tmp/pptx-table-margin/mixed-reopened.pptx
```

Require zero errors/warnings. Omitted/undefined/empty must have zero changed parts; source/edited may change only `/ppt/slides/slide1.xml`; edited/reopened must have zero changed parts.

- [ ] **Step 5: Render and inspect**

Convert `mixed-source.pptx`, `mixed-edited.pptx`, and `pptxgenjs-baseline.pptx` to isolated PDFs with LibreOffice, rasterize each with Poppler at 180 DPI, and run `slides_test.py` on edited and baseline decks.

Inspect every PNG at original detail. Require visible table/cell margin inheritance and override positions, clear/edit behavior, preserved vertical alignment, fill, borders, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict normalization, descriptor safety, per-side precedence, direct ownership, ordering with `valign`, omission, declarations, lifecycle isolation, PptxGenJS unit differences, smoke flags, docs boundaries, and invalid-input no-mutation assertions.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `?? .pnpm-store/`. QA-only success creates no empty commit.
