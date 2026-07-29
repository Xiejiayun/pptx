# Table Vertical Alignment Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user requires inline execution without subagents.

**Goal:** Add strict table-level `valign` creation that materializes top, middle, or bottom onto uncovered cells with PptxGenJS 4.0.1 final-state parity.

**Architecture:** Extend the existing descriptor-safe table option normalizer with one optional `TextBoxVerticalAlignment`, resolve it onto normalized cells that lack a cell value, and reuse the current direct cell anchor renderer. Keep the resolved state only on physical cells; do not add table metadata or a table-level editor.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler.

## Global Constraints

- Public creation property is `AddTableOptions.valign?: TextBoxVerticalAlignment` with exact values `top`, `middle`, and `bottom`.
- A valid cell-level `AddTableCellOptions.valign` overrides the table value; omitted or runtime-undefined cell values inherit.
- Resolved state is materialized on every uncovered physical cell as direct `a:tcPr@anchor="t|ctr|b"`; never write `a:bodyPr@anchor` or table metadata.
- Omitted/runtime-undefined table `valign` preserves current bytes and does not affect cell-level behavior.
- Serialization remains margins, optional anchor, L/R/T/B borders, then fill.
- Reuse `normalizeTextBoxVerticalAlignment()` and the existing cell renderer; do not accept aliases, coercion, case folding, whitespace trimming, `just`, `dist`, `mid`, `center`, or `distributed`.
- Clearing a materialized cell after creation removes its direct anchor and does not reapply the original table value.
- No `TableModel.verticalAlignment`, table-level editor, horizontal alignment creation, direction/fit creation, merge, hyperlink, rich text, auto-page, or layout recomputation in this slice.
- Every successful small item is reviewed, committed, pushed, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only gates do not create empty commits.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and materialize the table default

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing `normalizeTextBoxVerticalAlignment()`, normalized cell `verticalAlignment`, and `renderTableCellVerticalAlignmentAttribute()`.
- Produces: `NormalizedTableDefinition.rows` whose uncovered cells carry the resolved table value.

- [ ] **Step 1: Add failing precedence and exact-output tests**

Add a focused test adjacent to the cell-level vertical-alignment creation test:

```ts
it('materializes strict table vertical alignment onto uncovered cells', () => {
  const rows = [[
    'String',
    { text: 'Object' },
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { valign: undefined } },
    { text: 'Top', options: { valign: 'top' } },
    { text: 'Bottom', options: { valign: 'bottom' } },
  ]];
  const definition = normalizeTableDefinition(rows, { valign: 'middle' });
  expect(definition.rows[0]!.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
    'middle',
    'middle',
    'middle',
    'middle',
    'top',
    'bottom',
  ]);

  const xml = renderTableGraphicFrame(31, definition);
  const anchors = [...xml.matchAll(/<a:tcPr[^>]* anchor="([^"]+)"/g)]
    .map((match) => match[1]);
  expect(anchors).toEqual(['ctr', 'ctr', 'ctr', 'ctr', 't', 'b']);
  expect(xml).not.toMatch(/<a:bodyPr[^>]* anchor=/);

  const omitted = renderTableGraphicFrame(
    32,
    normalizeTableDefinition([['Same']], {}),
  );
  const runtimeUndefined = renderTableGraphicFrame(
    32,
    normalizeTableDefinition([['Same']], { valign: undefined }),
  );
  expect(runtimeUndefined).toBe(omitted);
  expect(omitted).not.toContain(' anchor=');
});
```

Extend malformed table options with every invalid strict value and an accessor-backed table option:

```ts
const accessorValignOptions = {};
Object.defineProperty(accessorValignOptions, 'valign', {
  get() {
    tableValignAccessorCalls += 1;
    return 'top';
  },
  enumerable: true,
});

const invalidTableValigns = [
  null, false, true, 0, '', 'Top', ' top ', 'mid', 'center',
  'just', 'dist', 'distributed', [], {}, Symbol('top'),
];
```

Require every `{ valign }` option and the accessor object to throw, `tableValignAccessorCalls` to stay zero, and slide bytes/mutation journal to remain unchanged.

- [ ] **Step 2: Run focused tests and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "table vertical alignment|malformed matrices"
```

Expected: table `valign` is rejected as an unsupported option.

- [ ] **Step 3: Implement the minimal normalization overlay**

Add `valign` to `OPTION_KEYS`. Immediately after `readOptions(options)`, normalize it:

```ts
const tableVerticalAlignment = normalizedOptions.valign === undefined
  ? undefined
  : normalizeTextBoxVerticalAlignment(
    normalizedOptions.valign,
    'Table valign',
  );
const resolvedRows = tableVerticalAlignment === undefined
  ? normalizedRows
  : normalizedRows.map((row) => row.map((cell) =>
    cell.verticalAlignment === undefined
      ? { ...cell, verticalAlignment: tableVerticalAlignment }
      : cell));
```

Use `resolvedRows` in the returned definition. Do not add a definition-level alignment field or change `renderTableCell()`.

- [ ] **Step 4: Run focused regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Require exact precedence, omitted-byte equality, getter-free rejection, no body anchor, and all existing cell-level tests to pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only the two files and commit:

```text
feat: propagate table valign during creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`; `.pnpm-store/` remains untracked.

---

### Task 2: Expose the public option and prove the model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1's normalization overlay.
- Produces: public `AddTableOptions.valign?: TextBoxVerticalAlignment` and create/read/edit/write/reopen evidence.

- [ ] **Step 1: Add the typed option and lifecycle assertions first**

In a native table creation test, use a typed option:

```ts
const options: AddTableOptions = {
  valign: 'middle',
  columnWidths: inches(2),
  rowHeights: inches(1),
};
const table = slide.addTable([[
  'Inherited string',
  { text: 'Inherited object' },
  { text: 'Top override', options: { valign: 'top' } },
  { text: 'Bottom override', options: { valign: 'bottom' } },
]], options);
expect(table.rows[0]!.cells.map(
  ({ verticalAlignment }) => verticalAlignment)).toEqual([
  'middle', 'middle', 'top', 'bottom',
]);
```

Clear the first cell, edit the second to bottom, duplicate the slide, write/reopen, and require:

```ts
expect(reopenedTable.rows[0]!.cells.map(
  ({ verticalAlignment }) => verticalAlignment)).toEqual([
  undefined, 'bottom', 'top', 'bottom',
]);
```

The duplicate must retain the pre-edit materialized values, proving no source-level default is consulted later.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: excess-property error for `AddTableOptions.valign`.

- [ ] **Step 3: Widen only the public table option**

In `packages/model/src/slide.ts`:

```ts
export interface AddTableOptions {
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly rowHeights?: number | readonly number[];
  readonly valign?: TextBoxVerticalAlignment;
}
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
feat: expose table valign creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 3: Cover the SDK lifecycle and invalid public inputs

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: public `AddTableOptions.valign` and existing `TableModel` cell editor.
- Produces: package-facing propagation, override, rollback, duplicate, and no-mutation evidence.

- [ ] **Step 1: Add table-level propagation to the created-table fixture**

Pass `valign: 'middle'` in the fixture's table options. Keep explicit top and bottom cell values, and change one previously omitted cell expectation from `undefined` to `middle`.

Require the immediate matrix to contain inherited middle values and explicit overrides, then clear one inherited cell and set another to bottom.

- [ ] **Step 2: Extend duplicate, rollback, and reopen assertions**

Require the duplicate to preserve the original materialized matrix. Require rollback to restore both resolved anchors and all existing text, margin, border, fill, width, and height state. After write/reopen, require cleared cells to remain `undefined` rather than regaining middle.

- [ ] **Step 3: Expand invalid table-option coverage**

Reuse the existing `invalidValigns` array. Add an accessor-backed `AddTableOptions` object and run:

```ts
expect(() => slide.addTable([['Accessor']], accessorOptions as AddTableOptions)).toThrow();
for (const valign of invalidValigns) {
  expect(() => slide.addTable(
    [['Invalid']],
    { valign } as unknown as AddTableOptions,
  )).toThrow(TypeError);
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
test: cover table valign creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 final-state parity

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: PptxGenJS 4.0.1 runtime and native table creation.
- Produces: exact supported-value equality and strict-difference evidence.

- [ ] **Step 1: Create equivalent native and PptxGenJS tables**

Create both with table `valign: 'middle'` and cells:

```ts
[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Top', options: { valign: 'top' } },
  { text: 'Bottom', options: { valign: 'bottom' } },
]
```

Use equal geometry, column widths, row heights, margins, borders, and fill.

- [ ] **Step 2: Assert final direct state**

Require both imported/native tables to expose:

```ts
['middle', 'middle', 'top', 'bottom']
```

Read each slide XML and require direct anchor tokens `ctr, ctr, t, b`, no `bodyPr@anchor`, and equal text, geometry, margins, border, and fill snapshots. Write/reopen both and require the same matrix.

Retain the existing PptxGenJS invalid `mid` / `distributed` passthrough assertions and native strict rejection; do not add aliases.

- [ ] **Step 3: Run adapter and model suites**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the adapter test and commit:

```text
test: compare table valign creation with pptxgenjs
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 5: Smoke the actual Node/browser/types package

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: packed `@jiayunxie/pptx` Node/browser/type surfaces.
- Produces: `tableVerticalAlignmentCreation: true` in smoke JSON.

- [ ] **Step 1: Extend Node smoke**

Pass `valign: 'middle'` in the created table options. Preserve explicit top and bottom cells, and require uncovered cells to be middle immediately and after reopen. Clear one inherited cell and require it to remain undefined after reopen.

Add:

```js
tableVerticalAlignmentCreation: true
```

only through real assertions, not a constant.

- [ ] **Step 2: Extend browser and declaration smoke**

Mirror the Node behavior in the browser bundle fixture. In the declaration fixture, use:

```ts
const tableOptions: AddTableOptions = {
  valign: cellAlignment,
};
```

The existing `cellAlignment: TextBoxVerticalAlignment = 'middle'` proves exact public type reuse.

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_valign_package_dir=$(mktemp -d /tmp/pptx-table-valign-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_valign_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_valign_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require every smoke flag true, including `tableVerticalAlignmentCreation`, `types`, and CLI `0.1.0`.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the smoke script and commit:

```text
test: smoke packed table valign creation
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

Add `valign: 'middle'` to a table options example while keeping a cell override. Document:

- `AddTableOptions.valign` exact values;
- cell override and omitted/undefined inheritance;
- direct per-cell materialization and `t/ctr/b` mapping;
- no retained table metadata, so a later clear does not reapply;
- omitted table bytes and existing serializer order;
- immediate per-cell access/editing.

- [ ] **Step 2: Update compatibility and unsupported lists**

Mark table-level `valign` creation as supported and PptxGenJS-final-state compatible. Keep a table-level editor/getter, horizontal alignment, direction/fit creation, merge, hyperlink, rich text, auto-page, and layout recomputation unsupported.

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'table-level `valign`.*(unsupported|尚未支持)|AddTableOptions.*valign.*unsupported' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
```

Review hits in context: creation support is valid, while table-level read/edit must remain unsupported.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the four documents and commit:

```text
docs: document table valign creation
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
  --outputFile=/tmp/pptx-table-valign-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-valign-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_valign_qa_package_dir=$(mktemp -d /tmp/pptx-table-valign-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_valign_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_valign_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
```

Require every smoke flag true, offline doctor success, and CLI `0.1.0`.

- [ ] **Step 3: Generate real native and PptxGenJS decks**

Create `/tmp/pptx-table-valign-qa.mjs` with `apply_patch`:

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
const output = '/tmp/pptx-table-valign';
await mkdir(output, { recursive: true });

const baseRows = [
  ['Inherited A', 'Inherited B', 'Top override', 'Bottom override'],
  ['Inherited C', 'Middle fill', 'Bottom border', 'Inherited D'],
];
const tableOptions = {
  name: 'Table valign creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(1.5), inches(2.5), inches(2), inches(2)],
  rowHeights: [inches(0.8), inches(1.6)],
};

function build(rows, options = {}) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const table = document.addSlide().addTable(rows, { ...tableOptions, ...options });
  return { document, table };
}

const omitted = build(baseRows);
const runtimeUndefined = build(baseRows, { valign: undefined });
await writeFile(output + '/omitted.pptx', await omitted.document.write());
await writeFile(output + '/undefined.pptx', await runtimeUndefined.document.write());

for (const valign of ['top', 'middle', 'bottom']) {
  const built = build(baseRows, { valign });
  assert.deepEqual(
    built.table.rows.map(({ cells }) => cells.map(({ verticalAlignment }) => verticalAlignment)),
    Array.from({ length: 2 }, () => Array(4).fill(valign)),
  );
  await writeFile(output + `/table-${valign}.pptx`, await built.document.write());
}

const mixedRows = [
  [
    'Inherited string',
    { text: 'Inherited object', options: {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    } },
    { text: 'Top override', options: {
      valign: 'top',
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
    } },
    { text: 'Bottom override', options: { valign: 'bottom' } },
  ],
  [
    { text: 'Inherited with margin', options: { margin: { top: 4, left: 18 } } },
    { text: 'Middle fill', options: {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
    } },
    { text: 'Bottom border', options: {
      valign: 'bottom',
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
const source = build(mixedRows, { valign: 'middle' });
const sourceAlignments = source.table.rows.map(({ cells }) =>
  cells.map(({ verticalAlignment }) => verticalAlignment));
assert.deepEqual(sourceAlignments, [
  ['middle', 'middle', 'top', 'bottom'],
  ['middle', 'middle', 'bottom', 'middle'],
]);
await writeFile(output + '/mixed-source.pptx', await source.document.write());

source.table.setCellVerticalAlignment(0, 0, undefined);
source.table.setCellVerticalAlignment(0, 1, 'bottom');
source.table.setCellText(0, 1, 'Edited inherited to bottom');
const editedBytes = await source.document.write();
await writeFile(output + '/mixed-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[0].cells[0].verticalAlignment, undefined);
assert.equal(reopenedTable.rows[0].cells[1].verticalAlignment, 'bottom');
assert.equal(reopenedTable.rows[0].cells[1].text, 'Edited inherited to bottom');
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
      { text: 'Top override', options: {
        valign: 'top',
        border: { type: 'solid', color: 'C00000', pt: 2 },
      } },
      { text: 'Bottom override', options: { valign: 'bottom' } },
    ],
    [
      { text: 'Inherited with margin', options: { margin: [4, 7.2, 3.6, 18] } },
      { text: 'Middle fill', options: { fill: { color: 'FFF2CC' } } },
      { text: 'Bottom border', options: {
        valign: 'bottom',
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
    valign: 'middle',
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceAlignments,
  reopenedAlignments: reopenedTable.rows.map(({ cells }) =>
    cells.map(({ verticalAlignment }) => verticalAlignment)),
  editedText: reopenedTable.rows[0].cells[1].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run:

```sh
node /tmp/pptx-table-valign-qa.mjs
```

Require source alignments `[["middle","middle","top","bottom"],["middle","middle","bottom","middle"]]`, reopened first row `[null,"bottom","top","bottom"]`, widths `[1371600,2286000,1828800,1828800]`, and heights `[731520,1463040]`.

- [ ] **Step 4: Validate and diff packages**

Validate all nine decks:

```sh
for deck in /tmp/pptx-table-valign/*.pptx; do
  node packages/pptx/dist/cli.js --json package validate \
    "$deck" --profile powerpoint-2010
done
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-valign/omitted.pptx \
  /tmp/pptx-table-valign/undefined.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-valign/mixed-source.pptx \
  /tmp/pptx-table-valign/mixed-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-valign/mixed-edited.pptx \
  /tmp/pptx-table-valign/mixed-reopened.pptx
```

Require zero errors/warnings. Diff omitted against runtime undefined and require zero changed parts. Diff source against edited and require only `/ppt/slides/slide1.xml`. Diff edited against reopened and require zero changed parts.

- [ ] **Step 5: Render and inspect**

Create an isolated render directory and profiles, then convert and rasterize:

```sh
table_valign_render_dir=$(mktemp -d /tmp/pptx-table-valign-rendered.XXXXXX)
mkdir -p "$table_valign_render_dir/source" \
  "$table_valign_render_dir/edited" \
  "$table_valign_render_dir/baseline"
table_valign_source_profile=$(mktemp -d /tmp/pptx-table-valign-source-profile.XXXXXX)
table_valign_edited_profile=$(mktemp -d /tmp/pptx-table-valign-edited-profile.XXXXXX)
table_valign_baseline_profile=$(mktemp -d /tmp/pptx-table-valign-baseline-profile.XXXXXX)
soffice -env:UserInstallation="file://$table_valign_source_profile" \
  --headless --convert-to pdf --outdir "$table_valign_render_dir/source" \
  /tmp/pptx-table-valign/mixed-source.pptx
soffice -env:UserInstallation="file://$table_valign_edited_profile" \
  --headless --convert-to pdf --outdir "$table_valign_render_dir/edited" \
  /tmp/pptx-table-valign/mixed-edited.pptx
soffice -env:UserInstallation="file://$table_valign_baseline_profile" \
  --headless --convert-to pdf --outdir "$table_valign_render_dir/baseline" \
  /tmp/pptx-table-valign/pptxgenjs-baseline.pptx
pdftoppm -png -r 180 "$table_valign_render_dir/source/mixed-source.pdf" \
  "$table_valign_render_dir/source/slide"
pdftoppm -png -r 180 "$table_valign_render_dir/edited/mixed-edited.pdf" \
  "$table_valign_render_dir/edited/slide"
pdftoppm -png -r 180 "$table_valign_render_dir/baseline/pptxgenjs-baseline.pdf" \
  "$table_valign_render_dir/baseline/slide"
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-valign/mixed-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-valign/pptxgenjs-baseline.pptx
```

Inspect every generated PNG at original detail. Require visible inheritance/override positions, clear/edit behavior, preserved margins/borders/fills, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict normalization, descriptor safety, precedence, direct ownership, ordering, omission, declarations, lifecycle isolation, PptxGenJS differences, smoke flags, docs boundaries, and invalid-input no-mutation assertions.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `?? .pnpm-store/`. QA-only success creates no empty commit.
