# Table Horizontal Alignment Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user requires autonomous inline execution without subagents or routine decision pauses.

**Goal:** Add strict table-level `align` creation that materializes left, center, right, or justify onto uncovered single-paragraph cells with PptxGenJS 4.0.1 final-state parity.

**Architecture:** Extend the descriptor-safe table option normalizer with one optional `TextAlignment`, resolve it onto normalized cells that lack a cell value, and reuse the existing direct paragraph alignment renderer. Keep only the resolved physical-cell paragraph state; do not add table metadata, a snapshot, or an editor.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler, and the presentation overflow checker.

## Global Constraints

- Public creation property is `AddTableOptions.align?: TextAlignment` with exact values `left`, `center`, `right`, and `justify`.
- A valid cell-level `AddTableCellOptions.align` overrides the table value; omitted or runtime-undefined cell values inherit.
- Resolved state is materialized on each uncovered physical cell's only paragraph as direct `a:pPr@algn="l|ctr|r|just"`; never write alignment to `a:tcPr`, `a:bodyPr`, table properties, or extensions.
- Omitted/runtime-undefined table `align` preserves current bytes and does not affect cell-level behavior.
- Reuse `normalizeTextAlignment()` and `renderRichTextParagraphs(..., { defaultAlign })`; do not copy token maps, add aliases, coerce, case-fold, or trim.
- PptxGenJS 4.0.1 silently ignores invalid table alignment; native creation deliberately rejects it before mutation.
- Existing table text, geometry, border, fill, margin, vertical alignment, direction, fit, identity, duplicate, and transaction behavior remains unchanged.
- No `TableModel.horizontalAlignment`, `TableCell.horizontalAlignment`, alignment editor, rich/multi-paragraph cells, merge, hyperlink, auto-page, repeated header, or layout recomputation is added in this slice.
- Every successful Task 1–6 item is tested, reviewed, committed, pushed over SSH port 443, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only success creates no empty commit.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only the named files for the current task.
- Use no subagents.

---

### Task 1: Normalize and materialize the table alignment default

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing `normalizeTextAlignment()`, normalized cell `alignment`, and `renderRichTextParagraphs()`.
- Produces: `NormalizedTableDefinition.rows` whose uncovered cells carry the resolved direct paragraph alignment.

- [ ] **Step 1: Add failing precedence, omission, and strict-input tests**

Add a focused test adjacent to the cell-level horizontal-alignment test:

```ts
it('materializes strict table horizontal alignment onto uncovered cells', () => {
  const rows = [[
    'String',
    { text: 'Object' },
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { align: undefined } },
    { text: 'Left', options: { align: 'left' } },
    { text: 'Right', options: { align: 'right' } },
    { text: 'Justify', options: { align: 'justify' } },
  ]];
  const definition = normalizeTableDefinition(rows, { align: 'center' });
  expect(definition.rows[0]!.map(({ alignment }) => alignment)).toEqual([
    'center', 'center', 'center', 'center', 'left', 'right', 'justify',
  ]);

  const xml = renderTableGraphicFrame(33, definition);
  expect([...xml.matchAll(/<a:pPr[^>]*\salgn="([^"]+)"/g)]
    .map((match) => match[1])).toEqual([
    'ctr', 'ctr', 'ctr', 'ctr', 'l', 'r', 'just',
  ]);
  expect(xml).not.toMatch(/<a:tcPr[^>]*\salgn=/);
  expect(xml).not.toMatch(/<a:bodyPr[^>]*\salgn=/);

  const expected = new Map([
    ['left', 'l'],
    ['center', 'ctr'],
    ['right', 'r'],
    ['justify', 'just'],
  ]);
  for (const [align, token] of expected) {
    const value = normalizeTableDefinition([['A']], { align });
    expect(value.rows[0]![0]!.alignment).toBe(align);
    expect(renderTableGraphicFrame(34, value)).toContain(`algn="${token}"`);
  }

  const omitted = renderTableGraphicFrame(
    35,
    normalizeTableDefinition([['Same']], {}),
  );
  const runtimeUndefined = renderTableGraphicFrame(
    35,
    normalizeTableDefinition([['Same']], { align: undefined }),
  );
  expect(runtimeUndefined).toBe(omitted);
  expect(omitted).not.toMatch(/<a:pPr[^>]*\salgn=/);
});
```

In the malformed-matrix test, add an accessor-backed table option and reuse the existing `invalidAlignments` array:

```ts
const accessorTableAlignOptions: Record<string, unknown> = {};
let tableAlignAccessorCalls = 0;
Object.defineProperty(accessorTableAlignOptions, 'align', {
  get() {
    tableAlignAccessorCalls += 1;
    return 'center';
  },
  enumerable: true,
  configurable: true,
});

expect(() => normalizeTableDefinition(
  [['A']],
  accessorTableAlignOptions,
)).toThrow(TypeError);
for (const align of invalidAlignments) {
  expect(() => normalizeTableDefinition([['A']], { align })).toThrow(TypeError);
}
expect(tableAlignAccessorCalls).toBe(0);
```

- [ ] **Step 2: Run the focused suite and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "table horizontal alignment|malformed matrices"
```

Expected: table `align` is rejected as unsupported and the new propagation assertion fails.

- [ ] **Step 3: Add the minimal normalization overlay**

Add `align` to `OPTION_KEYS`. Immediately after `readOptions(options)`, normalize and resolve it before existing border/fill/margin/valign overlays:

```ts
const tableAlignment = normalizedOptions.align === undefined
  ? undefined
  : normalizeTextAlignment(normalizedOptions.align, 'Table align');
const alignmentResolvedRows = tableAlignment === undefined
  ? normalizedRows
  : normalizedRows.map((row) => row.map((cell) =>
    cell.alignment === undefined
      ? { ...cell, alignment: tableAlignment }
      : cell));
```

Feed `alignmentResolvedRows` into the existing table-border overlay instead of `normalizedRows`. Do not change `NormalizedTableDefinition`, `renderTableCell()`, or the alignment token map.

- [ ] **Step 4: Run focused regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Require precedence, omission-byte equality, getter-free rejection, direct paragraph ownership, all four tokens, and every existing table creation test to pass.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review only the two task files for strict values, overlay order, renderer reuse, and absence of API/model expansion. Stage only those files and commit:

```text
feat: propagate table align during creation
```

Then run the standard SSH push/fetch commands and require `git rev-list --left-right --count origin/main...HEAD` to print `0 0`.

---

### Task 2: Expose the public option and prove the model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1's normalized overlay and existing public `TextAlignment`.
- Produces: public `AddTableOptions.align?: TextAlignment` plus duplicate/edit/rollback/write/reopen raw-OOXML evidence.

- [ ] **Step 1: Add the typed option and lifecycle assertions first**

In the main typed native table creation test, type the table options and add a center default:

```ts
const tableOptions: AddTableOptions = {
  align: 'center',
  name: 'Table "A"',
  x: inches(1),
  y: inches(1.5),
  width: inches(4),
  height: inches(2),
  columnWidths: [inches(1), inches(1), inches(2)],
  rowHeights: [inches(0.75), inches(1.25)],
};
const table = slide.addTable(tableRows, tableOptions);
```

Keep the first four cell values `center`, `left`, `right`, and `justify`. The fifth cell already supplies runtime `undefined`, and the sixth is a string. Change expected direct tokens everywhere in that test from:

```ts
['ctr', 'l', 'r', 'just', undefined, undefined]
```

to:

```ts
['ctr', 'l', 'r', 'just', 'ctr', 'ctr']
```

Keep the existing text, border, fill, margin, valign, transform, grid, duplicate, other-cell editor, rollback, write, reopen, opaque-part, and identity assertions. They prove the materialized paragraph state survives unrelated operations and remains isolated from the duplicate.

Extend the invalid public table option block with an accessor and every strict invalid value. Widen its declaration to `const invalidOptions: unknown[] = [...]` before pushing the new cases:

```ts
const accessorTableAlign = {};
let tableAlignAccessorCalls = 0;
Object.defineProperty(accessorTableAlign, 'align', {
  get() {
    tableAlignAccessorCalls += 1;
    return 'center';
  },
  enumerable: true,
});
invalidOptions.push(accessorTableAlign);
for (const align of [
  null, false, true, 0, '', 'Left', ' center ',
  'l', 'ctr', 'r', 'just', 'dist', 'thaiDist', 'justLow',
  [], {}, Symbol('center'),
]) {
  invalidOptions.push({ align });
}
```

Require `tableAlignAccessorCalls` to remain zero and retain the existing bytes/journal no-mutation assertions.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: excess-property error for `AddTableOptions.align`.

- [ ] **Step 3: Widen only the public table creation option**

In `slide.ts` add one field:

```ts
export interface AddTableOptions {
  readonly align?: TextAlignment;
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly rowHeights?: number | readonly number[];
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly valign?: TextBoxVerticalAlignment;
}
```

`TextAlignment` is already imported for text and cell creation. Do not add a model getter/editor or change `AddTableCellOptions`.

- [ ] **Step 4: Run lifecycle tests and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/model/src/table-create.internal.test.ts
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review public typing, the six-token precedence matrix, duplicate isolation, edit preservation, rollback, reopen, and invalid no-mutation behavior. Stage only both files and commit:

```text
feat: expose table align creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 3: Cover the public SDK lifecycle and strict failures

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: public `AddTableOptions.align` and the existing `tableCellHorizontalAlignmentTokens()` helper.
- Produces: SDK-level creation, preservation, isolation, and invalid-input proof.

- [ ] **Step 1: Add a table-level SDK lifecycle test**

Add a focused test next to the cell horizontal-alignment lifecycle test:

```ts
it('materializes table horizontal alignment through the SDK lifecycle', async () => {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const slide = document.addSlide();
  const table = slide.addTable([[
    'Inherited string',
    { text: 'Inherited object' },
    {
      text: 'Inherited undefined',
      options: { align: undefined } as unknown as AddTableCellOptions,
    },
    { text: 'Left override', options: { align: 'left' } },
    { text: 'Right override', options: { align: 'right' } },
    { text: 'Justify override', options: { align: 'justify' } },
  ]], {
    align: 'center',
    name: 'Table horizontal alignment lifecycle',
    columnWidths: inches(1.5),
    rowHeights: inches(1),
    margin: { top: 4, left: 8 },
    valign: 'middle',
  });
  const original = ['ctr', 'ctr', 'ctr', 'l', 'r', 'just'];
  expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
    .toEqual(original);

  const duplicate = document.duplicateSlide(0);
  const duplicateTable = duplicate.shapes[0] as TableModel;
  table.setCellText(0, 0, 'Inherited edited');
  table.setCellMargins(0, 1, { bottom: 9 });
  table.setCellVerticalAlignment(0, 2, 'bottom');
  table.setCellTextDirection(0, 3, 'vert270');
  table.setCellTextFit(0, 4, 'shrink');
  table.setColumnWidths(inches(1.25));
  table.setRowHeights(inches(1.25));
  expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
    .toEqual(original);
  expect(tableCellHorizontalAlignmentTokens(document, duplicate.partUri))
    .toEqual(original);
  expect(duplicateTable.rows[0]!.cells[0]!.text).toBe('Inherited string');

  const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
  const journal = [...document.opcPackage.mutations];
  expect(() => document.transaction(() => {
    slide.addTable([['Temporary']], { align: 'right' });
    throw new Error('restore table horizontal alignment');
  })).toThrow('restore table horizontal alignment');
  expect(document.opcPackage.requirePart(slide.partUri).bytes)
    .toEqual(beforeRollback);
  expect(document.opcPackage.mutations).toEqual(journal);

  const reopened = await PptxDocument.open(await document.write());
  expect(tableCellHorizontalAlignmentTokens(reopened, reopened.slides[0]!.partUri))
    .toEqual(original);
  expect(tableCellHorizontalAlignmentTokens(reopened, reopened.slides[1]!.partUri))
    .toEqual(original);
});
```

- [ ] **Step 2: Add SDK table-option rejection coverage**

Add a separate test using an existing aligned table as the mutation sentinel. Define an accessor-backed `AddTableOptions` and the complete invalid array in this test:

```ts
let getterCalls = 0;
const accessorOptions = {};
Object.defineProperty(accessorOptions, 'align', {
  get() {
    getterCalls += 1;
    return 'center';
  },
  enumerable: true,
});
const invalidAlignments: unknown[] = [
  null, false, true, 0, '', 'Left', ' center ',
  'l', 'ctr', 'r', 'just', 'dist', 'thaiDist', 'justLow',
  [], {}, Symbol('center'),
];
expect(() => slide.addTable(
  [['Accessor']],
  accessorOptions as AddTableOptions,
)).toThrow(TypeError);
for (const align of invalidAlignments) {
  expect(() => slide.addTable(
    [['Invalid']],
    { align } as unknown as AddTableOptions,
  )).toThrow(TypeError);
}
```

Require getter count zero, identical part URI list and every part's bytes, identical journal, slide/model identities, shape count, and original token array.

- [ ] **Step 3: Run SDK and model regressions**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts \
  packages/model/src/model.test.ts \
  packages/model/src/table-create.internal.test.ts
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Review target-slide isolation, inherited/override tokens, unrelated editors, rollback, reopen, accessor safety, and all-part byte preservation. Stage only the SDK test and commit:

```text
test: cover sdk table align creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 4: Compare supported and invalid PptxGenJS 4.0.1 behavior

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: native creation and adapter import of PptxGenJS-materialized paragraph XML.
- Produces: supported final-state parity and documented invalid fallback evidence without production adapter changes.

- [ ] **Step 1: Add the parity fixture**

Add a test adjacent to the cell horizontal-alignment comparison. Create both tables with a center table default and these six cells:

```ts
generated.addSlide().addTable([[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Inherited undefined', options: { align: undefined } },
  { text: 'Left override', options: { align: 'left' } },
  { text: 'Right override', options: { align: 'right' } },
  { text: 'Justify override', options: { align: 'justify' } },
]], {
  x: 0.5,
  y: 0.5,
  w: 9,
  h: 1,
  colW: Array(6).fill(1.5),
  rowH: 1,
  align: 'center',
  margin: 0.1,
  valign: 'middle',
});

const nativeTable = native.addSlide().addTable([[
  'Inherited string',
  { text: 'Inherited object' },
  {
    text: 'Inherited undefined',
    options: { align: undefined } as unknown as AddTableCellOptions,
  },
  { text: 'Left override', options: { align: 'left' } },
  { text: 'Right override', options: { align: 'right' } },
  { text: 'Justify override', options: { align: 'justify' } },
]], {
  x: inches(0.5),
  y: inches(0.5),
  columnWidths: inches(1.5),
  rowHeights: inches(1),
  align: 'center',
  margin: 7.2,
  valign: 'middle',
});
```

Import the generated deck and require both raw token arrays to equal:

```ts
['ctr', 'ctr', 'ctr', 'l', 'r', 'just']
```

Also require equal text, transform, column widths, row heights, margins, vertical alignment, border/fill snapshots, absence of `tcPr/bodyPr@algn`, and the same arrays after native/imported write and reopen.

- [ ] **Step 2: Record invalid-runtime divergence**

Generate a PptxGenJS table with `align: 'dist' as never`, one uncovered string cell, and one valid right cell override. Require imported direct tokens `[undefined, 'r']`. Then call native `addTable()` with the same invalid table value and require `TypeError`, unchanged target slide bytes/journal/shape count, and stable native table identity.

```ts
generatedInvalid.addSlide().addTable([
  ['Invalid inherited', { text: 'Right override', options: { align: 'right' } }],
], { align: 'dist' as never, x: 0.5, y: 0.5, w: 4, h: 1 });
```

- [ ] **Step 3: Run adapter, model, and type regressions**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Review supported token parity, precedence, geometry/property equality, invalid fallback difference, reopen preservation, and no production adapter diff. Stage only the adapter test and commit:

```text
test: compare table align creation with pptxgenjs
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 5: Prove packed Node, browser, declaration, and CLI surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: built `@jiayunxie/pptx` package and public declaration.
- Produces: actual-tarball runtime/type proof for inherited table horizontal alignment.

- [ ] **Step 1: Extend Node and browser runtime fixtures**

For both existing 2x2 creation tables:

- add table option `align: 'center'`;
- remove explicit `align: 'center'` from the second cell so it inherits;
- keep explicit left/right/justify on the other cells;
- keep expected token arrays `['l', 'ctr', 'r', 'just']` before edits and after reopen;
- add a `tableHorizontalAlignmentCreation` boolean gate and require it in the final JSON while retaining the cell-level gate.

The Node creation options must be:

```js
{
  name: 'Created smoke table',
  align: 'center',
  columnWidths: [inches(1), inches(3)],
  rowHeights: [inches(0.5), inches(1.5)],
  fill: tableCreationFill,
  margin: { top: 9, left: 18 },
  valign: 'middle',
}
```

Use the same `align` field in the browser fixture.

- [ ] **Step 2: Extend declaration compilation**

Add:

```ts
const tableHorizontalAlignment: TextAlignment = 'center';
```

and include it in typed table options:

```ts
const tableOptions: AddTableOptions = {
  align: tableHorizontalAlignment,
  name: 'Typed table',
  x: inches(1),
  columnWidths: [inches(1), inches(3)],
  rowHeights: [inches(0.5), inches(1.5)],
  border: cellBorderInput,
  fill: cellFill,
  margin: cellMargins,
  valign: cellAlignment,
};
```

Reference `tableHorizontalAlignment` in the final `void [...]` tuple.

- [ ] **Step 3: Build and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_align_package_dir=$(mktemp -d /tmp/pptx-table-align-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_align_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_align_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require every JSON flag true, including Node/browser table-level inheritance, declarations, and CLI doctor version `0.1.0`.

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Review that the tarball rather than workspace source is installed, center is inherited rather than cell-explicit, all existing smoke gates remain true, and only the script changed. Commit:

```text
test: smoke packed table align creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 6: Document the supported boundary

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: final public behavior and compatibility evidence.
- Produces: accurate supported/unsupported documentation.

- [ ] **Step 1: Update release and API documentation**

Add this exact contract in the repository's existing prose style:

```text
AddTableOptions.align accepts left, center, right, or justify during creation.
It is materialized onto each single-paragraph cell that omits cell align; a valid
cell value wins. Omitted or runtime-undefined table align preserves current bytes.
The default source is not retained after creation.
```

State that final ownership is direct `a:pPr@algn`, not `tcPr`, `bodyPr`, or table metadata. Keep existing-deck horizontal-alignment snapshot/editing and rich/multi-paragraph cell alignment listed as unsupported.

- [ ] **Step 2: Update PptxGenJS compatibility wording**

Mark the four supported table values and cell precedence as final-state compatible with PptxGenJS 4.0.1. Record that PptxGenJS silently drops invalid table runtime values while native rejects them before mutation. Do not claim byte-identical whole-slide XML.

- [ ] **Step 3: Run documentation boundary scans and full typecheck**

```sh
rg -n 'table-level `align`.*(unsupported|尚未支持)|AddTableOptions.*align.*unsupported|table default propagation.*unsupported' \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
git diff --check
```

Review every match in context: creation support is valid, while existing-deck getter/editor and rich/multi-paragraph alignment remain unsupported.

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Review public names, strict values, precedence, direct ownership, omission, invalid divergence, and unsupported boundaries. Stage only the four docs and commit:

```text
docs: document table align creation
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
  --outputFile=/tmp/pptx-table-align-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-align-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_align_qa_package_dir=$(mktemp -d /tmp/pptx-table-align-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_align_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_align_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
```

Require every smoke flag true, offline doctor success, and CLI version `0.1.0`.

- [ ] **Step 3: Generate real native and PptxGenJS decks**

Create `/tmp/pptx-table-align-qa.mjs` with `apply_patch`. Its fixed fixture is:

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
const output = '/tmp/pptx-table-align';
await mkdir(output, { recursive: true });

const baseRows = [
  ['Inherited A', 'Inherited B', 'Inherited C', 'Inherited D'],
  ['Inherited E', 'Inherited F', 'Inherited G', 'Inherited H'],
];
const tableOptions = {
  name: 'Table align creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(1.5), inches(2.5), inches(2), inches(2)],
  rowHeights: [inches(0.8), inches(1.6)],
};

function tokens(document) {
  const xml = new TextDecoder().decode(
    document.opcPackage.requirePart(document.slides[0].partUri).bytes,
  );
  return [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
    .map((match) => match[0].match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
}

function build(rows, options = {}) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const table = document.addSlide().addTable(rows, { ...tableOptions, ...options });
  return { document, table };
}

const omitted = build(baseRows);
const runtimeUndefined = build(baseRows, { align: undefined });
await writeFile(output + '/omitted.pptx', await omitted.document.write());
await writeFile(output + '/undefined.pptx', await runtimeUndefined.document.write());
assert.deepEqual(tokens(omitted.document), Array(8).fill(undefined));
assert.deepEqual(tokens(runtimeUndefined.document), Array(8).fill(undefined));

for (const [align, token] of [
  ['left', 'l'],
  ['center', 'ctr'],
  ['right', 'r'],
  ['justify', 'just'],
]) {
  const built = build(baseRows, { align });
  assert.deepEqual(tokens(built.document), Array(8).fill(token));
  await writeFile(output + `/table-${align}.pptx`, await built.document.write());
}

const mixedRows = [
  [
    'Inherited string',
    { text: 'Inherited object', options: {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    } },
    { text: 'Left override', options: {
      align: 'left',
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
    } },
    { text: 'Right override', options: { align: 'right' } },
  ],
  [
    { text: 'Inherited with margin', options: { margin: { top: 4, left: 18 } } },
    { text: 'Justify override sentence', options: {
      align: 'justify',
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
    } },
    { text: 'Undefined inherits', options: { align: undefined } },
    { text: 'Inherited D', options: { valign: 'bottom' } },
  ],
];
const source = build(mixedRows, { align: 'center', valign: 'middle' });
const sourceTokens = tokens(source.document);
assert.deepEqual(sourceTokens, [
  'ctr', 'ctr', 'l', 'r', 'ctr', 'just', 'ctr', 'ctr',
]);
await writeFile(output + '/mixed-source.pptx', await source.document.write());

source.table.setCellText(0, 0, 'Edited inherited center');
source.table.setCellMargins(0, 1, { bottom: 9 });
source.table.setCellVerticalAlignment(1, 3, 'top');
const editedBytes = await source.document.write();
await writeFile(output + '/mixed-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.deepEqual(tokens(reopened), sourceTokens);
assert.equal(reopenedTable.rows[0].cells[0].text, 'Edited inherited center');
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
      { text: 'Left override', options: {
        align: 'left',
        border: { type: 'solid', color: 'C00000', pt: 2 },
      } },
      { text: 'Right override', options: { align: 'right' } },
    ],
    [
      { text: 'Inherited with margin', options: { margin: [4, 7.2, 3.6, 18] } },
      { text: 'Justify override sentence', options: {
        align: 'justify',
        fill: { color: 'FFF2CC' },
      } },
      { text: 'Undefined inherits', options: { align: undefined } },
      { text: 'Inherited D', options: { valign: 'bottom' } },
    ],
  ],
  {
    x: 0.75,
    y: 0.75,
    w: 8,
    h: 2.4,
    colW: [1.5, 2.5, 2, 2],
    rowH: [0.8, 1.6],
    align: 'center',
    valign: 'middle',
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceTokens,
  reopenedTokens: tokens(reopened),
  editedText: reopenedTable.rows[0].cells[0].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run `node /tmp/pptx-table-align-qa.mjs`. Require source/reopened tokens `['ctr','ctr','l','r','ctr','just','ctr','ctr']`, edited text, widths `[1371600,2286000,1828800,1828800]`, and heights `[731520,1463040]`.

- [ ] **Step 4: Validate and diff packages**

Validate all ten decks with the PowerPoint 2010 profile:

```sh
for deck in /tmp/pptx-table-align/*.pptx; do
  node packages/pptx/dist/cli.js --json package validate \
    "$deck" --profile powerpoint-2010
done
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-align/omitted.pptx \
  /tmp/pptx-table-align/undefined.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-align/mixed-source.pptx \
  /tmp/pptx-table-align/mixed-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-align/mixed-edited.pptx \
  /tmp/pptx-table-align/mixed-reopened.pptx
```

Require zero errors/warnings; omitted versus undefined has zero changed parts; source versus edited changes only `/ppt/slides/slide1.xml`; edited versus reopened has zero changed parts.

- [ ] **Step 5: Render and inspect at original resolution**

Use isolated LibreOffice profiles, convert `mixed-source.pptx`, `mixed-edited.pptx`, and `pptxgenjs-baseline.pptx` to PDF, rasterize every page with `pdftoppm -png -r 180`, and inspect each PNG at original detail:

```sh
table_align_render_dir=$(mktemp -d /tmp/pptx-table-align-rendered.XXXXXX)
mkdir -p "$table_align_render_dir/source" \
  "$table_align_render_dir/edited" \
  "$table_align_render_dir/baseline"
table_align_source_profile=$(mktemp -d /tmp/pptx-table-align-source-profile.XXXXXX)
table_align_edited_profile=$(mktemp -d /tmp/pptx-table-align-edited-profile.XXXXXX)
table_align_baseline_profile=$(mktemp -d /tmp/pptx-table-align-baseline-profile.XXXXXX)
soffice -env:UserInstallation="file://$table_align_source_profile" \
  --headless --convert-to pdf --outdir "$table_align_render_dir/source" \
  /tmp/pptx-table-align/mixed-source.pptx
soffice -env:UserInstallation="file://$table_align_edited_profile" \
  --headless --convert-to pdf --outdir "$table_align_render_dir/edited" \
  /tmp/pptx-table-align/mixed-edited.pptx
soffice -env:UserInstallation="file://$table_align_baseline_profile" \
  --headless --convert-to pdf --outdir "$table_align_render_dir/baseline" \
  /tmp/pptx-table-align/pptxgenjs-baseline.pptx
pdftoppm -png -r 180 "$table_align_render_dir/source/mixed-source.pdf" \
  "$table_align_render_dir/source/slide"
pdftoppm -png -r 180 "$table_align_render_dir/edited/mixed-edited.pdf" \
  "$table_align_render_dir/edited/slide"
pdftoppm -png -r 180 "$table_align_render_dir/baseline/pptxgenjs-baseline.pdf" \
  "$table_align_render_dir/baseline/slide"
```

Run both overflow checks:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-align/mixed-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-align/pptxgenjs-baseline.pptx
```

Require visible center inheritance and left/right/justify overrides, preserved margin/border/fill/valign, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict normalization, descriptor safety, precedence, direct ownership, omission, declarations, lifecycle isolation, PptxGenJS divergence, smoke flags, docs boundaries, and invalid-input no-mutation assertions.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `?? .pnpm-store/`. If QA finds a defect, fix only the responsible task files, rerun the focused/full gates, review, commit, push, fetch, and re-prove synchronization; otherwise create no empty commit.
