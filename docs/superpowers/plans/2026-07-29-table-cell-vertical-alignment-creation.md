# Table Cell Vertical Alignment Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user requires inline execution without subagents.

**Goal:** Let native `slide.addTable()` create cells with strict top/middle/bottom direct vertical alignment through `{ text, options: { valign } }`, while preserving omitted bytes and the existing editor.

**Architecture:** Reuse the public `TextBoxVerticalAlignment` value model and normalizer, add a narrow creation renderer to the existing table-cell vertical-alignment codec, and carry an optional normalized alignment through the table creation definition. Keep direct ownership on `tcPr@anchor`; do not add table-level propagation in this slice.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler.

## Global Constraints

- Public creation property is `AddTableCellOptions.valign?: TextBoxVerticalAlignment` with exact values `top`, `middle`, and `bottom`.
- Storage is the created physical cell's direct `a:tcPr@anchor`; never write `a:bodyPr@anchor`.
- Omitted/undefined `valign`, string cells, `{ text }`, and empty cell options retain current bytes and direct anchor absence.
- `top` / `middle` / `bottom` render canonical `t` / `ctr` / `b` after `marL/marR/marT/marB` and before border/fill children.
- Reuse the existing strict value normalizer; do not add aliases, coercion, case folding, whitespace trimming, `just`, or `dist`.
- Existing `TableCell.verticalAlignment` and `setCellVerticalAlignment()` semantics remain unchanged.
- PptxGenJS invalid runtime passthrough is documented and tested, not copied into native behavior.
- Table-level `valign`, alignment/direction/fit creation, merge, hyperlink, rich text, auto-page, layout, and unrelated options remain outside this slice.
- Every successful small item is reviewed, committed, pushed, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only gates do not create empty commits.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and render cell valign during table creation

**Files:**
- Modify: `packages/model/src/table-cell-vertical-alignment.internal.ts`
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing `TextBoxVerticalAlignment`, `normalizeTextBoxVerticalAlignment()`, canonical margin renderer, border renderer, and fill renderer.
- Produces:

```ts
export function renderTableCellVerticalAlignmentAttribute(
  value: TextBoxVerticalAlignment | undefined,
): string;
```

- [ ] **Step 1: Add failing internal normalization and exact-output tests**

Add one test adjacent to the margin creation test:

```ts
it('normalizes and renders strict cell vertical alignment after margins', () => {
  const nullOptions = Object.assign(Object.create(null), { valign: 'bottom' });
  const rows = [[
    'String',
    { text: 'Empty options', options: {} },
    { text: 'Undefined', options: { valign: undefined } },
    { text: 'Top', options: { valign: 'top' } },
    { text: 'Middle', options: { valign: 'middle' } },
    { text: 'Bottom', options: { valign: 'bottom' } },
    { text: 'Null prototype', options: nullOptions },
    { text: 'Combined', options: {
      valign: 'middle',
      margin: { top: 4, left: 8 },
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
      },
    } },
  ]];
  const definition = normalizeTableDefinition(rows, undefined);
  expect(definition.rows[0]!.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
    undefined,
    undefined,
    undefined,
    'top',
    'middle',
    'bottom',
    'bottom',
    'middle',
  ]);

  const equivalent = [
    [['Same']],
    [[{ text: 'Same' }]],
    [[{ text: 'Same', options: {} }]],
    [[{ text: 'Same', options: { valign: undefined } }]],
  ].map((input) => renderTableGraphicFrame(
    20,
    normalizeTableDefinition(input, undefined),
  ));
  expect(new Set(equivalent).size).toBe(1);
  expect(equivalent[0]).not.toContain(' anchor=');

  const renderCell = (cell: unknown): string => renderTableGraphicFrame(
    21,
    normalizeTableDefinition([[cell]], undefined),
  );
  expect(renderCell(rows[0]![3])).toContain(
    '<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="t">',
  );
  expect(renderCell(rows[0]![4])).toContain(' marB="45720" anchor="ctr">');
  expect(renderCell(rows[0]![5])).toContain(' marB="45720" anchor="b">');
  expect(renderCell(rows[0]![7])).toMatch(
    /<a:tcPr marL="101600" marR="91440" marT="50800" marB="45720" anchor="ctr"><a:lnL[\s\S]*<\/a:lnB><a:solidFill>/,
  );
});
```

Extend the malformed-cell test with an accessor options property and invalid values:

```ts
const accessorValignOptions = {};
Object.defineProperty(accessorValignOptions, 'valign', {
  get() {
    cellAccessorCalls += 1;
    return 'top';
  },
  enumerable: true,
  configurable: true,
});

const invalidValigns = [
  null,
  false,
  true,
  0,
  '',
  'Top',
  ' top ',
  'mid',
  'center',
  'just',
  'dist',
  'distributed',
  [],
  {},
  Symbol('top'),
];
```

Add `{ text: 'A', options: accessorValignOptions }` and every `invalidValigns` item wrapped as `{ text: 'A', options: { valign } }` to `invalidCells`. Keep `expect(cellAccessorCalls).toBe(0)`.

- [ ] **Step 2: Run the focused internal suite and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "vertical alignment|malformed matrices"
```

Expected: `valign` is rejected as an unsupported cell option or normalized cells lack `verticalAlignment`.

- [ ] **Step 3: Add the narrow renderer and normalized field**

In `table-cell-vertical-alignment.internal.ts`, export:

```ts
export function renderTableCellVerticalAlignmentAttribute(
  value: TextBoxVerticalAlignment | undefined,
): string {
  return value === undefined ? '' : ` anchor="${TO_OOXML[value]}"`;
}
```

In `table-create.internal.ts`:

```ts
import {
  renderTableCellVerticalAlignmentAttribute,
} from './table-cell-vertical-alignment.internal.js';
import {
  normalizeTextBoxVerticalAlignment,
} from './text-box-vertical-alignment.internal.js';
import type {
  TextBoxMarginInput,
  TextBoxMargins,
  TextBoxVerticalAlignment,
} from './text.js';
```

Extend the normalized cell and options result:

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}

function normalizeTableCellOptions(
  value: unknown,
  context: string,
): Pick<
  NormalizedTableCell,
  'borders' | 'fill' | 'margins' | 'verticalAlignment'
> {
  if (value === undefined) return {};
  const options = readDataObject(
    value,
    `${context} options`,
    ['border', 'fill', 'margin', 'valign'],
  );
  const borders = normalizeTableCellBorders(options.border, `${context} border`);
  const fill = normalizeTableCellFill(options.fill, `${context} fill`);
  const margins = normalizeTextBoxMargins(
    options.margin as TextBoxMarginInput | undefined,
    `${context} margin`,
  );
  const verticalAlignment = options.valign === undefined
    ? undefined
    : normalizeTextBoxVerticalAlignment(options.valign, `${context} valign`);
  return {
    ...(borders === undefined ? {} : { borders }),
    ...(fill === undefined ? {} : { fill }),
    ...(margins === undefined ? {} : { margins }),
    ...(verticalAlignment === undefined ? {} : { verticalAlignment }),
  };
}
```

Return the defined field and render it after margin attributes:

```ts
const verticalAlignmentAttribute = renderTableCellVerticalAlignmentAttribute(
  cell.verticalAlignment,
);
return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody><a:tcPr${marginAttributes}${verticalAlignmentAttribute}>${borders}${fill}</a:tcPr></a:tc>`;
```

- [ ] **Step 4: Run focused regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: all internal creation and existing alignment editor tests pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Review exact attribute order, omitted bytes, strict input, no bodyPr anchor, and unchanged editor code. Stage only the three files, commit:

```text
feat: render table cell valign during creation
```

Push and fetch through the verified SSH-over-443 channel, then require:

```sh
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; `.pnpm-store/` remains untracked and unstaged.

---

### Task 2: Expose the public creation option and prove the model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 runtime creation path and existing table-cell snapshot/editor.
- Produces: public `AddTableCellOptions.valign?: TextBoxVerticalAlignment` with create/read/edit/write/reopen evidence.

- [ ] **Step 1: Add public typing and lifecycle assertions first**

Extend the existing native table creation model test so a typed input contains:

```ts
const sourceCellOptions: AddTableCellOptions = {
  valign: 'middle',
  margin: sourceNamedMargin,
  border: sourceBorder,
  fill: sourceFill,
};
```

Include cells for omitted, top, middle, bottom, and `valign: undefined`. Require immediate snapshots:

```ts
expect(table.rows.map(({ cells }) =>
  cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual([
  ['middle', 'top', 'bottom'],
  [undefined, undefined, undefined],
]);
```

After creation, call:

```ts
table.setCellVerticalAlignment(0, 0, 'bottom');
table.setCellVerticalAlignment(0, 1, undefined);
expect(table.rows[0]!.cells.map(
  ({ verticalAlignment }) => verticalAlignment)).toEqual([
  'bottom',
  undefined,
  'bottom',
]);
```

Retain the existing duplicate/write/reopen assertions and require the edited direct state, text, margins, borders, fill, geometry, widths, and heights to survive.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: excess-property error for `AddTableCellOptions.valign`.

- [ ] **Step 3: Widen only the public cell option**

In `slide.ts`:

```ts
export interface AddTableCellOptions {
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly valign?: TextBoxVerticalAlignment;
}
```

Do not change `AddTableOptions` or add a table-level default here.

- [ ] **Step 4: Run model typecheck and lifecycle tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/model/src/table-create.internal.test.ts
```

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Review the public declaration, lifecycle, existing editor isolation, and absence of table-level expansion. Stage only both files, commit:

```text
feat: expose table cell valign creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 3: Cover the SDK lifecycle and invalid public input

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 2 public cell creation surface.
- Produces: SDK transaction, duplicate, rollback, invalid-input, and reopen evidence.

- [ ] **Step 1: Add all three values to the created-table SDK fixture**

Give the existing 2x3 or 3x3 created table these explicit values while preserving its fill/border/margin combinations:

```ts
{ text: 'Region', options: {
  valign: 'top',
  margin: { top: 4, left: 8 },
  border: sourceBorder,
  fill: sourceFill,
} }
{ text: 'Revenue', options: { valign: 'middle' } }
{ text: 'East', options: { valign: 'bottom' } }
{ text: '', options: { valign: undefined } }
```

Assert immediate snapshots are top/middle/bottom/undefined and that detached source/options changes cannot affect them.

- [ ] **Step 2: Exercise immediate edit, duplicate, rollback, and reopen**

Duplicate before edits. Change top→bottom, clear middle, and set the omitted cell to top. Edit text/margins/borders/fill and row/column sizes after alignment changes. In an outer `document.transaction()`, change two alignments and throw; require exact bytes, journal, live `TableModel` identity, and snapshots to roll back.

Write/reopen both original and duplicate. Require edited and duplicate alignment matrices, all cell formatting, text, geometry, widths, and heights to remain isolated and correct.

- [ ] **Step 3: Reject invalid public values without mutation**

For each value below, call public `slide.addTable()` with `{ text: 'Invalid', options: { valign: value as never } }`:

```ts
[
  null,
  false,
  true,
  0,
  '',
  'Top',
  ' top ',
  'mid',
  'center',
  'just',
  'dist',
  'distributed',
  [],
  {},
  Symbol('top'),
]
```

Require `TypeError`, exact target slide bytes, unchanged mutation journal, slide/shape counts, existing table identity, and all fresh snapshots.

- [ ] **Step 4: Run focused SDK and model tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only `packages/sdk/src/index.test.ts`, commit:

```text
test: cover sdk table valign creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 conformance and intentional differences

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: Tasks 1–3 native runtime and the installed PptxGenJS 4.0.1 public writer.
- Produces: final direct-state equality for supported cell-level `valign`.

- [ ] **Step 1: Add native and PptxGenJS tables with equivalent direct values**

Create a PptxGenJS table through only public APIs:

```ts
const generated = new PptxGenJS();
expect(generated.version).toBe('4.0.1');
generated.layout = 'LAYOUT_WIDE';
generated.addSlide().addTable([[
  { text: 'Default', options: {} },
  { text: 'Top', options: { valign: 'top' } },
  { text: 'Middle', options: { valign: 'middle' } },
  { text: 'Bottom', options: { valign: 'bottom' } },
]], {
  x: 0.5,
  y: 0.5,
  w: 8,
  h: 1,
  colW: [2, 2, 2, 2],
  rowH: 1,
});
```

Create the native equivalent with `columnWidths: inches(2)` and `rowHeights: inches(1)`. Import the PptxGenJS output and require both snapshot arrays to equal:

```ts
[undefined, 'top', 'middle', 'bottom']
```

Require equal text, transform, widths, heights, and canonical direct margins.

- [ ] **Step 2: Assert exact direct anchors and strict difference**

For native and imported slide XML, require one each of `anchor="t"`, `anchor="ctr"`, and `anchor="b"`, and no `bodyPr@anchor`. Require the default cell's `tcPr` to omit anchor. Write/reopen and compare snapshots.

Generate one PptxGenJS runtime-invalid cell with `valign: 'mid' as never` and require its imported strict snapshot to be `undefined` while raw XML preserves `anchor="mid"`. Native creation with the same runtime value must throw `TypeError` before mutation.

- [ ] **Step 3: Run adapter and native model suites**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Review public-only PptxGenJS generation, explicit column array, direct ownership, omitted absence, and invalid difference. Stage only the adapter test, commit:

```text
test: compare table valign creation with pptxgenjs
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 5: Smoke the actual Node, browser, declaration, and CLI package surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: packed `@jiayunxie/pptx` Node/browser/type/CLI entry points.
- Produces: real tarball evidence for `AddTableCellOptions.valign`.

- [ ] **Step 1: Extend Node package smoke**

Add `valign: 'top'`, `'middle'`, and `'bottom'` to three existing created-table cells and leave the fourth omitted. Capture the initial snapshot matrix and require:

```js
[
  ['top', 'middle'],
  ['bottom', undefined],
]
```

After existing formatting checks, call:

```js
createdTable.setCellVerticalAlignment(0, 0, 'bottom');
```

Write/reopen and require bottom on the edited cell, middle/bottom on neighbors, omitted on the fourth cell, and all existing margin/border/fill checks unchanged. Add `tableCellVerticalAlignmentCreation: true` to the JSON API result.

- [ ] **Step 2: Extend browser and declaration smoke**

Mirror the Node assertions in the browser bundle path. In the declaration fixture, set:

```ts
const creationOptions: AddTableCellOptions = {
  border: creationBorder,
  fill: cellFill,
  margin: creationMargin,
  valign: cellAlignment,
};
```

The existing `cellAlignment: TextBoxVerticalAlignment = 'middle'` proves the public type without adding a duplicate union.

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
valign_package_dir=$(mktemp -d /tmp/pptx-table-cell-valign-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$valign_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$valign_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require every smoke flag true, including `tableCellVerticalAlignmentCreation`, `types`, and CLI `0.1.0`.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the smoke script, commit:

```text
test: smoke packed table valign creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 6: Document the public contract and unsupported boundary

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: completed public/runtime/package evidence.
- Produces: accurate supported and unsupported documentation.

- [ ] **Step 1: Update API examples and contract**

Add `valign: 'middle'` to the existing table cell example. Document:

- exact public values and `valign` / `verticalAlignment` naming;
- direct `tcPr@anchor` ownership and `t/ctr/b` mapping;
- omitted/undefined byte preservation and no effective-default synthesis;
- attribute/child order with margins, borders, and fill;
- immediate editing through `setCellVerticalAlignment()`.

- [ ] **Step 2: Update compatibility and unsupported lists**

Change cell creation option shapes from `{ border, fill, margin }` to `{ border, fill, margin, valign }`. Describe supported cell-level PptxGenJS final-state equality and invalid passthrough difference. Keep table-level `valign`, horizontal alignment, direction/fit creation, merge, hyperlink, rich text, auto-page, and layout recomputation listed as unsupported.

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'border/fill/margin(?!/valign)|\{ border, fill, margin \}|cell options.*vertical alignment.*unsupported' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
```

Review any long-line hits in context rather than treating them mechanically as errors.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the four documents, commit:

```text
docs: document table cell valign creation
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 7: Run full gates and real-deck QA

**Files:**
- Review every Task 1–6 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: all implementation/tests/docs, actual packed package, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces: a verified set of already-pushed commits; QA fixes, if any, receive their own review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-valign-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-cell-valign-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Require zero failed suites/tests and a passing performance gate. Do not change repository timeouts to hide host load.

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
valign_qa_package_dir=$(mktemp -d /tmp/pptx-table-cell-valign-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$valign_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$valign_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
```

Require every smoke flag true, offline doctor success, and CLI `0.1.0`.

- [ ] **Step 3: Generate six real PPTX files with the SDK under test**

Create `/tmp/pptx-table-cell-valign-qa.mjs` with `apply_patch` using this program:

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
const output = '/tmp/pptx-table-cell-valign';
await mkdir(output, { recursive: true });

const textRows = [
  ['Default', 'Top', 'Middle', 'Bottom'],
  ['Top combined', 'Middle fill', 'Bottom border', 'Omitted'],
];
const tableOptions = {
  name: 'Cell valign creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(1.5), inches(2.5), inches(2), inches(2)],
  rowHeights: [inches(0.8), inches(1.6)],
};

function build(rows) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const slide = document.addSlide();
  return { document, table: slide.addTable(rows, tableOptions) };
}

const strings = build(textRows);
const undefinedValign = build(textRows.map((row) => row.map((text) => ({
  text,
  options: { valign: undefined },
}))));
await writeFile(output + '/string-source.pptx', await strings.document.write());
await writeFile(
  output + '/undefined-valign-source.pptx',
  await undefinedValign.document.write(),
);

const rows = [
  [
    { text: 'Default' },
    { text: 'Top', options: {
      valign: 'top',
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    } },
    { text: 'Middle', options: {
      valign: 'middle',
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
    } },
    { text: 'Bottom', options: { valign: 'bottom' } },
  ],
  [
    { text: 'Top combined', options: {
      valign: 'top',
      margin: { top: 4, left: 18 },
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
    } },
    { text: 'Middle fill', options: {
      valign: 'middle',
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
    { text: 'Omitted' },
  ],
];

const source = build(rows);
const sourceAlignments = source.table.rows.map(({ cells }) =>
  cells.map(({ verticalAlignment }) => verticalAlignment));
await writeFile(output + '/valign-source.pptx', await source.document.write());

source.table.setCellVerticalAlignment(0, 1, 'bottom');
source.table.setCellText(0, 1, 'Edited to bottom');
const editedBytes = await source.document.write();
await writeFile(output + '/valign-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[0].cells[1].text, 'Edited to bottom');
assert.equal(reopenedTable.rows[0].cells[1].verticalAlignment, 'bottom');
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/valign-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      { text: 'Default', options: {} },
      { text: 'Top', options: {
        valign: 'top',
        fill: { color: 'D9EAF7' },
      } },
      { text: 'Middle', options: {
        valign: 'middle',
        border: { type: 'solid', color: 'C00000', pt: 2 },
      } },
      { text: 'Bottom', options: { valign: 'bottom' } },
    ],
    [
      { text: 'Top combined', options: {
        valign: 'top',
        margin: [4, 7.2, 3.6, 18],
        fill: { color: baseline.SchemeColor.accent1, transparency: 25 },
      } },
      { text: 'Middle fill', options: {
        valign: 'middle',
        fill: { color: 'FFF2CC' },
      } },
      { text: 'Bottom border', options: {
        valign: 'bottom',
        border: [
          { type: 'none' },
          { type: 'none' },
          { type: 'dash', color: '70AD47', pt: 3 },
          { type: 'none' },
        ],
      } },
      { text: 'Omitted', options: {} },
    ],
  ],
  {
    x: 0.75,
    y: 0.75,
    w: 8,
    h: 2.4,
    colW: [1.5, 2.5, 2, 2],
    rowH: [0.8, 1.6],
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

Run it and require source alignments `[[null,"top","middle","bottom"],["top","middle","bottom",null]]` in JSON, edited bottom after reopen, widths `[1371600,2286000,1828800,1828800]`, and heights `[731520,1463040]`.

- [ ] **Step 4: Validate packages and exact diff isolation**

Validate these six decks with the repository CLI and `--profile powerpoint-2010`:

```text
string-source.pptx
undefined-valign-source.pptx
valign-source.pptx
valign-edited.pptx
valign-reopened.pptx
pptxgenjs-baseline.pptx
```

Require zero errors/warnings. Run package diffs:

```sh
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-valign/string-source.pptx \
  /tmp/pptx-table-cell-valign/undefined-valign-source.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-valign/valign-source.pptx \
  /tmp/pptx-table-cell-valign/valign-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-valign/valign-edited.pptx \
  /tmp/pptx-table-cell-valign/valign-reopened.pptx
```

Require zero changed parts, only `/ppt/slides/slide1.xml`, and zero changed parts respectively.

- [ ] **Step 5: Render and inspect native/baseline decks**

Use isolated LibreOffice profiles to convert `valign-source.pptx`, `valign-edited.pptx`, and `pptxgenjs-baseline.pptx` to PDF. Rasterize every page with Poppler at 180 DPI. Run overflow checks on edited and baseline:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-valign/valign-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-valign/pptxgenjs-baseline.pptx
```

Inspect every PNG at original detail. Require visibly distinct top/middle/bottom positions, edited top→bottom movement, preserved margins/borders/fills, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict value normalization, descriptor-safe options, direct tcPr ownership, canonical attribute order, omitted bytes, public declarations, SDK rollback/isolation, PptxGenJS differences, tarball flags, docs boundaries, and invalid-input no-mutation assertions.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Expected final output is `0 0` and only `?? .pnpm-store/`. If QA reveals a defect, fix only that defect, run its focused/full gates, review, commit, push, fetch, and repeat this step.

---

After synchronization, design and implement native table-level `valign` creation/default propagation as the next independently reviewable PptxGenJS parity item without asking for routine decisions.
