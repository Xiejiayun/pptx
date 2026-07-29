# Table Cell Horizontal Alignment Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user requires autonomous inline execution without routine questions or subagents.

**Goal:** Let native `slide.addTable()` create single-paragraph cells with strict left/center/right/justify direct horizontal alignment through `{ text, options: { align } }`, with PptxGenJS 4.0.1 final-state parity.

**Architecture:** Reuse the existing public `TextAlignment`, `normalizeTextAlignment()`, and `renderRichTextParagraphs(..., { defaultAlign })`. Carry one optional detached alignment through the normalized cell definition and render it only as the created paragraph's direct `a:pPr@algn`; do not add table-level propagation, a snapshot, or an editor in this slice.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, artifact-tool rendering, Poppler, and the presentation overflow checker.

## Global Constraints

- Public creation property is `AddTableCellOptions.align?: TextAlignment` with exact values `left`, `center`, `right`, and `justify`.
- Storage is the created cell's only paragraph direct `a:p/a:pPr@algn`; never write alignment to `a:tcPr` or `a:bodyPr`.
- Omitted/undefined `align`, string cells, `{ text }`, and empty cell options retain current bytes and direct alignment absence.
- `left` / `center` / `right` / `justify` render canonical `l` / `ctr` / `r` / `just` through the existing rich-text renderer.
- Reuse the existing strict normalizer and renderer; do not copy token maps, add aliases, coerce values, case-fold, or trim whitespace.
- Existing table text, border, fill, margin, vertical-alignment, direction, fit, geometry, identity, and transaction semantics remain unchanged.
- PptxGenJS invalid runtime fallback is documented and tested, not copied into native behavior.
- Table-level `align`, existing-cell alignment snapshot/editor, rich or multi-paragraph cells, merge, hyperlink, auto-page, repeated headers, and layout recomputation remain outside this slice.
- Every successful small item is reviewed, committed, pushed over SSH port 443, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only gates do not create empty commits.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and render cell horizontal alignment

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing `TextAlignment`, `normalizeTextAlignment()`, `normalizeRichText()`, and `renderRichTextParagraphs()`.
- Produces: normalized `alignment?: TextAlignment` and canonical direct `pPr@algn` output for created table cells.

- [ ] **Step 1: Add failing internal normalization and exact-output tests**

Add a focused test adjacent to the cell vertical-alignment creation test:

```ts
it('normalizes and renders strict table cell horizontal alignment', () => {
  const nullOptions = Object.assign(Object.create(null), { align: 'right' });
  const rows = [[
    'String',
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { align: undefined } },
    { text: 'Left', options: { align: 'left' } },
    { text: 'Center', options: { align: 'center' } },
    { text: 'Right', options: { align: 'right' } },
    { text: 'Justify this sentence', options: { align: 'justify' } },
    { text: 'Null prototype', options: nullOptions },
    { text: 'Combined', options: {
      align: 'center',
      valign: 'bottom',
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
  expect(definition.rows[0]!.map(({ alignment }) => alignment)).toEqual([
    undefined,
    undefined,
    undefined,
    'left',
    'center',
    'right',
    'justify',
    'right',
    'center',
  ]);

  const equivalent = [
    [['Same']],
    [[{ text: 'Same' }]],
    [[{ text: 'Same', options: {} }]],
    [[{ text: 'Same', options: { align: undefined } }]],
  ].map((input) => renderTableGraphicFrame(
    20,
    normalizeTableDefinition(input, undefined),
  ));
  expect(new Set(equivalent).size).toBe(1);
  expect(equivalent[0]).not.toContain(' algn=');

  const renderCell = (cell: unknown): string => renderTableGraphicFrame(
    21,
    normalizeTableDefinition([[cell]], undefined),
  );
  expect(renderCell(rows[0]![3])).toContain('<a:pPr algn="l"/>');
  expect(renderCell(rows[0]![4])).toContain('<a:pPr algn="ctr"/>');
  expect(renderCell(rows[0]![5])).toContain('<a:pPr algn="r"/>');
  expect(renderCell(rows[0]![6])).toContain('<a:pPr algn="just"/>');
  const combined = renderCell(rows[0]![8]);
  expect(combined).toContain('<a:p><a:pPr algn="ctr"/><a:r>');
  expect(combined).toContain('marB="45720" anchor="b">');
  expect(combined).not.toMatch(/<a:tcPr[^>]*\salgn=/);
  expect(combined).not.toMatch(/<a:bodyPr[^>]*\salgn=/);
});
```

Extend the malformed-cell test with a getter-backed `align` and these invalid values:

```ts
const accessorAlignOptions = {};
Object.defineProperty(accessorAlignOptions, 'align', {
  get() {
    cellAccessorCalls += 1;
    return 'center';
  },
  enumerable: true,
  configurable: true,
});

const invalidAlignments = [
  null,
  false,
  true,
  0,
  '',
  'Left',
  ' center ',
  'l',
  'ctr',
  'r',
  'just',
  'dist',
  'thaiDist',
  'justLow',
  [],
  {},
  Symbol('center'),
];
```

Add `{ text: 'A', options: accessorAlignOptions }` and every invalid item wrapped as `{ text: 'A', options: { align } }` to `invalidCells`. Keep `expect(cellAccessorCalls).toBe(0)`.

- [ ] **Step 2: Run the focused internal suite and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "horizontal alignment|malformed matrices"
```

Expected: `align` is rejected as an unsupported cell option or normalized cells lack `alignment`.

- [ ] **Step 3: Add the minimal normalized field and renderer wiring**

In `table-create.internal.ts`, widen the existing imports:

```ts
import {
  normalizeRichText,
  normalizeTextAlignment,
  renderRichTextParagraphs,
} from './rich-text.internal.js';
import type {
  TextAlignment,
  TextBoxMarginInput,
  TextBoxMargins,
  TextBoxVerticalAlignment,
} from './text.js';
```

Extend the internal cell and cell option normalizer:

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly alignment?: TextAlignment;
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
  'alignment' | 'borders' | 'fill' | 'margins' | 'verticalAlignment'
> {
  if (value === undefined) return {};
  const options = readDataObject(
    value,
    `${context} options`,
    ['align', 'border', 'fill', 'margin', 'valign'],
  );
  const alignment = options.align === undefined
    ? undefined
    : normalizeTextAlignment(options.align, `${context} align`);
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
    ...(alignment === undefined ? {} : { alignment }),
    ...(borders === undefined ? {} : { borders }),
    ...(fill === undefined ? {} : { fill }),
    ...(margins === undefined ? {} : { margins }),
    ...(verticalAlignment === undefined ? {} : { verticalAlignment }),
  };
}
```

Pass only the normalized primitive to the existing paragraph renderer:

```ts
const paragraphs = renderRichTextParagraphs(normalizeRichText([
  { runs: [{ text: cell.text, style: {} }] },
]), { defaultAlign: cell.alignment });
```

Do not add a new token map, codec file, or paragraph mutation path.

- [ ] **Step 4: Run focused regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: all table creation and existing paragraph-alignment tests pass.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review strict values, getter-free input, omitted-byte equality, paragraph ownership, renderer reuse, and the absence of public API expansion. Stage only both files and commit:

```text
feat: render table cell horizontal alignment during creation
```

Then run:

```sh
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git HEAD:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; `.pnpm-store/` remains untracked and unstaged.

---

### Task 2: Expose the public cell creation option and prove model preservation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 runtime creation path and existing exported `TextAlignment`.
- Produces: public `AddTableCellOptions.align?: TextAlignment` plus create/edit/duplicate/write/reopen raw-OOXML evidence.

- [ ] **Step 1: Add public typing and lifecycle assertions first**

In the existing typed native table creation test, add alignment to the source options and row matrix:

```ts
const sourceCellOptions: AddTableCellOptions = {
  align: 'center',
  border: typedSourceBorder,
  fill: sourceFill,
  margin: sourceNamedMargin,
  valign: 'middle',
};
```

Give the six existing cells center/left/right/justify/undefined/omitted values. After creation, decode the target slide XML and extract cells in physical order:

```ts
const tableXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
const cellXml = [...tableXml.matchAll(
  /<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g,
)].map((match) => match[0]);
const directAlignment = (xml: string): string | undefined =>
  xml.match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1];
expect(cellXml.map(directAlignment)).toEqual([
  'ctr',
  'l',
  'r',
  'just',
  undefined,
  undefined,
]);
```

Edit text, fill, border, margin, vertical alignment, row heights, and column widths through existing APIs, then require the same six direct alignment tokens. Duplicate before target edits; require original and duplicate XML to preserve independent token arrays. After write/reopen, require the same tokens, text, geometry, widths, heights, and existing snapshots.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: excess-property error for `AddTableCellOptions.align`.

- [ ] **Step 3: Widen only the public cell option**

In `slide.ts`:

```ts
export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly valign?: TextBoxVerticalAlignment;
}
```

`TextAlignment` is already imported for `AddTextOptions`; do not change `AddTableOptions` or add a snapshot/editor here.

- [ ] **Step 4: Run model typecheck and lifecycle tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/model/src/table-create.internal.test.ts
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review the public declaration, raw token lifecycle, non-alignment formatting, duplicate isolation, and absence of table-level expansion. Stage only both files, commit:

```text
feat: expose table cell horizontal alignment creation
```

Push/fetch through SSH port 443 and require `origin/main...HEAD = 0 0`.

---

### Task 3: Cover the SDK lifecycle and invalid public input

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 2 public creation surface.
- Produces: SDK transaction, duplicate, rollback, invalid-input, non-target isolation, and reopen evidence.

- [ ] **Step 1: Add a focused public SDK creation test**

Create a 2x3 table with combined formatting:

```ts
const table = slide.addTable([
  [
    { text: 'Left', options: {
      align: 'left',
      valign: 'top',
      margin: { top: 4, left: 8 },
      border: sourceBorder,
      fill: sourceFill,
    } },
    { text: 'Center', options: { align: 'center', valign: 'middle' } },
    { text: 'Right', options: { align: 'right', valign: 'bottom' } },
  ],
  [
    { text: 'Justify this longer sentence', options: { align: 'justify' } },
    { text: 'Undefined', options: { align: undefined } },
    'Omitted',
  ],
], {
  name: 'Cell horizontal alignment creation',
  columnWidths: [inches(1.5), inches(2.5), inches(2)],
  rowHeights: [inches(0.75), inches(1.25)],
});
```

Use a local raw-XML helper to assert `['l', 'ctr', 'r', 'just', undefined, undefined]`. Mutate source formatting objects and require no effect on the created table.

- [ ] **Step 2: Exercise text/property edits, duplicate, rollback, and reopen**

Duplicate before edits. On the original, edit the four aligned cells' text, margins, borders, fills, vertical alignment, column widths, and row heights. Require all six alignment tokens unchanged.

In an outer `document.transaction()`, add a second aligned table and mutate the original, then throw. Require exact bytes, mutation journal, slide/shape counts, live `TableModel` identity, and both raw token arrays to roll back.

Write/reopen original and duplicate. Require target and duplicate tokens, text, non-alignment formatting, geometry, widths, and heights to remain isolated and correct. Require a non-target custom part fingerprint to remain unchanged.

- [ ] **Step 3: Reject invalid public values without mutation**

For each value below, call public `slide.addTable()` with `{ text: 'Invalid', options: { align: value as never } }`:

```ts
[
  null,
  false,
  true,
  0,
  '',
  'Left',
  ' center ',
  'l',
  'ctr',
  'r',
  'just',
  'dist',
  'thaiDist',
  'justLow',
  [],
  {},
  Symbol('center'),
]
```

Require `TypeError`, exact target slide bytes, unchanged mutation journal, slide/shape counts, existing table identity, and existing raw alignment tokens. Add a getter-backed `align` options object and require getter count zero.

- [ ] **Step 4: Run focused SDK and model tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Stage only `packages/sdk/src/index.test.ts`, commit:

```text
test: cover sdk table cell horizontal alignment creation
```

Push/fetch through SSH port 443 and require `origin/main...HEAD = 0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 conformance and strict differences

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: Tasks 1–3 native runtime and installed PptxGenJS 4.0.1 public writer.
- Produces: final direct paragraph-alignment equality for supported cell-level values.

- [ ] **Step 1: Add native and PptxGenJS tables with equivalent direct values**

Create a PptxGenJS table through public APIs only:

```ts
const generated = new PptxGenJS();
expect(generated.version).toBe('4.0.1');
generated.layout = 'LAYOUT_WIDE';
generated.addSlide().addTable([[
  { text: 'Default', options: {} },
  { text: 'Left', options: { align: 'left' } },
  { text: 'Center', options: { align: 'center' } },
  { text: 'Right', options: { align: 'right' } },
  { text: 'Justify this sentence', options: { align: 'justify' } },
]], {
  x: 0.5,
  y: 0.5,
  w: 10,
  h: 1,
  colW: [2, 2, 2, 2, 2],
  rowH: 1,
  margin: 0.1,
  valign: 'middle',
});
```

Create the native equivalent with `columnWidths: inches(2)`, `rowHeights: inches(1)`, `margin: 7.2`, and `valign: 'middle'`. Import PptxGenJS output and decode both slide parts.

- [ ] **Step 2: Assert exact direct tokens, omission, and strict invalid behavior**

Extract each cell's first direct `pPr@algn`. Require both native and imported arrays to equal:

```ts
[undefined, 'l', 'ctr', 'r', 'just']
```

Require equal cell text, transform, widths, heights, direct margins, vertical alignment, borders, and fills. Require no `algn` on `tcPr` or `bodyPr`. Write/reopen and compare the same token arrays.

Generate a PptxGenJS runtime-invalid cell with `align: 'dist' as never`; require final XML to omit `pPr@algn`. Native creation with the same runtime value must throw `TypeError` before mutation. This is an intentional strictness difference, not parity debt.

- [ ] **Step 3: Run adapter and native model suites**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Review public-only PptxGenJS generation, explicit geometry, paragraph ownership, omitted absence, strict invalid difference, and the lack of whole-slide byte comparison. Stage only the adapter test, commit:

```text
test: compare table cell horizontal alignment with pptxgenjs
```

Push/fetch through SSH port 443 and require `origin/main...HEAD = 0 0`.

---

### Task 5: Smoke the actual Node, browser, declaration, and CLI surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: packed `@jiayunxie/pptx` Node/browser/type/CLI entry points.
- Produces: actual tarball evidence for `AddTableCellOptions.align`.

- [ ] **Step 1: Extend Node package smoke**

Add `align: 'left'`, `'center'`, `'right'`, and `'justify'` to the existing four created-table cells. Capture `createdTableXml` before later edits, extract physical cells, and require direct alignment tokens `['l', 'ctr', 'r', 'just']`.

After existing text/fill/border/margin/vertical-alignment/size edits and write/reopen, decode the reopened table slide and require the same token array. Add `tableCellHorizontalAlignmentCreation: true` to the JSON API result, chained after the existing table vertical-alignment creation flag.

- [ ] **Step 2: Extend browser and declaration smoke**

Mirror the four cell values and raw XML assertion in the browser bundle path. In the declaration fixture, reuse the exported type:

```ts
const cellHorizontalAlignment: TextAlignment = 'center';
const creationOptions: AddTableCellOptions = {
  align: cellHorizontalAlignment,
  border: creationBorder,
  fill: cellFill,
  margin: creationMargin,
  valign: cellAlignment,
};
```

No new type export is needed because `TextAlignment` already ships.

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
align_package_dir=$(mktemp -d /tmp/pptx-table-cell-align-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$align_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$align_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require every smoke flag true, including `tableCellHorizontalAlignmentCreation`, `types`, browser, and CLI `0.1.0`.

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Stage only the smoke script, commit:

```text
test: smoke packed table cell horizontal alignment
```

Push/fetch through SSH port 443 and require `origin/main...HEAD = 0 0`.

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

Add `align: 'center'` to the existing table cell example. Document:

- exact public values and reuse of `TextAlignment`;
- direct `a:pPr@algn` ownership and `l/ctr/r/just` mapping;
- omitted/undefined byte preservation and no effective-left synthesis;
- preservation through existing text/property edits, duplicate, write, and reopen;
- the absence of a cell snapshot/editor in this slice.

- [ ] **Step 2: Update compatibility and unsupported lists**

Change cell creation option shapes from `{ border, fill, margin, valign }` to `{ align, border, fill, margin, valign }`. Mark cell-level horizontal alignment creation as supported and PptxGenJS-final-state compatible. Keep table-level `align`, existing-deck horizontal-alignment snapshot/editor, direction/fit creation, merge, hyperlink, rich text, auto-page, repeated headers, and layout recomputation unsupported.

Document the invalid-runtime difference: PptxGenJS silently drops unknown alignment; native creation rejects it.

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'cell options.*horizontal alignment.*unsupported|\{ border, fill, margin, valign \}|border/fill/margin/valign(?!/align)' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
```

Review hits in context; retain statements about table-level alignment and existing-deck editing.

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Stage only the four documents, commit:

```text
docs: document table cell horizontal alignment creation
```

Push/fetch through SSH port 443 and require `origin/main...HEAD = 0 0`.

---

### Task 7: Run full gates and real-deck QA

**Files:**
- Review every Task 1–6 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: implementation/tests/docs, actual packed package, repository CLI, PptxGenJS, artifact-tool renderer, Poppler, and overflow checker.
- Produces: a verified set of already-pushed commits; any defect fix receives its own review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-align-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-cell-align-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Require zero failed suites/tests and a passing performance gate. Do not weaken repository timeouts.

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
align_qa_package_dir=$(mktemp -d /tmp/pptx-table-cell-align-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$align_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$align_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
```

Require every smoke flag true, offline doctor success, and CLI `0.1.0`.

- [ ] **Step 3: Generate real native and PptxGenJS decks**

Create `/tmp/pptx-table-cell-align-qa.mjs` with `apply_patch` using this complete program:

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
const output = '/tmp/pptx-table-cell-align';
await mkdir(output, { recursive: true });

const plainRows = [
  ['Default', 'Left', 'Center', 'Right', 'Justify'],
  ['Justify second', 'Right second', 'Center second', 'Left second', 'Default second'],
];
const tableOptions = {
  name: 'Cell horizontal alignment QA',
  x: inches(0.5),
  y: inches(0.75),
  columnWidths: [
    inches(1.25),
    inches(1.75),
    inches(2),
    inches(2.25),
    inches(2.75),
  ],
  rowHeights: [inches(0.8), inches(1.6)],
};

function build(rows) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const slide = document.addSlide();
  const table = slide.addTable(rows, tableOptions);
  return { document, slide, table };
}

function cellXml(document, slide) {
  const xml = new TextDecoder().decode(
    document.opcPackage.requirePart(slide.partUri).bytes,
  );
  return [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
    .map((match) => match[0]);
}

function tokenMatrix(document, slide) {
  const tokens = cellXml(document, slide).map((xml) =>
    xml.match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
  return [tokens.slice(0, 5), tokens.slice(5, 10)];
}

const expected = [
  [undefined, 'l', 'ctr', 'r', 'just'],
  ['just', 'r', 'ctr', 'l', undefined],
];

const omitted = build(plainRows);
const runtimeUndefined = build(plainRows.map((row) => row.map((text) => ({
  text,
  options: { align: undefined },
}))));
await writeFile(output + '/omitted.pptx', await omitted.document.write());
await writeFile(output + '/undefined.pptx', await runtimeUndefined.document.write());

const nativeRows = [
  [
    { text: 'Default' },
    { text: 'Left', options: { align: 'left', valign: 'top' } },
    { text: 'Center', options: {
      align: 'center',
      valign: 'middle',
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 20,
      },
    } },
    { text: 'Right', options: {
      align: 'right',
      valign: 'bottom',
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
        style: 'solid',
      },
    } },
    { text: 'Justify a longer sentence across several words', options: {
      align: 'justify',
      valign: 'top',
      margin: 6,
    } },
  ],
  [
    { text: 'Justify another longer sentence across several words', options: {
      align: 'justify',
      valign: 'bottom',
    } },
    { text: 'Right second', options: {
      align: 'right',
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFF2CC' },
      },
    } },
    { text: 'Center second', options: {
      align: 'center',
      border: {
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 3,
          style: 'dash',
        },
      },
    } },
    { text: 'Left second', options: { align: 'left', margin: 10 } },
    'Default second',
  ],
];

const source = build(nativeRows);
const sourceTokens = tokenMatrix(source.document, source.slide);
assert.deepEqual(sourceTokens, expected);
await writeFile(output + '/align-source.pptx', await source.document.write());

source.table.setCellText(0, 2, 'Edited center');
source.table.setCellMargins(0, 2, { top: 4, left: 10 });
source.table.setCellFill(0, 2, {
  kind: 'solid',
  color: { kind: 'srgb', value: 'F4B183' },
  transparency: 25,
});
source.table.setCellBorders(0, 2, {
  left: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 4,
    style: 'solid',
  },
});
source.table.setCellVerticalAlignment(0, 2, 'bottom');
const editedTokens = tokenMatrix(source.document, source.slide);
assert.deepEqual(editedTokens, expected);
const editedBytes = await source.document.write();
await writeFile(output + '/align-edited.pptx', editedBytes);

const reopened = await PptxDocument.open(editedBytes);
const reopenedSlide = reopened.slides[0];
const reopenedTable = reopenedSlide.shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
const reopenedTokens = tokenMatrix(reopened, reopenedSlide);
assert.deepEqual(reopenedTokens, expected);
assert.equal(reopenedTable.rows[0].cells[2].text, 'Edited center');
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/align-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
assert.equal(baseline.version, '4.0.1');
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      { text: 'Default', options: {} },
      { text: 'Left', options: { align: 'left', valign: 'top' } },
      { text: 'Center', options: {
        align: 'center',
        valign: 'middle',
        fill: {
          color: baseline.SchemeColor.accent1,
          transparency: 20,
        },
      } },
      { text: 'Right', options: {
        align: 'right',
        valign: 'bottom',
        border: { type: 'solid', color: 'C00000', pt: 2 },
      } },
      { text: 'Justify a longer sentence across several words', options: {
        align: 'justify',
        valign: 'top',
        margin: 6,
      } },
    ],
    [
      { text: 'Justify another longer sentence across several words', options: {
        align: 'justify',
        valign: 'bottom',
      } },
      { text: 'Right second', options: {
        align: 'right',
        fill: { color: 'FFF2CC' },
      } },
      { text: 'Center second', options: {
        align: 'center',
        border: [
          { type: 'none' },
          { type: 'none' },
          { type: 'dash', color: '70AD47', pt: 3 },
          { type: 'none' },
        ],
      } },
      { text: 'Left second', options: { align: 'left', margin: 10 } },
      { text: 'Default second', options: {} },
    ],
  ],
  {
    x: 0.5,
    y: 0.75,
    w: 10,
    h: 2.4,
    colW: [1.25, 1.75, 2, 2.25, 2.75],
    rowH: [0.8, 1.6],
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceTokens,
  editedTokens,
  reopenedTokens,
  editedText: reopenedTable.rows[0].cells[2].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run it and require all three token matrices to match `expected`, edited text `Edited center`, widths `[1143000,1600200,1828800,2057400,2514600]`, and heights `[731520,1463040]`.

- [ ] **Step 4: Validate packages and exact diff isolation**

Validate all six decks with the repository CLI and `--profile powerpoint-2010`; require zero errors and warnings. Run package diffs:

```sh
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-align/omitted.pptx \
  /tmp/pptx-table-cell-align/undefined.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-align/align-source.pptx \
  /tmp/pptx-table-cell-align/align-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-align/align-edited.pptx \
  /tmp/pptx-table-cell-align/align-reopened.pptx
```

Require zero changed parts, only `/ppt/slides/slide1.xml`, and zero changed parts respectively.

- [ ] **Step 5: Render and inspect every deck**

Use the bundled runtime and presentation helpers:

```sh
RUNTIME_PY=/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3.12
PRESENTATION_SKILL=/Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations
for deck in /tmp/pptx-table-cell-align/*.pptx; do
  "$RUNTIME_PY" "$PRESENTATION_SKILL/container_tools/render_slides.py" "$deck"
  "$RUNTIME_PY" "$PRESENTATION_SKILL/container_tools/slides_test.py" "$deck"
done
```

If a single long loop is interrupted by the runtime, rerun each independent deck invocation and require exit 0. Inspect every `slide-1.png` at original detail. Require visibly distinct left/center/right alignment, justified long text, preserved vertical positions/margins/borders/fills, unequal rows/columns, all text, and no repair, clipping, unexpected wrapping, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict value normalization, descriptor-safe options, direct paragraph ownership, renderer reuse, omitted bytes, public declarations, SDK rollback/isolation, PptxGenJS difference, tarball flags, docs boundaries, and invalid-input no-mutation assertions.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Expected final output is `0 0` and only `?? .pnpm-store/`. If QA reveals a defect, fix only that defect, run focused/full gates, review, commit, push, fetch, and repeat this step. Pure QA success creates no empty commit.

---

After synchronization, design and implement native table-level `align` creation/default propagation as the next independently reviewable PptxGenJS parity item without asking for routine decisions.
