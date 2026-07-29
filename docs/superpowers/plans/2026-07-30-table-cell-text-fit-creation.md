# Table Cell Text-Fit Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let native `slide.addTable()` create plain-text cells with strict `none`/`shrink`/`resize` text fit through `{ text, options: { fit } }`, reusing the existing direct table-cell snapshot/editor without claiming a PptxGenJS table fit API.

**Architecture:** Extend the detached normalized table-cell definition with optional `textFit`, reuse `normalizeTextBoxFit()` and `renderTextBoxFitChild()`, and render the fit choice only inside each cell's direct `a:bodyPr`. Expose the input through `AddTableCellOptions.fit`; keep `AddTableOptions.fit`, dynamic font scaling, and layout recomputation out of scope.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarballs, `pptx-inspect`, LibreOffice headless, Poppler.

## Global Constraints

- Reuse `TextBoxFit = 'none' | 'shrink' | 'resize'`; do not add a second fit union.
- Public creation surface is only `AddTableCellOptions.fit?: TextBoxFit`; do not add `AddTableOptions.fit` in this item.
- Storage is the selected physical cell's direct `a:txBody/a:bodyPr` fit choice, never `a:tcPr`, table metadata, text length, or transform state.
- Omitted, runtime `undefined`, and `none` must preserve the existing self-closing `<a:bodyPr/>` bytes and snapshot as `undefined`.
- `shrink` and `resize` must write one direct `<a:normAutofit/>` and `<a:spAutoFit/>`, respectively.
- Never write `<a:noAutofit/>` during creation; existing explicit `noAutofit` remains readable as `none` through the existing editor model.
- Inputs must stay descriptor-safe, getter-free, ordinary/null-prototype-only, strict, and detached before package mutation.
- Do not alter table text, alignment, direction, margins, borders, fill, geometry, grid, rows, relationships, or identity outside the target fit behavior.
- PptxGenJS 4.0.1 has no table fit API and ignores runtime `fit`/`autoFit`/`shrinkText`; native shrink/resize creation is an intentional extension.
- Implement inline without subagent delegation, as required for this repository session.
- Never modify, delete, stage, or commit `.pnpm-store/`.
- Every task ends with review, a focused commit, SSH-over-443 push, fetch, and `origin/main...HEAD = 0 0`.

---

### Task 1: Normalize and render direct cell text fit

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Test: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing `normalizeTextBoxFit(value, context): TextBoxFit` and `renderTextBoxFitChild(value): string`.
- Produces: internal `NormalizedTableCell.textFit?: TextBoxFit` and exact direct `bodyPr` creation output.

- [ ] **Step 1: Add a focused failing normalization/render test**

Add this case beside the table-cell direction creation tests:

```ts
it('normalizes and renders strict table cell text fit', () => {
  const nullOptions = Object.assign(Object.create(null), { fit: 'resize' });
  const rows = [[
    'String',
    { text: 'Object' },
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { fit: undefined } },
    { text: 'None', options: { fit: 'none' } },
    { text: 'Shrink', options: { fit: 'shrink' } },
    { text: 'Resize', options: { fit: 'resize' } },
    { text: 'Null prototype', options: nullOptions },
    { text: 'Combined', options: {
      align: 'center',
      fit: 'shrink',
      margin: { top: 4, left: 8 },
      textDirection: 'vert270',
      valign: 'middle',
      border: { kind: 'none' },
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
    } },
  ]];
  const definition = normalizeTableDefinition(rows, undefined);

  expect(definition.rows[0]!.map(({ textFit }) => textFit)).toEqual([
    undefined,
    undefined,
    undefined,
    undefined,
    'none',
    'shrink',
    'resize',
    'resize',
    'shrink',
  ]);

  const equivalent = [
    [['Same']],
    [[{ text: 'Same' }]],
    [[{ text: 'Same', options: {} }]],
    [[{ text: 'Same', options: { fit: undefined } }]],
    [[{ text: 'Same', options: { fit: 'none' } }]],
  ].map((input) => renderTableGraphicFrame(
    70,
    normalizeTableDefinition(input, undefined),
  ));
  expect(new Set(equivalent).size).toBe(1);
  expect(equivalent[0]).toContain('<a:bodyPr/><a:lstStyle/>');

  const xml = renderTableGraphicFrame(71, definition);
  const cells = [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
    .map((match) => match[0]);
  expect(cells.map((cell) =>
    cell.match(/<a:(normAutofit|spAutoFit)\/>/)?.[1])).toEqual([
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    'normAutofit',
    'spAutoFit',
    'spAutoFit',
    'normAutofit',
  ]);
  expect(cells[5]).toContain(
    '<a:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>',
  );
  expect(cells[6]).toContain(
    '<a:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr><a:lstStyle/>',
  );
  expect(cells[8]).toMatch(
    /<a:bodyPr><a:normAutofit\/><\/a:bodyPr><a:lstStyle\/>[\s\S]*<a:tcPr marL="101600" marR="91440" marT="50800" marB="45720" anchor="ctr" vert="vert270">/,
  );
  expect(xml).not.toMatch(/<a:tcPr[^>]*\sfit=/);
  expect(xml).not.toContain('<a:noAutofit/>');

  nullOptions.fit = 'none';
  expect(definition.rows[0]![7]!.textFit).toBe('resize');
});
```

- [ ] **Step 2: Run the focused test and prove the unsupported boundary**

Run:

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts \
  -t 'normalizes and renders strict table cell text fit' --reporter=dot
```

Expected before implementation: FAIL because `fit` is rejected as an unsupported cell option or `textFit` is absent.

- [ ] **Step 3: Add strict invalid and getter-free cases**

In the existing adversarial normalization test, define:

```ts
const invalidFits = [
  null,
  false,
  true,
  0,
  '',
  'None',
  ' shrink',
  'resize ',
  'auto',
  [],
  {},
  Symbol('fit'),
];
```

Add every value as a cell option:

```ts
...invalidFits.map((fit) => ({ text: 'A', options: { fit } })),
```

Add an accessor and require zero calls:

```ts
let cellFitAccessorCalls = 0;
const accessorFitOptions: Record<string, unknown> = {};
Object.defineProperty(accessorFitOptions, 'fit', {
  get() {
    cellFitAccessorCalls += 1;
    return 'shrink';
  },
  enumerable: true,
});
```

Include `{ text: 'A', options: accessorFitOptions }` in `invalidCells` and finish with:

```ts
expect(cellFitAccessorCalls).toBe(0);
```

- [ ] **Step 4: Implement the minimum internal data flow**

In `table-create.internal.ts`, import the existing codec functions:

```ts
import {
  normalizeTextBoxFit,
  renderTextBoxFitChild,
} from './text-box-fit.internal.js';
```

Add `TextBoxFit` to the type imports and `textFit` to the normalized cell:

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly alignment?: TextAlignment;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}
```

Extend only the cell option allowlist and normalized pick:

```ts
['align', 'border', 'fill', 'fit', 'margin', 'textDirection', 'valign']
```

```ts
| 'textFit'
```

Normalize and conditionally return the value:

```ts
const textFit = options.fit === undefined
  ? undefined
  : normalizeTextBoxFit(options.fit, `${context} fit`);
```

```ts
...(textFit === undefined ? {} : { textFit }),
```

Do not add `fit` to table `OPTION_KEYS`.

- [ ] **Step 5: Render only the body-properties child**

In `renderTableCell()` add:

```ts
const textFitChild = cell.textFit === undefined
  ? ''
  : renderTextBoxFitChild(cell.textFit);
const bodyProperties = textFitChild === ''
  ? '<a:bodyPr/>'
  : `<a:bodyPr>${textFitChild}</a:bodyPr>`;
```

Use it in the existing return value:

```ts
return `<a:tc><a:txBody>${bodyProperties}<a:lstStyle/>${paragraphs}</a:txBody><a:tcPr${marginAttributes}${verticalAlignmentAttribute}${textDirectionAttribute}>${borders}${fill}</a:tcPr></a:tc>`;
```

- [ ] **Step 6: Run internal tests and review isolation**

Run:

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model/tsconfig.json --pretty false
git diff --check
```

Review that only cell `fit` is accepted, table `fit` remains rejected, none preserves bytes, shrink/resize own only `bodyPr`, no getter runs, and no unrelated renderer changes were introduced.

- [ ] **Step 7: Commit, push, fetch, and prove synchronization**

Stage only the implementation and internal test:

```sh
git add -- packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts
git commit -m 'feat: render table cell text fit during creation'
```

Push through SSH-over-443, fetch `main` into `origin/main`, and require `git rev-list --left-right --count origin/main...HEAD` to print `0 0`.

---

### Task 2: Expose the public creation option and prove lifecycle behavior

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 normalized `textFit` and direct renderer.
- Produces: `AddTableCellOptions.fit?: TextBoxFit` with live snapshot/edit/duplicate/rollback/reopen guarantees.

- [ ] **Step 1: Add a failing public SDK lifecycle test**

Add a focused case beside the existing basic-table creation test:

```ts
it('creates, edits, duplicates, rolls back, and reopens table cell text fit', async () => {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const slide = document.addSlide();
  const table = slide.addTable([[
    'String',
    { text: 'Omitted' },
    { text: 'Undefined', options: { fit: undefined } },
    { text: 'None', options: { fit: 'none' } },
    { text: 'Shrink', options: { fit: 'shrink' } },
    { text: 'Resize', options: {
      fit: 'resize',
      textDirection: 'vert270',
      valign: 'middle',
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
    } },
  ]], {
    columnWidths: inches(1.5),
    rowHeights: inches(1),
  });

  expect(table.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
    undefined,
    undefined,
    undefined,
    undefined,
    'shrink',
    'resize',
  ]);
  expect(table.rows[0]!.cells[5]!.textDirection).toBe('vert270');
  expect(table.rows[0]!.cells[5]!.verticalAlignment).toBe('middle');

  const sourceBytes = document.opcPackage.requirePart(slide.partUri).bytes;
  const sourceXml = new TextDecoder().decode(sourceBytes);
  expect(sourceXml.match(/<a:normAutofit\/>/g)).toHaveLength(1);
  expect(sourceXml.match(/<a:spAutoFit\/>/g)).toHaveLength(1);
  expect(sourceXml).not.toContain('<a:noAutofit/>');

  const duplicate = document.duplicateSlide(0);
  const duplicateTable = duplicate.shapes[0] as TableModel;
  expect(duplicateTable.rows[0]!.cells.map(({ textFit }) => textFit))
    .toEqual([undefined, undefined, undefined, undefined, 'shrink', 'resize']);

  table.setCellText(0, 4, 'Shrink edited');
  table.setCellTextDirection(0, 5, 'wordArtVert');
  table.setCellTextFit(0, 4, 'none');
  table.setCellTextFit(0, 3, 'resize');
  expect(table.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
    undefined,
    undefined,
    undefined,
    'resize',
    undefined,
    'resize',
  ]);

  const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
  const journal = [...document.opcPackage.mutations];
  expect(() => document.transaction(() => {
    table.setCellTextFit(0, 3, 'shrink');
    table.setCellTextFit(0, 5, undefined);
    throw new Error('restore created fits');
  })).toThrow('restore created fits');
  expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
  expect(document.opcPackage.mutations).toEqual(journal);

  const reopened = await PptxDocument.open(await document.write());
  const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
  expect(reopenedTable.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
    undefined,
    undefined,
    undefined,
    'resize',
    undefined,
    'resize',
  ]);
  expect(reopenedTable.rows[0]!.cells[4]!.text).toBe('Shrink edited');
  expect(reopenedTable.rows[0]!.cells[5]!.textDirection).toBe('wordArtVert');
  expect((reopened.slides[1]!.shapes[0] as TableModel).rows[0]!.cells[4]!.textFit)
    .toBe('shrink');
});
```

- [ ] **Step 2: Prove public typing fails before exposure**

Run:

```sh
node node_modules/typescript/bin/tsc -b packages/sdk/tsconfig.json --pretty false
```

Expected before the type change: FAIL because `fit` is not in `AddTableCellOptions`.

- [ ] **Step 3: Expose the exact public field**

In `packages/model/src/slide.ts` add one property in the existing interface:

```ts
export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly fit?: TextBoxFit;
  readonly margin?: TextBoxMarginInput;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}
```

- [ ] **Step 4: Add model-level atomic invalid-input coverage**

In the strict native table creation test, capture slide bytes, journal, shape count, and table identity, then add:

```ts
const beforeInvalidFit = pkg.requirePart(slide.partUri).bytes;
const invalidFitJournal = [...pkg.mutations];
const shapeCount = slide.shapes.length;
for (const fit of [null, false, 0, '', 'Shrink', ' shrink', 'auto', [], {}]) {
  expect(() => slide.addTable([[
    { text: 'Invalid fit', options: { fit: fit as never } },
  ]])).toThrow(TypeError);
}
expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalidFit);
expect(pkg.mutations).toEqual(invalidFitJournal);
expect(slide.shapes).toHaveLength(shapeCount);
expect(slide.shapes[1]).toBe(table);
```

- [ ] **Step 5: Run public/model suites and review lifecycle isolation**

Run:

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
git diff --check
```

Review field naming, strict type reuse, snapshot collapse for none, editor clear behavior, duplicate isolation, rollback, reopen, invalid zero mutation, and absence of `AddTableOptions.fit`.

- [ ] **Step 6: Commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/slide.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git commit -m 'feat: expose table cell text fit creation'
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 3: Lock the intentional PptxGenJS difference

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: Task 2 public native creation option and existing PptxGenJS ignored-runtime fixture.
- Produces: explicit evidence that PptxGenJS ignores table fit while native shrink/resize are strict creation extensions.

- [ ] **Step 1: Extend the existing ignored-runtime test with a native contrast**

After the PptxGenJS reopen assertions, add:

```ts
const native = PptxDocument.create({ slideSize: 'wide' });
const nativeSlide = native.addSlide();
const nativeTable = nativeSlide.addTable([[
  { text: 'Omitted' },
  { text: 'None', options: { fit: 'none' } },
  { text: 'Shrink', options: { fit: 'shrink' } },
  { text: 'Resize', options: { fit: 'resize', textDirection: 'vert' } },
]], {
  columnWidths: inches(2),
  rowHeights: inches(1),
});
expect(nativeTable.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
  undefined,
  undefined,
  'shrink',
  'resize',
]);
expect(nativeTable.rows[0]!.cells[3]!.textDirection).toBe('vert');

const nativeXml = new TextDecoder().decode(
  native.opcPackage.requirePart(nativeSlide.partUri).bytes,
);
expect(nativeXml.match(/<a:normAutofit\/>/g)).toHaveLength(1);
expect(nativeXml.match(/<a:spAutoFit\/>/g)).toHaveLength(1);
expect(nativeXml).not.toContain('<a:noAutofit/>');

const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes;
const invalidJournal = [...native.opcPackage.mutations];
expect(() => nativeSlide.addTable([[
  { text: 'Invalid', options: { fit: 'SHRINK' as never } },
]])).toThrow(TypeError);
expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
expect(native.opcPackage.mutations).toEqual(invalidJournal);
expect(nativeSlide.shapes[0]).toBe(nativeTable);
```

Use the already imported `inches`; do not access `_slides` or other private fields.

- [ ] **Step 2: Run adapter and dependency-boundary tests**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/pptxgenjs-adapter/src/dependency-boundary.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/pptxgenjs-adapter/tsconfig.json --pretty false
git diff --check
```

Review that PptxGenJS behavior is still described as ignored, native is not called parity, valid non-fit state remains unchanged, and invalid native input is atomic.

- [ ] **Step 3: Commit, push, fetch, and prove synchronization**

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git commit -m 'test: contrast table cell text fit creation with pptxgenjs'
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 4: Prove the packed Node/browser/type surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 2 public runtime and declarations.
- Produces: actual-tarball `tableCellTextFitCreation: true` for Node/browser plus compile-time `AddTableCellOptions.fit` coverage.

- [ ] **Step 1: Add Node creation values to the existing smoke table**

Add `fit` without creating another deck:

```js
{ text: 'Region', options: {
  align: 'left',
  border: creationBorder,
  fill: creationFill,
  fit: 'shrink',
  margin: creationMargin,
  textDirection: 'vert',
  valign: 'top',
} }
```

Set the remaining three cells to `resize`, `none`, and omitted. Capture:

```js
const initialCreatedFits = createdTable.rows.map(({ cells }) =>
  cells.map(({ textFit }) => textFit));
const createdTableFits = createdTableCells.map((cellXml) =>
  cellXml.match(/<a:(normAutofit|spAutoFit)\/>/)?.[1]);
const reopenedCreatedFits = reopenedCreatedTable instanceof TableModel
  ? reopenedCreatedTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))
  : undefined;
```

Define the flag after direction creation:

```js
const tableCellTextFitCreation = tableCellTextDirectionCreation &&
  JSON.stringify(initialCreatedFits) === JSON.stringify([
    ['shrink', 'resize'],
    [undefined, undefined],
  ]) &&
  JSON.stringify(createdTableFits) === JSON.stringify([
    'normAutofit',
    'spAutoFit',
    undefined,
    undefined,
  ]) &&
  JSON.stringify(reopenedCreatedFits) === JSON.stringify([
    ['shrink', 'resize'],
    [undefined, undefined],
  ]);
```

Emit `tableCellTextFitCreation` immediately before the existing `tableCellTextFit` flag.

- [ ] **Step 2: Mirror the runtime proof in the browser entry smoke**

Set the browser table's four cell inputs to shrink/resize/none/omitted and require:

```js
if (JSON.stringify(createdTable.rows.map(({ cells }) =>
  cells.map(({ textFit }) => textFit))) !== JSON.stringify([
  ['shrink', 'resize'],
  [undefined, undefined],
])) throw new Error('Browser table cell text fit creation failed');
```

After reading `browserCreatedTableCells`, require exact direct children:

```js
const browserCreatedTableFits = browserCreatedTableCells.map((cellXml) =>
  cellXml.match(/<a:(normAutofit|spAutoFit)\/>/)?.[1]);
if (JSON.stringify(browserCreatedTableFits) !== JSON.stringify([
  'normAutofit',
  'spAutoFit',
  undefined,
  undefined,
])) throw new Error('Browser table cell text fit XML creation failed');
```

- [ ] **Step 3: Compile the declaration surface**

In the generated declaration fixture, add the existing `cellFit` to creation options:

```ts
const creationOptions: AddTableCellOptions = {
  align: cellHorizontalAlignment,
  border: creationBorder,
  fill: cellFill,
  fit: cellFit,
  margin: creationMargin,
  textDirection: cellDirection,
  valign: cellAlignment,
};
```

- [ ] **Step 4: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_fit_creation_package_dir=$(mktemp -d /tmp/pptx-table-cell-fit-creation-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_fit_creation_package_dir"
table_cell_fit_creation_tarball=$(find "$table_cell_fit_creation_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$table_cell_fit_creation_tarball"
cd ../..
```

Require every smoke field true, including both `tableCellTextFitCreation` and `tableCellTextFit`.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review Node/browser/types coverage, exact child tokens, none collapse, existing editor flag retention, and no source-tree package dependency regression. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git commit -m 'test: smoke packed table cell text fit creation'
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 5: Document the creation contract and remaining boundary

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: completed runtime, adapter, declaration, and tarball evidence.
- Produces: accurate cell-level creation documentation without claiming PptxGenJS parity or table-level fit support.

- [ ] **Step 1: Update examples and public contract**

Add `fit: 'shrink'` to one `AddTableCellOptions` example. Change every cell option shape from:

```ts
{ align, border, fill, margin, textDirection, valign }
```

to:

```ts
{ align, border, fill, fit, margin, textDirection, valign }
```

Document all of these exact semantics:

- only cell-level `fit` is supported;
- strict values are `none`, `shrink`, and `resize`;
- omitted/undefined/none produce byte-identical self-closing `bodyPr` and direct snapshot `undefined`;
- shrink/resize write direct `normAutofit`/`spAutoFit` and snapshot immediately;
- no `noAutofit` creation, table metadata, font measurement, scale calculation, or geometry recomputation;
- direct editor none/undefined clear and do not restore creation input;
- fit is independent of cell direction in `tcPr`;
- PptxGenJS 4.0.1 ignores all table fit-like runtime inputs, so native shrink/resize are extensions.

- [ ] **Step 2: Update compatibility and unsupported lists**

Replace the matrix row with:

```md
| table-cell bodyPr autofit | `AddTableCellOptions.fit` / `TableCell.textFit` / `TableModel.setCellTextFit()` | 原生 cell 创建与 direct 编辑已支持；PptxGenJS 4.0.1 本身无 table fit API |
```

Update the `slide.addTable()` row to include `fit` in the cell-object shape. Remove only cell-level fit creation from unsupported lists; retain table-level fit creation, merge/span, hyperlinks, rich/multi-paragraph cells, advanced borders/fills, row insertion/deletion, styles, auto-page/repeated headers, content measurement, and layout recomputation.

Add one changelog bullet:

```md
- Added strict table-cell text-fit creation for none, shrink, and resize, reusing direct body-properties snapshots/editing while documenting PptxGenJS 4.0.1's lack of a table fit API.
```

- [ ] **Step 3: Scan contradictions and run full typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'cell(-level)? fit creation.*(unsupported|pending)|fit creation.*尚未支持|\{ align, border, fill, margin, textDirection, valign \}' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md || true
node node_modules/typescript/bin/tsc -b --pretty false
```

Review every hit in context: table-level fit and dynamic layout remain pending; cell-level creation does not.

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git commit -m 'docs: document table cell text fit creation'
```

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 6: Run full gates and real-deck QA

**Files:**
- Review every Task 1–5 path; never stage `.pnpm-store/`.

**Interfaces:**
- Consumes: implementation, tests, docs, actual tarball, repository CLI, LibreOffice, Poppler, and overflow checker.
- Produces: a fully verified pushed feature; any discovered defect gets a separate focused fix/review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-fit-creation-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-cell-fit-creation-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Require all functional suites and tests pass; the only default pending test is the separately executed performance test, which must stay below its 5-second budget.

- [ ] **Step 2: Rebuild and smoke the actual tarball plus offline doctor**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_fit_qa_package_dir=$(mktemp -d /tmp/pptx-table-cell-fit-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_fit_qa_package_dir"
table_cell_fit_qa_tarball=$(find "$table_cell_fit_qa_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$table_cell_fit_qa_tarball"
cd ../..
command -v pptx-inspect
pptx-inspect --json doctor
```

- [ ] **Step 3: Generate native, editor-baseline, edited, and reopened decks**

Create `/tmp/pptx-table-cell-fit-creation-qa.mjs` with `apply_patch`:

```js
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import {
  PptxDocument,
  TableModel,
  inches,
} from 'file:///Users/jeremy/workspace/pptx/packages/pptx/dist/index.js';

const output = '/tmp/pptx-table-cell-fit-creation-qa';
await mkdir(output, { recursive: true });
const geometry = {
  name: 'Table cell fit creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(2), inches(2), inches(2), inches(2)],
  rowHeights: [inches(1.5), inches(1.5)],
  margin: 7.2,
  valign: 'middle',
};

function build(rows) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const table = document.addSlide().addTable(rows, geometry);
  return { document, table };
}

for (const [name, fit] of [
  ['omitted', Symbol.for('omitted')],
  ['undefined', undefined],
  ['none', 'none'],
]) {
  const options = fit === Symbol.for('omitted') ? {} : { fit };
  const built = build([[{ text: 'Equivalent', options }]]);
  await writeFile(`${output}/${name}.pptx`, await built.document.write());
}

const rows = [[
  { text: 'Shrink a deliberately long line of text', options: { fit: 'shrink' } },
  { text: 'Resize', options: { fit: 'resize' } },
  { text: 'None', options: { fit: 'none' } },
  { text: 'Combined vertical', options: {
    fit: 'shrink',
    textDirection: 'vert270',
    fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    border: {
      kind: 'line',
      color: { kind: 'srgb', value: 'C00000' },
      width: 2,
    },
  } },
], [
  'String omitted',
  { text: 'Undefined', options: { fit: undefined } },
  { text: 'Shrink second row', options: { fit: 'shrink' } },
  { text: 'Resize second row', options: { fit: 'resize' } },
]];

const native = build(rows);
assert.deepEqual(native.table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit)), [
  ['shrink', 'resize', undefined, 'shrink'],
  [undefined, undefined, 'shrink', 'resize'],
]);
await writeFile(`${output}/native-source.pptx`, await native.document.write());

const baselineRows = rows.map((row) => row.map((cell) =>
  typeof cell === 'string'
    ? cell
    : { text: cell.text, options: Object.fromEntries(
      Object.entries(cell.options ?? {}).filter(([key]) => key !== 'fit'),
    ) }));
const baseline = build(baselineRows);
for (const [row, column, fit] of [
  [0, 0, 'shrink'],
  [0, 1, 'resize'],
  [0, 3, 'shrink'],
  [1, 2, 'shrink'],
  [1, 3, 'resize'],
]) baseline.table.setCellTextFit(row, column, fit);
await writeFile(`${output}/editor-baseline.pptx`, await baseline.document.write());

native.table.setCellTextFit(0, 0, 'none');
native.table.setCellTextFit(0, 2, 'resize');
native.table.setCellText(0, 2, 'Edited to resize');
const editedBytes = await native.document.write();
await writeFile(`${output}/native-edited.pptx`, editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.deepEqual(reopenedTable.rows[0].cells.map(({ textFit }) => textFit), [
  undefined,
  'resize',
  'resize',
  'shrink',
]);
assert.equal(reopenedTable.rows[0].cells[2].text, 'Edited to resize');
assert.equal(reopenedTable.rows[0].cells[3].textDirection, 'vert270');
await writeFile(`${output}/native-reopened.pptx`, await reopened.write());

process.stdout.write(JSON.stringify({
  source: native.table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit)),
  reopened: reopenedTable.rows.map(({ cells }) => cells.map(({ textFit }) => textFit)),
  widths: reopenedTable.columnWidths,
  heights: reopenedTable.rowHeights,
}));
```

Run `node /tmp/pptx-table-cell-fit-creation-qa.mjs` and require the assertions plus widths `[1828800,1828800,1828800,1828800]` and heights `[1371600,1371600]`.

- [ ] **Step 4: Validate packages, exact tokens, and mutation isolation**

```sh
for deck in /tmp/pptx-table-cell-fit-creation-qa/*.pptx; do
  pptx-inspect --json package validate "$deck" --profile powerpoint-2010
done
pptx-inspect --json package diff \
  /tmp/pptx-table-cell-fit-creation-qa/omitted.pptx \
  /tmp/pptx-table-cell-fit-creation-qa/undefined.pptx
pptx-inspect --json package diff \
  /tmp/pptx-table-cell-fit-creation-qa/omitted.pptx \
  /tmp/pptx-table-cell-fit-creation-qa/none.pptx
pptx-inspect --json package diff \
  /tmp/pptx-table-cell-fit-creation-qa/native-source.pptx \
  /tmp/pptx-table-cell-fit-creation-qa/editor-baseline.pptx
pptx-inspect --json package diff \
  /tmp/pptx-table-cell-fit-creation-qa/native-source.pptx \
  /tmp/pptx-table-cell-fit-creation-qa/native-edited.pptx
pptx-inspect --json package diff \
  /tmp/pptx-table-cell-fit-creation-qa/native-edited.pptx \
  /tmp/pptx-table-cell-fit-creation-qa/native-reopened.pptx
```

Require every deck to report zero errors/warnings; omitted→undefined, omitted→none, source→editor-baseline, and edited→reopened must have zero changed parts. Source→edited must change only `/ppt/slides/slide1.xml`.

Use the skill's exact-part read path:

```sh
pptx-inspect --json part read \
  /tmp/pptx-table-cell-fit-creation-qa/native-source.pptx \
  /ppt/slides/slide1.xml | jq -r '.data.content' | \
  rg -o '<a:(noAutofit|normAutofit|spAutoFit)\b' | sort | uniq -c
```

Require 3 `normAutofit`, 2 `spAutoFit`, and zero `noAutofit`; no fit token may occur under `tcPr`.

- [ ] **Step 5: Render every QA state and check overflow**

Use isolated LibreOffice profiles to export native source, native edited, and editor baseline to PDF; rasterize each PDF at 180 DPI with `pdftoppm`. Run:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-fit-creation-qa/native-source.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-fit-creation-qa/native-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-fit-creation-qa/editor-baseline.pptx
```

Inspect every PNG at original detail. Require all text, blue fill, red border, vertical direction, and cell geometry remain visible with no repair, clipping, off-slide content, overlap, blur, missing cell, or unexpected wrapping. Do not claim that LibreOffice computes the same final shrink factor as PowerPoint; only require native and editor-baseline rendering to match.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict normalization, descriptor safety, none collapse, direct bodyPr ownership, bodyPr/lstStyle order, declarations, lifecycle isolation, PptxGenJS difference wording, smoke flags, docs boundary, and invalid zero-mutation assertions.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `?? .pnpm-store/`. Successful QA creates no empty commit; any defect is fixed, reviewed, committed, pushed, fetched, and re-verified before the feature is complete.
