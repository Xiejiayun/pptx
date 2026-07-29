# Table Cell Border Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user explicitly requires inline execution without subagents.

**Goal:** Let native `slide.addTable()` create cells with strict direct four-side borders through `{ text, options: { border } }`, while preserving existing default bytes and PptxGenJS-compatible materialized sides.

**Architecture:** Upgrade the existing border normalizer to descriptor-safe deep reads, export one canonical side renderer, and let table creation overlay normalized scalar/TRBL/named border input on the four existing direct noFill sides. Keep public read/edit types, table transaction flow, and existing-deck mutation semantics unchanged.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler.

## Global Constraints

- Public creation property is singular `AddTableCellOptions.border?: TableCellBorderInput`.
- Creation border uses canonical four-side noFill as its base; undefined tuple/named sides preserve noFill.
- `setCellBorders()` remains a whole-replacement editor; this plan does not change its output semantics.
- Border sides support only none or strict solid sRGB/theme line, point width `0..1584`, and omitted/solid/dash style.
- Public tuple order is T/R/B/L; OOXML order is L/R/T/B; cell fill follows all four borders.
- All nested border input is descriptor-safe, getter-free, null-prototype-compatible, and detached.
- Omitted/undefined/empty border, all-undefined tuple, and scalar four-side none keep current native table bytes.
- No table-level border, diagonal, transparency, advanced line, merge, layout, or unrelated option work enters this slice.
- Every successful small item is reviewed, committed, pushed, and verified `origin/main...HEAD = 0 0`.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Make shared border normalization descriptor-safe and expose canonical creation rendering

**Files:**
- Modify: `packages/model/src/table-cell-borders.internal.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: existing `TableCellBorder`, `TableCellBorders`, `normalizeTableCellBorders()`, and editor XML.
- Produces:

```ts
export function renderTableCellBorders(
  borders: TableCellBorders | undefined,
  prefix: string,
): string;
```

The renderer overlays `borders` on four `{ kind: 'none' }` defaults and returns canonical L/R/T/B XML.

- [ ] **Step 1: Extend existing editor invalid-input coverage first**

In the existing `setCellBorders()` model test, add accessor/class/inherited/symbol cases at every layer:

```ts
const accessorBorder = {};
const accessorColor = { kind: 'srgb' };
let borderAccessorCalls = 0;
Object.defineProperty(accessorBorder, 'kind', {
  get() {
    borderAccessorCalls += 1;
    return 'none';
  },
  enumerable: true,
});
Object.defineProperty(accessorColor, 'value', {
  get() {
    borderAccessorCalls += 1;
    return 'FF0000';
  },
  enumerable: true,
});

const accessorTuple = [undefined, undefined, undefined, undefined];
Object.defineProperty(accessorTuple, '0', {
  get() {
    borderAccessorCalls += 1;
    return { kind: 'none' };
  },
  enumerable: true,
});

const accessorNamed = {};
Object.defineProperty(accessorNamed, 'top', {
  get() {
    borderAccessorCalls += 1;
    return { kind: 'none' };
  },
  enumerable: true,
});
```

Reject accessor scalar/named/tuple/color, inherited objects, class instances, arrays used as sides, and symbol extras. Add valid null-prototype scalar/named/color input and confirm normalized snapshots. Require `borderAccessorCalls === 0`, exact slide bytes, and unchanged mutation journal after invalid input.

- [ ] **Step 2: Run the model test and confirm red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t "table-cell borders"
```

Expected: at least one accessor getter is invoked or an exotic object is accepted.

- [ ] **Step 3: Replace direct property reads with descriptor-safe copies**

Add a strict helper inside `table-cell-borders.internal.ts`:

```ts
function readDataObject(
  value: unknown,
  context: string,
  supported: readonly string[],
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const allowed = new Set(supported);
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}
```

Use it for scalar/named detection, each side, and nested color. In `normalizeTuple()`, require `Object.getPrototypeOf(value) === Array.prototype` so array subclasses/exotic tuples are rejected. Rewrite tuple reads to inspect every own index descriptor before accessing a value:

```ts
const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
  throw new TypeError(`${context} TRBL tuple must contain only data items`);
}
const item = descriptor.value;
```

Do not use `candidate.kind`, `candidate[side]`, `value[index]`, or nested color fields until they have been copied by the helper.

- [ ] **Step 4: Export canonical border rendering**

Keep the existing single-side `renderBorder()` internal, then add:

```ts
const NONE_BORDER: TableCellBorder = { kind: 'none' };

export function renderTableCellBorders(
  borders: TableCellBorders | undefined,
  prefix: string,
): string {
  return XML_SIDES.map(([side, tag]) =>
    renderBorder(tag, borders?.[side] ?? NONE_BORDER, prefix)).join('');
}
```

This must reproduce the existing `NO_BORDERS` string exactly when `borders` is undefined or all values are none. Do not change editor insertion, equality, or whole-replacement behavior.

- [ ] **Step 5: Run typecheck and focused editor tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t "table-cell borders"
```

Expected: descriptor-safety and existing editor lifecycle pass.

---

### Task 2: Normalize and render cell border during native table creation

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 `normalizeTableCellBorders()` and `renderTableCellBorders()`.
- Produces: internal `NormalizedTableCell.borders?: TableCellBorders` and canonical border-before-fill XML.

- [ ] **Step 1: Add internal normalization and exact-XML tests first**

Extend the existing cell-options test with rows containing:

```ts
const sourceColor = { kind: 'srgb' as const, value: '#ff0000' };
const sourceLine = {
  kind: 'line' as const,
  color: sourceColor,
  width: 1.50004,
  style: 'solid' as const,
};

const rows = [[
  'String',
  { text: 'Empty options', options: {} },
  { text: 'Undefined border', options: { border: undefined } },
  { text: 'Empty border', options: { border: {} } },
  { text: 'None', options: { border: { kind: 'none' as const } } },
  { text: 'Scalar', options: { border: sourceLine } },
  { text: 'Tuple', options: { border: [
    { kind: 'line' as const, color: { kind: 'scheme' as const, value: 'accent1' }, width: 2, style: 'dash' as const },
    undefined,
    { kind: 'line' as const, color: { kind: 'srgb' as const, value: '00FF00' }, width: 0 },
    { kind: 'none' as const },
  ] } },
  { text: 'Named + fill', options: {
    border: { top: sourceLine, left: { kind: 'none' as const } },
    fill: { kind: 'solid' as const, color: { kind: 'scheme' as const, value: 'accent2' }, transparency: 25 },
  } },
]];
```

Require normalized point width `1.5`, uppercased color, detached nested data, and explicit overrides only in `NormalizedTableCell.borders`.

Build equivalent one-cell definitions for string, `{ text }`, empty options, undefined border, `{}`, all-undefined tuple, and scalar none. Require identical `renderTableGraphicFrame()` strings.

For styled cells require:

- L/R/T/B order even though tuple input is T/R/B/L.
- tuple undefined right becomes canonical right noFill.
- width 1.5pt writes `w="19050"`; width zero remains solid line `w="0"`.
- dash writes `sysDash`; omitted style writes no `prstDash`; explicit solid writes `solid`.
- theme and sRGB color tags are exact.
- `lnB` closes before optional cell fill begins.

- [ ] **Step 2: Add invalid border creation matrix**

In the getter-count matrix add:

- cell options `border` accessor;
- scalar border `kind` accessor;
- named `top` accessor;
- tuple index accessor, sparse/wrong-length/extra-key tuple;
- side class/inherited/array/symbol-extra objects;
- nested color accessor/class/inherited/symbol-extra;
- missing kind/color/width, unknown kind/style, invalid hex/theme;
- width `NaN`, infinities, negative, over 1584.

Require zero accessor calls and reject before table geometry/rendering.

- [ ] **Step 3: Run the internal suite and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts
```

Expected: border is rejected as an unsupported options property.

- [ ] **Step 4: Extend normalized cells and option keys**

In `table-create.internal.ts` import:

```ts
import {
  normalizeTableCellBorders,
  renderTableCellBorders,
} from './table-cell-borders.internal.js';
import type { TableCellBorders, TableCellFill } from './shapes.js';
```

Extend:

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
}
```

Allow options keys `['border', 'fill']`, normalize border first, and return only defined fields:

```ts
const borders = normalizeTableCellBorders(options.border, `${context} border`);
const fill = normalizeTableCellFill(options.fill, `${context} fill`);
return {
  ...(borders === undefined ? {} : { borders }),
  ...(fill === undefined ? {} : { fill }),
};
```

- [ ] **Step 5: Replace the local no-border constant with shared rendering**

Remove `NO_BORDERS` from `table-create.internal.ts`. Render:

```ts
const borders = renderTableCellBorders(cell.borders, 'a:');
const fill = cell.fill === undefined ? '' : renderTableCellFill(cell.fill, 'a:');
return `<a:tc>...<a:tcPr ...>${borders}${fill}</a:tcPr></a:tc>`;
```

No-border output must remain byte-identical.

- [ ] **Step 6: Run focused internal and editor suites**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: exact XML, invalid getter matrix, and existing border/fill editor tests pass.

---

### Task 3: Expose the public option and prove model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 2 internal border creation.
- Produces public `AddTableCellOptions.border?: TableCellBorderInput`.

- [ ] **Step 1: Add public typing and lifecycle assertions first**

Import `type TableCellBorderInput` through the public model surface. Type a mutable source border and create a matrix with:

- empty options;
- scalar sRGB solid;
- TRBL theme dash / undefined / zero-width / none;
- named partial border plus existing fill;
- scalar none.

Require immediate normalized `TableCell.borders` snapshots, caller deep mutation isolation, canonical noFill on unspecified sides, unchanged fill, transform, column widths, row heights, shape identity, and non-target package parts.

Immediately call `setCellBorders()` to replace and clear a created border, then require the existing whole-replacement semantics. Add invalid public border values and prove exact slide bytes/mutation journal isolation. Create a filled+bordered table inside an outer transaction and prove rollback removes it. Write/reopen and require remaining borders, fill, geometry, and identity-by-id.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: `AddTableCellOptions` rejects `border`.

- [ ] **Step 3: Widen only the existing public creation option**

In `slide.ts`, import `type TableCellBorderInput` from `shapes.ts` and update:

```ts
export interface AddTableCellOptions {
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
}
```

Do not change `AddTableCell`, `AddTableCellInput`, `addTable()`, `TableCell`, or `TableModel.setCellBorders()` signatures.

- [ ] **Step 4: Run model typecheck/tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: declarations and the public creation/edit/rollback/reopen lifecycle pass.

---

### Task 4: Prove SDK, PptxGenJS, packed, and documented behavior

**Files:**
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: Task 3 aggregate exports and border-capable cell objects.
- Produces SDK/adapter/browser/declaration/package coverage and accurate support boundaries.

- [ ] **Step 1: Extend the SDK table lifecycle**

Type one source option as `AddTableCellOptions`; create sRGB solid, theme dash, none, zero-width, TRBL, named, and border+fill cells in the existing mixed matrix. Mutate source side/color after creation and require detached snapshots. Keep duplicate isolation, editor mutation, rollback, geometry, non-target part, and reopen assertions; the duplicate must retain original creation borders after source table edits.

- [ ] **Step 2: Add supported PptxGenJS border creation conformance**

In the existing adapter border test, create a native table equivalent to public PptxGenJS cell border output:

```ts
const nativeTable = native.addSlide().addTable([[
  { text: 'Scalar', options: { border: {
    kind: 'line',
    color: { kind: 'srgb', value: 'FF0000' },
    width: 2,
    style: 'solid',
  } } },
  { text: 'Tuple', options: { border: [
    { kind: 'none' },
    { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 1.5, style: 'dash' },
    { kind: 'line', color: { kind: 'srgb', value: '0000FF' }, width: 0, style: 'solid' },
    { kind: 'line', color: { kind: 'srgb', value: '666666' }, width: 1, style: 'solid' },
  ] } },
]], geometry);
```

Generate the corresponding PptxGenJS table with only public APIs. Require equal final border snapshots, margins, fill, geometry, L/R/T/B XML ordering, colors, widths, and dash tokens after write/reopen.

Separately assert intentional direct-state differences:

- native empty border stays four noFill; PptxGenJS `{}` becomes 1pt gray solid;
- native omitted style has no direct dash; PptxGenJS default writes direct solid;
- native named/theme inputs are supported extensions;
- native width zero remains zero.

- [ ] **Step 3: Extend Node/browser/declaration tarball smoke**

Use created border+fill object cells in both Node and browser tables. Add `tableCellBorderCreation: true` only when creation snapshot, source detachment, scalar/TRBL/named, sRGB/theme, solid/dash/none/zero, edit, write/reopen, and existing fill creation all pass.

In declaration smoke, use:

```ts
const creationBorder: TableCellBorderInput = [
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1, style: 'dash' },
  { kind: 'none' },
  undefined,
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 2 },
];
const creationOptions: AddTableCellOptions = {
  border: creationBorder,
  fill: cellFill,
};
```

Keep both values in the final `void` expression.

- [ ] **Step 4: Update docs and compatibility boundaries**

Add one Unreleased bullet. Update API/package examples to include `options.border` and fill together. Document descriptor safety, detached normalization, scalar/TRBL/named inputs, canonical noFill overlay, point quantization/range, explicit zero line, dash mapping, XML order, PptxGenJS defaults/style differences, and remaining unsupported options.

- [ ] **Step 5: Run focused conformance and docs checks**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --check
rg -n "border.*AddTableCellOptions|options.*border|tableCellBorderCreation" \
  packages scripts CHANGELOG.md docs
```

Expected: focused suites/typecheck/docs scan pass; unsupported table-level/diagonal/advanced borders remain visible.

---

### Task 5: Full gates, real-deck QA, review, commit, and push

**Files:**
- Review every Task 1–4 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: implementation/tests/docs, actual packed package, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces reviewed `feat: create tables with cell borders` synchronized to `origin/main`.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-border-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Parse suite/test passed, failed, skipped, and todo totals. Retry unchanged default gates only for proven transient host load; never alter repository timeouts to hide a failure.

- [ ] **Step 2: Build and smoke the actual npm tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js \
  --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_border_package_dir=$(mktemp -d \
  /tmp/pptx-table-cell-border-package.XXXXXX)
npm pack --ignore-scripts \
  --pack-destination "$table_cell_border_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_cell_border_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require `tableCellBorderCreation: true`, `tableCellFillCreation: true`, all existing table flags true, declaration success, browser success, and CLI `0.1.0`.

- [ ] **Step 3: Verify repository CLI health**

```sh
node packages/pptx/dist/cli.js --json doctor
```

Require successful offline JSON.

- [ ] **Step 4: Generate native equivalent/bordered/edited/reopened and PptxGenJS decks**

Create `/tmp/pptx-table-cell-border-qa.mjs` with `apply_patch` using this complete program:

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
const output = '/tmp/pptx-table-cell-border';
await mkdir(output, { recursive: true });

const textRows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];
const tableOptions = {
  name: 'Cell border creation QA',
  x: inches(1),
  y: inches(1),
  columnWidths: [inches(2), inches(4), inches(2)],
  rowHeights: [inches(0.5), inches(1.25), inches(0.65)],
};

function build(rows) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const slide = document.addSlide();
  return { document, table: slide.addTable(rows, tableOptions) };
}

const strings = build(textRows);
const emptyBorder = build(textRows.map((row) => row.map((text) => ({
  text,
  options: { border: {} },
}))));
await writeFile(output + '/string-source.pptx', await strings.document.write());
await writeFile(
  output + '/empty-border-source.pptx',
  await emptyBorder.document.write(),
);

const solid = (color, width, style = 'solid') => ({
  kind: 'line',
  color,
  width,
  style,
});
const srgb = (value) => ({ kind: 'srgb', value });
const scheme = (value) => ({ kind: 'scheme', value });
const none = { kind: 'none' };
const borderRows = [
  [
    { text: 'Region', options: {
      border: solid(srgb('C00000'), 2),
      fill: { kind: 'solid', color: scheme('accent1') },
    } },
    { text: 'Revenue', options: {
      border: none,
      fill: { kind: 'solid', color: srgb('D9EAF7') },
    } },
    { text: 'Growth', options: {
      border: solid(scheme('accent2'), 2, 'dash'),
      fill: { kind: 'solid', color: scheme('accent2'), transparency: 25 },
    } },
  ],
  [
    'East',
    { text: '$1.2M', options: {
      border: [
        solid(srgb('C00000'), 2),
        solid(scheme('accent3'), 1.5, 'dash'),
        none,
        solid(srgb('70AD47'), 0),
      ],
      fill: { kind: 'solid', color: srgb('FFF2CC'), transparency: 33.333 },
    } },
    { text: 'Vertical', options: {
      border: {
        top: solid(scheme('accent1'), 1, 'dash'),
        left: solid(srgb('4472C4'), 3),
      },
    } },
  ],
  [
    { text: 'West', options: { border: none } },
    { text: '$980K', options: { border: solid(srgb('0000FF'), 0) } },
    { text: '', options: { border: {
      bottom: solid(srgb('70AD47'), 2, 'dash'),
      right: none,
    } } },
  ],
];
const bordered = build(borderRows);
bordered.table.setCellTextDirection(1, 2, 'vert270');
bordered.table.setCellVerticalAlignment(2, 1, 'bottom');
const sourceBorders = bordered.table.rows.map(({ cells }) =>
  cells.map(({ borders }) => borders));
await writeFile(output + '/border-source.pptx', await bordered.document.write());

bordered.table.setCellBorders(2, 2, solid(srgb('70AD47'), 3, 'dash'));
bordered.table.setCellText(2, 2, 'Edited empty');
const editedBytes = await bordered.document.write();
await writeFile(output + '/border-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[2].cells[2].text, 'Edited empty');
assert.deepEqual(reopenedTable.rows[2].cells[2].borders, {
  top: solid(srgb('70AD47'), 3, 'dash'),
  right: solid(srgb('70AD47'), 3, 'dash'),
  bottom: solid(srgb('70AD47'), 3, 'dash'),
  left: solid(srgb('70AD47'), 3, 'dash'),
});
assert.equal(reopenedTable.rows[1].cells[2].textDirection, 'vert270');
assert.equal(reopenedTable.rows[2].cells[1].verticalAlignment, 'bottom');
assert.deepEqual(reopenedTable.rows[0].cells[0].fill, {
  kind: 'solid',
  color: scheme('accent1'),
});
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/border-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      { text: 'Region', options: {
        border: { type: 'solid', color: 'C00000', pt: 2 },
        fill: { color: baseline.SchemeColor.accent1 },
      } },
      { text: 'Revenue', options: {
        border: { type: 'none' },
        fill: { color: 'D9EAF7' },
      } },
      { text: 'Growth', options: {
        border: { type: 'dash', color: 'ED7D31', pt: 2 },
        fill: { color: baseline.SchemeColor.accent2, transparency: 25 },
      } },
    ],
    [
      { text: 'East', options: {} },
      { text: '$1.2M', options: {
        border: [
          { type: 'solid', color: 'C00000', pt: 2 },
          { type: 'dash', color: 'A5A5A5', pt: 1.5 },
          { type: 'none' },
          { type: 'solid', color: '70AD47', pt: 0 },
        ],
        fill: { color: 'FFF2CC', transparency: 33.333 },
      } },
      { text: 'Vertical', options: { border: [
        { type: 'dash', color: '4472C4', pt: 1 },
        { type: 'none' },
        { type: 'none' },
        { type: 'solid', color: '4472C4', pt: 3 },
      ] } },
    ],
    [
      { text: 'West', options: { border: { type: 'none' } } },
      { text: '$980K', options: {
        border: { type: 'solid', color: '0000FF', pt: 0 },
      } },
      { text: '', options: { border: [
        { type: 'none' },
        { type: 'none' },
        { type: 'dash', color: '70AD47', pt: 2 },
        { type: 'none' },
      ] } },
    ],
  ],
  {
    x: 1,
    y: 1,
    w: 8,
    h: 2.4,
    colW: [2, 4, 2],
    rowH: [0.5, 1.25, 0.65],
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceBorders,
  reopenedBorders: reopenedTable.rows.map(({ cells }) =>
    cells.map(({ borders }) => borders)),
  editedText: reopenedTable.rows[2].cells[2].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run:

```sh
node /tmp/pptx-table-cell-border-qa.mjs
```

Require stable normalized border matrices, `editedText: "Edited empty"`, column widths `[1828800,3657600,1828800]`, and row heights `[457200,1143000,594360]`.

- [ ] **Step 5: Validate packages and exact diff isolation**

For all six decks run:

```sh
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-cell-border/<deck>.pptx \
  --profile powerpoint-2010
```

Require zero errors/warnings. Diff:

```sh
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-border/string-source.pptx \
  /tmp/pptx-table-cell-border/empty-border-source.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-border/border-source.pptx \
  /tmp/pptx-table-cell-border/border-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-border/border-edited.pptx \
  /tmp/pptx-table-cell-border/border-reopened.pptx
```

Require zero changed parts, only `/ppt/slides/slide1.xml`, and zero changed parts respectively.

- [ ] **Step 6: Render and inspect native/baseline decks**

Use bundled workspace dependencies. Convert `border-edited.pptx` and `pptxgenjs-baseline.pptx` to PDF with isolated LibreOffice user profiles, rasterize at 180 DPI with Poppler, and run:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-border/border-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-border/pptxgenjs-baseline.pptx
```

Inspect every PNG at original detail. Require visible solid/dash/none/per-side borders, distinguishable line widths/colors, preserved cell fill, all text, unequal rows/columns, no repair, clipping, unexpected wrap, overlap, blur, missing side/cell, or off-slide content.

- [ ] **Step 7: Final static/staged review**

```sh
git diff --check
git status --short
git diff --stat
```

Review descriptor reads, exact default bytes, L/R/T/B ordering, border-before-fill, editor behavior, public declarations, smoke output, docs boundaries, and every invalid-input no-mutation assertion.

- [ ] **Step 8: Stage only targets, commit, push, and prove synchronization**

```sh
git add \
  packages/model/src/table-cell-borders.internal.ts \
  packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "feat: create tables with cell borders"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
git status --short --branch
```

Expected: `0 0` and only `?? .pnpm-store/`.

---

After synchronization, design the next independently reviewable cell creation option without asking for routine decisions.
