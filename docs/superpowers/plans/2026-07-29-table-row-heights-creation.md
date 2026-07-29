# Table Row Heights Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend native `SlideModel.addTable()` with exact scalar or per-row EMU heights while keeping explicit table xfrm height equal to the OOXML row-height sum.

**Architecture:** Extend the focused table creation normalizer so it produces one detached exact `rowHeights` vector for auto, existing equal-distribution, and new explicit-height paths. Reuse a small positive-dimension vector normalizer for columns and rows, then make the renderer consume both normalized vectors directly; `SlideModel.addTable()` keeps its current transaction, shape-tree insertion, identity, and rollback behavior unchanged.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless source-span OOXML editing, PptxGenJS 4.0.1 public-output conformance, tsup, npm tarball smoke, repository CLI, LibreOffice headless, Poppler rendering.

## Global Constraints

- Public input is `AddTableOptions.rowHeights?: number | readonly number[]`, always in EMU.
- A scalar repeats exactly for every physical table row; an array must be dense, descriptor-safe, detached, and exactly match the rectangular matrix row count.
- Every normalized item must be a positive safe-integer EMU after `Math.round()`; the vector sum must also be a positive safe integer.
- When `rowHeights` is present and `height` is omitted, overall table height is the exact vector sum and auto-row-height is disabled.
- When both are present, rounded `height` must equal the vector sum or the call throws before package mutation.
- When `rowHeights` is omitted, preserve current auto-row-height and explicit equal-distribution output byte-for-byte.
- Preserve the auto case `xfrm height = 1 inch` with every `a:tr@h = 0`; for explicit rows always preserve `xfrm height === sum(row heights)`.
- Do not reproduce PptxGenJS single-array coercion, short/falsy fallback, long-array truncation, string coercion, invalid negative values, or xfrm/rows inconsistency.
- Do not add existing-table row snapshots/editing, transform/row resize, cell objects, rich cell text, merge, styles, auto-page, repeated headers, hyperlinks, or content measurement in this slice.
- Preserve strict matrix/options validation, only-target-slide mutation, stable identity, duplicate isolation, nested transaction rollback, write/reopen, existing column-width creation, and all cell editors.
- Implement inline without subagent delegation for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Commit and push implementation only after source review, typecheck, focused/full tests, performance, actual tarball smoke, CLI validation, PowerPoint 2010 validation, LibreOffice rendering, overflow checks, and full-size visual inspection pass.

---

### Task 1: Normalize exact row-height vectors and render them directly

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Test: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing descriptor-safe `readDenseArray()`, `normalizeCoordinate()`, `distributeTableDimension()`, column-width vector normalization, and `NormalizedTableDefinition`.
- Produces: `NormalizedTableDefinition.rowHeights`, exact scalar/array normalization, unchanged auto rows, and renderer output that uses the normalized row vector.

- [ ] **Step 1: Extend pure tests with red row-vector cases**

Update the two existing full-definition expectations:

```ts
expect(definition).toEqual({
  rows: [
    ['A & <1>', '', 'C1'],
    ['A2', 'B2', 'C2'],
  ],
  name: 'Table "A"',
  x: 457_200,
  y: 685_800,
  width: 2_743_201,
  height: 1_371_601,
  autoRowHeight: false,
  columnWidths: [914_401, 914_400, 914_400],
  rowHeights: [685_801, 685_800],
});

expect(defaults).toEqual({
  rows: [['A', 'B']],
  x: 457_200,
  y: 457_200,
  width: 1_828_800,
  height: 914_400,
  autoRowHeight: true,
  columnWidths: [914_400, 914_400],
  rowHeights: [0],
});
```

Add a focused test:

```ts
const unequalSource = [457_200.2, 914_399.7, 1_371_600];
const unequal = normalizeTableDefinition(
  [['A'], ['B'], ['C']],
  { rowHeights: unequalSource },
);
expect(unequal).toMatchObject({
  height: 2_743_200,
  autoRowHeight: false,
  rowHeights: [457_200, 914_400, 1_371_600],
});
expect(renderTableGraphicFrame(5, unequal)).toMatch(
  /<a:tr h="457200">[\s\S]*<a:tr h="914400">[\s\S]*<a:tr h="1371600">/,
);
unequalSource[0] = 1;
expect(unequal.rowHeights[0]).toBe(457_200);

const scalar = normalizeTableDefinition(
  [['A'], ['B'], ['C']],
  { rowHeights: 685_800 },
);
expect(scalar.height).toBe(2_057_400);
expect(scalar.rowHeights).toEqual([685_800, 685_800, 685_800]);

const matching = normalizeTableDefinition(
  [['A'], ['B'], ['C']],
  { height: 2_743_200, rowHeights: [457_200, 914_400, 1_371_600] },
);
expect(matching.height).toBe(2_743_200);
expect(matching.rowHeights).toEqual([457_200, 914_400, 1_371_600]);

expect(normalizeTableDefinition(
  [['A'], ['B']],
  { rowHeights: undefined },
).rowHeights).toEqual([0, 0]);
```

- [ ] **Step 2: Run the pure test and verify red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts
```

Expected: assertions fail because `rowHeights` is not accepted or returned yet.

- [ ] **Step 3: Extend option keys and normalized definition**

```ts
const OPTION_KEYS = [
  'name', 'x', 'y', 'width', 'height', 'columnWidths', 'rowHeights',
] as const;

export interface NormalizedTableDefinition {
  readonly rows: readonly (readonly string[])[];
  readonly name?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly autoRowHeight: boolean;
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
}
```

- [ ] **Step 4: Generalize the existing positive vector helpers**

Replace the column-only helpers with:

```ts
function normalizeDimensionVector(
  value: unknown,
  count: number,
  context: string,
  countContext: string,
): readonly number[] {
  const values = Array.isArray(value)
    ? readDenseArray(value, context)
    : Array.from({ length: count }, () => value);
  if (values.length !== count) {
    throw new TypeError(`${context} must match the table ${countContext}`);
  }
  return values.map((item, index) => {
    const dimension = normalizeCoordinate(item, `${context} ${index}`);
    if (dimension <= 0) {
      throw new RangeError(`${context} ${index} must be greater than zero`);
    }
    return dimension;
  });
}

function sumDimensions(dimensions: readonly number[], context: string): number {
  return dimensions.reduce((sum, dimension) => {
    if (dimension > Number.MAX_SAFE_INTEGER - sum) {
      throw new RangeError(`${context} sum must fit a safe integer EMU value`);
    }
    return sum + dimension;
  }, 0);
}
```

Update the column path without changing behavior:

```ts
columnWidths = normalizeDimensionVector(
  requestedColumnWidths,
  columnCount,
  'Table columnWidths',
  'column count',
);
const columnWidthSum = sumDimensions(columnWidths, 'Table columnWidths');
```

- [ ] **Step 5: Implement the strict row-height state machine**

Replace the current height-only path with:

```ts
const requestedRowHeights = normalizedOptions.rowHeights;
let height: number;
let autoRowHeight: boolean;
let rowHeights: readonly number[];
if (requestedRowHeights !== undefined) {
  rowHeights = normalizeDimensionVector(
    requestedRowHeights,
    normalizedRows.length,
    'Table rowHeights',
    'row count',
  );
  const rowHeightSum = sumDimensions(rowHeights, 'Table rowHeights');
  autoRowHeight = false;
  if (normalizedOptions.height === undefined) {
    height = rowHeightSum;
  } else {
    height = normalizeCoordinate(normalizedOptions.height, 'Table height');
    if (height !== rowHeightSum) {
      throw new RangeError('Table height must equal the sum of rowHeights');
    }
  }
} else {
  autoRowHeight = normalizedOptions.height === undefined;
  height = autoRowHeight
    ? DEFAULT_HEIGHT
    : normalizeCoordinate(normalizedOptions.height, 'Table height');
  if (height <= 0) throw new RangeError('Table height must be greater than zero');
  if (!autoRowHeight && height < normalizedRows.length) {
    throw new RangeError('Table height must provide at least one EMU per row');
  }
  rowHeights = autoRowHeight
    ? normalizedRows.map(() => 0)
    : distributeTableDimension(height, normalizedRows.length);
}
```

Return the complete normalized definition:

```ts
return {
  rows: normalizedRows,
  ...(name !== undefined ? { name } : {}),
  x,
  y,
  width,
  height,
  autoRowHeight,
  columnWidths,
  rowHeights,
};
```

The positive explicit vector already guarantees positive height, so do not add a second redundant validation branch.

- [ ] **Step 6: Render the exact normalized row vector**

Remove renderer-side row distribution and use:

```ts
const rows = definition.rows.map((row, rowIndex) => {
  const cells = row.map(renderTableCell).join('');
  return `<a:tr h="${definition.rowHeights[rowIndex]}">${cells}</a:tr>`;
}).join('');
```

Do not change grid values, cell text, margins, borders, namespaces, or graphic-frame serialization.

- [ ] **Step 7: Add exhaustive invalid row-height cases**

Add:

```ts
const sparseHeights = Array(3);
sparseHeights[0] = 1;
sparseHeights[2] = 1;
const extraHeights = Object.assign([1, 1, 1], { extra: true });
const symbolHeights = Object.assign([1, 1, 1], { [Symbol('extra')]: true });
const accessorHeights = [1, 1, 1];
let heightAccessorCalls = 0;
Object.defineProperty(accessorHeights, '1', {
  get() { heightAccessorCalls += 1; return 1; },
  enumerable: true,
  configurable: true,
});

const invalidRowHeights = [
  null, false, '', {}, new Uint32Array([1, 1, 1]),
  [], [1], [1, 1], [1, 1, 1, 1],
  sparseHeights, extraHeights, symbolHeights, accessorHeights,
  [1, '2', 3], [1, null, 3], [1, [2], 3],
  [1, Number.NaN, 3], [1, Number.POSITIVE_INFINITY, 3],
  [1, Number.MAX_SAFE_INTEGER + 2, 3], [1, 0, 3], [1, -1, 3],
];
for (const rowHeights of invalidRowHeights) {
  expect(() => normalizeTableDefinition(
    [['A'], ['B'], ['C']],
    { rowHeights },
  )).toThrow();
}
expect(heightAccessorCalls).toBe(0);
expect(() => normalizeTableDefinition(
  [['A'], ['B']],
  { rowHeights: [Number.MAX_SAFE_INTEGER, 1] },
)).toThrow(RangeError);
expect(() => normalizeTableDefinition(
  [['A'], ['B']],
  { height: 4, rowHeights: [1, 2] },
)).toThrow(RangeError);
```

- [ ] **Step 8: Run pure focused tests**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts
```

Expected: all internal normalization/render tests pass and omitted behavior remains exact.

---

### Task 2: Expose row heights through native creation lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `NormalizedTableDefinition.rowHeights` and current `SlideModel.addTable()` transaction.
- Produces: exported `AddTableOptions.rowHeights` and end-to-end unequal-row creation proof.

- [ ] **Step 1: Extend the public option**

```ts
export interface AddTableOptions {
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly rowHeights?: number | readonly number[];
}
```

No index export change is needed because `slide.ts` is already exported through model and aggregate SDK packages.

- [ ] **Step 2: Extend model-level exact XML and atomicity coverage**

Change the existing unequal-column table to omit overall height and add:

```ts
rowHeights: [inches(0.5), inches(1.5)],
```

Require `unequal.transform.height === inches(2)`, row values `[457200, 1371600]`, grid values `[914400, 1828800, 2743200]`, xfrm `cx="5486400" cy="1828800"`, six cells, and unchanged ordering. Add public invalid calls for height mismatch and an accessor row array, compare exact slide bytes and mutation journal before/after, and require the accessor call count to remain zero.

- [ ] **Step 3: Extend SDK lifecycle coverage**

Create the existing 3x3 revenue table with:

```ts
height: inches(2.25),
columnWidths: [inches(2), inches(4), inches(2)],
rowHeights: [inches(0.5), inches(0.75), inches(1)],
```

Require initial, duplicate, and reopened slide XML to contain row vector `[457200, 685800, 914400]`, grid vector `[1828800, 3657600, 1828800]`, and `cx="7315200" cy="2057400"`. Repeat existing cell edits, rollback, stable identity, duplicate isolation, validator, and only-target-part checks unchanged.

- [ ] **Step 4: Run model/SDK focused tests and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
```

Expected: declarations expose the new option and every lifecycle assertion passes.

---

### Task 3: Prove PptxGenJS parity and packed consumer coverage

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: public PptxGenJS 4.0.1 `rowH`, native `rowHeights`, aggregate Node/browser exports, and generated declarations.
- Produces: public-output conformance, explicit mismatch repair proof, and actual installed-package coverage.

- [ ] **Step 1: Add exact array conformance to the basic table case**

Generate the public PptxGenJS 2x3 table with:

```ts
{ x: 1, y: 1.5, w: 6, h: 2, colW: [1, 2, 3], rowH: [0.75, 1.25] }
```

Create the native counterpart with:

```ts
{
  x: inches(1),
  y: inches(1.5),
  width: inches(6),
  height: inches(2),
  columnWidths: [inches(1), inches(2), inches(3)],
  rowHeights: [inches(0.75), inches(1.25)],
}
```

Require both row vectors to equal `[685800, 1143000]`, both grid vectors to remain `[914400, 1828800, 2743200]`, both xfrm dimensions to remain exact, and existing text/margin/border/reopen assertions to remain equal.

- [ ] **Step 2: Capture the omitted-height repair**

Create a public PptxGenJS three-row table with `rowH: [0.5, 1, 1.5]` and no `h`, then create a native table with `rowHeights: [inches(0.5), inches(1), inches(1.5)]` and no `height`. Assert both row vectors equal `[457200, 914400, 1371600]`; PptxGenJS 4.0.1 retains xfrm height 1 inch while native derives 3 inches. Label the test as an intentional compatibility repair.

- [ ] **Step 3: Extend Node and browser smoke**

For Node, extend the existing two-row table:

```js
const createdTable = tableSlide.addTable(
  [['Region', 'Revenue'], ['East', '']],
  {
    name: 'Created smoke table',
    columnWidths: [inches(1), inches(3)],
    rowHeights: [inches(0.5), inches(1.5)],
  },
);
```

Require derived width 4 inches, derived height 2 inches, unequal grid and row values in the slide part, cell edit, write/reopen, and the existing adversarial table isolation. Keep `tableCreation: true` and `tableColumnWidths: true`; add `tableRowHeights: true`.

For browser, keep scalar column widths and add scalar `rowHeights: inches(0.75)` on two rows. Require derived width 2.5 inches, derived height 1.5 inches, both row values `685800`, fill edit, and write/reopen.

- [ ] **Step 4: Extend installed declaration smoke**

```ts
const tableOptions: AddTableOptions = {
  name: 'Typed table',
  x: inches(1),
  columnWidths: [inches(1), inches(3)],
  rowHeights: [inches(0.5), inches(1.5)],
};
```

Keep `tableOptions` and the returned `TableModel` in the final `void [...]` expression so strict installed-package compilation checks them.

- [ ] **Step 5: Run focused conformance**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
```

Expected: focused tests pass without reading PptxGenJS private fields.

---

### Task 4: Document the partial sizing slice

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: final option name, units, strict conflict rules, auto-row-height exception, and PptxGenJS baseline evidence.
- Produces: accurate examples without claiming existing-table row-height editing or auto-page support.

- [ ] **Step 1: Update API documentation**

Add to the table example:

```ts
rowHeights: [inches(0.5), inches(0.75), inches(1)],
```

Document scalar repetition, exact array length, derived overall height, matching explicit height, descriptor-safe validation, and explicit xfrm/row consistency. State that omission still produces automatic `tr@h=0`, `TableRow` does not expose height, and inherited `setTransform({ height })` does not resize rows.

- [ ] **Step 2: Update compatibility matrix and baseline**

Change the table row status to include `rowH` as partial support. Record PptxGenJS single-array coercion, short/falsy fallback, long-array truncation, and xfrm/rows inconsistency; explicitly state the native repairs and that existing-table row editing remains pending.

- [ ] **Step 3: Update npm README and changelog**

Add one concise Unreleased bullet and show `rowHeights` in the package example. Preserve the remaining unsupported list.

- [ ] **Step 4: Run docs checks**

```sh
git diff --check
rg -n "row heights|rowHeights|rowH|existing-table|部分支持" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
```

Expected: no claim says existing-table row editing or table sizing as a whole is complete.

---

### Task 5: Full gates, PPTX QA, review, commit, and push

**Files:**
- Review every Task 1-4 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed source/tests/docs, built aggregate package, repository CLI, LibreOffice, Poppler, and overflow helper.
- Produces: reviewed `feat: support table row heights on creation` commit synchronized to `origin/main`.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

- [ ] **Step 2: Build and smoke the actual npm tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_row_package_dir=$(mktemp -d /tmp/pptx-table-rows-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_row_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_row_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected output includes `tableCreation: true`, `tableColumnWidths: true`, `tableRowHeights: true`, `types: true`, and CLI `0.1.0`.

- [ ] **Step 3: Generate native and PptxGenJS QA decks**

Create `/tmp/pptx-table-row-heights-qa.mjs` with `apply_patch`. Use the same 3x3 matrix, x=1, y=1, `[2, 4, 2]` inch columns, and `[0.5, 1, 1.5]` inch rows. Native omits overall height; PptxGenJS supplies `h: 3` with `rowH: [0.5, 1, 1.5]`. Apply blue header fill, white header borders, one vertical cell, bottom alignment, and an empty cell through the same public APIs used by prior table QA.

Use this complete script:

```js
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import {
  PptxDocument,
  TableModel,
  inches,
} from 'file:///Users/jeremy/workspace/pptx/packages/pptx/dist/index.js';

const require = createRequire('/Users/jeremy/workspace/pptx/packages/pptxgenjs-adapter/package.json');
const PptxGenJS = require('pptxgenjs');
const nativeDirectory = '/tmp/pptx-table-row-heights-native';
const baselineDirectory = '/tmp/pptx-table-row-heights-baseline';
await mkdir(nativeDirectory, { recursive: true });
await mkdir(baselineDirectory, { recursive: true });

const rows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];
const expectedWidths = [inches(2), inches(4), inches(2)];
const expectedHeights = [inches(0.5), inches(1), inches(1.5)];

const native = PptxDocument.create({ slideSize: 'wide' });
const nativeSlide = native.addSlide();
const nativeTable = nativeSlide.addTable(rows, {
  name: 'Unequal revenue rows',
  x: inches(1),
  y: inches(1),
  columnWidths: expectedWidths,
  rowHeights: expectedHeights,
});
for (let column = 0; column < 3; column += 1) {
  nativeTable.setCellFill(0, column, {
    kind: 'solid',
    color: { kind: 'srgb', value: '4472C4' },
  });
  nativeTable.setCellBorders(0, column, {
    kind: 'line',
    color: { kind: 'srgb', value: 'FFFFFF' },
    width: 1,
    style: 'solid',
  });
}
nativeTable.setCellTextDirection(1, 2, 'vert');
nativeTable.setCellVerticalAlignment(2, 0, 'bottom');
assert.equal(nativeTable.transform.width, inches(8));
assert.equal(nativeTable.transform.height, inches(3));
const nativeXml = new TextDecoder().decode(
  native.opcPackage.requirePart(nativeSlide.partUri).bytes,
);
assert.deepEqual(
  [...nativeXml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((match) => Number(match[1])),
  expectedWidths,
);
assert.deepEqual(
  [...nativeXml.matchAll(/<a:tr h="(\d+)">/g)].map((match) => Number(match[1])),
  expectedHeights,
);
const nativePath = `${nativeDirectory}/native.pptx`;
const nativeBytes = await native.write();
await writeFile(nativePath, nativeBytes);

const reopened = await PptxDocument.open(nativePath);
const reopenedTable = reopened.slides[0].shapes[0];
assert(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.transform.height, inches(3));
assert.deepEqual(
  reopenedTable.rows.map(({ cells }) => cells.map(({ text }) => text)),
  rows,
);
assert.equal(reopenedTable.rows[1].cells[2].textDirection, 'vert');
assert.equal(reopenedTable.rows[2].cells[0].verticalAlignment, 'bottom');
const reopenedBytes = await reopened.write();
assert.deepEqual(reopenedBytes, nativeBytes);
await writeFile(`${nativeDirectory}/native-reopened.pptx`, reopenedBytes);

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
const baselineSlide = baseline.addSlide();
const baselineRows = rows.map((row, rowIndex) => row.map((text, columnIndex) => ({
  text,
  options: {
    ...(rowIndex === 0 ? {
      fill: { color: '4472C4' },
      border: { type: 'solid', color: 'FFFFFF', pt: 1 },
    } : {}),
    ...(rowIndex === 1 && columnIndex === 2 ? { textDirection: 'vert' } : {}),
    ...(rowIndex === 2 && columnIndex === 0 ? { valign: 'bottom' } : {}),
  },
})));
baselineSlide.addTable(baselineRows, {
  x: 1,
  y: 1,
  w: 8,
  h: 3,
  colW: [2, 4, 2],
  rowH: [0.5, 1, 1.5],
  margin: [0.05, 0.1, 0.05, 0.1],
  border: { type: 'none' },
});
await writeFile(
  `${baselineDirectory}/baseline.pptx`,
  await baseline.write({ outputType: 'uint8array', compression: false }),
);
```

Require native/reopened exact grid `[1828800, 3657600, 1828800]`, exact rows `[457200, 914400, 1371600]`, transform 8x3 inches, matrix and cell edits, and byte-identical reopen/write.

- [ ] **Step 4: Validate and diff packages**

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-row-heights-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-row-heights-native/native-reopened.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-row-heights-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-row-heights-native/native.pptx /tmp/pptx-table-row-heights-native/native-reopened.pptx
```

Expected: all validators report 0 errors/0 warnings and native vs reopened has no changed parts.

- [ ] **Step 5: Render, inspect, and run overflow checks**

Convert native and baseline PPTX to PDF with bundled LibreOffice, render every page to PNG with Poppler at 180 DPI, and inspect each PNG at original resolution. Require visibly unequal 1:2:3 rows and 1:2:1 columns, exact table bounds, header fill/borders, text order, empty cell, vertical/bottom alignment, and no repair, clipping, overlap, blur, missing cell, or off-slide content.

Run:

```sh
mkdir -p \
  /tmp/pptx-table-row-heights-native-pdf-qa \
  /tmp/pptx-table-row-heights-baseline-pdf-qa \
  /tmp/pptx-table-row-heights-native-render-qa \
  /tmp/pptx-table-row-heights-baseline-render-qa
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-row-heights-native-pdf-qa \
  /tmp/pptx-table-row-heights-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-row-heights-baseline-pdf-qa \
  /tmp/pptx-table-row-heights-baseline/baseline.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-row-heights-native-pdf-qa/native.pdf \
  /tmp/pptx-table-row-heights-native-render-qa/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-row-heights-baseline-pdf-qa/baseline.pdf \
  /tmp/pptx-table-row-heights-baseline-render-qa/slide
```

Run:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-row-heights-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-row-heights-baseline/baseline.pptx
```

Expected: both pass with no overflow.

- [ ] **Step 6: Perform final source review**

Verify descriptor-safe array reading, zero accessor calls, detached vectors, positive/safe rounding, overflow-safe sum, exact height conflict handling, unchanged auto/omitted behavior, explicit xfrm/row equality, atomic failure, stable identity, duplicate/reopen isolation, public PptxGenJS-only conformance, and partial-support docs. Confirm `.pnpm-store/` is still only untracked.

- [ ] **Step 7: Stage only intended paths, commit, push, and verify remote**

```sh
git add -- \
  CHANGELOG.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  packages/model/src/model.test.ts \
  packages/model/src/slide.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-create.internal.ts \
  packages/pptx/README.md \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/sdk/src/index.test.ts \
  scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "feat: support table row heights on creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
  git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
  git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: only the 11 intended implementation paths are committed, `.pnpm-store/` remains untracked, push succeeds, and remote divergence is `0 0`.
