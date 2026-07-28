# Table Column Widths Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend native `SlideModel.addTable()` with exact scalar or per-column EMU widths while keeping table xfrm width equal to the OOXML grid sum.

**Architecture:** Extend the focused table creation normalizer so it produces one detached exact `columnWidths` vector for both existing equal-distribution behavior and the new explicit-width path. The renderer consumes that vector directly; `SlideModel.addTable()` keeps its current transaction, shape-tree insertion, identity, and rollback behavior unchanged.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless source-span OOXML editing, PptxGenJS 4.0.1 public-output conformance, tsup, npm tarball smoke, repository CLI, LibreOffice headless, Poppler rendering.

## Global Constraints

- Public input is `AddTableOptions.columnWidths?: number | readonly number[]`, always in EMU.
- A scalar repeats exactly for every physical grid column; an array must be dense, descriptor-safe, detached, and exactly match the rectangular matrix column count.
- Every normalized item must be a positive safe-integer EMU after `Math.round()`; the vector sum must also be a positive safe integer.
- When `columnWidths` is present and `width` is omitted, overall table width is the exact vector sum.
- When both are present, rounded `width` must equal the vector sum or the call throws before package mutation.
- When `columnWidths` is omitted, preserve current default/explicit equal distribution byte-for-byte.
- Always preserve `xfrm width === sum(gridCol widths)`; do not reproduce PptxGenJS scalar floor, mismatch fallback, string coercion, warnings, or xfrm/grid inconsistency.
- Do not add row heights, cell objects, rich cell text, merge, styles, auto-page, repeated headers, hyperlinks, content measurement, existing-table grid editing, or transform/grid resize behavior in this slice.
- Preserve strict matrix/options validation, only-target-slide mutation, stable identity, duplicate isolation, nested transaction rollback, write/reopen, and all existing cell editors.
- Implement inline without subagent delegation for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Commit and push implementation only after source review, typecheck, focused/full tests, performance, actual tarball smoke, CLI validation, PowerPoint 2010 validation, LibreOffice rendering, overflow checks, and full-size visual inspection pass.

---

### Task 1: Normalize exact column-width vectors and render them directly

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Test: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing descriptor-safe `readDenseArray()`, `normalizeCoordinate()`, `distributeTableDimension()`, and `NormalizedTableDefinition`.
- Produces: `NormalizedTableDefinition.columnWidths`, exact scalar/array normalization, and renderer output that uses the normalized vector.

- [ ] **Step 1: Extend pure tests with red scalar/array cases**

Add assertions to `table-create.internal.test.ts` before implementation:

```ts
const unequalSource = [914_400.2, 1_828_799.7, 2_743_200];
const unequal = normalizeTableDefinition(
  [['A', 'B', 'C']],
  { columnWidths: unequalSource },
);
expect(unequal).toMatchObject({
  width: 5_486_400,
  columnWidths: [914_400, 1_828_800, 2_743_200],
});
expect(renderTableGraphicFrame(4, unequal)).toContain(
  '<a:gridCol w="914400"/><a:gridCol w="1828800"/><a:gridCol w="2743200"/>',
);
unequalSource[0] = 1;
expect(unequal.columnWidths[0]).toBe(914_400);

const scalar = normalizeTableDefinition(
  [['A', 'B', 'C']],
  { columnWidths: 1_143_000 },
);
expect(scalar.width).toBe(3_429_000);
expect(scalar.columnWidths).toEqual([1_143_000, 1_143_000, 1_143_000]);

const matching = normalizeTableDefinition(
  [['A', 'B', 'C']],
  { width: 5_486_400, columnWidths: [914_400, 1_828_800, 2_743_200] },
);
expect(matching.width).toBe(5_486_400);
```

Update existing default/explicit expected definitions to include their equal-distribution `columnWidths` arrays.

- [ ] **Step 2: Run the pure test and verify red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts
```

Expected: assertions fail because `columnWidths` is not accepted or returned yet.

- [ ] **Step 3: Extend option and normalized definition types**

In `table-create.internal.ts`, add the option key and normalized field:

```ts
const OPTION_KEYS = ['name', 'x', 'y', 'width', 'height', 'columnWidths'] as const;

export interface NormalizedTableDefinition {
  readonly rows: readonly (readonly string[])[];
  readonly name?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly autoRowHeight: boolean;
  readonly columnWidths: readonly number[];
}
```

- [ ] **Step 4: Implement strict scalar/array normalization and safe summation**

Add focused helpers:

```ts
function normalizeColumnWidths(value: unknown, columnCount: number): readonly number[] {
  const values = Array.isArray(value)
    ? readDenseArray(value, 'Table columnWidths')
    : Array.from({ length: columnCount }, () => value);
  if (values.length !== columnCount) {
    throw new TypeError('Table columnWidths must match the table column count');
  }
  return values.map((item, index) => {
    const width = normalizeCoordinate(item, `Table columnWidths ${index}`);
    if (width <= 0) {
      throw new RangeError(`Table columnWidths ${index} must be greater than zero`);
    }
    return width;
  });
}

function sumColumnWidths(widths: readonly number[]): number {
  return widths.reduce((sum, width) => {
    if (width > Number.MAX_SAFE_INTEGER - sum) {
      throw new RangeError('Table columnWidths sum must fit a safe integer EMU value');
    }
    return sum + width;
  }, 0);
}
```

After rows/options normalization, replace the current width-only path with:

```ts
const requestedColumnWidths = normalizedOptions.columnWidths;
let width: number;
let columnWidths: readonly number[];
if (requestedColumnWidths !== undefined) {
  columnWidths = normalizeColumnWidths(requestedColumnWidths, columnCount);
  const columnWidthSum = sumColumnWidths(columnWidths);
  if (normalizedOptions.width === undefined) {
    width = columnWidthSum;
  } else {
    width = normalizeCoordinate(normalizedOptions.width, 'Table width');
    if (width !== columnWidthSum) {
      throw new RangeError('Table width must equal the sum of columnWidths');
    }
  }
} else {
  const defaultWidth = columnCount * EMU_PER_INCH;
  if (!Number.isSafeInteger(defaultWidth)) {
    throw new RangeError('Table default width must fit a safe integer EMU value');
  }
  width = normalizedOptions.width === undefined
    ? defaultWidth
    : normalizeCoordinate(normalizedOptions.width, 'Table width');
  if (width < columnCount) {
    throw new RangeError('Table width must provide at least one EMU per column');
  }
  columnWidths = distributeTableDimension(width, columnCount);
}
```

Return `columnWidths` in the normalized definition. This preserves the old minimum-one-EMU rule only for equal distribution; explicit widths already require every item to be positive.

- [ ] **Step 5: Render the exact normalized vector**

Replace renderer-side width distribution:

```ts
const grid = definition.columnWidths
  .map((width) => `<a:gridCol w="${width}"/>`)
  .join('');
```

Do not change row height, text, margins, borders, namespaces, or graphic-frame serialization.

- [ ] **Step 6: Add exhaustive invalid cases**

Add tests for:

```ts
const sparseWidths = Array(3);
sparseWidths[0] = 1;
sparseWidths[2] = 1;
const extraWidths = Object.assign([1, 1, 1], { extra: true });
const symbolWidths = Object.assign([1, 1, 1], { [Symbol('extra')]: true });
const accessorWidths = [1, 1, 1];
let widthAccessorCalls = 0;
Object.defineProperty(accessorWidths, '1', {
  get() { widthAccessorCalls += 1; return 1; },
  enumerable: true,
  configurable: true,
});

const invalidColumnWidths = [
  null, false, '', {}, new Uint32Array([1, 1, 1]), [], [1], [1, 1], [1, 1, 1, 1],
  sparseWidths, extraWidths, symbolWidths, accessorWidths,
  [1, '2', 3], [1, null, 3], [1, [2], 3],
  [1, Number.NaN, 3], [1, Number.POSITIVE_INFINITY, 3],
  [1, Number.MAX_SAFE_INTEGER + 2, 3], [1, 0, 3], [1, -1, 3],
];
for (const columnWidths of invalidColumnWidths) {
  expect(() => normalizeTableDefinition(
    [['A', 'B', 'C']],
    { columnWidths },
  )).toThrow();
}
expect(widthAccessorCalls).toBe(0);
expect(() => normalizeTableDefinition(
  [['A', 'B']],
  { columnWidths: [Number.MAX_SAFE_INTEGER, 1] },
)).toThrow(RangeError);
expect(() => normalizeTableDefinition(
  [['A', 'B']],
  { width: 4, columnWidths: [1, 2] },
)).toThrow(RangeError);
```

Also require `{ columnWidths: undefined }` to preserve the current default vector.

- [ ] **Step 7: Run pure focused tests**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts
```

Expected: all internal normalization/render tests pass.

---

### Task 2: Expose column widths through native creation lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `NormalizedTableDefinition.columnWidths` and current `SlideModel.addTable()` transaction.
- Produces: exported `AddTableOptions.columnWidths` and end-to-end unequal-column creation proof.

- [ ] **Step 1: Extend the public option**

Add to `AddTableOptions`:

```ts
readonly columnWidths?: number | readonly number[];
```

No index export change is needed because `slide.ts` is already exported through the model and aggregate SDK.

- [ ] **Step 2: Add model-level exact XML and atomicity coverage**

Extend the existing basic-table model test with a second table:

```ts
const unequal = slide.addTable(
  [['A', 'B', 'C'], ['D', 'E', 'F']],
  {
    name: 'Unequal columns',
    columnWidths: [inches(1), inches(2), inches(3)],
    height: inches(2),
  },
);
expect(unequal.transform.width).toBe(inches(6));
```

Read the slide XML and isolate the `Unequal columns` graphic frame. Require grid values `[914400, 1828800, 2743200]`, xfrm `cx="5486400"`, exact sum, and unchanged row/cell XML. Add public invalid calls for width mismatch and accessor arrays, comparing exact slide bytes and mutation journal before/after.

- [ ] **Step 3: Extend SDK lifecycle coverage**

Create the existing 3x3 revenue table with:

```ts
columnWidths: [inches(2), inches(4), inches(2)],
```

Keep `width: inches(8)` as an equality assertion. Require initial, duplicate, and reopened slide XML to contain the same `[1828800, 3657600, 1828800]` grid vector and `cx="7315200"`. Repeat existing cell edits, rollback, stable identity, duplicate isolation, validator, and only-target-part checks unchanged.

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
- Consumes: public PptxGenJS 4.0.1 `colW`, native `columnWidths`, aggregate Node/browser exports, and generated declarations.
- Produces: public-output conformance and actual installed-package proof.

- [ ] **Step 1: Extend the PptxGenJS test interface**

Continue using `Record<string, unknown>` table options; no private fields or runtime source imports are allowed.

- [ ] **Step 2: Add exact array conformance**

Generate a PptxGenJS 2x3 table using only public API:

```ts
generatedSlide.addTable(objectRows, {
  x: 1,
  y: 1.5,
  w: 6,
  h: 2,
  colW: [1, 2, 3],
});
```

Create the native counterpart with:

```ts
const nativeTable = native.addSlide().addTable(rows, {
  x: inches(1),
  y: inches(1.5),
  width: inches(6),
  height: inches(2),
  columnWidths: [inches(1), inches(2), inches(3)],
});
```

Require both grid vectors to equal `[914400, 1828800, 2743200]`, both xfrm widths to equal `5486400`, and existing text/margin/border/row-height/reopen assertions to remain equal.

- [ ] **Step 3: Capture the scalar floor repair**

Create a public PptxGenJS three-column table with `colW: 1.25` and a native table with `columnWidths: inches(1.25)`. Assert PptxGenJS 4.0.1 produces three 1-inch grid columns due to its floor branch, while native produces three `1_143_000` columns and exact total `3_429_000`. Label the test as an intentional compatibility repair, not exact XML parity.

- [ ] **Step 4: Extend Node and browser smoke**

For Node, create a two-column table with an array and no overall width:

```js
const createdTable = tableSlide.addTable(
  [['Region', 'Revenue'], ['East', '']],
  { name: 'Created smoke table', columnWidths: [inches(1), inches(3)] },
);
```

Require returned transform width `inches(4)`, unequal grid values in the slide part, cell edit, write/reopen, and the existing adversarial table isolation. Keep `tableCreation: true` and add `tableColumnWidths: true`.

For browser, use scalar `columnWidths: inches(1.25)` on two columns and require derived width `inches(2.5)`, both grid values, fill edit, and write/reopen.

- [ ] **Step 5: Extend installed declaration smoke**

Use:

```ts
const tableOptions: AddTableOptions = {
  name: 'Typed table',
  x: inches(1),
  columnWidths: [inches(1), inches(3)],
};
const typedTable: TableModel = createdDocument.slides[0].addTable(tableRows, tableOptions);
```

Keep these values in the `void [...]` expression so strict installed-package compilation checks them.

- [ ] **Step 6: Run focused conformance**

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
- Consumes: final option name, units, strict conflict rules, and PptxGenJS baseline evidence.
- Produces: accurate examples without claiming row heights or existing-table width editing.

- [ ] **Step 1: Update API documentation**

Change the table example to:

```ts
columnWidths: [inches(2.5), inches(3.5), inches(2)],
```

Document scalar repetition, exact array length, derived overall width, matching explicit width, descriptor-safe validation, and xfrm/grid consistency. State that `TableModel` does not yet expose read/edit column widths and inherited `setTransform({ width })` does not resize the grid in this slice.

- [ ] **Step 2: Update compatibility matrix and baseline**

Change the table row status to include column widths as partial support. Record PptxGenJS scalar floor, single-array shortcut, mismatched-array fallback, and array xfrm/grid inconsistency. Explicitly state the native repairs and that row heights remain pending.

- [ ] **Step 3: Update npm README and changelog**

Add one concise Unreleased bullet and show `columnWidths` in the package example. Preserve the remaining unsupported list.

- [ ] **Step 4: Run docs checks**

```sh
git diff --check
rg -n "row heights|rowH|columnWidths|colW|partial|部分支持" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
```

Expected: no claim says table sizing as a whole or existing-table column editing is complete.

---

### Task 5: Full gates, PPTX QA, review, commit, and push

**Files:**
- Review every Task 1-4 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed source/tests/docs, built aggregate package, repository CLI, LibreOffice, Poppler, and overflow helper.
- Produces: reviewed `feat: support table column widths on creation` commit synchronized to `origin/main`.

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
table_column_package_dir=$(mktemp -d /tmp/pptx-table-columns-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_column_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_column_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: output includes `tableCreation: true`, `tableColumnWidths: true`, `types: true`, and CLI `0.1.0`.

- [ ] **Step 3: Generate native and PptxGenJS QA decks**

Create `/tmp/pptx-table-column-widths-qa.mjs` with `apply_patch`. Use the same 3x3 matrix, x=1, y=1, height=3 inches, and `[2, 4, 2]` inch columns. Native uses `columnWidths: [inches(2), inches(4), inches(2)]` with omitted overall width; PptxGenJS uses `colW: [2, 4, 2]` plus `w: 8`. Apply blue header fill, white header borders, one vertical cell, bottom alignment, and an empty cell through the same public APIs used by the basic-table QA.

Use this complete script body:

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
const nativeDirectory = '/tmp/pptx-table-column-widths-native';
const baselineDirectory = '/tmp/pptx-table-column-widths-baseline';
await mkdir(nativeDirectory, { recursive: true });
await mkdir(baselineDirectory, { recursive: true });

const rows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];
const expectedWidths = [inches(2), inches(4), inches(2)];

const native = PptxDocument.create({ slideSize: 'wide' });
const nativeSlide = native.addSlide();
const nativeTable = nativeSlide.addTable(rows, {
  name: 'Unequal revenue table',
  x: inches(1),
  y: inches(1),
  columnWidths: expectedWidths,
  height: inches(3),
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
const nativeXml = new TextDecoder().decode(
  native.opcPackage.requirePart(nativeSlide.partUri).bytes,
);
assert.deepEqual(
  [...nativeXml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((match) => Number(match[1])),
  expectedWidths,
);
const nativePath = `${nativeDirectory}/native.pptx`;
await writeFile(nativePath, await native.write());

const reopened = await PptxDocument.open(nativePath);
const reopenedTable = reopened.slides[0].shapes[0];
assert(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.transform.width, inches(8));
assert.deepEqual(
  reopenedTable.rows.map(({ cells }) => cells.map(({ text }) => text)),
  rows,
);
assert.equal(reopenedTable.rows[1].cells[2].textDirection, 'vert');
assert.equal(reopenedTable.rows[2].cells[0].verticalAlignment, 'bottom');
await writeFile(`${nativeDirectory}/native-reopened.pptx`, await reopened.write());

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
  margin: [0.05, 0.1, 0.05, 0.1],
  border: { type: 'none' },
});
await writeFile(
  `${baselineDirectory}/baseline.pptx`,
  await baseline.write({ outputType: 'uint8array', compression: false }),
);
```

Save:

```text
/tmp/pptx-table-column-widths-native/native.pptx
/tmp/pptx-table-column-widths-native/native-reopened.pptx
/tmp/pptx-table-column-widths-baseline/baseline.pptx
```

Assert native/reopened transform width `inches(8)`, exact grid `[1828800, 3657600, 1828800]`, matrix, cell edits, and byte-identical reopen/write.

- [ ] **Step 4: Validate and diff packages**

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-column-widths-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-column-widths-native/native-reopened.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-column-widths-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-column-widths-native/native.pptx /tmp/pptx-table-column-widths-native/native-reopened.pptx
```

Expected: all validators report 0 errors/0 warnings and native vs reopened has no changed parts.

- [ ] **Step 5: Render, inspect, and run overflow checks**

Convert native and baseline PPTX to PDF with the bundled LibreOffice, render every page to PNG with Poppler at 180 DPI, and inspect each PNG at original resolution. Require visibly unequal 25%/50%/25% columns, exact table bounds, header fill/borders, text order, empty cell, vertical/bottom alignment, and no repair, clipping, overlap, blur, missing cell, or off-slide content.

Run:

```sh
mkdir -p \
  /tmp/pptx-table-column-widths-native-pdf-qa \
  /tmp/pptx-table-column-widths-baseline-pdf-qa \
  /tmp/pptx-table-column-widths-native-render-qa \
  /tmp/pptx-table-column-widths-baseline-render-qa
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-column-widths-native-pdf-qa \
  /tmp/pptx-table-column-widths-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-column-widths-baseline-pdf-qa \
  /tmp/pptx-table-column-widths-baseline/baseline.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-column-widths-native-pdf-qa/native.pdf \
  /tmp/pptx-table-column-widths-native-render-qa/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-column-widths-baseline-pdf-qa/baseline.pdf \
  /tmp/pptx-table-column-widths-baseline-render-qa/slide
```

Run:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-column-widths-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-column-widths-baseline/baseline.pptx
```

Expected: both pass with no overflow.

- [ ] **Step 6: Perform final source review**

Verify descriptor-safe array reading, zero accessor calls, detached vectors, positive/safe rounding, overflow-safe sum, exact width conflict handling, unchanged omitted behavior, xfrm/grid equality, atomic failure, stable identity, duplicate/reopen isolation, public PptxGenJS-only conformance, and partial-support docs. Confirm `.pnpm-store/` is still only untracked.

- [ ] **Step 7: Stage only intended paths, commit, push, and verify remote**

```sh
git add -- \
  CHANGELOG.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  packages/model/src/model.test.ts \
  packages/model/src/slide.ts \
  packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/pptx/README.md \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/sdk/src/index.test.ts \
  scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "feat: support table column widths on creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
  git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
  git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; `git status --short` lists only untracked `.pnpm-store/`.
