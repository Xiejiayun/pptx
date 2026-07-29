# Table Cell Fill Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let native `slide.addTable()` create cells with strict direct no-fill or solid fill through `{ text, options: { fill } }`, while preserving byte-identical output for empty options.

**Architecture:** Harden the existing shared `TableCellFill` normalizer so creation and editing use one descriptor-safe value contract, export its existing encoder, and normalize creation rows into detached internal cell objects carrying optional fill. The table renderer writes an optional fill after the existing four direct borders in one pre-mutation pass; public, SDK, adapter, packed-package, OOXML, and visual tests prove the full lifecycle.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, descriptor inspection, existing native table renderer and fill editor, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarball smoke, repository pptx-inspect CLI, LibreOffice headless, Poppler.

## Global Constraints

- Public creation types are `AddTableCellOptions { readonly fill?: TableCellFill }` and `AddTableCell { readonly text: string; readonly options?: AddTableCellOptions }`.
- `AddTableCellInput` and the existing `SlideModel.addTable()` rows signature remain unchanged.
- String, `{ text }`, empty-options, undefined-fill, and filled cells may be mixed in one strict rectangular matrix.
- Cell/options/fill/color objects must use `Object.prototype` or null prototype and own data properties only; no accessor may be invoked.
- Empty options and undefined fill emit no new XML and remain byte-identical to equivalent string input.
- Fill reuses `TableCellFill`: direct none or solid strict sRGB/theme color with optional finite `0..100` transparency quantized to `0.001%`.
- Do not accept PptxGenJS-shaped `{ color, type, alpha }` directly; do not add table-level fill, gradients, patterns, pictures, group fill, border, margin, alignment, direction, fit, hyperlink, rich text, merge/span, auto-page, or repeated headers.
- Preserve table geometry, row/column sizing, stable identity, all existing editors, package atomicity, and non-target parts.
- Do not stage, remove, inspect recursively, or otherwise modify `.pnpm-store/`.
- Use inline execution in the current task; no subagents.
- Commit and push only after review, typecheck, focused/full tests, performance, actual tarball smoke, CLI/PowerPoint validation, package diffs, rendering, overflow checks, and original-resolution visual inspection pass.

---

### Task 1: Make the shared fill value path descriptor-safe

**Files:**
- Modify: `packages/model/src/table-cell-fill.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: existing `TableCellFill`, `RichTextColor`, read/replace behavior.
- Produces: unchanged `normalizeTableCellFill(value, context)` values plus exported `renderTableCellFill(fill, prefix)`.

- [ ] **Step 1: Add failing shared-normalizer safety coverage through `setCellFill()`**

Extend the existing invalid-fill section in `model.test.ts` with accessor, inherited, class, extra-symbol, and nested-color accessor inputs:

```ts
const accessorFill = {};
const accessorColor = {};
let fillAccessorCalls = 0;
Object.defineProperty(accessorFill, 'kind', {
  get() {
    fillAccessorCalls += 1;
    return 'none';
  },
  enumerable: true,
});
Object.defineProperty(accessorColor, 'value', {
  get() {
    fillAccessorCalls += 1;
    return 'FF0000';
  },
  enumerable: true,
});

class ExoticFill {
  kind = 'none';
}

const invalidValues = [
  accessorFill,
  Object.create({ kind: 'none' }),
  new ExoticFill(),
  Object.assign({ kind: 'none' }, { [Symbol('extra')]: true }),
  { kind: 'solid', color: accessorColor },
];
```

Require every call to throw before mutation and `fillAccessorCalls === 0`. Add one null-prototype solid fill and color to the existing valid lifecycle and require normalized uppercase sRGB output.

- [ ] **Step 2: Run the model test and confirm the accessor/class cases are red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: current direct property reads invoke one accessor or accept an exotic object, so the new assertions fail.

- [ ] **Step 3: Replace direct caller-property reads with one strict data-object reader**

In `table-cell-fill.internal.ts`, add:

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

Use it for the outer fill with `['kind', 'color', 'transparency']` and for color with `['kind', 'value']`. Keep the current kind, color, transparency, quantization, and detached return logic. The `kind: 'none'` branch must still reject copied `color` or `transparency` via `assertKeys(candidate, ['kind'], context)`.

- [ ] **Step 4: Export the existing encoder without changing bytes**

Rename the private encoder and its editor call site:

```ts
export function renderTableCellFill(fill: TableCellFill, prefix: string): string {
  if (fill.kind === 'none') return `<${prefix}noFill/>`;
  const tag = fill.color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const value = escapeXmlAttribute(fill.color.value);
  const color = fill.transparency === undefined
    ? `<${prefix}${tag} val="${value}"/>`
    : `<${prefix}${tag} val="${value}"><${prefix}alpha val="${Math.round(
        (100 - fill.transparency) * 1_000,
      )}"/></${prefix}${tag}>`;
  return `<${prefix}solidFill>${color}</${prefix}solidFill>`;
}
```

`replaceTableCellFill()` must call `renderTableCellFill(fill, prefix)`; no read or insertion behavior changes.

- [ ] **Step 5: Run model tests and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: descriptor safety, null-prototype valid input, all existing fill snapshots/edits/no-ops/rollback tests, and project references pass.

---

### Task 2: Normalize and render cell `options.fill` with TDD

**Files:**
- Modify: `packages/model/src/table-create.internal.test.ts`
- Modify: `packages/model/src/table-create.internal.ts`

**Interfaces:**
- Consumes: Task 1 `normalizeTableCellFill()` and `renderTableCellFill()`.
- Produces:

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly fill?: TableCellFill;
}
```

- [ ] **Step 1: Add failing normalization and exact-output tests**

Create one mixed matrix containing string, `{ text }`, empty options, undefined fill, no-fill, solid sRGB, theme, explicit zero, fractional, and full transparency cells. Require detached normalized snapshots:

```ts
const sourceFill = {
  kind: 'solid' as const,
  color: { kind: 'srgb' as const, value: '#ff0000' },
  transparency: 33.3334,
};
const definition = normalizeTableDefinition([[
  'String',
  { text: 'Plain' },
  { text: 'Empty', options: {} },
  { text: 'Undefined', options: { fill: undefined } },
  { text: 'None', options: { fill: { kind: 'none' } } },
  { text: 'Solid', options: { fill: sourceFill } },
]], undefined);
```

Require the solid fill to normalize to uppercase `FF0000` and transparency `33.333`, then mutate `sourceFill` and prove the definition is unchanged.

Render equal string, `{ text }`, empty-options, and undefined-fill single-cell definitions with the same id/options and require exact XML equality. For filled XML require border order followed by:

```xml
<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="66667"/></a:srgbClr></a:solidFill>
```

Also require `kind: 'none'` to emit a direct cell-level `<a:noFill/>` after all four border elements.

- [ ] **Step 2: Add failing cell-options descriptor matrix**

Reject without getter invocation:

```ts
const invalidCells = [
  { text: 'A', options: null },
  { text: 'A', options: [] },
  { text: 'A', options: new (class Options { fill = undefined; })() },
  { text: 'A', options: { unknown: true } },
  { text: 'A', options: Object.assign({}, { [Symbol('extra')]: true }) },
  { text: 'A', options: { fill: { kind: 'solid' } } },
  { text: 'A', options: { fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } } } },
];
```

Add accessor `options`, accessor `fill`, and nested fill/color accessor cases with a shared invocation counter that must remain zero. Continue rejecting unrelated cell keys such as `{ text, extra }`.

- [ ] **Step 3: Run internal tests and confirm red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts
```

Expected: current exact `{ text }` reader rejects `options`, and normalized rows lack fill state.

- [ ] **Step 4: Introduce the internal normalized cell and strict options reader**

Import `TableCellFill`, `normalizeTableCellFill`, and `renderTableCellFill`. Change `NormalizedTableDefinition.rows` to `readonly (readonly NormalizedTableCell[])[]`.

Refactor cell normalization to return `{ text }` for strings and plain objects. Read cell own data keys from the exact allowed set `['text', 'options']`, require `text`, and normalize options with:

```ts
function normalizeTableCellOptions(
  value: unknown,
  context: string,
): Pick<NormalizedTableCell, 'fill'> {
  if (value === undefined) return {};
  const options = readDataObject(value, `${context} options`, ['fill']);
  const fill = normalizeTableCellFill(options.fill, `${context} fill`);
  return fill === undefined ? {} : { fill };
}
```

Use a local descriptor-safe `readDataObject()` for cell and cell-options objects. Preserve the existing table-options reader behavior and error contexts.

- [ ] **Step 5: Render optional fill through the shared encoder**

Change `renderTableCell()` to accept `NormalizedTableCell`, render paragraphs from `cell.text`, and append:

```ts
const fill = cell.fill === undefined ? '' : renderTableCellFill(cell.fill, 'a:');
return `<a:tc>...<a:tcPr ...>${NO_BORDERS}${fill}</a:tcPr></a:tc>`;
```

Do not add whitespace and do not change unfilled bytes.

- [ ] **Step 6: Run focused internal/model tests and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: all normalization, exact XML, descriptor-safety, geometry, and existing editor tests pass.

---

### Task 3: Expose the public option and prove model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 2 internal cell fill.
- Produces:

```ts
export interface AddTableCellOptions {
  readonly fill?: TableCellFill;
}
```

- [ ] **Step 1: Add public typing and lifecycle assertions first**

Import `AddTableCellOptions` in the model test and type a mutable source fill. Create cells with empty options, direct none, sRGB, scheme, and transparency. Require:

- immediate `TableCell.fill` snapshots equal normalized inputs;
- caller mutation does not change the live table;
- `setCellFill()` can replace/clear a created fill immediately;
- shape identity, transform, column widths, and row heights remain stable;
- invalid options/fill cause zero slide bytes and zero journal changes;
- outer transaction rollback removes the created filled table;
- write/reopen preserves all remaining fills and geometry.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: `AddTableCellOptions` is absent and `AddTableCell` rejects `options`.

- [ ] **Step 3: Add the public option type and widen `AddTableCell` only**

In `slide.ts`, import `type TableCellFill` from `shapes.ts`, then add:

```ts
export interface AddTableCellOptions {
  readonly fill?: TableCellFill;
}

export interface AddTableCell {
  readonly text: string;
  readonly options?: AddTableCellOptions;
}
```

Do not change `AddTableCellInput` or `addTable()` again.

- [ ] **Step 4: Run model typecheck/tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: declarations and full public lifecycle pass.

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
- Consumes: Task 3 aggregate exports and filled-cell creation.
- Produces: SDK/adapter/browser/declaration/package coverage and accurate support matrix.

- [ ] **Step 1: Extend the SDK basic-table lifecycle**

Type one source option as `AddTableCellOptions`, create sRGB/theme/none/transparent cells in the existing mixed matrix, mutate the source after creation, and prove detached snapshots. Keep duplicate isolation, editing, rollback, geometry, non-target part, and reopen assertions; the duplicate must retain original fills after edits to the source table.

- [ ] **Step 2: Add supported PptxGenJS fill creation conformance**

In the existing adapter fill section, create a native table with equivalent supported cells:

```ts
[
  { text: 'Opaque', options: { fill: {
    kind: 'solid', color: { kind: 'srgb', value: 'FF0000' },
  } } },
  { text: 'Theme alpha', options: { fill: {
    kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25,
  } } },
]
```

Require equal final snapshots, direct color tags, `alpha=75000`, margins, borders, geometry, and reopen state against PptxGenJS `{ color, transparency }`. Separately assert the documented direct-state differences: native transparency zero contains `alpha=100000` while PptxGenJS omits alpha; native `kind: 'none'` contains direct cell noFill while PptxGenJS type none has no direct fill.

- [ ] **Step 3: Extend Node/browser/declaration tarball smoke**

Use one filled object cell in both Node and browser created tables. Add `tableCellFillCreation: true` only when creation snapshot, source detachment, edit, write/reopen, sRGB/theme/transparency, and existing cell-object creation all pass.

In declaration smoke import and use:

```ts
type AddTableCellOptions,

const creationFill: AddTableCellOptions = {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
};
```

Keep `creationFill` in the final `void` expression.

- [ ] **Step 4: Update docs and compatibility boundaries**

Add one Unreleased bullet. Update API/package examples to include `{ text, options: { fill } }`. Document descriptor safety, detached normalization, direct none/solid semantics, transparency quantization, empty-options byte identity, PptxGenJS zero/none differences, and remaining unsupported cell options.

- [ ] **Step 5: Run focused conformance and docs checks**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --check
rg -n "AddTableCellOptions|options.*fill|tableCellFillCreation" \
  packages scripts CHANGELOG.md docs
```

Expected: focused tests/typecheck/docs scan pass and unsupported border/margin/merge/etc. remain visible.

---

### Task 5: Full gates, real-deck QA, review, commit, and push

**Files:**
- Review every Task 1-4 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: all implementation/tests/docs, actual packed package, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces: reviewed `feat: create tables with cell fills` commit synchronized to `origin/main`.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-fill-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Parse and record suite/test passed, failed, skipped, and todo totals.

- [ ] **Step 2: Build and smoke the actual npm tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js \
  --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_fill_package_dir=$(mktemp -d \
  /tmp/pptx-table-cell-fill-package.XXXXXX)
npm pack --ignore-scripts \
  --pack-destination "$table_cell_fill_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_cell_fill_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require `tableCellFillCreation: true`, all existing table flags true, declaration success, browser success, and CLI `0.1.0`.

- [ ] **Step 3: Verify repository CLI health**

```sh
node packages/pptx/dist/cli.js --json doctor
```

Require successful offline JSON.

- [ ] **Step 4: Generate native equivalent/filled/edited/reopened and PptxGenJS decks**

Create `/tmp/pptx-table-cell-fill-qa.mjs` with `apply_patch` using this complete program:

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
const output = '/tmp/pptx-table-cell-fill';
await mkdir(output, { recursive: true });

const textRows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];
const tableOptions = {
  name: 'Cell fill creation QA',
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
const emptyOptions = build(textRows.map((row) =>
  row.map((text) => ({ text, options: {} }))));
await writeFile(output + '/string-source.pptx', await strings.document.write());
await writeFile(
  output + '/empty-options-source.pptx',
  await emptyOptions.document.write(),
);

const fillRows = [
  [
    { text: 'Region', options: { fill: {
      kind: 'solid', color: { kind: 'scheme', value: 'accent1' },
    } } },
    { text: 'Revenue', options: { fill: {
      kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' },
    } } },
    { text: 'Growth', options: { fill: {
      kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 25,
    } } },
  ],
  [
    'East',
    { text: '$1.2M', options: { fill: {
      kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' }, transparency: 33.333,
    } } },
    { text: 'Vertical', options: { fill: {
      kind: 'solid', color: { kind: 'srgb', value: 'E2F0D9' }, transparency: 100,
    } } },
  ],
  [
    { text: 'West', options: { fill: { kind: 'none' } } },
    '$980K',
    { text: '', options: {} },
  ],
];
const filled = build(fillRows);
filled.table.setCellTextDirection(1, 2, 'vert270');
filled.table.setCellVerticalAlignment(2, 1, 'bottom');
const sourceFills = filled.table.rows.map(({ cells }) =>
  cells.map(({ fill }) => fill));
await writeFile(output + '/fill-source.pptx', await filled.document.write());

filled.table.setCellFill(2, 2, {
  kind: 'solid',
  color: { kind: 'srgb', value: '70AD47' },
  transparency: 25,
});
filled.table.setCellText(2, 2, 'Edited empty');
const editedBytes = await filled.document.write();
await writeFile(output + '/fill-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[2].cells[2].text, 'Edited empty');
assert.deepEqual(reopenedTable.rows[2].cells[2].fill, {
  kind: 'solid',
  color: { kind: 'srgb', value: '70AD47' },
  transparency: 25,
});
assert.equal(reopenedTable.rows[1].cells[2].textDirection, 'vert270');
assert.equal(reopenedTable.rows[2].cells[1].verticalAlignment, 'bottom');
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/fill-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      { text: 'Region', options: { fill: { color: baseline.SchemeColor.accent1 } } },
      { text: 'Revenue', options: { fill: { color: 'D9EAF7' } } },
      { text: 'Growth', options: {
        fill: { color: baseline.SchemeColor.accent2, transparency: 25 },
      } },
    ],
    [
      { text: 'East', options: {} },
      { text: '$1.2M', options: {
        fill: { color: 'FFF2CC', transparency: 33.333 },
      } },
      { text: 'Vertical', options: {
        fill: { color: 'E2F0D9', transparency: 100 },
      } },
    ],
    [
      { text: 'West', options: { fill: { type: 'none' } } },
      { text: '$980K', options: {} },
      { text: '', options: {} },
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
  sourceFills,
  reopenedFills: reopenedTable.rows.map(({ cells }) =>
    cells.map(({ fill }) => fill)),
  editedText: reopenedTable.rows[2].cells[2].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run:

```sh
node /tmp/pptx-table-cell-fill-qa.mjs
```

Require the stable normalized source/reopened fill matrices, `editedText: "Edited empty"`, column widths `[1828800,3657600,1828800]`, and row heights `[457200,1143000,594360]`.

- [ ] **Step 5: Validate packages and exact diff isolation**

Run PowerPoint 2010 validation on all six decks:

```sh
for deck in \
  string-source empty-options-source fill-source fill-edited fill-reopened \
  pptxgenjs-baseline
do
  node packages/pptx/dist/cli.js --json package validate \
    "/tmp/pptx-table-cell-fill/$deck.pptx" \
    --profile powerpoint-2010
done
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-fill/string-source.pptx \
  /tmp/pptx-table-cell-fill/empty-options-source.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-fill/fill-source.pptx \
  /tmp/pptx-table-cell-fill/fill-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-fill/fill-edited.pptx \
  /tmp/pptx-table-cell-fill/fill-reopened.pptx
```

Require:

- string-source → empty-options-source: zero changed decompressed parts;
- fill-source → fill-edited: only `/ppt/slides/slide1.xml` changes;
- fill-edited → fill-reopened: zero changed decompressed parts;
- every validation: zero errors and zero warnings.

- [ ] **Step 6: Render and inspect native/baseline decks**

Use the bundled LibreOffice and Poppler binaries to render `fill-edited.pptx` and `pptxgenjs-baseline.pptx` at 180 DPI:

```sh
mkdir -p \
  /tmp/pptx-table-cell-fill/rendered-native \
  /tmp/pptx-table-cell-fill/rendered-baseline
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-cell-fill/rendered-native \
  /tmp/pptx-table-cell-fill/fill-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-cell-fill/rendered-baseline \
  /tmp/pptx-table-cell-fill/pptxgenjs-baseline.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-cell-fill/rendered-native/fill-edited.pdf \
  /tmp/pptx-table-cell-fill/rendered-native/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-cell-fill/rendered-baseline/pptxgenjs-baseline.pdf \
  /tmp/pptx-table-cell-fill/rendered-baseline/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-fill/fill-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-fill/pptxgenjs-baseline.pptx
```

Inspect every PNG with original detail and require visible sRGB/theme/transparent fill behavior, direct no-fill cell, all text, unequal rows/columns, existing direction/alignment, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 7: Review, stage only targets, commit, push, and prove synchronization**

```sh
git diff --check
git status --short
git diff --stat
git add \
  packages/model/src/table-cell-fill.internal.ts \
  packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "feat: create tables with cell fills"
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
