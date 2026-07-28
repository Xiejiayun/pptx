# Basic Table Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native `SlideModel.addTable()` creation for strict rectangular single-paragraph string matrices with deterministic geometry, immediately returning an editable `TableModel`.

**Architecture:** Add a focused `table-create.internal.ts` module that validates/detaches the matrix and options, distributes integer EMUs without losing totals, and renders one canonical DrawingML table graphic frame. `SlideModel.addTable()` owns the package transaction, unique shape-tree lookup, shared shape-id allocation, schema-order insertion, write-back, and live model resolution; all existing `TableModel` cell and transform mutations operate on the created XML unchanged.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless source-span OOXML editing, PptxGenJS 4.0.1 public-output conformance, tsup, npm tarball smoke, repository CLI, LibreOffice headless, Poppler rendering.

## Global Constraints

- Public API is `slide.addTable(rows, options): TableModel`, where this slice accepts only non-empty strict rectangular `readonly (readonly string[])[]`.
- Cells are single paragraph in this slice: CR or LF is rejected. Empty string is valid and must contain an editable direct `a:t` node.
- `AddTableOptions` owns only `name`, `x`, `y`, `width`, and `height`; geometry is EMU and works with the existing `inches()` helper.
- Defaults are x/y `0.5 inch`, width `columnCount * 1 inch`, and xfrm height `1 inch`; omitted height writes `tr@h="0"`, explicit height is exactly distributed over rows.
- x/y may be negative. All geometry must be finite and round to safe integers; width/height must be positive and large enough to distribute at least one EMU per column/row.
- Normalize and detach all input before mutation. Reject sparse/accessor/extra-key arrays, non-string cells, invalid XML characters, exotic/accessor/symbol/unknown options, and invalid geometry atomically.
- Grid widths and explicit row heights use quotient/remainder distribution; sums must exactly match xfrm width/height.
- Canonical cells materialize 7.2pt left/right and 3.6pt top/bottom margins plus four zero-width noFill borders. They do not materialize fill, anchor, direction, or fit.
- Reuse current rich-text validation/rendering for plain cell runs; use an explicit empty style object so an empty cell still renders an `a:t` node.
- New table XML has deterministic namespaces and no random `p14:modId`, relationship, package part, style id, banding, header flag, merge, or auto-page state.
- Preserve all existing slide/package content, direct `spTree/extLst` ordering, stable `TableModel` identity, duplicate isolation, nested transaction rollback, and every existing cell editor.
- Do not add cell objects, numbers, rich/multi-paragraph cell text, row/column sizing options, styles, merge, row/column mutation, hyperlinks, table styles, auto-page, repeated headers, percent coordinates, or content measurement in this slice.
- Implement inline without subagent delegation for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Commit and push implementation only after source review, typecheck, focused/full tests, performance, actual tarball smoke, CLI validation, PowerPoint 2010 validation, LibreOffice rendering, overflow checks, and full-size visual inspection pass.

---

### Task 1: Strict table input normalization and canonical XML rendering

**Files:**
- Create: `packages/model/src/table-create.internal.ts`
- Create: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: `escapeXmlAttribute`, `normalizeRichText()`, `renderRichTextParagraphs()`, and the established `914_400` EMU/inch unit.
- Produces: `NormalizedTableDefinition`, `normalizeTableDefinition()`, `distributeTableDimension()`, and `renderTableGraphicFrame()` for `SlideModel.addTable()`.

- [ ] **Step 1: Add failing pure normalization and rendering coverage**

In `packages/model/src/table-create.internal.test.ts`, import `normalizeTableDefinition()`, `distributeTableDimension()`, and `renderTableGraphicFrame()` from the not-yet-created module. Normalize a 2x3 matrix:

```ts
const sourceRows = [
  ['A & <1>', '', 'C1'],
  ['A2', 'B2', 'C2'],
];
const sourceOptions = {
  name: 'Table "A"',
  x: 457_200,
  y: 685_800,
  width: 2_743_201,
  height: 1_371_601,
};
const definition = normalizeTableDefinition(sourceRows, sourceOptions);
const xml = renderTableGraphicFrame(7, definition);

expect(definition.rows).toEqual(sourceRows);
expect(definition).toMatchObject({
  x: 457_200,
  y: 685_800,
  width: 2_743_201,
  height: 1_371_601,
});
```

Inspect slide XML and require:

```ts
expect(xml).toContain('<a:gridCol w="914401"/><a:gridCol w="914400"/><a:gridCol w="914400"/>');
expect(xml).toContain('<a:tr h="685801">');
expect(xml).toContain('<a:tr h="685800">');
expect(xml).toContain('<a:t xml:space="preserve">A &amp; &lt;1&gt;</a:t>');
expect(xml).toMatch(/<a:t xml:space="preserve"><\/a:t>/);
expect(xml).toContain('marL="91440" marR="91440" marT="45720" marB="45720"');
expect(xml).toContain('<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>');
```

Mutate `sourceRows[0][0]` and `sourceOptions.width` after normalization; prove `definition` and rendered XML remain detached. Add a default 1x2 definition and assert x/y `457_200`, width `1_828_800`, height `914_400`, two grid columns of `914_400`, and every `tr@h="0"`.

- [ ] **Step 2: Run the model test and verify red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts
```

Expected: import/compile fails because the internal table creation module does not exist.

- [ ] **Step 3: Implement descriptor-safe matrix and option normalization**

Create `packages/model/src/table-create.internal.ts` with these public internal shapes:

```ts
export interface NormalizedTableDefinition {
  readonly rows: readonly (readonly string[])[];
  readonly name?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly autoRowHeight: boolean;
}

export function normalizeTableDefinition(
  rows: unknown,
  options: unknown,
): NormalizedTableDefinition;

export function distributeTableDimension(total: number, count: number): readonly number[];

export function renderTableGraphicFrame(
  id: number,
  definition: NormalizedTableDefinition,
): string;
```

Use constants:

```ts
const EMU_PER_INCH = 914_400;
const DEFAULT_OFFSET = EMU_PER_INCH / 2;
const DEFAULT_HEIGHT = EMU_PER_INCH;
const CELL_MARGIN_HORIZONTAL = 91_440;
const CELL_MARGIN_VERTICAL = 45_720;
```

Read arrays through own property descriptors, never through an accessor:

```ts
function readDenseArray(value: unknown, context: string): readonly unknown[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new TypeError(`${context} must be a non-empty array`);
  }
  const descriptors = Object.getOwnPropertyDescriptors(value);
  const allowed = new Set(['length', ...Array.from({ length: value.length }, (_, i) => String(i))]);
  for (const key of Reflect.ownKeys(descriptors)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
  }
  return Array.from({ length: value.length }, (_, index) => {
    const descriptor = descriptors[String(index)];
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} must be dense and contain only data items`);
    }
    return descriptor.value;
  });
}
```

Apply it to outer rows and each row. Require equal non-zero row lengths; require every cell to be a string with no `\r`, `\n`, or XML 1.0 forbidden control character. Copy strings into fresh frozen-independent arrays.

For options, require `Object.getPrototypeOf(value)` to be `Object.prototype` or `null`, reject arrays, symbol keys, accessors, and every key outside `name/x/y/width/height`. Read only data descriptor values. Treat absent or data-value `undefined` as omitted. Validate name and geometry before returning. Use:

```ts
function normalizeCoordinate(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(`${context} must round to a safe integer EMU value`);
  }
  return rounded;
}
```

Default width to `columnCount * EMU_PER_INCH`, rejecting an unsafe product. Default height to one inch with `autoRowHeight: true`; any numeric height uses `autoRowHeight: false`. Reject width `< columnCount` and explicit height `< rowCount` so every distributed explicit dimension remains positive.

- [ ] **Step 4: Implement exact integer distribution**

Implement:

```ts
export function distributeTableDimension(total: number, count: number): readonly number[] {
  const quotient = Math.floor(total / count);
  const remainder = total % count;
  return Array.from({ length: count }, (_, index) => quotient + (index < remainder ? 1 : 0));
}
```

The normalizer guarantees positive safe `total` and positive integer `count`. Unit assertions must cover exact division and `2_743_201 / 3 -> [914_401, 914_400, 914_400]`; do not use floating-point inches inside distribution.

- [ ] **Step 5: Implement deterministic table XML rendering**

Render a cell paragraph by reusing the existing engine and forcing an empty `a:t` node:

```ts
function renderCellText(text: string): string {
  return renderRichTextParagraphs(normalizeRichText([
    { runs: [{ text, style: {} }] },
  ]));
}
```

Render canonical no-border sides in schema order:

```ts
const noBorders = ['lnL', 'lnR', 'lnT', 'lnB']
  .map((tag) => `<a:${tag} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:${tag}>`)
  .join('');
```

Each cell must be:

```xml
<a:tc>
  <a:txBody><a:bodyPr/><a:lstStyle/>...</a:txBody>
  <a:tcPr marL="91440" marR="91440" marT="45720" marB="45720">...</a:tcPr>
</a:tc>
```

Build `tblGrid` from `distributeTableDimension(width, columnCount)`. Build row heights from explicit distribution or zeros. Render the graphic frame exactly with deterministic `p`/`a` namespace declarations, escaped `name ?? \`Table ${id}\``, `graphicFrameLocks noGrp="1"`, `p:nvPr/`, xfrm, table graphicData URI, empty `a:tblPr`, grid, and rows. Do not add `p14:modId`, relationships, style id, fill, table flags, or extra package state.

- [ ] **Step 6: Add exhaustive invalid normalization tests**

Test these matrices directly against the pure normalizer:

```ts
const sparseOuter = Array(1);
const sparseRow = [Array(2)];
sparseRow[0]![0] = 'A';
const extraOuter = Object.assign([['A']], { extra: true });
const extraRow = [Object.assign(['A'], { extra: true })];
const accessorRow = [['A']];
let accessorCalls = 0;
Object.defineProperty(accessorRow[0]!, '0', {
  get() { accessorCalls += 1; return 'A'; },
  enumerable: true,
  configurable: true,
});

const invalidRows = [
  null, false, '', [], ['A'], [[], []], [['A'], ['B', 'C']],
  sparseOuter, sparseRow, extraOuter, extraRow, accessorRow,
  [[1]], [[null]], [['line\nbreak']], [['carriage\rreturn']], [['bad\u0000xml']],
  Symbol('rows'),
];
```

Require every case to throw and `accessorCalls === 0`. Add options cases for null/array/exotic prototype/symbol/unknown/accessor key; non-string name; invalid XML name; non-number/non-finite/unsafe x/y/width/height; zero/negative/too-small width and explicit height. Public pre-parse atomicity is covered in Task 2.

- [ ] **Step 7: Run focused model tests**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-create.internal.test.ts
```

Expected: every pure normalizer/distributor/renderer test passes independently before public integration.

---

### Task 2: Integrate `SlideModel.addTable()` with transactions and live identity

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 normalizer/renderer, `allocateShapeId()`, source-span insertion, OPC transactions, `decodeShape()`, `TableModel`, and `BaseShapeModel.setTransform()`.
- Produces: exported `AddTableOptions` and `SlideModel.addTable()` returning the created live `TableModel`.

- [ ] **Step 1: Define the public options and import the focused renderer**

In `packages/model/src/slide.ts`, import `TableModel` from `shapes.ts` and the three Task 1 functions/types. Add:

```ts
export interface AddTableOptions {
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
}
```

No `index.ts` edit is needed because `packages/model/src/index.ts` already exports `./slide.js`, and the aggregate SDK/package already re-export model.

- [ ] **Step 2: Implement unique direct shape-tree lookup and insertion**

Add a table-specific helper that requires exactly one direct `sld/cSld/spTree` and at most one direct `extLst`:

```ts
function requireTableShapeTree(xml: LosslessXmlDocument, partUri: string): XmlElement {
  const candidates = xml.elements('spTree').filter((tree) =>
    tree.parent?.localName === 'cSld'
    && tree.parent.parent?.localName === 'sld');
  if (candidates.length !== 1) {
    throw new ModelParseError('Slide must contain exactly one direct shape tree', partUri);
  }
  const extensionLists = directChildren(candidates[0]!, 'extLst');
  if (extensionLists.length > 1) {
    throw new ModelParseError('Slide shape tree contains repeated extension lists', partUri);
  }
  return candidates[0]!;
}
```

Keep existing `addText()` behavior unchanged. Insert a table immediately before the sole direct `extLst`, or append to the tree when absent.

- [ ] **Step 3: Implement `SlideModel.addTable()` atomically**

Add:

```ts
addTable(
  rows: readonly (readonly string[])[],
  options: AddTableOptions = {},
): TableModel {
  return this.presentation.opcPackage.transaction(() => {
    const definition = normalizeTableDefinition(rows, options);
    const { xml } = this.parse();
    const shapeTree = requireTableShapeTree(xml, this.partUri);
    const nextId = allocateShapeId(xml);
    const tableXml = renderTableGraphicFrame(nextId, definition);
    const extensionList = directChildren(shapeTree, 'extLst')[0];
    if (extensionList) xml.replace(extensionList.start, extensionList.start, tableXml);
    else xml.appendChildXml(shapeTree, tableXml);
    this.setXml(xml.serialize());
    const table = this.shapes.find((candidate) => candidate.id === nextId);
    if (!(table instanceof TableModel) || table.kind !== 'table') {
      throw new ModelParseError(`Created table ${nextId} could not be resolved`, this.partUri);
    }
    return table;
  });
}
```

Do not catch errors; nested OPC transactions must restore bytes/journal on render, write-back, or model-resolution failure.

- [ ] **Step 4: Complete model identity, order, malformed structure, and rollback tests**

In `packages/model/src/model.test.ts`, create text/table/text in one slide and assert shared allocator ids `[2, 3, 4]`, shape order, table insertion before opaque `extLst`, and exact preservation of the extension. Require repeated `cSld/spTree` or repeated direct `extLst` to throw `ModelParseError` with unchanged bytes/journal.

Repeat representative Task 1 invalid rows/options through public `slide.addTable()` and compare exact part bytes and mutation journal. Combine invalid input with a slide missing `spTree` and require the input error, proving normalization precedes parsing; combine valid input with the malformed slide and require `ModelParseError` without mutation.

Check identity and outer rollback:

```ts
expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
table.setCellText(0, 0, 'Edited');
table.setTransform({ x: inches(2) });
expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);

let rolledBack: TableModel | undefined;
expect(() => document.transaction(() => {
  rolledBack = slide.addTable([['rollback']]);
  throw new Error('restore table');
})).toThrow('restore table');
expect(() => rolledBack!.rows).toThrow(ModelParseError);
```

- [ ] **Step 5: Add full SDK lifecycle and cross-capability coverage**

In `packages/sdk/src/index.test.ts`, create a 3x3 table from zero with explicit geometry, then immediately apply every existing cell editor:

```ts
const table = slide.addTable([
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', '12%'],
  ['West', '$980K', '8%'],
], {
  name: 'Revenue table',
  x: inches(1),
  y: inches(1.25),
  width: inches(8),
  height: inches(2.25),
});

table.setCellText(1, 0, 'Eastern');
table.setCellTextDirection(1, 1, 'vert270');
table.setCellTextFit(1, 2, 'shrink');
table.setCellVerticalAlignment(2, 0, 'bottom');
table.setCellMargins(2, 1, [2, 4, 6, 8]);
table.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
});
table.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: 'FFFFFF' },
  width: 1,
  style: 'solid',
});
table.setTransform({ x: inches(1.5) });
```

Assert the created defaults first: text matrix, default margins `{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }`, four none borders, and undefined fill/direction/fit/alignment. Then assert edited snapshots, live identity, package validator zero errors, duplicate isolation, nested rollback, write/reopen, and that only target slide XML changed.

- [ ] **Step 6: Run focused model/SDK tests and typecheck**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: model and SDK suites pass; declarations expose `AddTableOptions` and `addTable()` without an explicit export edit.

---

### Task 3: PptxGenJS conformance and packed Node/browser/type coverage

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: public PptxGenJS 4.0.1 `addTable()` output, native `PptxDocument.create()`, aggregate Node/browser exports, and generated declarations.
- Produces: semantic parity evidence for the strict plain matrix subset and consumer-level package proof.

- [ ] **Step 1: Add a public-output PptxGenJS/native conformance test**

Create one PptxGenJS table with object cells and one native table with string cells, both 2x3 at x=1, y=1.5, w=6, h=2 inches. Import the PptxGenJS output only through `importPptxGenJS()` and compare:

```ts
expect(nativeTable.transform).toMatchObject(importedTable.transform);
expect(nativeTable.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual(
  importedTable.rows.map(({ cells }) => cells.map(({ text }) => text)),
);
expect(nativeTable.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual(
  importedTable.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
);
expect(nativeTable.rows.map(({ cells }) => cells.map(({ borders }) => borders))).toEqual(
  importedTable.rows.map(({ cells }) => cells.map(({ borders }) => borders)),
);
```

Parse both slide parts and assert table URI, grid count, grid sum `5_486_400`, two row heights of `914_400`, six physical cells, direct margin attributes, and L/R/T/B noFill order. Also assert native output omits `p14:modId` and its xfrm width equals grid sum. Reopen both through `PptxDocument.open()` and repeat semantic assertions.

- [ ] **Step 2: Replace smoke-test XML injection with one real native-created table where possible**

Keep the existing adversarial manually injected table because it covers unsupported XML preservation. Add a separate native creation smoke:

```ts
const createdTable = tableSlide.addTable(
  [['Header', 'Value'], ['Count', '42']],
  { name: 'Created table', x: inches(1), y: inches(1), width: inches(4), height: inches(1.5) },
);
createdTable.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
});
createdTable.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: 'FFFFFF' },
  width: 1,
});
```

Add `tableCreation: true` only when the returned object is `TableModel`, its text/default margin/default none borders/geometry are correct, edits survive `created.write()` and reopen, and the adversarial table remains unchanged.

- [ ] **Step 3: Add equivalent browser smoke**

In the browser-condition script string, call the same public API without Node globals. Assert returned type, matrix, default margin/borders, geometry, cell edit, write/reopen, and another slide shape remains intact. The check must fail if browser bundling omits `table-create.internal.ts`.

- [ ] **Step 4: Extend the installed-package declaration consumer**

Import `type AddTableOptions`, create typed rows, and use the returned `TableModel`:

```ts
const tableOptions: AddTableOptions = {
  name: 'Typed table',
  x: inches(1),
  y: inches(1),
  width: inches(4),
  height: inches(2),
};
const typedRows: readonly (readonly string[])[] = [['A', 'B'], ['C', 'D']];
const typedTable: TableModel = createdDocument.slides[0]!.addTable(typedRows, tableOptions);
typedTable.setCellText(0, 0, 'Edited');
```

Include these values in the existing `void [...]` expression so `tsc --noEmit` checks all imports and calls.

- [ ] **Step 5: Run focused conformance, smoke source typecheck, and actual tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
```

Expected: all relevant tests pass and the new PptxGenJS comparison uses no private fields.

---

### Task 4: Document the supported subset without overstating full tables

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: final public API/defaults/limits and PptxGenJS baseline evidence.
- Produces: accurate API examples and a compatibility matrix row that remains partial until later table slices land.

- [ ] **Step 1: Add API examples**

Add the exact from-zero flow to `docs/api/README.md`:

```ts
const table = document.addSlide().addTable(
  [['Region', 'Revenue'], ['East', '$1.2M']],
  { x: inches(1), y: inches(1), width: inches(8), height: inches(1.5) },
);
table.setCellText(1, 0, 'Eastern');
```

Document strict non-empty rectangular strings, no line breaks, EMU geometry/defaults, exact distribution, direct cell margin/no-border defaults, returned live model, and immediate compatibility with every existing cell editor.

- [ ] **Step 2: Update the parity matrix precisely**

Add:

```markdown
| `slide.addTable()` plain rectangular string matrix + x/y/w/h | `slide.addTable(string[][], options)` | 部分支持：基础创建已支持；cell objects、rich text、merge、row/column sizing、style 与 auto-page 待补齐 |
```

Explain that native output repairs PptxGenJS's omitted-width xfrm/grid inconsistency and rejects permissive coercion. Do not mark the entire PptxGenJS `TableProps` surface as supported.

- [ ] **Step 3: Update package README and changelog**

Add one concise npm README paragraph showing native table creation and one Unreleased changelog bullet. Preserve the existing cell-editing descriptions; change “table creation remains future work” to enumerate the now-supported basic slice and remaining advanced slices.

- [ ] **Step 4: Run docs and focused source checks**

```sh
git diff --check
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
```

Expected: no whitespace errors, declarations compile, and focused tests pass.

---

### Task 5: Full gates, tarball, PPTX QA, review, commit, and push

**Files:**
- Review every Task 1-4 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed source/tests/docs, built aggregate package, repository CLI, LibreOffice, Poppler, and presentation overflow helper.
- Produces: reviewed `feat: support basic table creation` commit synchronized to `origin/main`.

- [ ] **Step 1: Run full functional and performance gates without pnpm refresh**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: all functional tests pass, only the default performance case is skipped in the full run, and isolated performance passes.

- [ ] **Step 2: Build and smoke the actual npm tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
basic_table_package_dir=$(mktemp -d /tmp/pptx-basic-table-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$basic_table_package_dir"
node ../../scripts/smoke-npm-package.mjs "$basic_table_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declaration consumer, and CLI smoke all pass with `tableCreation: true`.

- [ ] **Step 3: Generate native and PptxGenJS public baseline decks**

Create `/tmp/pptx-basic-table-creation-qa.mjs` with `apply_patch`, using only public APIs. The native deck must create a 3x3 table at x=1, y=1, w=8, h=3 inches, then use existing setters for a blue header fill, white header borders, bottom alignment, one vertical cell, and an empty cell. The PptxGenJS 4.0.1 deck must create the same matrix/geometry/effective formatting through public `addTable()` cell options. Save:

```text
/tmp/pptx-basic-table-native/native.pptx
/tmp/pptx-basic-table-native/native-reopened.pptx
/tmp/pptx-basic-table-baseline/baseline.pptx
```

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
const nativeDirectory = '/tmp/pptx-basic-table-native';
const baselineDirectory = '/tmp/pptx-basic-table-baseline';
await mkdir(nativeDirectory, { recursive: true });
await mkdir(baselineDirectory, { recursive: true });

const rows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];

const native = PptxDocument.create({ slideSize: 'wide' });
const nativeSlide = native.addSlide();
const nativeTable = nativeSlide.addTable(rows, {
  name: 'Revenue table',
  x: inches(1),
  y: inches(1),
  width: inches(8),
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
const nativePath = `${nativeDirectory}/native.pptx`;
await writeFile(nativePath, await native.write());

const reopened = await PptxDocument.open(nativePath);
const reopenedTable = reopened.slides[0].shapes[0];
assert(reopenedTable instanceof TableModel);
assert.deepEqual(
  reopenedTable.rows.map(({ cells }) => cells.map(({ text }) => text)),
  rows,
);
assert.deepEqual(reopenedTable.transform, {
  x: inches(1),
  y: inches(1),
  width: inches(8),
  height: inches(3),
  rotation: 0,
  flipHorizontal: false,
  flipVertical: false,
});
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
  fontSize: 18,
  margin: [0.05, 0.1, 0.05, 0.1],
  border: { type: 'none' },
});
await writeFile(
  `${baselineDirectory}/baseline.pptx`,
  await baseline.write({ outputType: 'uint8array', compression: false }),
);
```

Open the native file with `PptxDocument.open()`, assert matrix/geometry/defaults/edits, and write `native-reopened.pptx`. Run:

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-basic-table-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-basic-table-native/native-reopened.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-basic-table-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-basic-table-native/native.pptx /tmp/pptx-basic-table-native/native-reopened.pptx
```

Expected: all three files report zero errors/warnings; native vs reopened package diff has empty added/removed/changed arrays. Do not require native/PptxGenJS package identity because deterministic namespaces, text defaults, and `p14:modId` intentionally differ.

- [ ] **Step 4: Render and inspect both implementations**

```sh
native_pdf_dir=$(mktemp -d /tmp/pptx-basic-table-native-pdf.XXXXXX)
baseline_pdf_dir=$(mktemp -d /tmp/pptx-basic-table-baseline-pdf.XXXXXX)
native_render_dir=$(mktemp -d /tmp/pptx-basic-table-native-render.XXXXXX)
baseline_render_dir=$(mktemp -d /tmp/pptx-basic-table-baseline-render.XXXXXX)
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir "$native_pdf_dir" /tmp/pptx-basic-table-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir "$baseline_pdf_dir" /tmp/pptx-basic-table-baseline/baseline.pptx
pdftoppm -png "$native_pdf_dir/native.pdf" "$native_render_dir/slide"
pdftoppm -png "$baseline_pdf_dir/baseline.pdf" "$baseline_render_dir/slide"
```

Inspect every PNG at full size. Confirm the 3x3 bounds, exact row/column split, header fill/borders, text order, empty cell, vertical/bottom alignment, and absence of repair, clipping, overlap, blur, missing cells, or off-slide content. Differences limited to documented theme/font defaults are acceptable; do not claim raster equality unless hashes actually match.

Run overflow checks:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-basic-table-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-basic-table-baseline/baseline.pptx
```

Expected: both overflow checks pass.

- [ ] **Step 5: Perform final source review**

Run `git diff --check`, inspect the entire diff, and verify:

- descriptor-safe validation never invokes input accessors and always precedes package mutation;
- string matrix is detached, rectangular, non-empty, and single-paragraph;
- geometry defaults/ranges and quotient/remainder distribution match the design exactly;
- grid/row totals, XML escaping, namespaces, noFill order, margins, and empty `a:t` are correct;
- shape tree/extLst ordering, shared id allocation, transaction rollback, stable identity, duplicate isolation, write/reopen, and existing cell mutations are proven;
- docs say partial table creation and do not imply cell objects, styles, merge, sizing, or auto-page support;
- `.pnpm-store/` remains untracked and untouched.

- [ ] **Step 6: Stage only intended paths, commit, push, and verify remote**

The intended implementation paths are:

```text
CHANGELOG.md
docs/api/README.md
docs/compatibility/pptxgenjs-baseline.md
packages/model/src/model.test.ts
packages/model/src/slide.ts
packages/model/src/table-create.internal.ts
packages/model/src/table-create.internal.test.ts
packages/pptx/README.md
packages/pptxgenjs-adapter/src/index.test.ts
packages/sdk/src/index.test.ts
scripts/smoke-npm-package.mjs
```

Stage exactly those files, run `git diff --cached --check`, inspect the staged diff, then:

```sh
git commit -m "feat: support basic table creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; `git status --short` lists only untracked `.pnpm-store/`.
