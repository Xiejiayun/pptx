# Plain-Text Table Cell Object Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let native `slide.addTable()` accept strict `{ text: string }` cells mixed with strings while preserving byte-identical output for equivalent text.

**Architecture:** Normalize every string or exact ordinary text object into the existing detached `string[][]` table definition, so the renderer and all geometry/editing paths remain unchanged. Export a narrow creation input type from the model surface, then prove equivalent native/PptxGenJS output, packed-consumer behavior, rollback, package stability, and rendering.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, strict descriptor inspection, existing native table renderer, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarball smoke, repository pptx-inspect CLI, LibreOffice headless, Poppler.

## Global Constraints

- Public types are exactly `AddTableCell { readonly text: string }` and `AddTableCellInput = string | AddTableCell`.
- `SlideModel.addTable()` accepts `readonly (readonly AddTableCellInput[])[]` and keeps the existing `AddTableOptions` surface.
- String and object cells may be mixed in any rectangular matrix; empty text remains valid.
- Cell objects must have `Object.prototype` or null prototype, exactly one own key `text`, and an own data descriptor.
- Missing/extra/symbol keys, arrays, null, exotic prototypes, accessors, inherited text, non-string text, CR/LF, and invalid XML characters are rejected before package writes.
- Normalization never retains caller arrays or cell objects and never invokes a JavaScript property getter.
- Equivalent string and `{ text }` inputs normalize to the same `readonly string[][]` and render byte-identical table XML.
- Do not accept or ignore `options`, rich-text arrays, hyperlinks, merge/span properties, styles, row/column mutation, auto-page, or repeated headers in this item.
- Do not change table geometry, existing editing APIs, `TableCell` snapshots, transaction behavior, or renderer defaults.
- Do not stage, remove, inspect recursively, or otherwise modify `.pnpm-store/`.
- Use inline execution in the current task; no subagents.
- Commit and push only after review, typecheck, focused/full tests, performance, actual tarball smoke, CLI/PowerPoint validation, package diffs, rendering, overflow checks, and original-resolution visual inspection pass.

---

### Task 1: Normalize strict plain-text cell objects with TDD

**Files:**
- Modify: `packages/model/src/table-create.internal.test.ts`
- Modify: `packages/model/src/table-create.internal.ts`

**Interfaces:**
- Consumes: existing dense matrix reader and single-paragraph/XML text contract.
- Produces: unchanged `NormalizedTableDefinition.rows: readonly (readonly string[])[]` from string or object inputs.

- [ ] **Step 1: Add failing mixed-object normalization tests**

~~~ts
const objectCell = { text: 'A & <1>' };
const nullPrototype = Object.assign(Object.create(null), { text: 'B1' });
const definition = normalizeTableDefinition([
  [objectCell, nullPrototype, ''],
  ['A2', { text: 'B2' }, { text: 'C2' }],
], {
  columnWidths: [914_400, 1_828_800, 914_400],
  rowHeights: [457_200, 914_400],
});
expect(definition.rows).toEqual([
  ['A & <1>', 'B1', ''],
  ['A2', 'B2', 'C2'],
]);
objectCell.text = 'MUTATED';
nullPrototype.text = 'MUTATED';
expect(definition.rows[0]).toEqual(['A & <1>', 'B1', '']);
~~~

Render one all-string definition and one equivalent all-object definition with the same id/options, then require exact XML equality:

~~~ts
expect(renderTableGraphicFrame(8, objectDefinition))
  .toBe(renderTableGraphicFrame(8, stringDefinition));
~~~

Also assert XML escaping, empty `<a:t>`, default margins, borders, row heights, and column widths remain unchanged.

- [ ] **Step 2: Add failing descriptor and value rejection tests**

Create and reject each cell form:

~~~ts
const inherited = Object.create({ text: 'Inherited' });
const accessor = {};
let accessorCalls = 0;
Object.defineProperty(accessor, 'text', {
  get() {
    accessorCalls += 1;
    return 'Accessor';
  },
  enumerable: true,
});

const invalidCells = [
  null,
  1,
  true,
  [],
  new Date(0),
  new (class Cell { text = 'class'; })(),
  {},
  inherited,
  accessor,
  { text: 'A', options: {} },
  { text: 'A', extra: true },
  Object.assign({ text: 'A' }, { [Symbol('extra')]: true }),
  { text: 1 },
  { text: ['rich'] },
  { text: 'line\nbreak' },
  { text: 'carriage\rreturn' },
  { text: 'bad\u0000xml' },
];
~~~

For every value, call `normalizeTableDefinition([[value]], undefined)` and require `TypeError`. Require `accessorCalls === 0` and keep the existing outer/inner array accessor tests passing.

- [ ] **Step 3: Run the focused internal test and confirm red**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts
~~~

Expected: object cells fail with the current “must be a string” validation.

- [ ] **Step 4: Add one focused cell normalization helper**

Replace the inline string-only branch with:

~~~ts
function normalizeTableCell(
  cell: unknown,
  rowIndex: number,
  columnIndex: number,
): string {
  const context = `Table cell ${rowIndex},${columnIndex}`;
  if (typeof cell === 'string') return normalizeTableCellText(cell, context);
  if (!cell || typeof cell !== 'object' || Array.isArray(cell)) {
    throw new TypeError(`${context} must be a string or text object`);
  }
  const prototype = Object.getPrototypeOf(cell);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary text object`);
  }
  const keys = Reflect.ownKeys(cell);
  if (keys.length !== 1 || keys[0] !== 'text') {
    throw new TypeError(`${context} must contain only text`);
  }
  const descriptor = Object.getOwnPropertyDescriptor(cell, 'text');
  if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
    throw new TypeError(`${context} text must be a data property`);
  }
  return normalizeTableCellText(descriptor.value, context);
}
~~~

Share validation with strings:

~~~ts
function normalizeTableCellText(value: unknown, context: string): string {
  if (typeof value !== 'string') {
    throw new TypeError(`${context} text must be a string`);
  }
  if (/\r|\n/.test(value)) {
    throw new TypeError(`${context} must contain one paragraph`);
  }
  if (containsInvalidXmlCharacter(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value;
}
~~~

Do not alter `renderTableCell()` or normalized row types.

- [ ] **Step 5: Run internal tests**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts
~~~

Expected: mixed/object/equivalent/invalid tests and all geometry tests pass.

---

### Task 2: Expose the public type and model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 unknown-input normalizer.
- Produces:

~~~ts
export interface AddTableCell {
  readonly text: string;
}

export type AddTableCellInput = string | AddTableCell;
~~~

- [ ] **Step 1: Add failing public model lifecycle coverage**

Create a table through the public API:

~~~ts
const cell = { text: 'Object A' };
const table = slide.addTable([
  [cell, 'String B'],
  [{ text: '' }, { text: 'Object D' }],
], {
  columnWidths: [inches(2), inches(3)],
  rowHeights: [inches(0.75), inches(1.25)],
});
expect(table.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
  ['Object A', 'String B'],
  ['', 'Object D'],
]);
cell.text = 'MUTATED';
expect(table.rows[0]!.cells[0]!.text).toBe('Object A');
table.setCellText(1, 0, 'Edited empty');
~~~

Require stable table identity, exact transform geometry, outer transaction rollback, write/reopen persistence, and no mutation for `{ text: 'A', options: {} }` or accessor input.

- [ ] **Step 2: Run model tests and confirm public typing/behavior is red**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
~~~

Expected: typecheck rejects object rows and runtime tests fail before Task 1 implementation/public signature changes are complete.

- [ ] **Step 3: Add the public types and widen only the rows signature**

Place `AddTableCell` and `AddTableCellInput` beside `AddTableOptions`, then change:

~~~ts
addTable(
  rows: readonly (readonly AddTableCellInput[])[],
  options: AddTableOptions = {},
): TableModel
~~~

No cast or overload is needed because `normalizeTableDefinition()` already consumes unknown.

- [ ] **Step 4: Run model tests and typecheck**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
~~~

Expected: declarations expose both types and public lifecycle/rollback tests pass.

---

### Task 3: Prove SDK, PptxGenJS, browser, and packed behavior

**Files:**
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 2 aggregate exports and object-capable `addTable()`.
- Produces: typed mixed-cell, duplicate, PptxGenJS-equivalence, Node/browser, declaration, and CLI coverage.

- [ ] **Step 1: Use typed mixed cells in the SDK lifecycle**

Import `type AddTableCellInput` and define:

~~~ts
const rows: readonly (readonly AddTableCellInput[])[] = [
  [{ text: 'Region' }, 'Revenue', { text: 'Growth' }],
  ['East', { text: '$1.2M' }, '12%'],
  [{ text: 'West' }, '$980K', { text: '' }],
];
const table = slide.addTable(rows, options);
~~~

Keep all existing cell-editor, duplicate-isolation, rollback, geometry, and reopen assertions. Mutate one source object after creation and prove the live table remains detached.

- [ ] **Step 2: Compare native object cells with PptxGenJS object output**

In basic table conformance, create the native table with:

~~~ts
rows.map((row) => row.map((text) => ({ text })))
~~~

PptxGenJS continues to receive `{ text, options: {} }`. Require identical text, geometry, margins, borders, direct row heights/grid widths, and valid reopen behavior. Add one assertion that native XML equals the prior native string-cell XML shape aside from no differences at all for the cell subtree.

- [ ] **Step 3: Extend Node and browser smoke**

Use mixed inputs in both created tables:

~~~js
const createdTable = tableSlide.addTable([
  [{ text: 'Region' }, 'Revenue'],
  ['East', { text: '' }],
], options);
~~~

Keep the existing edits/reopen assertions and add `tableCellObjectCreation: true` to smoke JSON by checking all four normalized cell texts.

- [ ] **Step 4: Extend declaration smoke**

Import `AddTableCell` and `AddTableCellInput`, then add:

~~~ts
const objectCell: AddTableCell = { text: 'Revenue' };
const tableRows: readonly (readonly AddTableCellInput[])[] = [
  ['Region', objectCell],
  [{ text: 'East' }, { text: '' }],
];
~~~

Keep both values in the final `void` expression and preserve all row-height/column-width declarations.

- [ ] **Step 5: Run focused conformance**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
~~~

---

### Task 4: Document the exact supported slice

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: final public names and strict object limitations.
- Produces: accurate docs that keep object options/rich text/merge gaps visible.

- [ ] **Step 1: Update API examples and contract**

~~~ts
const createdTable = document.addSlide().addTable([
  [{ text: 'Region' }, 'Revenue', { text: 'Growth' }],
  ['East', { text: '$1.2M' }, '12%'],
], tableOptions);
~~~

Document exact `{ text: string }`, mixed cells, detached normalization, one-paragraph/XML validation, descriptor safety, byte equivalence, and rejection of `options` or rich arrays.

- [ ] **Step 2: Update compatibility baseline, README, and changelog**

Mark plain-text cell object containers supported. Keep object `options`, rich/multi-paragraph text, hyperlinks, merges, styles, auto-page, repeated headers, and content measurement unsupported. Add one Unreleased bullet and package example.

- [ ] **Step 3: Run docs checks**

~~~sh
git diff --check
rg -n "AddTableCell|AddTableCellInput|text:|cell object|cell objects|对象" \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
~~~

---

### Task 5: Full gates, real-deck QA, review, commit, and push

**Files:**
- Review every Task 1-4 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: all implementation/tests/docs, actual packed package, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces: reviewed `feat: create tables with text cell objects` commit synchronized to `origin/main`.

- [ ] **Step 1: Run full functional and performance gates**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-object-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
~~~

Parse and record suite/test passed, failed, and skipped totals.

- [ ] **Step 2: Build and smoke the actual npm tarball**

~~~sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js \
  --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_object_package_dir=$(mktemp -d \
  /tmp/pptx-table-cell-object-package.XXXXXX)
npm pack --ignore-scripts \
  --pack-destination "$table_cell_object_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_cell_object_package_dir/jiayunxie-pptx-0.1.0.tgz"
~~~

Expected JSON includes `tableCellObjectCreation: true`, existing table creation/editing flags, declarations, browser success, and CLI `0.1.0`.

- [ ] **Step 3: Verify repository CLI health**

~~~sh
node packages/pptx/dist/cli.js --json doctor
~~~

Require successful offline JSON.

- [ ] **Step 4: Generate equivalent string/object, edited, reopened, and baseline decks**

Create `/tmp/pptx-table-cell-object-qa.mjs` with `apply_patch` and this complete program:

~~~js
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
const output = '/tmp/pptx-table-cell-object';
await mkdir(output, { recursive: true });

const textRows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];
const objectRows = [
  [{ text: 'Region' }, 'Revenue', { text: 'Growth' }],
  ['East', { text: '$1.2M' }, 'Vertical'],
  [{ text: 'West' }, '$980K', { text: '' }],
];

function build(rows) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const slide = document.addSlide();
  const table = slide.addTable(rows, {
    name: 'Cell object creation QA',
    x: inches(1),
    y: inches(1),
    columnWidths: [inches(2), inches(4), inches(2)],
    rowHeights: [inches(0.5), inches(1.25), inches(0.65)],
  });
  table.setCellFill(0, 0, {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
  });
  table.setCellTextDirection(1, 2, 'vert270');
  table.setCellVerticalAlignment(2, 1, 'bottom');
  return { document, table };
}

const strings = build(textRows);
const objects = build(objectRows);
await writeFile(output + '/string-source.pptx', await strings.document.write());
await writeFile(output + '/object-source.pptx', await objects.document.write());
assert.deepEqual(
  objects.table.rows.map(({ cells }) => cells.map(({ text }) => text)),
  textRows,
);

objects.table.setCellText(2, 2, 'Edited empty');
const editedBytes = await objects.document.write();
await writeFile(output + '/object-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[2].cells[2].text, 'Edited empty');
assert.equal(reopenedTable.rows[1].cells[2].textDirection, 'vert270');
assert.equal(reopenedTable.rows[2].cells[1].verticalAlignment, 'bottom');
await writeFile(output + '/object-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  textRows.map((row) => row.map((text) => ({ text, options: {} }))),
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
  sourceTexts: objects.table.rows.map(
    ({ cells }) => cells.map(({ text }) => text),
  ),
  editedEmpty: reopenedTable.rows[2].cells[2].text,
  verticalCell: reopenedTable.rows[1].cells[2].textDirection,
  bottomCell: reopenedTable.rows[2].cells[1].verticalAlignment,
}));
~~~

Run `node /tmp/pptx-table-cell-object-qa.mjs`. Require the original 3×3 text matrix, `editedEmpty: "Edited empty"`, `verticalCell: "vert270"`, and `bottomCell: "bottom"`.

- [ ] **Step 5: Validate packages and exact diff isolation**

Run PowerPoint 2010 validation on all five decks:

~~~sh
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-cell-object/string-source.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-cell-object/object-source.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-cell-object/object-edited.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-cell-object/object-reopened.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-cell-object/pptxgenjs-baseline.pptx \
  --profile powerpoint-2010
~~~

Run diffs:

~~~sh
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-object/string-source.pptx \
  /tmp/pptx-table-cell-object/object-source.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-object/object-source.pptx \
  /tmp/pptx-table-cell-object/object-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-object/object-edited.pptx \
  /tmp/pptx-table-cell-object/object-reopened.pptx
~~~

Require zero errors/warnings; string→object and edited→reopened must have zero changed decompressed parts; object-source→edited may change only `/ppt/slides/slide1.xml`.

- [ ] **Step 6: Render and inspect native/baseline decks**

~~~sh
mkdir -p /tmp/pptx-table-cell-object/rendered-native \
  /tmp/pptx-table-cell-object/rendered-baseline
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-cell-object/rendered-native \
  /tmp/pptx-table-cell-object/object-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-cell-object/rendered-baseline \
  /tmp/pptx-table-cell-object/pptxgenjs-baseline.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-cell-object/rendered-native/object-edited.pdf \
  /tmp/pptx-table-cell-object/rendered-native/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-cell-object/rendered-baseline/pptxgenjs-baseline.pdf \
  /tmp/pptx-table-cell-object/rendered-baseline/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-object/object-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-object/pptxgenjs-baseline.pptx
~~~

Inspect every PNG at original detail. Require all object/string text, edited former-empty cell, `0.5:1.25:0.65` rows, `2:4:2` columns, vertical text, bottom alignment, and no repair, clipping, wrap, unintended overlap, blur, missing cell, or off-slide content.

- [ ] **Step 7: Review, stage, commit, push, and prove synchronization**

~~~sh
git diff --check
git status --short
git diff --stat
git add \
  packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "feat: create tables with text cell objects"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
git status --short --branch
~~~

Expected: `0 0` and only `?? .pnpm-store/`.

---

After synchronization, design the first independently reviewable cell `options` creation capability without asking for routine decisions.
