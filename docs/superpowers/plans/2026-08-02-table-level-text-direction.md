# Table-Level Text Direction Implementation Plan

> **For agentic workers:** Execute this plan task-by-task inline in the current session. The standing user direction requires autonomous inline execution; review, commit, push, fetch, and verify each task before starting its successor. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a live `TableModel.textDirection` consensus getter and atomic all-cell setter over direct DrawingML cell direction attributes, share exact physical-cell resolution with table-level vertical alignment, and prove the final API in source, PptxGenJS import, actual-package Node/browser/CLI, and real Chrome.

**Architecture:** A focused `table-physical-cells.internal.ts` module owns exact direct `graphicFrame -> graphic -> graphicData -> tbl -> tr -> tc` resolution. The text-direction and vertical-alignment modules keep separate value readers/replacers but consume the same physical-cell set. `TableModel` exposes the new live property; no table metadata or new OOXML part is introduced.

**Tech Stack:** TypeScript 5.8, Vitest, lossless OOXML model editing, PptxGenJS 4.0.1, tsup, actual npm tarball smoke, installed `pptx-inspect`, Playwright, and installed Google Chrome.

## Global Constraints

- Public API is exactly `TableModel.textDirection: TableCellTextDirection | undefined` with both getter and setter.
- Getter returns a value only for one or more physical cells with one identical safe explicit direct token; absent, mixed, empty, and unsafe state return `undefined`.
- Getter never synthesizes absent state as `horz`; only explicit direct `vert="horz"` returns `horz`.
- Setter accepts only `horz`, `vert`, `vert270`, `wordArtVert`, or `undefined`, and validates before package access.
- Setter writes explicit direct `horz`, clears only for `undefined`, edits every physical `<a:tc>` including merge continuations, serializes the slide at most once, and is an exact no-op when no token changes.
- Missing or ambiguous table paths, empty physical-cell sets, missing or repeated direct `tcPr`, or repeated direct `vert` reject edits with `ModelParseError` and no package mutation.
- One unsupported direct token may be repaired or cleared, matching current single-cell replacement semantics.
- The shared physical-cell resolver contains no property parsing, mutation, model error, or visible-merge inference.
- Do not add a `mixed` token, array getter, table-default metadata, extension element, alias type, additional DrawingML token, or PptxGenJS compatibility facade.
- Do not change table creation normalization/output, `TableCell.textDirection`, `setCellTextDirection()`, or unrelated table properties.
- Existing `TableModel.verticalAlignment` behavior and tests must remain unchanged after moving structural resolution.
- Every task ends with review, scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`, generated workspace tarballs, or `/tmp` evidence.

---

## File map

- Create `packages/model/src/table-physical-cells.internal.ts`: exact direct table-path and physical-cell resolution only.
- Create `packages/model/src/table-cell-text-direction.internal.test.ts`: direct consensus read, bulk patch, explicit horizontal, no-op, preservation, malformed state, and shared-path coverage.
- Modify `packages/model/src/table-cell-text-direction.internal.ts`: table-level read/replace helpers that reuse existing cell logic.
- Modify `packages/model/src/table-cell-vertical-alignment.internal.ts`: replace its private physical-cell parser with the shared resolver without changing public behavior.
- Modify `packages/model/src/shapes.ts`: public live `TableModel.textDirection` getter/setter.
- Modify `packages/model/src/model.test.ts`: lifecycle, isolation, rollback, malformed package, and type behavior.
- Modify `packages/sdk/src/index.test.ts`: aggregate SDK/root creation/edit/write/reopen proof and public type closure.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: real PptxGenJS uniform, horizontal-collapse, mixed import, and native normalization proof.
- Modify `scripts/build-npm-package-types.mjs`: require getter/setter inside the packed `TableModel` declaration block.
- Modify `scripts/smoke-npm-package.mjs`: declarations, Node, TypeScript, browser condition, CLI part inspection, and top-level stable state.
- Modify `scripts/playwright-browser-smoke.js`: real-Chrome uniform/mixed/overwrite/horizontal/clear/reopen and zero-error state.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: public contract, direct-state boundary, evidence, progress, and next item.

---

### Task 1: Shared physical-cell resolver, core property, and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/table-physical-cells.internal.ts`
- Create: `packages/model/src/table-cell-text-direction.internal.test.ts`
- Modify: `packages/model/src/table-cell-text-direction.internal.ts`
- Modify: `packages/model/src/table-cell-vertical-alignment.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: `readTableCellTextDirection()`, `replaceTableCellTextDirection()`, `normalizeTableCellTextDirection()`, live `TableModel.resolve()`, and existing table creation/cell editing APIs.
- Produces: `readDirectTablePhysicalCells(frame)`, `readTableTextDirection(xml, frame)`, `replaceTableTextDirection(xml, frame, value, partUri)`, and public `TableModel.textDirection` getter/setter.

- [ ] **Step 1: Add direct helper tests for consensus and structure**

Create `packages/model/src/table-cell-text-direction.internal.test.ts` with exact frame and arbitrary-source fixtures:

```ts
import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableTextDirection,
  replaceTableTextDirection,
} from './table-cell-text-direction.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';

function parseSource(source: string): {
  xml: LosslessXmlDocument;
  frame: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(source);
  const frame = xml.roots[0];
  if (!frame) throw new Error('Fixture has no root frame');
  return { xml, frame };
}

function parseFrame(rows: string) {
  return parseSource(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
      '<a:graphic><a:graphicData><a:tbl>' + rows +
      '</a:tbl></a:graphicData></a:graphic>' +
      '</p:graphicFrame>',
  );
}
```

Require each supported explicit token to round-trip through the getter, including merge continuations and two rows:

```ts
for (const direction of ['horz', 'vert', 'vert270', 'wordArtVert'] as const) {
  const fixture = parseFrame(
    '<a:tr h="1">' +
      `<a:tc gridSpan="2"><a:tcPr vert="${direction}"/></a:tc>` +
      `<a:tc hMerge="1"><a:tcPr vert="${direction}"/></a:tc>` +
    '</a:tr>' +
    `<a:tr h="2"><a:tc vMerge="1"><a:tcPr vert="${direction}"/></a:tc></a:tr>`,
  );
  expect(readTableTextDirection(fixture.xml, fixture.frame)).toBe(direction);
  expect(fixture.xml.changed).toBe(false);
}
```

Use explicit fixtures for all absent, legal mixed, present/absent mixed, unsupported `eaVert`, repeated `vert`, repeated/missing `tcPr`, empty table, empty row, non-`graphicFrame` root, repeated direct `graphic`, repeated direct `graphicData`, repeated direct `tbl`, and opaque nested `tr/tc/tcPr` lookalikes. Every unsafe or non-consensus read returns `undefined` and leaves `xml.changed === false`.

- [ ] **Step 2: Add direct helper tests for bulk writes and atomic failure**

Require all-cell replacement across direct physical cells while preserving unrelated XML:

```ts
const target = parseFrame(
  '<a:tr h="1">' +
    '<a:tc gridSpan="2"><a:tcPr keep="A" vert="vert"><a:solidFill/></a:tcPr></a:tc>' +
    '<a:tc hMerge="1"><a:tcPr keep="B"/></a:tc>' +
  '</a:tr>' +
  '<a:tr h="2"><a:tc vMerge="1"><a:tcPr keep="C" vert="eaVert"/></a:tc></a:tr>',
);
expect(replaceTableTextDirection(
  target.xml,
  target.frame,
  'wordArtVert',
  PART_URI,
)).toBe(true);
expect(target.xml.serialize().match(/ vert="wordArtVert"/g)).toHaveLength(3);
expect(target.xml.serialize()).toContain('gridSpan="2"');
expect(target.xml.serialize()).toContain('hMerge="1"');
expect(target.xml.serialize()).toContain('vMerge="1"');
expect(target.xml.serialize()).toContain('keep="A"');
expect(target.xml.serialize()).toContain('<a:solidFill/>');
```

Reparse the serialized XML before exact no-op assertions so element offsets match the new source. Require same-value `wordArtVert` to return `false` and preserve exact bytes. Require `horz` to replace every token with explicit `vert="horz"`; require `undefined` to remove every direct `vert` while preserving all `keep` attributes and child nodes; require a second clear to return `false` with exact bytes.

Require empty/ambiguous physical paths, missing/repeated direct `tcPr`, and repeated direct `vert` to throw `ModelParseError`. Put the repeated attribute on the last of three cells, call the bulk replacer on one in-memory document, and require that no package-level assertion is made from the partially changed document; the public model test in Step 3 proves package atomicity. Require a single unsupported token to be replaceable and clearable.

- [ ] **Step 3: Add failing public model lifecycle and type tests**

Beside the existing table text-direction lifecycle test in `packages/model/src/model.test.ts`, add a dedicated two-row/two-column table using non-default borders, fill, margins, fit, horizontal alignment, vertical alignment, widths, and heights. Require:

```ts
expect(table.textDirection).toBe('vert270');

const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
const noOpJournal = [...pkg.mutations];
table.textDirection = 'vert270';
expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
expect(pkg.mutations).toEqual(noOpJournal);

table.setCellTextDirection(0, 1, 'vert');
expect(table.textDirection).toBeUndefined();
table.textDirection = 'wordArtVert';
expect(table.textDirection).toBe('wordArtVert');
expect(table.rows.flatMap(({ cells }) => cells)
  .map(({ textDirection }) => textDirection))
  .toEqual(['wordArtVert', 'wordArtVert', 'wordArtVert', 'wordArtVert']);

table.textDirection = 'horz';
expect(table.textDirection).toBe('horz');
expect(table.rows.flatMap(({ cells }) => cells)
  .every(({ textDirection }) => textDirection === 'horz')).toBe(true);

table.textDirection = undefined;
expect(table.textDirection).toBeUndefined();
expect(table.rows.flatMap(({ cells }) => cells)
  .every(({ textDirection }) => textDirection === undefined)).toBe(true);
```

Continue with `vert` assignment, duplicate isolation, source-only edit, slide move, write/reopen, stable object identity, and exact preservation of text, borders, fill, margins, fit, horizontal/vertical alignment, grid/row dimensions, transform, relationships, and an untouched package part hash.

Add a transaction that assigns `wordArtVert` and then throws; require bytes, mutation journal, and prior direction to roll back. Pass each runtime value below through the setter and require `TypeError` plus unchanged bytes/journal:

```ts
[
  null,
  false,
  true,
  0,
  '',
  'horizontal',
  'Horz',
  ' vert',
  'vert90',
  'eaVert',
  [],
  {},
  Symbol('vert'),
]
```

Inject repeated direct `vert` into the final physical cell, require the getter to return `undefined`, then require assignment to throw `ModelParseError` with slide bytes and journal unchanged. Inject an empty direct table and require assignment to throw without mutation. These checks prove that earlier in-memory patches never become partial package edits.

In a compile-only branch, require the public type closure:

```ts
if (false) {
  const direction: TableCellTextDirection | undefined = table.textDirection;
  table.textDirection = 'horz';
  table.textDirection = 'vert';
  table.textDirection = 'vert270';
  table.textDirection = 'wordArtVert';
  table.textDirection = undefined;
  // @ts-expect-error unsupported table text direction
  table.textDirection = 'eaVert';
  void direction;
}
```

- [ ] **Step 4: Add failing SDK/root and real-PptxGenJS tests**

In `packages/sdk/src/index.test.ts`, add a root lifecycle test using `PptxDocument.create()` and `TableModel`. Require initial uniform read, a cell-created mixed state, bulk stacked normalization, explicit horizontal state, clear, final vertical state, write/reopen, duplicate isolation, invalid input failure isolation, and unchanged validation diagnostics.

In `packages/pptxgenjs-adapter/src/index.test.ts`, add a dedicated real PptxGenJS 4.0.1 test with these tables:

```ts
slide.addTable([['A', 'B']], {
  x: 0.5, y: 0.5, w: 4, h: 1,
  textDirection: 'vert270',
});
slide.addTable([['A', 'B']], {
  x: 0.5, y: 2, w: 4, h: 1,
  textDirection: 'horz',
});
slide.addTable([[
  { text: 'Inherited', options: {} },
  { text: 'Override', options: { textDirection: 'vert' } },
]], {
  x: 0.5, y: 3.5, w: 4, h: 1,
  textDirection: 'wordArtVert',
});
```

Require the first property to be `vert270`, the horizontal table property and both horizontal cell snapshots to be `undefined`, and the overridden table property to be `undefined` with cell matrix `['wordArtVert', 'vert']`. Set the mixed imported table to `horz`, require explicit `['horz', 'horz']`, then set it to `wordArtVert`, write/reopen, and require both the property and physical cell tokens to remain `wordArtVert`.

- [ ] **Step 5: Run RED and require missing-helper/property failures**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
```

Expected: FAIL because `readTableTextDirection`, `replaceTableTextDirection`, and `TableModel.textDirection` do not exist. Existing vertical-alignment tests remain runnable and must not be weakened.

- [ ] **Step 6: Implement the shared exact direct physical-cell resolver**

Create `packages/model/src/table-physical-cells.internal.ts`:

```ts
import type { XmlElement } from '@pptx/lossless-xml';

export function readDirectTablePhysicalCells(
  frame: XmlElement,
): readonly XmlElement[] | undefined {
  if (frame.localName !== 'graphicFrame') return undefined;
  const graphic = exactDirectChild(frame, 'graphic');
  const graphicData = graphic ? exactDirectChild(graphic, 'graphicData') : undefined;
  const table = graphicData ? exactDirectChild(graphicData, 'tbl') : undefined;
  if (!table) return undefined;
  const rows = directChildren(table, 'tr');
  if (rows.length === 0) return undefined;
  const matrix = rows.map((row) => directChildren(row, 'tc'));
  return matrix.some((cells) => cells.length === 0) ? undefined : matrix.flat();
}

function exactDirectChild(
  element: XmlElement,
  localName: string,
): XmlElement | undefined {
  const matches = directChildren(element, localName);
  return matches.length === 1 ? matches[0] : undefined;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
```

In `table-cell-vertical-alignment.internal.ts`, import `readDirectTablePhysicalCells`, replace both `physicalTableCells(frame)` calls, and remove only the now-duplicated private `physicalTableCells()` and `exactDirectChild()` functions. Retain the module's local `directChildren()` because cell-level `tcPr` parsing still uses it.

- [ ] **Step 7: Implement text-direction table helpers**

In `packages/model/src/table-cell-text-direction.internal.ts`, import the shared resolver and add:

```ts
export function readTableTextDirection(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): TableCellTextDirection | undefined {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) return undefined;
  const first = readTableCellTextDirection(xml, cells[0]!);
  if (first === undefined) return undefined;
  return cells.every(
    (cell) => readTableCellTextDirection(xml, cell) === first,
  ) ? first : undefined;
}

export function replaceTableTextDirection(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  value: TableCellTextDirection | undefined,
  partUri: string,
): boolean {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) {
    throw new ModelParseError(
      'Table must contain one complete set of direct physical cells',
      partUri,
    );
  }
  let changed = false;
  for (const cell of cells) {
    changed = replaceTableCellTextDirection(xml, cell, value, partUri) || changed;
  }
  return changed;
}
```

Do not change `renderTableCellTextDirectionAttribute()`: creation must continue collapsing `horz` to absence while editing writes explicit `horz` through the existing replacer.

- [ ] **Step 8: Implement the live `TableModel` property**

In `packages/model/src/shapes.ts`, import the two table-level helpers with the existing cell helpers and add after `rows`:

```ts
get textDirection(): TableCellTextDirection | undefined {
  const { xml, element } = this.resolve();
  return readTableTextDirection(xml, element);
}

set textDirection(value: TableCellTextDirection | undefined) {
  const direction = value === undefined
    ? undefined
    : normalizeTableCellTextDirection(value, 'Table text direction');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableTextDirection(
      xml,
      element,
      direction,
      this.slide.partUri,
    )) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

Do not modify table creation normalization, rendering, row snapshots, or the existing single-cell setter.

- [ ] **Step 9: Run focused GREEN, type, bundle, full regression, and performance gates**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
node_modules/.bin/vitest run --reporter=dot --maxWorkers=4
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Expected: every command passes. Record focused/full file and test totals plus performance elapsed time. Inspect `packages/pptx/dist/types/model/shapes.d.ts` and require both text-direction accessors inside `TableModel`. Confirm the browser bundle has no static `node:` import.

- [ ] **Step 10: Review, commit, push, and verify**

Review shared-helper scope, exact direct-path ownership, explicit-horizontal behavior, no-op behavior, late-cell package atomicity, error types/messages, unrelated-state preservation, vertical-alignment regression, duplicate/rollback/reopen, public types, and real PptxGenJS final output. Then:

```sh
git add \
  packages/model/src/table-physical-cells.internal.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.ts \
  packages/model/src/table-cell-vertical-alignment.internal.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit table-level text direction"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0` before Task 2.

---

### Task 2: Actual-package Node, declarations, browser, CLI, and Chrome proof

**Files:**
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: Task 1 `TableModel.textDirection`, packed `model/shapes.d.ts`, actual root/browser exports, and installed CLI.
- Produces: stable `tableTextDirection: true`, detailed `tableTextDirectionState`, and `tableTextDirectionInspect: true` evidence.

- [ ] **Step 1: Lock the packed `TableModel` declaration block**

In `scripts/build-npm-package-types.mjs`, after declaration copying, read `dist/types/model/shapes.d.ts`, slice from `export declare class TableModel` to `export declare class ChartModel`, and require:

```ts
get textDirection(): TableCellTextDirection | undefined;
set textDirection(value: TableCellTextDirection | undefined);
```

Throw `Packed TableModel declaration is missing table-level text direction` if the class block or either accessor is absent.

Mirror the same class-block check in `scripts/smoke-npm-package.mjs` using the installed `shapeDeclarationSource`. This check must not pass from `TableCell.textDirection`, `TextShapeModel.textDirection`, or another declaration outside `TableModel`.

- [ ] **Step 2: Add installed Node runtime evidence**

In the generated Node consumer, create a dedicated two-by-two table with `textDirection: 'vert270'`. Capture this exact stable state:

```ts
const tableTextDirectionState = {
  uniform: 'vert270',
  readIsolation: true,
  noOp: true,
  mixed: null,
  overwritten: 'wordArtVert',
  overwrittenCells: ['wordArtVert', 'wordArtVert', 'wordArtVert', 'wordArtVert'],
  horizontal: 'horz',
  horizontalCells: ['horz', 'horz', 'horz', 'horz'],
  cleared: null,
  clearedCells: [null, null, null, null],
  reopened: 'vert',
  reopenedCells: ['vert', 'vert', 'vert', 'vert'],
  invalidError: {
    name: 'TypeError',
    message: 'Table text direction must be horz, vert, vert270, or wordArtVert',
  },
  failureIsolation: true,
  validationErrors: 0,
};
```

Use the dedicated output path `table-text-direction-smoke.pptx`. Read isolation compares slide bytes and mutation journal before and after the getter. No-op compares the same state before and after assigning `vert270`. Mixed state comes from one cell set to `vert`; bulk overwrite sets `wordArtVert`; explicit horizontal sets `horz`; clear uses `undefined`; final `vert` is written and reopened. Invalid assignment occurs before final write, leaves bytes/journal unchanged, and does not prevent the later valid write.

Add `tableTextDirection` and `tableTextDirectionState` to `apiChecks`, the explicit failure guard, and top-level output JSON.

- [ ] **Step 3: Add installed TypeScript and browser-condition evidence**

In the generated TypeScript consumer, use the existing typed table:

```ts
const typedTableTextDirection: TableCellTextDirection | undefined =
  typedTable.textDirection;
typedTable.textDirection = 'horz';
typedTable.textDirection = 'vert';
typedTable.textDirection = 'vert270';
typedTable.textDirection = 'wordArtVert';
typedTable.textDirection = undefined;
// @ts-expect-error unsupported table-level text direction
typedTable.textDirection = 'eaVert';
```

Add `typedTableTextDirection` to the terminal `void` list so unused-variable settings cannot hide a missing type surface.

In the generated `browser-smoke.mjs`, repeat uniform read, one-cell mixed projection, bulk stacked, explicit horizontal, clear, final vertical `writeBlob()`/reopen, invalid failure isolation, and zero diagnostics. Throw `Browser table-level text direction failed` unless every scalar and four-cell matrix matches the Node state.

- [ ] **Step 4: Add installed CLI package inspection**

After installed CLI setup, inspect `table-text-direction-smoke.pptx`:

```sh
pptx-inspect --json package validate table-text-direction-smoke.pptx --profile powerpoint-2010
pptx-inspect --json slides list table-text-direction-smoke.pptx
pptx-inspect --json part read table-text-direction-smoke.pptx /ppt/slides/slide1.xml
```

Require zero validation errors and warnings, one table shape, exactly four physical cell property elements with direct `vert="vert"`, no direct `vert="horz"`, `vert="vert270"`, `vert="wordArtVert"`, or unsupported token, and no `a:bodyPr@vert` false positive. Add `tableTextDirectionInspect: true` to the top-level smoke summary.

- [ ] **Step 5: Extend the real-Chrome callback**

In `scripts/playwright-browser-smoke.js`, add the same stable state fields, using `writeBlob()` for reopen and `null` for absent/mixed JSON values. Require exact four-cell matrices, explicit `horz`, invalid error/failure isolation, getter/no-op mutation isolation, and zero document validation errors. Add both fields to actual and expected objects:

```json
{
  "tableTextDirection": true,
  "tableTextDirectionState": {
    "uniform": "vert270",
    "readIsolation": true,
    "noOp": true,
    "mixed": null,
    "overwritten": "wordArtVert",
    "overwrittenCells": ["wordArtVert", "wordArtVert", "wordArtVert", "wordArtVert"],
    "horizontal": "horz",
    "horizontalCells": ["horz", "horz", "horz", "horz"],
    "cleared": null,
    "clearedCells": [null, null, null, null],
    "reopened": "vert",
    "reopenedCells": ["vert", "vert", "vert", "vert"],
    "invalidError": {
      "name": "TypeError",
      "message": "Table text direction must be horz, vert, vert270, or wordArtVert"
    },
    "failureIsolation": true,
    "validationErrors": 0
  }
}
```

Retain global `errorCounts: { console: 0, page: 0, network: 0 }` and every pre-existing expected smoke field.

- [ ] **Step 6: Run syntax, build, actual-tarball, and installed smoke gates**

Run:

```sh
node --check scripts/build-npm-package-types.mjs
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
git diff --check
node scripts/build-npm-package-types.mjs
(cd packages/pptx && npm pack --ignore-scripts)
node scripts/smoke-npm-package.mjs packages/pptx/jiayunxie-pptx-0.1.0.tgz
```

Require `tableTextDirection: true`, `tableTextDirectionInspect: true`, exact `tableTextDirectionState`, and all prior smoke fields unchanged. Record actual tarball file count and SHA-256.

- [ ] **Step 7: Run installed Google Chrome against the extracted tarball**

Create a fresh evidence directory with `mktemp -d /tmp/pptx-table-text-direction-artifacts.XXXXXX`. Extract the actual tarball under `site/packages/pptx`, copy the browser callback, serve the package over loopback with a valid favicon response, and run installed Google Chrome through the bundled Playwright runtime.

Require exact `tableTextDirectionState`, `tableTextDirection: true`, and validation/console/page/network counts all zero. Retain compact and full result JSON under the evidence directory. Move the workspace tarball into that directory only after hashing so no generated tarball remains in the repository worktree.

- [ ] **Step 8: Review, commit, push, and verify**

Review class-block declaration isolation, actual packed imports, stable JSON null handling, explicit-horizontal state, invalid failure isolation, CLI direct-cell specificity, previous-field preservation, loopback responses, and zero Chrome errors. Then:

```sh
git add \
  scripts/build-npm-package-types.mjs \
  scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed table-level text direction"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0` before Task 3.

---

### Task 3: Public documentation and progress closeout

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: final public property, PptxGenJS horizontal-collapse/import behavior, full/performance/package/Chrome evidence, and retained artifact path.
- Produces: public examples, exact direct-consensus semantics, supported compatibility state, completed progress entry, and next smallest advanced-table item.

- [ ] **Step 1: Document usage and direct-state semantics**

In both public READMEs and `docs/api/README.md`, add a concise example with uniform `vert270` read, mixed `undefined`, bulk `wordArtVert`, explicit `horz`, and clear. State that DrawingML stores only physical per-cell direct `tcPr@vert`; the property is a live consensus projection and does not remember a creation default.

Document the four legal values, absent-versus-explicit-horizontal distinction, merge-continuation participation, exact no-ops, one-transaction all-cell writes, unrelated-property preservation, and unsafe edit rejection. Direct callers who need mixed detail should inspect `rows[].cells[].textDirection`.

- [ ] **Step 2: Update compatibility and deliberate divergence**

In `docs/compatibility/pptxgenjs-baseline.md`, change stale “table-level direction getter/editor pending” statements to supported direct consensus getter/bulk editor plus the existing cell editor. Record that PptxGenJS 4.0.1 exposes only creation options and collapses `horz` to absence; native post-creation explicit horizontal and bulk editing are lossless extensions over the same final cell attributes.

Do not claim effective style/default resolution, a serializable table-level field, inherited horizontal state, visible-cell merge semantics, extra direction tokens, or support for table-level fit, horizontal alignment, margins, borders, fill, auto-page, repeated headers, or `tableToSlides` from this item.

- [ ] **Step 3: Record exact verification and select the next item**

In `docs/implementation-progress.md`, record focused/full test totals, performance, tarball file count and SHA-256, Node/browser/CLI/Chrome state, zero browser errors, retained `/tmp` artifact path, and commit boundaries. Keep overall parity at approximately 97% unless the audited denominator is recalculated from the compatibility inventory.

Choose table-level direct horizontal-alignment consensus/bulk editing as the next smallest advanced-table item because strict table/cell creation and single-cell editing already exist, while its paragraph-level ownership rules need a dedicated design.

In `CHANGELOG.md`, add one precise entry for live direct table text-direction consensus/bulk editing, including explicit `horz`, clear, atomicity, and PptxGenJS final-state import behavior.

- [ ] **Step 4: Run final verification, review, commit, push, and verify**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/vitest run --reporter=dot --maxWorkers=4
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Scan for stale “table-level direction getter/editor pending” claims, verify all examples against declarations, and stage only the six documentation files:

```sh
git add \
  README.md \
  packages/pptx/README.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  docs/implementation-progress.md \
  CHANGELOG.md
git diff --cached --check
git commit -m "docs: document table-level text direction"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, a clean tracked worktree, and only the known untracked `.pnpm-store/` cache before beginning the next capability.
