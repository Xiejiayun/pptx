# Table-Level Vertical Alignment Implementation Plan

> **For agentic workers:** Execute this plan task-by-task inline in the current session. The standing user direction prohibits subagents; review, commit, push, fetch, and verify each task before starting its successor.

**Goal:** Add a live `TableModel.verticalAlignment` consensus getter and atomic all-cell setter over direct DrawingML cell anchors, then prove it in source, PptxGenJS import, actual-package Node/browser/CLI, and real Chrome.

**Architecture:** The existing table-cell vertical-alignment module owns table-path parsing and reuses its single-cell reader/replacer for one safe direct physical-cell set. `TableModel` exposes a property that delegates to those helpers; no table metadata or new OOXML is introduced. Package/browser smokes exercise final packed declarations and real write/reopen behavior.

**Tech Stack:** TypeScript 5.8, Vitest, lossless OOXML model editing, PptxGenJS 4.0.1, tsup, actual npm tarball smoke, installed `pptx-inspect`, Playwright, and installed Google Chrome.

## Global Constraints

- Public API is exactly `TableModel.verticalAlignment: TextBoxVerticalAlignment | undefined` with both getter and setter.
- Getter returns a value only for one or more physical cells with one identical safe direct value; absent, mixed, empty, and unsafe state return `undefined`.
- Setter accepts only `top`, `middle`, `bottom`, or `undefined`, and validates before package access.
- Setter edits every physical direct `<a:tc>` including merge continuations, serializes the slide at most once, and is an exact no-op when no anchor changes.
- Missing/ambiguous table paths, empty physical-cell sets, missing/repeated direct `tcPr`, or repeated direct anchors reject edits with `ModelParseError` and no package mutation.
- A single unsupported direct anchor token may be repaired or cleared, matching current cell replacement semantics.
- Do not add a `mixed` token, array getter, table-default metadata, extension element, alias type, or PptxGenJS compatibility facade.
- Do not change `TableCell.verticalAlignment`, `setCellVerticalAlignment()`, table creation precedence, or unrelated table properties.
- Every task ends with review, scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`, generated workspace tarballs, or `/tmp` evidence.

---

## File map

- Create `packages/model/src/table-cell-vertical-alignment.internal.test.ts`: direct path, consensus read, bulk patch, no-op, preservation, and malformed-state coverage.
- Modify `packages/model/src/table-cell-vertical-alignment.internal.ts`: exact table-path parsing plus table-level read/replace helpers.
- Modify `packages/model/src/shapes.ts`: public live getter/setter on `TableModel`.
- Modify `packages/model/src/model.test.ts`: lifecycle, isolation, rollback, malformed package, and type behavior.
- Modify `packages/sdk/src/index.test.ts`: aggregate SDK/root creation/edit/write/reopen proof and public type closure.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: real PptxGenJS uniform/mixed import and native normalization.
- Modify `scripts/build-npm-package-types.mjs`: require getter/setter inside the packed `TableModel` declaration block.
- Modify `scripts/smoke-npm-package.mjs`: declarations, Node, TypeScript, browser condition, CLI part inspection, and top-level stable state.
- Modify `scripts/playwright-browser-smoke.js`: real-Chrome uniform/mixed/overwrite/clear/reopen and zero-error state.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: public contract, PptxGenJS boundary, evidence, and next item.

---

### Task 1: Core consensus property and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/table-cell-vertical-alignment.internal.test.ts`
- Modify: `packages/model/src/table-cell-vertical-alignment.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: `readTableCellVerticalAlignment()`, `replaceTableCellVerticalAlignment()`, `normalizeTextBoxVerticalAlignment()`, live `TableModel.resolve()`, and existing table creation/cell editing APIs.
- Produces: `readTableVerticalAlignment(xml, frame)`, `replaceTableVerticalAlignment(xml, frame, value, partUri)`, and public `TableModel.verticalAlignment` getter/setter.

- [ ] **Step 1: Add direct helper tests**

Create `packages/model/src/table-cell-vertical-alignment.internal.test.ts` with a direct frame fixture:

```ts
import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableVerticalAlignment,
  replaceTableVerticalAlignment,
} from './table-cell-vertical-alignment.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';

function parseFrame(rows: string): {
  xml: LosslessXmlDocument;
  frame: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
      '<a:graphic><a:graphicData><a:tbl>' + rows +
      '</a:tbl></a:graphicData></a:graphic>' +
      '</p:graphicFrame>',
  );
  const frame = xml.roots[0];
  if (!frame) throw new Error('Fixture has no root frame');
  return { xml, frame };
}
```

Add exact cases that require:

```ts
const uniform = parseFrame(
  '<a:tr h="1"><a:tc><a:tcPr anchor="ctr"/></a:tc>' +
  '<a:tc hMerge="1"><a:tcPr anchor="ctr"/></a:tc></a:tr>',
);
expect(readTableVerticalAlignment(uniform.xml, uniform.frame)).toBe('middle');
```

Use explicit fixtures for uniform `t/ctr/b`, all absent, legal mixed, present/absent mixed, unsupported `mid`, repeated anchor, repeated/missing `tcPr`, empty table, repeated direct `graphic`/`graphicData`/`tbl`, and opaque nested lookalikes. Reads must not mark XML changed.

For writes, require `replaceTableVerticalAlignment()` to:

```ts
const target = parseFrame(
  '<a:tr h="1">' +
    '<a:tc gridSpan="2"><a:tcPr keep="A" anchor="t"><a:solidFill/></a:tcPr></a:tc>' +
    '<a:tc hMerge="1"><a:tcPr keep="B"/></a:tc>' +
  '</a:tr>' +
  '<a:tr h="2"><a:tc vMerge="1"><a:tcPr keep="C" anchor="b"/></a:tc></a:tr>',
);
expect(replaceTableVerticalAlignment(
  target.xml,
  target.frame,
  'middle',
  PART_URI,
)).toBe(true);
expect(target.xml.serialize().match(/ anchor="ctr"/g)).toHaveLength(3);
expect(target.xml.serialize()).toContain('keep="A"');
expect(target.xml.serialize()).toContain('<a:solidFill/>');
expect(replaceTableVerticalAlignment(
  target.xml,
  target.frame,
  'middle',
  PART_URI,
)).toBe(false);
expect(replaceTableVerticalAlignment(
  target.xml,
  target.frame,
  undefined,
  PART_URI,
)).toBe(true);
expect(target.xml.serialize()).not.toContain(' anchor=');
```

Require empty/ambiguous physical cell paths, missing/repeated `tcPr`, and repeated anchor to throw `ModelParseError`. Require a single invalid token to be replaceable and clearable.

- [ ] **Step 2: Add failing public model lifecycle tests**

Beside the existing `materializes table valign` test in `packages/model/src/model.test.ts`, add a dedicated test that creates a two-row/two-column table with `valign: 'middle'` and requires:

```ts
expect(table.verticalAlignment).toBe('middle');

const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
const noOpJournal = [...pkg.mutations];
table.verticalAlignment = 'middle';
expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
expect(pkg.mutations).toEqual(noOpJournal);

table.setCellVerticalAlignment(0, 1, 'top');
expect(table.verticalAlignment).toBeUndefined();
table.verticalAlignment = 'bottom';
expect(table.verticalAlignment).toBe('bottom');
expect(table.rows.flatMap(({ cells }) => cells)
  .map(({ verticalAlignment }) => verticalAlignment))
  .toEqual(['bottom', 'bottom', 'bottom', 'bottom']);

table.verticalAlignment = undefined;
expect(table.verticalAlignment).toBeUndefined();
expect(table.rows.flatMap(({ cells }) => cells)
  .every(({ verticalAlignment }) => verticalAlignment === undefined)).toBe(true);
```

Continue the same test with `top` assignment, duplicate isolation, source-only edit, move, write/reopen, stable object identity, and exact preservation of text, borders, fill, margins, text direction, text fit, horizontal alignment, grid/row dimensions, transform, relationships, and an untouched part hash.

Add a transaction that assigns `middle` and then throws; require bytes, mutation journal, and prior alignment to roll back. Pass each invalid runtime value below through the setter and require `TypeError` plus unchanged bytes/journal:

```ts
[null, false, true, 0, '', 'Top', ' top', 't', 'ctr', 'distributed', [], {}, Symbol('top')]
```

Inject a repeated anchor into a later physical cell, require the getter to return `undefined`, then require assignment to throw `ModelParseError` without committing the earlier cell change. Inject an empty direct table and require assignment to throw without mutation.

In a compile-only branch, require the public type closure:

```ts
if (false) {
  const alignment: TextBoxVerticalAlignment | undefined = table.verticalAlignment;
  table.verticalAlignment = 'top';
  table.verticalAlignment = 'middle';
  table.verticalAlignment = 'bottom';
  table.verticalAlignment = undefined;
  // @ts-expect-error unsupported table vertical alignment
  table.verticalAlignment = 'distributed';
  void alignment;
}
```

- [ ] **Step 3: Add failing SDK/root and real-PptxGenJS tests**

In `packages/sdk/src/index.test.ts`, add a root lifecycle test using `PptxDocument.create()` and `TableModel`. Require initial uniform read, cell-created mixed state, bulk normalization, clear, write/reopen, duplicate isolation, invalid input failure isolation, and validation diagnostics unchanged.

In `packages/pptxgenjs-adapter/src/index.test.ts`, extend the existing table-level `valign` conformance test with a second PptxGenJS slide whose table has only `valign: 'middle'`. Require:

```ts
expect(uniformImportedTable.verticalAlignment).toBe('middle');
expect(importedTable.verticalAlignment).toBeUndefined();
importedTable.verticalAlignment = 'bottom';
expect(importedTable.verticalAlignment).toBe('bottom');
expect(importedTable.rows[0]!.cells.every(
  ({ verticalAlignment }) => verticalAlignment === 'bottom',
)).toBe(true);
const reopenedImported = await PptxDocument.open(await imported.write());
expect((reopenedImported.slides[0]!.shapes[0] as TableModel).verticalAlignment)
  .toBe('bottom');
```

Also assert that the mixed PptxGenJS input retains its exact per-cell matrix before native normalization.

- [ ] **Step 4: Run RED and require missing-helper/property failures**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
```

Expected: FAIL because `readTableVerticalAlignment`, `replaceTableVerticalAlignment`, and `TableModel.verticalAlignment` do not exist. Do not weaken test expectations.

- [ ] **Step 5: Implement exact table-path parsing and bulk helpers**

In `packages/model/src/table-cell-vertical-alignment.internal.ts`, retain existing cell helpers and add:

```ts
export function readTableVerticalAlignment(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): TextBoxVerticalAlignment | undefined {
  const cells = physicalTableCells(frame);
  if (!cells) return undefined;
  const first = readTableCellVerticalAlignment(xml, cells[0]!);
  if (first === undefined) return undefined;
  return cells.every(
    (cell) => readTableCellVerticalAlignment(xml, cell) === first,
  ) ? first : undefined;
}

export function replaceTableVerticalAlignment(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  value: TextBoxVerticalAlignment | undefined,
  partUri: string,
): boolean {
  const cells = physicalTableCells(frame);
  if (!cells) {
    throw new ModelParseError(
      'Table must contain one complete set of direct physical cells',
      partUri,
    );
  }
  let changed = false;
  for (const cell of cells) {
    changed = replaceTableCellVerticalAlignment(
      xml,
      cell,
      value,
      partUri,
    ) || changed;
  }
  return changed;
}
```

Import `ModelParseError` and add private `physicalTableCells()`, `exactDirectChild()`, and `directChildren()` helpers. `physicalTableCells()` must require exactly one direct `graphic -> graphicData -> tbl` path, one or more direct rows, and one or more direct cells in every row:

```ts
function physicalTableCells(frame: XmlElement): readonly XmlElement[] | undefined {
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
```

- [ ] **Step 6: Implement the live `TableModel` property**

In `packages/model/src/shapes.ts`, import the two table-level helpers and add after `rows`:

```ts
get verticalAlignment(): TextBoxVerticalAlignment | undefined {
  const { xml, element } = this.resolve();
  return readTableVerticalAlignment(xml, element);
}

set verticalAlignment(value: TextBoxVerticalAlignment | undefined) {
  const alignment = value === undefined
    ? undefined
    : normalizeTextBoxVerticalAlignment(value, 'Table vertical alignment');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableVerticalAlignment(
      xml,
      element,
      alignment,
      this.slide.partUri,
    )) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

Do not modify table creation normalization or the existing cell getter/setter.

- [ ] **Step 7: Run focused GREEN, type, bundle, full regression, and performance gates**

Run:

```sh
node_modules/.bin/vitest run \
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

Expected: every command passes; record focused/full file and test totals plus performance elapsed time. Inspect `packages/pptx/dist/types/model/shapes.d.ts` and require both accessors inside `TableModel`. Confirm the browser bundle has no static `node:` import.

- [ ] **Step 8: Review, commit, push, and verify**

Review exact direct-path ownership, no-op behavior, late-cell atomicity, error types/messages, unrelated-state preservation, duplicate/rollback/reopen, public types, and real PptxGenJS final output. Then:

```sh
git add \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit table-level vertical alignment"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
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
- Consumes: Task 1 `TableModel.verticalAlignment`, packed `model/shapes.d.ts`, actual root/browser exports, and installed CLI.
- Produces: stable `tableVerticalAlignment: true`, detailed `tableVerticalAlignmentState`, and `tableVerticalAlignmentInspect: true` evidence.

- [ ] **Step 1: Lock the packed `TableModel` declaration block**

In `scripts/build-npm-package-types.mjs`, after declaration copying, read `dist/types/model/shapes.d.ts`, slice from `export declare class TableModel` to `export declare class ChartModel`, and require:

```ts
get verticalAlignment(): TextBoxVerticalAlignment | undefined;
set verticalAlignment(value: TextBoxVerticalAlignment | undefined);
```

Throw `Packed TableModel declaration is missing table-level vertical alignment` if the class block or either accessor is absent.

Mirror the same class-block check in `scripts/smoke-npm-package.mjs` using the installed `shapeDeclarationSource`. This must not pass merely because `ShapeModel` or `TableCell` has a property of the same name.

- [ ] **Step 2: Add installed Node runtime and TypeScript evidence**

In the generated Node consumer, create a two-by-two table with `valign: 'middle'`. Capture this exact stable state sequence:

```ts
const tableVerticalAlignmentState = {
  uniform: 'middle',
  readIsolation: true,
  noOp: true,
  mixed: null,
  overwritten: 'bottom',
  overwrittenCells: ['bottom', 'bottom', 'bottom', 'bottom'],
  cleared: null,
  clearedCells: [null, null, null, null],
  reopened: 'top',
  reopenedCells: ['top', 'top', 'top', 'top'],
  invalidError: {
    name: 'TypeError',
    message: 'Table vertical alignment must be top, middle, or bottom',
  },
  failureIsolation: true,
  validationErrors: 0,
};
```

The generated deck path is `table-vertical-alignment-smoke.pptx`. Read isolation compares bytes/journal before and after the getter. No-op compares bytes/journal before and after assigning `middle`. Mixed state comes from one cell set to `top`; bulk overwrite sets `bottom`; clear uses `undefined`; final `top` is written and reopened. Invalid assignment must occur before final write, leave bytes/journal unchanged, and permit the later valid write.

Add `tableVerticalAlignment` and `tableVerticalAlignmentState` to `apiChecks`, the explicit failure guard, and the top-level output JSON.

In the generated TypeScript consumer, use the existing `typedTable`:

```ts
const typedTableVerticalAlignment: TextBoxVerticalAlignment | undefined =
  typedTable.verticalAlignment;
typedTable.verticalAlignment = 'top';
typedTable.verticalAlignment = 'middle';
typedTable.verticalAlignment = 'bottom';
typedTable.verticalAlignment = undefined;
// @ts-expect-error unsupported table-level vertical alignment
typedTable.verticalAlignment = 'distributed';
```

Add every variable to the terminal `void` list so `noUnusedLocals`-equivalent drift cannot hide a missing type surface.

- [ ] **Step 3: Add browser-condition and installed CLI proof**

In the generated `browser-smoke.mjs`, repeat uniform read, one-cell mixed projection, bulk bottom, clear, final top writeBlob/reopen, invalid failure isolation, and zero diagnostics. Throw `Browser table-level vertical alignment failed` unless all values match.

After installed CLI setup, inspect `table-vertical-alignment-smoke.pptx`:

```sh
pptx-inspect --json package validate table-vertical-alignment-smoke.pptx --profile powerpoint-2010
pptx-inspect --json slides list table-vertical-alignment-smoke.pptx
pptx-inspect --json part read table-vertical-alignment-smoke.pptx /ppt/slides/slide1.xml
```

Require 0 validation errors/warnings, one table shape, exactly four direct `anchor="t"` attributes, no `anchor="ctr"` or `anchor="b"`, and no `a:bodyPr@anchor` false positives. Add `tableVerticalAlignmentInspect: true` to the top-level smoke summary.

- [ ] **Step 4: Extend the real-Chrome callback**

In `scripts/playwright-browser-smoke.js`, add the same stable state fields, using `writeBlob()` for reopen. Store absent/mixed values as `null` so JSON comparison preserves them. Require exact four-cell matrices, invalid error/failure isolation, mutation-read isolation, and zero document validation errors. Add both fields to actual and expected objects:

```json
{
  "tableVerticalAlignment": true,
  "tableVerticalAlignmentState": {
    "uniform": "middle",
    "readIsolation": true,
    "noOp": true,
    "mixed": null,
    "overwritten": "bottom",
    "overwrittenCells": ["bottom", "bottom", "bottom", "bottom"],
    "cleared": null,
    "clearedCells": [null, null, null, null],
    "reopened": "top",
    "reopenedCells": ["top", "top", "top", "top"],
    "invalidError": {
      "name": "TypeError",
      "message": "Table vertical alignment must be top, middle, or bottom"
    },
    "failureIsolation": true,
    "validationErrors": 0
  }
}
```

Retain global `errorCounts: { console: 0, page: 0, network: 0 }`.

- [ ] **Step 5: Run syntax, build, actual-tarball, and installed smoke gates**

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

Require `tableVerticalAlignment: true`, `tableVerticalAlignmentInspect: true`, all prior smoke fields unchanged, and record actual tarball file count and SHA-256.

- [ ] **Step 6: Run installed Google Chrome against the extracted tarball**

Extract the actual tarball into a fresh `/tmp/pptx-table-vertical-alignment-artifacts.XXXXXX/site/packages/pptx`, copy the browser callback, serve the package over loopback with a valid favicon response, and run installed Google Chrome through the bundled Playwright runtime.

Require exact `tableVerticalAlignmentState`, `tableVerticalAlignment: true`, and `errorCounts` all zero. Retain compact and full result JSON under the evidence directory. Move the workspace tarball into that directory after hashing.

- [ ] **Step 7: Review, commit, push, and verify**

Review class-block declaration isolation, actual packed imports, stable JSON null handling, invalid failure isolation, CLI anchor specificity, prior-field preservation, all HTTP responses, and zero Chrome errors. Then:

```sh
git add \
  scripts/build-npm-package-types.mjs \
  scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed table-level vertical alignment"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
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
- Consumes: final public property, PptxGenJS import behavior, full/performance/package/Chrome evidence, and retained artifact path.
- Produces: public examples, exact consensus semantics, supported compatibility state, completed progress entry, and next smallest advanced-table item.

- [ ] **Step 1: Document usage and the recoverable-state boundary**

Add a concise example showing uniform read, mixed `undefined`, bulk bottom assignment, and clear. State that DrawingML stores only per-cell direct anchors; the property is a consensus projection and does not remember a creation default.

Document physical-cell/merge-continuation participation, legal values, exact no-ops, all-cell atomicity, unrelated-property preservation, and unsafe edit rejection. Direct callers who need mixed detail should inspect `rows[].cells[].verticalAlignment`.

- [ ] **Step 2: Update compatibility and deliberate divergence**

Change the table/table-cell `valign` row from “table getter/editor pending” to supported table consensus getter/bulk editor plus existing cell editor. Record that PptxGenJS 4.0.1 exposes only creation options; native post-creation editing is a lossless extension over the same final cell anchors.

Do not claim table-level OOXML ownership, inherited defaults, merge semantics, horizontal alignment, margins, borders, fill, fit, auto-page, or `tableToSlides` support from this item.

- [ ] **Step 3: Record exact verification and next item**

In `docs/implementation-progress.md`, record focused/full test totals, performance, tarball count/SHA-256, Node/browser/CLI/Chrome state, zero browser errors, `/tmp` artifact path, and commit boundaries. Keep overall parity at approximately 97% unless the audited denominator is recalculated.

Choose table-level direct text-direction consensus/bulk editing as the next smallest advanced-table item because it uses the same physical-cell projection pattern and creation support already exists.

- [ ] **Step 4: Review, commit, push, and verify**

Run `git diff --check`, scan for stale “table valign getter/editor pending” claims, verify all examples against declarations, and stage only the six documentation files:

```sh
git add \
  README.md \
  packages/pptx/README.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  docs/implementation-progress.md \
  CHANGELOG.md
git diff --cached --check
git commit -m "docs: document table-level vertical alignment"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, a clean tracked worktree, and only the known untracked `.pnpm-store/` cache before selecting the next item.
