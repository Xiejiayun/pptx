# Table-Level Horizontal Alignment Implementation Plan

> **For agentic workers:** Execute this plan task-by-task inline in the current session. The standing user direction requires autonomous inline execution without subagents; review, commit, push, fetch, and verify each task before starting its successor. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a live `TableModel.horizontalAlignment` consensus getter and atomic all-cell setter over direct single-paragraph DrawingML cell alignment, then prove the final API in source, PptxGenJS import, actual-package Node/browser/CLI, and real Chrome.

**Architecture:** `table-cell-horizontal-alignment.internal.ts` keeps ownership of paragraph structure, token parsing, and lossless replacement while reusing `readDirectTablePhysicalCells()` for exact table traversal. `TableModel` exposes the live property; no table metadata, style calculation, or new OOXML part is introduced.

**Tech Stack:** TypeScript 5.8, Vitest, lossless OOXML model editing, PptxGenJS 4.0.1, tsup, actual npm tarball smoke, installed `pptx-inspect`, Playwright, and installed Google Chrome.

## Global Constraints

- Public API is exactly `TableModel.horizontalAlignment: TextAlignment | undefined` with both getter and setter.
- Getter returns a value only when one or more physical cells all expose one identical safe explicit direct single-paragraph token; absent, mixed, empty, multi-paragraph, and unsafe state return `undefined`.
- Getter never synthesizes absent state as `left`; only explicit direct `algn="l"` returns `left`.
- Setter accepts only `left`, `center`, `right`, `justify`, or `undefined`, and validates before package access.
- Setter writes explicit direct `left`, clears only for `undefined`, edits every physical `<a:tc>` including merge continuations, serializes the slide at most once, and is an exact no-op when no token changes.
- Missing `pPr` may be created for a legal value; clearing an absent alignment must not create `pPr`.
- Missing/ambiguous table paths, empty physical-cell sets, missing/multiple direct text bodies, zero/multiple direct paragraphs, repeated direct `pPr`, repeated direct `algn`, or invalid paragraph templates reject edits with `ModelParseError` and no package mutation.
- One unsupported direct `algn` may be repaired or cleared, matching current single-cell replacement semantics.
- Do not add a `mixed` token, array getter, table-default metadata, extension element, alias type, additional DrawingML token, rich/multi-paragraph bulk editing, or PptxGenJS compatibility facade.
- Do not change table creation normalization/output, `TableCell.horizontalAlignment`, `setCellHorizontalAlignment()`, or unrelated table properties.
- Existing `TableModel.verticalAlignment` and `TableModel.textDirection` behavior and tests must remain unchanged.
- Every task ends with review, scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`, generated workspace tarballs, or `/tmp` evidence.

---

## File map

- Create `packages/model/src/table-cell-horizontal-alignment.internal.test.ts`: direct consensus read, bulk patch, explicit-left, no-op, preservation, malformed-state, and shared-path coverage.
- Modify `packages/model/src/table-cell-horizontal-alignment.internal.ts`: table-level read/replace helpers using the shared physical-cell resolver.
- Modify `packages/model/src/shapes.ts`: public live `TableModel.horizontalAlignment` getter/setter.
- Modify `packages/model/src/model.test.ts`: lifecycle, isolation, rollback, malformed package, and type behavior.
- Modify `packages/sdk/src/index.test.ts`: aggregate SDK/root creation/edit/write/reopen proof and public type closure.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: real PptxGenJS uniform, omitted, mixed import, and native normalization proof.
- Modify `scripts/build-npm-package-types.mjs`: require getter/setter inside the packed `TableModel` declaration block.
- Modify `scripts/smoke-npm-package.mjs`: declarations, Node, TypeScript, browser condition, CLI part inspection, and stable top-level state.
- Modify `scripts/playwright-browser-smoke.js`: real-Chrome uniform/mixed/overwrite/explicit-left/clear/reopen and zero-error state.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: public contract, evidence, progress, and next item.

---

### Task 1: Core property, source tests, and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/table-cell-horizontal-alignment.internal.test.ts`
- Modify: `packages/model/src/table-cell-horizontal-alignment.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: `readDirectTablePhysicalCells()`, `readTableCellHorizontalAlignment()`, `replaceTableCellHorizontalAlignment()`, `normalizeTextAlignment()`, live `TableModel.resolve()`, and existing table creation/cell editing APIs.
- Produces: `readTableHorizontalAlignment(xml, frame)`, `replaceTableHorizontalAlignment(xml, frame, value, partUri)`, and public `TableModel.horizontalAlignment` getter/setter.

- [ ] **Step 1: Write failing internal consensus tests**

Create `packages/model/src/table-cell-horizontal-alignment.internal.test.ts` with a `parseSource()` fixture using `LosslessXmlDocument` and direct `graphicFrame/graphic/graphicData/tbl/tr/tc/txBody/p/pPr` XML. Import the not-yet-implemented helpers:

```ts
import {
  readTableHorizontalAlignment,
  replaceTableHorizontalAlignment,
} from './table-cell-horizontal-alignment.internal.js';
```

Require uniform `l/ctr/r/just` to read as `left/center/right/justify`; all-absent, present/absent mixed, mixed legal values, unsupported tokens, repeated `algn`, missing/repeated `pPr`, missing/multiple `txBody`, zero/multiple paragraphs, empty rows/table, ambiguous direct table paths, and descendant lookalikes to read `undefined`. Include physical merge-continuation cells in the consensus.

- [ ] **Step 2: Write failing internal bulk-edit and atomicity tests**

Use a two-by-two fixture containing text, runs, unrelated paragraph attributes/children, `tcPr` direction/anchor/margins/borders/fill, fit, grid, rows, and transform. Require:

```ts
expect(replaceTableHorizontalAlignment(xml, frame, 'center', PART_URI)).toBe(true);
expect(readTableHorizontalAlignment(xml, frame)).toBe('center');
expect(replaceTableHorizontalAlignment(xml, frame, 'center', PART_URI)).toBe(false);
expect(replaceTableHorizontalAlignment(xml, frame, 'left', PART_URI)).toBe(true);
expect(readTableHorizontalAlignment(xml, frame)).toBe('left');
expect(replaceTableHorizontalAlignment(xml, frame, undefined, PART_URI)).toBe(true);
expect(readTableHorizontalAlignment(xml, frame)).toBeUndefined();
expect(replaceTableHorizontalAlignment(xml, frame, undefined, PART_URI)).toBe(false);
```

Assert explicit left emits four direct `algn="l"` values, a value assignment inserts missing `pPr`, clear does not create absent `pPr`, and all unrelated XML remains exact. A malformed final physical cell must throw `ModelParseError`; because only the in-memory document was changed, the public-model test in Step 3 must prove no package bytes or journal entries commit.

- [ ] **Step 3: Write failing public model and SDK lifecycle tests**

In `packages/model/src/model.test.ts`, add a multi-row table with uniform center alignment. Require read isolation, exact center no-op, one-cell mixed projection, bulk justify, explicit left, clear, final right, stable table identity, duplicate isolation, source-only edit, slide move, all six formats, write/reopen, transaction rollback, and unchanged text/direction/vertical alignment/margins/borders/fill/fit/grid/rows/transform/relationships plus one untouched package-part hash.

Pass these runtime values through the setter and require `TypeError` plus unchanged slide bytes/journal:

```ts
[null, false, true, 0, '', 'l', 'ctr', 'dist', 'Left', ' left', [], {}, Symbol('align')]
```

Inject multiple paragraphs into the final physical cell, require getter `undefined`, then require assignment `ModelParseError` with zero package partial mutation. Add this compile-only type closure:

```ts
if (false) {
  const alignment: TextAlignment | undefined = table.horizontalAlignment;
  table.horizontalAlignment = 'left';
  table.horizontalAlignment = 'center';
  table.horizontalAlignment = 'right';
  table.horizontalAlignment = 'justify';
  table.horizontalAlignment = undefined;
  // @ts-expect-error unsupported table horizontal alignment
  table.horizontalAlignment = 'dist';
  void alignment;
}
```

In `packages/sdk/src/index.test.ts`, add the equivalent aggregate root lifecycle proof with validation diagnostics unchanged.

- [ ] **Step 4: Write failing real-PptxGenJS import tests**

In `packages/pptxgenjs-adapter/src/index.test.ts`, create real PptxGenJS 4.0.1 tables for uniform left/center/right/justify, omitted alignment, and table center with one right cell override:

```ts
slide.addTable([['A', 'B']], { x: 0.5, y: 0.5, w: 4, h: 1, align: 'center' });
slide.addTable([['A', 'B']], { x: 0.5, y: 2, w: 4, h: 1 });
slide.addTable([[
  { text: 'Inherited', options: {} },
  { text: 'Override', options: { align: 'right' } },
]], { x: 0.5, y: 3.5, w: 4, h: 1, align: 'center' });
```

Require uniform supported values to import exactly, omitted to project `undefined`, mixed to project `undefined` with cell matrix `['center', 'right']`, then normalize the mixed table to explicit `left`, write/reopen, and require table/cells to remain `left`.

- [ ] **Step 5: Run RED and require missing-helper/property failures**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-horizontal-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
```

Expected: FAIL because `readTableHorizontalAlignment`, `replaceTableHorizontalAlignment`, and `TableModel.horizontalAlignment` do not exist. Existing vertical-alignment and text-direction tests remain unchanged.

- [ ] **Step 6: Implement alignment-specific table helpers**

In `table-cell-horizontal-alignment.internal.ts`, import `readDirectTablePhysicalCells` and add:

```ts
export function readTableHorizontalAlignment(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): TextAlignment | undefined {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) return undefined;
  const first = readTableCellHorizontalAlignment(xml, cells[0]!);
  if (first === undefined) return undefined;
  return cells.every(
    (cell) => readTableCellHorizontalAlignment(xml, cell) === first,
  ) ? first : undefined;
}

export function replaceTableHorizontalAlignment(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  value: TextAlignment | undefined,
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
    changed = replaceTableCellHorizontalAlignment(xml, cell, value, partUri) || changed;
  }
  return changed;
}
```

Do not alter the existing cell reader/replacer or table creation renderer.

- [ ] **Step 7: Implement the live `TableModel` property**

In `packages/model/src/shapes.ts`, import both helpers with the existing cell functions and add before `textDirection`:

```ts
get horizontalAlignment(): TextAlignment | undefined {
  const { xml, element } = this.resolve();
  return readTableHorizontalAlignment(xml, element);
}

set horizontalAlignment(value: TextAlignment | undefined) {
  const alignment = value === undefined
    ? undefined
    : normalizeTextAlignment(value, 'Table horizontal alignment');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableHorizontalAlignment(
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

- [ ] **Step 8: Run focused GREEN, type, bundle, full regression, and performance gates**

Run the six-file focused command from Step 5, then:

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
node_modules/.bin/vitest run --reporter=dot --maxWorkers=4
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Record focused/full totals and performance time. Inspect `packages/pptx/dist/types/model/shapes.d.ts` and require both accessors inside `TableModel`.

- [ ] **Step 9: Review, commit, push, and verify**

Review direct-path ownership, explicit-left behavior, no-op behavior, missing-`pPr` insertion, multi-paragraph rejection, late-cell package atomicity, error types/messages, unrelated-state preservation, completed-property regression, duplicate/rollback/reopen, public types, and real PptxGenJS output. Then stage only the six declared files and run:

```sh
git diff --cached --check
git commit -m "feat: edit table-level horizontal alignment"
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
- Consumes: Task 1 `TableModel.horizontalAlignment`, packed `model/shapes.d.ts`, actual root/browser exports, and installed CLI.
- Produces: stable `tableHorizontalAlignment: true`, detailed `tableHorizontalAlignmentState`, and `tableHorizontalAlignmentInspect: true` evidence.

- [ ] **Step 1: Lock the packed `TableModel` declaration block**

In `scripts/build-npm-package-types.mjs` and installed-package smoke, slice the declaration from `export declare class TableModel` to `export declare class ChartModel` and require:

```ts
get horizontalAlignment(): TextAlignment | undefined;
set horizontalAlignment(value: TextAlignment | undefined);
```

Throw a horizontal-alignment-specific error if either accessor is absent. The check must not pass from `TableCell.horizontalAlignment` or another class.

- [ ] **Step 2: Add installed Node runtime evidence**

Create a dedicated two-by-two center-aligned table. Record this exact state:

```ts
const tableHorizontalAlignmentState = {
  uniform: 'center',
  readIsolation: true,
  noOp: true,
  mixed: null,
  overwritten: 'justify',
  overwrittenCells: ['justify', 'justify', 'justify', 'justify'],
  explicitLeft: 'left',
  explicitLeftCells: ['left', 'left', 'left', 'left'],
  cleared: null,
  clearedCells: [null, null, null, null],
  reopened: 'right',
  reopenedCells: ['right', 'right', 'right', 'right'],
  invalidError: {
    name: 'TypeError',
    message: 'Table horizontal alignment must be left, center, right, or justify',
  },
  failureIsolation: true,
  validationErrors: 0,
};
```

Use `table-horizontal-alignment-smoke.pptx`, preserve bytes/journal for getter and no-op, create mixed state with one cell right, overwrite justify, set explicit left, clear, reject `dist`, set final right, write/reopen, and add stable state to `apiChecks` plus top-level JSON.

- [ ] **Step 3: Add installed TypeScript and browser-condition evidence**

In the generated TypeScript consumer:

```ts
const typedTableHorizontalAlignment: TextAlignment | undefined =
  typedTable.horizontalAlignment;
typedTable.horizontalAlignment = 'left';
typedTable.horizontalAlignment = 'center';
typedTable.horizontalAlignment = 'right';
typedTable.horizontalAlignment = 'justify';
typedTable.horizontalAlignment = undefined;
// @ts-expect-error unsupported table-level horizontal alignment
typedTable.horizontalAlignment = 'dist';
```

Add the variable to the terminal `void` list. In the generated browser consumer, reproduce the exact Node state using `writeBlob()` and require the same matrices, invalid error, failure isolation, and zero diagnostics.

- [ ] **Step 4: Add installed CLI package inspection**

Validate, list slides, and read the final slide part from `table-horizontal-alignment-smoke.pptx`. Require one table shape, exactly four direct single-paragraph `pPr@algn="r"` tokens, and no `tcPr@algn` or `bodyPr@algn` false positive. Add `tableHorizontalAlignmentInspect: true` to the final summary.

- [ ] **Step 5: Extend the real-Chrome callback**

Add the same `tableHorizontalAlignmentState` and expected JSON to `scripts/playwright-browser-smoke.js`. Require exact read/no-op isolation, mixed projection, justify overwrite, explicit left, clear, right reopen, invalid failure isolation, zero document diagnostics, and zero validation/console/page/network errors.

- [ ] **Step 6: Build, pack, install, and run all proof gates**

Run:

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
node_modules/.bin/vitest run --reporter=dot --maxWorkers=4
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Retain actual tarball file count, SHA-256, installed Node/types/browser/CLI output, Chrome state, and evidence directory under `/tmp`; do not stage them.

- [ ] **Step 7: Review, commit, push, and verify**

Review declaration scoping, top-level JSON propagation, exact stable state, browser/Node parity, CLI regex ownership, final-right tokens, and absence of generated artifacts. Stage only the three scripts, then:

```sh
git diff --cached --check
git commit -m "test: verify packed table-level horizontal alignment"
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
- Consumes: final Task 1 API semantics and Task 2 measured source/package/browser evidence.
- Produces: one consistent public support statement, direct-state boundary, final evidence record, revised remaining-work list, and next advanced-table item.

- [ ] **Step 1: Add public examples and direct-state semantics**

In all three public README/API surfaces, add a compact example using uniform center, one-cell mixed state, bulk justify, explicit left, and `undefined` clear. State that the getter returns only a uniform safe direct single-paragraph value, never synthesizes absent state as left, and that per-cell mixed detail remains in `rows[].cells[].horizontalAlignment`.

- [ ] **Step 2: Update compatibility and progress records**

Move table-level horizontal-alignment consensus/bulk editing to supported in the compatibility matrix and detailed table section. Record real PptxGenJS legal final-state imports, omitted-to-`undefined`, native existing-deck extension, actual tarball count/SHA, focused/full/performance totals, installed Node/types/browser/CLI, real Chrome, error counts, evidence path, and all implementation commit hashes.

- [ ] **Step 3: Update changelog and remaining-work statements**

Add three changelog bullets for core semantics, PptxGenJS boundary, and actual-package/browser proof. Remove every stale statement that table-level horizontal-alignment getter/editor is pending. Select the next smallest advanced-table item after comparing table-level margin, border, and fill consensus/bulk editors; document the selected next item consistently while keeping overall parity approximately 97% until the final audit.

- [ ] **Step 4: Run final documentation and regression gates**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/table-cell-horizontal-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/vitest run --reporter=dot --maxWorkers=4
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 5: Review, commit, push, verify, and report**

Review examples, direct-state wording, PptxGenJS boundary, measured totals, hashes, evidence path, remaining-work consistency, and `git diff --check`. Stage only the six documentation files, then:

```sh
git commit -m "docs: document table-level horizontal alignment"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, report the completed item, remaining items, verification totals, commit hashes, and overall progress, then begin the selected next item.
