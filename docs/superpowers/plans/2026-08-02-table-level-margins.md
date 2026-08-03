# Table-Level Margins Implementation Plan

> **For agentic workers:** Execute this plan task-by-task inline in the current session. The standing user direction requires autonomous inline execution without subagents; review, commit, push, fetch, and verify each task before starting its successor. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a live `TableModel.margins` strict direct-state consensus getter and atomic all-physical-cell whole-replacement setter, then prove it through source, PptxGenJS import, actual-package Node/browser/CLI, and real Chrome.

**Architecture:** `table-cell-margins.internal.ts` retains ownership of signed-Int32 point/EMU parsing and lossless direct `tcPr` attribute replacement while adding a stricter table-consensus read and reusing `readDirectTablePhysicalCells()`. `TableModel` exposes separate getter/setter types over the existing public margin values; no table metadata, effective-style calculation, or layout recomputation is introduced.

**Tech Stack:** TypeScript 5.8, Vitest, lossless OOXML source-span editing, PptxGenJS 4.0.1, tsup, actual npm tarball smoke, installed `pptx-inspect`, Playwright, and installed Google Chrome.

## Global Constraints

- Public API is exactly `get margins(): TextBoxMargins | undefined` and `set margins(value: TextBoxMarginInput | undefined)`.
- Getter returns a value only when one or more physical cells all expose the same non-empty, structurally safe direct complete or partial margin vector.
- Getter compares parsed signed Int32 values, so lexically different equal integer tokens are equivalent.
- All-absent, mixed side sets, mixed values, empty, ambiguous, malformed, repeated, or out-of-range owned state returns `undefined`.
- Getter does not synthesize canonical defaults, table style values, creation metadata, or effective inherited values.
- Setter normalizes once with `normalizeTextBoxMargins(value, 'Table margins')` before package access.
- Scalar and TRBL assignment write all four sides; partial named input whole-replaces all cells and clears omitted sides; `{}` and `undefined` clear all four sides.
- Setter edits every physical `a:tc`, including merge continuations, serializes the slide at most once, and is an exact no-op when no direct numeric state changes.
- Missing/repeated direct `tcPr`, repeated owned attributes, incomplete direct table paths, or empty physical-cell sets reject edits with `ModelParseError` and zero package partial mutation.
- One malformed single margin token may be replaced by a legal value or cleared, matching the existing single-cell editor.
- Do not change `TableCell.margins`, `setCellMargins()`, table creation defaults/precedence/output, point-only native units, or unrelated table properties.
- Do not add effective margin calculation, a `mixed` token, matrix getter, table metadata, extension element, alias type, layout recomputation, or higher-order table-property framework.
- Existing `TableModel.horizontalAlignment`, `textDirection`, and `verticalAlignment` behavior and tests must remain unchanged.
- Every task ends with review, scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`, generated workspace tarballs, or retained `/tmp` evidence.

---

## File map

- Create `packages/model/src/table-cell-margins.internal.test.ts`: strict direct consensus, numeric equivalence, bulk replacement, clearing, preservation, unsafe-state, and physical-cell coverage.
- Modify `packages/model/src/table-cell-margins.internal.ts`: strict table-level reader and all-cell replacer.
- Modify `packages/model/src/shapes.ts`: public live `TableModel.margins` getter/setter.
- Modify `packages/model/src/model.test.ts`: lifecycle, detachment, isolation, rollback, malformed package, and type behavior.
- Modify `packages/sdk/src/index.test.ts`: aggregate SDK/root creation/edit/write/reopen proof and public type closure.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: real PptxGenJS omitted, uniform, mixed import, unit boundary, and native normalization proof.
- Modify `scripts/build-npm-package-types.mjs`: require both accessors inside the packed `TableModel` declaration block.
- Modify `scripts/smoke-npm-package.mjs`: declarations, Node, TypeScript, browser condition, CLI part inspection, and stable top-level state.
- Modify `scripts/playwright-browser-smoke.js`: real-Chrome uniform/mixed/scalar/partial/clear/reopen and zero-error state.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: public contract, PptxGenJS unit boundary, evidence, progress, and next item.

---

### Task 1: Core property, source tests, and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/table-cell-margins.internal.test.ts`
- Modify: `packages/model/src/table-cell-margins.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: `readDirectTablePhysicalCells()`, `readTableCellMargins()`, `replaceTableCellMargins()`, `normalizeTextBoxMargins()`, live `TableModel.resolve()`, and existing table creation/cell editing APIs.
- Produces: `readTableMargins(xml, frame)`, `replaceTableMargins(xml, frame, margins, partUri)`, and public `TableModel.margins`.

- [ ] **Step 1: Write failing strict consensus tests**

Create `packages/model/src/table-cell-margins.internal.test.ts` with a `LosslessXmlDocument` fixture using the exact direct `graphicFrame/graphic/graphicData/tbl/tr/tc/tcPr` path. Import the not-yet-implemented functions:

```ts
import {
  readTableMargins,
  replaceTableMargins,
} from './table-cell-margins.internal.js';
```

Require complete uniform vectors and partial uniform vectors to read exactly in points. Require `12700` and `0012700` to compare as the same 1pt direct value. Require all-absent, present/absent mixed, differing keys, differing values, malformed integer, out-of-range integer, repeated margin attribute, missing/repeated `tcPr`, empty row/table, ambiguous direct table paths, and descendant/qualified lookalikes to return `undefined`. Include physical merge-continuation cells in the consensus.

- [ ] **Step 2: Write failing bulk replacement and atomicity tests**

Use a two-by-two fixture containing text, runs, body insets, alignments, direction, fit, borders, fill, grid, rows, transform, extensions, and one merge continuation. Require:

```ts
expect(replaceTableMargins(xml, frame, {
  top: 6, right: 6, bottom: 6, left: 6,
}, PART_URI)).toBe(true);
expect(readTableMargins(xml, frame)).toEqual({
  top: 6, right: 6, bottom: 6, left: 6,
});
expect(replaceTableMargins(xml, frame, {
  top: 6, right: 6, bottom: 6, left: 6,
}, PART_URI)).toBe(false);
expect(replaceTableMargins(xml, frame, { top: 2, left: 4 }, PART_URI))
  .toBe(true);
expect(readTableMargins(xml, frame)).toEqual({ top: 2, left: 4 });
expect(replaceTableMargins(xml, frame, undefined, PART_URI)).toBe(true);
expect(readTableMargins(xml, frame)).toBeUndefined();
expect(replaceTableMargins(xml, frame, undefined, PART_URI)).toBe(false);
```

Assert that partial input removes right/bottom on every cell, clear removes only `marL/marR/marT/marB`, valid assignment repairs one malformed single token, and all unrelated XML remains exact. A repeated owned attribute on the final physical cell must throw `ModelParseError`; the public-model test in Step 3 proves that no partially edited slide bytes or journal entries commit.

- [ ] **Step 3: Write failing public model and SDK lifecycle tests**

In `packages/model/src/model.test.ts`, create a multi-row table with uniform `[3.6, 7.2, 10.8, 14.4]` margins. Require detached getter snapshots, read isolation, exact uniform no-op, one-cell mixed projection, scalar 6pt overwrite, partial `{ top: 2, left: 4 }` whole replacement, `{}` clear, final `[1, 2, 3, 4]`, stable table identity, duplicate isolation, source-only edit, slide move, all six formats, write/reopen, transaction rollback, and unchanged text/alignment/direction/vertical alignment/borders/fill/fit/grid/rows/transform/relationships plus one untouched package-part hash.

Pass these runtime values through the setter and require the existing precise `TypeError` or `RangeError` plus unchanged slide bytes/journal:

```ts
[
  null,
  false,
  true,
  '',
  [1, 2, 3],
  [1, 2, 3, 4, 5],
  { middle: 1 },
  { top: Number.NaN },
  Object.create({ top: 1 }),
  Symbol('margins'),
]
```

Inject a repeated `marT` on the final physical cell, require getter `undefined`, then require assignment `ModelParseError` with zero package partial mutation. Add this compile-only type closure:

```ts
if (false) {
  const margins: TextBoxMargins | undefined = table.margins;
  table.margins = 6;
  table.margins = [1, 2, 3, 4];
  table.margins = { top: 2, left: 4 };
  table.margins = {};
  table.margins = undefined;
  // @ts-expect-error table margins reject null
  table.margins = null;
  // @ts-expect-error table margin tuple requires four values
  table.margins = [1, 2, 3];
  void margins;
}
```

In `packages/sdk/src/index.test.ts`, add the equivalent aggregate root lifecycle proof with unchanged validation diagnostics.

- [ ] **Step 4: Write failing real-PptxGenJS import tests**

In `packages/pptxgenjs-adapter/src/index.test.ts`, create real PptxGenJS 4.0.1 tables for:

```ts
slide.addTable([['A', 'B']], {
  x: 0.5, y: 0.5, w: 4, h: 1,
  margin: [0.05, 0.1, 0.15, 0.2],
});
slide.addTable([['A', 'B']], {
  x: 0.5, y: 2, w: 4, h: 1,
});
slide.addTable([[
  { text: 'Inherited', options: {} },
  { text: 'Override', options: { margin: 0 } },
]], {
  x: 0.5, y: 3.5, w: 4, h: 1,
  margin: 0.1,
});
```

Require the first table to project `{ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 }`, omitted creation to project canonical `{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }`, and the override table to project `undefined` with the exact inherited 7.2pt/override 0pt cell matrix. Normalize the mixed table to native `[1, 2, 3, 4]`, write/reopen, and require table/cells to retain the exact point vector.

- [ ] **Step 5: Run RED and require missing-helper/property failures**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-margins.internal.test.ts \
  packages/model/src/table-cell-horizontal-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
```

Expected: FAIL because `readTableMargins`, `replaceTableMargins`, and `TableModel.margins` do not exist. Existing table-level alignment/direction tests remain unchanged.

- [ ] **Step 6: Implement strict table-margin helpers**

In `table-cell-margins.internal.ts`, import `readDirectTablePhysicalCells` and add:

```ts
export function readTableMargins(
  _xml: LosslessXmlDocument,
  frame: XmlElement,
): TextBoxMargins | undefined {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) return undefined;
  const first = readStrictTableCellMargins(cells[0]!);
  if (!first) return undefined;
  return cells.slice(1).every((cell) => {
    const candidate = readStrictTableCellMargins(cell);
    return candidate !== undefined && SIDES.every(
      ([side]) => candidate[side] === first[side],
    );
  }) ? first : undefined;
}

export function replaceTableMargins(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  margins: TextBoxMargins | undefined,
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
    changed = replaceTableCellMargins(xml, cell, margins, partUri) || changed;
  }
  return changed;
}
```

Add a private strict cell reader that requires exactly one direct `tcPr`, zero or one exact unqualified attribute per side, valid signed Int32 integer tokens for every present owned attribute, and at least one present side:

```ts
function readStrictTableCellMargins(cell: XmlElement): TextBoxMargins | undefined {
  const directProperties = directChildren(cell, 'tcPr');
  if (directProperties.length !== 1) return undefined;
  const margins: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const [side, attributeName] of SIDES) {
    const attributes = directProperties[0]!.attributes.filter(
      ({ name }) => name === attributeName,
    );
    if (attributes.length > 1) return undefined;
    if (attributes.length === 0) continue;
    const raw = parseRaw(attributes[0]!.value);
    if (raw === undefined) return undefined;
    margins[side] = raw / EMU_PER_POINT;
  }
  return Object.keys(margins).length > 0 ? margins : undefined;
}
```

Do not alter `readTableCellMargins()`, `replaceTableCellMargins()`, normalization, rendering, or creation defaults.

- [ ] **Step 7: Implement the live `TableModel` property**

In `packages/model/src/shapes.ts`, import both new table helpers and add after `horizontalAlignment`:

```ts
get margins(): TextBoxMargins | undefined {
  const { xml, element } = this.resolve();
  return readTableMargins(xml, element);
}

set margins(value: TextBoxMarginInput | undefined) {
  const margins = normalizeTextBoxMargins(value, 'Table margins');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableMargins(xml, element, margins, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

- [ ] **Step 8: Run focused GREEN, type, bundle, full regression, and performance gates**

Run the seven-file focused command from Step 5, then:

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
node_modules/.bin/vitest run --reporter=dot --maxWorkers=4
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Record focused/full totals and performance time. Inspect `packages/pptx/dist/types/model/shapes.d.ts` and require both margin accessors inside `TableModel`.

- [ ] **Step 9: Review, commit, push, and verify**

Review strict-versus-tolerant reader isolation, complete/partial consensus, numeric lexical equality, scalar/TRBL/partial/clear behavior, no-ops, late-cell package atomicity, error types/messages, unrelated-state preservation, completed-property regression, duplicate/rollback/reopen, public types, and real PptxGenJS output. Stage only the six declared files and run:

```sh
git diff --cached --check
git commit -m "feat: edit table-level margins"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
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
- Consumes: Task 1 `TableModel.margins`, packed `model/shapes.d.ts`, actual root/browser exports, and installed CLI.
- Produces: stable `tableMargins: true`, detailed `tableMarginsState`, and `tableMarginsInspect: true` evidence.

- [ ] **Step 1: Lock the packed `TableModel` declaration block**

In `scripts/build-npm-package-types.mjs` and installed-package smoke, slice from `export declare class TableModel` to `export declare class ChartModel` and require:

```ts
get margins(): TextBoxMargins | undefined;
set margins(value: TextBoxMarginInput | undefined);
```

Throw a table-margin-specific error if either accessor is absent. The check must not pass from `TableCell.margins` or `ShapeModel.textMargins`.

- [ ] **Step 2: Add installed Node runtime evidence**

Create a dedicated two-by-two table with `[3.6, 7.2, 10.8, 14.4]`. Record this exact state:

```ts
const tableMarginsState = {
  uniform: { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
  readIsolation: true,
  noOp: true,
  mixed: null,
  overwritten: { top: 6, right: 6, bottom: 6, left: 6 },
  overwrittenCells: [
    { top: 6, right: 6, bottom: 6, left: 6 },
    { top: 6, right: 6, bottom: 6, left: 6 },
    { top: 6, right: 6, bottom: 6, left: 6 },
    { top: 6, right: 6, bottom: 6, left: 6 },
  ],
  partial: { top: 2, left: 4 },
  partialCells: [
    { top: 2, left: 4 },
    { top: 2, left: 4 },
    { top: 2, left: 4 },
    { top: 2, left: 4 },
  ],
  cleared: null,
  clearedCells: [null, null, null, null],
  reopened: { top: 1, right: 2, bottom: 3, left: 4 },
  reopenedCells: [
    { top: 1, right: 2, bottom: 3, left: 4 },
    { top: 1, right: 2, bottom: 3, left: 4 },
    { top: 1, right: 2, bottom: 3, left: 4 },
    { top: 1, right: 2, bottom: 3, left: 4 },
  ],
  invalidError: {
    name: 'TypeError',
    message: 'Table margins must be a number, four-value tuple, or margin object',
  },
  failureIsolation: true,
  validationErrors: 0,
};
```

Use `table-margins-smoke.pptx`. Mutate the first detached getter result and preserve bytes/journal to prove read isolation; assign the equivalent initial tuple for no-op; create mixed state with one cell `{ top: 9 }`; overwrite scalar 6; assign partial `{ top: 2, left: 4 }`; clear; reject `null`; set final `[1, 2, 3, 4]`; write/reopen. Add the boolean and state to `apiChecks` plus top-level JSON.

Canonicalize every getter/cell snapshot to explicit TRBL key order before exact JSON comparison:

```js
const packedTableMarginsSnapshot = (value) => value === undefined
  ? null
  : {
      top: value.top,
      right: value.right,
      bottom: value.bottom,
      left: value.left,
    };
```

- [ ] **Step 3: Add installed TypeScript and browser-condition evidence**

In the generated TypeScript consumer:

```ts
const typedTableMargins: TextBoxMargins | undefined = typedTable.margins;
typedTable.margins = 6;
typedTable.margins = [1, 2, 3, 4];
typedTable.margins = { top: 2, left: 4 };
typedTable.margins = {};
typedTable.margins = undefined;
// @ts-expect-error table margins reject null
typedTable.margins = null;
// @ts-expect-error table margin tuple requires four values
typedTable.margins = [1, 2, 3];
```

Add the variable to the terminal `void` list. In the generated browser consumer, reproduce the exact Node state using `writeBlob()` and require identical vectors, invalid error, failure isolation, and zero diagnostics.

- [ ] **Step 4: Add installed CLI package inspection**

Validate, list slides, and read the final slide part from `table-margins-smoke.pptx`. Require one table shape and exactly four direct `tcPr` elements whose owned values are:

```json
{ "marL": "50800", "marR": "25400", "marT": "12700", "marB": "38100" }
```

Reject missing/extra owned values and any text-body `lIns/rIns/tIns/bIns` false positive. Add `tableMarginsInspect: true` to the final summary.

- [ ] **Step 5: Extend the real-Chrome callback**

Add the same `tableMarginsState` and expected JSON to `scripts/playwright-browser-smoke.js`. Require detached/read/no-op isolation, mixed projection, scalar overwrite, partial whole replacement, clear, tuple reopen, invalid failure isolation, zero document diagnostics, and global `errorCounts: { console: 0, page: 0, network: 0 }`.

- [ ] **Step 6: Build, pack, install, and run all proof gates**

Run:

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
table_margins_artifacts=$(mktemp -d /tmp/pptx-table-margins-artifacts.XXXXXX)
(cd packages/pptx && npm pack --ignore-scripts --pack-destination "$table_margins_artifacts")
node scripts/smoke-npm-package.mjs "$table_margins_artifacts/jiayunxie-pptx-0.1.0.tgz"
node_modules/.bin/vitest run --reporter=dot --maxWorkers=4
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Extract the actual tarball beneath the same fresh `/tmp/pptx-table-margins-artifacts.XXXXXX` directory, serve its browser module over loopback, and run the checked-in callback in installed Google Chrome through the bundled Playwright runtime. Retain actual tarball file count, SHA-256, installed Node/types/browser/CLI output, full and compact Chrome state, and evidence directory; do not stage them.

- [ ] **Step 7: Review, commit, push, and verify**

Review declaration scoping, exact stable state/key order, detached getter proof, browser/Node parity, CLI ownership, final TRBL vector, prior-field preservation, and absence of generated repository artifacts. Stage only the three scripts, then:

```sh
git diff --cached --check
git commit -m "test: verify packed table-level margins"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
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
- Produces: one consistent public support statement, direct-state and unit boundary, final evidence record, revised remaining-work list, and next advanced-table item.

- [ ] **Step 1: Add public examples and direct-state semantics**

In all three public README/API surfaces, add a compact example using uniform TRBL, one-cell mixed state, scalar overwrite, partial whole replacement, and `undefined` clear. State that the getter returns only one uniform non-empty safe direct complete/partial vector, does not synthesize canonical defaults for absent attributes, and that per-cell mixed detail remains in `rows[].cells[].margins`.

- [ ] **Step 2: Update compatibility and progress records**

Move table-level margin consensus/bulk editing to supported in the compatibility matrix and detailed table section. Record PptxGenJS's less-than-one inch input boundary versus native points, legal final-state imports, omitted canonical vector, native existing-deck extension, actual tarball count/SHA, focused/full/performance totals, installed Node/types/browser/CLI, real Chrome, error counts, evidence path, and all implementation commit hashes.

- [ ] **Step 3: Update changelog and remaining-work statements**

Add three changelog bullets for core semantics, PptxGenJS unit/direct-state boundary, and actual-package/browser proof. Remove every stale statement that table-level margin getter/editor is pending. Select table-level border or fill as the next item by comparing remaining helper complexity, and document the selection consistently while keeping overall parity approximately 97% until the final audit.

- [ ] **Step 4: Run final documentation and regression gates**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/table-cell-margins.internal.test.ts \
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

Review examples, getter/setter type distinction, complete/partial direct-state wording, PptxGenJS unit boundary, measured totals, hashes, evidence path, remaining-work consistency, and `git diff --check`. Stage only the six documentation files, then:

```sh
git commit -m "docs: document table-level margins"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, report the completed item, remaining items, verification totals, commit hashes, and overall progress, then begin the selected next item.
