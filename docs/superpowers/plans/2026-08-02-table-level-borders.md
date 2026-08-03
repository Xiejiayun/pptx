# Table-Level Borders Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `TableModel.borders` uniform direct-state reading plus atomic all-physical-cell border replacement/clear, with PptxGenJS 4.0.1 final-state conformance and actual-package/browser proof.

**Architecture:** Extend the existing focused table-cell border helper with a strict table consensus reader and bulk replacer. Expose those functions through `TableModel.borders`, reuse the current descriptor-safe `TableCellBorderInput` normalizer and cell whole-replacement editor, and keep table styles, shared-edge resolution, diagonals, and advanced line state outside the public projection.

**Tech Stack:** TypeScript, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public API, tsup, npm pack, installed CLI, `pptx-inspect`, Playwright, and Google Chrome.

## Global Constraints

- Getter: `get borders(): TableCellBorders | undefined`.
- Setter: `set borders(value: TableCellBorderInput | undefined)`.
- Getter accepts only one uniform non-empty safe direct complete or partial L/R/T/B vector across every physical cell, including merge continuations.
- All-absent, mixed, malformed, unsupported, repeated, or ambiguous direct state reads as `undefined`; no table style, creation metadata, adjacent edge, or effective border is synthesized.
- Setter normalizes before package access and atomically whole-replaces all four owned sides on every physical cell in one slide commit.
- Scalar writes four sides; partial named/TRBL input clears omitted sides; empty/all-undefined/`undefined` clears all four sides.
- Direct absence, direct none, zero-width line, omitted style, and explicit solid remain distinct.
- Unique malformed or advanced side state may be replaced/cleared; repeated owned sides or missing/repeated `tcPr` reject with zero package mutation.
- Existing `TableCell.borders`, `setCellBorders()`, creation precedence/output, diagonals, fill, margins, alignments, direction, fit, geometry, relationships, and stable identity remain unchanged.
- Every task ends with review, one scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`, generated workspace tarballs, or retained `/tmp` evidence.

---

### Task 1: Core helper, public model, lifecycle, and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/table-cell-borders.internal.test.ts`
- Modify: `packages/model/src/table-cell-borders.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: existing `normalizeTableCellBorders()`, `readTableCellBorders()`, `replaceTableCellBorders()`, `readDirectTablePhysicalCells()`, `TableCellBorders`, and `TableCellBorderInput`.
- Produces: `readTableBorders(xml, frame)`, `replaceTableBorders(xml, frame, borders, partUri)`, and live `TableModel.borders` accessors.

- [ ] **Step 1: Add focused failing consensus-reader tests**

Create `packages/model/src/table-cell-borders.internal.test.ts` with local `cell()`, `parseFrame()`, and `parseSource()` fixtures modeled on the fill/margins focused tests. Import the not-yet-existing functions:

```ts
import {
  readTableBorders,
  replaceTableBorders,
} from './table-cell-borders.internal.js';
```

Define stable helpers for supported direct lines:

```ts
const none = (tag: string, width = '0') =>
  `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr">` +
  `<a:noFill/></a:${tag}>`;

const line = (
  tag: string,
  color = '4472C4',
  width = '12700',
  dash?: 'solid' | 'sysDash',
) =>
  `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr">` +
  `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
  `${dash ? `<a:prstDash val="${dash}"/>` : ''}` +
  `<a:round/><a:headEnd type="none" w="med" len="med"/>` +
  `<a:tailEnd type="none" w="med" len="med"/></a:${tag}>`;
```

Require these reader cases:

- uniform complete four-side none, scalar sRGB/theme line, omitted/solid/dash style, zero/fractional/max width;
- uniform partial `{ top, left }` vectors;
- numeric lexical equivalence (`12700` versus `012700`) and sRGB case equivalence;
- merge continuation cells through `gridSpan`, `hMerge`, and `vMerge`;
- all absent, present/absent, different side sets, none/line mixtures, different colors/widths/styles, omitted versus explicit solid;
- malformed width/color/dash, unsupported gradient/pattern line, duplicate owned side, missing/repeated `tcPr`, empty row/table, and ambiguous direct table path;
- diagonal, wrong-prefix, descendant, text-outline, and unrelated lookalikes ignored without changing `xml.changed`.

Expected examples:

```ts
expect(readTableBorders(complete.xml, complete.frame)).toEqual({
  top: { kind: 'line', color: { kind: 'srgb', value: '4472C4' }, width: 1 },
  right: { kind: 'none' },
  bottom: { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 2, style: 'dash' },
  left: { kind: 'none' },
});
expect(readTableBorders(absent.xml, absent.frame)).toBeUndefined();
expect(readTableBorders(mixed.xml, mixed.frame)).toBeUndefined();
```

- [ ] **Step 2: Add focused failing bulk-editor tests**

In the same file, cover a two-row/two-column table plus later unsafe cells. Assert:

- scalar none/line writes all four sides to all physical cells;
- exact TRBL and partial named input use whole replacement and remove omitted sides;
- `undefined` clears every L/R/T/B side;
- same semantic vector and all-absent clear preserve serialized bytes and `xml.changed`;
- one unique gradient/malformed side can be replaced or cleared;
- diagonal lines, fill, margins, alignment, direction, fit, text, extensions, transform, grid, and row bytes are preserved;
- an unsafe late cell throws `ModelParseError`; the later public-model test proves package bytes and the mutation journal stay unchanged because `TableModel` never commits the temporary XML.

Use exact final order assertions:

```ts
expect(serialized).toMatch(
  /<a:lnL[\s\S]*<a:lnR[\s\S]*<a:lnT[\s\S]*<a:lnB[\s\S]*<a:solidFill/,
);
```

- [ ] **Step 3: Run the new focused test and verify the expected failure**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-borders.internal.test.ts \
  --reporter=dot
```

Expected: FAIL because `readTableBorders` and `replaceTableBorders` are not exported.

- [ ] **Step 4: Implement the strict table consensus reader**

Modify `packages/model/src/table-cell-borders.internal.ts` to import the physical-cell helper:

```ts
import { readDirectTablePhysicalCells } from './table-physical-cells.internal.js';
```

Add the public helper:

```ts
export function readTableBorders(
  _xml: LosslessXmlDocument,
  frame: XmlElement,
): TableCellBorders | undefined {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) return undefined;
  const first = readStrictTableCellBorders(cells[0]!);
  if (!first) return undefined;
  return cells.slice(1).every((cell) => {
    const candidate = readStrictTableCellBorders(cell);
    return candidate !== undefined && tableCellBordersEqual(candidate, first);
  }) ? first : undefined;
}
```

Add `readStrictTableCellBorders(cell)` beside the tolerant cell reader. It must:

1. require exactly one direct local-name `tcPr`;
2. derive the lexical prefix from that owner;
3. for each public side, find only exact same-prefix direct `lnL/lnR/lnT/lnB` children;
4. reject repeated exact owned sides;
5. parse one present side with existing `readBorder()` and reject if unsupported;
6. return a non-empty partial object or `undefined` for all-absent/unsafe state.

Add semantic comparison without changing tolerant cell snapshots:

```ts
function tableCellBordersEqual(left: TableCellBorders, right: TableCellBorders): boolean {
  return PUBLIC_SIDES.every((side) => {
    const leftBorder = left[side];
    const rightBorder = right[side];
    if (leftBorder === undefined || rightBorder === undefined) {
      return leftBorder === rightBorder;
    }
    return bordersEqual(leftBorder, rightBorder);
  });
}
```

- [ ] **Step 5: Implement the atomic all-cell replacer**

Add:

```ts
export function replaceTableBorders(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  borders: TableCellBorders | undefined,
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
    changed = replaceTableCellBorders(xml, cell, borders, partUri) || changed;
  }
  return changed;
}
```

Keep all mutations on the in-memory slide XML. A late throw must occur before `TableModel` calls `setXml()`, preserving package bytes and the mutation journal.

- [ ] **Step 6: Run internal border tests and existing cell-border regression**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-borders.internal.test.ts \
  packages/model/src/model.test.ts \
  --reporter=dot
```

Expected: all tests pass; existing tolerant `TableCell.borders` and `setCellBorders()` behavior is unchanged.

- [ ] **Step 7: Add live `TableModel.borders` accessors**

Modify the border imports in `packages/model/src/shapes.ts`:

```ts
import {
  normalizeTableCellBorders,
  readTableBorders,
  readTableCellBorders,
  replaceTableBorders,
  replaceTableCellBorders,
} from './table-cell-borders.internal.js';
```

Add the property before `fill`:

```ts
get borders(): TableCellBorders | undefined {
  const { xml, element } = this.resolve();
  return readTableBorders(xml, element);
}

set borders(value: TableCellBorderInput | undefined) {
  const borders = normalizeTableCellBorders(value, 'Table borders');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableBorders(xml, element, borders, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

- [ ] **Step 8: Add public model typing and lifecycle coverage**

Extend `packages/model/src/model.test.ts` with a dedicated table-level borders test. Require:

```ts
const typedBorders: TableCellBorders | undefined = table.borders;
table.borders = { kind: 'none' };
table.borders = [
  { kind: 'none' },
  undefined,
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 2 },
  undefined,
];
table.borders = undefined;
// @ts-expect-error null is not a table border input
table.borders = null;
```

At runtime cover:

- immediate uniform creation read and detached nested-color snapshot;
- one-cell override to mixed `undefined` while `rows[].cells[].borders` retains detail;
- scalar, TRBL, partial, explicit none, empty, and clear broadcast;
- exact no-op and invalid-input failure isolation;
- unique advanced side replacement, late-cell repeated-side atomicity, transaction rollback;
- stable `TableModel` identity, duplicate isolation, slide move, all six formats, write/reopen;
- preservation of fill, margins, alignments, direction, fit, text, geometry, diagonal lines, unrelated shapes/parts, relationships, and diagnostics.

- [ ] **Step 9: Add SDK integration coverage**

Extend `packages/sdk/src/index.test.ts` with the same public sequence through `PptxDocument`. Test root exports, transaction rollback, invalid null/PptxGenJS-shaped input, duplicate/source isolation, write/reopen, and all-six-format state. Require no mutation for reads, exact no-op, and rejected input.

- [ ] **Step 10: Add real PptxGenJS final-state conformance**

Extend `packages/pptxgenjs-adapter/src/index.test.ts`. Generate with public PptxGenJS 4.0.1:

```ts
generatedSlide.addTable(
  [[{ text: 'Inherited', options: {} }, { text: 'Override', options: {
    border: { type: 'none' },
  } }]],
  { border: { type: 'dash', color: '4472C4', pt: 1.5 } },
);
generatedSlide.addTable([['Omitted']], {});
generatedSlide.addTable([['TRBL']], {
  border: [
    { type: 'solid', color: 'FF0000', pt: 1 },
    { type: 'none' },
    { type: 'dash', color: '70AD47', pt: 2 },
    { type: 'solid', color: '4472C4', pt: 3 },
  ],
});
```

Require uniform scalar/TRBL consensus, omitted four-side none, mixed `undefined`, exact per-cell snapshots, native normalization of the mixed imported table, and write/reopen preservation. Keep the existing empty/default-style/TRBL-zero difference assertions unchanged.

- [ ] **Step 11: Run focused, type, bundle, full, and performance gates**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-borders.internal.test.ts \
  packages/model/src/table-cell-fill.internal.test.ts \
  packages/model/src/table-cell-margins.internal.test.ts \
  packages/model/src/table-cell-horizontal-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Inspect `packages/pptx/dist/types/model/shapes.d.ts` and require both accessors inside `TableModel`.

- [ ] **Step 12: Review, commit, push, and verify core behavior**

Review strict side safety, partial equality, none/absence and omitted/solid distinctions, zero width, wrong-prefix/diagonal isolation, advanced replacement, no-ops, late-cell atomicity, public types, lifecycle, and real PptxGenJS output. Stage only the six declared files:

```sh
git add \
  packages/model/src/table-cell-borders.internal.test.ts \
  packages/model/src/table-cell-borders.internal.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit table-level borders"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0` before Task 2.

---

### Task 2: Actual-package declarations, Node/browser/CLI, and real-Chrome proof

**Files:**
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: Task 1 `TableModel.borders`, packed `model/shapes.d.ts`, root/browser exports, installed CLI, and checked-in Chrome callback.
- Produces: stable `tableBorders: true`, detailed `tableBordersState`, `tableBordersInspect: true`, and one retained packed PPTX.

- [ ] **Step 1: Lock the packed `TableModel` declaration block**

In both declaration checks, slice from `export declare class TableModel` to `export declare class ChartModel` and require:

```ts
get borders(): TableCellBorders | undefined;
set borders(value: TableCellBorderInput | undefined);
```

Throw a table-level-borders-specific error if either accessor is absent. The check must not pass from `TableCell.borders` or `AddTableOptions.border`.

- [ ] **Step 2: Add installed Node runtime evidence**

In `scripts/smoke-npm-package.mjs`, create a two-by-two table with a uniform scheme dash line. Canonicalize snapshots in public TRBL key order and nested color key order. Record this exact state shape:

```ts
const noneBorder = { kind: 'none' };
const fourBorders = (border) => ({
  top: border,
  right: border,
  bottom: border,
  left: border,
});
const dashSchemeLine = {
  kind: 'line',
  color: { kind: 'scheme', value: 'accent1' },
  width: 1.5,
  style: 'dash',
};
const partialSnapshot = {
  top: {
    kind: 'line',
    color: { kind: 'srgb', value: 'D9EAF7' },
    width: 2,
  },
  bottom: noneBorder,
};
const tableBordersState = {
  uniform: fourBorders(dashSchemeLine),
  readIsolation: true,
  noOp: true,
  mixed: null,
  partial: partialSnapshot,
  partialCells: Array(4).fill(partialSnapshot),
  none: fourBorders(noneBorder),
  noneCells: Array(4).fill(fourBorders(noneBorder)),
  cleared: null,
  clearedCells: [null, null, null, null],
  reopened: fourBorders(noneBorder),
  reopenedCells: Array(4).fill(fourBorders(noneBorder)),
  invalidError: {
    name: 'TypeError',
    message: 'Table borders must be an object',
  },
  failureIsolation: true,
  validationErrors: 0,
};
```

Mutate a detached nested color snapshot and preserve bytes/journal. Assign the equivalent initial scalar for no-op, make one cell none for mixed state, broadcast partial named state, broadcast scalar none, clear, reject `null`, set final scalar none, write `table-borders-smoke.pptx`, and reopen. Add the boolean/state to `apiChecks` and top-level JSON.

Before removing the temporary consumer workspace, copy `table-borders-smoke.pptx` beside the input tarball for retained host inspection.

- [ ] **Step 3: Add installed TypeScript and browser-condition evidence**

Add `type TableCellBorders` to the generated consumer import, then use:

```ts
const typedTableBorders: TableCellBorders | undefined = typedTable.borders;
typedTable.borders = { kind: 'none' };
typedTable.borders = [
  { kind: 'none' },
  undefined,
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 2 },
  undefined,
];
typedTable.borders = undefined;
// @ts-expect-error table borders reject null
typedTable.borders = null;
// @ts-expect-error table borders reject PptxGenJS-shaped input
typedTable.borders = { type: 'solid', color: '4472C4', pt: 1 };
void typedTableBorders;
```

In the generated browser consumer, reproduce the exact Node state with `writeBlob()` and require identical snapshots, invalid error, failure isolation, and zero diagnostics.

- [ ] **Step 4: Add installed CLI and `pptx-inspect` evidence**

Use the installed CLI to inspect, validate, list slides, and read the final slide part from `table-borders-smoke.pptx`. Require one table shape, four physical cells, and exactly one same-prefix direct `lnL/lnR/lnT/lnB` none line per cell. Inspect only those four direct line owners: reject diagonal/text-line false positives and any remaining solid/gradient/pattern fill beneath an owned L/R/T/B line without rejecting the independent cell fill.

Run the required broad-to-narrow host workflow:

```sh
pptx-inspect --json package inspect table-borders-smoke.pptx
pptx-inspect --json package validate table-borders-smoke.pptx --profile powerpoint-2010
pptx-inspect --json slides list table-borders-smoke.pptx
pptx-inspect --json part read table-borders-smoke.pptx /ppt/slides/slide1.xml
```

Require zero validation errors and warnings before accepting the exact slide-part read. Add `tableBordersInspect: true` to the compact result.

- [ ] **Step 5: Extend the real-Chrome callback**

Add the same `tableBordersState` and expected JSON to `scripts/playwright-browser-smoke.js`. Require detached-read/no-op isolation, mixed projection, partial/none overwrite, clear, none reopen, invalid failure isolation, zero document diagnostics, and global `errorCounts: { console: 0, page: 0, network: 0 }`.

- [ ] **Step 6: Build, pack, install, and run all proof gates**

Run:

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
table_borders_artifacts=$(mktemp -d /tmp/pptx-table-borders-artifacts.XXXXXX)
(cd packages/pptx && npm pack --ignore-scripts --pack-destination "$table_borders_artifacts")
node scripts/smoke-npm-package.mjs "$table_borders_artifacts/jiayunxie-pptx-0.1.0.tgz"
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Extract the actual tarball beneath the same fresh evidence directory, serve its browser module over loopback, and run the checked-in callback in installed Google Chrome through the bundled Playwright runtime. Retain tarball file count, SHA-256, installed Node/types/browser/CLI output, `pptx-inspect` JSON, full/compact Chrome state, and evidence directory outside the repository.

- [ ] **Step 7: Review, commit, push, and verify package proof**

Review declaration scoping, stable state/key order, detached getter proof, Node/browser parity, exact direct-cell ownership, diagonal/text false-positive rejection, prior-field preservation, zero browser errors, and absence of generated repository artifacts. Stage only the three scripts:

```sh
git add \
  scripts/build-npm-package-types.mjs \
  scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed table-level borders"
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
- Consumes: final Task 1 semantics and Task 2 measured source/package/browser/CLI evidence.
- Produces: one consistent public support statement, PptxGenJS direct-state boundary, final evidence record, revised remaining-work list, and the next advanced-table item.

- [ ] **Step 1: Add public examples and direct-state semantics**

In all three public README/API surfaces, add one compact example with uniform line read, one-cell mixed state, partial broadcast, scalar none broadcast, and `undefined` clear. State that the getter returns only one uniform supported safe non-empty direct vector, never resolves table styles/shared edges/defaults, and that mixed detail remains in `rows[].cells[].borders`.

- [ ] **Step 2: Update compatibility and progress records**

Move table-level borders consensus/bulk editing to supported in the compatibility matrix and detailed table section. Record PptxGenJS legal uniform/mixed/omitted final states, native existing-deck extension, direct absence/none, omitted/solid, zero-width differences, focused/full/performance totals, actual tarball count/SHA, installed Node/types/browser/CLI, `pptx-inspect`, real Chrome, zero error counts, evidence path, and all implementation commit hashes.

- [ ] **Step 3: Update changelog and remaining-work statements**

Add three changelog bullets for core semantics, PptxGenJS direct-state boundary, and actual-package/browser proof. Remove every stale statement that table-level border getter/editor is pending. Choose the next smallest independent parity item by comparing table-level fit, advanced cell text/style, merge/hyperlink, row/column CRUD, auto-page/repeated headers, and `tableToSlides`; keep overall parity approximately 97% until the final peer/client audit.

- [ ] **Step 4: Run final documentation and regression gates**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/table-cell-borders.internal.test.ts \
  packages/model/src/table-cell-fill.internal.test.ts \
  packages/model/src/table-cell-margins.internal.test.ts \
  packages/model/src/table-cell-horizontal-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 5: Review, commit, push, verify, and report**

Review examples, exact getter/setter types, partial whole replacement, direct none versus absence, omitted versus explicit solid, zero-width line, PptxGenJS boundary, measured totals/hashes/evidence, remaining-work consistency, and `git diff --check`. Stage only the six documentation files:

```sh
git add \
  README.md \
  packages/pptx/README.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  docs/implementation-progress.md \
  CHANGELOG.md
git diff --cached --check
git commit -m "docs: document table-level borders"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, report completed/remaining items and overall progress, then begin the selected next parity design item.
