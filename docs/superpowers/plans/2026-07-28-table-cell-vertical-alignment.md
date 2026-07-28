# Table Cell Vertical Alignment Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, lossless table-cell vertical-alignment snapshots and physical-cell editing for existing PPTX tables with PptxGenJS 4.0.1 output conformance.

**Architecture:** Add a dedicated table-cell codec for direct `tcPr@anchor`, while reusing the existing public `TextBoxVerticalAlignment` union and its input normalizer. Expose the direct state through immutable `TableCell.verticalAlignment` snapshots and mutate it through `TableModel.setCellVerticalAlignment()`, preserving text-body `bodyPr@anchor`, fit, direction, cell formatting, and all neighboring content.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Reuse `TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom'`; do not add a duplicate table-only union.
- Public read surface is optional `TableCell.verticalAlignment`; public write surface is `TableModel.setCellVerticalAlignment(rowIndex, columnIndex, value)`.
- Storage is the selected physical cell's direct `a:tcPr@anchor`, not `a:bodyPr@anchor`, a table-level default, or an effective inherited value.
- Strict read maps only exact unqualified `t`, `ctr`, and `b`; absent, malformed, repeated, namespaced, long-form, justified/distributed, or unknown state returns `undefined`.
- Setter maps top/middle/bottom to canonical `t/ctr/b`; `undefined` removes only the direct unqualified anchor.
- Setter requires exactly one direct `tcPr`; same canonical assignment is an exact no-op, and a single unknown token may be replaced or cleared.
- PptxGenJS 4.0.1 materializes omitted table/cell `valign` as direct `anchor="ctr"`; table-level valid values copy to cells and cell-level values override them.
- Do not add table creation, table-level valign mutation, effective inheritance, just/dist support, margin/fill/border APIs, merge mutation, rich text, row/column mutation, or dynamic layout.
- Preserve `tcPr` ordering/children, direction, bodyPr anchor/fit, text, paragraphs/runs, merge state, neighboring cells, relationships, and `TableModel` identity.
- Implement inline without subagent delegation, as required for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, empty native/baseline package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Strict direct-anchor snapshot and lossless model mutation

**Files:**
- Create: `packages/model/src/table-cell-vertical-alignment.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `TextBoxVerticalAlignment`, `normalizeTextBoxVerticalAlignment()`, `TableModel.rows`, and nested OPC transactions.
- Produces: `TableCell.verticalAlignment`, `readTableCellVerticalAlignment()`, `replaceTableCellVerticalAlignment()`, and `TableModel.setCellVerticalAlignment()`.

- [ ] **Step 1: Add failing strict-read tests**

Replace one fixture table row with physical cells whose direct `tcPr@anchor` values are `t`, `ctr`, `b`, absent, empty, uppercase, whitespace, long-form `top/middle/bottom`, `just`, `dist`, unknown, namespaced, and duplicate. Include a cell with `bodyPr@anchor="b"` but no `tcPr@anchor`, plus repeated and missing direct `tcPr`. Assert:

```ts
const alignments = table.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment);
expect(alignments.slice(0, 3)).toEqual(['top', 'middle', 'bottom']);
expect(alignments.slice(3)).toEqual(Array(15).fill(undefined));
expect(pkg.mutations).toEqual(journal);
```

Mutate a returned snapshot through a test-only cast and prove a fresh `table.rows` read still returns the source value.

- [ ] **Step 2: Run the model suite and verify red**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new snapshot assertions fail because `TableCell.verticalAlignment` does not exist.

- [ ] **Step 3: Implement the dedicated codec and snapshot field**

Create the table-cell codec with exact mappings:

```ts
const FROM_OOXML = new Map<string, TextBoxVerticalAlignment>([
  ['t', 'top'],
  ['ctr', 'middle'],
  ['b', 'bottom'],
]);

const TO_OOXML: Readonly<Record<TextBoxVerticalAlignment, string>> = {
  top: 't',
  middle: 'ctr',
  bottom: 'b',
};
```

`readTableCellVerticalAlignment()` must require one direct `tcPr`, exactly one unqualified `anchor`, and one mapped token. `replaceTableCellVerticalAlignment()` must parse only that `tcPr` source slice, reject repeated direct anchor attributes, perform a no-op before patching when the canonical token already matches, and otherwise add/replace/remove only that attribute. In `TableCell`, add `readonly verticalAlignment?: TextBoxVerticalAlignment`; in `TableModel.rows`, conditionally spread only a defined value.

- [ ] **Step 4: Add failing edit, isolation, and rollback tests**

Build cells with self-closing and expanded `tcPr`, single/double quote styles, margins, `vert`, `horzOverflow`, fill/borders, unknown attributes/children, direct bodyPr anchor and fit, unknown anchor, a neighboring bottom-aligned cell, a merged placeholder, repeated `anchor`, and repeated/missing `tcPr`.

Cover middle same-mode no-op, absent→top, top→middle, middle→bottom, bottom→clear, unknown→canonical, unknown→clear, and merged-cell editing. Run `setCellText()`, `setCellTextDirection()`, `setCellTextFit()`, and transform edits and prove anchor survives. Invalid indices/structures must change no bytes. Roll back two alignment edits inside an outer transaction and compare bytes, journal, table identity, and fresh snapshots.

- [ ] **Step 5: Implement indexed validation and transactional mutation**

Add the public method:

```ts
setCellVerticalAlignment(
  rowIndex: number,
  columnIndex: number,
  value: TextBoxVerticalAlignment | undefined,
): void {
  const alignment = value === undefined
    ? undefined
    : normalizeTextBoxVerticalAlignment(value, 'Table cell vertical alignment');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    if (replaceTableCellVerticalAlignment(xml, cell, alignment, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

Use the existing value normalizer; do not modify text-box vertical-alignment behavior.

- [ ] **Step 6: Re-run the model suite**

Expected: strict snapshots, canonical edits, same-mode/absent no-ops, malformed structure, bodyPr/format isolation, merged cell, and rollback tests pass.

### Task 2: SDK lifecycle and invalid-input isolation

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public field/method, the existing editable table fixture, duplicate, write/reopen, `PptxDocument.transaction()`, cell text/direction/fit edits, and transforms.
- Produces: public lifecycle, cross-property preservation, and package-isolation evidence without table creation APIs.

- [ ] **Step 1: Extend the existing table fixture with direct anchors**

On the current fit row, retain fit/body metadata and give its cells `anchor="t"`, `"ctr"`, `"b"`, absent, and a merged `anchor="ctr"`. Keep unknown tcPr attributes and the existing direction values so all three table-cell properties coexist in one public fixture.

- [ ] **Step 2: Cover the public edit lifecycle**

Open the fixture, retain slide/table identity, and duplicate before editing. Read top/middle/bottom/undefined/middle; set all three canonical values, clear through `undefined`, edit fit/direction/text and transform, then mutate a detached rows snapshot. Assert current versus duplicate alignments, text, fit, direction, merged placeholder, unknown XML, neighbor isolation, and stable identity.

Roll back alignment changes inside `document.transaction()` and assert exact bytes/journal/live state. Write/reopen and compare edited versus duplicated rows.

- [ ] **Step 3: Reject invalid values and coordinates before mutation**

For `null`, booleans, numbers, empty/case/whitespace strings, `mid`, `center`, `just`, `dist`, objects, arrays, and symbols, assert `TypeError`. For negative, fractional, NaN, infinite, and out-of-range coordinates, assert `RangeError`. Confirm bytes, journal, slide/shape counts, table identity, text, directions, fits, and alignments remain unchanged.

- [ ] **Step 4: Run typecheck and focused suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck and both suites pass.

### Task 3: PptxGenJS conformance, docs, and packed surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public table-cell surface and PptxGenJS 4.0.1 public `addTable()` output.
- Produces: real default/override conformance, accurate compatibility status, and packed Node/browser/declaration coverage.

- [ ] **Step 1: Add real PptxGenJS valign conformance**

Generate public tables covering omitted table/cell values, table-level top/middle/bottom, cell-level top/middle/bottom overrides, and runtime invalid `mid`/`distributed`. Import through `importPptxGenJS()` and assert valid snapshots, omitted→middle materialization, invalid→undefined, all text survives, valid direct `tcPr@anchor` counts match, and no cell `bodyPr@anchor` is created. Write/reopen and compare snapshots without accessing private fields.

- [ ] **Step 2: Update API, compatibility, npm README, and changelog**

Add the public snapshot/setter example and document physical indices, unique direct `tcPr`, strict `t/ctr/b`, clear semantics, separation from bodyPr anchor/fit/direction, and PptxGenJS omitted→middle materialization. Add this compatibility row:

```md
| table-cell `valign: top/middle/bottom` | `TableCell.verticalAlignment` / `TableModel.setCellVerticalAlignment()` | 已支持 direct 编辑 |
```

Do not claim table creation, table-level mutation, effective inheritance, or just/dist support.

- [ ] **Step 3: Extend actual-package Node/browser/declaration smoke**

Give the existing injected target/neighbor cells direct `anchor="ctr"` / `"b"`. In Node and browser conditions, verify middle snapshot, top/middle/bottom switching, clear, text/direction/fit preservation, and adjacent-cell isolation.

Compile this declaration fixture:

```ts
const cellAlignment: TextBoxVerticalAlignment = 'middle';
const table = createdDocument.slides[0].shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
const snapshotAlignment: TextBoxVerticalAlignment | undefined =
  table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, cellAlignment);
table?.setCellVerticalAlignment(0, 0, undefined);
```

- [ ] **Step 4: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: all functional tests pass, only default performance is skipped in the full run, and isolated performance passes.

### Task 4: Tarball, native/baseline validation, visual QA, and delivery

**Files:**
- Review all Task 1–3 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and all prior evidence.
- Produces: reviewed `feat: support table cell vertical alignment` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_alignment_package_dir=$(mktemp -d /tmp/pptx-table-cell-alignment-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_alignment_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_cell_alignment_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke all report true, including `tableCellVerticalAlignment`.

- [ ] **Step 2: Generate same-source native and hand-patched table files**

Use public PptxGenJS 4.0.1 to create one fixed five-cell table whose omitted valign is materialized as `ctr`. Open one copy through `PptxDocument` and apply top, middle, bottom, clear, and an untouched neighbor through `setCellVerticalAlignment()`; patch the same `tcPr@anchor` attributes by hand in another copy. Save explicit outputs under `/tmp/pptx-table-cell-alignment-native/native.pptx` and `/tmp/pptx-table-cell-alignment-baseline/baseline.pptx`.

Validate and compare:

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-alignment-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-alignment-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-cell-alignment-native/native.pptx /tmp/pptx-table-cell-alignment-baseline/baseline.pptx
```

Expected: both packages report zero errors/warnings and added/removed/changed diff arrays are empty.

- [ ] **Step 3: Render, inspect, and run overflow checks**

Export both through `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf`, rasterize every page with `pdftoppm -png`, and inspect each image at full size. Confirm native/baseline renders are pixel-identical, top/middle/bottom positions are visibly distinct, labels/text remain intact, and no repair, clipping, overlap, or neighboring-cell change occurs.

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-alignment-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-alignment-baseline/baseline.pptx
```

Expected: both overflow checks pass.

- [ ] **Step 4: Final review, commit, push, and verify remote state**

Run `git diff --check`, inspect the complete diff, and verify the shared three-value type, direct tcPr ownership, omitted PptxGen middle materialization, canonical/clear behavior, strict snapshots, physical indices, malformed/neighbor isolation, and no table-creation/effective-layout expansion. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, run `git diff --cached --check`, commit, and push through the verified SSH-over-443 channel:

```sh
git commit -m "feat: support table cell vertical alignment"
git rev-list --left-right --count origin/main...HEAD
```

Expected: remote count is `0 0`.
