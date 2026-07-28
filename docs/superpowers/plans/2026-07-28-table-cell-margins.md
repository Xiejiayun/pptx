# Table Cell Margins Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, lossless point-based table-cell margin snapshots and physical-cell editing for existing PPTX tables, with PptxGenJS 4.0.1 output conformance.

**Architecture:** Reuse the existing public `TextBoxMargins` / `TextBoxMarginInput` value types and normalizer, but add a dedicated table-cell codec that owns only direct `tcPr@marL/marR/marT/marB`. Expose detached direct-state snapshots through `TableCell.margins` and replace the four managed attributes through `TableModel.setCellMargins()`, preserving text-body insets, cell formatting, other table-cell properties, and neighboring content.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public values use point and reuse `TextBoxMargins` / `TextBoxMarginInput`; do not add a PptxGenJS-style conditional inches/points native input.
- Public read surface is optional `TableCell.margins`; public write surface is `TableModel.setCellMargins(rowIndex, columnIndex, value)`.
- Storage is the selected physical cell's direct `a:tcPr@marL/marR/marT/marB`, not text-box `bodyPr@*Ins`, table-level defaults, styles, or effective values.
- Scalar writes four sides; tuple order is top/right/bottom/left; object input is a whole replacement and keeps only provided sides; `{}` / `undefined` clear all four managed direct attributes.
- Point values quantize through `Math.round(value * 12700)` and must fit signed Int32; strict reads accept only unique unqualified decimal-integer attributes in that range.
- A unique `tcPr` may return a partial snapshot when some sides are absent or malformed; missing/repeated `tcPr` returns `undefined`. Setter requires exactly one direct `tcPr` and no repeated managed unqualified attribute.
- A same numeric direct state is an exact no-op, including valid non-canonical integer spellings such as leading zeros; a single malformed value may be replaced or cleared.
- PptxGenJS 4.0.1 materializes omitted narrow defaults and table-level margin into each ordinary cell, and uses its legacy first-value `<1` inches / `>=1` points branch; adapter tests must assert actual emitted OOXML, not inferred input intent.
- Do not add table creation, table-level margin mutation, effective inheritance, row/column sizing, fill, border, merge mutation, rich text, auto paging, or dynamic layout.
- Preserve anchor, direction, fit, text-body insets, text, paragraphs/runs, fill/border/effects, merge state, neighboring cells, relationships, and `TableModel` identity.
- Implement inline without subagent delegation, as required for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, empty native/baseline package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Strict direct-margin snapshot and lossless model mutation

**Files:**
- Create: `packages/model/src/table-cell-margins.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `TextBoxMargins`, `TextBoxMarginInput`, `normalizeTextBoxMargins()`, `TableModel.rows`, and nested OPC transactions.
- Produces: `TableCell.margins`, `readTableCellMargins()`, `replaceTableCellMargins()`, and `TableModel.setCellMargins()`.

- [ ] **Step 1: Add failing strict-read tests**

Replace one fixture table row with cells containing complete, partial, zero, negative, fractional-point, minimum and maximum signed-Int32 raw margins. Add cells with absent, empty, decimal, scientific, whitespace, wrong-case, namespaced, duplicate, and out-of-range attributes; include text-body `*Ins` descendants plus repeated/missing direct `tcPr`.

Assert exact point snapshots without mutation:

```ts
expect(table.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
  { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
  { top: 0, left: -1 },
  { right: 1_588 / 12_700 },
  { top: -2_147_483_648 / 12_700, bottom: 2_147_483_647 / 12_700 },
  undefined,
  { right: 8 },
  undefined,
  undefined,
]);
expect(pkg.mutations).toEqual(journal);
```

Mutate a returned margin object through a test-only cast and prove a fresh `table.rows` read still reflects source XML.

- [ ] **Step 2: Run the model suite and verify red**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: new assertions fail because `TableCell.margins` and `setCellMargins()` do not exist.

- [ ] **Step 3: Implement normalization reuse and the dedicated codec**

Import the already-exported `normalizeTextBoxMargins()` from `text-box-margins.internal.ts` without changing that module. Define stable cell-side ownership:

```ts
const EMU_PER_POINT = 12_700;
const MIN_INT32 = -2_147_483_648;
const MAX_INT32 = 2_147_483_647;
const SIDES = [
  ['left', 'marL'],
  ['right', 'marR'],
  ['top', 'marT'],
  ['bottom', 'marB'],
] as const;

export function readTableCellMargins(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TextBoxMargins | undefined;

export function replaceTableCellMargins(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  margins: TextBoxMargins | undefined,
  partUri: string,
): boolean;
```

`readTableCellMargins()` must require one direct `tcPr`, then independently require exactly one unqualified side attribute matching `/^-?\d+$/` and signed Int32. `replaceTableCellMargins()` must parse only the selected `tcPr` source slice, reject repeated managed attributes, compare strict integer numeric values with normalized raw targets for a whole-operation no-op, and otherwise apply all four replacements/removals to that slice. Reparse between side updates so multiple inserts at the end of the same start tag remain non-overlapping and end in `marL`, `marR`, `marT`, `marB` order.

Use a private updater equivalent to:

```ts
function updateAttribute(template: string, name: string, value: string | undefined): string {
  const properties = LosslessXmlDocument.parse(template);
  const root = properties.roots[0];
  if (!root || root.localName !== 'tcPr') {
    throw new ModelParseError('Invalid table cell properties template');
  }
  const attributes = root.attributes.filter((attribute) => attribute.name === name);
  if (attributes.length > 1) {
    throw new ModelParseError(`Table cell contains repeated direct ${name} attributes`);
  }
  const attribute = attributes[0];
  if (value !== undefined) {
    if (attribute) properties.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? properties.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      properties.replace(insertionPoint, insertionPoint, ` ${name}="${value}"`);
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(properties.source[start - 1] ?? '')) start -= 1;
    properties.replace(start, attribute.end, '');
  }
  return properties.serialize();
}
```

In `TableCell`, add `readonly margins?: TextBoxMargins`; in `TableModel.rows`, spread only defined snapshots.

- [ ] **Step 4: Add failing edit, isolation, and rollback tests**

Build target cells with self-closing and expanded `tcPr`, mixed quote styles/order, complete/partial/unknown/repeated margins, anchor, direction, overflow, fill, borders, unknown attributes/children, text-body insets and fit, a neighboring cell, and a merged placeholder.

Cover:

```ts
table.setCellMargins(0, 0, 7.2);
table.setCellMargins(0, 1, [3.6, 7.2, 10.8, 14.4]);
table.setCellMargins(0, 2, { top: 4, left: 8 });
table.setCellMargins(0, 3, {});
table.setCellMargins(0, 4, undefined);
```

Assert scalar/TRBL/object mapping, partial replacement clears omitted managed sides, same numeric values preserve exact bytes, malformed single values canonicalize/clear, and selected physical merged cells can be edited. Run text/direction/fit/vertical-alignment/transform edits and prove margins survive. Invalid indices/structures must change no bytes. Roll back two margin edits in an outer transaction and compare bytes, journal, table identity, and fresh snapshots.

- [ ] **Step 5: Implement indexed validation and transactional mutation**

Add the public method:

```ts
setCellMargins(
  rowIndex: number,
  columnIndex: number,
  value: TextBoxMarginInput | undefined,
): void {
  const margins = normalizeTextBoxMargins(value, 'Table cell margins');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    if (replaceTableCellMargins(xml, cell, margins, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

Do not modify text-box margin serialization or its public `ShapeModel.textMargins` behavior.

- [ ] **Step 6: Re-run the model suite**

Expected: strict partial snapshots, all input forms, no-op, malformed structure, bodyPr/format isolation, merged cell, invalid input/index, and rollback tests pass.

### Task 2: SDK lifecycle and invalid-input isolation

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public field/method, the existing editable table fixture, duplicate, write/reopen, `PptxDocument.transaction()`, existing cell text/direction/fit/alignment edits, and transforms.
- Produces: public lifecycle, cross-property preservation, and package-isolation evidence without table creation APIs.

- [ ] **Step 1: Extend the table fixture with direct margins**

On the existing combined table-cell row, keep fit/direction/alignment metadata and add complete, partial, zero, absent, and merged-cell margins. Use canonical values that read as:

```ts
[
  { top: 4, right: 8, bottom: 12, left: 16 },
  { top: 0, left: 2 },
  { right: 7.2 },
  undefined,
  { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
]
```

- [ ] **Step 2: Cover the public edit lifecycle**

Open the fixture, retain slide/table identity, and duplicate before editing. Apply scalar, TRBL, partial object, clear, and merged-cell edits, then edit fit/direction/alignment/text and transform. Mutate a detached rows snapshot. Assert current versus duplicate margins and all existing properties, XML ownership, neighboring-cell isolation, and stable identity.

Roll back margin changes inside `document.transaction()` and assert exact bytes/journal/live state. Write/reopen and compare edited versus duplicated rows.

- [ ] **Step 3: Reject invalid values and coordinates before mutation**

Test `null`, booleans, strings, wrong-length arrays, arrays containing invalid values, unknown object keys, NaN, infinities, signed-Int32 overflow, and symbols. Test negative, fractional, NaN, infinite, and out-of-range coordinates. Confirm bytes, journal, slide/shape counts, table identity, text, direction, fit, alignment, and margins remain unchanged.

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
- Produces: real default/materialization/unit-branch conformance, accurate compatibility status, and packed Node/browser/declaration coverage.

- [ ] **Step 1: Add real PptxGenJS margin conformance**

Generate public tables with fresh cell objects for omitted table/cell margin, table scalar `0` and `0.1`, table tuple `[0.05, 0.1, 0.15, 0.2]`, cell overrides `0`, `0.25`, `[0.05, 0.1, 0.15, 0.2]`, scalar `1`, `[1, 2, 3, 4]`, and runtime negative input. Import through `importPptxGenJS()` and assert point snapshots plus exact direct raw attribute counts. Prove all margins live on `tcPr`, not `bodyPr`; all text survives; write/reopen snapshots match; and no private PptxGenJS field is used.

- [ ] **Step 2: Update API, compatibility, npm README, and changelog**

Add a public snapshot/setter example and document point units, scalar/TRBL/object input, whole-replacement/clear semantics, physical indices, unique direct `tcPr`, strict signed-Int32 reads, separation from bodyPr insets and other cell properties, and the PptxGenJS legacy dual-unit branch. Add this compatibility row:

```md
| table-cell `margin` scalar/TRBL | `TableCell.margins` / `TableModel.setCellMargins()` | 已支持 direct point snapshot 与编辑 |
```

Do not claim table creation, table-level mutation, style/effective default resolution, or layout recomputation.

- [ ] **Step 3: Extend actual-package Node/browser/declaration smoke**

Give the existing injected target/neighbor cells distinct direct margins. In Node and browser conditions, verify snapshot, scalar, TRBL, partial object, clear, text/direction/fit/alignment preservation, and adjacent-cell isolation. Add a package check named `tableCellMargins`.

Compile this declaration fixture:

```ts
const cellMargins: TextBoxMarginInput = { top: 4, left: 8 };
const snapshotMargins: TextBoxMargins | undefined = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, cellMargins);
table?.setCellMargins(0, 0, [3.6, 7.2, 10.8, 14.4]);
table?.setCellMargins(0, 0, undefined);
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
- Produces: reviewed `feat: support table cell margins` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_margins_package_dir=$(mktemp -d /tmp/pptx-table-cell-margins-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_margins_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_cell_margins_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke all report true, including `tableCellMargins`.

- [ ] **Step 2: Generate same-source native and PptxGenJS baseline files**

Use public PptxGenJS 4.0.1 to create one fixed five-cell table whose source cells all have explicit zero margins. Open one copy through `PptxDocument` and set zero, uniform 7.2pt, asymmetric `[3.6, 7.2, 10.8, 14.4]`, clear, and an untouched neighbor. Create the baseline from the same source by applying the exact expected `tcPr@mar*` values, preserving package timestamps and every unrelated part. Save outputs under `/tmp/pptx-table-cell-margins-native/native.pptx` and `/tmp/pptx-table-cell-margins-baseline/baseline.pptx`.

Validate and compare:

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-margins-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-margins-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-cell-margins-native/native.pptx /tmp/pptx-table-cell-margins-baseline/baseline.pptx
```

Expected: both packages report zero errors/warnings and added/removed/changed diff arrays are empty.

- [ ] **Step 3: Render, inspect, and run overflow checks**

Export both through `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf`, rasterize every page with `pdftoppm -png`, and inspect each image at full size. Confirm native/baseline renders are pixel-identical; zero, uniform, asymmetric, clear/default, and untouched-neighbor layouts are visibly correct; labels/text remain intact; and no repair, clipping, overlap, or neighboring-cell change occurs.

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-margins-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-margins-baseline/baseline.pptx
```

Expected: both overflow checks pass and raster hashes match.

- [ ] **Step 4: Final review, commit, push, and verify remote state**

Run `git diff --check`, inspect the complete diff, and verify point units, direct tcPr ownership, omitted PptxGen narrow defaults, table-level materialization, dual-unit legacy evidence, whole-replacement/clear behavior, partial strict snapshots, physical indices, malformed/neighbor isolation, and no table-creation/effective-layout expansion. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, run `git diff --cached --check`, commit, and push through the verified SSH-over-443 channel:

```sh
git commit -m "feat: support table cell margins"
git rev-list --left-right --count origin/main...HEAD
```

Expected: remote count is `0 0`.
