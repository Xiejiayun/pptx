# Table Cell Fit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, lossless table-cell text-fit snapshots and indexed editing for existing PPTX cells while documenting that PptxGenJS 4.0.1 has no corresponding table fit API.

**Architecture:** Add a table-cell boundary wrapper that requires a unique direct `txBody/bodyPr` and delegates the actual autofit choice read/patch to the existing text-box fit codec. Expose the result through immutable `TableCell.textFit` snapshots and mutate it through `TableModel.setCellTextFit()`, preserving all non-fit cell state and avoiding table creation or dynamic layout calculations.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Reuse the existing `TextBoxFit` values exactly: `none | shrink | resize`; do not add a duplicate public union.
- Public read surface is optional `TableCell.textFit`; public write surface is `TableModel.setCellTextFit(rowIndex, columnIndex, value)`.
- Storage is the selected cell's direct `a:txBody/a:bodyPr` fit choice, not `a:tcPr`, table defaults, text length, or shape transform.
- Getter requires one direct `txBody`, one direct `bodyPr`, and exactly one supported same-prefix direct child; malformed, absent, descendant, namespaced, or ambiguous state returns `undefined`.
- Setter `none` and `undefined` both remove all direct supported fit children; it never writes `a:noAutofit`.
- Setter shrink/resize writes one direct `a:normAutofit` / `a:spAutoFit`; same-mode assignment preserves calculated attributes byte-for-byte.
- Setter requires a unique direct `txBody/bodyPr` and throws `ModelParseError` rather than creating guessed structure.
- PptxGenJS 4.0.1 `TableCellProps` / `TableProps` expose no fit API; runtime `fit`, `autoFit`, and `shrinkText` passthrough is ignored and must import as `undefined`.
- Do not add table creation, table-level fit defaults, dynamic font scaling, shape resizing, cell overflow, vertical alignment, margins, fill, border, merge mutation, rich text, or row/column mutation.
- Preserve `tcPr`, direction, text body metadata, paragraphs/runs, merge state, neighboring cells, relationships, and `TableModel` identity.
- Implement inline without subagent delegation, as required for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, empty native/baseline package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Strict cell boundary, snapshot read, and lossless model mutation

**Files:**
- Create: `packages/model/src/table-cell-text-fit.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `TextBoxFit`, `normalizeTextBoxFit()`, `readTextBoxFit()`, `replaceTextBoxFit()`, `TableModel.rows`, and nested OPC transactions.
- Produces: `TableCell.textFit`, `readTableCellTextFit()`, `replaceTableCellTextFit()`, and `TableModel.setCellTextFit()`.

- [ ] **Step 1: Add failing strict-read tests**

Replace the fixture table row with physical cells whose unique direct bodyPr contains sole `noAutofit`, calculated `normAutofit`, or `spAutoFit`, followed by absent, duplicate same choice, mixed choice, case variant, different-prefix choice, descendant choice, repeated `txBody`, repeated `bodyPr`, missing `bodyPr`, and missing `txBody`. Keep valid text/direction metadata on malformed cells. Assert:

```ts
const fits = table.rows[0]!.cells.map(({ textFit }) => textFit);
expect(fits.slice(0, 3)).toEqual(['none', 'shrink', 'resize']);
expect(fits.slice(3)).toEqual(Array(10).fill(undefined));
expect(pkg.mutations).toEqual(journal);
```

Mutate a returned snapshot through a test-only cast and prove a fresh `table.rows` read still returns the source fit.

- [ ] **Step 2: Run the model suite and verify red**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new snapshot assertions fail because `TableCell.textFit` does not exist.

- [ ] **Step 3: Add the narrow table-cell wrapper and snapshot field**

Create the internal wrapper with unique direct-child checks:

```ts
export function readTableCellTextFit(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  partUri: string,
): TextBoxFit | undefined {
  if (!uniqueTextBodyProperties(cell)) return undefined;
  return readTextBoxFit(xml, cell, partUri);
}

export function replaceTableCellTextFit(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  value: TextBoxFit | undefined,
  partUri: string,
): boolean {
  if (!uniqueTextBodyProperties(cell)) {
    throw new ModelParseError(
      'Table cell must contain one direct text body with one body properties element',
      partUri,
    );
  }
  replaceTextBoxFit(xml, cell, value, partUri);
  return xml.changed;
}
```

`uniqueTextBodyProperties()` must count exact direct `txBody` and exact direct `bodyPr` by local name; do not inspect descendants. In `TableCell`, add `readonly textFit?: TextBoxFit`. In `TableModel.rows`, read once and conditionally spread only a defined property so unrelated object-shape expectations stay unchanged.

- [ ] **Step 4: Add failing edit, no-op, isolation, and rollback tests**

Build cells with:

- a calculated single-quoted `normAutofit` plus wrap/margin/anchor/vert/warp/3D/unknown XML;
- a self-closing `bodyPr`;
- an expanded fit-less `bodyPr` with `prstTxWarp`, `scene3d`, and `extLst` ordering anchors;
- mixed supported choices;
- an adjacent `spAutoFit` cell with unique metadata;
- a merged placeholder with explicit `noAutofit`;
- repeated/missing text-body structures.

Cover same-mode shrink no-op, shrink→resize, add shrink to self-closing, schema-ordered resize insertion, conflict→shrink normalization, explicit noAutofit→none removal, absent clear no-op, and clear with `undefined`. Assert exact target snippets, calculated metadata preservation, neighboring cell bytes, `tcPr@vert`, text, and unknown XML.

Run `setCellText()`, `setCellTextDirection()`, and transform edits and prove fit survives. Invalid indices/structures must change no bytes. Roll back two fit edits inside an outer transaction and compare bytes, journal, table identity, and fresh fit snapshots.

- [ ] **Step 5: Implement indexed validation and transactional mutation**

Add the method:

```ts
setCellTextFit(
  rowIndex: number,
  columnIndex: number,
  value: TextBoxFit | undefined,
): void {
  const fit = value === undefined
    ? undefined
    : normalizeTextBoxFit(value, 'Table cell text fit');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    if (replaceTableCellTextFit(xml, cell, fit, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

Use the existing text-box fit codec without changing its public behavior or duplicating its child-order algorithm.

- [ ] **Step 6: Re-run the model suite**

Expected: strict snapshots, lossless edits, same-mode/absent no-ops, conflict normalization, malformed structure, non-fit isolation, and rollback tests pass.

### Task 2: SDK lifecycle and invalid-input isolation

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public field/method, the existing editable table fixture, duplicate, write/reopen, `PptxDocument.transaction()`, cell text/direction edits, and transforms.
- Produces: public lifecycle and package-isolation evidence without table creation APIs.

- [ ] **Step 1: Extend the dedicated table fixture with fit bodies**

Let its cell helper accept a bodyPr template while preserving the current direction rows. Add a fit row with sole no/norm/sp, absent, and merged cells; retain unknown bodyPr and tcPr metadata. Keep the fixture independent from title/create shape-index tests.

- [ ] **Step 2: Cover public edit lifecycle**

Open the fixture, retain slide/table identity, and duplicate before editing. Read `none/shrink/resize/undefined`; set shrink/resize across self-closing and expanded bodies, clear noAutofit through none, clear another through undefined, edit cell text/direction and table transform, and mutate a detached rows snapshot. Assert current versus duplicated fits, text, direction, merged placeholder, unknown XML, neighbor isolation, and stable identity.

Roll back a fit change inside `document.transaction()` and assert exact bytes/journal/live state. Write/reopen and compare edited versus duplicated rows.

- [ ] **Step 3: Reject invalid values and coordinates before mutation**

For `null`, booleans, numbers, empty/case/whitespace strings, unknown strings, objects, arrays, and symbols, assert `TypeError`. For negative, fractional, NaN, infinite, and out-of-range coordinates, assert `RangeError`. Confirm bytes, journal, slide/shape counts, table identity, text, directions, and fits remain unchanged.

- [ ] **Step 4: Run typecheck and focused suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck and both suites pass.

### Task 3: PptxGenJS ignored-output conformance, docs, and packed surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public table-cell surface and PptxGenJS 4.0.1 public `addTable()` output.
- Produces: evidence that unsupported PptxGen table fit inputs are ignored, accurate compatibility status, and packed Node/browser/declaration coverage.

- [ ] **Step 1: Add real PptxGenJS ignored-runtime conformance**

Generate a public table whose outer options contain runtime `fit: 'resize'`, `autoFit: true`, and `shrinkText: true`; cells cover omitted, none, shrink, resize, both legacy flags, and all-conflict runtime objects. Import through `importPptxGenJS()` and assert every `textFit` is `undefined`, every raw `bodyPr` is fit-less, all cell text/directions survive, write/reopen is unchanged, and no private field is accessed.

- [ ] **Step 2: Update API, compatibility, npm README, and changelog**

Add a `TableModel` example for direct fit snapshots, shrink/resize, none, and clear. Document physical indices, unique direct `txBody/bodyPr` strict reads, same-mode calculated metadata, none/undefined clear behavior, lack of dynamic scaling, and separation from `tcPr@vert`.

Replace the remaining compatibility row with:

```md
| table-cell bodyPr autofit | `TableCell.textFit` / `TableModel.setCellTextFit()` | 原生编辑已支持；PptxGenJS 4.0.1 本身无 table fit API |
```

Add one precise changelog bullet and npm README capability note. Do not claim table creation, table-level fit, PptxGenJS creation parity, or dynamic layout.

- [ ] **Step 3: Extend actual-package Node/browser/declaration smoke**

Add direct `bodyPr` to the existing injected table cells. In Node and browser conditions, verify noAutofit snapshot, same-mode shrink, shrink/resize switching, none/undefined clear, text/direction preservation, and adjacent-cell isolation through the packed `TableModel`.

Compile this declaration fixture:

```ts
const cellFit: TextBoxFit = 'shrink';
const table = createdDocument.slides[0].shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
const snapshotFit: TextBoxFit | undefined = table?.rows[0]?.cells[0]?.textFit;
table?.setCellTextFit(0, 0, cellFit);
table?.setCellTextFit(0, 0, 'none');
table?.setCellTextFit(0, 0, undefined);
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
- Produces: reviewed `feat: support table cell fit` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_fit_package_dir=$(mktemp -d /tmp/pptx-table-cell-fit-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_fit_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_cell_fit_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke all report true, including `tableCellTextFit`.

- [ ] **Step 2: Generate same-source native and hand-patched table files**

Use public PptxGenJS 4.0.1 to create one fixed table with fit-less bodyPr cells. Open one copy through `PptxDocument` and apply none, shrink, resize, and clear with `setCellTextFit()`; patch the same bodyPr choice children by hand in another copy. Save explicit outputs under `/tmp/pptx-table-cell-fit-native/native.pptx` and `/tmp/pptx-table-cell-fit-baseline/baseline.pptx`.

Validate and compare:

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-fit-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-fit-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-cell-fit-native/native.pptx /tmp/pptx-table-cell-fit-baseline/baseline.pptx
```

Expected: both packages report zero errors/warnings and added/removed/changed diff arrays are empty.

- [ ] **Step 3: Render, inspect, and run overflow checks**

Export both through `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf`, rasterize every page with `pdftoppm -png`, and inspect each image at full size. Confirm native/baseline renders are pixel-identical, labels/text remain intact, and no repair, clipping, overlap, or neighboring-cell change occurs. Record that final shrink factors/shape sizes may be recalculated only by PowerPoint after edit/resize.

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-fit-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-fit-baseline/baseline.pptx
```

Expected: both overflow checks pass.

- [ ] **Step 4: Final review, commit, push, and verify remote state**

Run `git diff --check`, inspect the complete diff, and verify the three-value reuse, direct bodyPr ownership, none/clear behavior, strict snapshots, calculated metadata preservation, physical indices, malformed/neighbor isolation, PptxGenJS ignored behavior, and no table-creation/dynamic-layout expansion. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, run `git diff --cached --check`, commit, and push through the verified SSH-over-443 channel:

```sh
git commit -m "feat: support table cell fit"
git rev-list --left-right --count origin/main...HEAD
```

Expected: remote count is `0 0`.
