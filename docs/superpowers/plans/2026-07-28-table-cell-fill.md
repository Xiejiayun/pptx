# Table Cell Fill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, lossless direct table-cell no-fill/solid-fill snapshots and physical-cell editing, including sRGB/theme colors and percentage transparency, with PptxGenJS 4.0.1 output conformance.

**Architecture:** Add a focused table-cell fill codec that owns only the selected direct `tcPr` fill choice and reuses the public `RichTextColor` value shape without coupling to rich-text storage. Expose detached `none | solid` snapshots through `TableCell.fill`; replace, insert, or remove one direct fill choice through `TableModel.setCellFill()`, preserving borders, text fill, all other cell properties, and neighboring content.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, source-span OOXML editing, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public type is `TableCellFill = { kind: 'none' } | { kind: 'solid'; color: RichTextColor; transparency?: number }`; do not expose raw XML, PptxGenJS runtime strings, or unstable effective-style values.
- `kind: 'none'` writes direct `a:noFill`; `undefined` removes the direct fill choice. These states must remain distinct.
- Solid colors support strict six-digit sRGB and the existing `RichTextColor` scheme token set. Native input may accept optional leading `#` for sRGB and must normalize it away.
- Transparency uses finite `0..100` percent, quantized with `Math.round((100 - value) * 1000)`; omission writes no alpha while explicit zero writes `alpha="100000"`.
- Getter ownership is only direct children of a unique direct `tcPr`; border/text descendant fills, table styles, inheritance, and effective colors are out of scope.
- Getter accepts only one unambiguous direct noFill or strict solidFill; unsupported/malformed/multiple fill states return `undefined` without mutation.
- Setter requires one direct `tcPr` and at most one direct fill choice. It may replace or clear one unsupported/malformed choice, but rejects multiple direct choices.
- Same supported semantic state is an exact no-op, including lowercase sRGB and leading-zero alpha tokens.
- Preserve margins, anchor, direction, fit, overflow, borders, cell3D, extensions, unknown XML, text body, text/run fill, merge state, relationships, neighbors, and live model identity.
- Do not add table creation, table-level defaults, advanced fill editing, border, merge, sizing, rich-cell-text, auto-page, or layout behavior.
- Implement inline without subagent delegation, as required for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, empty native/baseline package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Strict direct-fill snapshot and lossless model mutation

**Files:**
- Create: `packages/model/src/table-cell-fill.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `RichTextColor`, direct table-cell lookup, and nested OPC transactions.
- Produces: `TableCellFill`, `TableCell.fill`, `normalizeTableCellFill()`, `readTableCellFill()`, `replaceTableCellFill()`, and `TableModel.setCellFill()`.

- [ ] **Step 1: Add failing strict-read tests**

Replace one fixture table row with cells whose direct `tcPr` states include:

```xml
<a:tcPr><a:noFill/></a:tcPr>
<a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>
<a:tcPr><a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr>
<a:tcPr><a:solidFill><a:srgbClr val="00FF00"><a:alpha val="100000"/></a:srgbClr></a:solidFill></a:tcPr>
<a:tcPr><a:solidFill><a:srgbClr val="0000FF"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:tcPr>
<a:tcPr><a:solidFill><a:srgbClr val="abcdef"><a:alpha val="075000"/></a:srgbClr></a:solidFill></a:tcPr>
```

Add absent fill, border-only nested fills, `gradFill`, `pattFill`, two direct fill choices, invalid/short/non-hex sRGB, unsupported scheme token, empty/decimal/negative/over-range/repeated alpha, another transform beside alpha, namespaced `x:val` attributes, wrong-prefix `x:solidFill` / `x:alpha` elements, extra non-namespace attributes, malformed noFill, and repeated/missing direct `tcPr`.

Assert exact detached snapshots and zero mutation:

```ts
expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
  { kind: 'none' },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
  {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
  {
    kind: 'solid',
    color: { kind: 'srgb', value: '00FF00' },
    transparency: 0,
  },
  {
    kind: 'solid',
    color: { kind: 'srgb', value: '0000FF' },
    transparency: 100,
  },
  {
    kind: 'solid',
    color: { kind: 'srgb', value: 'ABCDEF' },
    transparency: 25,
  },
  undefined,
  undefined,
  undefined,
]);
expect(pkg.mutations).toEqual(journal);
```

Mutate both the returned fill object and its nested color through test-only casts; assert a fresh `table.rows` read remains source-backed.

- [ ] **Step 2: Run the model test and verify red**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new assertions fail because `TableCell.fill` and `setCellFill()` do not exist.

- [ ] **Step 3: Implement the public type and strict normalizer**

In `packages/model/src/shapes.ts`, import `RichTextColor` and define:

```ts
export type TableCellFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };
```

In `packages/model/src/table-cell-fill.internal.ts`, define the supported scheme set exactly as the existing rich-text public contract and normalize before mutation:

```ts
export function normalizeTableCellFill(
  value: unknown,
  context: string,
): TableCellFill | undefined {
  if (value === undefined) return undefined;
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be a fill object or undefined`);
  }
  const candidate = value as Record<string, unknown>;
  if (candidate.kind === 'none') {
    assertKeys(candidate, ['kind'], context);
    return { kind: 'none' };
  }
  if (candidate.kind !== 'solid') {
    throw new TypeError(`${context} kind must be none or solid`);
  }
  assertKeys(candidate, ['kind', 'color', 'transparency'], context);
  if (candidate.color === undefined) {
    throw new TypeError(`${context} solid fill must provide color`);
  }
  const color = normalizeColor(candidate.color, `${context} color`);
  const transparency = candidate.transparency === undefined
    ? undefined
    : normalizeTransparency(candidate.transparency, `${context} transparency`);
  return {
    kind: 'solid',
    color,
    ...(transparency !== undefined ? { transparency } : {}),
  };
}
```

`normalizeColor()` must accept only `{ kind: 'srgb', value: '#?RRGGBB' }` or `{ kind: 'scheme', value: supportedToken }`, reject extra keys, and uppercase normalized sRGB. `normalizeTransparency()` must reject non-number, non-finite, negative, and over-100 values, then return `(100_000 - Math.round((100 - value) * 1_000)) / 1_000` so snapshots and renders use the same 0.001% quantization.

Do not refactor `rich-text.internal.ts`; storage ownership remains independent even though the public color value shape matches.

- [ ] **Step 4: Implement strict direct snapshot parsing**

Use direct-child helpers and these fill choices:

```ts
const FILL_CHOICES = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

export function readTableCellFill(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TableCellFill | undefined;
```

Require exactly one direct `tcPr`, derive its lexical prefix (for example `a:`), and recognize fill/color/alpha elements only when they use that same prefix. Then require exactly one direct fill choice. Return `undefined` for no choice, wrong-prefix elements, multiple choices, or unsupported choice. For `noFill`, require no non-namespace attributes and no element children. For `solidFill`, require no non-namespace attributes and exactly one same-prefix direct `srgbClr` or `schemeClr`.

The color element must have exactly one non-namespace `val` attribute. Accept six hex digits or a supported scheme token. It may have zero element children or exactly one `alpha`; alpha must have exactly one non-namespace `val`, no element children, and strict integer `0..100000`. Return transparency only when alpha exists:

```ts
const transparency = 100 - rawAlpha / 1_000;
```

Reject other transforms, repeated alpha, namespaced substitutes, extra attributes, and malformed values as one opaque `undefined` snapshot. Never descend into `lnL/lnR/lnT/lnB/lnTlToBr/lnBlToTr`.

- [ ] **Step 5: Add failing edit, no-op, isolation, and rollback tests**

Build target cells with:

- lowercase sRGB plus `alpha val="075000"` for exact semantic no-op;
- self-closing `tcPr` for insertion;
- expanded noFill for replacement;
- one unsupported gradient for supported replacement;
- one direct solid fill for clear;
- one malformed but unique solid fill for canonical replacement;
- one merged physical placeholder;
- one adjacent cell with distinct borders, text fill, margins, direction, fit, alignment, and unknown XML;
- multiple-fill, repeated-`tcPr`, and missing-`tcPr` invalid targets.

Exercise:

```ts
table.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'srgb', value: '#ABCDEF' },
  transparency: 25,
});
table.setCellFill(0, 1, {
  kind: 'solid',
  color: { kind: 'srgb', value: 'FF0000' },
});
table.setCellFill(0, 2, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 33.333,
});
table.setCellFill(0, 3, { kind: 'none' });
table.setCellFill(0, 4, undefined);
```

Assert same-value exact bytes/journal, canonical sRGB/theme/alpha output, direct none versus clear, unsupported/malformed replacement, namespace-prefix reuse, insertion before `extLst`, self-closing expansion, merged physical-cell editing, and exact preservation of borders/text fill/other cell state/neighbor.

Run text, direction, fit, vertical-alignment, margins, and transform edits after fill mutation and prove the fills survive. Reject invalid indices and structures without mutation. Roll back two fill edits inside an outer transaction and compare exact bytes, journal, slide/table identity, and fresh snapshots.

- [ ] **Step 6: Implement direct replacement and public indexed mutation**

Implement:

```ts
export function replaceTableCellFill(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  fill: TableCellFill | undefined,
  partUri: string,
): boolean;
```

Parse only the selected direct `tcPr` source slice. Reject missing/repeated `tcPr` and multiple direct fill choices with `ModelParseError`. If no choice and `fill === undefined`, return false. If the one supported choice parses to the same normalized semantic state, return false without serialization. Otherwise render with the current `tcPr` prefix:

```ts
function renderFill(fill: TableCellFill, prefix: string): string {
  if (fill.kind === 'none') return `<${prefix}noFill/>`;
  const tag = fill.color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const color = fill.transparency === undefined
    ? `<${prefix}${tag} val="${fill.color.value}"/>`
    : `<${prefix}${tag} val="${fill.color.value}"><${prefix}alpha val="${Math.round(
        (100 - fill.transparency) * 1_000,
      )}"/></${prefix}${tag}>`;
  return `<${prefix}solidFill>${color}</${prefix}solidFill>`;
}
```

Replace/remove the existing choice as one source span. Insert new fill before same-prefix direct `extLst`, otherwise use `appendChildXml()` so self-closing `tcPr` expands safely. Replace only the selected `tcPr` in the parent XML.

Add `fill?: TableCellFill` in `TableCell`, spread only defined snapshots in `TableModel.rows`, and add:

```ts
setCellFill(
  rowIndex: number,
  columnIndex: number,
  value: TableCellFill | undefined,
): void {
  const fill = normalizeTableCellFill(value, 'Table cell fill');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    if (replaceTableCellFill(xml, cell, fill, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

- [ ] **Step 7: Re-run the model suite**

Expected: strict supported/opaque snapshots, all edit modes, exact no-op, prefix/order preservation, cross-property isolation, invalid input/index/structure, and rollback tests pass.

### Task 2: SDK lifecycle and invalid-input isolation

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public fill type/field/method, the existing combined table-cell fixture, duplicate, write/reopen, `PptxDocument.transaction()`, current cell property edits, and transforms.
- Produces: public lifecycle, nested detached-snapshot, and package-isolation evidence.

- [ ] **Step 1: Extend the combined table fixture with direct fills**

Keep current direction/alignment/fit/margins and give the five combined cells:

```ts
const originalFills = [
  { kind: 'solid', color: { kind: 'srgb', value: '4472C4' } },
  {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 25,
  },
  { kind: 'none' },
  undefined,
  {
    kind: 'solid',
    color: { kind: 'srgb', value: '70AD47' },
    transparency: 50,
  },
];
```

Use cell-level direct children after borders and before any `extLst`; do not place fills inside `bodyPr`, text `rPr`, or line elements.

- [ ] **Step 2: Cover edit, duplicate, rollback, and reopen lifecycles**

Open the fixture, retain slide/table identity, and duplicate before editing. Apply opaque sRGB, theme plus explicit zero transparency, explicit none, clear, and merged-cell full transparency. Then edit margins, fit, direction, alignment, text, and transform.

Mutate both outer and nested color objects in an old `rows` snapshot. Assert current versus duplicate values, all cross-properties, exact XML ownership, adjacent-cell isolation, and stable identity. Roll back fill changes inside `document.transaction()` and compare exact bytes/journal/live state. Write/reopen and compare edited versus duplicated rows.

- [ ] **Step 3: Reject invalid public values and coordinates before mutation**

Cover these invalid runtime values through casts:

```ts
const invalidValues = [
  null,
  false,
  true,
  '',
  [],
  {},
  { kind: 'none', color: { kind: 'srgb', value: 'FF0000' } },
  { kind: 'unknown' },
  { kind: 'solid' },
  { kind: 'solid', color: null },
  { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
  { kind: 'solid', color: { kind: 'srgb', value: 'GG0000' } },
  { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', extra: true } },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: -0.001 },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: 100.001 },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: Number.NaN },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: Number.POSITIVE_INFINITY },
  Symbol('table cell fill'),
];
```

Reuse negative/fractional/NaN/infinite/out-of-range coordinate coverage. Confirm bytes, journal, slide/shape counts, table identity, and text/direction/fit/alignment/margins/fills remain unchanged.

- [ ] **Step 4: Run typecheck and focused model/SDK suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck and both suites pass.

### Task 3: PptxGenJS conformance, docs, and packed public surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1-2 public surface and real PptxGenJS 4.0.1 `addTable()` output.
- Produces: actual materialization/transparency conformance, public documentation, Node/browser/declaration smoke, and package check `tableCellFill`.

- [ ] **Step 1: Add real PptxGenJS table-fill conformance**

Generate public one-cell tables for omitted fill, `{ type: 'none' }`, table-level sRGB, table-level SchemeColor with 25% transparency, cell override, explicit zero transparency, 33.333% transparency, 100% transparency, and deprecated `alpha: 25`. Add runtime invalid -1 and 101 transparency through a test-only cast while still calling only public `addTable()` / `write()`.

Import through `importPptxGenJS()` and assert:

- omitted and type-none snapshots are `undefined`;
- table-level values are materialized into direct cell fills;
- cell-level values override table-level values;
- sRGB/theme tokens and 25/33.333/100 transparency read exactly;
- PptxGenJS explicit zero collapses to absent alpha and therefore snapshot transparency is omitted;
- out-of-range direct alpha makes the strict fill snapshot `undefined` while text and XML remain preserved;
- direct fill counts exclude `ln*` border fills and text-run fills;
- write/reopen snapshots match; no private PptxGenJS fields are accessed.

- [ ] **Step 2: Update docs and compatibility status**

Add `TableCellFill` and `setCellFill()` examples to `docs/api/README.md`. Document direct none versus clear, sRGB/theme color, 0..100 percentage, explicit-zero alpha, strict unique `tcPr`/fill choice, physical indices, border/text separation, PptxGenJS materialization and invalid-runtime preservation.

Add this matrix row:

```md
| table-cell `fill` solid/none/transparency | `TableCell.fill` / `TableModel.setCellFill()` | 已支持 direct 读取、编辑与清除 |
```

Update npm README and changelog without claiming table creation, effective table styles, gradients/patterns/pictures, or table-level mutation.

- [ ] **Step 3: Extend actual-package Node/browser/declaration smoke**

Give the existing injected target and neighbor cells distinct direct fills in addition to their current properties. In Node and browser conditions, verify snapshot, sRGB/theme set, omitted/explicit-zero/fractional transparency, direct none, clear, text/direction/fit/alignment/margin preservation, and neighbor isolation.

Add a package check named `tableCellFill` and compile:

```ts
const cellFill: TableCellFill = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
};
const fillSnapshot: TableCellFill | undefined = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, cellFill);
table?.setCellFill(0, 0, { kind: 'none' });
table?.setCellFill(0, 0, undefined);
```

- [ ] **Step 4: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: all functional tests pass, only default performance is skipped in the full run, and isolated performance passes.

### Task 4: Tarball, native/baseline validation, visual QA, review, and delivery

**Files:**
- Review all Task 1-3 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation, built package, repository CLI, presentation overflow helper, LibreOffice, and same-source native/baseline files.
- Produces: reviewed `feat: support table cell fill` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual npm tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_fill_package_dir=$(mktemp -d /tmp/pptx-table-cell-fill-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_fill_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_cell_fill_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke all report true, including `tableCellFill`.

- [ ] **Step 2: Generate same-source native and baseline files**

Use public PptxGenJS 4.0.1 to create one fixed six-cell table with zero borders and no direct cell fill. Open one copy through `PptxDocument` and set:

1. opaque sRGB red;
2. scheme `accent1`;
3. sRGB blue at 50% transparency;
4. explicit direct noFill;
5. solid green and then clear it;
6. untouched neighbor.

Create the baseline from the exact same source by applying the expected direct `tcPr` fill children while preserving package timestamps and every unrelated part. Save:

- `/tmp/pptx-table-cell-fill-native/native.pptx`
- `/tmp/pptx-table-cell-fill-baseline/baseline.pptx`

Validate and compare:

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-fill-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-fill-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-cell-fill-native/native.pptx /tmp/pptx-table-cell-fill-baseline/baseline.pptx
```

Expected: both packages report zero errors/warnings and added/removed/changed arrays are empty.

- [ ] **Step 3: Render every output, inspect at full size, and run overflow checks**

Export both PPTX files to PDF with:

```sh
native_pdf_dir=$(mktemp -d /tmp/pptx-table-cell-fill-native-pdf.XXXXXX)
baseline_pdf_dir=$(mktemp -d /tmp/pptx-table-cell-fill-baseline-pdf.XXXXXX)
native_render_dir=$(mktemp -d /tmp/pptx-table-cell-fill-native-render.XXXXXX)
baseline_render_dir=$(mktemp -d /tmp/pptx-table-cell-fill-baseline-render.XXXXXX)
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir "$native_pdf_dir" /tmp/pptx-table-cell-fill-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir "$baseline_pdf_dir" /tmp/pptx-table-cell-fill-baseline/baseline.pptx
pdftoppm -png "$native_pdf_dir/native.pdf" "$native_render_dir/slide"
pdftoppm -png "$baseline_pdf_dir/baseline.pdf" "$baseline_render_dir/slide"
```

Inspect every page individually at full size. Confirm opaque sRGB, resolved scheme color, 50% transparency, explicit noFill, cleared fill, and neighbor are visibly correct; labels, text fill, borders, and geometry remain intact; no repair, clipping, overlap, blur, or unintended neighboring-cell difference occurs. Compare SHA-256 hashes of corresponding PNGs and require equality.

Run:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-fill-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-fill-baseline/baseline.pptx
```

Expected: both overflow checks pass and raster hashes match.

- [ ] **Step 4: Final review, explicit staging, commit, push, and remote verification**

Run `git diff --check` and inspect the complete diff. Verify direct-child ownership, strict color/alpha parsing, none versus clear, exact semantic no-op, PptxGenJS materialization/zero/runtime-invalid evidence, nested snapshot isolation, border/text/neighbor preservation, and absence of table-creation/effective-style claims.

Confirm status lists only the intended implementation files plus `.pnpm-store/`. Stage the explicit intended paths, run `git diff --cached --check`, and commit:

```sh
git commit -m "feat: support table cell fill"
```

Push through the verified SSH-over-443 channel, fetch, and run:

```sh
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; `.pnpm-store/` remains untracked and untouched.
