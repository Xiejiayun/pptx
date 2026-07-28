# Table Cell Borders Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, lossless direct table-cell left/right/top/bottom border snapshots and physical-cell editing, including explicit none, solid/dash style, point width, sRGB/theme color, and PptxGenJS 4.0.1 output conformance.

**Architecture:** Add a focused table-cell borders codec that owns only same-prefix direct `tcPr` children `lnL`, `lnR`, `lnT`, and `lnB`. Expose detached partial snapshots through `TableCell.borders`; normalize scalar/TRBL/named whole-replacement input before mutation; replace, insert, or remove only the four managed line elements while preserving diagonals, fill, text, all unrelated cell properties, and neighboring content.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, source-span OOXML editing, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public side value is `TableCellBorder = { kind: 'none' } | { kind: 'line'; color: RichTextColor; width: number; style?: 'solid' | 'dash' }`.
- Public setter accepts one scalar border for all sides, one exact `[top, right, bottom, left]` tuple, one partial named whole-replacement object, or `undefined` to clear all four sides.
- `kind: 'none'` writes a direct zero-width line with `noFill`; an omitted/undefined side removes the direct line element. These states must remain distinct.
- Width uses finite `0..1584` points, quantized to the nearest EMU (`12_700` EMU/point); a zero-width line remains a line value.
- Solid colors support strict six-digit sRGB and the existing `RichTextColor` scheme token set. Native input may accept optional leading `#` for sRGB and must normalize it away.
- Style omission writes no `prstDash`; `solid` writes `prstDash val="solid"`; `dash` writes `prstDash val="sysDash"`.
- Getter ownership is only same-prefix direct `lnL/lnR/lnT/lnB` children of a unique direct `tcPr`; diagonals, cell fill, text outline, table styles, inheritance, shared-edge resolution, and effective colors are out of scope.
- Getter parses each side independently. Unsupported/malformed/missing/repeated sides are omitted while other valid sides remain readable; no valid side returns `undefined`.
- Setter requires one direct `tcPr` and at most one same-prefix managed element per side. It may replace or clear one unsupported/malformed side, but rejects repeated managed elements.
- Same supported normalized state is an exact no-op, including lowercase sRGB, leading-zero width, omitted versus explicit dash style, and neutral PptxGenJS metadata.
- Preserve fill, margins, anchor, direction, fit, overflow, diagonal borders, cell3D, extensions, unknown XML, text body, text/run style, merge state, relationships, neighbors, and live model identity.
- Do not add table creation, table-level defaults, diagonal editing, other dash presets, compound/cap/join/alignment/arrow/effect editing, transparency, merge/sizing/rich-cell-text/auto-page/layout behavior, or a cross-feature line-engine refactor.
- Implement inline without subagent delegation, as required for this repository session.
- Never stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, empty native/baseline package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Strict direct-border snapshots and lossless model mutation

**Files:**
- Create: `packages/model/src/table-cell-borders.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `RichTextColor`, direct table-cell lookup, source-span replacement, and nested OPC transactions.
- Produces: `TableCellBorderStyle`, `TableCellBorder`, `TableCellBorders`, `TableCellBorderInput`, `TableCell.borders`, `normalizeTableCellBorders()`, `readTableCellBorders()`, `replaceTableCellBorders()`, and `TableModel.setCellBorders()`.

- [ ] **Step 1: Add failing strict-read tests**

Replace one model fixture table row with cells whose unique direct `tcPr` contains these supported forms:

```xml
<a:tcPr>
  <a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>
  <a:lnR w="12700"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:lnR>
  <a:lnT w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnT>
  <a:lnB w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="sysDash"/></a:lnB>
</a:tcPr>
```

Add separate cells for partial valid sides, `w="00012700"`, `w="0"` solid line, `w="20116800"`, and theme colors. Assert detached snapshots:

```ts
expect(table.rows[0]!.cells[0]!.borders).toEqual({
  top: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 1.5,
    style: 'solid',
  },
  right: {
    kind: 'line',
    color: { kind: 'srgb', value: 'FF0000' },
    width: 1,
  },
  bottom: {
    kind: 'line',
    color: { kind: 'srgb', value: '0000FF' },
    width: 2,
    style: 'dash',
  },
  left: { kind: 'none' },
});
expect(pkg.mutations).toEqual(journal);
```

Add one invalid side per case while keeping another valid side: missing/repeated/namespaced `w`, decimal/negative/over-range width, wrong-prefix line, repeated same side, `gradFill`, two fill choices, invalid/short sRGB, unsupported scheme, alpha transform, unsupported `prstDash`, repeated dash, bevel/miter, non-neutral arrow, extra/duplicate attributes, unknown child, malformed noFill, diagonal-only lines, and repeated/missing direct `tcPr`. Assert only valid sides survive and no read mutation occurs.

Mutate the returned borders object, one nested side, and its nested color through test-only casts; assert a fresh `table.rows` read remains source-backed.

- [ ] **Step 2: Run the model test and verify red**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new assertions fail because `TableCell.borders` and `setCellBorders()` do not exist.

- [ ] **Step 3: Define the public border value and input types**

In `packages/model/src/shapes.ts`, add exactly:

```ts
export type TableCellBorderStyle = 'solid' | 'dash';

export type TableCellBorder =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly width: number;
      readonly style?: TableCellBorderStyle;
    };

export interface TableCellBorders {
  readonly top?: TableCellBorder;
  readonly right?: TableCellBorder;
  readonly bottom?: TableCellBorder;
  readonly left?: TableCellBorder;
}

export type TableCellBorderInput =
  | TableCellBorder
  | readonly [
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
    ]
  | TableCellBorders;
```

Add `readonly borders?: TableCellBorders` to `TableCell`; do not add table creation options.

- [ ] **Step 4: Implement complete pre-mutation normalization**

In `packages/model/src/table-cell-borders.internal.ts`, define:

```ts
export function normalizeTableCellBorders(
  value: unknown,
  context: string,
): TableCellBorders | undefined;
```

Normalize with these exact rules:

```ts
const PUBLIC_SIDES = ['top', 'right', 'bottom', 'left'] as const;
const EMU_PER_POINT = 12_700;
const MAX_LINE_WIDTH_EMU = 20_116_800;
```

- a scalar object with `kind` is normalized once and cloned to all four sides;
- an array must have length 4, own every numeric slot, and contain only strict border objects or explicit `undefined`;
- a named object may contain only top/right/bottom/left and is a whole replacement; an empty result returns `undefined`;
- `none` allows only `kind`;
- `line` allows only `kind`, `color`, `width`, and `style`, requires color/width, accepts optional `#` before six hex digits, accepts only the established scheme set, and rejects extra nested color keys;
- width must be a finite number in `0..1584` before rounding and must round into `0..20_116_800` EMU; return `widthEmu / 12_700` so read/write use the same quantization;
- style must be absent, `solid`, or `dash`.

Reject symbols and extra keys with `Reflect.ownKeys()`. Do not refactor rich-text color storage; only the public value shape is shared.

- [ ] **Step 5: Implement independent strict side parsing**

Define:

```ts
export function readTableCellBorders(
  xml: LosslessXmlDocument,
  cell: XmlElement,
): TableCellBorders | undefined;
```

Require exactly one direct `tcPr`, derive its lexical prefix, and inspect same-prefix direct elements through these mappings:

```ts
const XML_SIDES = [
  ['left', 'lnL'],
  ['right', 'lnR'],
  ['top', 'lnT'],
  ['bottom', 'lnB'],
] as const;
```

For each side independently:

1. require exactly one unqualified strict integer `w` and allow only optional `cap="flat"`, `cmpd="sng"`, and `algn="ctr"` attributes;
2. require `w` in `0..20_116_800`;
3. if the sole same-prefix element child is empty `noFill`, return `{ kind: 'none' }`;
4. otherwise require one strict `solidFill` with one six-hex `srgbClr` or supported `schemeClr`, no transforms, and no extra attributes;
5. allow zero or one empty `prstDash`, mapping absent to omitted style, `solid` to `solid`, and `sysDash` to `dash`;
6. allow zero or one empty `round`; allow zero or one `headEnd` and `tailEnd` only when each has exactly `type="none"`, `w="med"`, `len="med"` and no element children;
7. reject all other/duplicate/wrong-prefix children for that side.

Build the public snapshot in top/right/bottom/left key order. Return `undefined` when no side parsed. Never inspect `lnTlToBr`, `lnBlToTr`, cell fill choices, or text descendants.

- [ ] **Step 6: Add failing edit, no-op, ordering, isolation, and rollback tests**

Build target cells with:

- lowercase sRGB, leading-zero width and PptxGenJS neutral children for exact same-value no-op;
- self-closing `tcPr` and partially populated L/R/T/B children for ordered insertion;
- four noFill sides for scalar replacement;
- supported lines for named/tuple clear;
- one unsupported dash, one malformed line, and one gradient line fill for whole-side replacement;
- direct `lnTlToBr` and `lnBlToTr`, cell solid fill, `cell3D`, `extLst`, unknown attributes/children, and distinct text-body/run state;
- one merged physical placeholder and one byte-identical adjacent cell;
- repeated `lnL`, repeated `tcPr`, and missing `tcPr` invalid targets.

Exercise:

```ts
table.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: '#ABCDEF' },
  width: 1,
  style: 'solid',
});
table.setCellBorders(0, 1, [
  { kind: 'line', color: { kind: 'scheme', value: 'accent2' }, width: 1.5, style: 'dash' },
  { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 },
  { kind: 'none' },
  undefined,
]);
table.setCellBorders(0, 2, {
  top: { kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 2 },
  left: { kind: 'none' },
});
table.setCellBorders(0, 3, {});
table.setCellBorders(0, 4, undefined);
```

Assert scalar replication, TRBL input, partial named whole replacement, none/clear distinction, 0/fractional/max widths, style omitted/solid/dash, canonical line XML, L/R/T/B schema order, insertion before diagonals/fill/cell3D/extLst, namespace-prefix reuse, self-closing expansion, merged physical-cell editing, and exact preservation of diagonals/fill/text/other cell state/neighbor.

Run text, direction, fit, vertical-alignment, margins, fill, and transform edits after border mutation and prove borders survive. Reject invalid coordinates, wrong-length/sparse arrays, unknown keys/kinds/styles, malformed colors, missing width/color, non-finite/negative/over-range width, repeated managed elements, and invalid `tcPr` without mutation. Roll back two border edits inside an outer transaction and compare exact bytes, journal, slide/table identity, and fresh snapshots.

- [ ] **Step 7: Implement lossless whole-replacement mutation**

Define:

```ts
export function replaceTableCellBorders(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  borders: TableCellBorders | undefined,
  partUri: string,
): boolean;
```

Parse only the selected `tcPr` source slice. Reject missing/repeated direct `tcPr` and repeated same-prefix managed elements with `ModelParseError`. Before editing, compare every current supported side with the normalized desired side; when all four match, return false without serialization.

Render canonical sides with the current `tcPr` prefix:

```ts
function renderBorder(
  tag: 'lnL' | 'lnR' | 'lnT' | 'lnB',
  border: TableCellBorder,
  prefix: string,
): string {
  const open = `<${prefix}${tag} w="${
    border.kind === 'none' ? 0 : Math.round(border.width * 12_700)
  }" cap="flat" cmpd="sng" algn="ctr">`;
  if (border.kind === 'none') {
    return `${open}<${prefix}noFill/></${prefix}${tag}>`;
  }
  const colorTag = border.color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  const dash = border.style === undefined
    ? ''
    : `<${prefix}prstDash val="${border.style === 'dash' ? 'sysDash' : 'solid'}"/>`;
  return `${open}<${prefix}solidFill><${prefix}${colorTag} val="${
    escapeXmlAttribute(border.color.value)
  }"/></${prefix}solidFill>${dash}<${prefix}round/><${prefix}headEnd type="none" w="med" len="med"/><${prefix}tailEnd type="none" w="med" len="med"/></${prefix}${tag}>`;
}
```

Update sides one at a time from L/R/T/B against a freshly parsed current `tcPr` string. Replace/delete an existing managed element in place. For a missing desired side, insert before the first later same-prefix schema child using this order:

```ts
lnL, lnR, lnT, lnB, lnTlToBr, lnBlToTr,
noFill | solidFill | gradFill | blipFill | pattFill | grpFill,
cell3D, extLst
```

If no later anchor exists, use `appendChildXml()` so a self-closing `tcPr` expands safely. Preserve wrong-prefix lines and every non-managed source span.

In `TableModel.rows`, spread `borders` only when defined. Add:

```ts
setCellBorders(
  rowIndex: number,
  columnIndex: number,
  value: TableCellBorderInput | undefined,
): void {
  const borders = normalizeTableCellBorders(value, 'Table cell borders');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    if (replaceTableCellBorders(xml, cell, borders, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

- [ ] **Step 8: Run focused model and type checks**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
pnpm typecheck
```

Expected: all model tests and TypeScript project references pass; no unrelated snapshot changes occur.

---

### Task 2: Public SDK isolation, transaction, and round-trip coverage

**Files:**
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: exported model border types, `PptxDocument.open()`, `PptxDocument.write()`, duplicate slides, public transactions, stable wrapper identity, and existing table fixture.
- Produces: end-to-end proof that public border snapshots and setters remain isolated, atomic, and round-trip stable.

- [ ] **Step 1: Extend the existing table fixture with PptxGenJS-compatible borders**

Give the third fixture row these states without removing existing direction/fit/alignment/margin/fill coverage:

1. all four red 1pt solid;
2. top theme dash, right green zero-width line, bottom none, left absent;
3. four none;
4. no managed sides;
5. distinct borders on a merged physical placeholder.

Keep one duplicated slide so edits can prove clone-on-write isolation.

- [ ] **Step 2: Add public edit and detached-snapshot assertions**

Read `originalBorders`, mutate one returned side and its nested color through test casts, and prove fresh snapshots are unchanged. Apply scalar, TRBL, named partial, `{}`, and `undefined` edits while also changing margins, fill, text, direction, fit, alignment, and transform.

Assert:

```ts
expect(document.slides[0]).toBe(slide);
expect(slide.shapes[0]).toBe(table);
expect(table.rows[2]!.cells.map(({ borders }) => borders)).toEqual(editedBorders);
expect(duplicateTable.rows[2]!.cells.map(({ borders }) => borders)).toEqual(originalBorders);
```

Inspect source XML for canonical L/R/T/B output and exact survival of diagonal/fill/body state. Write/reopen and assert edited/duplicate snapshots again.

- [ ] **Step 3: Add invalid-input and rollback coverage**

Reject before mutation:

```ts
const invalidValues = [
  null,
  false,
  '',
  [],
  [undefined, undefined, undefined],
  { top: undefined, extra: true },
  { kind: 'none', width: 1 },
  { kind: 'line' },
  { kind: 'line', color: null, width: 1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FFF' }, width: 1 },
  { kind: 'line', color: { kind: 'scheme', value: 'unknown' }, width: 1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: -0.001 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1584.001 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: Number.NaN },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1, style: 'dot' },
  Symbol('table cell borders'),
];
```

Also reject negative/fractional/non-finite/out-of-range physical coordinates. Compare exact bytes, mutation journal, slides/shapes identity, borders, fill, margins, alignment, fit, direction, and text before/after. Roll back two valid border edits inside `document.transaction()` and repeat the same checks.

- [ ] **Step 4: Run model and SDK tests together**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: both suites pass with stable identity and no cross-slide leakage.

---

### Task 3: PptxGenJS 4.0.1 public-output conformance

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addTable()` / `write()`, typed `BorderProps`, `importPptxGenJS()`, and `TableModel.borders`.
- Produces: evidence for materialized defaults, scalar/TRBL/override semantics, dash mapping, point quirks, and strict malformed-output handling.

- [ ] **Step 1: Extend the local PptxGenJS test interface only with public border types**

Represent `border` as:

```ts
type BorderProps = {
  type?: 'none' | 'dash' | 'solid';
  color?: string;
  pt?: number;
};

border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps];
```

Do not access `_slides` or other private runtime fields.

- [ ] **Step 2: Generate representative tables and assert imported snapshots**

Use public PptxGenJS 4.0.1 to create:

- omitted border;
- table scalar red 2pt solid;
- table TRBL with none/dash/solid/default values;
- cell scalar override over a distinct table border;
- cell tuple with explicit missing values supplied through a test-only runtime cast;
- cell `pt: 0` solid;
- table tuple `pt: 0` to capture its 1pt fallback;
- default type/color/point values;
- negative and over-range runtime point values that serialize but must not become valid strict snapshots.

Assert omitted border imports as four `{ kind: 'none' }`; scalar values materialize to four sides; tuple order is top/right/bottom/left; dash maps from `sysDash`; defaults are `666666`/1pt/solid; cell zero width remains a line; table tuple zero becomes 1pt; malformed widths return `undefined` only for affected sides.

- [ ] **Step 3: Assert actual XML evidence and reopen stability**

Decode the imported slide part and assert:

```xml
<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>
<a:prstDash val="solid"/>
<a:prstDash val="sysDash"/>
```

Verify each table writes managed children in L/R/T/B order even though tuple input is TRBL. Write/reopen through `PptxDocument` and compare all border snapshots exactly.

- [ ] **Step 4: Run adapter conformance**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
```

Expected: all adapter tests pass against installed PptxGenJS 4.0.1 public output.

---

### Task 4: Documentation, compatibility matrix, and packed-package surfaces

**Files:**
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `CHANGELOG.md`
- Modify: `scripts/smoke-npm-package.mjs`
- Test: actual npm tarball Node/browser/declaration/CLI smoke

**Interfaces:**
- Consumes: public border types/setter and validated PptxGenJS 4.0.1 behavior.
- Produces: accurate user documentation and evidence that the aggregate package exposes the feature in every supported runtime/type entry point.

- [ ] **Step 1: Extend Node and browser package smoke**

Add direct PptxGenJS-compatible four-side lines to the existing injected table XML. Read initial snapshots, then exercise:

```js
table.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: '#FF0000' },
  width: 2,
  style: 'solid',
});
table.setCellBorders(0, 0, [
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'dash' },
  { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 0 },
  { kind: 'none' },
  undefined,
]);
table.setCellBorders(0, 0, { left: { kind: 'none' } });
table.setCellBorders(0, 0, undefined);
```

Add `tableCellBorders: true` only when scalar/TRBL/named/clear snapshots, neighboring borders, text, fill, margins, direction, fit, and alignment are all preserved. Repeat equivalent browser assertions.

- [ ] **Step 2: Extend declaration smoke**

Import and instantiate:

```ts
type TableCellBorder,
type TableCellBorderInput,
type TableCellBorders,
type TableCellBorderStyle,
```

Type snapshot reads and scalar/TRBL/named/undefined setter calls. Include all new values in the existing `void [...]` expression so no unused declaration masks a missing export.

- [ ] **Step 3: Update public docs and compatibility claims**

Add examples to `docs/api/README.md` for scalar, TRBL, named partial, none, and clear. Document:

- point range and EMU quantization;
- TRBL public order versus L/R/T/B XML order;
- explicit none versus absence;
- omitted style versus direct solid/dash;
- strict partial snapshots and physical indices;
- preservation but non-editability of diagonals/advanced line properties;
- PptxGenJS omitted-border materialization and table/cell normalization quirks.

Update `packages/pptx/README.md`, `CHANGELOG.md`, and the compatibility row without claiming table creation, effective shared-edge resolution, advanced dash/line editing, or table-level mutation.

- [ ] **Step 4: Run focused smoke source and typecheck**

```sh
pnpm typecheck
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
```

Expected: type exports and focused runtime suites pass before the expensive full gates.

---

### Task 5: Full gates, tarball, native/baseline validation, visual QA, review, and delivery

**Files:**
- Review all Task 1-4 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation, built package, repository CLI, presentation overflow helper, LibreOffice, and same-source native/baseline files.
- Produces: reviewed `feat: support table cell borders` commit synchronized to `origin/main`.

- [ ] **Step 1: Run full functional and performance gates**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
```

Expected: all functional tests pass, only default performance is skipped in the full run, and isolated performance passes.

- [ ] **Step 2: Build and smoke the actual npm tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_borders_package_dir=$(mktemp -d /tmp/pptx-table-cell-borders-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_borders_package_dir"
node ../../scripts/smoke-npm-package.mjs "$table_cell_borders_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke all report true, including `tableCellBorders`.

- [ ] **Step 3: Generate same-source native and baseline files**

Use public PptxGenJS 4.0.1 to create one fixed six-cell table with direct noFill borders, stable geometry, labels, and distinct cell fills. Open one copy through `PptxDocument` and set:

1. all-side 2pt red solid;
2. mixed top theme dash, right zero-width green line, bottom none, left clear;
3. named top 3pt blue + left 1pt black, other sides clear;
4. four explicit none sides;
5. solid green borders and then clear all four;
6. untouched neighbor.

Create the baseline from the exact same source by applying the expected direct L/R/T/B line elements while preserving package timestamps and every unrelated part. Save:

- `/tmp/pptx-table-cell-borders-native/native.pptx`
- `/tmp/pptx-table-cell-borders-baseline/baseline.pptx`

Validate and compare:

```sh
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-borders-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-table-cell-borders-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-table-cell-borders-native/native.pptx /tmp/pptx-table-cell-borders-baseline/baseline.pptx
```

Expected: both packages report zero errors/warnings and added/removed/changed arrays are empty.

- [ ] **Step 4: Render every output, inspect at full size, and run overflow checks**

```sh
native_pdf_dir=$(mktemp -d /tmp/pptx-table-cell-borders-native-pdf.XXXXXX)
baseline_pdf_dir=$(mktemp -d /tmp/pptx-table-cell-borders-baseline-pdf.XXXXXX)
native_render_dir=$(mktemp -d /tmp/pptx-table-cell-borders-native-render.XXXXXX)
baseline_render_dir=$(mktemp -d /tmp/pptx-table-cell-borders-baseline-render.XXXXXX)
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir "$native_pdf_dir" /tmp/pptx-table-cell-borders-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir "$baseline_pdf_dir" /tmp/pptx-table-cell-borders-baseline/baseline.pptx
pdftoppm -png "$native_pdf_dir/native.pdf" "$native_render_dir/slide"
pdftoppm -png "$baseline_pdf_dir/baseline.pdf" "$baseline_render_dir/slide"
```

Inspect every page individually at full size. Confirm the four sides, red/blue/green/theme/black colors, point-width differences, dashed versus solid, explicit none, clear, zero-width line, cell fills, labels, and untouched neighbor are visibly correct; no repair, clipping, overlap, blur, missing edge, or unintended neighboring-cell difference occurs. Compare SHA-256 hashes of corresponding PNGs and require equality.

Run:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-borders-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-table-cell-borders-baseline/baseline.pptx
```

Expected: both overflow checks pass and raster hashes match.

- [ ] **Step 5: Final review, explicit staging, commit, push, and remote verification**

Run `git diff --check` and inspect the complete diff. Verify direct-side ownership, independent strict parsing, point quantization/range, none versus absence, style omission versus direct solid/dash, exact semantic no-op, PptxGenJS default/scalar/TRBL/zero quirks, nested snapshot isolation, diagonal/fill/text/neighbor preservation, schema-order insertion, and absence of table-creation/effective-style claims.

Confirm status lists only the intended implementation files plus `.pnpm-store/`. Stage these explicit intended paths:

```text
CHANGELOG.md
docs/api/README.md
docs/compatibility/pptxgenjs-baseline.md
packages/model/src/model.test.ts
packages/model/src/shapes.ts
packages/model/src/table-cell-borders.internal.ts
packages/pptx/README.md
packages/pptxgenjs-adapter/src/index.test.ts
packages/sdk/src/index.test.ts
scripts/smoke-npm-package.mjs
```

Run `git diff --cached --check`, inspect the staged diff, and commit:

```sh
git commit -m "feat: support table cell borders"
```

Push through the verified SSH-over-443 channel, fetch, and run:

```sh
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; `.pnpm-store/` remains untracked and untouched.
