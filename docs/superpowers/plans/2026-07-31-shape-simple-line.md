# Shape Simple Line Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking, and every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict native none/solid shape-line creation and lossless direct-state editing, including color, transparency, width, and all eight PptxGenJS 4.0.1 preset dash values.

**Architecture:** Add an internal `simple-line` value/OOXML codec that reuses the proven `simple-fill` color and transparency contract, while keeping `a:ln` container discovery, child ordering, and preservation in a shape-specific adapter. `AddShapeOptions` normalizes detached line state before rendering, and `ShapeModel.line` delegates read/edit through `SlideModel` so transactions, live identity, no-op behavior, and rollback remain centralized.

**Tech Stack:** TypeScript strict mode, Vitest, lossless source-span OOXML editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke tests, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- Public `ShapeLine` supports only `{ kind: 'none' }` or `{ kind: 'line', color, transparency?, width?, dash? }`.
- Public `ShapeLineDash` is exactly `solid | dash | dashDot | lgDash | lgDashDot | lgDashDotDot | sysDash | sysDot`.
- Color is strict six-digit sRGB or a supported scheme color; transparency is finite `0..100` quantized to `0.001%`.
- Width is finite `0..1584` points quantized to one EMU; omitted width materializes as `1` point and omitted dash materializes as `solid`.
- Inputs accept only ordinary or null-prototype objects with own data properties; reject arrays, class instances, symbols, accessors, unknown keys, missing line color, malformed colors, invalid transparency/width/dash, and PptxGenJS/deprecated aliases before package mutation.
- Creation omitted/runtime-`undefined` preserves the existing canonical empty `<a:ln/>`; explicit none writes direct `<a:noFill/>`; explicit zero transparency writes direct `a:alpha val="100000"`; explicit zero width writes `a:ln w="0"`.
- Read only the unique direct `p:spPr/a:ln`. Unsupported, absent, empty, malformed, namespace-lookalike, noncanonical-attribute, or ambiguous owned state returns `undefined`; do not calculate inherited/effective line state.
- Editing `undefined` removes owned `w`, fill-choice, and dash-choice state but preserves the `a:ln` container, canonical unrelated attributes, arrowheads, joins, extensions, and their relative order.
- Editing none removes owned width/dash and writes direct no-fill. Editing line writes normalized width, whole-replaces the line fill choice, and writes one preset dash.
- Same supported value is an exact bytes/journal no-op. Repeated `spPr`/`ln`, repeated/conflicting fill or dash choices, or unsafe insertion order throws `ModelParseError` with zero mutation.
- Existing gradient/picture/pattern/group line fill or custom dash is preserved by unrelated edits and may be explicitly replaced or cleared when unique.
- Do not change table-border behavior or add arrows, arrow sizes, line cap, compound/alignment editing, joins, custom dash creation, advanced line fills, shadow, hyperlink, adjustments, custom geometry, shape text, or percentage positions.
- Every failure and outer transaction rollback preserves exact slide bytes, relationships, parts, mutation journal, shape order, and live model identity.
- Node 20+ and the browser bundle expose the same public behavior from `@jiayunxie/pptx`.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed, fetched, and verified at divergence `0 0`.
- Execute inline without subagents or routine decision pauses.

---

### Task 1: Add the internal simple-line value and OOXML codec

**Files:**
- Create: `packages/model/src/simple-line.internal.ts`
- Create: `packages/model/src/simple-line.internal.test.ts`

**Interfaces:**
- Consumes: `RichTextColor`, `SimpleFill`, `normalizeSimpleFill()`, `readSimpleFillChoice()`, `renderSimpleFill()`, `simpleFillsEqual()`, and `XmlElement`.
- Produces: internal `SimpleLineDash`, `NormalizedSimpleLine`, `normalizeSimpleLine()`, `readSimpleLine()`, `renderSimpleLine()`, `simpleLinesEqual()`, `SIMPLE_LINE_FILL_CHOICE_NAMES`, and `SIMPLE_LINE_DASH_CHOICE_NAMES` for Tasks 2 and 3; Task 1 has no dependency on the later public shape-line types.

- [ ] **Step 1: Write descriptor-safe normalization tests**

Cover undefined, none, sRGB, scheme, explicit zero/50/100/fractional transparency, omitted/zero/fractional/max width, and all eight dashes. Require exact defaults and detachment:

```ts
expect(normalizeSimpleLine(undefined, 'Shape line')).toBeUndefined();
expect(normalizeSimpleLine({ kind: 'none' }, 'Shape line')).toEqual({ kind: 'none' });
expect(normalizeSimpleLine({
  kind: 'line',
  color: { kind: 'srgb', value: '#ff0000' },
}, 'Shape line')).toEqual({
  kind: 'line',
  color: { kind: 'srgb', value: 'FF0000' },
  width: 1,
  dash: 'solid',
});
expect(normalizeSimpleLine({
  kind: 'line',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 33.3334,
  width: 0.333333,
  dash: 'lgDashDotDot',
}, 'Shape line')).toEqual({
  kind: 'line',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 33.333,
  width: 0.33330708661417324,
  dash: 'lgDashDotDot',
});
```

Reject null/primitives/arrays/dates/class instances, inherited-only fields, own unknown string keys, all symbol keys, getters/setters without invoking them, invalid `kind`, extra keys on none, missing color, invalid color object/kind/value, NaN/infinities, transparency outside `0..100`, width outside `0..1584`, unknown/case-mismatched dash, and `type`/`alpha`/`dashType`/`lineDash`/`lineHead`/`lineTail` aliases.

- [ ] **Step 2: Write strict direct-line decode/render/equality tests**

Parse canonical and alternate-prefix `a:ln` elements and require:

```ts
renderSimpleLine({ kind: 'none' }, 'a:') === '<a:noFill/>';
renderSimpleLine({
  kind: 'line',
  color: { kind: 'srgb', value: 'FF0000' },
  transparency: 50,
  width: 2.5,
  dash: 'dashDot',
}, 'a:') ===
  '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>' +
  '<a:prstDash val="dashDot"/>';
```

Require the reader to combine `a:ln@w` with one strict simple fill and one `a:prstDash`, materializing absent `w`/dash as `1`/`solid` only for a nonempty supported solid line. Accept canonical unrelated `cap="flat"`, `cmpd="sng"`, and `algn="ctr"`, plus arrow/join/ext children. Return `undefined` for empty lines, line attributes outside the accepted canonical set, malformed/out-of-range/noninteger width, absent/conflicting/repeated fills, malformed solid fill, unsupported/custom/repeated dash, wrong namespace lookalikes, and multiple owned choices. Equality must compare kind, color, transparency presence/value, quantized width, and dash.

- [ ] **Step 3: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/simple-line.internal.test.ts --reporter=dot
```

Expected: FAIL because `simple-line.internal.ts` does not exist.

- [ ] **Step 4: Implement the internal codec**

Define the exact internal contract:

```ts
export type SimpleLineDash =
  | 'solid' | 'dash' | 'dashDot' | 'lgDash'
  | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot';

export type NormalizedSimpleLine =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly transparency?: number;
      readonly width: number;
      readonly dash: SimpleLineDash;
    };

export const SIMPLE_LINE_FILL_CHOICE_NAMES = SIMPLE_FILL_CHOICE_NAMES;
export const SIMPLE_LINE_DASH_CHOICE_NAMES = Object.freeze(['prstDash', 'custDash'] as const);

export function normalizeSimpleLine(
  value: unknown,
  context: string,
): NormalizedSimpleLine | undefined;

export function readSimpleLine(
  lineElement: XmlElement,
  prefix: string,
): NormalizedSimpleLine | undefined;

export function renderSimpleLine(
  line: NormalizedSimpleLine,
  prefix: string,
): string;

export function simpleLinesEqual(
  left: NormalizedSimpleLine | undefined,
  right: NormalizedSimpleLine,
): boolean;
```

Normalize `width` through `Math.round(points * 12700) / 12700`, reject values whose rounded EMU count exceeds `20_116_800`, and render `a:ln@w` in the container adapter rather than in `renderSimpleLine()`. Reuse simple-fill for nested color/transparency validation and exact XML; do not copy its scheme-color table or alpha rules.

- [ ] **Step 5: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/simple-line.internal.test.ts \
  packages/model/src/simple-fill.internal.test.ts \
  packages/model/src/table-create.internal.test.ts --reporter=dot
pnpm vitest run packages/model/src/model.test.ts -t 'table cell border|shape fill' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; table border/fill output and public declarations remain unchanged.

- [ ] **Step 6: Review, commit, push, and verify**

Review descriptor safety, one-EMU quantization, max width, eight-dash exhaustiveness, canonical unrelated attributes, alternate prefixes, simple-fill reuse, and absence of public exports. Then:

```sh
git add -- packages/model/src/simple-line.internal.ts \
  packages/model/src/simple-line.internal.test.ts
git diff --cached --check
git commit -m "feat: add simple line codec"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Create preset shapes with strict none or solid lines

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `SimpleLineDash`, `NormalizedSimpleLine`, `normalizeSimpleLine()`, and `renderSimpleLine()` plus the existing preset-shape normalizer/renderer.
- Produces: public `ShapeLineDash`, public `ShapeLine`, `AddShapeOptions.line?: ShapeLine`, and detached normalized creation state rendered into the existing shape skeleton.

- [ ] **Step 1: Add public compile-time and runtime creation tests**

Add exported-type assertions and create shapes with omitted, runtime-undefined, none, sRGB, scheme, explicit zero/50 transparency, omitted/zero/fractional/max width, and every dash:

```ts
const line = slide.addShape('rect', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 25,
    width: 2.5,
    dash: 'dashDot',
  },
});
const none = slide.addShape('ellipse', { line: { kind: 'none' } });
```

Require canonical empty `<a:ln/>` for omitted/runtime undefined, direct no-fill for none, exact solid fill/alpha/width/dash for line, unchanged geometry/transform/fill ordering, returned `ShapeModel`, stable order/identity, and caller detachment after nested input mutation.

- [ ] **Step 2: Add invalid and zero-mutation creation tests**

For all Task 1 invalid inputs, plus empty/missing-color line and PptxGenJS aliases, snapshot part bytes, relationships, parts, shape models, next shape ID, and mutation journal before `addShape()`; require identical state after rejection. Add `line` and nested field accessors to prove getters are never invoked.

- [ ] **Step 3: Extend exact preset normalization and renderer tests**

Require `normalizePresetShape()` to preserve omitted `line` as `undefined`, detach explicit line, and accept only the new `line` key. Keep the default renderer byte-identical. Add exact none and solid render assertions with line children ordered fill choice, dash, then any future preserved siblings:

```ts
export interface NormalizedPresetShape {
  readonly line: NormalizedSimpleLine | undefined;
}
```

The renderer must output `<a:ln/>` for undefined, `<a:ln><a:noFill/></a:ln>` for none, and `<a:ln w="31750">...<a:prstDash val="dashDot"/></a:ln>` for a 2.5-point dashed line.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape line' \
  packages/sdk/src/index.test.ts -t 'preset shape line' --reporter=dot
```

Expected: FAIL because `ShapeLine`, `ShapeLineDash`, and `AddShapeOptions.line` are absent.

- [ ] **Step 5: Implement the public contract and creation renderer**

Add the public types exactly:

```ts
export type ShapeLineDash =
  | 'solid' | 'dash' | 'dashDot' | 'lgDash'
  | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot';

export type ShapeLine =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly transparency?: number;
      readonly width?: number;
      readonly dash?: ShapeLineDash;
    };
```

Keep the public dash union structurally identical to Task 1 `SimpleLineDash`, with a compile-time bidirectional assignability assertion in the internal test. Add `line?: ShapeLine` to `AddShapeOptions`, `line` to `OPTION_KEYS`, and `line: normalizeSimpleLine(values.line, 'Preset shape line')` to the frozen normalized record. Render the existing empty line when undefined; otherwise render a canonical `a:ln` with `w` only for solid line and codec-owned children. No other shape skeleton bytes change.

- [ ] **Step 6: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/simple-line.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts -t 'preset shape|rollback' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 7: Review, commit, push, and verify**

Review public type scope, validation-before-mutation, exact default-byte compatibility, detached values, width/dash defaults, OOXML order, next-ID isolation, and rollback. Then:

```sh
git add -- packages/model/src/preset-shape.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create shape simple lines"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Read, replace, and clear direct shape lines

**Files:**
- Create: `packages/model/src/shape-line.internal.ts`
- Create: `packages/model/src/shape-line.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: public `ShapeLine`, Task 1 codec helpers, `LosslessXmlDocument`, `XmlElement`, and `ModelParseError`.
- Produces: `readShapeLine()`, `replaceShapeLine()`, `SlideModel.getShapeLine()`/`setShapeLine()`, and `ShapeModel.line`.

- [ ] **Step 1: Add strict shape-container read tests**

Build `p:sp` fixtures with canonical and alternate prefixes. Require detached snapshots for unique direct none and solid sRGB/scheme/alpha/width/every dash. Require `undefined` for absent or empty line, gradient/blip/pattern/group line fill, custom dash, nested lookalike, wrong namespace, missing/repeated direct `spPr`, repeated direct `ln`, repeated/conflicting owned choices, malformed line/fill/color/alpha/width/dash, and supported line nested under effects/text. Require arrowheads, round/bevel/miter joins, and ext children to coexist without appearing in the snapshot.

- [ ] **Step 2: Add replace/clear/order/isolation tests**

Require:

```ts
replaceShapeLine(xml, shape, sameLine, PART_URI) === false;
replaceShapeLine(xml, shape, { kind: 'none' }, PART_URI) === true;
replaceShapeLine(xml, shape, undefined, PART_URI) === true;
```

Test replacement and clear of unique no/solid/gradient/blip/pattern/group fill plus preset/custom dash. Test insertion when `a:ln` is absent: after shape fill choices and before `effectLst`, `effectDag`, `scene3d`, `sp3d`, or `extLst`. Preserve xfrm, geometry/adjustments, shape fill, effects, arrows, joins, extensions, text/nonvisual siblings, neighbor shapes, namespace declarations, canonical unrelated line attributes, and unknown bytes. Repeated `spPr`/`ln`, repeated/conflicting owned choices, or unsafe absent-line insertion must throw before patching.

- [ ] **Step 3: Add model/SDK identity, lifecycle, and rollback tests**

Require immediate `ShapeModel.line` on new/existing preset shapes and existing text-box `ShapeModel` instances; same-value byte/journal no-op; solid→none→clear→solid; unique advanced-line replacement and clear; arrow/join/ext preservation; duplicate independence; stable model object; outer rollback; six-format write/reopen; invalid setters with unchanged bytes/relationships/parts/journal. Do not add line creation options to `addText()`.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/shape-line.internal.test.ts \
  packages/model/src/model.test.ts -t 'shape line' \
  packages/sdk/src/index.test.ts -t 'shape line' --reporter=dot
```

Expected: FAIL because the shape-line adapter and model property are absent.

- [ ] **Step 5: Implement the strict shape-container adapter**

Expose only internal functions:

```ts
export function readShapeLine(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): ShapeLine | undefined;

export function replaceShapeLine(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  line: NormalizedSimpleLine | undefined,
  partUri: string,
): boolean;
```

Navigate direct presentation/drawing children by namespace URI. For existing `a:ln`, patch only owned `w`, fill-choice, and dash-choice bytes; keep the container and all unrelated attributes/children. For absent `a:ln`, derive the in-scope drawing prefix and insert at the schema-safe boundary. Never serialize or rebuild the full shape. Compute the current supported snapshot before patching so same-value assignment is an exact no-op.

- [ ] **Step 6: Add slide-owned API and public model property**

Add:

```ts
getShapeLine(id: number): ShapeLine | undefined;
setShapeLine(id: number, value: ShapeLine | undefined): void;
```

Normalize outside the package transaction when non-undefined; inside, resolve the current shape, call the strict replacer, and call `setXml()` only when it returns true. Add only to `ShapeModel`:

```ts
get line(): ShapeLine | undefined {
  return this.slide.getShapeLine(this.id);
}

set line(value: ShapeLine | undefined) {
  this.slide.setShapeLine(this.id, value);
}
```

Do not add this property to image/table/chart/group models.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/simple-line.internal.test.ts \
  packages/model/src/shape-line.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  -t 'shape line|shape fill|preset shape|duplicate slide|rollback' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 8: Review, commit, push, and verify**

Review direct namespace ownership, strict unsupported reads, schema-safe insertion, unique unsupported replacement, preservation of arrows/joins/extensions, no-op/journal behavior, identity, duplicate independence, six formats, and rollback. Then:

```sh
git add -- packages/model/src/shape-line.internal.ts \
  packages/model/src/shape-line.internal.test.ts packages/model/src/slide.ts \
  packages/model/src/shapes.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit shape simple lines"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public `PptxGenJS.ShapeType`, `slide.addShape()`, `write()`, native `ShapeLine`, creation/read/edit APIs, and direct OOXML inspection.
- Produces: public-only evidence for supported line semantics and explicit evidence for intentional strict divergences.

- [ ] **Step 1: Add a public-output shape-line fixture matrix**

Generate one shape per case using only public PptxGenJS APIs: omitted, explicit none, empty, solid missing color, sRGB, scheme, transparency 0/50, width 0/positive, all eight `dashType` values, deprecated `alpha`/`lineDash`, and begin/end arrows coexisting with line style. Capture packages through public `write()` without importing PptxGenJS internals.

- [ ] **Step 2: Compare supported semantic state**

For sRGB, scheme, 50% transparency, positive width, and every dash, open the public output with `PptxDocument`, compare shape name/geometry/transform and final line snapshot to native equivalents. Verify native edit of imported supported PptxGenJS line and write/reopen while preserving arrow nodes.

- [ ] **Step 3: Assert strict divergences explicitly**

Require PptxGenJS omitted and explicit none to contain empty `a:ln`, while native omitted keeps empty `a:ln` and native none writes direct no-fill. Require PptxGenJS empty/missing-color and width zero to fall back to `333333`/1pt/solid, while native rejects missing color and preserves zero width. Require PptxGenJS zero transparency to omit alpha while native writes `100000`. Require `alpha` and `lineDash` to be ignored by PptxGenJS shape lines and rejected by native input.

- [ ] **Step 4: Run conformance and dependency gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'shape line public output' --reporter=dot
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; `pptxgenjs` remains absent from non-adapter manifests.

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only evidence, supported semantic comparison, strict divergence assertions, arrow coexistence, output-warning isolation, and dependency boundaries. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare shape lines with pptxgenjs"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: the actual `@jiayunxie/pptx` tarball, Node ESM export, browser conditional export, generated declarations, and current package smoke harness.
- Produces: tarball-level `shapeLines: true` evidence in the smoke JSON.

- [ ] **Step 1: Extend the installed-package Node smoke**

In generated `smoke.mjs`, create sRGB/scheme/none/zero-width/eight-dash shapes, mutate original caller objects, read detached `ShapeModel.line`, perform same-value/replace/clear edits, preserve injected arrow/join/ext bytes, duplicate, write/reopen, and require zero relationship/part growth. Add `shapeLines` to the final JSON envelope without removing existing fields.

- [ ] **Step 2: Extend browser and declaration smoke**

In browser smoke, import `PptxDocument` from the package browser condition, create/read/edit/clear/reopen a dashed line, and verify no `node:` dependency. In generated typecheck source, import `ShapeLine`/`ShapeLineDash`, assign valid none/line values, use `AddShapeOptions.line`, and require invalid kind/color/transparency/width/dash/alias cases with `@ts-expect-error`.

- [ ] **Step 3: Build, pack, and run the actual tarball smoke**

```sh
pnpm build
mkdir -p /tmp/pptx-shape-line-pack
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-shape-line-pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-line-pack/jiayunxie-pptx-0.1.0.tgz
```

Expected JSON includes `"shapeLines":true`, existing `"shapeFills":true`, `"presetShapes":true`, `"types":true`, and CLI version `0.1.0`.

- [ ] **Step 4: Run script and repository gates**

```sh
node --check scripts/smoke-npm-package.mjs
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 5: Review, commit, push, and verify**

Review installed-tarball isolation, Node/browser parity, public type coverage, all eight dashes, existing smoke preservation, temporary cleanup, and absence of internal imports. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed shape lines"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Document the supported boundary and remaining shape gaps

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `CHANGELOG.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: Tasks 1–5 final public API and conformance evidence.
- Produces: user-facing creation/edit examples, changelog entry, and an honest compatibility-matrix update.

- [ ] **Step 1: Add concise public examples**

Document exact native syntax:

```ts
const shape = slide.addShape('roundRect', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 20,
    width: 2.5,
    dash: 'dashDot',
  },
});

shape.line = { kind: 'line', color: { kind: 'srgb', value: 'FF0000' } };
shape.line = { kind: 'none' };
shape.line = undefined;
```

Explain 1pt/solid defaults, zero width, `undefined` clear semantics, explicit none, direct-state snapshots, and preservation of arrow/join/extension state.

- [ ] **Step 2: Update compatibility and changelog**

Change the `slide.addShape()` row to include none/solid line creation/read/edit/clear, sRGB/theme, transparency, width, and eight preset dashes. Record PptxGenJS empty/missing-color, explicit-none, zero-width, zero-transparency, and deprecated-alias divergences without claiming byte identity. Keep arrows/sizes, cap/compound/alignment/join editing, advanced line fills/custom dash, shadow, hyperlink, adjustments, custom geometry, shape text, percentage positions, and broader drawing objects explicitly remaining.

- [ ] **Step 3: Validate examples and documentation consistency**

```sh
rg -n "ShapeLine|shape\.line|dashDot|zero width|arrow|advanced line" \
  README.md packages/pptx/README.md CHANGELOG.md \
  docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
git diff --check
```

Expected: examples match exported types and all remaining gaps remain visible.

- [ ] **Step 4: Review, commit, push, and verify**

Review API spelling, width units, none-versus-clear semantics, strict divergences, remaining-gap accuracy, and parity between root/package docs. Then:

```sh
git add -- README.md packages/pptx/README.md CHANGELOG.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document shape simple lines"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run release, real-PPTX, structural, visual, and import QA

**Files:**
- No repository file changes expected; use `/tmp/pptx-shape-line-qa-20260731`.

**Interfaces:**
- Consumes: Tasks 1–6 complete repository state.
- Produces: final evidence for package integrity, mutation isolation, real-office rendering, visual correctness, import compatibility, and remote parity.

- [ ] **Step 1: Run complete repository gates**

```sh
pnpm check
pnpm test:performance
pnpm build
```

Expected: strict typecheck/build, all tests, and independent performance budget pass.

- [ ] **Step 2: Pack and smoke the final tarball**

```sh
mkdir -p /tmp/pptx-shape-line-qa-20260731/pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-shape-line-qa-20260731/pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-line-qa-20260731/pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: identity/version, Node/browser/types, preset shapes, shape fills, shape lines, dependency boundaries, and CLI smoke all pass.

- [ ] **Step 3: Generate representative lifecycle decks**

Create native decks for omitted/none, sRGB/scheme, transparency 0/25/50/100, width 0/1/fractional/max, eight dashes, alternate geometries, existing-line read/edit/clear, advanced→simple replacement, arrows/joins/extensions preservation, duplicate, reopen, and second write. Also create PptxGenJS legal comparison and strict-divergence fixtures with labels and visible fills.

- [ ] **Step 4: Validate every legal package and mutation boundary**

For each legal deck run:

```sh
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json slides list deck.pptx
pptx-inspect --json part read deck.pptx /ppt/slides/slide1.xml
```

Use `package diff` to prove source→edit changes only the target slide part and edit→reopen→second-write has zero part changes. Verify exact direct fill/alpha/width/dash choices, preserved arrows/joins/extensions, shape IDs/order/names, and neighbor bytes.

- [ ] **Step 5: Render and inspect every slide**

Open/export all legal decks with an isolated LibreOffice profile, run `pdfinfo`, render PDF pages with `pdftoppm`, then run `render_slides.py` and `slides_test.py` for every PPTX. Inspect every rendered page at full size for color/theme/transparency, eight dash patterns, width ordering including hairline zero, none/clear visibility, arrow preservation, clipping, overlap, and label readability.

- [ ] **Step 6: Verify artifact-tool import behavior**

Import native representative/lifecycle decks with `PresentationFile.importPptx()`, require expected slide/shape counts, and inspect named shapes. Import PptxGenJS comparison fixtures and confirm supported files load without mutation.

- [ ] **Step 7: Verify repository and remote final state**

```sh
git diff --check
git status --short
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected: only the pre-existing untracked `.pnpm-store/` remains ignored, and divergence is `0 0`. Create no empty commit when Task 7 produces no repository change.
