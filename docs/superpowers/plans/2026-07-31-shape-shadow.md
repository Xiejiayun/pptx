# Shape Shadow Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking, and every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict native outer/inner preset-shape shadow creation and lossless direct-state editing compatible with valid PptxGenJS 4.0.1 shape-shadow output.

**Architecture:** Add a focused `simple-shadow` value/element codec for descriptor-safe normalization, quantized shadow parameters, color/alpha decoding, rendering, and semantic equality. Add a separate `shape-shadow` owner adapter for `p:spPr/a:effectLst` schema-order inspection and source-span edits; preset-shape creation renders the codec directly, while `ShapeModel.shadow` delegates existing-deck reads and transactional replacement through `SlideModel`.

**Tech Stack:** TypeScript strict mode, Vitest, lossless source-span OOXML editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke tests, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- Public `ShapeShadow` is a strict `kind: 'outer' | 'inner'` union. `undefined` is the only no-shadow/clear value; it covers PptxGenJS omitted and `{ type: 'none' }` behavior without inventing a direct OOXML no-shadow element.
- Outer/inner share optional `color`, `opacity`, `blur`, `angle`, and `distance`; outer alone accepts `rotateWithShape`. Inner with a defined rotate field fails before mutation.
- Defaults are the real PptxGenJS 4.0.1 shape output: sRGB black, opacity `0.75`, blur `8pt`, angle `270°`, distance `4pt`, and outer `rotateWithShape: false`.
- Preserve explicit zero. Opacity is finite `0..1` quantized to `1/100000`; blur is finite `0..100pt`, distance finite `0..200pt`, both quantized to 1 EMU; angle is finite `0 <= value < 360°` quantized to `1/60000°`.
- Color reuses strict sRGB/theme `RichTextColor`. Inputs, nested color, and getter snapshots are detached and deep-frozen.
- Only ordinary/null-prototype own data properties are accepted. Aliases, symbols, accessors, inherited fields, arrays/classes, coercions, unknown keys, invalid XML color values, and out-of-range numbers fail before package mutation.
- Read only a unique direct namespace-correct `a:innerShdw` or `a:outerShdw` from at most one direct `a:effectLst`; `a:effectDag`, multiple lists/shadows, wrong namespaces, malformed owned attributes/color/alpha, or unsafe schema order are unsupported.
- `blurRad`/`dist`/`dir` omission reads as OOXML zero defaults; alpha omission reads opacity 1; outer `rotWithShape` omission reads true. Compatible neutral outer `sx/sy/kx/ky/algn` bytes are preserved but not exposed.
- `ShapeModel.shadow` owns only one direct inner/outer child. It preserves other legal effects, effect-list container, geometry, fill, line/arrows, hyperlink, text, style, scene/3D/extensions, neighbors, relationships, parts, and unrelated lexical bytes.
- Same supported kind and normalized values is an exact bytes/journal no-op. Same-kind field changes patch owned spans; kind switches replace only the shadow child. Clear removes only the child and preserves the effect-list container.
- Effect-list child order is `blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge`. Shape-property insertion is after line and before scene3d/sp3d/extLst.
- Omitted/runtime-undefined creation leaves previously published preset-shape XML byte-for-byte unchanged.
- PptxGenJS malformed inner closing tag, ignored rotate flag, falsy defaults, hash color, invalid type, and invalid range passthrough are recorded as strict divergences, never copied.
- Do not add generic effect-stack, preset shadow, reflection, soft edge, effect DAG, custom shadow transform, or image/text/table/chart/media shadow APIs in this plan.
- Node 20+ and the browser bundle expose identical public behavior from `@jiayunxie/pptx`.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed, fetched, and verified at divergence `0 0`.
- Execute inline without subagents or routine decision pauses.

---

### Task 1: Add the strict simple-shadow value and element codec

**Files:**
- Create: `packages/model/src/simple-shadow.internal.ts`
- Create: `packages/model/src/simple-shadow.internal.test.ts`

**Interfaces:**
- Consumes: `RichTextColor`, `XmlElement`, DrawingML namespace/prefix helpers, point/EMU and degree/OOXML-angle conversions, and current descriptor-safe normalization conventions.
- Produces: `NormalizedShapeShadow`, `normalizeShapeShadow()`, `readSimpleShadow()`, `renderSimpleShadow()`, and `shapeShadowsEqual()` for Tasks 2–4.

- [ ] **Step 1: Write descriptor-safe normalization tests**

Cover outer/inner defaults, fully custom values, explicit zeros, min/max/fractional quantization, sRGB/theme color, null-prototype input, own runtime undefined, deep freeze, and immediate caller detachment:

```ts
expect(normalizeShapeShadow({ kind: 'outer' }, 'Shape shadow')).toEqual({
  kind: 'outer',
  color: { kind: 'srgb', value: '000000' },
  opacity: 0.75,
  blur: 8,
  angle: 270,
  distance: 4,
  rotateWithShape: false,
});

expect(normalizeShapeShadow({
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0,
  blur: 0,
  angle: 0,
  distance: 0,
}, 'Shape shadow')).toEqual({
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0,
  blur: 0,
  angle: 0,
  distance: 0,
});
```

Reject null/primitives/arrays/dates/class instances, inherited-only fields, own unknown string keys, every symbol key, getters/setters without invocation, missing/`none`/unknown/case-variant kind, PptxGenJS `type`/`offset` aliases, invalid color, non-number/NaN/infinity/negative/high values, angle 360, and inner rotate-with-shape.

- [ ] **Step 2: Write exact outer/inner render tests**

Require deterministic DrawingML using the supplied prefix:

```ts
expect(renderSimpleShadow(normalizeShapeShadow({ kind: 'outer' }, 'x'), 'a:'))
  .toBe(
    '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
    'rotWithShape="0" blurRad="101600" dist="50800" dir="16200000">' +
    '<a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr>' +
    '</a:outerShdw>',
  );
```

Require legal matching `innerShdw` tags without outer-only attributes, theme color, opacity `0/1`, maximum geometry, alternate prefix, and exact XML attribute escaping.

- [ ] **Step 3: Write strict element reader tests**

Read canonical outer/inner, alternate prefixes, optional `blurRad`/`dist`/`dir`, alpha omission, outer rotate omission/boolean lexical forms, and compatible neutral outer attributes. Require fully normalized detached deep-frozen snapshots.

Return `undefined` for wrong element/prefix/namespace, unknown/repeated/qualified attributes, invalid integer/range/decimal raw values, inner outer-only attributes, partial/invalid neutral outer attributes, zero/multiple color children, preset/system color, malformed sRGB/theme value, extra color transform, repeated/qualified/invalid alpha, child elements under alpha, and unexpected shadow children.

- [ ] **Step 4: Write semantic equality tests**

Require equal cloned colors/values to compare true, kind/color/opacity/blur/angle/distance/rotate differences false, and `undefined` equality only with `undefined`:

```ts
expect(shapeShadowsEqual(outer, { ...outer, color: { ...outer.color } })).toBe(true);
expect(shapeShadowsEqual(outer, { ...outer, blur: outer.blur + 1 })).toBe(false);
expect(shapeShadowsEqual(undefined, undefined)).toBe(true);
```

- [ ] **Step 5: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/simple-shadow.internal.test.ts --reporter=dot
```

Expected: FAIL because `simple-shadow.internal.ts` does not exist.

- [ ] **Step 6: Implement the minimal codec**

Define the exact surface:

```ts
export type NormalizedShapeShadow = Readonly<
  | {
      readonly kind: 'outer';
      readonly color: RichTextColor;
      readonly opacity: number;
      readonly blur: number;
      readonly angle: number;
      readonly distance: number;
      readonly rotateWithShape: boolean;
    }
  | {
      readonly kind: 'inner';
      readonly color: RichTextColor;
      readonly opacity: number;
      readonly blur: number;
      readonly angle: number;
      readonly distance: number;
    }
>;

export function normalizeShapeShadow(
  value: unknown,
  context: string,
): NormalizedShapeShadow;

export function readSimpleShadow(
  element: XmlElement,
  prefix: string,
): NormalizedShapeShadow | undefined;

export function renderSimpleShadow(
  shadow: NormalizedShapeShadow,
  prefix: string,
): string;

export function shapeShadowsEqual(
  left: NormalizedShapeShadow | undefined,
  right: NormalizedShapeShadow | undefined,
): boolean;
```

Use property descriptors rather than public reads. Freeze nested color before freezing the final value. Keep reading non-throwing. Parse only exact expanded-name state, use integer lexical checks, and quantize public values once during normalization.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/simple-shadow.internal.test.ts \
  packages/model/src/simple-fill.internal.test.ts \
  packages/model/src/simple-line.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; no public export or shape mutation exists yet.

- [ ] **Step 8: Review, commit, push, and verify**

Review descriptor safety, exact defaults, quantization boundaries, zero preservation, deep freeze, theme catalog, XML escaping, namespace handling, optional OOXML defaults, strict invalid-state rejection, and equality completeness. Then:

```sh
git add -- packages/model/src/simple-shadow.internal.ts \
  packages/model/src/simple-shadow.internal.test.ts
git diff --cached --check
git commit -m "feat: add shape shadow codec"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Create preset shapes with strict outer and inner shadows

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `NormalizedShapeShadow`, `normalizeShapeShadow()`, and `renderSimpleShadow()` plus current preset-shape normalization/rendering and transaction behavior.
- Produces: public `ShapeShadow`, `AddShapeOptions.shadow?: ShapeShadow`, normalized creation state, and canonical effect-list XML that Task 3 exposes through `ShapeModel.shadow`.

- [ ] **Step 1: Add public type and creation tests**

Export the exact public union from `preset-shape.ts` and exercise it through `addShape()`:

```ts
const outer = slide.addShape('roundRect', {
  name: 'Outer shadow',
  shadow: {
    kind: 'outer',
    color: { kind: 'srgb', value: '123ABC' },
    opacity: 0.42,
    blur: 7.25,
    angle: 123.4,
    distance: 5.5,
    rotateWithShape: true,
  },
});

const inner = slide.addShape('ellipse', {
  shadow: { kind: 'inner', color: { kind: 'scheme', value: 'accent2' } },
});
```

Require exact canonical outer/inner XML, immediate stable shape identity, caller/nested-color detachment, coexistence with fill/line/arrows/hyperlink, shape order/IDs, and all 178 preset geometries remaining creatable.

- [ ] **Step 2: Prove omitted behavior and invalid zero mutation**

Require omitted and runtime-undefined shadow creation to preserve the previously published preset-shape XML exactly. For every Task 1 invalid input, snapshot slide bytes, parts, relationships, shape identity map behavior, next shape ID, and mutation journal before `addShape()`; require identical state after rejection.

Add compile-time `@ts-expect-error` cases for missing/`none` kind, `type`/`offset` aliases, inner rotate, invalid color/value types, unknown key, and symbol-like incompatible values. Runtime accessor tests must prove no getter invocation.

- [ ] **Step 3: Extend preset normalization and rendering tests**

Add `shadow` to accepted option keys and normalized state:

```ts
export interface NormalizedPresetShape {
  readonly shadow: NormalizedShapeShadow | undefined;
}
```

Require no-shadow bytes unchanged and supplied shadow rendered immediately after the existing line and before `</p:spPr>`. Verify outer defaults, inner defaults, fully custom/zero/theme cases, matching tags, canonical attribute order, and no effect list for omission.

- [ ] **Step 4: Add transaction and format tests**

Wrap shadow creation in an outer transaction that throws after `addShape()` and require exact rollback of slide bytes, mutation journal, IDs, live model cache, and unrelated relationships. Write/reopen outer and inner shapes in `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`, then inspect exact shadow snapshots through Task 1 until Task 3 adds the public getter.

- [ ] **Step 5: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape shadow' \
  packages/sdk/src/index.test.ts -t 'preset shape shadow' --reporter=dot
```

Expected: FAIL because `ShapeShadow` and creation integration are absent.

- [ ] **Step 6: Implement public type, normalization, and rendering integration**

Define `ShapeShadowBase`/`ShapeShadow` in `preset-shape.ts`, add `shadow` to `AddShapeOptions`, normalize through Task 1 before transaction mutation, and store it on `NormalizedPresetShape`. In `renderPresetShapeXml()` append:

```ts
const effect = shape.shadow === undefined
  ? ''
  : `<a:effectLst>${renderSimpleShadow(shape.shadow, 'a:')}</a:effectLst>`;
```

Do not change the no-shadow renderer or add a second effect list. Keep shadow independent from line/arrows and hyperlink relationship creation.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape|shape shadow' \
  packages/sdk/src/index.test.ts -t 'preset shape|shape shadow' --reporter=dot
pnpm vitest run packages/model/src/shape-fill.internal.test.ts \
  packages/model/src/shape-line.internal.test.ts \
  packages/model/src/shape-arrows.internal.test.ts \
  packages/model/src/shape-hyperlink.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; existing no-shadow/fill/line/arrows/hyperlink bytes remain exact.

- [ ] **Step 8: Review, commit, push, and verify**

Review public union correctness, export reachability, option-key strictness, validation-before-ID allocation, default/zero semantics, inner well-formedness, effect-list placement, no-shadow byte stability, rollback, identity, six formats, and absence of non-shape APIs. Then:

```sh
git add -- packages/model/src/preset-shape.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create shape shadows"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Read, edit, preserve siblings, and clear shape shadows

**Files:**
- Create: `packages/model/src/shape-shadow.internal.ts`
- Create: `packages/model/src/shape-shadow.internal.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Tasks 1–2 public/normalized values, lossless XML source spans, shape-property/effect-list schema order, `ModelParseError`, and existing OPC transaction/model delegation.
- Produces: `readShapeShadow()`, `replaceShapeShadow()`, `ShapeModel.shadow`, `SlideModel.getShapeShadow()`, and `SlideModel.setShapeShadow()`.

- [ ] **Step 1: Write shape-owner reader and insertion tests**

Read a unique direct inner/outer shadow under the unique `p:spPr`. Return `undefined` for no shadow, wrong shape root, repeated/wrong-namespace `spPr`, effect DAG, repeated/wrong-namespace effect list, multiple inner/outer shadow, malformed Task 1 element, or unsafe effect order.

For insertion, cover no effect list, self-closing effect list, empty expanded list, and every legal sibling stage. Require inner inserted after glow and before outer/prst/reflection/soft-edge; outer inserted after inner and before preset/reflection/soft-edge. Preserve all sibling bytes and alternate prefixes.

- [ ] **Step 2: Write exact same-kind lossless patch tests**

Start with alternate prefixes, single quotes, optional attrs omitted, compatible neutral outer attrs, alpha omitted, namespace declarations, and legal sibling effects. Change one field at a time and require only the owned value/element span to change:

- blur/distance/angle existing value replacement or missing-attribute insertion;
- outer rotate existing replacement or missing insertion;
- same-kind sRGB/theme value replacement;
- alpha existing replacement or missing child insertion;
- color-kind switch replacing only the color child while retaining a valid DrawingML binding.

Same normalized value must return false and preserve exact source. Explicit default/zero assignments must not collapse through truthiness.

- [ ] **Step 3: Write kind-switch, clear, and preservation tests**

Switch outer↔inner and require one legal matching-tag child at the correct schema stage. Clear only the shadow child, preserve empty `effectLst`, and make absent clear exact no-op. Preserve glow/reflection/soft-edge/preset shadow, effect-list lexical bytes, geometry/adjustments, fill, advanced line/arrows, hyperlink, text, scene3d/sp3d/extLst, neighbors, relationships, and unrelated parts.

- [ ] **Step 4: Write malformed zero-mutation tests**

For every unsupported state from Tasks 1 and 3, require getter `undefined` without mutation and setter `ModelParseError` with exact equality of slide bytes, parts, relationships, graph, journal, and live identity. Cover malformed color/alpha/attrs, partial neutral outer transforms, effectDag, multiple lists/shadows, wrong namespace, invalid sibling stage, repeated shape properties, and unsafe property insertion order.

- [ ] **Step 5: Add public lifecycle tests**

Expose exact methods:

```ts
class ShapeModel {
  get shadow(): ShapeShadow | undefined;
  set shadow(value: ShapeShadow | undefined);
}

class SlideModel {
  getShapeShadow(id: number): ShapeShadow | undefined;
  setShapeShadow(id: number, value: ShapeShadow | undefined): void;
}
```

Test imported/native outer/inner, deep-frozen detached snapshots, same-value no-op, every field edit, color-kind/kind switch, clear, stable identity, duplicate, move, source deletion, outer rollback, write/reopen/second-write, and six formats. Duplicate must preserve shadow/sibling effects without aliasing future edits.

- [ ] **Step 6: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/model/src/shape-shadow.internal.test.ts \
  packages/model/src/model.test.ts -t 'shape shadow lifecycle' \
  packages/sdk/src/index.test.ts -t 'shape shadow lifecycle' --reporter=dot
```

Expected: FAIL because the owner adapter and public getter/setter are absent.

- [ ] **Step 7: Implement strict owner inspection and source-span editing**

Define:

```ts
export function readShapeShadow(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedShapeShadow | undefined;

export function replaceShapeShadow(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  shadow: NormalizedShapeShadow | undefined,
  partUri: string,
): boolean;
```

Separate non-throwing read inspection from mutation inspection. Mutation must validate the complete owned state and insertion stages before creating edits. Use descending source-span edits for same-kind values; render only absent attrs/alpha, switched color child, kind-switched shadow, or newly inserted effect list. Preserve an empty effect-list container on clear.

- [ ] **Step 8: Implement model delegation**

Normalize public input before entering the transaction. `SlideModel.getShapeShadow()` resolves and reads the shape; `setShapeShadow()` resolves inside the transaction, calls `replaceShapeShadow()`, and serializes only when it returns true. Add direct delegation to `ShapeModel`; do not add relationship or presentation lifecycle logic.

- [ ] **Step 9: Run focused, lifecycle, and full model gates**

```sh
pnpm vitest run packages/model/src/simple-shadow.internal.test.ts \
  packages/model/src/shape-shadow.internal.test.ts \
  packages/model/src/model.test.ts -t 'shape shadow' \
  packages/sdk/src/index.test.ts -t 'shape shadow' --reporter=dot
pnpm vitest run packages/model/src/model.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; existing shape and effect-preservation tests remain green.

- [ ] **Step 10: Review, commit, push, and verify**

Review complete-state validation before edits, schema stages, alternate prefixes, optional/default attrs, no-op comparison, field-span isolation, color/alpha handling, kind-switch tags/order, empty-container preservation, sibling/unrelated byte preservation, malformed rejection, rollback, identity, duplicate, reopen, and six formats. Then:

```sh
git add -- packages/model/src/shape-shadow.internal.ts \
  packages/model/src/shape-shadow.internal.test.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit shape shadows"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Prove PptxGenJS and packed Node/browser/type compatibility

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 3 final public behavior, locked `pptxgenjs@4.0.1`, actual package tarball, Node/browser conditional exports, generated declarations, and current smoke harness.
- Produces: public-output conformance/strict-divergence evidence and tarball-level `shapeShadows: true` evidence.

- [ ] **Step 1: Add public-only PptxGenJS shape-shadow cases**

Generate each structural class in an isolated deck through public `addShape()` and `write()` only:

```ts
const cases = [
  ['omitted', undefined],
  ['none', { type: 'none' }],
  ['outer defaults', { type: 'outer' }],
  ['outer custom', {
    type: 'outer', color: '123ABC', opacity: 0.42,
    blur: 7.25, angle: 123.4, offset: 5.5, rotateWithShape: true,
  }],
  ['outer zero', {
    type: 'outer', color: '000000', opacity: 0,
    blur: 0, angle: 0, offset: 0, rotateWithShape: false,
  }],
  ['inner', { type: 'inner' }],
  ['hash color', { type: 'outer', color: '#ABCDEF' }],
  ['unknown type', { type: 'bogus', color: 'FF0000' }],
  ['invalid ranges', {
    type: 'outer', color: '00FF00', opacity: 2,
    blur: -1, angle: 400, offset: 201,
  }],
];
```

Assert valid outer exact XML/defaults/custom quantization, ignored rotate flag, and falsy fallback. Keep inner/unknown/invalid cases isolated so malformed output does not contaminate valid fixtures; require the exact wrong inner closing tag and validator/XML parser failure.

- [ ] **Step 2: Compare supported native semantics and strict divergences**

Create native equivalents for omitted, outer defaults/custom, and a valid inner. Compare valid outer normalized color/opacity/blur/angle/distance and visual semantics without requiring byte-identical attribute order or ignored rotate behavior. Require native explicit zero to remain zero, theme color to work, inner to be well-formed, and invalid/alias/accessor inputs to throw before mutation.

Document that native `undefined` covers PptxGenJS none, `distance` replaces legacy offset, and PptxGenJS 4.0.1 shape runtime emits malformed inner plus invalid passthrough.

- [ ] **Step 3: Extend installed-package Node smoke**

In generated `smoke.mjs`, create outer/inner/default/zero/theme shadows, mutate caller objects, read deep-frozen detached snapshots, edit each field and kind, preserve a sibling glow fixture, duplicate, clear, write/reopen, and require valid diagnostics. Add `shapeShadows` to the final JSON envelope without removing existing fields.

- [ ] **Step 4: Extend browser and declaration smoke**

In browser smoke, import `PptxDocument`/`ShapeModel` from the browser condition and create/read/edit/clear/reopen outer/inner shadows without `node:` dependencies. In generated typecheck source, import `ShapeShadow`, use both union branches, and add `@ts-expect-error` cases for missing/none kind, inner rotate, offset/type aliases, unknown key, and invalid field types.

- [ ] **Step 5: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'shape shadow public output' --reporter=dot
```

Expected: FAIL because conformance coverage and the packed smoke field are absent.

- [ ] **Step 6: Implement conformance and packed smoke coverage**

Add only tests/smoke behavior; production adapter code remains unchanged because `importPptxGenJS()` already imports public `write()` bytes generically. Parse valid outer output through `PptxDocument`; inspect malformed inner/invalid output at raw ZIP/XML level and with the validator rather than presenting it as importable native state.

- [ ] **Step 7: Build, pack, and run compatibility gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
pnpm typecheck
pnpm build
mkdir -p /tmp/pptx-shape-shadow-pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-shape-shadow-pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-shadow-pack/jiayunxie-pptx-0.1.0.tgz
node --check scripts/smoke-npm-package.mjs
git diff --check
```

Expected JSON includes `"shapeShadows":true`, `"shapeHyperlinks":true`, `"shapeArrows":true`, `"shapeLines":true`, `"shapeFills":true`, `"presetShapes":true`, `"types":true`, and CLI version `0.1.0`.

- [ ] **Step 8: Review, commit, push, and verify**

Review public-only generation, isolated malformed fixtures, actual shape defaults, explicit-zero divergence, ignored rotate flag, valid native inner, strict zero-mutation rejection, installed-tarball isolation, Node/browser parity, declaration checks, existing smoke preservation, and dependency boundaries. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify shape shadow compatibility"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Document supported shape shadows and remaining effect gaps

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: Tasks 1–4 final API, direct-state semantics, conformance evidence, and strict divergence results.
- Produces: create/edit examples, changelog entry, API contract, and an honest PptxGenJS parity update.

- [ ] **Step 1: Add concise public examples**

Document exact native syntax:

```ts
const shape = slide.addShape('roundRect', {
  shadow: {
    kind: 'outer',
    color: { kind: 'srgb', value: '000000' },
    opacity: 0.35,
    blur: 6,
    angle: 45,
    distance: 4,
  },
});

shape.shadow = {
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0.5,
  blur: 3,
  angle: 270,
  distance: 2,
};
shape.shadow = undefined;
```

Explain points/degrees, opacity `0..1`, defaults, zero preservation, outer-only rotate, detached deep-frozen snapshots, whole replacement, same-value no-op, effect-sibling preservation, and clear semantics.

- [ ] **Step 2: Update compatibility, API, and changelog**

Change the `slide.addShape()` row to include outer/inner shadow create/read/edit/clear/duplicate/reopen. Record PptxGenJS actual defaults, none/omission mapping, legacy offset versus native distance, zero fallback, ignored rotate flag, invalid passthrough, and malformed inner output without claiming byte identity.

Remove simple shape shadow from the pending shape list while retaining generic/advanced effects, custom shadow transforms, non-shape shadow, custom geometry, adjustments, advanced line state, shape-text options, and percentage positions. Change the next shape priority from shadow to custom geometry/adjustments.

- [ ] **Step 3: Validate examples and documentation consistency**

```sh
rg -n "ShapeShadow|shadow|outer|inner|opacity|distance|custom geometry" \
  README.md packages/pptx/README.md CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
git diff --check
```

Expected: examples compile against public names, numeric ranges/defaults are consistent, PptxGenJS defects are explicit, and every remaining effect gap stays visible.

- [ ] **Step 4: Review, commit, push, and verify**

Review API spelling, unit/range/default wording, `undefined`/none mapping, zero behavior, outer/inner legality, whole replacement, sibling preservation, PptxGenJS divergences, and remaining-gap accuracy across all five docs. Then:

```sh
git add -- README.md packages/pptx/README.md CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document shape shadows"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Run release, structural, visual, and import QA

**Files:**
- No repository file changes expected; use `/tmp/pptx-shape-shadow-qa-20260731`.

**Interfaces:**
- Consumes: Tasks 1–5 complete repository state.
- Produces: final evidence for package integrity, mutation isolation, Office-compatible shadows, visual correctness, import compatibility, and remote parity.

- [ ] **Step 1: Run complete repository gates**

```sh
pnpm check
pnpm test:performance
pnpm build
```

Expected: strict typecheck/build, all tests, and independent performance budget pass.

- [ ] **Step 2: Pack and smoke the final tarball**

```sh
mkdir -p /tmp/pptx-shape-shadow-qa-20260731/pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-shape-shadow-qa-20260731/pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-shadow-qa-20260731/pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: identity/version, Node/browser/types, preset shapes, fills, lines, arrows, hyperlinks, shadows, dependency boundaries, and CLI all pass.

- [ ] **Step 3: Generate representative and lifecycle decks**

Create native decks for outer/inner defaults, sRGB/theme, opacity/blur/angle/distance boundaries, explicit zero, rotate true/false, outer↔inner, field edits, clear, sibling glow/reflection preservation, duplicate, rollback, malformed zero-mutation, reopen, and second write. Create isolated PptxGenJS valid outer/default/custom/none comparisons and malformed inner/invalid-divergence fixtures with visibly labeled shapes.

- [ ] **Step 4: Validate packages and exact mutation boundaries**

For each legal deck run:

```sh
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json slides list deck.pptx
pptx-inspect --json part read deck.pptx /ppt/slides/slide1.xml
```

Use `package diff` to prove same-value/reopen/second-write change no parts; each shadow edit changes only its owning slide part; clear and kind switch preserve other effect/shape bytes; malformed rejection changes no parts. Require exact tags/attributes/color/alpha/order, names/counts, valid relationships, zero errors/warnings for legal native/PptxGenJS outer fixtures, and exact XML/validator failure for isolated PptxGenJS inner/invalid fixtures.

- [ ] **Step 5: Round-trip through LibreOffice and inspect visual output**

Open/export every legal native and PptxGenJS outer deck with isolated LibreOffice profiles. Reinspect round-tripped PPTX shadow XML for kind/color/alpha/blur/distance/direction/rotate behavior and record any client normalization separately. Export to PDF, run `pdfinfo`, rasterize with `pdftoppm`, then run Presentations `render_slides.py` and `slides_test.py` for every legal PPTX.

Inspect every page at original detail for visibly distinct outer versus inner shadows, zero/no-shadow, opacity, blur, distance, angle, rotate behavior, unchanged fill/line/arrows/text, all labels, clipping, overlap, blur artifacts, and overflow. Structural OOXML is the semantic authority; raster output proves visual non-regression and client rendering.

- [ ] **Step 6: Verify artifact-tool import behavior**

Import native representative/lifecycle decks and valid PptxGenJS outer comparison with `PresentationFile.importPptx()`. Require expected slide/object counts, named shadow shapes, and unchanged source SHA-256 hashes. Do not import malformed PptxGenJS inner as a legal deck.

- [ ] **Step 7: Verify repository and remote final state**

```sh
git diff --check
git status --short
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: only the pre-existing untracked `.pnpm-store/` remains, and divergence is `0 0`. Create no empty commit when Task 6 produces no repository change; commit and push only if real QA findings require documentation or code changes.
