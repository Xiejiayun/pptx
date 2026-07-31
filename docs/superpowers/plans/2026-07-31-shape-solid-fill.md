# Shape Solid/No-Fill Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking, and every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict native solid/no-fill creation and lossless direct-state editing for preset shapes, with PptxGenJS 4.0.1 conformance and packed Node/browser/type coverage.

**Architecture:** Extract the already-proven color/transparency/value logic from table-cell fill into an internal `simple-fill` module, while keeping container navigation and child ordering in table-cell- and shape-specific adapters. `AddShapeOptions` normalizes a detached fill before rendering, and `ShapeModel.fill` delegates read/edit through `SlideModel` so OPC transactions, live identity, no-op behavior, and rollback stay centralized.

**Tech Stack:** TypeScript strict mode, Vitest, lossless source-span OOXML editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke tests, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- Public `ShapeFill` supports only `{ kind: 'none' }` and `{ kind: 'solid', color, transparency? }`; color is strict sRGB or supported scheme color, transparency is finite `0..100` quantized to `0.001%`.
- Creation omitted/`undefined`/none produces canonical direct `a:noFill`; explicit solid transparency zero writes direct `a:alpha val="100000"`.
- Inputs accept only ordinary or null-prototype objects with own data properties; reject arrays, class instances, symbols, accessors, unknown keys, missing solid color, malformed colors, and invalid transparency before package mutation.
- Read only the unique direct `p:spPr` fill choice. Unsupported, absent, malformed, namespace-lookalike, or ambiguous direct state returns `undefined`; do not calculate inherited/effective fill.
- Editing `undefined` deletes one unique direct fill choice, none/solid whole-replaces it, and absent insertion occurs directly after the unique direct preset/custom geometry.
- Same supported value is an exact no-op. Repeated `spPr`/fill choices or unsafe insertion structure throws `ModelParseError` with zero mutation.
- Existing gradient/picture/pattern/group fill is preserved by unrelated edits; explicit simple replacement or clear may replace/remove one unique unsupported fill choice.
- Do not add line, arrows, shadow, hyperlink, adjustment, custom geometry, shape text, advanced fill, slide background, or a combined simple/gradient facade.
- Every failure and outer transaction rollback preserves exact slide bytes, relationships, parts, mutation journal, shape order, and live model identity.
- Node 20+ and the browser bundle expose the same public behavior from `@jiayunxie/pptx`.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed, fetched, and verified at divergence `0 0`.
- The user has authorized inline autonomous execution; do not dispatch subagents or pause for routine choices.

---

### Task 1: Extract the shared simple-fill value codec without changing table behavior

**Files:**
- Create: `packages/model/src/simple-fill.internal.ts`
- Create: `packages/model/src/simple-fill.internal.test.ts`
- Modify: `packages/model/src/table-cell-fill.internal.ts`

**Interfaces:**
- Consumes: `RichTextColor`, `LosslessXmlDocument`, `XmlElement`, and XML attribute escaping.
- Produces: internal `SimpleFill`, `normalizeSimpleFill()`, `readSimpleFillChoice()`, `renderSimpleFill()`, `simpleFillsEqual()`, and `SIMPLE_FILL_CHOICE_NAMES` for Tasks 2 and 3; existing `normalizeTableCellFill()`, `readTableCellFill()`, `replaceTableCellFill()`, and `renderTableCellFill()` retain their signatures and behavior.

- [ ] **Step 1: Write focused shared-value tests**

Create `simple-fill.internal.test.ts` with exact normalize cases:

```ts
expect(normalizeSimpleFill(undefined, 'Fill')).toBeUndefined();
expect(normalizeSimpleFill({ kind: 'none' }, 'Fill')).toEqual({ kind: 'none' });
expect(normalizeSimpleFill({
  kind: 'solid',
  color: { kind: 'srgb', value: '#ff0000' },
  transparency: 33.3334,
}, 'Fill')).toEqual({
  kind: 'solid',
  color: { kind: 'srgb', value: 'FF0000' },
  transparency: 33.333,
});
expect(normalizeSimpleFill({
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 0,
}, 'Fill')).toEqual({
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 0,
});
```

Reject null/primitives/arrays/dates/class instances, inherited-only fields, own unknown string keys, all symbol keys, getters/setters without invoking them, invalid kind, extra keys on none, missing color, invalid color object/kind/value, NaN/infinities, and transparency outside `0..100`. Prove normalization returns new nested values detached from later caller mutation.

- [ ] **Step 2: Write strict choice decode/render/equality tests**

Parse fixtures using `a:` and alternate valid prefixes. Require:

```ts
renderSimpleFill({ kind: 'none' }, 'a:') === '<a:noFill/>';
renderSimpleFill({
  kind: 'solid',
  color: { kind: 'srgb', value: 'FF0000' },
  transparency: 50,
}, 'a:') ===
  '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>';
```

Read strict no-fill, sRGB, scheme, absent alpha, alpha 0/100000, and fractional transparency. Return `undefined` for attributes on fill/color/alpha, multiple colors/alpha transforms, non-alpha transforms, wrong prefix/namespace lookalikes, invalid hex/scheme, empty/decimal/out-of-range alpha, and unsupported choice names. Equality must compare kind, normalized color kind/value, and presence/value of transparency.

- [ ] **Step 3: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/simple-fill.internal.test.ts --reporter=dot
```

Expected: FAIL because `simple-fill.internal.ts` does not exist.

- [ ] **Step 4: Implement the internal shared codec**

Define the internal structural contract and functions:

```ts
export type SimpleFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export const SIMPLE_FILL_CHOICE_NAMES = Object.freeze([
  'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill',
] as const);

export function normalizeSimpleFill(value: unknown, context: string): SimpleFill | undefined;
export function readSimpleFillChoice(choice: XmlElement, prefix: string): SimpleFill | undefined;
export function renderSimpleFill(fill: SimpleFill, prefix: string): string;
export function simpleFillsEqual(left: SimpleFill | undefined, right: SimpleFill): boolean;
```

Move only the descriptor-safe object reading, scheme set, color/transparency normalization, fill-choice decoding, deterministic rendering, and equality logic from `table-cell-fill.internal.ts`. Preserve error types/messages except for the caller-provided context.

- [ ] **Step 5: Rewire table-cell fill as a container adapter**

Keep `TableCellFill` public ownership in `shapes.ts`. Make `normalizeTableCellFill()` delegate to `normalizeSimpleFill()` and return the structurally identical result. Keep direct `tcPr` discovery, fill-choice multiplicity, extension-list ordering, and `ModelParseError` behavior local; replace private decode/render/equality calls with shared functions. Keep `renderTableCellFill(fill, prefix)` as the existing exported wrapper.

- [ ] **Step 6: Run focused and table regression gates**

```sh
pnpm vitest run packages/model/src/simple-fill.internal.test.ts --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  packages/model/src/table-create.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; public declarations and table XML snapshots do not change.

- [ ] **Step 7: Review, commit, push, and verify**

Review descriptor safety, color set, transparency quantization, alpha presence, alternate prefixes, exact table behavior, and absence of public exports. Then:

```sh
git add -- packages/model/src/simple-fill.internal.ts \
  packages/model/src/simple-fill.internal.test.ts \
  packages/model/src/table-cell-fill.internal.ts
git diff --cached --check
git commit -m "refactor: share simple fill codec"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Create preset shapes with strict solid or none fill

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `normalizeSimpleFill()`/`renderSimpleFill()` and existing preset-shape normalizer/renderer.
- Produces: public `ShapeFill`, `AddShapeOptions.fill?: ShapeFill`, and detached normalized creation state rendered into the existing shape skeleton.

- [ ] **Step 1: Add public compile-time and runtime creation tests**

Extend `AddShapeOptions` type assertions and create shapes with omitted, runtime-undefined, none, sRGB, scheme, and transparent solid fills:

```ts
const red = slide.addShape('rect', {
  fill: { kind: 'solid', color: { kind: 'srgb', value: '#ff0000' } },
});
const themed = slide.addShape('ellipse', {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 25,
  },
});
const none = slide.addShape('star5', { fill: { kind: 'none' } });
```

Require direct `a:noFill`, exact `solidFill` color/alpha, unchanged geometry/transform/empty line, returned `ShapeModel`, stable order/identity, and caller detachment after nested input mutation.

- [ ] **Step 2: Add invalid and zero-mutation creation tests**

For invalid fill objects from Task 1 plus `alpha`, `type`, missing solid color, unsupported scheme, invalid hex, and invalid transparency, snapshot part bytes, relationships, parts, shape models, and mutation journal before `addShape()`; require the same state after rejection. Add `fill` accessor cases to prove the getter is not invoked.

- [ ] **Step 3: Extend preset internal exact tests**

Require `normalizePresetShape()` to default `fill` to `{ kind: 'none' }`, detach explicit fill, and allow only the new `fill` option key. Update the exact default renderer assertion to remain byte-identical. Add exact sRGB/scheme/explicit-zero render assertions and keep fill immediately after geometry and before `a:ln`.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape fill' \
  packages/sdk/src/index.test.ts -t 'preset shape fill' --reporter=dot
```

Expected: FAIL because `ShapeFill` and `AddShapeOptions.fill` are absent.

- [ ] **Step 5: Implement the public contract and creation renderer**

In `preset-shape.ts` add the exact union from the design and `fill?: ShapeFill`. In `preset-shape.internal.ts` add `fill` to `OPTION_KEYS` and `NormalizedPresetShape`, normalize with:

```ts
const fill = normalizeSimpleFill(values.fill, 'Preset shape fill') ?? { kind: 'none' };
```

Store the detached value in the frozen normalized record. Replace the hard-coded `<a:noFill/>` in `renderPresetShapeXml()` with `renderSimpleFill(shape.fill, 'a:')`; no other shape skeleton bytes change.

- [ ] **Step 6: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/simple-fill.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts -t 'preset shape|rollback' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 7: Review, commit, push, and verify**

Review public type scope, validation-before-mutation, default-byte compatibility, detached values, exact OOXML order, and rollback. Then:

```sh
git add -- packages/model/src/preset-shape.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create shape solid fills"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Read, replace, and clear direct shape fill

**Files:**
- Create: `packages/model/src/shape-fill.internal.ts`
- Create: `packages/model/src/shape-fill.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: public `ShapeFill`, Task 1 simple-fill helpers, `LosslessXmlDocument`, `XmlElement`, and `ModelParseError`.
- Produces: `readShapeFill()`, `replaceShapeFill()`, `SlideModel.getShapeFill()`/`setShapeFill()`, and `ShapeModel.fill`.

- [ ] **Step 1: Add strict container read tests**

Build `p:sp` fixtures with canonical and alternate prefixes. Require detached snapshots for unique direct no-fill and solid sRGB/scheme/alpha. Require `undefined` for absent fill, grad/blip/pattern/group fill, nested lookalike, wrong namespace/prefix, missing/repeated direct `spPr`, repeated fill choices, malformed no-fill/solid/color/alpha, and supported fill nested under line/effects/text.

- [ ] **Step 2: Add replace/clear/order/isolation tests**

Require:

```ts
replaceShapeFill(xml, shape, sameFill, partUri) === false;
replaceShapeFill(xml, shape, { kind: 'none' }, partUri) === true;
replaceShapeFill(xml, shape, undefined, partUri) === true;
```

Test replacement of unique no/solid/grad/blip/pattern/group choices, clear of each unique choice, and insertion into absent fill after unique direct `prstGeom` or `custGeom`. Preserve transform, geometry/adjustments, line/effects/3D/extLst, text/non-visual siblings, neighbor shapes, namespace declarations, and unknown bytes. Repeated fill choices, repeated/missing `spPr`, and unsafe absent-fill insertion must throw before patching.

- [ ] **Step 3: Add model/SDK identity, lifecycle, and rollback tests**

Require immediate `ShapeModel.fill` on existing/new preset shapes and existing text-box `ShapeModel` instances; same-value byte/journal no-op; solid→none→clear→solid; unique gradient replacement and clear; duplicate independence; stable model object; outer rollback; six-format write/reopen; and invalid setters that leave part/relationships/parts/journal unchanged. Do not add fill creation options to `addText()` in this task.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/shape-fill.internal.test.ts \
  packages/model/src/model.test.ts -t 'shape fill' \
  packages/sdk/src/index.test.ts -t 'shape fill' --reporter=dot
```

Expected: FAIL because the shape fill adapter and model property are absent.

- [ ] **Step 5: Implement the strict shape container adapter**

Expose only internal functions:

```ts
export function readShapeFill(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): ShapeFill | undefined;

export function replaceShapeFill(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  fill: ShapeFill | undefined,
  partUri: string,
): boolean;
```

Navigate direct `p:spPr` and DrawingML children by namespace URI. On insertion, derive the in-scope prefix from the unique direct geometry and place the encoded fill at `geometry.end`; do not serialize or rebuild the full shape. Use shared equality for no-op and clone only the local properties element if needed to calculate safe edits.

- [ ] **Step 6: Add slide-owned API and public model property**

Add:

```ts
getShapeFill(id: number): ShapeFill | undefined;
setShapeFill(id: number, value: ShapeFill | undefined): void;
```

Normalize outside the package transaction when value is not `undefined`; inside, resolve the current shape, call the strict replacer, and call `setXml()` only when it returns true. Add only to `ShapeModel`:

```ts
get fill(): ShapeFill | undefined {
  return this.slide.getShapeFill(this.id);
}

set fill(value: ShapeFill | undefined) {
  this.slide.setShapeFill(this.id, value);
}
```

Do not add this property to image/table/chart/group models.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/simple-fill.internal.test.ts \
  packages/model/src/shape-fill.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  -t 'shape fill|preset shape|duplicate slide|rollback' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 8: Review, commit, push, and verify**

Review direct namespace ownership, unsupported-state reads, insertion order, unique unsupported replacement, no-op/journal behavior, identity, duplicate independence, and rollback. Then:

```sh
git add -- packages/model/src/shape-fill.internal.ts \
  packages/model/src/shape-fill.internal.test.ts packages/model/src/slide.ts \
  packages/model/src/shapes.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit shape solid fills"
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
- Consumes: public `PptxGenJS.ShapeType`, `slide.addShape()`, `write()`, native `ShapeFill`, creation, read, and edit APIs.
- Produces: public-only evidence for supported fill semantics and explicit evidence for strict divergences.

- [ ] **Step 1: Add a public-output fill fixture matrix**

Generate one shape per case using only public PptxGenJS APIs: omitted, `{ type: 'none' }`, sRGB, scheme, 50% transparency, explicit zero transparency, deprecated alpha, empty object, and solid missing color. Capture public-output packages without importing PptxGenJS internals.

- [ ] **Step 2: Compare supported semantic state**

For omitted, sRGB, scheme, and 50% transparency, open with `PptxDocument`, compare shape name/geometry/transform and final simple fill snapshot or direct state to native equivalents. Verify native edit of imported valid PptxGenJS solid fill and write/reopen.

- [ ] **Step 3: Assert strict divergences rather than hiding them**

Require PptxGenJS explicit none to contain no direct fill choice while native none contains direct `a:noFill`; effective rendering is no-fill in both. Require PptxGenJS explicit zero to omit alpha while native writes `100000`; effective opacity is identical. Require PptxGenJS empty/missing-color output black while native rejects before mutation. Require native rejection of deprecated `alpha` and adapter preservation/read of its final generated alpha.

- [ ] **Step 4: Run conformance and dependency gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'shape fill public output' --reporter=dot
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; `pptxgenjs` remains absent from non-adapter manifests.

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only evidence, supported semantic comparison, strict divergence assertions, output warnings isolation, and dependency boundaries. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare shape fills with pptxgenjs"
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
- Produces: tarball-level `shapeFills: true` evidence in the smoke JSON.

- [ ] **Step 1: Extend the installed-package Node smoke**

In the generated `smoke.mjs`, create sRGB/scheme/none shapes, mutate the original caller objects, read detached `ShapeModel.fill`, perform same-value/replace/clear edits, duplicate, write/reopen, and require zero relationship/part growth. Add `shapeFills` to the final JSON envelope without removing existing fields.

- [ ] **Step 2: Extend browser and declaration smoke**

In the generated browser smoke, import `PptxDocument` from the package browser condition, create a filled shape, read/edit its fill, write/open, and verify no `node:` dependency. In the generated typecheck source, import `ShapeFill`, assign valid none/solid values, use `AddShapeOptions.fill`, and require invalid kind/color/transparency cases with `@ts-expect-error`.

- [ ] **Step 3: Build, pack, and run the real tarball smoke**

```sh
pnpm build
mkdir -p /tmp/pptx-shape-fill-pack
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-shape-fill-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-shape-fill-pack/jiayunxie-pptx-0.1.0.tgz
```

Expected JSON includes `"shapeFills":true`, existing `"presetShapes":true`, `"types":true`, and CLI version `0.1.0`.

- [ ] **Step 4: Run script and repository gates**

```sh
node --check scripts/smoke-npm-package.mjs
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 5: Review, commit, push, and verify**

Review use of installed tarball only, Node/browser parity, public type coverage, existing smoke preservation, temporary cleanup, and absence of internal imports. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed shape fills"
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

Document native creation and editing with exact public syntax:

```ts
const shape = slide.addShape('roundRect', {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 20,
  },
});

shape.fill = { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } };
shape.fill = { kind: 'none' };
shape.fill = undefined;
```

Explain that `undefined` clears direct state, none writes direct no-fill, and gradient remains a separate API.

- [ ] **Step 2: Update compatibility and changelog**

Change the `slide.addShape()` row from geometry-only wording to geometry plus solid/no-fill creation/read/edit/clear. Keep line/dash/arrows, shadow, hyperlink, adjustments, custom geometry, shape text, percentage positions, and advanced fills explicitly remaining. Record PptxGenJS empty/missing-color, explicit-none, zero-transparency, and deprecated-alpha divergences without claiming byte identity.

- [ ] **Step 3: Validate examples and documentation consistency**

```sh
rg -n "ShapeFill|shape\.fill|solid/no-fill|line/dash|advanced fill" \
  README.md packages/pptx/README.md CHANGELOG.md \
  docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
git diff --check
```

Expected: all examples match the exported API and all remaining gaps stay visible.

- [ ] **Step 4: Review, commit, push, and verify**

Review API spelling, semantics of none versus clear, strict divergences, no overclaiming, and parity between root/package docs. Then:

```sh
git add -- README.md packages/pptx/README.md CHANGELOG.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document shape solid fills"
git push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run release, real-PPTX, structural, visual, and import QA

**Files:**
- No repository file changes expected; use a conversation-specific `/tmp/pptx-shape-fill-qa-20260731` workspace.

**Interfaces:**
- Consumes: Tasks 1–6 complete repository state.
- Produces: final gate evidence for package integrity, mutation isolation, real-office rendering, visual correctness, and remote parity.

- [ ] **Step 1: Run complete repository gates**

```sh
pnpm check
pnpm test:performance
pnpm build
```

Expected: strict typecheck/build, all tests, and the independent performance budget pass.

- [ ] **Step 2: Pack and smoke the final tarball**

```sh
mkdir -p /tmp/pptx-shape-fill-qa-20260731/pack
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-shape-fill-qa-20260731/pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-fill-qa-20260731/pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: identity/version, Node/browser/types, preset shapes, shape fills, dependency boundaries, and CLI smoke all pass.

- [ ] **Step 3: Generate representative lifecycle decks**

Create native decks for omitted/none/sRGB/scheme/0/25/50/100 transparency, alternate geometries, existing fill read/edit/clear, gradient→simple replacement, duplicate, reopen, second write, plus PptxGenJS legal comparison and strict-divergence fixtures. Include labels and visible line outlines so none/clear cases are inspectable.

- [ ] **Step 4: Validate every legal package and mutation boundary**

For each legal deck run:

```sh
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json slides list deck.pptx
pptx-inspect --json part read deck.pptx /ppt/slides/slide1.xml
```

Use `package diff` to prove source→edit changes only the target slide part, and edit→reopen→second-write has zero part changes. Verify exact direct fill choices, color, alpha, shape IDs, order, names, and neighbor bytes.

- [ ] **Step 5: Render and inspect every slide**

Open/export all legal decks with an isolated LibreOffice profile, run `pdfinfo`, render PDF pages with `pdftoppm`, then run `render_slides.py` and `slides_test.py` for every PPTX. Inspect every rendered page at full size for solid color, theme resolution, transparency ordering, none/clear visibility, clipping, overlap, and label readability.

- [ ] **Step 6: Verify artifact-tool import behavior**

Import native representative and lifecycle decks with `PresentationFile.importPptx()`, require expected slide/shape counts, and inspect named shapes. Import PptxGenJS comparison fixtures and confirm supported files load without mutation.

- [ ] **Step 7: Verify the repository and remote are final**

```sh
git diff --check
git status --short
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected: only the pre-existing untracked `.pnpm-store/` is ignored, and divergence is `0 0`. Task 7 creates no empty commit when the repository has no changes.
