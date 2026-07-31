# Preset Shape Creation Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking, and every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add a complete native preset-shape vertical slice that can create, read, replace, duplicate, write, reopen, package, and visually validate all 178 canonical OOXML preset shape types represented by PptxGenJS 4.0.1.

**Architecture:** Keep the frozen public token/type/options contract in `preset-shape.ts` and put descriptor-safe normalization plus namespace-aware OOXML read/replace/render behavior in `preset-shape.internal.ts`. `SlideModel` owns transaction, shape-ID allocation, shape-tree insertion, persistence, and live-model resolution; `ShapeModel` delegates preset geometry access back through the slide so stable identity and rollback semantics remain centralized.

**Tech Stack:** TypeScript strict mode, Vitest, lossless source-span OOXML editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, pnpm pack, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- The canonical public set contains exactly 178 unique OOXML `ST_ShapeType` strings in PptxGenJS declaration order with invalid `folderCorner` replaced in place by valid `foldedCorner`; it excludes `folderCorner` and `custGeom`.
- `PRESET_SHAPE_TYPES` is frozen at runtime; `PresetShapeType` is derived from its tuple; `AddShapeOptions` contains only `name` plus `Partial<Transform>`.
- Creation defaults are x/y/width/height = 914400 EMU, rotation = 0, both flips = false, name = `Shape ${id}`, direct `a:noFill`, and an empty direct `a:ln`.
- Inputs accept only ordinary or null-prototype objects with own data properties; reject arrays, class instances, symbols, accessors, unknown keys, invalid XML strings, non-finite numbers, invalid extents/rotation, and non-boolean flips before package mutation.
- Read only the unique direct `p:spPr/a:prstGeom` with one unqualified canonical `prst`; missing, unknown, qualified-lookalike, duplicate, or ambiguous geometry reads as `undefined`.
- Replacing the current canonical value is an exact no-op; changing it whole-replaces only the unique direct geometry with canonical `a:prstGeom + a:avLst`, clearing stale adjustments while preserving every unrelated byte.
- Do not implement fill configuration, line styling, shadow, hyperlink, adjustment handles, custom geometry, shape text, or a PptxGenJS namespace-compatible `ShapeType` facade in this slice.
- Every failure and outer transaction rollback preserves exact slide bytes, relationships, parts, mutation journal, shape order, and live model identity.
- Node 20+ and the browser bundle expose the same behavior from the single `@jiayunxie/pptx` package.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed over SSH port 443, fetched, and verified at divergence `0 0`.
- The user has authorized inline autonomous execution; do not dispatch subagents or pause for routine choices.

---

### Task 1: Add the frozen token catalog, strict normalizer, and deterministic codec

**Files:**
- Create: `packages/model/src/preset-shape.ts`
- Create: `packages/model/src/preset-shape.internal.ts`
- Create: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/index.ts`

**Interfaces:**
- Consumes: `Transform`, `inches()`, `LosslessXmlDocument`, `XmlElement`, XML attribute escaping, and `ModelParseError`.
- Produces: `PRESET_SHAPE_TYPES`, `PresetShapeType`, `AddShapeOptions`, `NormalizedPresetShape`, `normalizePresetShape()`, `renderPresetShapeXml()`, `readPresetShapeType()`, and `replacePresetShapeType()` for Tasks 2 and 3.

- [ ] **Step 1: Write the public catalog and compile-time contract test**

Define the catalog as an `as const` tuple and freeze the actual exported array. The tuple is this exact ordered sequence:

```ts
export const PRESET_SHAPE_TYPES = Object.freeze([
  'accentBorderCallout1', 'accentBorderCallout2', 'accentBorderCallout3',
  'accentCallout1', 'accentCallout2', 'accentCallout3',
  'actionButtonBackPrevious', 'actionButtonBeginning', 'actionButtonBlank',
  'actionButtonDocument', 'actionButtonEnd', 'actionButtonForwardNext',
  'actionButtonHelp', 'actionButtonHome', 'actionButtonInformation',
  'actionButtonMovie', 'actionButtonReturn', 'actionButtonSound',
  'arc', 'bentArrow', 'bentUpArrow', 'bevel', 'blockArc',
  'borderCallout1', 'borderCallout2', 'borderCallout3', 'bracePair',
  'bracketPair', 'callout1', 'callout2', 'callout3', 'can', 'chartPlus',
  'chartStar', 'chartX', 'chevron', 'chord', 'circularArrow', 'cloud',
  'cloudCallout', 'corner', 'cornerTabs', 'cube', 'curvedDownArrow',
  'curvedLeftArrow', 'curvedRightArrow', 'curvedUpArrow', 'decagon',
  'diagStripe', 'diamond', 'dodecagon', 'donut', 'doubleWave',
  'downArrow', 'downArrowCallout', 'ellipse', 'ellipseRibbon',
  'ellipseRibbon2', 'flowChartAlternateProcess', 'flowChartCollate',
  'flowChartConnector', 'flowChartDecision', 'flowChartDelay',
  'flowChartDisplay', 'flowChartDocument', 'flowChartExtract',
  'flowChartInputOutput', 'flowChartInternalStorage',
  'flowChartMagneticDisk', 'flowChartMagneticDrum',
  'flowChartMagneticTape', 'flowChartManualInput',
  'flowChartManualOperation', 'flowChartMerge', 'flowChartMultidocument',
  'flowChartOfflineStorage', 'flowChartOffpageConnector',
  'flowChartOnlineStorage', 'flowChartOr', 'flowChartPredefinedProcess',
  'flowChartPreparation', 'flowChartProcess', 'flowChartPunchedCard',
  'flowChartPunchedTape', 'flowChartSort', 'flowChartSummingJunction',
  'flowChartTerminator', 'foldedCorner', 'frame', 'funnel', 'gear6',
  'gear9', 'halfFrame', 'heart', 'heptagon', 'hexagon', 'homePlate',
  'horizontalScroll', 'irregularSeal1', 'irregularSeal2', 'leftArrow',
  'leftArrowCallout', 'leftBrace', 'leftBracket', 'leftCircularArrow',
  'leftRightArrow', 'leftRightArrowCallout', 'leftRightCircularArrow',
  'leftRightRibbon', 'leftRightUpArrow', 'leftUpArrow', 'lightningBolt',
  'line', 'lineInv', 'mathDivide', 'mathEqual', 'mathMinus',
  'mathMultiply', 'mathNotEqual', 'mathPlus', 'moon', 'noSmoking',
  'nonIsoscelesTrapezoid', 'notchedRightArrow', 'octagon',
  'parallelogram', 'pentagon', 'pie', 'pieWedge', 'plaque',
  'plaqueTabs', 'plus', 'quadArrow', 'quadArrowCallout', 'rect',
  'ribbon', 'ribbon2', 'rightArrow', 'rightArrowCallout', 'rightBrace',
  'rightBracket', 'round1Rect', 'round2DiagRect', 'round2SameRect',
  'roundRect', 'rtTriangle', 'smileyFace', 'snip1Rect', 'snip2DiagRect',
  'snip2SameRect', 'snipRoundRect', 'squareTabs', 'star10', 'star12',
  'star16', 'star24', 'star32', 'star4', 'star5', 'star6', 'star7',
  'star8', 'stripedRightArrow', 'sun', 'swooshArrow', 'teardrop',
  'trapezoid', 'triangle', 'upArrow', 'upArrowCallout', 'upDownArrow',
  'upDownArrowCallout', 'uturnArrow', 'verticalScroll', 'wave',
  'wedgeEllipseCallout', 'wedgeRectCallout', 'wedgeRoundRectCallout',
] as const);

export type PresetShapeType = (typeof PRESET_SHAPE_TYPES)[number];

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
}
```

In `preset-shape.internal.test.ts`, require length 178, uniqueness 178, `Object.isFrozen(...) === true`, canonical `foldedCorner`, absent `folderCorner/custGeom`, stable first/last tokens, and a rejected runtime `push` without changing the array.

- [ ] **Step 2: Write strict normalization tests**

Require exact defaults and quantization:

```ts
expect(normalizePresetShape('rect', undefined)).toEqual({
  type: 'rect', name: undefined,
  x: 914400, y: 914400, width: 914400, height: 914400,
  rotation: 0, flipHorizontal: false, flipVertical: false,
});
expect(normalizePresetShape('ellipse', {
  name: 'A & <B>', x: 1.4, y: -1.5, width: 2.6, height: 3.5,
  rotation: 2.5, flipHorizontal: true, flipVertical: true,
})).toMatchObject({ x: 1, y: -1, width: 3, height: 4, rotation: 3 });
```

Reject non-string/unknown types, `folderCorner`, `custGeom`, null/primitive options, arrays, dates, class instances, inherited known/unknown fields, own unknown string fields, all symbol keys, own getters/setters without invoking them, non-string/illegal-XML names, NaN/infinities, values that round outside safe integers, rounded width/height <= 0, rotation outside `-21600000..21600000`, and truthy/falsy non-booleans. Accept an ordinary object and `Object.create(null)` with own data properties only; prove the normalized result is detached from later caller mutation.

- [ ] **Step 3: Write render/read/replace tests**

For every catalog token render ID 7 and parse the result. Require one `p:sp`, direct `p:nvSpPr`, `p:spPr`, one direct DrawingML `a:xfrm`, one direct `a:prstGeom@prst`, one direct empty `a:avLst`, direct `a:noFill`, direct empty `a:ln`, no text body, relationship, extension, or timestamp.

Use exact assertions for defaults, custom name escaping, negative coordinates, rotation, and flips:

```ts
expect(renderPresetShapeXml(2, normalizePresetShape('rect', undefined))).toBe(
  '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<p:nvSpPr><p:cNvPr id="2" name="Shape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="914400" y="914400"/>' +
  '<a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr></p:sp>',
);
```

Create fixtures with alternate valid prefixes and require canonical reads. Require `undefined` without source changes for wrong shape namespace/name, missing/repeated direct `spPr`, missing/repeated direct `prstGeom`, absent/repeated/qualified-only `prst`, unknown token, nested lookalike, and malformed geometry. For replace, require same-value exact no-op preserving custom `avLst`; different-value whole-replace preserving transform/fill/line/effects/text/non-visual/ext siblings; unsafe states throw `ModelParseError` before any patch.

- [ ] **Step 4: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts --reporter=dot
```

Expected: FAIL because the new modules do not exist.

- [ ] **Step 5: Implement the minimal public contract and strict normalizer**

Use property descriptors before reading values. Determine the prototype with `Object.getPrototypeOf`; accept only `Object.prototype` or `null`. Enumerate `Reflect.ownKeys`, reject symbols and keys outside the exact option set, reject any descriptor lacking `value`, then normalize primitives into a frozen internal record. Use `Math.round`, `Number.isSafeInteger`, the specified rotation interval, and the existing invalid XML 1.0 control range.

- [ ] **Step 6: Implement deterministic rendering and direct geometry editing**

Render only normalized primitives, omit zero/false transform attributes, and use `escapeXmlAttribute()` for name/type. Navigate through direct, namespace-resolved children; never use a descendant-only guess for `spPr` or `prstGeom`. `replacePresetShapeType()` returns `false` for the same value, otherwise replaces exactly the geometry element span with `<a:prstGeom prst="TYPE"><a:avLst/></a:prstGeom>` using the geometry element's in-scope DrawingML prefix.

- [ ] **Step 7: Export the public symbols and run focused gates**

Add `export * from './preset-shape.js';` to `packages/model/src/index.ts`, then run:

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass and no internal helper is publicly exported.

- [ ] **Step 8: Review, commit, push, and verify**

Review token count/spelling/order, runtime freezing, descriptor safety, default units, numeric boundaries, exact XML, namespace/direct ownership, same-value no-op, unknown-byte preservation, and public export scope. Then:

```sh
git add -- packages/model/src/preset-shape.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts packages/model/src/index.ts
git diff --cached --check
git commit -m "feat: add preset shape codec"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Create preset shapes through `SlideModel.addShape()`

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `PresetShapeType`, `AddShapeOptions`, `normalizePresetShape()`, and `renderPresetShapeXml()`; existing `allocateShapeId()`, `ShapeModel`, OPC transactions, and shape-model cache.
- Produces: `SlideModel.addShape(type, options?)` returning the unique live `ShapeModel` for the allocated ID.

- [ ] **Step 1: Add focused creation and full-token tests**

In `packages/sdk/src/index.test.ts`, create default, named/custom-transform, line, inverse line, ellipse, flowchart, star, action-button, and folded-corner shapes. Require returned instances to have `kind === 'shape'`, IDs `[2..]`, stable `name`, exact native transform units, no text body, zero relationships, source order, and the same object identity from repeated `slide.shapes` reads.

Create one shape for every `PRESET_SHAPE_TYPES` entry on a blank slide; require 178 returned models, unique IDs/names, catalog order, and a parseable direct geometry token matching the request.

- [ ] **Step 2: Add insertion, invalid-state, and zero-mutation tests**

Insert an opaque direct `p:extLst` under the shape tree, create two shapes, and require both before `extLst` with the opaque bytes unchanged. Reject a missing direct shape tree, repeated direct shape trees, repeated direct `extLst`, exhausted/unsafe/non-integer/duplicate shape IDs, and every invalid input from Task 1. Snapshot part bytes, relationships, parts, and journal before each failure and require equality afterward.

- [ ] **Step 3: Add lifecycle and outer rollback tests**

Create representative shapes on visible/hidden slides in sections, duplicate the slide, move source and duplicate, delete another slide, write/reopen, and require geometry/type/name/transform/order unchanged. In an outer `document.transaction()`, create a shape and throw; require exact package/journal restoration, absence from `slide.shapes`, and the rolled-back returned model to throw `ModelParseError` on access.

Repeat write/reopen for `pptx/pptm/ppsx/ppsm/potx/potm` and require identical semantic snapshots.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/model.test.ts -t 'preset shape' \
  packages/sdk/src/index.test.ts -t 'preset shape' --reporter=dot
```

Expected: FAIL because `SlideModel.addShape()` is absent.

- [ ] **Step 5: Implement the transactional API**

Add imports and this exact public signature:

```ts
addShape(type: PresetShapeType, options?: AddShapeOptions): ShapeModel;
```

Inside one package transaction: normalize first; parse the slide; require exactly one direct `p:sld/p:cSld/p:spTree`; require at most one direct `p:extLst`; allocate the next safe shape ID; render; insert immediately before `extLst` or append; save once; resolve through `this.shapes`; and require an instance of `ShapeModel` with `kind === 'shape'` and the same ID. Throwing at any stage rolls back.

- [ ] **Step 6: Harden shared shape-tree/ID helpers only as required**

Reuse one strict helper for table and shape creation only if its existing behavior already matches the new ownership contract. `allocateShapeId()` must reject malformed direct `cNvPr@id` states rather than silently skipping them, must never return above `Number.MAX_SAFE_INTEGER`, and must not renumber existing shapes. Do not refactor text creation or unrelated slide code.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape|duplicate slide|delete slide' \
  packages/sdk/src/index.test.ts -t 'preset shape|section|hidden|rollback' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 8: Review, commit, push, and verify**

Review validation-before-mutation, shape-tree ownership, safe ID allocation, `extLst` ordering, exact defaults, live identity, six formats, lifecycle preservation, and rollback. Then:

```sh
git add -- packages/model/src/slide.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create preset shapes"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Read and replace preset geometry through `ShapeModel.presetType`

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `readPresetShapeType()`/`replacePresetShapeType()` and Task 2-created live shapes.
- Produces: `ShapeModel.presetType: PresetShapeType | undefined` with strict transactional replacement.

- [ ] **Step 1: Add read tests for existing and newly created shapes**

Require correct reads for every catalog token, alternate valid prefixes, existing text boxes with direct preset geometry, and newly created shapes. Require `undefined` without mutation for missing/unknown/qualified-lookalike/repeated/ambiguous geometry. Confirm malformed PptxGenJS `folderCorner` remains losslessly openable and unrelated edits preserve its bytes while `presetType` is `undefined`.

- [ ] **Step 2: Add replace/no-op/isolation tests**

Set `rect → ellipse` and require only the direct geometry element to change. Preserve name, ID, transform, `a:noFill`, `a:ln`, effects, text body, non-visual properties, extensions, relationships, neighboring shapes, and unknown package parts. Require `ellipse → ellipse` to preserve exact slide bytes/journal including non-empty `a:avLst`. Require changed type to reset old adjustments to one empty direct `a:avLst`.

- [ ] **Step 3: Add invalid, lifecycle, identity, and rollback tests**

Reject runtime `folderCorner`, `custGeom`, unknown/non-string values, and unsafe existing geometry before mutation. Replace on source and duplicate independently; move/write/reopen without semantic drift. Require the same cached `ShapeModel` before and after replacement. Roll an edit back in an outer transaction and require exact bytes/journal plus the original `presetType` on the same model object.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/model.test.ts -t 'preset type' \
  packages/sdk/src/index.test.ts -t 'preset type' --reporter=dot
```

Expected: FAIL because `ShapeModel.presetType` is absent.

- [ ] **Step 5: Add slide-owned read and mutation methods**

Add:

```ts
getShapePresetType(id: number): PresetShapeType | undefined;
setShapePresetType(id: number, value: PresetShapeType): void;
```

The getter resolves the current XML element and calls the strict reader. The setter first normalizes the token without touching the package, then runs in a package transaction, resolves the unique shape, invokes the strict replacer, and saves only when it returns `true`.

- [ ] **Step 6: Expose `ShapeModel.presetType`**

Add only:

```ts
get presetType(): PresetShapeType | undefined {
  return this.slide.getShapePresetType(this.id);
}

set presetType(value: PresetShapeType) {
  this.slide.setShapePresetType(this.id, value);
}
```

Do not add this property to `BaseShapeModel`, `ImageModel`, `TableModel`, or `ChartModel`.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape|preset type|duplicate slide' \
  packages/sdk/src/index.test.ts -t 'preset shape|preset type|rollback' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 8: Review, commit, push, and verify**

Review strict direct reads, same-value byte no-op, adjustment clearing, isolation, malformed-source preservation, duplicate independence, live identity, and rollback. Then:

```sh
git add -- packages/model/src/slide.ts packages/model/src/shapes.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit preset shape geometry"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: actual public `PptxGenJS.ShapeType`, `slide.addShape()`, `write()` output, and native create/read APIs.
- Produces: public-output evidence for valid shape skeletons and the documented `folderCorner` upstream defect.

- [ ] **Step 1: Add a public-output fixture extractor**

Use only `new PptxGenJS()`, `pptx.ShapeType`, `slide.addShape()`, and `write({ outputType: 'nodebuffer' })`. Open the result through `PptxDocument.open()` and the package API; do not import PptxGenJS internals or copy its generated XML constants.

- [ ] **Step 2: Compare representative valid outputs**

Cover omitted options, `{}`, `rect`, `ellipse`, `line`, `lineInv`, `flowChartDecision`, `star5`, `actionButtonHome`, and object name plus x/y/w/h/45-degree rotate/flips. Convert PptxGenJS inch/degrees inputs to expected native EMU/OOXML angles, then compare final public `name`, `transform`, `presetType`, direct no-fill state, and empty line state. Do not require byte-identical packages or copy PptxGenJS metadata.

- [ ] **Step 3: Prove full legal-token and defect boundaries**

For each `PRESET_SHAPE_TYPES` token except `foldedCorner`, index public `pptx.ShapeType[token]`, generate a PptxGenJS shape, open it natively, and require `presetType === token`. Generate PptxGenJS `folderCorner` and require its raw direct token to be `folderCorner`, native `presetType` to be `undefined`, and the Open XML/PowerPoint validator evidence to classify it as invalid. Generate native `foldedCorner`, require successful read/reopen/validation, and require native creation to reject `folderCorner`.

- [ ] **Step 4: Run the conformance test and verify its initial result**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'preset shape public output' --reporter=dot
```

Expected before adding the test: no matching conformance test; after adding it: PASS against Tasks 1–3.

- [ ] **Step 5: Run adapter and dependency gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'keeps pptxgenjs out of every non-adapter package dependency list' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; `pptxgenjs` remains absent from non-adapter manifests.

- [ ] **Step 6: Review, commit, push, and verify**

Review public-only evidence, legal-token matrix, unit conversion, semantic rather than byte comparison, defect isolation, and dependency boundaries. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare preset shapes with pptxgenjs"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball public API.
- Produces: runtime/type proof and JSON check `presetShapes: true`.

- [ ] **Step 1: Add packed Node runtime coverage**

In the generated consumer, assert a frozen 178-token catalog, create representative/default/custom shapes, read/replace `presetType`, duplicate/edit independently, write/reopen, and inspect only public state. Require exact transforms/types/names/order and no relationship/part growth for geometry-only operations.

- [ ] **Step 2: Add browser bundle runtime coverage**

Use this public flow in the Playwright bundle:

```js
if (PRESET_SHAPE_TYPES.length !== 178 || !Object.isFrozen(PRESET_SHAPE_TYPES)) {
  throw new Error('Browser preset catalog failed');
}
const browserShapeDeck = PptxDocument.create();
const browserShape = browserShapeDeck.addSlide().addShape('foldedCorner');
browserShape.presetType = 'star5';
const reopenedBrowserShape = await PptxDocument.open(await browserShapeDeck.writeBlob());
if (reopenedBrowserShape.slides[0]?.shapes[0]?.presetType !== 'star5') {
  throw new Error('Browser preset shape failed');
}
```

- [ ] **Step 3: Add declaration coverage**

```ts
const typedPreset: PresetShapeType = 'foldedCorner';
const typedShapeOptions: AddShapeOptions = {
  x: inches(1), y: inches(2), width: inches(3), height: inches(4),
  rotation: degrees(45), flipHorizontal: true, name: 'Typed shape',
};
const typedShape: ShapeModel = createdDocument.addSlide().addShape(typedPreset, typedShapeOptions);
const typedRead: PresetShapeType | undefined = typedShape.presetType;
typedShape.presetType = 'rect';
```

Add all symbols to the existing no-unused aggregate and compile negative fixtures for `folderCorner`, `custGeom`, unknown option fields, and invalid transform value types.

- [ ] **Step 4: Build, pack, and run the actual tarball**

```sh
pnpm build
preset_shape_pack_dir=$(mktemp -d /tmp/pptx-preset-shape-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$preset_shape_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$preset_shape_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require JSON containing `"presetShapes":true`, `"types":true`, and `"cli":"0.1.0"`.

- [ ] **Step 5: Review, commit, push, and verify**

Review actual tarball use, Node/browser parity, catalog freezing, declarations, negative compile cases, no workspace import leakage, and unchanged CLI behavior. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed preset shapes"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Document preset-shape creation, editing, and parity boundaries

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: complete native behavior, conformance evidence, and packed public surface.
- Produces: discoverable usage, precise compatibility status, strict divergences, and remaining shape gaps.

- [ ] **Step 1: Update the parity matrix and focused baseline**

Add:

```md
| `slide.addShape(ShapeType, options)` | `SlideModel.addShape(type, options)` / `ShapeModel.presetType` | 已支持 178 个 canonical preset 的 create/read/replace/duplicate/reopen；native transform 使用 EMU/OOXML angle |
```

Explain the frozen catalog, default 1-inch geometry, strict option descriptors, stable identity, same-value no-op, adjustment clearing on type change, and public-output comparison boundary. Record PptxGenJS `folderCorner` as an upstream invalid-token defect and native `foldedCorner` as the canonical supported spelling. Keep fill/line/shadow/hyperlink/adjustments/custom geometry/shape text explicitly pending.

- [ ] **Step 2: Update API README, npm README, and changelog**

Use this public example:

```ts
const slide = document.addSlide();
const shape = slide.addShape('roundRect', {
  x: inches(1), y: inches(1), width: inches(3), height: inches(2),
  rotation: degrees(15), name: 'Feature card',
});
shape.presetType = 'hexagon';
console.log(shape.presetType); // 'hexagon'
```

Document that `PRESET_SHAPE_TYPES` drives discovery and is frozen. Add one changelog bullet without claiming style/text/custom-geometry parity.

- [ ] **Step 3: Run documentation and type gates**

```sh
rg -n 'addShape|presetType|PRESET_SHAPE_TYPES|foldedCorner|folderCorner' \
  CHANGELOG.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
rg -n 'preset shape.*(pending|unsupported|尚未支持|仍待)' \
  packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
git diff --check
```

The second search may only find the explicit style/text/adjustment/custom-geometry non-goals, not basic preset creation/read/replace.

- [ ] **Step 4: Review, commit, push, and verify**

Review API names, native units, catalog count, defect wording, no overclaim, removed stale gap text, remaining gaps, migration example, and changelog scope. Then:

```sh
git add -- CHANGELOG.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document preset shapes"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run full release, package, validator, and visual QA

**Files:**
- No repository changes expected; generate only under `/tmp/pptx-preset-shape-qa-20260731` and a `mktemp` pack directory.

**Interfaces:**
- Consumes: the complete preset-shape slice and installed `pptx-inspect`/LibreOffice/Presentations QA tools.
- Produces: release evidence, full-token validity proof, mutation-isolation proof, render/open proof, and a clean synced repository.

- [ ] **Step 1: Run complete repository and package gates**

```sh
pnpm check
pnpm test:performance
pnpm build
preset_shape_qa_pack_dir=$(mktemp -d /tmp/pptx-preset-shape-qa-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$preset_shape_qa_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$preset_shape_qa_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require all tests/type gates/builds, the 5-second performance budget, and packed JSON checks to pass.

- [ ] **Step 2: Generate representative and full-token public-API decks**

Under `/tmp/pptx-preset-shape-qa-20260731`, generate: native defaults; native transform/flip/name; native representative categories; native full-token pages arranged in a labeled grid; source/edited/reopened/second-write lifecycle decks; native `foldedCorner`; PptxGenJS representative legal shapes; and PptxGenJS malformed `folderCorner`. To make geometry visible without expanding this API, apply the existing `gradientFill` editor after creation. Keep generator/verifier `.mjs` files under `/tmp` and create them with `apply_patch`.

- [ ] **Step 3: Validate OOXML and mutation isolation with `pptx-inspect`**

Use the `pptx-inspect` skill's dry-run/read-first workflow. For every legal deck run `pptx-inspect --json package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and exact slide-part reads. Require 0 errors/warnings; exact catalog/token/name/transform/no-fill/line state before visible-fill editing; unique IDs; no orphan parts/relationships; and one canonical direct geometry per shape.

Run package diffs:

```sh
pptx-inspect --json package diff native-source.pptx native-edited.pptx
pptx-inspect --json package diff native-edited.pptx native-reopened.pptx
pptx-inspect --json package diff native-edited.pptx native-second-write.pptx
```

The first diff may change only the target slide XML; the latter two require zero added/removed/changed parts. Validate `folderCorner` separately and retain its expected validator failure as upstream evidence, never as a passing native fixture.

- [ ] **Step 4: Open, render, and inspect every legal shape**

Use isolated LibreOffice profiles, `pdfinfo`, `pdftoppm`, Presentations `render_slides.py`, and `slides_test.py`. Require no repair output, expected page counts/sizes, no clipping/overflow/unintended overlaps, correct rotation/flip direction, and a visually distinct/rendered geometry for every labeled canonical token. Inspect every PNG individually at original detail; a contact sheet is only a navigation aid. Confirm `foldedCorner` renders successfully and document PptxGenJS `folderCorner` failure without delivering it as a valid deck.

- [ ] **Step 5: Verify artifact-tool import boundary**

Import the legal representative/full-token native decks through the installed presentation artifact tooling and require `foldedCorner` to parse. Attempt the PptxGenJS `folderCorner` fixture and require the known invalid-enum failure. This is compatibility evidence only; do not generate the product deck with artifact-tool or add repository dependencies.

- [ ] **Step 6: Final repository and remote audit**

```sh
git diff --check
git diff --cached --check
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git status --short --branch
git rev-list --left-right --count origin/main...HEAD
```

Expected: only untracked `.pnpm-store/`, divergence `0 0`. Do not create an empty QA commit.
