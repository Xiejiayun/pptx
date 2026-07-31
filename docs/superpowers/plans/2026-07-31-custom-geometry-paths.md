# Custom Geometry Paths Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox
> (`- [ ]`) syntax for tracking, and every repository-changing task ends with an
> independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add native typed custom path geometry that creates, reads, whole-replaces,
converts, duplicates, writes, reopens, packages, and visually validates every legal
PptxGenJS 4.0.1 custom-points command.

**Architecture:** Keep public path/command types in a focused `custom-geometry.ts`
module and descriptor-safe normalization plus namespace-aware OOXML ownership in
`custom-geometry.internal.ts`. Reuse the existing shape option/style renderer for
creation, delegate live reads and edits through `SlideModel`, and extend preset
replacement only for safely parsed custom geometry so stable identity and rollback
remain centralized.

**Tech Stack:** TypeScript strict mode, Vitest, source-span lossless OOXML editing,
OPC transactions, PptxGenJS 4.0.1 public output, pnpm pack, `pptx-inspect`,
LibreOffice, Poppler, and the Presentations render/overflow/artifact-tool workflow.

## Global Constraints

- Follow `docs/superpowers/specs/2026-07-31-custom-geometry-paths-design.md` exactly.
- Public coordinates/radii are safe-integer EMU; public angles are safe-integer
  OOXML angles. Native input never accepts inch/percentage/string shortcuts.
- Support `moveTo`, `lineTo`, `arcTo`, `quadraticBezierTo`, `cubicBezierTo`, and
  `close`, multiple subpaths, multiple paths, empty command lists, and direct path
  fill/stroke/extrusion flags.
- Custom geometry creation uses `SlideModel.addCustomShape()`; `addShape()` and
  `PRESET_SHAPE_TYPES` remain preset-only and continue rejecting `custGeom`.
- `AddCustomShapeOptions` includes the existing transform/name/fill/line/arrows/
  shadow/hyperlink surface and excludes preset-only `adjustments`.
- Read only one direct namespace-correct geometry choice. Non-empty guides,
  adjustment lists, handles, connections, custom text rectangles, coordinate
  formulas, malformed ownership, and mixed geometry return `undefined` and reject
  editing without mutation.
- Setter replacement is whole-geometry, same-value is exact bytes/journal no-op,
  and supported preset/custom conversion replaces only the geometry choice.
- Every public input is descriptor-safe, getter-free, detached, and deep-frozen
  before ID allocation, relationship creation, or package mutation.
- Preserve transform, name, fill, line, arrows, shadow/effects, hyperlink, text,
  shape order, extensions, relationships, opaque parts, and live model identity.
- Cover duplicate isolation, outer rollback, all six presentation formats,
  write/reopen, Node/browser/types/CLI packed surfaces, validator, round-trip,
  visual QA, overflow, and artifact-tool import.
- Do not implement arbitrary formulas, handles, connection sites, custom text
  rectangle formulas, or geometry evaluation in this plan.
- Never stage or commit `.pnpm-store/`. Execute inline without subagents. Every
  completed task is reviewed, committed, pushed, fetched, and verified at `0 0`.

---

### Task 1: Add the typed custom path codec

**Files:**
- Create: `packages/model/src/custom-geometry.ts`
- Create: `packages/model/src/custom-geometry.internal.ts`
- Create: `packages/model/src/custom-geometry.internal.test.ts`
- Modify: `packages/model/src/index.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, XML namespace helpers,
  `ModelParseError`, and existing descriptor-safe conventions.
- Produces: `CustomGeometryPoint`, `CustomGeometryCommand`,
  `CustomGeometryPathFill`, `CustomGeometryPath`, `CustomGeometry`,
  `AddCustomShapeOptions`, `NormalizedCustomGeometry`,
  `normalizeCustomGeometry()`, `renderCustomGeometry()`,
  `readCustomGeometry()`, `replaceCustomGeometry()`, and
  `customGeometryEqual()` for Tasks 2–4.

- [ ] **Step 1: Write the public types and compile-time contract**

Create `custom-geometry.ts` with the exact discriminated unions from the design:

```ts
import type { AddShapeOptions } from './preset-shape.js';

export interface CustomGeometryPoint {
  readonly x: number;
  readonly y: number;
}

export type CustomGeometryCommand =
  | { readonly kind: 'moveTo'; readonly point: CustomGeometryPoint }
  | { readonly kind: 'lineTo'; readonly point: CustomGeometryPoint }
  | {
      readonly kind: 'arcTo';
      readonly widthRadius: number;
      readonly heightRadius: number;
      readonly startAngle: number;
      readonly sweepAngle: number;
    }
  | {
      readonly kind: 'quadraticBezierTo';
      readonly control: CustomGeometryPoint;
      readonly end: CustomGeometryPoint;
    }
  | {
      readonly kind: 'cubicBezierTo';
      readonly control1: CustomGeometryPoint;
      readonly control2: CustomGeometryPoint;
      readonly end: CustomGeometryPoint;
    }
  | { readonly kind: 'close' };

export type CustomGeometryPathFill =
  | 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

export interface CustomGeometryPath {
  readonly width: number;
  readonly height: number;
  readonly fill?: CustomGeometryPathFill;
  readonly stroke?: boolean;
  readonly extrusionOk?: boolean;
  readonly commands: readonly CustomGeometryCommand[];
}

export interface CustomGeometry {
  readonly paths: readonly CustomGeometryPath[];
}

export type AddCustomShapeOptions = Omit<AddShapeOptions, 'adjustments'>;
```

Export the module from `packages/model/src/index.ts`. In the test, assign every
union branch to `CustomGeometryCommand`, exercise `AddCustomShapeOptions`, and add
`@ts-expect-error` assertions for `adjustments`, unknown kinds, missing points,
arc endpoints, and string coordinates.

- [ ] **Step 2: Write failing normalization tests**

Add a representative valid definition:

```ts
const input: CustomGeometry = {
  paths: [{
    width: 3_657_600,
    height: 2_743_200,
    fill: 'norm',
    stroke: true,
    extrusionOk: false,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 914_400, y: 0 } },
      {
        kind: 'quadraticBezierTo',
        control: { x: 1_371_600, y: 0 },
        end: { x: 1_828_800, y: 914_400 },
      },
      {
        kind: 'cubicBezierTo',
        control1: { x: 2_057_400, y: 914_400 },
        control2: { x: 2_514_600, y: 1_828_800 },
        end: { x: 2_743_200, y: 1_828_800 },
      },
      {
        kind: 'arcTo', widthRadius: 914_400, heightRadius: 457_200,
        startAngle: 1_800_000, sweepAngle: 7_200_000,
      },
      { kind: 'close' },
    ],
  }],
};
const normalized = normalizeCustomGeometry(input, 'Custom geometry');
expect(normalized).toEqual(input);
expect(Object.isFrozen(normalized)).toBe(true);
expect(Object.isFrozen(normalized.paths)).toBe(true);
expect(normalized.paths[0]?.commands.every(Object.isFrozen)).toBe(true);
```

Mutate every caller object/array after normalization and prove the snapshot does
not change. Accept null-prototype geometry/path/command/point objects, negative
point coordinates, zero angles, multiple paths, later `moveTo`, and an empty
command list.

Reject primitive/array/class/proxy-like containers, inherited properties,
symbols, accessors without invoking them, sparse/exotic arrays, extra/missing
keys, empty paths, non-positive/unsafe path extents, non-positive/unsafe arc
radii, non-number/fractional/non-finite/unsafe coordinates or angles, invalid
path-fill tokens, non-boolean flags, unknown command kinds, `close` first, and a
non-empty path whose first command is not `moveTo`.

- [ ] **Step 3: Implement minimal descriptor-safe normalization**

Define the exact internal type and entry point:

```ts
export type NormalizedCustomGeometry = Readonly<{
  readonly paths: readonly Readonly<{
    readonly width: number;
    readonly height: number;
    readonly fill?: CustomGeometryPathFill;
    readonly stroke?: boolean;
    readonly extrusionOk?: boolean;
    readonly commands: readonly Readonly<CustomGeometryCommand>[];
  }>[];
}>;

export function normalizeCustomGeometry(
  value: unknown,
  context: string,
): NormalizedCustomGeometry;
```

Use `Reflect.ownKeys()` and `Object.getOwnPropertyDescriptor()` at every level.
Copy and freeze each point, command, command array, path, path array, and root.
Normalize negative zero to zero. Do not round numeric input: coordinates,
radii, extents, and angles must already be safe integers.

- [ ] **Step 4: Write failing render/read/equality/replacement tests**

Use one `p:sp` fixture with direct `p:spPr`. Assert exact compact rendering for
all six commands, multiple paths, empty path, flags, negative coordinates, and
alternate prefixes. Assert `customGeometryEqual()` is ordered and distinguishes
path dimensions, optional-property absence, flags, command kind, control points,
and command order.

Reader success cases include canonical PptxGenJS output, self-closing or expanded
empty list containers, absent empty list containers, absent/default rect,
alternate prefixes, namespace declarations, XML whitespace, boolean lexical
forms, and signed decimal integers.

Reader rejection cases include wrong root/spPr/geometry/path namespaces,
missing/repeated/mixed geometry, missing/repeated `pathLst`, zero paths, missing/
duplicate/qualified/unsupported attributes, unsafe integers, formulas, non-empty
`avLst/gdLst/ahLst/cxnLst`, custom rect, wrong command child count/order,
nested lookalikes, non-whitespace text, unknown command, and invalid sequence.
For every rejected source, call `replaceCustomGeometry()` and assert
`ModelParseError`, exact source bytes, and no XML patch.

Assert same-value replacement returns `false` with exact bytes. Assert different
replacement keeps the existing DrawingML prefix, adds a local namespace only when
needed, replaces only the geometry span, and preserves transform/fill/line/effects/
text/extensions outside it.

- [ ] **Step 5: Implement the namespace-aware codec**

Add these exact functions:

```ts
export function renderCustomGeometry(
  geometry: NormalizedCustomGeometry,
  prefix: string,
): string;

export function readCustomGeometry(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedCustomGeometry | undefined;

export function replaceCustomGeometry(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  geometry: NormalizedCustomGeometry,
  partUri: string,
): boolean;

export function customGeometryEqual(
  left: NormalizedCustomGeometry | undefined,
  right: NormalizedCustomGeometry | undefined,
): boolean;
```

Render canonical `avLst/gdLst/ahLst/cxnLst/rect/pathLst` order. Parse expanded
namespace URI rather than lexical prefix. Preserve optional path-property absence
in snapshots. `replaceCustomGeometry()` initially edits only supported direct
`custGeom`; preset-to-custom conversion is added in Task 3.

- [ ] **Step 6: Run focused gates**

```sh
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts --reporter=dot
pnpm typecheck
pnpm build
git diff --check
```

Expected: codec tests, public type exports, typecheck, and build pass.

- [ ] **Step 7: Review, commit, push, and verify**

Review descriptor safety, no getter invocation, deep freeze, exact units, command
sequence, optional-property semantics, namespace ownership, list/rect rejection,
same-value no-op, prefix retention, and unrelated-byte preservation. Then:

```sh
git add -- packages/model/src/custom-geometry.ts \
  packages/model/src/custom-geometry.internal.ts \
  packages/model/src/custom-geometry.internal.test.ts packages/model/src/index.ts
git diff --cached --check
git commit -m "feat: add custom geometry path codec"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: `0 0`.

---

### Task 2: Create styled custom shapes from zero

**Files:**
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 geometry normalizer/renderer and existing preset shape style,
  transform, hyperlink, ID allocation, shape-tree insertion, and transaction code.
- Produces: `NormalizedCustomShape`, `normalizeCustomShape()`,
  `renderCustomShapeXml()`, and `SlideModel.addCustomShape()`.

- [ ] **Step 1: Write failing shared-option and renderer tests**

Add the internal shape definition:

```ts
export interface NormalizedCustomShape {
  readonly geometry: NormalizedCustomGeometry;
  readonly name: string | undefined;
  readonly fill: ShapeFill;
  readonly line: NormalizedSimpleLine | undefined;
  readonly arrows: NormalizedShapeArrows | undefined;
  readonly hyperlink: NormalizedHyperlink | undefined;
  readonly shadow: NormalizedShapeShadow | undefined;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}
```

Test `normalizeCustomShape(geometry, options)` with omitted/default options and a
fully styled shape. Assert it accepts the same transform/name/fill/line/arrows/
shadow/hyperlink values as preset creation, rejects `adjustments` even when
runtime value is `undefined`, rejects every unknown/accessor/symbol option before
geometry caller mutation can affect the normalized result, and resolves no slide
target yet.

Assert `renderCustomShapeXml(id, normalized, relationshipId)` produces the same
non-visual, transform, fill, line, arrow, effect, and hyperlink bytes as
`renderPresetShapeXml`, with only `custGeom` replacing `prstGeom`.

- [ ] **Step 2: Extract the minimal shared shape renderer and normalizer**

Inside `preset-shape.internal.ts`, keep public preset behavior byte-identical and
factor only the creation code needed by both paths:

```ts
export function normalizeCustomShape(
  geometry: unknown,
  options: unknown = undefined,
): NormalizedCustomShape;

export function renderCustomShapeXml(
  id: number,
  shape: NormalizedCustomShape,
  hyperlinkRelationshipId?: string,
): string;
```

Use one private `normalizeShapeOptions()` for the shared option fields and one
private `renderShapeXml()` that accepts a rendered geometry string. Keep
`normalizePresetShape()` responsible for type/adjustments and keep every existing
preset exact-XML test passing. Normalize geometry and options before returning the
frozen custom definition.

- [ ] **Step 3: Write failing public creation tests**

In model and SDK tests, create:

```ts
const geometry: CustomGeometry = {
  paths: [{
    width: inches(4),
    height: inches(3),
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: inches(4), y: 0 } },
      { kind: 'lineTo', point: { x: inches(2), y: inches(3) } },
      { kind: 'close' },
    ],
  }],
};
const shape = slide.addCustomShape(geometry, {
  name: 'Custom triangle',
  x: inches(1), y: inches(1), width: inches(4), height: inches(3),
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
  line: { kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 2 },
  arrows: { end: 'triangle' },
  shadow: { kind: 'outer' },
  hyperlink: { url: 'https://example.com/custom' },
});
```

Assert immediate ID/name/kind/transform/style/link state, `presetType === undefined`,
and the custom geometry bytes. Prove caller mutation cannot change created XML.
Create one multi-path and one empty-command shape. Assert insertion before direct
`p:extLst`.

Invalid geometry/options/slide targets must preserve exact slide bytes,
relationships, part set, mutation journal, and next shape ID.

- [ ] **Step 4: Implement `SlideModel.addCustomShape()`**

Add the exact method beside `addShape()`:

```ts
addCustomShape(
  geometry: CustomGeometry,
  options: AddCustomShapeOptions = {},
): ShapeModel;
```

Use the existing transaction order: normalize → resolve internal slide target →
parse and require one shape tree → allocate next ID → create external/internal
hyperlink relationship → render → insert before `extLst` → save → resolve the same
live `ShapeModel`. Error text must say `Custom shape` rather than `Preset shape`.

- [ ] **Step 5: Add lifecycle and six-format coverage**

Create styled custom shapes, duplicate the slide, edit unrelated transform/style,
move slides, delete a duplicate, write/reopen, and assert geometry bytes/snapshots.
Wrap creation in an outer transaction that throws and prove shape ID, relationship,
model cache, and journal rollback. Repeat representative creation/reopen for
`pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`.

- [ ] **Step 6: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  -t 'custom geometry|custom shape|preset shape' --reporter=dot
pnpm typecheck
pnpm build
git diff --check
```

Expected: custom creation and every existing preset creation regression pass.

- [ ] **Step 7: Review, commit, push, and verify**

Review option ownership, preset byte stability, geometry detachment, normalization
before allocation/relationships, hyperlink target identity, insertion order,
returned live model, rollback, duplicate lifecycle, and all formats. Then:

```sh
git add -- packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts packages/model/src/slide.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create custom geometry shapes"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: `0 0`.

---

### Task 3: Edit custom paths and convert geometry kinds

**Files:**
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Modify: `packages/model/src/custom-geometry.internal.test.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 codec, Task 2 creation, existing `ShapeModel.presetType`, and
  transaction-based slide mutation.
- Produces: `ShapeModel.customGeometry`, `SlideModel.getShapeCustomGeometry()`,
  `SlideModel.setShapeCustomGeometry()`, and safe preset/custom conversion.

- [ ] **Step 1: Write failing live getter/setter tests**

Create a custom shape and assert:

```ts
const first = shape.customGeometry;
const second = shape.customGeometry;
expect(first).toEqual(geometry);
expect(first).not.toBe(second);
expect(Object.isFrozen(first)).toBe(true);
expect(Object.isFrozen(first?.paths[0]?.commands)).toBe(true);
```

Assign the same snapshot and assert exact slide bytes and mutation journal. Assign
a changed multi-path snapshot and assert only `a:custGeom` changes; transform,
fill, line, arrows, effects, hyperlink relationship, text, shape order, extensions,
and model identity remain. Assign back the original and clear commands with an
explicit empty list.

Inject every unsupported source state from Task 1, assert getter `undefined`, then
assert setter throws `ModelParseError` with exact bytes/journal/relationships.
Invalid setter input must fail before resolving/mutating the package.

- [ ] **Step 2: Implement live delegation and transaction placement**

In `ShapeModel`:

```ts
get customGeometry(): CustomGeometry | undefined {
  return this.slide.getShapeCustomGeometry(this.id);
}

set customGeometry(value: CustomGeometry) {
  this.slide.setShapeCustomGeometry(this.id, value);
}
```

In `SlideModel`:

```ts
getShapeCustomGeometry(id: number): CustomGeometry | undefined {
  const { xml, element } = this.resolveShape(id);
  return readCustomGeometry(xml, element);
}

setShapeCustomGeometry(id: number, value: CustomGeometry): void {
  this.presentation.opcPackage.transaction(() => {
    const geometry = normalizeCustomGeometry(value, 'Custom geometry');
    const { xml, element } = this.resolveShape(id);
    if (replaceCustomGeometry(xml, element, geometry, this.partUri)) {
      this.setXml(xml.serialize());
    }
  });
}
```

Return newly detached deep-frozen snapshots from each read. Avoid `setXml()` when
replacement returns `false` so same-value journal remains exact.

- [ ] **Step 3: Write failing preset/custom conversion tests**

For preset → custom, start with adjusted `blockArc`, set `customGeometry`, and
assert `presetType`/`adjustments` become `undefined`, custom snapshot is readable,
and all non-geometry state remains. For custom → preset, set `presetType = 'ellipse'`
and assert custom snapshot becomes `undefined` and canonical empty `avLst` appears.

Repeat across duplicate isolation, outer rollback, reopen, and stable live identity.
Reject conversions from unsupported/mixed custom state and malformed preset state
without mutation. Same preset assignment remains exact no-op.

- [ ] **Step 4: Extend both replacement functions safely**

Update `replaceCustomGeometry()` so it accepts either a safely parsed supported
custom geometry or a unique canonical preset geometry; same custom remains no-op,
and preset input is whole-replaced by rendered custom geometry.

Update `replacePresetShapeType()` so it accepts either a canonical preset geometry
or a custom geometry for which `readCustomGeometry()` succeeds. Supported custom
input is whole-replaced by canonical `prstGeom + empty avLst`; unsupported custom
input throws. Do not make `readPresetShapeType()` report custom shapes.

- [ ] **Step 5: Add SDK type/lifecycle regression coverage**

Exercise public setter/getter types, compile-time rejection of `undefined`, all
command branches, multiple paths, duplicate edit isolation, rollback, all formats,
write/reopen, and interaction with `fill`, `line`, `arrows`, `shadow`, `hyperlink`,
`text`, and `presetType`.

- [ ] **Step 6: Run focused and full model gates**

```sh
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  -t 'custom geometry|custom shape|preset type' --reporter=dot
pnpm check
pnpm build
git diff --check
```

Expected: live editing, conversion, full model/SDK regressions, typecheck, and build pass.

- [ ] **Step 7: Review, commit, push, and verify**

Review detached snapshots, exact no-op, setter transaction, unsupported-state
isolation, conversion ownership, preset adjustment reset, style/text/relationship
preservation, duplicate/rollback/reopen, and stable identity. Then:

```sh
git add -- packages/model/src/custom-geometry.internal.ts \
  packages/model/src/custom-geometry.internal.test.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit custom geometry paths"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: `0 0`.

---

### Task 4: Add PptxGenJS conformance and packed public-surface coverage

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 3 public API and PptxGenJS 4.0.1 public
  `ShapeType.custGeom/addShape(points)/write` output.
- Produces: real-output compatibility evidence and installed Node/browser/type/CLI
  verification.

- [ ] **Step 1: Add all-command public-output fixtures**

Extend the local PptxGenJS test interface with the exact public points union. Build
named shapes using public APIs only:

```ts
generatedSlide.addShape(generated.ShapeType.custGeom!, {
  objectName: 'All custom commands', x: 1, y: 1, w: 4, h: 3,
  points: [
    { x: 0, y: 0 },
    { x: 1, y: 0 },
    { x: 2, y: 1, curve: { type: 'quadratic', x1: 1.5, y1: 0 } },
    {
      x: 3, y: 2,
      curve: { type: 'cubic', x1: 2.25, y1: 1, x2: 2.75, y2: 2 },
    },
    {
      x: 999, y: 999,
      curve: { type: 'arc', wR: 1, hR: 0.5, stAng: 30, swAng: 120 },
    },
    { x: 0.5, y: 0.5, moveTo: true },
    { x: 1.25, y: 1.25 },
    { close: true },
  ],
});
```

Import actual written bytes and assert the exact native `customGeometry` snapshot:
inch inputs converted to EMU, angles converted to OOXML units, arc `x/y` absent,
and shape extents copied to path width/height. Write through native, reopen, and
assert the same snapshot.

- [ ] **Step 2: Add empty, numeric, percentage, and divergence probes**

Generate empty points, direct numbers >=100, numeric strings, percentages, later
`moveTo`, first-curve, invalid kind, missing fields, zero/negative radii, and unsafe
values. Classify which outputs are schema-valid and supported snapshots. Assert
valid PptxGenJS final state imports; assert malformed/unsupported output remains
byte-preserved with `customGeometry === undefined`.

Create native explicit-EMU equivalents and compare snapshots. Assert native rejects
string/percentage input, mixed-unit heuristics, arc endpoints, first-curve paths,
invalid branches, and unsafe values before mutation.

- [ ] **Step 3: Extend actual-tarball smoke**

In `scripts/smoke-npm-package.mjs`, create one all-command and one multi-path custom
shape, assert immediate deep-frozen snapshots, edit a path, convert custom → preset
→ custom, write/reopen, and assert values. Add compile-time source importing every
custom geometry type and calling `addCustomShape()`/`shape.customGeometry`.

Require Node ESM, browser bundle, declaration compiler, and CLI smoke summary to
include `customGeometryPaths: true`.

- [ ] **Step 4: Run focused, packed, and build gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  -t 'custom geometry|custom shape' --reporter=dot
pnpm typecheck
pnpm build
node --check scripts/smoke-npm-package.mjs
mkdir -p /tmp/pptx-custom-geometry-pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-custom-geometry-pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-custom-geometry-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected: public-output fixtures and installed Node/browser/types/CLI all pass.

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only PptxGenJS generation, exact conversion math, arc endpoint
evidence, percentage/numeric heuristic evidence, malformed preservation, native
strictness, packed dependency isolation, declaration reachability, browser parity,
and CLI output. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify custom geometry path compatibility"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: `0 0`.

---

### Task 5: Document support and complete release QA

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: verified behavior from Tasks 1–4.
- Produces: accurate public documentation, updated remaining-gap status, and final
  compatibility evidence.

- [ ] **Step 1: Update usage and capability documentation**

Add a concise native example:

```ts
const custom = slide.addCustomShape({
  paths: [{
    width: inches(4),
    height: inches(3),
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: inches(4), y: 0 } },
      { kind: 'lineTo', point: { x: inches(2), y: inches(3) } },
      { kind: 'close' },
    ],
  }],
}, { fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } } });
custom.customGeometry = { ...custom.customGeometry!, paths: updatedPaths };
custom.presetType = 'triangle';
```

Document direct units, all command branches, multiple paths/subpaths, path flags,
deep-frozen snapshots, whole replacement, no clear, exact no-op, preset/custom
conversion, unsupported formulas/handles/connections/custom rect, PptxGenJS final
output mappings, and strict runtime divergences. Change the compatibility row from
generic custom geometry pending to custom paths supported with formulas/handles next.

- [ ] **Step 2: Run the complete automated gate**

```sh
pnpm check
pnpm test:performance
pnpm build
node --check scripts/smoke-npm-package.mjs
mkdir -p /tmp/pptx-custom-geometry-pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-custom-geometry-pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-custom-geometry-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected: all tests, typecheck, performance, build, installed package, and diff checks pass.

- [ ] **Step 3: Generate a packed-package custom geometry gallery**

Using only the packed public package, create a wide deck containing labeled open
and closed polylines, arc, quadratic and cubic curves, multiple subpaths, multiple
paths, empty path, path flags, style coexistence, edit, preset/custom conversion,
duplicate, and write/reopen cases. Use explicit Arial for visible labels. Keep all
gallery source and outputs in one task-specific `/tmp` directory.

Use `pptx-inspect` stable JSON to verify titles, slide/shape counts, direct custom
geometry path commands, relationships, and compatibility diagnostics. Validate the
original legal gallery against PowerPoint 2010 with 0 errors and 0 warnings.

- [ ] **Step 4: Render, round-trip, and inspect every slide**

Render the original with the Presentations helper, run `slides_test.py`, and inspect
every full-size slide. Open/export through an isolated LibreOffice profile, validate
the round-tripped file, compare every custom path snapshot, render again, export PDF,
verify with `pdfinfo`, rasterize with `pdftoppm`, rerun overflow, and inspect every
slide/page. Record client normalization separately from native behavior and fix any
unintended overlap, clipping, shape corruption, font substitution, or missing glyph.

- [ ] **Step 5: Verify artifact-tool import preservation**

Initialize a temporary Artifact Tool workspace and import original and round-tripped
galleries through JavaScript ES modules. Record SHA-256, slide/element counts, shape
names, bounds, and stable repeated-inspect IDs. Require no dropped named custom
shapes and no unexpected object-count or identity changes.

- [ ] **Step 6: Review, commit, push, and verify**

Review docs against exact public names and units, remaining-gap accuracy, PptxGenJS
divergences, test counts, packed output, validator diagnostics, original/round-trip
path snapshots, visual inspection, overflow, PDF, and Artifact Tool evidence. Then:

```sh
git add -- CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document custom geometry paths"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: `0 0`; only `.pnpm-store/` remains untracked. Custom geometry paths are
complete, and arbitrary formulas/handles/connections/custom text rectangles become
the next custom-geometry item.
