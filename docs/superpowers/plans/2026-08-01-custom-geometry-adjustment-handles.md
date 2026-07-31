# Custom Geometry Adjustment Handles Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict ordered DrawingML XY/polar custom-geometry adjustment handles for zero-input creation and lossless existing-deck editing.

**Architecture:** Extend the existing `CustomGeometry` direct-state tree with one ordered discriminated handle union. Keep normalization, rendering, parsing, semantic equality, and unsupported-state gating in `custom-geometry.internal.ts`; reuse the current `SlideModel.addCustomShape()` / `ShapeModel.customGeometry` transaction path for public lifecycle behavior.

**Tech Stack:** TypeScript 5.9, Vitest, `@pptx/lossless-xml`, OPC package transactions, PptxGenJS 4.0.1 output fixtures, Node/browser package smoke tests, LibreOffice headless, PowerPoint 2010 Open XML validation.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-01-custom-geometry-adjustment-handles-design.md` exactly.
- Do not add connection sites, custom text rectangles, formula/handle evaluation, drag APIs, or resolved bounds.
- Preserve existing formula/path source compatibility: omitted or empty `handles` does not appear as an own snapshot property.
- Preserve XY/polar cross-kind list order and independent optional-attribute presence.
- Reuse `CustomGeometryValue` and the existing token contract; do not infer guide existence, min/max pairing, range order, units, or arithmetic validity.
- Keep unsupported OOXML lossless and reject unsafe edits before package mutation.
- Execute inline in the root task; do not dispatch subagents.
- Do not stage `.pnpm-store/`.
- End every task with review, an isolated commit, push to `main`, fetch, and `HEAD...origin/main == 0 0`.

---

### Task 1: Public handle types, strict normalization, rendering, and equality

**Files:**
- Modify: `packages/model/src/custom-geometry.ts`
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Produces: `CustomGeometryXyHandle`, `CustomGeometryPolarHandle`, `CustomGeometryHandle`.
- Produces: optional `CustomGeometry.handles` ordered direct-state list.
- Preserves: `normalizeCustomGeometry()`, `renderCustomGeometry()`, and `customGeometryEqual()` signatures.

- [ ] **Step 1: Add failing public-type and normalization tests**

Add a typed fixture beside `formulaGeometry`:

```ts
const handleGeometry: CustomGeometry = {
  adjustments: [
    { name: 'adjX', formula: { operator: 'val', operands: [25_000] } },
    { name: 'adjY', formula: { operator: 'val', operands: [50_000] } },
    { name: 'adjR', formula: { operator: 'val', operands: [30_000] } },
    { name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } },
  ],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adjR', 100_000] } },
    { name: 'y1', formula: { operator: '*/', operands: ['h', 'adjR', 100_000] } },
  ],
  handles: [
    {
      kind: 'xy',
      position: { x: 'adjX', y: 'adjY' },
      xGuide: 'adjX',
      minX: 0,
      maxX: 100_000,
      yGuide: 'adjY',
      minY: 't',
      maxY: 'b',
    },
    {
      kind: 'polar',
      position: { x: 'x1', y: 'y1' },
      radiusGuide: 'adjR',
      minRadius: 0,
      maxRadius: 'ss',
      angleGuide: 'adjAng',
      minAngle: 0,
      maxAngle: 'cd',
    },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'adjX', y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'adjY' } },
      { kind: 'close' },
    ],
  }],
};
```

Assert normalization detaches and recursively freezes the root, handle list, each handle, and each
position. Mutate caller-owned handles/positions after normalization and prove the snapshot is unchanged.
Assert omitted and `handles: []` normalize to no own property. Add compile-only declarations proving the
union narrows by `kind` and rejects cross-kind fields.

- [ ] **Step 2: Add failing strict-input cases**

Cover non-object handles, unknown/missing kind, missing position, wrong-kind fields, unknown/accessor/symbol
properties, sparse/subclass arrays, unsafe numbers, decimal/whitespace/empty/invalid-XML guide refs or
bounds, and point accessors. Assert accessors are never invoked. Explicitly prove min/max may be independently
present, may use tokens, may be numerically reversed, and do not require a guide ref.

- [ ] **Step 3: Run the focused test and verify the red state**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: type/transpile or assertions fail because handle types and the root field do not exist.

- [ ] **Step 4: Add the public type model**

In `custom-geometry.ts`, add exactly:

```ts
export interface CustomGeometryXyHandle {
  readonly kind: 'xy';
  readonly position: CustomGeometryPoint;
  readonly xGuide?: string;
  readonly minX?: CustomGeometryValue;
  readonly maxX?: CustomGeometryValue;
  readonly yGuide?: string;
  readonly minY?: CustomGeometryValue;
  readonly maxY?: CustomGeometryValue;
}

export interface CustomGeometryPolarHandle {
  readonly kind: 'polar';
  readonly position: CustomGeometryPoint;
  readonly radiusGuide?: string;
  readonly minRadius?: CustomGeometryValue;
  readonly maxRadius?: CustomGeometryValue;
  readonly angleGuide?: string;
  readonly minAngle?: CustomGeometryValue;
  readonly maxAngle?: CustomGeometryValue;
}

export type CustomGeometryHandle = CustomGeometryXyHandle | CustomGeometryPolarHandle;
```

Add `readonly handles?: readonly CustomGeometryHandle[]` before `paths` in `CustomGeometry`.

- [ ] **Step 5: Implement descriptor-safe handle normalization**

Import the three new types into `custom-geometry.internal.ts`. Add exact key sets:

```ts
const XY_HANDLE_KEYS = new Set([
  'kind', 'position', 'xGuide', 'minX', 'maxX', 'yGuide', 'minY', 'maxY',
]);
const POLAR_HANDLE_KEYS = new Set([
  'kind', 'position', 'radiusGuide', 'minRadius', 'maxRadius',
  'angleGuide', 'minAngle', 'maxAngle',
]);
const HANDLE_REQUIRED_KEYS = new Set(['kind', 'position']);
```

Add `handles` to `ROOT_KEYS`. Normalize it after guides and before paths:

```ts
const handles = Object.hasOwn(root, 'handles')
  ? normalizeHandleList(root.handles, `${context} handles`)
  : undefined;
```

`normalizeHandleList()` must call existing `readArray()`, branch on a descriptor-safe `kind`, call
`requireKeys()` with the correct set, normalize `position` through `normalizePoint()`, guide refs through
`normalizeCustomGeometryToken()`, bounds through `normalizeCustomGeometryValue(..., false)`, copy only
present optional properties, and freeze every object/list. Include `{ handles }` in the root only when
non-empty.

- [ ] **Step 6: Render handles canonically**

Replace the fixed `ahLst` output with `renderHandleList(geometry.handles, prefix)`. Render absent/empty as
self-closing. Render non-empty values in exact list order with these attribute mappings and order:

```ts
const XY_ATTRIBUTE_FIELDS = [
  ['xGuide', 'gdRefX'], ['minX', 'minX'], ['maxX', 'maxX'],
  ['yGuide', 'gdRefY'], ['minY', 'minY'], ['maxY', 'maxY'],
] as const;

const POLAR_ATTRIBUTE_FIELDS = [
  ['radiusGuide', 'gdRefR'], ['minRadius', 'minR'], ['maxRadius', 'maxR'],
  ['angleGuide', 'gdRefAng'], ['minAngle', 'minAng'], ['maxAngle', 'maxAng'],
] as const;
```

Only present properties emit attributes. Use `escapeXmlAttribute(String(value))`. Each handle contains one
`<prefix:pos x="..." y="..."/>` rendered through a focused helper that reuses the point value escaping.

- [ ] **Step 7: Extend semantic equality**

Compare optional handle-list presence, length, cross-kind order, kind, position, and every optional property's
presence/value before comparing paths. Add variants for absent list, reversed order, kind change, position
change, removed optional property, and changed bound. Do not compare effective movement behavior.

- [ ] **Step 8: Run focused tests and typecheck**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm typecheck
git diff --check
```

Expected: focused suite and typecheck pass; existing formula/path snapshots remain unchanged.

- [ ] **Step 9: Review, commit, and push Task 1**

Review only the three task files and confirm every changed line maps to types, normalization, rendering, or
equality. Then:

```bash
git add packages/model/src/custom-geometry.ts packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: add custom geometry adjustment handle codec"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 2: Strict OOXML reader and live whole-replacement editing

**Files:**
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 handle union and canonical renderer.
- Produces: `readCustomGeometry()` snapshots for supported non-empty `a:ahLst`.
- Preserves: exact semantic no-op, alternate prefix, unsupported connection/custom-rect isolation.

- [ ] **Step 1: Add failing supported-reader fixtures**

Render `handleGeometry` with `a:` and `d:` prefixes and assert both read back identically. Add a lexical
fixture using `+0`, `+100000`, escaped refs/tokens, mixed XY/polar order, all attributes absent on one handle,
and whitespace around elements. Assign the normalized snapshot back and prove serialized bytes are unchanged.

- [ ] **Step 2: Add failing malformed-reader cases**

Add one case for each of: repeated `ahLst`, list attribute/text, wrong-namespace/unknown handle child,
missing/repeated/wrong-namespace `pos`, extra handle child/text, unknown/qualified handle attribute,
missing/qualified/extra position attribute, unsafe integer, decimal/whitespace/empty token, and illegal XML
character. Retain non-empty `cxnLst` and non-default `rect` rejection. For every case, assert getter
`undefined`, replacement throws `ModelParseError`, and serialized XML is unchanged.

- [ ] **Step 3: Run the focused test and verify failure**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: non-empty handles still return `undefined` before reader implementation.

- [ ] **Step 4: Parse the ordered handle list**

Change `EMPTY_CUSTOM_LIST_NAMES` to contain only `cxnLst`. In `parseCustomGeometryElement()`:

```ts
const handles = parseHandleList(children);
if (handles === undefined) return undefined;
```

Add `...(handles.length ? { handles } : {})` to the root passed to `normalizeCustomGeometry()`.

`parseHandleList()` must distinguish absent/empty success from malformed failure, validate list attributes/text,
iterate direct children in source order, require the DrawingML namespace, and dispatch only `ahXY` or
`ahPolar`.

- [ ] **Step 5: Parse exact handle attributes and position**

For XY call `readXmlAttributes()` with allowed
`gdRefX/minX/maxX/gdRefY/minY/maxY` and no required attributes. For polar use
`gdRefR/minR/maxR/gdRefAng/minAng/maxAng`. Build public objects with this exact mapping:

```ts
gdRefX -> xGuide       gdRefY -> yGuide
gdRefR -> radiusGuide  gdRefAng -> angleGuide
minR   -> minRadius    maxR -> maxRadius
minAng -> minAngle     maxAng -> maxAngle
```

Guide refs use `parseCustomGeometryToken()`. Bounds use
`parseCustomGeometryValue(value, false)`. Preserve optional absence with `Object.hasOwn(attributes, xmlName)`.

Generalize `parsePointElement(element)` to `parsePointElement(element, expectedName = 'pt')`, or add a
focused equivalent, so handle parsing requires exactly one direct `pos` while path parsing still requires
`pt`. In both cases require same namespace, no children/text, and exact unqualified `x/y`.

- [ ] **Step 6: Run reader/edit isolation tests**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm vitest run packages/model/src/model.test.ts -t "custom geometry"
pnpm typecheck
git diff --check
```

Expected: supported handles read/edit/no-op; malformed handles and remaining connection/rect states reject with
unchanged bytes.

- [ ] **Step 7: Review, commit, and push Task 2**

```bash
git add packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: edit custom geometry adjustment handles"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 3: Public lifecycle, six-format, SDK, and PptxGenJS boundary coverage

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public `SlideModel.addCustomShape()` and `ShapeModel.customGeometry` with handles.
- Proves: create/read/edit/conversion/duplicate/rollback/write/reopen behavior and native-extension boundary.

- [ ] **Step 1: Add public create/edit and detachment tests**

Add a `customHandleGeometry` fixture with formulas, one fully populated XY handle, one fully populated polar
handle, and guide-referenced paths. Create a styled/hyperlinked shape from mutable input, mutate the caller's
handle list/objects/positions, and assert the live snapshot remains equal and recursively frozen.

Capture package bytes/journal before assigning `structuredClone(snapshot)` and prove exact no-op. Replace a
handle bound, position token, and cross-kind order; prove stable shape identity plus transform, fill, line,
arrows, shadow, hyperlink, text/effects/extensions, relationships, and neighbor bytes are preserved.

- [ ] **Step 2: Extend lifecycle and conversion coverage**

For each `PRESENTATION_FORMAT_PROFILES` entry, create handle geometry, duplicate its slide, edit only the
source handles, move/delete the duplicate, write/reopen, and assert source/duplicate isolation. Wrap creation
and edit in throwing outer transactions and assert package snapshots, mutation journal, and next shape ID roll
back. Convert preset -> handle custom -> preset while preserving the live model object and unrelated state.

- [ ] **Step 3: Keep remaining unsupported targets isolated**

Inject non-empty connection lists and non-default rect independently around otherwise valid handles. Assert
getter `undefined`, setter `ModelParseError`, and zero package mutation. Do not weaken those gates to make
client round-trips pass.

- [ ] **Step 4: Add SDK zero-input integration coverage**

In `packages/sdk/src/index.test.ts`, create a blank document, add handle geometry, write/reopen, edit handle
position/bounds/order, write/reopen again, and assert the final snapshot plus `validatePackage()` zero errors.

- [ ] **Step 5: Preserve PptxGenJS 4.0.1 behavior**

In the custom-geometry adapter tests, assert legal PptxGenJS shapes continue to have no own `handles` property
and all existing path snapshots stay equal. Keep malformed runtime outputs returning `undefined`. Add an
explicit assertion/comment that the public PptxGenJS 4.0.1 shape input exposes no arbitrary handle field, so
production adapter code remains unchanged.

- [ ] **Step 6: Run public and full suites**

Run:

```bash
pnpm vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm test
pnpm typecheck
git diff --check
```

Expected: all tests pass with only the repository's established skip count.

- [ ] **Step 7: Review, commit, and push Task 3**

```bash
git add packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: verify custom geometry adjustment handle lifecycle"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 4: Actual tarball Node/browser/type/CLI verification

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves: aggregate package exports handle types and packed Node/browser builds preserve runtime behavior.
- Produces: `customGeometryAdjustmentHandles: true` in package-smoke JSON.

- [ ] **Step 1: Extend the generated TypeScript consumer**

Import `CustomGeometryXyHandle`, `CustomGeometryPolarHandle`, and `CustomGeometryHandle`. Add valid typed XY
and polar objects and a `CustomGeometry` using their ordered union. Add `@ts-expect-error` cases for a missing
position and an XY object carrying `radiusGuide`; keep runtime lexical failures out of compile assertions.

- [ ] **Step 2: Extend the installed Node package smoke**

Through the installed tarball, create mixed handle geometry, verify recursive freeze/detachment, exact no-op,
handle edit/order change, preset/custom conversion, write/reopen, and expected compact `ahXY/ahPolar/pos`
XML. Add `customGeometryAdjustmentHandles` to both the internal check object and final JSON summary.

- [ ] **Step 3: Extend browser-condition smoke**

Update the browser custom-geometry block to create/read/edit/reopen mixed handles through `dist/browser.js`.
Require frozen handle snapshots, optional-property preservation, and final geometry equality.

- [ ] **Step 4: Build, pack, and run actual package smoke**

Run:

```bash
pnpm build
mkdir -p /tmp/pptx-adjustment-handle-pack
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-adjustment-handle-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-adjustment-handle-pack/jiayunxie-pptx-0.1.0.tgz
git diff --exit-code -- packages/pptx/dist
git diff --check
```

If pnpm emits a different exact tarball filename, pass that emitted path. Expected JSON contains
`"customGeometryAdjustmentHandles":true`, `"types":true`, and CLI version `0.1.0`.

- [ ] **Step 5: Review, commit, and push Task 4**

```bash
git add scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify packaged custom geometry adjustment handles"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 5: Documentation, handle gallery, compatibility validation, and release gate

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Temporary only: handle gallery script and outputs under a new `/tmp/pptx-custom-geometry-handles-*` directory

**Interfaces:**
- Documents: ordered handle union, direct fields, units/tokens, optional-state semantics, and remaining boundary.
- Proves: generated handle PPTX is valid, renderable, reopenable, and client round-trip behavior is understood.

- [ ] **Step 1: Generate a representative gallery from the actual tarball**

Create a 16:9 deck with four slides:

1. XY x-only, y-only, and two-axis handles driving visible guide-based polygons;
2. polar radius-only, angle-only, and two-axis polar handles driving curves/arcs;
3. mixed XY/polar ordered list, independently omitted optional attributes, numeric and token bounds;
4. handle edit, list reorder, and preset/custom conversion lifecycle evidence.

Use audience-facing labels that distinguish stored PowerPoint editing metadata from rendered geometry. Save
source snapshots and SHA-256 beside the temporary deck.

- [ ] **Step 2: Validate and visually inspect every slide**

Run the presentation overflow helper, export PDF through LibreOffice, rasterize every page, and inspect every
full-size slide. Fix clipping, overlap, broken paths, wrapping, or unreadable labels. Handles are not expected
to appear in slideshow output; visible geometry must still prove current guide values render correctly.

- [ ] **Step 3: Round-trip and structurally compare**

LibreOffice-save a copy, reopen original and round-trip with the actual tarball, and compare slide/shape counts,
names, adjustments, guides, formulas, handle kinds/order/attributes/positions, path values, and commands. If a
client inserts a non-default rect or evaluates tokens, inspect raw OOXML and record that normalization instead
of broadening the strict reader.

- [ ] **Step 4: Run PowerPoint 2010 validation**

Validate original and round-trip with:

```bash
pptx-inspect --json package validate <deck> --profile powerpoint-2010
```

Required result: 0 errors / 0 warnings for both. Any schema diagnostic blocks release until emitted OOXML is
fixed.

- [ ] **Step 5: Update public documentation**

Document the three handle types, ordered `handles`, public-to-OOXML field mapping, position/bound units and
tokens, independent optional presence, strict parsing, detachment/freeze/no-op, PptxGenJS native-extension
status, and remaining connection/rect/evaluator boundary. Update the compatibility row so handles are complete
and connection sites are next.

- [ ] **Step 6: Run the complete release gate**

Run:

```bash
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --exit-code -- packages/pptx/dist
node scripts/smoke-npm-package.mjs /tmp/pptx-adjustment-handle-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected: 603-or-higher tests pass with the established skip count, 1,000-part performance stays within five
seconds, dist is reproducible, actual tarball Node/browser/types/CLI checks pass, and no stale unsupported-handle
claim remains.

- [ ] **Step 7: Final review, commit, and push Task 5**

Stage only the five documentation files:

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document custom geometry adjustment handles"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected divergence: `0 0`; only `.pnpm-store/` may remain untracked.

## Plan Self-Review

- Spec coverage: Tasks 1-2 implement the complete ordered XY/polar direct-state codec and reader/editor; Task 3
  proves lifecycle/PptxGenJS boundaries; Task 4 proves packed runtimes/types; Task 5 proves real-file
  compatibility and documentation.
- Scope: connection sites, custom rect, evaluator, drag behavior, and resolved bounds remain explicitly excluded.
- Type consistency: every task uses `CustomGeometryXyHandle`, `CustomGeometryPolarHandle`,
  `CustomGeometryHandle`, and `handles` exactly as defined by the design.
- Optional-state consistency: omitted properties remain omitted; empty handle lists normalize to root absence;
  cross-kind order is never split or sorted.
- Mutation safety: invalid input and unsafe existing OOXML have explicit zero-mutation assertions before public
  release claims.
- Execution: standing user delegation selects inline execution; no additional choice or pause is needed.
