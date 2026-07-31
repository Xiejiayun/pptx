# Preset Shape Adjustments Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with independent review, commit, push, fetch, and remote-divergence verification.

**Goal:** Add strict native creation and lossless existing-deck editing for ordered preset-shape `val` adjustment guides, with valid PptxGenJS 4.0.1 shortcut-output parity.

**Architecture:** A focused `shape-adjustments.internal.ts` module owns descriptor-safe list normalization, immutable snapshots, canonical rendering, strict namespace-aware source inspection, semantic equality, and direct `a:avLst` replacement. Preset-shape creation consumes its normalizer and renderer; `ShapeModel` delegates existing-deck reads and transactional writes through `SlideModel`.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML source-span editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke tests, `pptx-inspect`, LibreOffice, Poppler, and Presentations render/overflow tools.

## Global Constraints

- Public values are ordered `readonly ShapeAdjustment[]`, where each entry is exactly `{ readonly name: string; readonly value: number }`.
- `AddShapeOptions.adjustments` is optional. `ShapeModel.adjustments` returns a detached deeply frozen list or `undefined` for unsupported source state; assignment accepts a list only and `[]` clears.
- Lists must be dense ordinary arrays. Entries must be ordinary/null-prototype objects containing exactly own data properties `name` and `value`; no accessors, symbols, inherited values, aliases, extra fields, sparse slots, or array subclasses.
- Names are non-empty XML-safe strings and unique within the list. Preserve list order and XML-escape names.
- Values are finite safe integers and are the direct integer operands of `fmla="val N"`; perform no shape-specific unit conversion or range guessing.
- Read only one namespace-correct direct `p:spPr/a:prstGeom/a:avLst`. Guides must be direct same-namespace `a:gd` elements with unique unqualified `name`/`fmla`, no unsupported attributes or child elements, and formula grammar `val` + XML whitespace + one signed safe integer.
- Missing/repeated/wrong-namespace structures, duplicate names, non-guide children, complex formulas, ambiguous attributes, unsafe integers, and malformed geometry read as `undefined`; mutation throws `ModelParseError` before package changes.
- Same ordered normalized value is an exact bytes/journal no-op. Changed values replace only the direct `a:avLst` span using its current DrawingML prefix. Clear retains the list container.
- Preset-type changes keep existing behavior: a different type resets the complete geometry to empty adjustments; the same type preserves exact bytes.
- Omitted and explicit-empty creation keep the previously published compact `<a:avLst/>` bytes.
- Valid PptxGenJS `rectRadius`, `angleRange`, and `arcThicknessRatio` final guide lists must import to the same snapshots. Native does not copy PptxGenJS truthiness loss, string coercion, invalid-range passthrough, or shortcut precedence.
- Custom geometry, arbitrary guide formulas, handles, connection sites, paths, and geometry evaluation remain outside this plan.
- Node 20+ and browser bundles expose identical behavior from `@jiayunxie/pptx`.
- Never stage or commit `.pnpm-store/`. Each task is independently reviewed and committed. Retry pending pushes non-interactively whenever remote connectivity is available.
- Execute inline without subagents or routine decision pauses.

---

### Task 1: Add the strict adjustment value and OOXML owner codec

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Create: `packages/model/src/shape-adjustments.internal.ts`
- Create: `packages/model/src/shape-adjustments.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `escapeXmlAttribute()`, `ModelParseError`, and existing direct-child/namespace/source-span conventions.
- Produces: public `ShapeAdjustment`, `NormalizedShapeAdjustments`, `normalizeShapeAdjustments()`, `renderShapeAdjustmentList()`, `readShapeAdjustments()`, `replaceShapeAdjustments()`, and `shapeAdjustmentsEqual()` for Tasks 2–4.

- [ ] **Step 1: Write normalization and immutability tests**

Cover empty, one-guide, ordered multi-guide, explicit zero, negative value,
`Number.MIN_SAFE_INTEGER`, `Number.MAX_SAFE_INTEGER`, XML-sensitive names,
null-prototype entries, immediate caller detachment, and deep freeze:

```ts
const input = [
  { name: 'adj1', value: 16_200_000 },
  Object.assign(Object.create(null), { name: 'adj2', value: 0 }),
];
const normalized = normalizeShapeAdjustments(input, 'Shape adjustments');
expect(normalized).toEqual([
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
]);
expect(Object.isFrozen(normalized)).toBe(true);
expect(normalized.every(Object.isFrozen)).toBe(true);
input[0]!.value = 7;
expect(normalized[0]!.value).toBe(16_200_000);
```

Reject null/primitives/objects, frozen-but-non-array objects, array subclasses,
sparse arrays, non-index properties, symbol properties, accessors without
invocation, entry arrays/dates/classes, missing/unknown/inherited/accessor
fields, empty/non-string/XML-invalid names, duplicate names, non-number,
fractional, NaN/infinite, and unsafe values.

- [ ] **Step 2: Write canonical renderer and equality tests**

Require supplied-prefix rendering and ordered equality:

```ts
expect(renderShapeAdjustmentList([
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
], 'a:')).toBe(
  '<a:avLst><a:gd name="adj1" fmla="val 16200000"/>' +
  '<a:gd name="adj2" fmla="val 0"/></a:avLst>',
);
expect(renderShapeAdjustmentList([], 'd:')).toBe('<d:avLst/>');
```

Verify escaped names, negative and safe-bound integers, equal cloned lists,
order sensitivity, name/value differences, length differences, and
`undefined` equality only with `undefined`.

- [ ] **Step 3: Write strict namespace-aware reader tests**

Build `p:sp` fixtures and require canonical snapshots for empty and multi-guide
lists, alternate DrawingML prefixes, namespace declarations, `+7`, `-0`, tabs,
line feeds, and carriage returns between `val` and operand. Snapshots must be
detached and deeply frozen.

Return `undefined` without mutation for wrong root/property/geometry/list/guide
namespaces, missing/repeated `spPr`, `prstGeom`, or `avLst`, nested lookalikes,
non-guide children, non-whitespace text, duplicate guide names,
missing/repeated/qualified `name` or `fmla`, unknown attributes, guide child
elements, leading/trailing formula tokens, missing/multiple operands, decimals,
scientific notation, unsafe integers, and complex operators such as `*/`, `+-`,
`pin`, `min`, and guide references.

- [ ] **Step 4: Write direct-list replacement tests**

Require an exact no-op for the same ordered normalized list, including source
with alternate accepted whitespace/sign lexical forms. Require a changed value
to replace only `a:avLst`, preserve `prstGeom@prst`, its unknown attributes,
namespace declarations, geometry-local comments/elements, fill, line, effects,
text, names, IDs, neighbors, and unrelated bytes. Require clear to write one
same-prefix empty list.

For every unsupported reader fixture, call `replaceShapeAdjustments()` and
require `ModelParseError` containing the part URI, `xml.changed === false`, and
byte-identical serialization.

- [ ] **Step 5: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/shape-adjustments.internal.test.ts --reporter=dot
```

Expected: FAIL because `shape-adjustments.internal.ts` does not exist.

- [ ] **Step 6: Implement the minimal codec**

Add the public value type to `preset-shape.ts`:

```ts
export interface ShapeAdjustment {
  readonly name: string;
  readonly value: number;
}
```

Define the exact internal surface:

```ts
export type NormalizedShapeAdjustments = readonly Readonly<ShapeAdjustment>[];

export function normalizeShapeAdjustments(
  value: unknown,
  context: string,
): NormalizedShapeAdjustments;

export function renderShapeAdjustmentList(
  adjustments: NormalizedShapeAdjustments,
  prefix: string,
): string;

export function readShapeAdjustments(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedShapeAdjustments | undefined;

export function replaceShapeAdjustments(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  adjustments: NormalizedShapeAdjustments,
  partUri: string,
): boolean;

export function shapeAdjustmentsEqual(
  left: NormalizedShapeAdjustments | undefined,
  right: NormalizedShapeAdjustments | undefined,
): boolean;
```

Use descriptors and `Reflect.ownKeys()` for input validation. Freeze each copied
entry and the copied array. Parse formulas with one anchored expression; reject
unsafe integers after lexical validation. Keep the reader non-throwing. Inspect
expanded names through in-scope namespace bindings, and replace only the
resolved list element span.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/shape-adjustments.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/shape-fill.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; no public model behavior changes yet.

- [ ] **Step 8: Review, commit, push, and verify**

Review descriptor safety, array density, getter-free behavior, detachment,
freeze depth, duplicate detection, safe-integer boundaries, formula grammar,
expanded-name checks, same-value semantics, prefix retention, source-span
isolation, and unsupported-state zero mutation. Then:

```sh
git add -- packages/model/src/shape-adjustments.internal.ts \
  packages/model/src/shape-adjustments.internal.test.ts \
  packages/model/src/preset-shape.ts
git diff --cached --check
git commit -m "feat: add preset shape adjustment codec"
GIT_TERMINAL_PROMPT=0 git -c credential.interactive=never push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence after connectivity is available: `0 0`.

---

### Task 2: Create preset shapes with ordered adjustments

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public type, normalization, and rendering plus current preset-shape creation.
- Produces: `AddShapeOptions.adjustments`, normalized creation state, and canonical guide-list XML for Task 3.

- [ ] **Step 1: Add the creation option and compile-time contract tests**

Extend `AddShapeOptions` in `preset-shape.ts`:

```ts
export interface AddShapeOptions extends Partial<Transform> {
  readonly adjustments?: readonly ShapeAdjustment[];
}
```

Add valid compile-time uses and `@ts-expect-error` cases for missing/extra
fields, wrong value types, scalar input, and `undefined` list entries. Confirm
the type remains reachable from `@pptx/model`, `@pptx/sdk`, and the bundled
package through existing star exports.

- [ ] **Step 2: Add creation behavior tests**

Create round-rectangle, pie, arc, block-arc, and an additional adjustable preset:

```ts
const shape = slide.addShape('blockArc', {
  adjustments: [
    { name: 'adj1', value: 16_200_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 25_000 },
  ],
});
const { xml, element } = slide.resolveShape(shape.id);
expect(readShapeAdjustments(xml, element)).toEqual([
  { name: 'adj1', value: 16_200_000 },
  { name: 'adj2', value: 0 },
  { name: 'adj3', value: 25_000 },
]);
```

Until Task 3 adds the public getter, inspect with the Task 1 reader. Require
exact guide XML, caller detachment, coexistence with name/transform/fill/line/
arrows/hyperlink/shadow, stable IDs/order, and all 178 preset types remaining
creatable with omitted adjustments.

- [ ] **Step 3: Prove omitted/empty stability and invalid zero mutation**

Require omitted, runtime `undefined`, and explicit `[]` to preserve the existing
compact empty-list bytes. For every invalid Task 1 value, snapshot slide bytes,
parts, relationships, mutation journal, cached shape identities, and next ID
before `addShape()`; require identical state after rejection. Accessor fixtures
must prove getters were not invoked.

- [ ] **Step 4: Add transaction, duplication, and all-format creation tests**

Create adjusted shapes inside a transaction that later throws and require exact
rollback. Duplicate an adjusted slide and prove initially equal lists and
independent bytes. Write/reopen adjusted shapes in `pptx`, `pptm`, `potx`,
`potm`, `ppsx`, and `ppsm`, preserving names, values, order, and unrelated
properties.

- [ ] **Step 5: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape adjustment' \
  packages/sdk/src/index.test.ts -t 'preset shape adjustment' --reporter=dot
```

Expected: FAIL because the public type and creation integration are absent.

- [ ] **Step 6: Implement creation integration**

Add `adjustments` to `OPTION_KEYS` and `NormalizedPresetShape`; normalize with
`normalizeShapeAdjustments(values.adjustments ?? [], 'Preset shape adjustments')`
without treating an own runtime `undefined` as an invalid list. Replace the
hard-coded list inside the geometry renderer with:

```ts
`<a:prstGeom prst="${type}">` +
  renderShapeAdjustmentList(shape.adjustments, 'a:') +
  '</a:prstGeom>'
```

Do not alter no-adjustment bytes or any non-geometry renderer.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/shape-adjustments.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape|shape adjustment' \
  packages/sdk/src/index.test.ts -t 'preset shape|shape adjustment' --reporter=dot
pnpm vitest run packages/model/src/shape-fill.internal.test.ts \
  packages/model/src/shape-line.internal.test.ts \
  packages/model/src/shape-arrows.internal.test.ts \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/shape-shadow.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; omitted creation bytes and existing shape features remain exact.

- [ ] **Step 8: Review, commit, push, and verify**

Review public typing, option-key strictness, runtime-undefined behavior,
normalization before allocation/relationships, canonical order, no-adjustment
stability, detachment, rollback, all-format reopen, and absence of custom
geometry behavior. Then:

```sh
git add -- packages/model/src/preset-shape.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create preset shape adjustments"
GIT_TERMINAL_PROMPT=0 git -c credential.interactive=never push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence after connectivity is available: `0 0`.

---

### Task 3: Read, replace, and clear adjustments through the live model

**Files:**
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 read/replace/normalize functions and Task 2 public type.
- Produces: `ShapeModel.adjustments`, `SlideModel.getShapeAdjustments()`, and `SlideModel.setShapeAdjustments()`.

- [ ] **Step 1: Add live getter/setter tests**

Require immediate snapshots after native creation, detached deep-frozen getter
values, ordered whole-list replacement, value/name/order changes, and explicit
empty clear:

```ts
expect(shape.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
shape.adjustments = [{ name: 'adj', value: 40_000 }];
expect(shape.adjustments).toEqual([{ name: 'adj', value: 40_000 }]);
shape.adjustments = [];
expect(shape.adjustments).toEqual([]);
```

Runtime `undefined`, malformed entries, accessors, duplicates, and invalid
numbers must reject before package mutation.

- [ ] **Step 2: Add exact no-op and ownership-isolation tests**

Inject an existing list with accepted alternate whitespace and sign lexical
forms. Assign the normalized same value and require exact slide bytes and
journal. Change the list and require only `avLst` bytes to differ. Preserve
geometry attributes/content, preset type, transform, fill, line, arrows,
hyperlink, shadow/effect siblings, text, extensions, names, IDs, shape order,
relationships, and opaque package parts.

Exercise changes in both directions with fill, line, arrows, hyperlink,
shadow, transform, text, and preset-type edits. A same preset type preserves
adjustments; a changed preset type resets them to `[]`.

- [ ] **Step 3: Add malformed-source and lifecycle tests**

For every unsupported Task 1 fixture, require `shape.adjustments === undefined`.
Setter calls must throw `ModelParseError` with the slide URI and preserve exact
bytes/journal. Unrelated property edits must leave unsupported guide XML intact.

Prove duplicate isolation, outer-transaction rollback, stable `ShapeModel`
identity, source independence, six-format write/reopen, and preserved order/
values after repeated writes.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts -t 'shape adjustments' --reporter=dot
```

Expected: FAIL because live model accessors and slide delegation are absent.

- [ ] **Step 5: Implement model delegation transactionally**

Add to `ShapeModel`:

```ts
get adjustments(): readonly ShapeAdjustment[] | undefined {
  return this.slide.getShapeAdjustments(this.id);
}

set adjustments(value: readonly ShapeAdjustment[]) {
  this.slide.setShapeAdjustments(this.id, value);
}
```

Add to `SlideModel`:

```ts
getShapeAdjustments(id: number): readonly ShapeAdjustment[] | undefined {
  const { xml, element } = this.resolveShape(id);
  return readShapeAdjustments(xml, element);
}

setShapeAdjustments(id: number, value: readonly ShapeAdjustment[]): void {
  this.presentation.opcPackage.transaction(() => {
    const adjustments = normalizeShapeAdjustments(value, 'Shape adjustments');
    const { xml, element } = this.resolveShape(id);
    if (replaceShapeAdjustments(xml, element, adjustments, this.partUri)) {
      this.setXml(xml.serialize());
    }
  });
}
```

Skip `setXml()` on an exact no-op so the mutation journal stays unchanged.

- [ ] **Step 6: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/shape-adjustments.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape|shape adjustment|shape shadow' \
  packages/sdk/src/index.test.ts -t 'preset shape|shape adjustment' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; live object identity and exact no-ops hold.

- [ ] **Step 7: Review, commit, push, and verify**

Review public setter/getter types, detached snapshots, transaction placement,
no-op journal behavior, malformed zero mutation, property ownership, prefix
retention, preset-type interaction, duplicate isolation, rollback, stable
identity, and six-format reopen. Then:

```sh
git add -- packages/model/src/shapes.ts packages/model/src/slide.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit preset shape adjustments"
GIT_TERMINAL_PROMPT=0 git -c credential.interactive=never push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence after connectivity is available: `0 0`.

---

### Task 4: Add PptxGenJS conformance and packed public-surface coverage

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 3 public API and PptxGenJS 4.0.1 public `addShape()`/`write()`.
- Produces: real-output compatibility evidence and installed Node/browser/type/CLI verification.

- [ ] **Step 1: Add valid PptxGenJS public-output fixtures**

Generate named shapes through public APIs only:

- `roundRect` with `rectRadius: 0.5`, width 4, height 2 → `adj=25000`;
- `pie` with `angleRange: [270, 0]` → `adj1=16200000`, `adj2=0`;
- `arc` with fractional angles → exact rounded OOXML integer values;
- `blockArc` with angle range and `arcThicknessRatio: 0.5` → `adj3=25000`;
- omitted and explicit-zero shortcut cases that produce empty or partial lists.

Import actual written bytes, assert `ShapeModel.adjustments`, write with native,
reopen, and assert identical normalized snapshots. Create native paired shapes
with the expected direct integer lists and compare final supported state.

- [ ] **Step 2: Add explicit runtime-divergence probes**

Assert PptxGenJS drops zero `rectRadius` and zero thickness, accepts string
coercion and invalid negative/over-range values, gives `rectRadius` precedence,
and ignores thickness without angles. Assert native retains explicit zero guide
values, rejects non-number input, and accepts a deliberate final list without
shortcut precedence. Preserve malformed/unsafe imported XML even when its
snapshot is `undefined`.

- [ ] **Step 3: Extend the actual-tarball smoke**

In `scripts/smoke-npm-package.mjs`, create an adjusted block arc, assert the
immediate snapshot, edit one value, clear, restore, write/reopen, and assert
deep-frozen snapshots. Add compile-time source that imports `ShapeAdjustment`
and exercises `AddShapeOptions.adjustments` and `ShapeModel.adjustments`.

Require Node ESM, browser bundle, declaration compiler, and CLI smoke summary to
include `shapeAdjustments: true`.

- [ ] **Step 4: Run focused, packed, and build gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts -t 'adjustment|preset shape' --reporter=dot
pnpm typecheck
pnpm build
node --check scripts/smoke-npm-package.mjs
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-adjustment-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-adjustment-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected: public PptxGenJS fixtures, installed Node/browser/types/CLI, and build pass.

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only fixture generation, expected integer math, explicit-zero
differences, coercion/range evidence, adapter round-trip preservation, packed
dependency isolation, declaration reachability, browser parity, and CLI output.
Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify preset shape adjustment compatibility"
GIT_TERMINAL_PROMPT=0 git -c credential.interactive=never push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence after connectivity is available: `0 0`.

---

### Task 5: Document support and complete full release QA

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: verified behavior from Tasks 1–4.
- Produces: accurate public documentation, remaining-gap status, and final release evidence.

- [ ] **Step 1: Update public usage and capability documentation**

Add a concise example:

```ts
const arc = slide.addShape('blockArc', {
  adjustments: [
    { name: 'adj1', value: 16_200_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 25_000 },
  ],
});
arc.adjustments = [{ name: 'adj1', value: 10_800_000 }];
arc.adjustments = [];
```

Document direct integer units, ordered whole-list replacement, immutable
snapshots, exact no-op, malformed/complex-state behavior, preset-type reset,
PptxGenJS shortcut mappings/divergences, and the remaining custom-geometry/
arbitrary-formula scope. Remove generic "adjustment editing pending" wording.

- [ ] **Step 2: Run the complete automated gate**

```sh
pnpm check
pnpm test:performance
pnpm build
node --check scripts/smoke-npm-package.mjs
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-adjustment-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-adjustment-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected: all tests, typecheck, performance, build, installed package, and diff checks pass.

- [ ] **Step 3: Generate a representative legal gallery**

Using only the packed public package, create a wide deck containing labeled
round-rectangle, pie, arc, block-arc, and another adjustable preset. Include
empty, zero, fractional-source rounded, edit,
clear/restore, fill/line/arrows/shadow coexistence, and write/reopen cases. Keep
the deck in a task-specific `/tmp` directory, not the repository.

Use `pptx-inspect` stable JSON to verify titles, package structure, direct guide
lists, relationships, and compatibility diagnostics. Validate against the
PowerPoint 2010 profile with zero errors for every legal deck.

- [ ] **Step 4: Render and visually inspect every legal slide**

Open/export the gallery with an isolated LibreOffice profile. Reinspect the
round-tripped guide XML and record client normalization separately from native
behavior. Export to PDF, verify with `pdfinfo`, rasterize with `pdftoppm`, run
Presentations `render_slides.py` and `slides_test.py`, and inspect every slide at
full size. Fix any unexpected overlap, clipping, corrupt geometry, or rendering
regression before completion.

- [ ] **Step 5: Verify artifact-tool import preservation**

Import the original and LibreOffice-round-tripped galleries through
`@oai/artifact-tool` from JavaScript ES modules. Record slide/element counts,
names, bounds, and source hashes. Require no dropped adjusted shapes or
unexpected object identity changes.

- [ ] **Step 6: Review, commit, push, and verify**

Review docs against exact implemented behavior, all public names, remaining-gap
accuracy, PptxGenJS divergences, test counts, packed output, validator results,
render inspection, overflow, and artifact import evidence. Then:

```sh
git add -- CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document preset shape adjustments"
GIT_TERMINAL_PROMPT=0 git -c credential.interactive=never push origin main
git fetch origin main
git rev-list --left-right --count origin/main...HEAD
git status --short
```

Expected divergence after connectivity is available: `0 0`; only the existing
untracked `.pnpm-store/` cache remains. Preset-shape adjustments are complete,
and custom geometry becomes the next item.
