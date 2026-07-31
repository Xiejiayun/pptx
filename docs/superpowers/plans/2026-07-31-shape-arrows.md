# Shape Arrows Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking, and every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict native begin/end shape-arrow type creation and lossless direct-state editing for all six PptxGenJS 4.0.1 arrow types without changing the published `ShapeLine` ownership contract.

**Architecture:** Add a focused internal `shape-arrows` codec/adapter that normalizes detached endpoint snapshots, reads unique direct `a:headEnd` / `a:tailEnd`, and patches only those children inside the existing shape `a:ln` container. `AddShapeOptions.arrows` composes with the preset-shape renderer, while `ShapeModel.arrows` delegates read/edit through `SlideModel`; line width/fill/dash and arrow endpoints remain independent owners.

**Tech Stack:** TypeScript strict mode, Vitest, lossless source-span OOXML editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke tests, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- Public `ShapeArrowType` is exactly `none | arrow | diamond | oval | stealth | triangle`.
- Public `ShapeArrows` has optional readonly `begin` and `end`; it is a whole-replacement direct-state snapshot.
- Missing or runtime-`undefined` sides normalize to absence. An empty snapshot, both undefined sides, or a property assignment of `undefined` clears both endpoints.
- Inputs accept only ordinary or null-prototype objects with own data properties; reject arrays, class instances, symbols, accessors, inherited-only fields, unknown keys, aliases, null, empty strings, case variants, and unknown tokens before package mutation.
- Getter absence is `undefined`; explicit OOXML `type="none"` is the public token `'none'` and must not collapse into absence.
- Read only the unique direct `p:spPr/a:ln/a:headEnd` and `a:tailEnd`. Missing endpoints return `undefined`; malformed, namespace-lookalike, repeated, reversed, or ambiguous endpoint state is never exposed as supported.
- Endpoint attributes are one required unqualified canonical `type` plus optional unique unqualified `w` / `len` in `sm | med | lg`; endpoint children and other attributes are unsupported.
- Existing legal `w` / `len` survive same-value and type-replacement edits. New endpoints omit size. Clearing a side removes its whole endpoint child.
- `ShapeModel.arrows` owns only the direct endpoint children. It preserves line width/fill/dash, line attributes, join, extensions, unknown unrelated bytes, relationships, and other parts.
- `ShapeModel.line` keeps its existing promise: replacement or clear owns only width/fill/dash and preserves arrows.
- Creation omitted arrows leaves current line-only bytes unchanged. Arrows-only creation writes endpoints without synthesizing a color, width, or dash.
- Endpoint order is head then tail, after fill/dash/join and before `extLst`. Unsafe existing or insertion order throws `ModelParseError` before mutation.
- Same supported value is an exact bytes/journal no-op. Every failure and outer transaction rollback preserves exact slide bytes, relationships, parts, mutation journal, shape order, and live model identity.
- Do not add arrow size creation/editing, line cap/compound/alignment/join editing, advanced line fills/custom dash creation, shadow, hyperlink, adjustments, custom geometry, shape text options, percentage positions, or arrows on other drawing object APIs.
- Node 20+ and the browser bundle expose the same public behavior from `@jiayunxie/pptx`.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed, fetched, and verified at divergence `0 0`.
- Execute inline without subagents or routine decision pauses.

---

### Task 1: Add the internal arrow value and OOXML container adapter

**Files:**
- Create: `packages/model/src/shape-arrows.internal.ts`
- Create: `packages/model/src/shape-arrows.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlAttribute`, `XmlElement`, `ModelParseError`, the existing direct-shape/line OOXML conventions, and `a:ln` schema order.
- Produces: internal `ShapeArrowTypeValue`, `NormalizedShapeArrows`, `normalizeShapeArrows()`, `readShapeArrows()`, `renderShapeArrows()`, `replaceShapeArrows()`, and `shapeArrowsEqual()` for Tasks 2 and 3.

- [ ] **Step 1: Write descriptor-safe normalization tests**

Cover undefined, empty, begin-only, end-only, both, all six values, own undefined sides, null-prototype values, frozen output, and immediate caller detachment:

```ts
expect(normalizeShapeArrows(undefined, 'Shape arrows')).toBeUndefined();
expect(normalizeShapeArrows({}, 'Shape arrows')).toEqual({});
expect(normalizeShapeArrows({ begin: undefined, end: 'arrow' }, 'Shape arrows'))
  .toEqual({ end: 'arrow' });
expect(normalizeShapeArrows({ begin: 'triangle', end: 'oval' }, 'Shape arrows'))
  .toEqual({ begin: 'triangle', end: 'oval' });
```

Reject null/primitives/arrays/dates/class instances, inherited-only fields, own unknown string keys, all symbol keys, getters/setters without invoking them, invalid tokens, empty strings, case variants, and `beginArrowType` / `endArrowType` / `lineHead` / `lineTail` aliases.

- [ ] **Step 2: Write strict render/read tests**

Require deterministic endpoint order and alternate-prefix support:

```ts
expect(renderShapeArrows({ begin: 'triangle', end: 'arrow' }, 'a:')).toBe(
  '<a:headEnd type="triangle"/><a:tailEnd type="arrow"/>',
);
expect(renderShapeArrows({}, 'a:')).toBe('');
```

Read begin-only, end-only, both, explicit none, and endpoints with omitted/`med`/`sm`/`lg` size. Require detached output and no XML mutation. Return `undefined` for no endpoints, missing/unknown/qualified/repeated `type`, illegal/repeated/qualified `w` or `len`, endpoint children, wrong namespace, duplicate head/tail, reversed tail/head, nested endpoints, repeated `spPr`/`ln`, non-shape roots, and namespace lookalikes.

- [ ] **Step 3: Write lossless replacement and insertion tests**

Exercise exact same-value no-op, in-place type value replacement preserving lexical prefix plus legal size, whole-replacement clearing of one or both sides, insertion before a peer or `extLst`, and both-end creation after join:

```xml
<a:ln w="31750" cap="flat" data-keep="LINE">
  <a:gradFill><a:gsLst/></a:gradFill>
  <a:custDash><a:ds d="1" sp="1"/></a:custDash>
  <a:round/>
  <a:headEnd type="triangle" w="lg" len="sm"/>
  <a:tailEnd type="arrow"/>
  <a:extLst><a:ext uri="urn:keep"/></a:extLst>
</a:ln>
```

After endpoint edits, require the width, attributes, advanced fill/dash, join, legal size, extension, prefixes, and surrounding shape bytes to be identical. When `a:ln` is absent, create one after the unique shape fill and before effect/scene/3D/ext children. Clearing absent arrows must not create a line. Repeated/malformed endpoints and unsafe absent-line insertion must throw `ModelParseError` with exact source bytes unchanged.

- [ ] **Step 4: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/shape-arrows.internal.test.ts --reporter=dot
```

Expected: FAIL because `shape-arrows.internal.ts` does not exist.

- [ ] **Step 5: Implement the internal contract**

Define the exact internal surface:

```ts
export type ShapeArrowTypeValue =
  | 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle';

export interface NormalizedShapeArrows {
  readonly begin?: ShapeArrowTypeValue;
  readonly end?: ShapeArrowTypeValue;
}

export function normalizeShapeArrows(
  value: unknown,
  context: string,
): NormalizedShapeArrows | undefined;

export function renderShapeArrows(
  arrows: NormalizedShapeArrows | undefined,
  prefix: string,
): string;

export function readShapeArrows(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): NormalizedShapeArrows | undefined;

export function replaceShapeArrows(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  arrows: NormalizedShapeArrows | undefined,
  partUri: string,
): boolean;

export function shapeArrowsEqual(
  left: NormalizedShapeArrows | undefined,
  right: NormalizedShapeArrows | undefined,
): boolean;
```

Use descriptor inspection rather than property access. Freeze every normalized/read snapshot. Reader failures return `undefined`; replacement performs a separate strict analysis and throws before creating edits. For existing endpoints, replace only `type.valueStart..valueEnd`; remove a missing side by exact element span. For new endpoints, use the in-scope DrawingML prefix, insert head before tail and both before `extLst`, and never rewrite the rest of `a:ln`.

- [ ] **Step 6: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/shape-arrows.internal.test.ts \
  packages/model/src/shape-line.internal.test.ts \
  packages/model/src/simple-line.internal.test.ts --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  -t 'shape line|shape fill|preset shape' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; existing line setter still preserves endpoint bytes exactly.

- [ ] **Step 7: Review, commit, push, and verify**

Review descriptor safety, six-token exhaustiveness, strict endpoint attributes, legal size preservation, schema ordering, alternate prefixes, no-op behavior, unsafe insertion rejection, and absence of public exports. Then:

```sh
git add -- packages/model/src/shape-arrows.internal.ts \
  packages/model/src/shape-arrows.internal.test.ts
git diff --cached --check
git commit -m "feat: add shape arrow codec"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Create preset shapes with strict begin/end arrows

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `ShapeArrowTypeValue`, `NormalizedShapeArrows`, `normalizeShapeArrows()`, and `renderShapeArrows()` plus the existing preset-shape normalizer/renderer.
- Produces: public `ShapeArrowType`, public `ShapeArrows`, `AddShapeOptions.arrows?: ShapeArrows`, and detached normalized creation state rendered into the existing shape line container.

- [ ] **Step 1: Add public compile-time and runtime creation tests**

Add exported-type assertions and create shapes with omitted, runtime-undefined, empty, begin-only, end-only, both, explicit none, and every token:

```ts
const arrows = slide.addShape('line', {
  name: 'Native arrows',
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
});
const arrowOnly = slide.addShape('lineInv', {
  arrows: { end: 'stealth' },
});
const explicitNone = slide.addShape('line', {
  arrows: { begin: 'none', end: 'none' },
});
```

Require canonical head/tail order, no implicit line state for arrow-only creation, unchanged line-only bytes, exact line+arrow composition, returned `ShapeModel`, stable shape order/identity, caller detachment, and all 178 geometries remaining creatable.

- [ ] **Step 2: Add invalid and zero-mutation creation tests**

For every Task 1 invalid input, snapshot part bytes, relationships, parts, shape models, next shape ID, and mutation journal before `addShape()`; require identical state after rejection. Add top-level `arrows` and nested endpoint accessors to prove getters are never invoked. Require `beginArrowType`, `endArrowType`, `lineHead`, and `lineTail` to fail as unsupported `AddShapeOptions` keys or arrow keys.

- [ ] **Step 3: Extend exact preset normalization and renderer tests**

Require `normalizePresetShape()` to preserve omitted arrows as `undefined`, keep `{}` as an empty frozen normalized snapshot, detach explicit endpoints, and accept only the new `arrows` key:

```ts
export interface NormalizedPresetShape {
  readonly arrows: NormalizedShapeArrows | undefined;
}
```

The renderer must keep `<a:ln/>` for no line/no arrows, output `<a:ln><a:headEnd .../></a:ln>` for arrows-only, append arrows after no-fill for native none, and append arrows after preset dash for solid line. Include a compile-time bidirectional assignability assertion between public `ShapeArrowType` and internal `ShapeArrowTypeValue`.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape arrows' \
  packages/sdk/src/index.test.ts -t 'preset shape arrows' --reporter=dot
```

Expected: FAIL because public arrow types and `AddShapeOptions.arrows` are absent.

- [ ] **Step 5: Implement the public contract and creation renderer**

Add the public types exactly:

```ts
export type ShapeArrowType =
  | 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle';

export interface ShapeArrows {
  readonly begin?: ShapeArrowType;
  readonly end?: ShapeArrowType;
}

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly arrows?: ShapeArrows;
}
```

Add `arrows` to `OPTION_KEYS`, normalize it with `normalizeShapeArrows(values.arrows, 'Preset shape arrows')`, and include it in the frozen record. Change only the private preset-line renderer signature:

```ts
function renderPresetLine(
  line: NormalizedSimpleLine | undefined,
  arrows: NormalizedShapeArrows | undefined,
): string;
```

Build line-owned children first, append `renderShapeArrows(arrows, 'a:')`, use a self-closing line only when both are empty, and write `w` only for a solid line. No other shape skeleton bytes change.

- [ ] **Step 6: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape' \
  packages/sdk/src/index.test.ts -t 'preset shape' --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  -t 'shape line|shape fill|duplicate slide|rollback' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; omitted arrows leave all prior creation fixtures byte-identical.

- [ ] **Step 7: Review, commit, push, and verify**

Review public names, whole-snapshot semantics, no hidden line defaults, exact child order, line-only byte stability, descriptor-safe zero mutation, all geometry regression, and generated declarations. Then:

```sh
git add -- packages/model/src/preset-shape.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create shape arrows"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Expose live arrow reading, replacement, and clearing

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 internal normalize/read/replace functions and Task 2 public `ShapeArrows`.
- Produces: `ShapeModel.arrows`, `SlideModel.getShapeArrows()`, and `SlideModel.setShapeArrows()` with transaction, stable identity, no-op, and rollback behavior.

- [ ] **Step 1: Add live-model getter and editor tests**

Require immediate snapshots on newly created preset shapes and existing text-box `ShapeModel` instances; two getter calls must be equal but not the same object. Cover both→begin-only→end-only→explicit-none→empty→undefined→both, and prove missing side whole-replacement clearing.

```ts
expect(shape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
shape.arrows = { begin: 'diamond' };
expect(shape.arrows).toEqual({ begin: 'diamond' });
shape.arrows = undefined;
expect(shape.arrows).toBeUndefined();
```

Require same-value exact bytes/journal no-op, caller detachment, stable model identity, shape order, neighbor bytes, relationships, and parts.

- [ ] **Step 2: Add ownership-isolation and advanced-line tests**

Inject a legal existing line with gradient/picture/pattern/group fill, custom dash, cap/compound/alignment attributes, join, sized endpoints, extension, and opaque sibling bytes. Change endpoint types and require only the `type` value spans to differ. Then clear arrows and require line width/fill/dash/attributes/join/ext to remain exact. Separately set and clear `shape.line` and require endpoint type/size bytes to remain exact.

- [ ] **Step 3: Add malformed, lifecycle, and rollback tests**

For missing type, invalid token, illegal size, unknown attribute, endpoint child, wrong namespace, duplicate head/tail, tail-before-head, repeated `ln`/`spPr`, and unsafe insertion order, require setter `ModelParseError` with exact part/relationship/parts/journal state. Cover absent-line insertion, duplicate-slide independence, outer transaction rollback, write/reopen, second write, move, and all six formats.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/model/src/model.test.ts \
  -t 'shape arrows' --reporter=dot
pnpm vitest run packages/sdk/src/index.test.ts \
  -t 'shape arrows' --reporter=dot
```

Expected: FAIL because `ShapeModel.arrows` and the slide delegates are absent.

- [ ] **Step 5: Implement the live public surface**

Add the model property:

```ts
get arrows(): ShapeArrows | undefined {
  return this.slide.getShapeArrows(this.id);
}

set arrows(value: ShapeArrows | undefined) {
  this.slide.setShapeArrows(this.id, value);
}
```

Add slide delegates using the exact transaction pattern:

```ts
getShapeArrows(id: number): ShapeArrows | undefined {
  const { xml, element } = this.resolveShape(id);
  return readShapeArrows(xml, element);
}

setShapeArrows(id: number, value: ShapeArrows | undefined): void {
  const arrows = value === undefined
    ? undefined
    : normalizeShapeArrows(value, 'Shape arrows');
  this.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolveShape(id);
    if (replaceShapeArrows(xml, element, arrows, this.partUri)) {
      this.setXml(xml.serialize());
    }
  });
}
```

Keep normalization before transaction and package mutation. Do not add arrows to `addText()` creation options or other drawing model classes.

- [ ] **Step 6: Run focused, package, and regression gates**

```sh
pnpm vitest run packages/model/src/shape-arrows.internal.test.ts \
  packages/model/src/model.test.ts -t 'shape arrows|shape line|shape fill' \
  packages/sdk/src/index.test.ts -t 'shape arrows|rollback' --reporter=dot
pnpm vitest run packages/model/src/model.test.ts \
  -t 'preset shape|duplicate slide|move slide|write and reopen' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; arrow and line editors preserve each other's ownership.

- [ ] **Step 7: Review, commit, push, and verify**

Review strict namespace ownership, whole-replacement semantics, sized-arrow preservation, bidirectional line isolation, no-op/journal behavior, stable identity, duplicate independence, six formats, and rollback. Then:

```sh
git add -- packages/model/src/slide.ts packages/model/src/shapes.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit shape arrows"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public `PptxGenJS.ShapeType`, `slide.addShape()`, `write()`, native `ShapeLine`/`ShapeArrows`, creation/read/edit APIs, and direct OOXML inspection.
- Produces: public-only evidence for all supported endpoint semantics and explicit evidence for intentional strict/default divergences.

- [ ] **Step 1: Add a public-output arrow fixture matrix**

Generate one shape per case using only public PptxGenJS APIs: every begin token, every end token, both ends, arrow-only defaults, `{ type: 'none' }` with arrow, empty strings, nested deprecated aliases, top-level deprecated aliases, and runtime invalid passthrough. Capture packages through public `write()` without importing PptxGenJS internals.

- [ ] **Step 2: Compare supported semantic state**

For every six-token begin/end case, open public output with `PptxDocument`, compare shape name/geometry/transform, `ShapeModel.line`, `ShapeModel.arrows`, and direct child order to native equivalents. Verify native type replacement and clear on imported PptxGenJS output, write/reopen, and preservation of line width/fill/dash.

- [ ] **Step 3: Assert strict/default divergences explicitly**

Require PptxGenJS arrow-only output to materialize `333333`/1pt/solid, while native arrows-only output owns no line state. Require an explicitly configured native default line plus arrows to match that final supported state. Require PptxGenJS `type: 'none'` plus arrow and native arrows-only creation to match endpoint-only state; native explicit line none intentionally writes direct `a:noFill`. Require PptxGenJS empty arrow values to disappear and invalid tokens to pass through, while native rejects both before mutation. Record nested alias ignore and top-level alias mapping without accepting either alias natively.

- [ ] **Step 4: Run conformance and dependency gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'shape arrow public output' --reporter=dot
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; `pptxgenjs` remains absent from non-adapter manifests.

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only generation, all twelve single-end cases, both-end order, default-line mapping, explicit-none mapping, strict invalid behavior, alias evidence, edit/reopen, and dependency boundaries. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare shape arrows with pptxgenjs"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: the actual `@jiayunxie/pptx` tarball, Node ESM export, browser conditional export, generated declarations, and current package smoke harness.
- Produces: tarball-level `shapeArrows: true` evidence in the smoke JSON.

- [ ] **Step 1: Extend the installed-package Node smoke**

In generated `smoke.mjs`, create begin/end/both/explicit-none/all-token arrows, mutate original caller objects, read detached snapshots, perform same-value/whole-replacement/clear edits, preserve line state, inject legal `w="lg" len="sm"` plus advanced line siblings, duplicate, write/reopen, and require zero relationship/part growth. Add `shapeArrows` to the final JSON envelope without removing existing fields.

- [ ] **Step 2: Extend browser and declaration smoke**

In browser smoke, import `PptxDocument`/`ShapeModel` from the package browser condition, create/read/edit/clear/reopen both endpoints with a solid dashed line, and verify no `node:` dependency. In generated typecheck source, import `ShapeArrowType`/`ShapeArrows`, assign all valid tokens, use `AddShapeOptions.arrows` and `ShapeModel.arrows`, and require invalid token/empty/alias/unknown-key cases with `@ts-expect-error`.

- [ ] **Step 3: Build, pack, and run the actual tarball smoke**

```sh
pnpm build
mkdir -p /tmp/pptx-shape-arrows-pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-shape-arrows-pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-arrows-pack/jiayunxie-pptx-0.1.0.tgz
```

Expected JSON includes `"shapeArrows":true`, `"shapeLines":true`, `"shapeFills":true`, `"presetShapes":true`, `"types":true`, and CLI version `0.1.0`.

- [ ] **Step 4: Run script and repository gates**

```sh
node --check scripts/smoke-npm-package.mjs
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 5: Review, commit, push, and verify**

Review installed-tarball isolation, Node/browser parity, six-token type coverage, endpoint detachment, line isolation, legal-size preservation, existing smoke preservation, temporary cleanup, and absence of internal imports. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed shape arrows"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Document the supported boundary and remaining shape gaps

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: Tasks 1–5 final public API and conformance evidence.
- Produces: user-facing creation/edit examples, changelog entry, API contract, and an honest compatibility-matrix update.

- [ ] **Step 1: Add concise public examples**

Document exact native syntax:

```ts
const shape = slide.addShape('line', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
});

shape.arrows = { begin: 'diamond' }; // clears the omitted end
shape.arrows = { begin: 'none', end: 'oval' };
shape.arrows = undefined; // clears both endpoints, preserves line style
```

Explain explicit none versus absence, whole-replacement sides, arrows-only line behavior, detached snapshots, legal size preservation, and bidirectional line/arrows ownership isolation.

- [ ] **Step 2: Update compatibility, API, and changelog**

Change the `slide.addShape()` row to include six begin/end arrow types across native creation/read/edit/clear/duplicate/reopen. Record PptxGenJS arrow-only implicit line defaults, native arrows-only direct state, explicit-none mapping, invalid passthrough, and deprecated-alias behavior without claiming byte identity. Remove arrow type from the pending list while keeping arrow size, cap/compound/alignment/join, advanced line fill/custom dash creation, shadow, hyperlink, adjustments, custom geometry, shape text, percentage positions, and broader drawing objects explicitly remaining. Change the next shape priority from arrows to hyperlink.

- [ ] **Step 3: Validate examples and documentation consistency**

```sh
rg -n "ShapeArrows|ShapeArrowType|shape\.arrows|begin|end|arrow size|hyperlink" \
  README.md packages/pptx/README.md CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
git diff --check
```

Expected: examples match exported types, explicit none and clear are not conflated, and all remaining gaps remain visible.

- [ ] **Step 4: Review, commit, push, and verify**

Review API spelling, whole-replacement semantics, arrows-only default mapping, line ownership, size wording, strict divergences, remaining-gap accuracy, and parity between root/package/API docs. Then:

```sh
git add -- README.md packages/pptx/README.md CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document shape arrows"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run release, real-PPTX, structural, visual, and import QA

**Files:**
- No repository file changes expected; use `/tmp/pptx-shape-arrows-qa-20260731`.

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
mkdir -p /tmp/pptx-shape-arrows-qa-20260731/pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-shape-arrows-qa-20260731/pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-arrows-qa-20260731/pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: identity/version, Node/browser/types, preset shapes, shape fills, shape lines, shape arrows, dependency boundaries, and CLI smoke all pass.

- [ ] **Step 3: Generate representative lifecycle decks**

Create native decks for omitted/empty, begin/end/both, explicit none, all six types, arrows-only, none-line arrows, sRGB/scheme/transparent/dashed line combinations, sized existing endpoints, advanced-line preservation, whole replacement, clear, malformed zero-mutation source, duplicate, rollback, reopen, and second write. Create PptxGenJS legal comparison and strict/default-divergence fixtures with labeled visible lines.

- [ ] **Step 4: Validate every legal package and mutation boundary**

For each legal deck run:

```sh
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json slides list deck.pptx
pptx-inspect --json part read deck.pptx /ppt/slides/slide1.xml
```

Use `package diff` to prove source→edit changes only the target slide part and edit→reopen→second-write has zero part changes. Verify exact head/tail tokens/order, preserved legal size, preserved width/fill/dash/join/extensions, shape IDs/order/names, and neighbor bytes. Require zero PowerPoint 2010 profile errors and warnings for every legal fixture.

- [ ] **Step 5: Render and inspect every slide**

Open/export all legal decks with an isolated LibreOffice profile, run `pdfinfo`, render PDF pages with `pdftoppm`, then run `render_slides.py` and `slides_test.py` for every PPTX. Inspect every rendered page at full size for all arrow silhouettes, begin/end orientation, explicit none, line visibility, size preservation, theme/transparency/dash composition, clipping, overlap, and label readability.

- [ ] **Step 6: Verify artifact-tool import behavior**

Import native representative/lifecycle decks with `PresentationFile.importPptx()`, require expected slide/shape counts, and inspect named shapes. Import PptxGenJS comparison fixtures and confirm supported files load without mutation.

- [ ] **Step 7: Verify repository and remote final state**

```sh
git diff --check
git status --short
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: only the pre-existing untracked `.pnpm-store/` remains ignored, and divergence is `0 0`. Create no empty commit when Task 7 produces no repository change.
