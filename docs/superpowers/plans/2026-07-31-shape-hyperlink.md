# Shape Hyperlink Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking, and every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict native whole-shape URL/internal-slide hyperlink creation and lossless direct-state editing compatible with valid PptxGenJS 4.0.1 `HyperlinkProps` output.

**Architecture:** Add a focused `shape-hyperlink` value/OOXML adapter for descriptor-safe normalization, `p:cNvPr/a:hlinkClick` inspection and source-span edits. `SlideModel.addShape()` creates the matching OPC relationship before rendering the shape, while `ShapeModel.hyperlink` coordinates reference-aware relationship update/clone/GC; presentation slide deletion removes DrawingML references to the deleted target before generic incoming-relationship cleanup.

**Tech Stack:** TypeScript strict mode, Vitest, lossless source-span OOXML editing, OPC relationships/transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke tests, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- Public `Hyperlink` is a mutually exclusive union with exactly one defined target: non-empty string `url` or one-based positive safe-integer `slide`; optional `tooltip` is an XML-safe string and may be empty.
- Omitted and own runtime-`undefined` fields collapse to absence. Missing targets, both defined targets, zero/negative/fractional/out-of-range slides, empty URLs, coercions, aliases, symbols, accessors, inherited fields, and unknown keys fail before package mutation.
- `slide` resolves at assignment time to an actual presentation slide part. Getter recalculates its current one-based ordinal from the relationship target, so order changes never retarget the link.
- Getter snapshots are detached and frozen. Tooltip attribute absence is omitted; direct `tooltip=""` returns explicit empty and is not collapsed.
- Read only the unique direct DrawingML `hlinkClick` under the shape's unique PresentationML non-visual properties. URL requires the exact external hyperlink relationship and no action; internal slide requires the exact internal slide relationship and `ppaction://hlinksldjump`.
- Existing unsupported, dangling, malformed, namespace-lookalike, duplicate, or ambiguous link state is not exposed as supported and is rejected by mutation with `ModelParseError`.
- `ShapeModel.hyperlink` owns only the direct click target/action/tooltip and its relationship reference. It preserves `hlinkHover`, extra legal click attributes/children, non-visual metadata, shape geometry/fill/line/arrows/effects/text, neighbor XML, and unrelated parts/relationships.
- Same supported target and tooltip is an exact bytes/journal no-op. Tooltip-only edits never touch the relationship.
- Target replacement updates an unshared relationship in place and clone-on-writes a shared relationship. Clear/retarget removes an old relationship only when no relationship-namespace ID attribute in the source slide XML references it.
- Duplicate preserves external/other-slide links and retargets self-links to the duplicate. Move/insert/delete before a target changes only the getter ordinal. Target deletion removes DrawingML click/hover references before the target part and incoming relationship disappear.
- Creation omitted/runtime-undefined hyperlink leaves the published preset-shape XML and relationship bytes unchanged.
- External URL validation may report the expected `OPC_EXTERNAL_RELATIONSHIP` portability warning; it must never report an error. Internal links must validate with zero errors and zero warnings.
- Do not add text-run, table-cell, image, chart, media, group, graphic-frame, hover, custom-show, macro, program, action-only, sound, history, or highlight hyperlink APIs in this plan.
- Transitional OOXML relationship namespaces remain the supported package contract; Strict-package relationship URIs wait for the unified strict-package audit.
- Node 20+ and the browser bundle expose identical public behavior from `@jiayunxie/pptx`.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed, fetched, and verified at divergence `0 0`.
- Execute inline without subagents or routine decision pauses.

---

### Task 1: Add the strict hyperlink value and direct OOXML adapter

**Files:**
- Create: `packages/model/src/shape-hyperlink.internal.ts`
- Create: `packages/model/src/shape-hyperlink.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlAttribute`, `XmlElement`, `Relationship`, `ModelParseError`, DrawingML/PresentationML/relationship namespace rules, and current source-span mutation conventions.
- Produces: `NormalizedHyperlink`, `ShapeHyperlinkReadContext`, `normalizeHyperlink()`, `readShapeHyperlink()`, `renderShapeHyperlink()`, `replaceShapeHyperlinkElement()`, `shapeHyperlinksEqual()`, `relationshipReferenceCount()`, and `removeDrawingHyperlinkReferences()` for Tasks 2 and 3.

- [ ] **Step 1: Write descriptor-safe value tests**

Cover URL, internal slide, tooltip absent/empty/Unicode/XML metacharacters, own undefined fields, null-prototype input, frozen output, and immediate caller detachment:

```ts
expect(normalizeHyperlink({ url: 'https://example.com?a=1&b=2' }, 'Shape hyperlink'))
  .toEqual({ url: 'https://example.com?a=1&b=2' });
expect(normalizeHyperlink({ slide: 2, tooltip: '' }, 'Shape hyperlink'))
  .toEqual({ slide: 2, tooltip: '' });
expect(normalizeHyperlink({ url: undefined, slide: 1 }, 'Shape hyperlink'))
  .toEqual({ slide: 1 });
```

Reject null/primitives/arrays/dates/class instances, inherited-only fields, own unknown string keys, all symbol keys, getters/setters without invocation, missing targets, both defined targets, empty URL, numeric URL, slide zero/negative/fraction/unsafe integer/NaN/infinity/string/boolean, invalid XML controls, invalid tooltip, and `_rId`/`target`/`kind` aliases.

- [ ] **Step 2: Write strict external/internal reader tests**

Use real namespace URIs and a read context containing exact relationships plus ordered slide part URIs:

```ts
export type NormalizedHyperlink = Readonly<
  | { readonly url: string; readonly slide?: never; readonly tooltip?: string }
  | { readonly slide: number; readonly url?: never; readonly tooltip?: string }
>;

export interface ShapeHyperlinkReadContext {
  readonly relationships: readonly Relationship[];
  readonly slidePartUris: readonly string[];
}
```

Read URL with/without/empty tooltip, internal slide/self-link, and alternate legal DrawingML/relationship prefixes. Require current slide ordinal mapping, frozen detached snapshots, no XML/package mutation, and preservation of direct empty tooltip.

Return `undefined` for no click, wrong shape root, repeated non-visual containers/clicks, wrong namespace, missing/repeated/qualified tooltip or action, missing/repeated relationship ID, missing/duplicate relationship ID, empty external target, wrong type/mode, URL with action, internal slide without exact action, dangling slide target, relationship targets outside the presentation slide list, and a target part that appears at multiple presentation ordinals.

- [ ] **Step 3: Write render and lossless element replacement tests**

Require exact target-specific output:

```ts
expect(renderShapeHyperlink(
  { url: 'https://example.com', tooltip: '' },
  'rId7',
  { drawing: 'a', relationship: 'r' },
)).toBe('<a:hlinkClick r:id="rId7" tooltip=""/>');

expect(renderShapeHyperlink(
  { slide: 2, tooltip: 'Next' },
  'rId8',
  { drawing: 'a', relationship: 'r' },
)).toBe(
  '<a:hlinkClick r:id="rId8" tooltip="Next" action="ppaction://hlinksldjump"/>',
);
```

Start from a click containing `history`, `highlightClick`, `tgtFrame`, a sound child, an extension child, alternate prefixes, a sibling `hlinkHover`, and surrounding unknown bytes. Verify same element value is exact no-op; URL/slide switching patches only relationship ID/action/tooltip; tooltip add/replace/empty/remove preserves extra state; clear removes only `hlinkClick`; absent creation expands a self-closing `cNvPr` or inserts before hover/extLst without reserializing neighbors.

- [ ] **Step 4: Write reference counting and target-deletion cleanup tests**

Count only expanded-name-correct relationship ID attributes across the whole slide XML, including shape, text-run, image, and opaque containers. Ignore same-local-name attributes in wrong namespaces.

For `removeDrawingHyperlinkReferences()`, remove DrawingML `hlinkClick` and `hlinkHover` elements whose relationship ID is in the supplied set, preserve all other elements and bytes, and return false without mutation when none match. Cover shared ID references, alternate prefixes, self-closing/expanded elements, target and non-target links side by side, and nested sound/ext children.

- [ ] **Step 5: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/shape-hyperlink.internal.test.ts --reporter=dot
```

Expected: FAIL because `shape-hyperlink.internal.ts` does not exist.

- [ ] **Step 6: Implement the minimal internal contract**

Define the exact surface:

```ts
export interface HyperlinkPrefixes {
  readonly drawing: string;
  readonly relationship: string;
}

export function normalizeHyperlink(
  value: unknown,
  context: string,
): NormalizedHyperlink;

export function readShapeHyperlink(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  context: ShapeHyperlinkReadContext,
): NormalizedHyperlink | undefined;

export function renderShapeHyperlink(
  hyperlink: NormalizedHyperlink,
  relationshipId: string,
  prefixes: HyperlinkPrefixes,
): string;

export function replaceShapeHyperlinkElement(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  hyperlink: NormalizedHyperlink | undefined,
  relationshipId: string | undefined,
  partUri: string,
): boolean;

export function shapeHyperlinksEqual(
  left: NormalizedHyperlink | undefined,
  right: NormalizedHyperlink | undefined,
): boolean;

export function relationshipReferenceCount(
  xml: LosslessXmlDocument,
  relationshipId: string,
): number;

export function removeDrawingHyperlinkReferences(
  xml: LosslessXmlDocument,
  relationshipIds: ReadonlySet<string>,
): boolean;
```

Use descriptors rather than public property reads. Freeze normalized/read values. Keep read as a non-throwing supported-state probe; replacement performs its own strict inspection and throws `ModelParseError` before creating edits. Resolve namespaces by expanded name, preserve existing lexical prefixes, and add local namespace declarations only when no usable in-scope prefix exists.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/shape-arrows.internal.test.ts \
  packages/model/src/shape-line.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; the internal file has no public export and no package relationship mutation.

- [ ] **Step 8: Review, commit, push, and verify**

Review descriptor safety, exactly-one-target enforcement, XML character checks, one-based slide mapping, exact relationship type/mode/action matching, prefix handling, tooltip absence/empty distinction, extra-state preservation, reference counting, cleanup isolation, no-op behavior, and zero mutation on malformed input. Then:

```sh
git add -- packages/model/src/shape-hyperlink.internal.ts \
  packages/model/src/shape-hyperlink.internal.test.ts
git diff --cached --check
git commit -m "feat: add shape hyperlink codec"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Create preset shapes with strict URL and slide hyperlinks

**Files:**
- Create: `packages/model/src/hyperlink.ts`
- Modify: `packages/model/src/index.ts`
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `NormalizedHyperlink`, `normalizeHyperlink()`, and `renderShapeHyperlink()` plus `OpcPackage.addRelationship()`, `relativeRelationshipTarget()`, current preset-shape normalization/rendering, and the ordered presentation slide models.
- Produces: public `Hyperlink`, `AddShapeOptions.hyperlink?: Hyperlink`, normalized creation state, exact external/internal relationships, and valid click XML that Task 3 exposes through `ShapeModel.hyperlink`.

- [ ] **Step 1: Add public compile-time and runtime creation tests**

Export and exercise the exact type:

```ts
export type Hyperlink =
  | {
      readonly url: string;
      readonly slide?: never;
      readonly tooltip?: string;
    }
  | {
      readonly slide: number;
      readonly url?: never;
      readonly tooltip?: string;
    };

const website = first.addShape('rect', {
  name: 'Website',
  hyperlink: { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' },
});
const next = first.addShape('actionButtonForwardNext', {
  name: 'Next slide',
  hyperlink: { slide: 2, tooltip: '' },
});
const self = first.addShape('actionButtonHome', {
  hyperlink: { slide: 1 },
});
```

Require URL/internal/self target semantics, tooltip absence/empty distinction, XML escaping, exact relationship type/target/mode, shape order/ID/stable identity, caller detachment after creation, and all 178 preset geometries remaining creatable. Inspect created click state through Task 1 `readShapeHyperlink()` until Task 3 adds the public getter.

- [ ] **Step 2: Prove omitted behavior and invalid zero mutation**

Require omitted and runtime-undefined hyperlink creation to preserve the previously published preset-shape bytes and relationship list exactly. For every Task 1 invalid value plus a valid-range slide that becomes out of range, snapshot slide bytes, all parts, relationships, shape-model map behavior, next shape ID, and mutation journal before `addShape()`; require identical state after rejection.

Add compile-time `@ts-expect-error` cases for empty object, both URL/slide, numeric URL, string slide, unknown key, `_rId`, and invalid tooltip. Runtime tests must prove URL/slide/tooltip accessors are never invoked.

- [ ] **Step 3: Extend preset normalization and rendering tests**

Add `hyperlink` to the only accepted shape option keys and normalized record:

```ts
export interface NormalizedPresetShape {
  readonly hyperlink: NormalizedHyperlink | undefined;
}

export function renderPresetShapeXml(
  id: number,
  shape: NormalizedPresetShape,
  hyperlinkRelationshipId?: string,
): string;
```

Require no-link bytes to remain exact, URL click to render without action, internal click to render with exact action, tooltip omission/empty to stay distinct, and link markup to expand only `p:cNvPr`. Reject a normalized hyperlink without a relationship ID and an ID without a normalized hyperlink.

- [ ] **Step 4: Write relationship creation and rollback tests**

For external URL, require exact hyperlink type, raw decoded target, and `External` mode. For internal slide, require exact slide type, relative target to the resolved target part, `Internal` mode, and no dangling graph edge. Create links when layout/notes/opaque relationships already occupy non-contiguous IDs and require `allocateRelationshipId()` behavior.

Wrap creation in an outer transaction that throws after `addShape()` and require exact rollback of shape XML, relationship part creation/update, graph, mutation journal, next ID, and returned live identity cache.

- [ ] **Step 5: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape hyperlink' \
  packages/sdk/src/index.test.ts -t 'preset shape hyperlink' --reporter=dot
```

Expected: FAIL because the public `Hyperlink` type and creation integration are absent.

- [ ] **Step 6: Implement public export, creation normalization, relationship creation, and rendering**

Create `hyperlink.ts`, export it from the model barrel, and import it into `preset-shape.ts`. Normalize `values.hyperlink` through Task 1. In `SlideModel.addShape()` resolve a slide target from `this.presentation.slides[hyperlink.slide - 1]` before adding a relationship.

Use exact constants:

```ts
const HYPERLINK_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const SLIDE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
```

Create the relationship inside the existing transaction, pass its ID to `renderPresetShapeXml()`, append the shape, and resolve the returned `ShapeModel` exactly as before. Do not share relationships between new shapes implicitly.

- [ ] **Step 7: Run focused and regression gates**

```sh
pnpm vitest run packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/model.test.ts -t 'preset shape|shape hyperlink' \
  packages/sdk/src/index.test.ts -t 'preset shape|shape hyperlink' --reporter=dot
pnpm vitest run packages/model/src/shape-fill.internal.test.ts \
  packages/model/src/shape-line.internal.test.ts \
  packages/model/src/shape-arrows.internal.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; omitted-link shape bytes and existing fill/line/arrows behavior remain exact.

- [ ] **Step 8: Review, commit, push, and verify**

Review public union exclusivity, export reachability, option-key strictness, target range validation before mutation, external/internal relationship correctness, local namespace declarations, XML escaping, no-link byte stability, rollback, identity, and absence of text/image/table APIs. Then:

```sh
git add -- packages/model/src/hyperlink.ts packages/model/src/index.ts \
  packages/model/src/preset-shape.ts packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts packages/model/src/slide.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create shape hyperlinks"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Read, edit, clear, duplicate, and clean up shape hyperlinks

**Files:**
- Modify: `packages/model/src/shape-hyperlink.internal.ts`
- Modify: `packages/model/src/shape-hyperlink.internal.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Tasks 1–2 public/normalized values, direct-element operations, relationship reference count/cleanup, `OpcPackage.updateRelationship()`/`addRelationship()`/`removeRelationship()`, slide dependency cloning, and presentation delete/move/duplicate transactions.
- Produces: `ShapeModel.hyperlink`, `SlideModel.getShapeHyperlink()`, `SlideModel.setShapeHyperlink()`, reference-aware target replacement/GC, and target-slide deletion cleanup.

- [ ] **Step 1: Add getter/setter public contract tests**

Expose exact methods:

```ts
class ShapeModel {
  get hyperlink(): Hyperlink | undefined;
  set hyperlink(value: Hyperlink | undefined);
}

class SlideModel {
  getShapeHyperlink(id: number): Hyperlink | undefined;
  setShapeHyperlink(id: number, value: Hyperlink | undefined): void;
}
```

Test imported and native URL/internal/self links, frozen detached getter values, current ordinal after slide moves, exact same-value no-op, tooltip add/replace/explicit-empty/remove, URL replacement, slide replacement, URL↔slide switching, and whole clear. Require `ShapeModel` identity to remain stable after every edit.

- [ ] **Step 2: Add relationship update, clone-on-write, and GC tests**

Build a slide where two shapes and a text run share one URL relationship ID. Changing only one shape target must allocate one new relationship, retarget only that click, and leave the other two references unchanged. Changing a unique relationship may retain its ID. Tooltip-only edits must retain relationship bytes exactly.

Clear the first shape and require the shared relationship to remain for the second shape and text run. Clear the second shape and require it to remain for the text run. Remove the run hyperlink from the fixture by exact XML element span and then replace/clear a new unique second-shape link; require the newly allocated relationship to disappear only after its final correct reference is gone.

Require wrong-namespace `id` attributes not to retain a relationship, while any correct relationship ID attribute anywhere in slide XML does retain it. Verify relationship part deletion is not required when unrelated layout/notes relationships remain.

- [ ] **Step 3: Add lossless preservation and malformed zero-change tests**

Inject a supported click with hover, non-owned attributes, sound/ext children, alternate prefixes, `descr/title/hidden` non-visual metadata, custom geometry adjustments, fill, line/arrows, effects, rich text, neighbor shapes, opaque relationships, and extension lists. Replace target/tooltip and clear; require only owned click/relationship spans to change.

For each unsupported state from Task 1, snapshot slide bytes, relationship bytes/list, parts, graph, model identity, and journal before setter calls; require `ModelParseError` and exact equality afterward. A getter returns `undefined` without mutation. Invalid public value rejection must happen before parsing or mutation.

- [ ] **Step 4: Add duplicate, move, insert/delete, target deletion, rollback, and format tests**

Cover:

- duplicate source with external link;
- duplicate source linking another slide;
- duplicate self-link retargeting to duplicate itself;
- move source and target, with getter ordinal update and unchanged relationship target part;
- add/delete a slide before target, preserving target identity;
- delete source, preserving target and other slides;
- delete target, removing matching DrawingML click/hover references and incoming slide relationship while preserving unrelated links/bytes;
- target relationship shared by shape/text/image link elements, all matching elements removed on target deletion;
- outer rollback after edit/clear/delete target;
- write/reopen/second write in `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`.

After target deletion, require validator zero dangling relationships and `shape.hyperlink === undefined` for cleaned shapes.

- [ ] **Step 5: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/model/src/model.test.ts \
  -t 'shape hyperlink lifecycle' --reporter=dot
pnpm vitest run packages/sdk/src/index.test.ts \
  -t 'shape hyperlink lifecycle' --reporter=dot
```

Expected: FAIL because `ShapeModel.hyperlink` and lifecycle coordination are absent.

- [ ] **Step 6: Implement model delegation and relationship-aware replacement**

Normalize public input before transaction. Inside `setShapeHyperlink()` parse/resolve the shape, inspect supported current state, compare semantic target part plus direct tooltip, and return before mutation on equality.

For target changes:

1. Count the current ID's correct relationship references in the original slide XML.
2. If count is one and the relationship is unique, update type/target/mode in place.
3. Otherwise create a new relationship and pass its ID to `replaceShapeHyperlinkElement()`.
4. Set changed slide XML.
5. Remove an old relationship only when the updated XML reference count is zero.

For tooltip-only changes, patch only the click. For clear, remove the click, set XML, and remove the relationship only after the updated reference count reaches zero. Keep all operations in one OPC transaction.

- [ ] **Step 7: Implement target deletion cleanup before generic part deletion**

Before `PresentationModel.deleteSlide()` deletes the target part, collect every internal exact slide relationship from every remaining slide whose `resolvedTarget` is the target URI. Parse each source slide, call `removeDrawingHyperlinkReferences()` with those relationship IDs, and set only changed source parts. Continue through existing presentation entry/section/relationship/part deletion and dependency GC in the same transaction.

Do not remove action-only links, wrong-namespace lookalikes, unrelated target links, or opaque content without a DrawingML click/hover element. Generic `deletePart()` remains responsible for removing incoming relationships after XML cleanup.

- [ ] **Step 8: Run focused, lifecycle, and full model gates**

```sh
pnpm vitest run packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/model.test.ts -t 'shape hyperlink' \
  packages/sdk/src/index.test.ts -t 'shape hyperlink' --reporter=dot
pnpm vitest run packages/model/src/model.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass; existing slide duplicate/delete/move, notes, media/image dependency, and preset-shape tests remain green.

- [ ] **Step 9: Review, commit, push, and verify**

Review semantic no-op comparison, one-based/current-order getter semantics, unique/shared relationship behavior, update ordering, reference-aware GC, external↔internal mode transitions, extra-state preservation, malformed rejection, self-link duplication, target deletion coverage, transaction rollback, six formats, and unchanged unrelated dependency lifecycles. Then:

```sh
git add -- packages/model/src/shape-hyperlink.internal.ts \
  packages/model/src/shape-hyperlink.internal.test.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/presentation.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit shape hyperlinks"
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
- Consumes: Task 3 final public behavior, locked `pptxgenjs@4.0.1`, the actual `@jiayunxie/pptx` tarball, Node ESM export, browser conditional export, generated declarations, and current smoke harness.
- Produces: public-output conformance/strict-divergence evidence and tarball-level `shapeHyperlinks: true` evidence.

- [ ] **Step 1: Add public-only PptxGenJS output cases**

Generate through public `addShape()` and `write()` only:

```ts
const cases = [
  { name: 'URL', hyperlink: { url: 'https://example.com?a=1&b=2' } },
  { name: 'URL tooltip', hyperlink: { url: 'mailto:test@example.com', tooltip: 'Mail & help' } },
  { name: 'Slide', hyperlink: { slide: 2 } },
  { name: 'Slide tooltip', hyperlink: { slide: 3, tooltip: '' } },
  { name: 'Empty', hyperlink: {} },
  { name: 'Both', hyperlink: { url: 'https://example.com', slide: 2 } },
  { name: 'Zero', hyperlink: { slide: 0 } },
  { name: 'Negative', hyperlink: { slide: -1 } },
  { name: 'Fraction', hyperlink: { slide: 1.5 } },
  { name: 'Out of range', hyperlink: { slide: 99 } },
  { name: 'Numeric URL', hyperlink: { url: 42 } },
  { name: 'String value', hyperlink: 'https://example.com' },
];
```

Spy on console output without changing runtime behavior. Import actual bytes, assert valid URL/slide snapshots and exact XML/relationship state, assert omitted tooltip becomes direct empty in PptxGenJS, and document ignored/coerced malformed states. Validate dangling `Both`/negative/fraction/out-of-range outputs separately instead of importing them as native supported state.

- [ ] **Step 2: Compare supported native semantics and strict divergences**

Create native equivalents for valid URL, URL tooltip, slide, slide tooltip, and self-link. Compare normalized snapshot, resolved relationship target/type/mode, action, tooltip, duplicate/reopen behavior, and presentation order semantics. Do not require matching relationship IDs or `tooltip=""` for native omitted tooltip.

Require native missing/both targets, empty/numeric URL, invalid slide, non-object, `_rId`, aliases, symbols/accessors, and out-of-range target to throw before mutation. Keep PptxGenJS console-ignore, coercion, duplicate-click, and dangling-output behavior explicitly asserted as divergences.

- [ ] **Step 3: Extend installed-package Node smoke**

In generated `smoke.mjs`, create URL/internal/self links, mutate caller objects, read frozen detached snapshots, edit tooltip and targets, exercise shared relationship clone-on-write, duplicate self-link, move target, clear, delete target, write/reopen, and require expected relationship counts plus zero dangling diagnostics. Add `shapeHyperlinks` to the final JSON envelope without removing existing fields.

- [ ] **Step 4: Extend browser and declaration smoke**

In browser smoke, import `PptxDocument`/`ShapeModel` from the package browser condition, create/read/edit/clear/reopen URL and internal links, and verify no `node:` dependency. In generated typecheck source, import `Hyperlink`, use both union branches through `AddShapeOptions.hyperlink` and `ShapeModel.hyperlink`, and require missing/both/invalid target plus unknown-key cases with `@ts-expect-error`.

- [ ] **Step 5: Run focused tests and verify expected failure**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'shape hyperlink public output' --reporter=dot
```

Expected: FAIL because the conformance test and packed smoke field are not implemented.

- [ ] **Step 6: Implement conformance and packed smoke coverage**

Add only tests/smoke behavior; production adapter code remains unchanged because `importPptxGenJS()` already imports public `write()` bytes generically. Keep `pptxgenjs` absent from non-adapter manifests and avoid internal workspace imports in installed-package scripts.

- [ ] **Step 7: Build, pack, and run all compatibility gates**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
pnpm typecheck
pnpm build
mkdir -p /tmp/pptx-shape-hyperlink-pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-shape-hyperlink-pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-hyperlink-pack/jiayunxie-pptx-0.1.0.tgz
node --check scripts/smoke-npm-package.mjs
git diff --check
```

Expected JSON includes `"shapeHyperlinks":true`, `"shapeArrows":true`, `"shapeLines":true`, `"shapeFills":true`, `"presetShapes":true`, `"types":true`, and CLI version `0.1.0`.

- [ ] **Step 8: Review, commit, push, and verify**

Review public-only generation, valid URL/internal parity, tooltip direct-state difference, actual malformed 4.0.1 evidence, strict zero-mutation rejection, imported editing/reopen, installed-tarball isolation, Node/browser parity, union declaration checks, lifecycle coverage, existing smoke preservation, and dependency boundaries. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify shape hyperlink compatibility"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Document the supported boundary and remaining hyperlink gaps

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: Tasks 1–4 final public API, lifecycle semantics, conformance evidence, and strict divergence results.
- Produces: user-facing create/edit examples, changelog entry, API contract, and an honest compatibility-matrix update.

- [ ] **Step 1: Add concise public examples**

Document exact native syntax:

```ts
const target = document.addSlide();
const shape = document.slides[0]!.addShape('roundRect', {
  hyperlink: {
    url: 'https://example.com/docs',
    tooltip: 'Open documentation',
  },
});

shape.hyperlink = { slide: 2, tooltip: 'Go to details' };
shape.hyperlink = { url: 'mailto:team@example.com', tooltip: '' };
shape.hyperlink = undefined;
```

Explain exactly-one-target input, one-based slide numbers, target identity across reorder, tooltip absent versus empty, detached snapshots, whole replacement, same-value no-op, shared relationship clone-on-write/GC, duplicate self-link, and target deletion cleanup.

- [ ] **Step 2: Update compatibility, API, and changelog**

Change the `slide.addShape()` row to include URL/internal-slide hyperlink create/read/edit/clear/duplicate/reopen. Record PptxGenJS omitted-tooltip materialization, native direct absence, expected external portability warning, console-ignore/coercion, and invalid both/slide outputs without claiming byte identity.

Remove shape-level hyperlink from the pending shape list while keeping hover, text-run/table/image/chart/media hyperlink creation, action navigation, shadow, adjustments, custom geometry, advanced line state, shape text options, and percentage positions explicit. Change the next shape priority from hyperlink to shadow.

- [ ] **Step 3: Validate examples and documentation consistency**

```sh
rg -n "Hyperlink|hyperlink|tooltip|one-based|slide target|shadow" \
  README.md packages/pptx/README.md CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
git diff --check
```

Expected: examples compile against exported names, URL/slide exclusivity is unambiguous, tooltip absence/empty and external validator warning are accurate, and all remaining gaps remain visible.

- [ ] **Step 4: Review, commit, push, and verify**

Review API spelling, exactly-one-target semantics, one-based/current-order behavior, direct tooltip wording, whole replacement, relationship lifecycle, target deletion, PptxGenJS divergences, expected warning language, and remaining-gap accuracy across root/package/API/baseline docs. Then:

```sh
git add -- README.md packages/pptx/README.md CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document shape hyperlinks"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Run release, real-PPTX, structural, visual, and import QA

**Files:**
- No repository file changes expected; use `/tmp/pptx-shape-hyperlink-qa-20260731`.

**Interfaces:**
- Consumes: Tasks 1–5 complete repository state.
- Produces: final evidence for package integrity, relationship/mutation isolation, Office-compatible hyperlink persistence, visual non-regression, import compatibility, and remote parity.

- [ ] **Step 1: Run complete repository gates**

```sh
pnpm check
pnpm test:performance
pnpm build
```

Expected: strict typecheck/build, all tests, and independent performance budget pass.

- [ ] **Step 2: Pack and smoke the final tarball**

```sh
mkdir -p /tmp/pptx-shape-hyperlink-qa-20260731/pack
pnpm --filter @jiayunxie/pptx pack \
  --pack-destination /tmp/pptx-shape-hyperlink-qa-20260731/pack
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-shape-hyperlink-qa-20260731/pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: identity/version, Node/browser/types, preset shapes, fills, lines, arrows, hyperlinks, dependency boundaries, and CLI smoke all pass.

- [ ] **Step 3: Generate representative lifecycle decks**

Create native decks for URL with metacharacters, URL tooltip absent/empty/custom, mailto/custom URI, internal first/middle/last/self target, URL↔slide editing, shared relationship clone-on-write, clear/GC, duplicate external/other/self, move/insert/delete ordinal changes, target deletion cleanup, malformed zero-mutation source, rollback, reopen, and second write. Create PptxGenJS valid comparison and isolated malformed-divergence fixtures with visibly labeled shapes.

- [ ] **Step 4: Validate packages and exact mutation boundaries**

For each legal deck run:

```sh
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json slides list deck.pptx
pptx-inspect --json part read deck.pptx /ppt/slides/slide1.xml
pptx-inspect --json part read deck.pptx /ppt/slides/_rels/slide1.xml.rels
```

Use `package diff` to prove creation/edit changes only the intended slide XML and its relationship part, tooltip-only edit changes only slide XML, same-value changes no parts, source inputs remain unchanged, and reopen→second-write changes no parts. Verify exact action/type/mode/target, XML escaping, current ordinal snapshots, shared references, GC, self-duplicate retarget, target deletion cleanup, shape IDs/order/names, and neighbor bytes.

Require zero errors for every legal fixture; internal-only fixtures require zero warnings, while each external URL fixture may contain only the expected `OPC_EXTERNAL_RELATIONSHIP` warning. Malformed PptxGenJS fixtures must produce the exact expected dangling diagnostics and remain outside supported-native claims.

- [ ] **Step 5: Round-trip through LibreOffice and inspect visual output**

Open/export every legal deck with an isolated LibreOffice profile. Reinspect the resulting PPTX relationship and click XML to prove URL/internal targets and tooltips survive. Export to PDF, run `pdfinfo`, render pages with `pdftoppm`, then run `render_slides.py` and `slides_test.py` for every PPTX.

Inspect every rendered page at full size for unchanged shape geometry/fill/line/arrows/text, all labels, clipping, overlap, and overflow. PDF is a visual non-regression check; hyperlink semantics are established by OOXML/relationship inspection before and after LibreOffice round-trip.

- [ ] **Step 6: Verify artifact-tool import behavior**

Import native representative/lifecycle decks with `PresentationFile.importPptx()`, require expected slide/shape counts, and inspect named hyperlink shapes. Import valid PptxGenJS comparison files and confirm supported files load without source mutation.

- [ ] **Step 7: Verify repository and remote final state**

```sh
git diff --check
git status --short
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: only the pre-existing untracked `.pnpm-store/` remains ignored, and divergence is `0 0`. Create no empty commit when Task 6 produces no repository change.
