# Master、Layout 与 Placeholder Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict named-layout definition, selection, semantic master/layout editing, and placeholder materialization/population with PptxGenJS 4.0.1 public-output parity.

**Architecture:** PptxGenJS `defineSlideMaster()` is modeled as a named `slideLayout` under a real parent master. Low-level package graph operations stay in `@pptx/codecs`; SDK stable wrappers compose owner-neutral `SlideModel` content for master/layout parts, while placeholder identity and background codecs become owner-aware. Definition sources and chart workbooks are prepared asynchronously before one synchronous OPC transaction commits the layout.

**Tech Stack:** TypeScript 5.8, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, JSZip, PptxGenJS 4.0.1, tsup, real Chrome, `pptx-inspect`, LibreOffice/Poppler, PowerPoint automation.

## Global Constraints

- Execute inline in the current task; do not delegate implementation or review.
- PptxGenJS `masterName` selects a unique direct layout `p:cSld@name`; it does not create or select one real master per name.
- Unknown and duplicate names fail before mutation; do not copy PptxGenJS silent default/first-match behavior.
- Ordinary layout objects remain in layout parts. Only empty placeholder owners and direct slide-number visibility state materialize into a new slide.
- Placeholder creation types are exactly `title | body | pic | chart | tbl | media`; index `4294967295` is never allocated.
- Placeholder name and `type + index` are each unique within one layout; selectors must resolve exactly one layout and one slide owner.
- Population domains are strict: title/body → text or shape, pic → image, chart → chart, tbl → table, media → audio/video.
- Master/layout direct background owns only `p:cSld/p:bg`; unsupported `p:bgRef`, pattern/group fill, or unsafe relationships remain byte-preserved until explicit safe replacement.
- Layout margin is transient frozen EMU state. It is not serialized; reopen returns `undefined`.
- Async image/background resolution and chart workbook construction finish before package mutation.
- Define/replace/delete, placeholder population, dependency cleanup, transient state, model caches, and mutation journal are transactionally isolated.
- Preserve low-level `document.masterLayoutTheme` raw codec access; SDK wrappers do not create a dependency from codecs back to model.
- Existing slide APIs keep their signatures unless this plan explicitly adds `placeholder?: PlaceholderSelector`.
- Apply the same package behavior to `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`.
- Production packages never import PptxGenJS. Locked 4.0.1 usage stays in adapter/tests and generated evidence.
- Use fresh `/tmp/pptx-master-layout-placeholder-*` directories for generated decks, package copies, tarballs, renders, validator output, and client round trips.
- Every task ends with focused tests, relevant type/build gates, diff review, one isolated commit, SSH push, fetch, and divergence `0 0`.
- Never stage `.pnpm-store/`, generated PPTX/PDF/PNG, tarballs, manifests from `/tmp`, or client round-trip artifacts.

---

## File Map

- Create `packages/model/src/presentation-layout.internal.ts`: strict direct layout-name discovery and named/default resolution.
- Create `packages/model/src/placeholder.ts`: public placeholder types, selector, and identity.
- Create `packages/model/src/placeholder.internal.ts`: strict identity read/normalize/render, empty materialization, selector/domain resolution, and owner replacement helpers.
- Create `packages/model/src/chart-create.internal.ts`: prepared chart creation plan and synchronous commit helper shared by slides and layouts.
- Modify `packages/model/src/presentation-sections.internal.ts`: accept strict `masterName` beside `sectionTitle`.
- Modify `packages/model/src/presentation.ts`: named layout selection, placeholder materialization, and lifecycle hooks.
- Modify `packages/model/src/slide.ts`: owner-neutral content delegation and placeholder-aware creators.
- Modify `packages/model/src/shapes.ts`: live read-only `ShapeModel.placeholder` identity.
- Modify `packages/model/src/slide-background.internal.ts`: `slide | layout | master` direct background roots.
- Modify `packages/model/src/slide-background.ts`, `image.ts`, `chart.ts`, `preset-shape.ts`, and table creation types: placeholder-aware public options.
- Modify `packages/model/src/index.ts`: export public placeholder types.
- Modify `packages/codecs/src/master.ts`: strict attached/raw ownership helpers and safe parent/header-footer operations.
- Modify `packages/codecs/src/media.ts`: placeholder selector in media options and canonical media picture ownership.
- Create `packages/sdk/src/master-layout.ts`: stable SDK wrappers, definition types/normalizers, transient margin map, and definition plans.
- Modify `packages/sdk/src/index.ts`: wrapper collections and define/replace/delete orchestration.
- Modify `packages/sdk/src/raster-image-source.ts`: allow placeholder selectors through high-level image options.
- Modify `packages/codecs/src/codecs.test.ts`, `packages/model/src/model.test.ts`, `packages/sdk/src/index.test.ts`: focused TDD coverage.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: locked public-output conformance and intentional differences.
- Modify `packages/validator/src/index.test.ts`: validator coverage for native/imported named layouts and populated placeholders.
- Modify `packages/pptx/src/index.test.ts`: root package exports and runtime flags.
- Modify `scripts/build-npm-package-types.mjs`: installed TypeScript consumer.
- Modify `scripts/smoke-npm-package.mjs`: actual tarball Node/CLI/gallery evidence and `masterLayouts: true`.
- Modify `scripts/playwright-browser-smoke.js`: real-Chrome definition/selection/population/reopen state.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: final contract, evidence, remaining scope, and next roadmap item.

---

### Task 1: Strict named layout selection

**Files:**

- Create: `packages/model/src/presentation-layout.internal.ts`
- Modify: `packages/model/src/presentation-sections.internal.ts`
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: presentation→master→layout package graph, direct layout XML names, existing section normalization and blank-slide creation.
- Produces:

```ts
export interface AddSlideOptions {
  readonly masterName?: string;
  readonly sectionTitle?: string;
}

export interface NormalizedAddPresentationSlideOptions {
  readonly masterName?: string;
  readonly sectionTitle?: string;
}

export function resolveSlideLayoutPartUri(
  pkg: OpcPackage,
  presentationPartUri: string,
  masterName: string | undefined,
  inheritedLayoutPartUri: string | undefined,
): string;
```

- [ ] **Step 1: Add failing option and named-selection tests**

In `packages/sdk/src/index.test.ts`, create one second layout by copying default XML and changing only direct `p:cSld@name`:

```ts
const document = PptxDocument.create();
const rawMaster = document.masterLayoutTheme.masters[0]!;
const defaultLayout = document.masterLayoutTheme.layouts[0]!;
const defaultXml = new TextDecoder().decode(
  document.opcPackage.requirePart(defaultLayout.partUri).bytes,
);
const brand = document.masterLayoutTheme.createLayout(
  rawMaster.partUri,
  defaultXml.replace('name="DEFAULT"', 'name="BRAND"'),
);

document.addSection({ title: 'Named' });
const slide = document.addSlide({ masterName: 'BRAND', sectionTitle: 'Named' });
expect(slide.relationships.find(({ type }) => type.endsWith('/slideLayout'))?.resolvedTarget)
  .toBe(brand.partUri);
expect(document.sections?.[0]?.slideIds).toContain(slide.slideId);
```

Add strict cases for whitespace/invalid XML/unknown name, duplicate direct names, wrong namespace name lookalikes, external/dangling/wrong-content-type relationships, inherited default selection, and options with accessor/symbol/extra keys. Assert `packageSnapshot()` and mutation journal stay unchanged on every failure.

- [ ] **Step 2: Run the focused test and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "named slide layout"
```

Expected: FAIL because `masterName` is rejected as an unknown option and `addSlide()` always chooses the inherited/default layout.

- [ ] **Step 3: Normalize `masterName` as a strict title**

Update `normalizeAddPresentationSlideOptions()`:

```ts
const data = readDataObject(
  value,
  'Add slide options',
  ['masterName', 'sectionTitle'],
);
const masterName = data.masterName === undefined
  ? undefined
  : normalizePresentationSectionTitle(data.masterName);
const sectionTitle = data.sectionTitle === undefined
  ? undefined
  : normalizePresentationSectionTitle(data.sectionTitle);
return {
  ...(masterName !== undefined ? { masterName } : {}),
  ...(sectionTitle !== undefined ? { sectionTitle } : {}),
};
```

Rename the shared string error context inside the normalizer call so invalid names report `Slide master name` rather than `Presentation section title`.

- [ ] **Step 4: Implement strict graph/name resolution**

In `presentation-layout.internal.ts`, enumerate only internal content-type-correct relationships reachable from the presentation. Parse each layout and accept only one namespace-correct direct `p:sldLayout/p:cSld` with one direct unqualified `name` attribute. Return the unique exact match, otherwise throw:

```ts
if (matches.length === 0) {
  throw new RangeError(`Slide master ${masterName} was not found`);
}
if (matches.length !== 1) {
  throw new ModelParseError(`Slide master ${masterName} is ambiguous`, presentationPartUri);
}
return matches[0]!;
```

When `masterName` is absent, prefer the source slide's unique safe internal layout, then the first safe attached layout; throw `ModelParseError` if neither exists.

- [ ] **Step 5: Select before slide allocation and preserve section behavior**

At the start of the existing `addSlide()` transaction, compute:

```ts
const inheritedLayoutPartUri = this.slides[0]?.relationships
  .find(({ type }) => type === SLIDE_LAYOUT_RELATIONSHIP)?.resolvedTarget;
const layoutPartUri = resolveSlideLayoutPartUri(
  this.opcPackage,
  this.presentationPartUri,
  normalized.masterName,
  inheritedLayoutPartUri,
);
```

Use only that validated URI when creating the slide relationship. Keep the existing section transaction unchanged.

- [ ] **Step 6: Cover all formats and run gates**

Add a loop over `pptx/pptm/potx/potm/ppsx/ppsm`: create, add a raw named layout, select it, write/reopen, and assert the relationship target/name. Run:

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "named slide layout"
pnpm typecheck
pnpm --filter @pptx/model build
pnpm --filter @pptx/sdk build
git diff --check
```

Expected: focused tests, typecheck, and builds pass.

- [ ] **Step 7: Review, commit, push, and verify**

Review that no default/section behavior changed without `masterName`, then:

```bash
git add packages/model/src/presentation-layout.internal.ts packages/model/src/presentation-sections.internal.ts packages/model/src/presentation.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: select named slide layouts"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 2: Stable semantic master/layout wrappers

**Files:**

- Create: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**

- Consumes: raw codec models and a package-backed internal `SlideModel` content owner.
- Produces:

```ts
export class SlideLayoutModel {
  readonly partUri: string;
  get name(): string;
  get masterPartUri(): string | undefined;
  get shapes(): readonly SemanticShape[];
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}

export class SlideMasterModel {
  readonly partUri: string;
  get layouts(): readonly SlideLayoutModel[];
  get shapes(): readonly SemanticShape[];
  get theme(): ThemeModel | undefined;
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}
```

- [ ] **Step 1: Add failing wrapper identity/content tests**

```ts
const document = PptxDocument.create();
const layout = document.layouts[0]!;
const master = document.masters[0]!;
expect(document.layouts[0]).toBe(layout);
expect(document.masters[0]).toBe(master);
expect(master.layouts[0]).toBe(layout);

const text = layout.addText('Inherited layout text');
const shape = master.addShape('rect', {
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
});
expect(layout.shapes.find(({ id }) => id === text.id)).toBe(text);
expect(master.shapes.find(({ id }) => id === shape.id)).toBe(shape);
expect(document.masterLayoutTheme.layouts[0]?.partUri).toBe(layout.partUri);
```

Also assert stable shape identity, transform/text edit, raw copy/relink visibility, wrong-document wrapper rejection, deleted-part stale handles, and current slide-number getter/setter behavior.

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "semantic master layout models"
```

Expected: FAIL because `document.layouts` and `masters` still expose raw codec models without content creators.

- [ ] **Step 3: Implement one owner-neutral wrapper base**

In `master-layout.ts`, create a package-private base holding raw model and internal content:

```ts
abstract class CommonSlideOwnerModel {
  protected constructor(
    protected readonly document: PptxDocument,
    readonly partUri: string,
    protected readonly content: SlideModel,
  ) {}

  get shapes(): readonly SemanticShape[] {
    return this.content.shapes;
  }

  addText(value: string, options: AddTextOptions = {}): ShapeModel {
    return this.content.addText(value, options);
  }

  addRichText(
    value: readonly RichTextParagraph[],
    options: AddTextOptions = {},
  ): ShapeModel {
    return this.content.addRichText(value, options);
  }

  addShape(type: PresetShapeType, options: AddShapeOptions = {}): ShapeModel {
    return this.content.addShape(type, options);
  }

  addImage(bytes: Uint8Array, options: AddImageOptions): ImageModel {
    return this.content.addImage(bytes, options);
  }
}
```

Add explicit chart overloads matching `SlideModel.addChart()` and delegate to `content`. Construct content with the real document/presentation and owner part URI, but never expose slide-only properties.

- [ ] **Step 4: Cache wrappers by part URI**

Add `#layoutModels`, `#masterModels`, `modelForLayout(raw)`, and `modelForMaster(raw)` to `PptxDocument`. Change `layouts`/`masters` getters to map raw collections through these caches. Wrapper `master.layouts` calls back into the same cache, so all paths return identical objects.

Keep `masterLayoutTheme`, raw `LayoutModel`, raw `MasterModel`, and raw codec tests unchanged.

- [ ] **Step 5: Run wrapper and regression gates**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "master|layout"
pnpm typecheck
pnpm --filter @pptx/sdk build
pnpm --filter @jiayunxie/pptx build
git diff --check
```

- [ ] **Step 6: Review, commit, and push**

```bash
git add packages/sdk/src/master-layout.ts packages/sdk/src/index.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts
git diff --cached --check
git commit -m "feat: expose semantic master layout models"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 3: Direct layout/master backgrounds

**Files:**

- Modify: `packages/model/src/slide-background.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/model/src/slide-background.internal.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

```ts
export type BackgroundOwnerKind = 'slide' | 'layout' | 'master';

export function readSlideBackground(
  pkg: OpcPackage,
  partUri: string,
  ownerKind?: BackgroundOwnerKind,
): SlideBackground | undefined;

export function replaceSlideBackground(
  pkg: OpcPackage,
  partUri: string,
  value: unknown,
  ownerKind?: BackgroundOwnerKind,
): void;
```

- [ ] **Step 1: Add failing owner-root tests**

For each `layout` and `master`, test inherited clear, legal no-fill, sRGB/theme solid+transparency, linear/path gradient, PNG image, unsupported `p:bgRef`, wrong namespace, multiple `cSld`, equal no-op, opaque replacement, shared target COW, clear/delete GC, and rollback.

Public SDK assertion:

```ts
const document = PptxDocument.create();
const layout = document.layouts[0]!;
const master = document.masters[0]!;
layout.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
};
master.background = { kind: 'none' };
expect(layout.background).toEqual({
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
});
expect(master.background).toEqual({ kind: 'none' });
```

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/slide-background.internal.test.ts packages/sdk/src/index.test.ts -t "layout master background"
```

Expected: FAIL because the background reader requires a `p:sld` root.

- [ ] **Step 3: Parameterize only the root contract**

Add a constant map:

```ts
const OWNER_ROOTS = {
  slide: 'sld',
  layout: 'sldLayout',
  master: 'sldMaster',
} as const;
```

Use `OWNER_ROOTS[ownerKind]` in strict root resolution while retaining the same direct `cSld/p:bg`, fill, relationship, resource replacement, COW, and GC code. `SlideModel` continues calling with the default `slide`; wrappers call with `layout` or `master`.

- [ ] **Step 4: Lock PptxGenJS layout-background import**

Generate public `defineSlideMaster()` solid, transparency, PNG data, inherited, and deprecated `bkgd` cases. Assert the matching imported `SlideLayoutModel.background`; record empty/invalid `bgPr` as unsupported without mutation.

- [ ] **Step 5: Run gates and commit**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/slide-background.internal.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "layout master background"
pnpm typecheck
pnpm --filter @pptx/model build
pnpm --filter @pptx/sdk build
git diff --check
git add packages/model/src/slide-background.internal.ts packages/model/src/slide.ts packages/sdk/src/master-layout.ts packages/model/src/slide-background.internal.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit master layout backgrounds"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 4: Placeholder identity and empty materialization

**Files:**

- Create: `packages/model/src/placeholder.ts`
- Create: `packages/model/src/placeholder.internal.ts`
- Modify: `packages/model/src/index.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
export const PLACEHOLDER_TYPES = [
  'title', 'body', 'pic', 'chart', 'tbl', 'media',
] as const;
export type PlaceholderType = typeof PLACEHOLDER_TYPES[number];
export interface PlaceholderIdentity {
  readonly type: PlaceholderType;
  readonly index: number;
}
export type PlaceholderSelector = string | PlaceholderIdentity;
export interface AddPlaceholderOptions extends Omit<AddTextOptions, 'name'> {
  readonly name: string;
  readonly type: PlaceholderType;
  readonly index?: number;
}
```

- [ ] **Step 1: Add failing normalization/identity tests**

```ts
const layout = document.layouts[0]!;
const prompt = layout.addPlaceholder('Quarterly title', {
  name: 'title_box',
  type: 'title',
  index: 103,
  x: 100,
  y: 200,
  width: 300,
  height: 400,
});
expect(prompt.placeholder).toEqual({ type: 'title', index: 103 });
expect(Object.isFrozen(prompt.placeholder)).toBe(true);
expect(layout.placeholders).toEqual([prompt]);

const slide = document.addSlide({ masterName: 'DEFAULT' });
const empty = slide.shapes.find(({ name }) => name === 'title_box')!;
expect(empty.placeholder).toEqual({ type: 'title', index: 103 });
expect(empty.transform).toEqual(prompt.transform);
expect((empty as ShapeModel).richText.flatMap(({ runs }) => runs)).toEqual([]);
```

Test all six types, omitted index allocation by `100 + ordinal`, explicit zero, max accepted value, reserved max rejection, duplicate name, duplicate identity, prompt rich text, alternate prefixes, wrong namespace, inherited unknown OOXML types, exact no-op, all formats, reopen, and rollback.

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "placeholder identity|empty layout placeholder"
```

- [ ] **Step 3: Implement strict identity and renderer**

`normalizePlaceholderIdentity()` accepts only own-data `type/index`, validates the closed token set and `0..4294967294`, returns a frozen copy. `readShapePlaceholder()` requires one direct namespace-correct `p:nvPr/p:ph` under the shape's own nonvisual container and rejects duplicate/qualified attributes.

Render canonical layout placeholder with direct `cNvPr@name`, `p:ph@type/idx`, transform, prompt text, and supported body/text state. Add:

```ts
get placeholder(): Readonly<PlaceholderIdentity> | undefined {
  return this.slide.getShapePlaceholder(this.id);
}
```

- [ ] **Step 4: Materialize empty owners inside `addSlide()`**

After adding the validated slide→layout relationship and before attaching the slide to presentation, call:

```ts
materializeLayoutPlaceholders(
  this.opcPackage,
  layoutPartUri,
  slideUri,
);
```

The helper validates the complete supported layout placeholder set first, allocates unique slide shape IDs, copies name/type/index/transform/body properties, removes prompt runs, and appends canonical empty paragraphs. Ordinary layout objects never copy.

- [ ] **Step 5: Expose wrapper placeholder creation/collection**

`SlideLayoutModel.addPlaceholder()` and `SlideMasterModel.addPlaceholder()` delegate to owner-neutral insertion. `placeholders` filters `shapes` by `shape.placeholder !== undefined` and returns the same stable shape instances.

- [ ] **Step 6: Run gates, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "placeholder"
pnpm typecheck
pnpm --filter @pptx/model build
pnpm --filter @pptx/sdk build
git diff --check
git add packages/model/src/placeholder.ts packages/model/src/placeholder.internal.ts packages/model/src/index.ts packages/model/src/shapes.ts packages/model/src/slide.ts packages/model/src/presentation.ts packages/sdk/src/master-layout.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: materialize layout placeholders"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 5: Populate text and shape placeholders

**Files:**

- Modify: `packages/model/src/placeholder.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly placeholder?: PlaceholderSelector;
}

export interface AddShapeOptions extends Partial<Transform> {
  readonly placeholder?: PlaceholderSelector;
}
```

- [ ] **Step 1: Add failing text/rich/shape population tests**

```ts
const slide = document.addSlide({ masterName: 'BRAND' });
const empty = slide.shapes.find(({ name }) => name === 'title_box')!;
const title = slide.addText('Filled title', { placeholder: 'title_box' });
expect(title.id).toBe(empty.id);
expect(title.placeholder).toEqual({ type: 'title', index: 103 });
expect(title.transform).toEqual(layout.placeholders[0]!.transform);
expect(title.richText[0]?.runs[0]?.text).toBe('Filled title');
expect(() => empty.transform).toThrow();
```

Add rich text, identity selector, body placeholder, line/rect shape population, explicit transform ignored in favor of layout geometry, unknown/ambiguous selector, missing slide owner, second fill, wrong domain (`pic` with text), malformed layout relationship, rollback, sibling isolation, duplicate/move/delete, and reopen.

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "populate text shape placeholder"
```

- [ ] **Step 3: Resolve layout and slide owners before mutation**

Implement:

```ts
export interface ResolvedPlaceholderOwner {
  readonly identity: Readonly<PlaceholderIdentity>;
  readonly name: string;
  readonly shapeId: number;
  readonly transform: Transform;
  readonly slideElement: XmlElement;
  readonly layoutElement: XmlElement;
}

export function resolvePlaceholderOwner(
  pkg: OpcPackage,
  slidePartUri: string,
  selector: PlaceholderSelector,
  domain: 'text-shape' | 'image' | 'chart' | 'table' | 'media',
): ResolvedPlaceholderOwner;
```

String selector matches direct layout `cNvPr@name`; identity matches normalized `type + index`. Require the slide owner to be empty, direct, and matching. Domain mismatch throws before any package write.

- [ ] **Step 4: Replace at the same tree slot and ID**

Add an optional resolved owner to text and preset-shape internal insertion. Render with owner transform/name/identity, use the existing slide shape ID, replace `slideElement.start..end`, and preserve tree order. Ignore caller x/y/width/height/rotation/flip only when placeholder is present; all content/style options continue through existing strict normalization.

- [ ] **Step 5: Run gates, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "placeholder"
pnpm typecheck
pnpm --filter @pptx/model build
git diff --check
git add packages/model/src/placeholder.internal.ts packages/model/src/slide.ts packages/model/src/preset-shape.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: populate text shape placeholders"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 6: Populate image and chart placeholders

**Files:**

- Create: `packages/model/src/chart-create.internal.ts`
- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/image.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
export interface PreparedChartCreation {
  readonly definition: Readonly<ChartDefinition>;
  readonly options: Readonly<AddChartOptions>;
  readonly workbookBytes: Uint8Array;
  readonly formulas: readonly string[];
}

export async function prepareChartCreation(
  groups: readonly ChartGroupInput[],
  options?: AddChartOptions,
): Promise<PreparedChartCreation>;

export function commitPreparedChart(
  slide: SlideModel,
  prepared: PreparedChartCreation,
): ChartModel;
```

- [ ] **Step 1: Add failing picture/chart placeholder tests**

Create `pic` and `chart` placeholders, then populate them with raster bytes, SVG+fallback, high-level data URI/Blob source, single chart, and combo chart. Assert same shape ID/tree slot, inherited geometry, correct `p:ph`, media/chart/workbook relationships, dedup/COW, source rectangle, chart definition, reopen, duplicate isolation, delete GC, and rollback.

```ts
const picture = await document.addImage(slideIndex, PNG_BYTES, {
  contentType: 'image/png',
  placeholder: 'hero_image',
});
expect(picture.placeholder).toEqual({ type: 'pic', index: 104 });

const chart = await slide.addChart([
  { type: 'bar', series: [{ name: 'Revenue', categories: ['Q1'], values: [10] }] },
], { placeholder: { type: 'chart', index: 105 } });
expect(chart.placeholder).toEqual({ type: 'chart', index: 105 });
```

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "populate image chart placeholder"
```

- [ ] **Step 3: Split chart preparation from commit**

Move only normalization, workbook build, and formula planning into `prepareChartCreation()`. Move the existing synchronous chart/workbook parts, relationships, frame insertion, and model resolution into `commitPreparedChart()`. Public `SlideModel.addChart()` becomes:

```ts
const prepared = await prepareChartCreation(groups, normalizedOptions);
return this.presentation.opcPackage.transaction(() =>
  commitPreparedChart(this, prepared));
```

Preserve existing chart bytes/tests when no placeholder is supplied.

- [ ] **Step 4: Make image/chart frames placeholder-aware**

Add `placeholder?: PlaceholderSelector` to raster/SVG image and chart options and their strict key lists. Resolve before allocating parts. Commit media/chart resources and replace the empty owner with matching `p:ph`; on any failure the outer transaction restores resources and XML.

- [ ] **Step 5: Run full chart/image regressions and commit**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/model/src/chart-render.internal.test.ts packages/model/src/chart-edit.internal.test.ts packages/sdk/src/index.test.ts -t "image|chart|placeholder"
pnpm typecheck
pnpm --filter @pptx/model build
pnpm --filter @pptx/sdk build
git diff --check
git add packages/model/src/chart-create.internal.ts packages/model/src/chart.ts packages/model/src/image.ts packages/model/src/slide.ts packages/sdk/src/raster-image-source.ts packages/sdk/src/index.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: populate image chart placeholders"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 7: Populate table and media placeholders

**Files:**

- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
export interface AddTableOptions {
  readonly placeholder?: PlaceholderSelector;
}

export interface AddMediaOptions {
  readonly placeholder?: PlaceholderSelector;
}
```

- [ ] **Step 1: Add failing table/audio/video population tests**

Use `tbl` and `media` placeholders. Cover basic/rich table state already supported by native creators, embedded/external audio, embedded video, poster replacement, timing settings, geometry inheritance, matching `p:ph`, same ID/tree order, duplicate COW, object/slide delete GC, reopen, and rollback.

```ts
const table = slide.addTable([['A', 'B']], { placeholder: 'data_table' });
expect(table.placeholder).toEqual({ type: 'tbl', index: 106 });

const video = await slide.addVideo(MP4_BYTES, {
  contentType: 'video/mp4',
  placeholder: { type: 'media', index: 107 },
});
expect(video.placeholder).toEqual({ type: 'media', index: 107 });
```

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "populate table media placeholder"
```

- [ ] **Step 3: Add placeholder-aware frame/picture commit**

Normalize the selector with table/media options before source I/O or part allocation. Table replaces empty owner with a `p:graphicFrame` whose `p:nvGraphicFramePr/p:nvPr` contains matching `p:ph`. Media replaces with `p:pic` whose own `p:nvPr` contains matching `p:ph`; poster/media relationships and native timing target the preserved shape ID.

- [ ] **Step 4: Reject mismatches and preserve existing bytes**

Lock `tbl`→table and `media`→audio/video. Reject title/body/pic/chart selectors, already-filled owners, ambiguous layout identity, missing owner, and invalid source without mutation. Confirm no-placeholder table/media bytes remain unchanged against existing expected snapshots.

- [ ] **Step 5: Run media/table regressions, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/codecs/src/codecs.test.ts -t "table|media|placeholder"
pnpm typecheck
pnpm --filter @pptx/codecs build
pnpm --filter @pptx/model build
git diff --check
git add packages/model/src/slide.ts packages/model/src/table-create.internal.ts packages/codecs/src/media.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: populate table media placeholders"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 8: Declarative synchronous slide-master definition

**Files:**

- Modify: `packages/codecs/src/master.ts`
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**

Implement the design types `SlideMasterMargin`, `SlideMasterMarginInput`, sync `SlideMasterObject` variants (`rect`, `line`, `text`, `placeholder`), `DefineSlideMasterOptions`, and:

```ts
defineSlideMaster(options: DefineSlideMasterOptions): Promise<SlideLayoutModel>;
```

- [ ] **Step 1: Add failing strict-definition tests**

```ts
const layout = await document.defineSlideMaster({
  title: 'BRAND',
  background: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 10,
  },
  margin: [100, 200, 300, 400],
  slideNumber: { x: 800, y: 500, width: 100, height: 30 },
  objects: [
    { kind: 'rect', options: { x: 10, y: 20, width: 30, height: 40 } },
    { kind: 'line', options: { x: 50, y: 60, width: 70, height: 1 } },
    { kind: 'text', text: 'Brand', options: { x: 80, y: 90 } },
    {
      kind: 'placeholder',
      text: 'Title prompt',
      options: { name: 'title_box', type: 'title', x: 100, y: 110 },
    },
  ],
});
expect(layout.name).toBe('BRAND');
expect(layout.margin).toEqual({ top: 100, right: 200, bottom: 300, left: 400 });
expect(layout.shapes.map(({ kind }) => kind)).toEqual(['shape', 'shape', 'text', 'text', 'text']);
```

The fifth object is direct layout slide number. Assert parent master relation/header-footer enablement, selected slide direct number/cache, placeholders, ordering, unique IDs, stable wrappers, six formats, write/reopen margin boundary, and strict invalid/rollback/no-op cases.

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "define slide master"
```

- [ ] **Step 3: Normalize the complete synchronous definition before mutation**

In `master-layout.ts`, closed-read title/master/background/margin/slideNumber/objects. Normalize every object with its existing native normalizer, normalize scalar margin to four sides, require pair sums smaller than slide width/height, allocate omitted placeholder index from its object ordinal, and freeze detached results.

Track margin in `PptxDocument` by layout URI:

```ts
readonly #layoutMargins = new Map<string, Readonly<SlideMasterMargin>>();
```

- [ ] **Step 4: Commit a canonical layout in one transaction**

Create one canonical blank `p:sldLayout preserve="1"` with direct `p:cSld@name`, internal parent-master relationship, and master layout ID. Within the same outer transaction apply background, ordered sync objects, placeholder, slide number, master `p:hf@sldNum="1"`, and margin map only after package operations succeed.

Nested object transactions are allowed; no async work occurs inside the outer callback.

- [ ] **Step 5: Integrate named slide number materialization**

When `addSlide()` selects a layout with a supported direct slide number, materialize a direct slide field with current canonical cache in the same transaction after placeholder owners. Do not create a fixed-id or disabled master field.

- [ ] **Step 6: Run gates, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "define slide master|named slide layout"
pnpm typecheck
pnpm --filter @pptx/codecs build
pnpm --filter @pptx/sdk build
pnpm --filter @jiayunxie/pptx build
git diff --check
git add packages/codecs/src/master.ts packages/sdk/src/master-layout.ts packages/sdk/src/index.ts packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts
git diff --cached --check
git commit -m "feat: define named slide masters"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 9: Async background/image/chart definition objects

**Files:**

- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/model/src/chart-create.internal.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

Add design variants `SlideMasterBackground` with `kind: 'image-source'`, `SlideMasterObject.kind: 'image'`, and `SlideMasterObject.kind: 'chart'`.

- [ ] **Step 1: Add failing async definition tests**

Cover PNG/JPEG/GIF/SVG data/path/URL/Blob/stream/bytes, SVG fallback, image sizing/source rectangle, image background source, all nine chart types, combo chart, two image dedup, two chart workbooks, and mixed object order.

```ts
const layout = await document.defineSlideMaster({
  title: 'ASYNC',
  background: { kind: 'image-source', source: PNG_DATA_URI },
  objects: [
    { kind: 'image', source: PNG_DATA_URI, options: { x: 100, y: 100 } },
    {
      kind: 'chart',
      groups: [{
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      }],
      options: { x: 200, y: 200, width: 400, height: 300 },
    },
  ],
});
expect(layout.shapes.map(({ kind }) => kind)).toEqual(['image', 'chart']);
```

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/sdk/src/index.test.ts packages/model/src/model.test.ts -t "async slide master definition"
```

- [ ] **Step 3: Prepare every async object before the transaction**

Normalize the full definition first. Resolve all image/background sources with existing SDK loaders and MIME assertions. Build all chart workbooks through `prepareChartCreation()`. Preserve object order in an immutable prepared array:

```ts
type PreparedSlideMasterObject =
  | NormalizedRectObject
  | NormalizedLineObject
  | NormalizedTextObject
  | NormalizedPlaceholderObject
  | { readonly kind: 'image'; readonly source: ResolvedImageSource; readonly options: AddImageSourceOptions }
  | { readonly kind: 'chart'; readonly chart: PreparedChartCreation };
```

Do not allocate any OPC part until every promise resolves.

- [ ] **Step 4: Commit prepared resources synchronously**

Inside the existing definition transaction, route raster/SVG bytes to wrapper image commit, `commitPreparedChart()` to layout content, and resolved background bytes to owner-aware background replacement. Verify failure injection after each part/relationship/XML operation rolls everything back.

- [ ] **Step 5: Run gates, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/sdk/src/index.test.ts packages/model/src/model.test.ts -t "slide master definition|chart|image"
pnpm typecheck
pnpm --filter @pptx/model build
pnpm --filter @pptx/sdk build
git diff --check
git add packages/sdk/src/master-layout.ts packages/sdk/src/index.ts packages/model/src/chart-create.internal.ts packages/sdk/src/index.test.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: add async master layout objects"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 10: Whole-definition replace and delete lifecycle

**Files:**

- Modify: `packages/codecs/src/master.ts`
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
replaceSlideMaster(
  layout: SlideLayoutModel,
  options: DefineSlideMasterOptions,
): Promise<void>;

deleteSlideMaster(
  layout: SlideLayoutModel,
  replacement?: SlideLayoutModel,
): void;
```

- [ ] **Step 1: Add failing replace/delete tests**

Test rename, parent relink, complete object/background/margin/number replacement, stable layout wrapper, stable incoming slide relationship, equivalent no-op, shared/exclusive image and chart dependencies, failed replace rollback, delete unused, used-without-replacement rejection, used-with-replacement retarget, default layout deletion, margin cleanup, stale wrapper, URI reuse, and raw codec coexistence.

```ts
const original = await document.defineSlideMaster({ title: 'ORIGINAL' });
const slide = document.addSlide({ masterName: 'ORIGINAL' });
await document.replaceSlideMaster(original, {
  title: 'RENAMED',
  objects: [{ kind: 'text', text: 'Replacement' }],
});
expect(document.layouts.find(({ name }) => name === 'RENAMED')).toBe(original);
expect(slide.relationships.find(({ type }) => type.endsWith('/slideLayout'))?.resolvedTarget)
  .toBe(original.partUri);
```

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts -t "replace delete slide master"
```

- [ ] **Step 3: Prepare replace exactly like define**

Reuse the full normalization/async preparation function. Before mutation, validate target/replacement wrappers belong to this document, are attached, distinct, and still resolve to expected part/content types. Compare a supported canonical snapshot; return without changing package or transient map when equivalent.

- [ ] **Step 4: Replace one layout atomically**

Keep target part URI and master layout ID. Whole-replace only owned `cSld` background/content, direct layout slide number, layout relationships other than its parent-master edge, and transient margin. Reuse compatible exclusive targets; clone/dedup shared targets; collect superseded targets only when graph incoming reaches zero. Parent change updates both layout backlink and old/new master ID lists.

- [ ] **Step 5: Delete with safe retarget and cache invalidation**

Use low-level `deleteLayout()` after validating replacement. Retarget every incoming slide first inside one transaction, update master lists/relationships, delete owned dependencies, then remove margin and SDK cache entry only after success. Mark old wrapper generation stale so later URI reuse cannot revive it.

- [ ] **Step 6: Run lifecycle gates, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts -t "slide master|layout lifecycle"
pnpm typecheck
pnpm --filter @pptx/codecs build
pnpm --filter @pptx/sdk build
git diff --check
git add packages/codecs/src/master.ts packages/sdk/src/master-layout.ts packages/sdk/src/index.ts packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: replace named slide masters"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 11: PptxGenJS 4.0.1 public-output conformance and diagnostics

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/validator/src/index.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS public `defineSlideMaster()`, `addSlide({ masterName })`, public creators, and `write()` bytes only.
- Produces diagnostics `LAYOUT_NAME_DUPLICATE`, `LAYOUT_RELATIONSHIP_INVALID`, `PLACEHOLDER_IDENTITY_AMBIGUOUS`, `PLACEHOLDER_OWNER_MISSING`, and `PLACEHOLDER_DOMAIN_MISMATCH`.

- [ ] **Step 1: Extend only the local public-test facade**

Give the test facade exact declared PptxGenJS 4.0.1 types for six master object variants and placeholder options. Do not add native-only fields to that facade.

- [ ] **Step 2: Generate the valid public matrix**

Create independent PptxGenJS outputs for default/multiple named layouts, solid/transparency/PNG background, scalar/TRBL margin, slide number, rect/line/plain/rich text, PNG/SVG image, nine chart types/combo, six placeholder types, empty placeholders, and populated text/image/chart/table/media placeholders.

- [ ] **Step 3: Compare imported and native final semantics**

For each case compare layout/master/slide relationship topology, direct background, layout object order/kind/transform/text/style, placeholder type/index and inherited geometry, populated owner kind, payload hashes, chart definition/workbook agreement, slide-number visible state, and validator diagnostics. Ignore random GUID/rId, XML formatting, and PptxGen fixed caches.

- [ ] **Step 4: Lock intentional corrections**

Publicly generate unknown name, duplicate title, invalid placeholder type, duplicate placeholder name, falsy zero geometry, fixed slide-number ID, disabled master, and delayed write mutation. Assert PptxGenJS behavior from output/warnings, then assert native strict rejection or canonical correction with zero pre-error mutation.

- [ ] **Step 5: Add diagnostic profile tests**

Construct each unsafe package state directly, run all five compatibility profiles, and assert exact code/severity/part URI/object ID. Healthy native and PptxGenJS valid imports must have zero new errors/warnings.

- [ ] **Step 6: Run gates, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/validator/src/index.test.ts packages/sdk/src/index.test.ts -t "slide master|layout|placeholder"
pnpm typecheck
pnpm --filter @pptx/pptxgenjs-adapter build
pnpm --filter @pptx/validator build
git diff --check
git add packages/pptxgenjs-adapter/src/index.test.ts packages/validator/src/index.test.ts packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "test: compare master layouts with pptxgenjs"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 12: Actual package, browser, gallery, and client verification

**Files:**

- Modify: `packages/pptx/src/index.test.ts`
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Produces root/runtime `masterLayouts: true`, installed declaration evidence, deterministic manifests, browser state, CLI booleans, native/control galleries, validator/render/client evidence.

- [ ] **Step 1: Add root package runtime/type test**

Import `PLACEHOLDER_TYPES`, `PlaceholderType`, `PlaceholderSelector`, `SlideMasterObject`, `DefineSlideMasterOptions`, `SlideLayoutModel`, and `SlideMasterModel` from `@jiayunxie/pptx`. Define a layout, select it, fill at least text/image/chart placeholders, write/reopen, and assert exported classes/constants and persistent final state.

- [ ] **Step 2: Extend installed TypeScript consumer**

Compile an actual tarball consumer that uses the exact public types, calls async `defineSlideMaster()`, checks wrapper methods/properties, uses name and identity selectors, and performs `write()` without workspace imports.

- [ ] **Step 3: Extend Node and installed CLI smoke**

Create one layout with all six object kinds and all six placeholder kinds, add/select slides, populate all domains, replace another layout, delete with retarget, and reopen. Compute `masterLayouts` from exact semantic checks. Add CLI inspect/validate/slides/part-read/diff booleans for layout/master/slide parts and require 0/0 validation for native gallery.

- [ ] **Step 4: Extend real-Chrome smoke**

Use Blob/data URI/Uint8Array sources and one chart. Return a JSON state containing live wrapper identity, layout names, background kinds, margin before write, placeholder identities/kinds, selected relationship targets, reopened margin `null`, payload hashes, chart definition, validation errors, and console/page/network error counts. Add the exact expected object to the smoke assertion.

- [ ] **Step 5: Run focused, full, performance, type, and build gates**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts packages/validator/src/index.test.ts packages/pptx/src/index.test.ts -t "slide master|layout|placeholder"
pnpm --config.verify-deps-before-run=false exec vitest run
RUN_PERF=1 pnpm --config.verify-deps-before-run=false exec vitest run packages/testkit/src/performance.test.ts
pnpm typecheck
pnpm build
```

Record exact passed/skipped counts and performance duration from output.

- [ ] **Step 6: Pack twice and prove determinism**

Run `pnpm --filter @jiayunxie/pptx pack` in two clean temporary copies. For each tarball, list files, hash every `package/dist` file in sorted order, hash the manifest, and hash the tarball. Require equal dist file sets/manifests; record actual file counts and hashes verbatim.

Install the second tarball into a fresh consumer with no workspace links, run Node/type/browser/CLI smoke, and require every new boolean plus all existing booleans to be true.

- [ ] **Step 7: Build native and PptxGenJS galleries**

Generate a native gallery and independent PptxGenJS control covering every definition/object/placeholder/background/number/lifecycle case. Record slide/part/relationship/layout/master/media/chart/workbook counts, relationship ownership, unique IDs, and zero orphans. Run:

```bash
pptx-inspect --json package inspect native-master-layout-gallery.pptx
pptx-inspect --json package validate native-master-layout-gallery.pptx --profile powerpoint-2010
pptx-inspect --json slides list native-master-layout-gallery.pptx
pptx-inspect --json package diff native-master-layout-gallery.pptx pptxgenjs-master-layout-gallery.pptx
```

- [ ] **Step 8: Render and inspect every page**

Render native/control at 180 DPI, run overflow/edge detection, compute minimum non-empty-pixel margins, build montage/contact sheets, and visually inspect every page for inherited objects, prompt/filled placeholders, clipping, font fallback, chart/image/media/table placement, and slide numbers.

- [ ] **Step 9: Run factual LibreOffice and PowerPoint checks**

Save both galleries through LibreOffice 26.8, reopen with the library, revalidate, rerender, and compare page order, relationship chain, layout names, placeholder identities/kinds, text, styles, payload hashes, chart caches/workbooks, backgrounds, and slide numbers. Run PowerPoint 16.112 native/control automation with a minimal control; report only actual loaded/saved/PDF/PPTX outputs and error codes.

- [ ] **Step 10: Commit source verification changes only**

```bash
git status --short
git diff --check
git add packages/pptx/src/index.test.ts scripts/build-npm-package-types.mjs scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed master layout support"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only pre-existing `.pnpm-store/` remains untracked.

---

### Task 13: Publish docs and close the specialty

**Files:**

- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**

- Consumes: final verified API and exact evidence from Tasks 1–12.
- Produces: public examples, persistence/inheritance warnings, compatibility status, exact evidence, and next roadmap item.

- [ ] **Step 1: Add concise public examples**

Document a complete flow:

```ts
const layout = await document.defineSlideMaster({
  title: 'BRAND',
  background: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
  },
  margin: [inches(0.5), inches(0.5), inches(0.5), inches(0.5)],
  objects: [{
    kind: 'placeholder',
    text: 'Presentation title',
    options: {
      name: 'title_box',
      type: 'title',
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
    },
  }],
});

const slide = document.addSlide({ masterName: layout.name });
slide.addText('Quarterly results', { placeholder: 'title_box' });
```

State that PptxGenJS “masterName” selects a layout, ordinary objects remain inherited, placeholder owners materialize on slide creation, layout margin is transient/reopens `undefined`, and background/content/placeholder relationships are editable.

- [ ] **Step 2: Update compatibility and progress matrices**

Move `defineSlideMaster()`/`masterName`/placeholder creation and population, layout/master backgrounds, wrappers, lifecycle, six formats, packed runtimes, PptxGenJS valid output, and client verification into supported. Keep full theme text cascade, percentage coordinates, advanced text/table/media/chart styling, and broader client certification in later items.

Set the remaining sequence to:

```text
advanced text → advanced table/tableToSlides → output/runtime helpers → peer-range full-suite audit
```

- [ ] **Step 3: Publish only actual evidence**

Copy exact focused/full/performance counts, tarball/dist counts and SHA-256, deterministic results, browser JSON state, gallery package counts, validation diagnostics, render margins, LibreOffice normalization/degradation, and PowerPoint result from Task 12. Do not copy Default Color counts or claim a client pass without output files.

- [ ] **Step 4: Run stale-language and diff review**

```bash
git diff --check
rg -n -i --glob '!docs/superpowers/**' "next .*master|master.*next|下一.*master/layout|defineSlideMaster.*未支持|masterName.*未支持" README.md packages/pptx/README.md docs CHANGELOG.md
git diff -- README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

Current public roadmap references must say master/layout/placeholder is complete and advanced text is next. Historical specs/plans remain unchanged.

- [ ] **Step 5: Commit, push, and verify closure**

```bash
git add README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --cached --check
git commit -m "docs: complete master layout placeholder support"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked. Report specialty 15/15 (100%), completed/remaining items and commits, then automatically begin advanced text.

---

## Plan Self-Review

- Spec coverage: named selection, semantic wrappers, direct backgrounds, identity/materialization, six-domain population, sync/async definitions, replace/delete, PptxGenJS, package/browser/client verification, and docs each map to an independently reviewable task.
- Task sizing: placeholder population is split into text/shape, image/chart, and table/media so each reviewer gate has one resource/lifecycle profile.
- Type consistency: `PlaceholderType`, `PlaceholderIdentity`, `PlaceholderSelector`, `AddPlaceholderOptions`, `SlideLayoutModel`, `SlideMasterModel`, `DefineSlideMasterOptions`, and prepared chart signatures are identical across producer and consumer tasks.
- File boundaries: codecs retain package graph, model owns semantic shape/placeholder/background logic, SDK owns source resolution/wrappers/definition orchestration, and tests keep PptxGenJS out of production.
- Placeholder scan: all steps name exact files, interfaces, commands, assertions, expected failure/pass states, and commit messages; no unresolved implementation choices remain.
- Scope: full theme text cascade, percentage coordinates, advanced object styles, auto-page/`tableToSlides`, and broad client certification remain outside this specialty while required placement/margin foundations are included.
