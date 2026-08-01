# Slide Number Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add complete direct slide, layout, and master slide-number create/read/edit/clear support, presentation start-number and lifecycle cache synchronization, and PptxGenJS 4.0.1 public-output parity.

**Architecture:** A package-independent strict codec owns slide-number values, validation, namespace-aware direct-field projection, canonical rendering, and local mutation for slide/layout/master parts. `SlideModel`, `LayoutModel`, and `MasterModel` expose three explicit owners, while `PresentationModel` owns `firstSlideNumber` and atomically synchronizes safely recognized direct slide caches after reorder operations.

**Tech Stack:** TypeScript 5.8, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, JSZip, PptxGenJS 4.0.1, tsup, real-Chrome browser smoke, `pptx-inspect`, LibreOffice/Poppler.

## Global Constraints

- Execute inline in the current task; do not delegate implementation or review.
- Public units are EMU for x/y/width/height, points for margins and font size, and percent for transparency.
- Default geometry is x `0`, y `0`, width `800000`, height `300000`; default text is left aligned, LTR, `en-US`, not bold, and not italic.
- `SlideModel.slideNumber` edits only the direct slide part; `LayoutModel.slideNumber` edits only the layout part; `MasterModel.slideNumber` edits only the master part and its direct `p:hf@sldNum` flag.
- Do not reproduce PptxGenJS 4.0.1 hidden writes into the unique master and `DEFAULT` layout.
- Direct slide fields use cached text `(firstSlideNumber ?? 1) + zero-based slide index`; layout and master fields use `‹#›`.
- `firstSlideNumber` accepts only signed Int32 safe integers; `undefined` removes the direct attribute and restores the OOXML default of `1`.
- Input is descriptor-safe, closed, copied, deep-frozen, and fully validated before package access or mutation.
- Reads are strict, namespace-aware, direct-only, and non-mutating; unsupported or ambiguous state projects to `undefined`.
- Equal semantic assignment, absent clear, and same `firstSlideNumber` are exact no-ops for parts, relationships, graph, and mutation journal.
- Safely recognized supported edits patch only owned spans and preserve unknown attributes, children, extensions, ids, indexes, GUIDs, and neighboring order.
- Explicit assignment may canonicalize one unique unsafe slide-number placeholder only when its id and index are unambiguous; multiple placeholders are rejected before mutation.
- Apply identical behavior to `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm` without changing format profiles or macro/template/slideshow content types.
- Production packages do not import or call PptxGenJS; the locked 4.0.1 package remains adapter/test evidence only.
- Use fresh `/tmp/pptx-slide-number-*` directories for tarballs, fixtures, galleries, renders, validation output, and client round trips.
- Every task ends with focused tests, typecheck where applicable, diff review, one isolated commit, SSH push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage or commit `.pnpm-store/`, generated decks, renders, packed tarballs, or temporary client artifacts.

---

## File Map

- Create `packages/codecs/src/slide-number.ts`: public color, margin, style, options, and normalized value types.
- Create `packages/codecs/src/slide-number.internal.ts`: strict validation, direct-field reader, semantic equality, canonical renderer, local patching, master `p:hf`, and cached-text helpers.
- Create `packages/codecs/src/slide-number.internal.test.ts`: normalization, reader, writer, local-patch, canonicalization, master-flag, no-op, and rollback tests.
- Modify `packages/codecs/src/index.ts`: export the public slide-number types and the shared runtime helpers consumed by model/SDK packages.
- Modify `packages/codecs/src/master.ts`: add live direct `slideNumber` properties to layout and master models.
- Modify `packages/codecs/src/codecs.test.ts`: cover layout/master copy, relink, live identity, and codec registry behavior.
- Modify `packages/model/src/slide.ts`: add live direct `SlideModel.slideNumber` projection.
- Create `packages/model/src/presentation-slide-number.internal.ts`: strict `firstSlideNum` projection and direct slide-cache synchronization.
- Create `packages/model/src/presentation-slide-number.internal.test.ts`: start-number lexical/range, no-op, repair, cache, and rollback tests.
- Modify `packages/model/src/presentation.ts`: add `firstSlideNumber` and invoke cache synchronization after duplicate/move/delete.
- Modify `packages/model/src/model.test.ts`: public model lifecycle, sections, stable identity, six-format, and mutation-journal tests.
- Modify `packages/sdk/src/create.ts`: accept and validate `CreatePresentationOptions.firstSlideNumber` before creating parts.
- Modify `packages/sdk/src/index.ts`: expose created first-slide number and include slide-number compatibility diagnostics in writes.
- Modify `packages/sdk/src/index.test.ts`: zero-input, all-format, diagnostics, write/reopen, and public declaration behavior.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: permanent PptxGenJS public-output import and intentional-difference evidence.
- Modify `packages/validator/src/index.test.ts`: validate canonical and imported slide-number packages.
- Modify `packages/pptx/src/index.test.ts`: root-package public export and runtime coverage.
- Modify `scripts/smoke-npm-package.mjs`: packed Node ESM/CJS, declarations, CLI, and `slideNumbers: true` smoke.
- Modify `scripts/playwright-browser-smoke.js`: real-browser create/edit/reorder/reopen slide-number smoke.
- Modify root/package/API/compatibility/progress/changelog documentation for final closure.

---

### Task 1: Public values, strict normalization, and direct-field reader

**Files:**

- Create: `packages/codecs/src/slide-number.ts`
- Create: `packages/codecs/src/slide-number.internal.ts`
- Create: `packages/codecs/src/slide-number.internal.test.ts`
- Modify: `packages/codecs/src/index.ts`

**Interfaces:**

- Consumes: `LosslessXmlDocument`, `XmlElement`, XML namespace resolution, DrawingML signed/positive coordinate ranges, and OPC part lookup.
- Produces:

```ts
export type SlideNumberColor =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string };

export interface SlideNumberMargins {
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly left?: number;
}

export type SlideNumberMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | SlideNumberMargins;

export interface SlideNumberTextStyleOptions {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

export interface SlideNumberTextStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

export interface SlideNumberOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly align?: 'left' | 'center' | 'right' | 'justify';
  readonly rtl?: boolean;
  readonly valign?: 'top' | 'middle' | 'bottom';
  readonly margin?: SlideNumberMarginInput;
  readonly style?: SlideNumberTextStyleOptions;
}

export interface SlideNumber {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly align: 'left' | 'center' | 'right' | 'justify';
  readonly rtl: boolean;
  readonly valign?: 'top' | 'middle' | 'bottom';
  readonly margin?: SlideNumberMargins;
  readonly style: SlideNumberTextStyle;
}

export type SlideNumberOwnerKind = 'slide' | 'layout' | 'master';

export function normalizeSlideNumberOptions(value: unknown): Readonly<SlideNumber>;

export function readSlideNumber(
  pkg: OpcPackage,
  ownerPartUri: string,
  ownerKind: SlideNumberOwnerKind,
): Readonly<SlideNumber> | undefined;
```

`packages/codecs/src/index.ts` must export `normalizeSlideNumberOptions`, `readSlideNumber`, and the Task 2 mutation
helpers from the root entry because `@pptx/model` consumes the shared implementation through the package boundary.

- [ ] **Step 1: Add the public type declarations and failing defaults test**

Create the interfaces above and compile this exact default expectation:

```ts
expect(normalizeSlideNumberOptions({})).toEqual({
  x: 0,
  y: 0,
  width: 800_000,
  height: 300_000,
  align: 'left',
  rtl: false,
  style: { lang: 'en-US', bold: false, italic: false },
});
```

Run:

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/slide-number.internal.test.ts
```

Expected: FAIL because `normalizeSlideNumberOptions()` has not been implemented.

- [ ] **Step 2: Implement descriptor-safe normalization and immutable snapshots**

Implement the normalizer around one closed key set and one frozen return path:

```ts
const OPTION_KEYS = new Set([
  'x', 'y', 'width', 'height', 'align', 'rtl', 'valign', 'margin', 'style',
]);

export function normalizeSlideNumberOptions(value: unknown): Readonly<SlideNumber> {
  const input = ownDataRecord(value, OPTION_KEYS, 'Slide number options');
  const style = normalizeSlideNumberStyle(input.style);
  const margin = normalizeMargins(input.margin);
  return deepFreeze({
    x: coordinate(input.x, 0, 'x'),
    y: coordinate(input.y, 0, 'y'),
    width: extent(input.width, 800_000, 'width'),
    height: extent(input.height, 300_000, 'height'),
    align: closedEnum(input.align, ['left', 'center', 'right', 'justify'], 'left'),
    rtl: strictBoolean(input.rtl, false, 'rtl'),
    ...(input.valign === undefined ? {} : { valign: closedEnum(input.valign, ['top', 'middle', 'bottom']) }),
    ...(margin === undefined ? {} : { margin }),
    style,
  });
}
```

Cover explicit zero; integer EMU quantization; positive width/height; `1..4000pt` font-size quantization to `0.01pt`;
signed Int32 point-margin conversion; six-digit uppercase sRGB; supported scheme tokens; `0..100` transparency quantized to
`0.001%`; and transparency-only `tx1`. Reject null, array, class/inherited object, accessor, symbol key, unknown key,
sparse/nonordinary tuple, nested exotic object, invalid XML string, wrong boolean, invalid enum, non-finite/range value, and invalid color.

- [ ] **Step 3: Add strict canonical and PptxGenJS reader fixtures**

Use real namespace declarations and assert both native `a:rPr` and PptxGenJS `a:defRPr` + field override precedence:

```ts
expect(readSlideNumber(pkg, '/ppt/slides/slide1.xml', 'slide')).toEqual({
  x: 0,
  y: 0,
  width: 800_000,
  height: 300_000,
  align: 'center',
  rtl: false,
  valign: 'middle',
  margin: { top: 2, right: 3, bottom: 4, left: 5 },
  style: {
    fontFamily: 'Aptos',
    fontSize: 18,
    lang: 'zh-CN',
    bold: true,
    italic: true,
    color: { kind: 'srgb', value: 'FF3399' },
    transparency: 25,
  },
});
```

Require detached, deeply frozen results and ignore cached `a:t` in semantic equality.

- [ ] **Step 4: Implement the namespace-aware direct reader**

Resolve only this direct chain and require uniqueness at every owned level:

```ts
const target = directSlideNumberTarget(xml, ownerKind);
if (!target || (ownerKind === 'master' && target.masterFlag === false)) return undefined;
return deepFreeze({
  x: target.offX,
  y: target.offY,
  width: target.extentX,
  height: target.extentY,
  align: target.align,
  rtl: target.rtl,
  ...(target.valign === undefined ? {} : { valign: target.valign }),
  ...(target.margin === undefined ? {} : { margin: target.margin }),
  style: target.style,
});
```

Reject wrong namespaces, descendant traps, duplicate `cSld`/`spTree`/placeholder/`spPr`/`xfrm`/off/ext/`txBody`/paragraph/
field/text, invalid or duplicate shape id, unsupported field type, ordinary run/break/additional field, malformed lexical
coordinate/style, conflicting defaults, extra paragraphs, and master `p:hf@sldNum` false tokens.

Export the shared runtime functions explicitly from `packages/codecs/src/index.ts`:

```ts
export * from './slide-number.js';
export {
  normalizeSlideNumberOptions,
  readSlideNumber,
  replaceSlideNumber,
  replaceSlideNumberCachedText,
} from './slide-number.internal.js';
```

- [ ] **Step 5: Run focused gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/slide-number.internal.test.ts
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
git add packages/codecs/src/slide-number.ts packages/codecs/src/slide-number.internal.ts \
  packages/codecs/src/slide-number.internal.test.ts packages/codecs/src/index.ts
git diff --cached --check
git commit -m "feat: read strict slide numbers"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review every design value against the tests, confirm all invalid inputs fail before package access, and confirm reads leave
`xml.changed === false` and source bytes unchanged.

---

### Task 2: Canonical creation, lossless local patching, clear, and master flag

**Files:**

- Modify: `packages/codecs/src/slide-number.internal.ts`
- Modify: `packages/codecs/src/slide-number.internal.test.ts`

**Interfaces:**

- Consumes: Task 1 `normalizeSlideNumberOptions()`, strict target inspection, immutable `SlideNumber`, and owner kind.
- Produces:

```ts
export function replaceSlideNumber(
  pkg: OpcPackage,
  ownerPartUri: string,
  ownerKind: SlideNumberOwnerKind,
  value: SlideNumberOptions | undefined,
  cachedText: string,
): void;

export function replaceSlideNumberCachedText(
  pkg: OpcPackage,
  slidePartUri: string,
  cachedText: string,
): boolean;
```

- [ ] **Step 1: Add failing create and canonical XML tests**

Assign a complete value and assert the direct structure, schema order, unique ids, and cache:

```ts
replaceSlideNumber(pkg, slideUri, 'slide', {
  x: 0,
  y: 10,
  width: 900_000,
  height: 400_000,
  align: 'justify',
  rtl: true,
  valign: 'bottom',
  margin: [1, 2, 3, 4],
  style: {
    fontFamily: 'Aptos',
    fontSize: 20,
    lang: 'zh-CN',
    bold: true,
    italic: true,
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 12.5,
  },
}, '7');
expect(readSlideNumber(pkg, slideUri, 'slide')).toMatchObject({
  align: 'justify', rtl: true, valign: 'bottom',
});
```

Assert a minimum available positive `p:cNvPr@id`, preferred placeholder index `4294967295` with minimum-free fallback,
fixed field GUID, `a:alpha`, latin/ea/cs family, and insertion immediately before direct `p:extLst`.

- [ ] **Step 2: Implement canonical rendering and owner insertion**

Render one canonical field with escaped values and explicit supported attributes:

```ts
function renderSlideNumberShape(
  value: SlideNumber,
  shapeId: number,
  placeholderIndex: number,
  cachedText: string,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="Slide Number ${shapeId}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>` +
    `<p:ph type="sldNum" sz="quarter" idx="${placeholderIndex}"/></p:nvPr></p:nvSpPr>` +
    renderSlideNumberProperties(value) + renderSlideNumberTextBody(value, cachedText) + `</p:sp>`;
}
```

Use `a:noFill` shape/line, rectangle preset geometry, `anchor`, margin attributes, paragraph `algn`/`rtl`, one
`a:fld type="slidenum"`, supported field `a:rPr`, and `a:endParaRPr lang`.

- [ ] **Step 3: Add failing no-op, local-patch, clear, and opaque-repair tests**

Start with unknown attributes/children/extensions around every supported region, perform one supported edit, and require
all unknown source fragments and shape order to survive. Require:

```ts
const before = snapshotPackage(pkg);
replaceSlideNumber(pkg, slideUri, 'slide', readSlideNumber(pkg, slideUri, 'slide'), '7');
expect(snapshotPackage(pkg)).toEqual(before);

replaceSlideNumber(pkg, slideUri, 'slide', undefined, '7');
expect(readSlideNumber(pkg, slideUri, 'slide')).toBeUndefined();
```

Also cover unique unsafe-placeholder canonicalization with preserved valid id/index, rejection of ambiguous id/index,
multiple-placeholder clear rejection, absent clear no-op, PptxGenJS defaults removal, and unrelated ordinary fields.

- [ ] **Step 4: Implement supported-span patching and explicit canonicalization**

Select the write path without guessing:

```ts
if (value === undefined) {
  removeUniqueDirectSlideNumberPlaceholder(xml, ownerKind);
} else if (target.kind === 'supported') {
  patchSupportedSlideNumber(xml, target, normalized, cachedText);
} else if (target.kind === 'unique-opaque' && target.validId !== undefined && target.validIndex !== undefined) {
  xml.replaceElement(target.shape, renderSlideNumberShape(
    normalized, target.validId, target.validIndex, cachedText,
  ));
} else {
  throw new Error(`Slide number state is ambiguous: ${ownerPartUri}`);
}
```

Patch transform, body properties, paragraph attributes, cache, field style, and owned default-style spans only. Update or
remove PptxGenJS `defRPr` values that would otherwise override the assigned field style while retaining unknown defaults.

- [ ] **Step 5: Implement master `p:hf@sldNum` enable/disable and rollback tests**

For master assignment set the unique/created direct flag to `1`; for clear set it to `0` while preserving `hdr`, `ftr`,
`dt`, qualified lookalikes, and unknown attributes:

```ts
replaceSlideNumber(pkg, masterUri, 'master', {}, '‹#›');
expect(readDirectMasterFlag(pkg, masterUri)).toBe(true);
replaceSlideNumber(pkg, masterUri, 'master', undefined, '‹#›');
expect(readDirectMasterFlag(pkg, masterUri)).toBe(false);
```

Insert missing `p:hf` before direct `p:txStyles` or `p:extLst`, otherwise after `p:sldLayoutIdLst` or at root tail. Inject
failure after shape/flag patches and require part bytes, relationships, graph, and mutation journal to roll back exactly.

- [ ] **Step 6: Run focused gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/codecs/src/slide-number.internal.test.ts
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
git add packages/codecs/src/slide-number.internal.ts packages/codecs/src/slide-number.internal.test.ts
git diff --cached --check
git commit -m "feat: edit direct slide number fields"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review id/index allocation, OOXML schema order, escape handling, PptxGenJS style precedence, local-span isolation, exact no-op,
opaque-repair boundaries, master flag behavior, and rollback before committing.

---

### Task 3: Slide API, presentation start number, and lifecycle cache synchronization

**Files:**

- Modify: `packages/model/src/slide.ts`
- Create: `packages/model/src/presentation-slide-number.internal.ts`
- Create: `packages/model/src/presentation-slide-number.internal.test.ts`
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/model/src/index.ts`

**Interfaces:**

- Consumes: Task 2 `readSlideNumber()`, `replaceSlideNumber()`, and `replaceSlideNumberCachedText()`.
- Produces:

```ts
export function normalizeFirstSlideNumber(value: unknown): number;
export function readFirstSlideNumber(xml: LosslessXmlDocument): number | undefined;
export function replaceFirstSlideNumber(xml: LosslessXmlDocument, value: number | undefined): boolean;
export function synchronizeSlideNumberCaches(model: PresentationModel): void;

export class SlideModel {
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}

export class PresentationModel {
  get firstSlideNumber(): number | undefined;
  set firstSlideNumber(value: number | undefined);
}
```

`packages/model/src/index.ts` exports `normalizeFirstSlideNumber` so `@pptx/sdk` uses the same signed-Int32 rule before
creating package parts; the other presentation slide-number helpers remain model-internal.

- [ ] **Step 1: Add failing live `SlideModel.slideNumber` tests**

Create a zero-input package, add a slide, assign defaults, reopen, edit one property, exact-no-op the same value, clear,
and verify the direct field cache uses the current live ordinal:

```ts
const slide = model.addSlide();
slide.slideNumber = {};
expect(slide.slideNumber).toEqual({
  x: 0, y: 0, width: 800_000, height: 300_000,
  align: 'left', rtl: false,
  style: { lang: 'en-US', bold: false, italic: false },
});
expect(readDirectSlideNumberCache(pkg, slide.partUri)).toBe('1');
```

- [ ] **Step 2: Delegate the slide property through the shared codec**

Add imports from `@pptx/codecs` and use the live presentation order:

```ts
get slideNumber(): Readonly<SlideNumber> | undefined {
  return readSlideNumber(this.presentation.opcPackage, this.partUri, 'slide');
}

set slideNumber(value: SlideNumberOptions | undefined) {
  replaceSlideNumber(
    this.presentation.opcPackage,
    this.partUri,
    'slide',
    value,
    String(this.presentation.effectiveSlideNumber(this)),
  );
}
```

Expose an internal safe-integer `effectiveSlideNumber(slide)` and reject a detached slide model not found in the current
presentation order.

- [ ] **Step 3: Add failing `firstSlideNumber` lexical, repair, and no-op tests**

Cover absent/custom/zero/negative/Int32 min/max, invalid lexical direct attribute, duplicate direct attribute, range errors,
same-value no-op, repair, and clear:

```ts
expect(model.firstSlideNumber).toBeUndefined();
model.firstSlideNumber = -10;
expect(model.firstSlideNumber).toBe(-10);
model.firstSlideNumber = undefined;
expect(model.firstSlideNumber).toBeUndefined();
```

Require `normalizeFirstSlideNumber()` to reject floats, `NaN`, infinities, strings, booleans, and values outside
`[-2147483648, 2147483647]` before parsing any package part.

- [ ] **Step 4: Implement direct start-number projection and atomic cache synchronization**

Use one outer transaction and safe-integer cache arithmetic:

```ts
set firstSlideNumber(value: number | undefined) {
  const normalized = value === undefined ? undefined : normalizeFirstSlideNumber(value);
  this.opcPackage.transaction(() => {
    const { xml } = this.parsePresentation();
    if (!replaceFirstSlideNumber(xml, normalized)) return;
    this.setXmlPart(this.presentationPartUri, xml.serialize());
    synchronizeSlideNumberCaches(this);
  });
}
```

Getter returns only one legal direct unqualified `firstSlideNum`; invalid/ambiguous lexical state returns `undefined` and
remains byte-identical until explicit assignment.

- [ ] **Step 5: Add failing duplicate/move/delete/section lifecycle tests**

Build four slides with a mix of supported and unsupported direct fields, a custom start, and sections. Require duplicate
to change only the duplicate cache; move/delete to update only recognized affected slide caches; section operations to
leave caches untouched; source and live identities to remain stable:

```ts
model.firstSlideNumber = 10;
const duplicate = model.duplicateSlide(0);
expect(readDirectSlideNumberCache(pkg, duplicate.partUri)).toBe('14');
model.moveSlide(model.slides.indexOf(duplicate), 0);
expect(model.slides.map((slide) => readDirectSlideNumberCache(pkg, slide.partUri)))
  .toEqual(['10', '11', '12', '13', '14']);
```

- [ ] **Step 6: Hook lifecycle mutations and prove total rollback**

At the end of the existing outer duplicate/move/delete transactions invoke cache synchronization after the presentation
order is final:

```ts
const duplicate = this.attachSlide(slideUri);
synchronizeSlideNumberCaches(this);
return duplicate;
```

Do the equivalent after move and delete. Inject failure during the final affected slide patch and assert restoration of
presentation XML, every slide part, relationships, graph, model identities, sections, and mutation journal.

- [ ] **Step 7: Run focused gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/codecs/src/slide-number.internal.test.ts \
  packages/model/src/presentation-slide-number.internal.test.ts packages/model/src/model.test.ts
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
git add packages/model/src/slide.ts packages/model/src/presentation-slide-number.internal.ts \
  packages/model/src/presentation-slide-number.internal.test.ts packages/model/src/presentation.ts \
  packages/model/src/model.test.ts packages/model/src/index.ts
git diff --cached --check
git commit -m "feat: synchronize slide number lifecycle"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review model liveness, index arithmetic, unsupported-field preservation, affected-slide isolation, nested transaction
behavior, stable sections, and complete rollback before committing.

---

### Task 4: Layout/master owner APIs, SDK creation, and all-six-format behavior

**Files:**

- Modify: `packages/codecs/src/master.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: Tasks 1–3 shared codec and `PresentationModel.firstSlideNumber`.
- Produces:

```ts
export class LayoutModel {
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}

export class MasterModel {
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}

export interface CreatePresentationOptions {
  readonly firstSlideNumber?: number;
}
```

- [ ] **Step 1: Add failing explicit layout/master owner tests**

Assign three distinct values to slide, layout, and master, then clear each owner separately:

```ts
slide.slideNumber = { x: 100 };
layout.slideNumber = { x: 200 };
master.slideNumber = { x: 300 };
expect([slide.slideNumber?.x, layout.slideNumber?.x, master.slideNumber?.x])
  .toEqual([100, 200, 300]);
layout.slideNumber = undefined;
expect(slide.slideNumber?.x).toBe(100);
expect(master.slideNumber?.x).toBe(300);
```

Require layout/master cache `‹#›`, master flag enable/disable, no hidden cross-owner mutation, detached/frozen reads, copy
preservation, relink independence, model identity stability, and delete behavior.

- [ ] **Step 2: Add live layout/master properties**

Delegate both models within the codec package transaction:

```ts
get slideNumber(): Readonly<SlideNumber> | undefined {
  return readSlideNumber(this.codec.pkg, this.partUri, 'layout');
}

set slideNumber(value: SlideNumberOptions | undefined) {
  replaceSlideNumber(this.codec.pkg, this.partUri, 'layout', value, '‹#›');
}
```

Use `'master'` for `MasterModel` and keep the same cache. Do not search for or write any related slide/layout/master part.

- [ ] **Step 3: Add failing create-option validation tests**

Require validation before the first `setPart()` call and exact public output:

```ts
const document = PptxDocument.create({ firstSlideNumber: 0 });
expect(document.firstSlideNumber).toBe(0);
expect(new TextDecoder().decode(
  document.opcPackage.requirePart(document.presentationPartUri).bytes,
)).toContain(' firstSlideNum="0"');
```

Spy on `OpcPackage.create()`/mutation journal for invalid values and require zero parts/relationships/mutations.

- [ ] **Step 4: Thread `firstSlideNumber` into SDK creation**

Normalize at the top of `createPresentationPackage()` before `OpcPackage.create()`, pass the normalized value to
`presentationXml()`, and render the direct root attribute exactly once:

```ts
const firstSlideNumber = options.firstSlideNumber === undefined
  ? undefined
  : normalizeFirstSlideNumber(options.firstSlideNumber);
const pkg = OpcPackage.create();

function renderFirstSlideNumberAttribute(value: number | undefined): string {
  return value === undefined
    ? ''
    : ` firstSlideNum="${value}"`;
}

const firstSlideAttribute = renderFirstSlideNumberAttribute(firstSlideNumber);
```

Add `firstSlideNumber: number | undefined` to the existing `presentationXml()` parameters and interpolate
`${firstSlideAttribute}` in its existing `p:presentation` opening tag immediately after `${rtlAttribute}`. Keep
`PptxDocument.create()` from setting the value a second time, so creation has one deterministic part-write history.

- [ ] **Step 5: Add all-six-format and root-surface tests**

For `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`, create with start `5`, add direct slide/layout/master numbers, write,
open twice, edit, duplicate/move/delete, clear, and require preserved format/content type, exact public values, valid caches,
master flag, declarations, and `validatePackage()` errors `[]`.

```ts
for (const format of ['pptx', 'pptm', 'potx', 'potm', 'ppsx', 'ppsm'] as const) {
  const first = PptxDocument.create({ format, firstSlideNumber: 5 });
  first.addSlide().slideNumber = {};
  const second = await PptxDocument.open(await first.write());
  expect(second.format).toBe(format);
  expect(second.firstSlideNumber).toBe(5);
  expect(second.slides[0]?.slideNumber).toBeDefined();
}
```

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/codecs/src/slide-number.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/presentation-slide-number.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false build
git diff --check
git add packages/codecs/src/master.ts packages/codecs/src/codecs.test.ts \
  packages/sdk/src/create.ts packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: expose slide number owners"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review direct ownership, master flag semantics, copy/relink behavior, early SDK validation, deterministic creation history,
six-format profile preservation, and exported declarations before committing.

---

### Task 5: PptxGenJS public-output conformance and compatibility diagnostics

**Files:**

- Modify: `packages/codecs/src/slide-number.internal.ts`
- Modify: `packages/codecs/src/slide-number.internal.test.ts`
- Modify: `packages/codecs/src/index.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/validator/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**

- Consumes: locked `pptxgenjs@4.0.1`, adapter `importPptxGenJS()`, public SDK/model/codec APIs, and validator profiles.
- Produces:

```ts
export function slideNumberDiagnostics(
  pkg: OpcPackage,
  ownerPartUri: string,
  ownerKind: SlideNumberOwnerKind,
  expectedCachedText: string,
  compatibility: string,
): CodecDiagnostic[];
```

and permanent public-output conformance tests with explicit intentional-difference assertions.

- [ ] **Step 1: Generate PptxGenJS fixtures using public APIs only**

Create omitted, empty, sRGB styled, theme-color, left/center/right, uniform/four-side margin, and top/middle/bottom cases,
plus `defineSlideMaster({ slideNumber })` and a named-layout slide:

```ts
const source = new PptxGenJS();
source.layout = 'LAYOUT_WIDE';
const slide = source.addSlide();
slide.slideNumber = {
  x: 0, y: 0, w: 1, h: 0.3,
  align: 'center', valign: 'mid',
  fontFace: 'Aptos', fontSize: 18, color: 'FF3399', bold: true,
  margin: [1, 2, 3, 4], lang: 'zh-CN',
};
const bytes = new Uint8Array(await source.write({ outputType: 'arraybuffer' }));
const imported = await importPptxGenJS(bytes);
expect(imported.slides[0]?.slideNumber).toMatchObject({ align: 'center' });
```

Do not read `_slides`, `_slideNumberProps`, or another PptxGenJS private field.

- [ ] **Step 2: Compare supported final state structurally and semantically**

Compare direct owner count, geometry in EMU, horizontal/vertical alignment, margins, font family/size/lang, sRGB/theme color,
bold, field type, cached slide text, and owner locality. Import PptxGenJS layout fields even when the renderer does not display
them. Require its master shape with `p:hf@sldNum="0"` to remain an unsupported/disabled master getter state.

- [ ] **Step 3: Lock intentional corrections**

Assert the observed PptxGenJS differences without reproducing them:

```ts
expect(pptxGenShapeId).toBe(25);
expect(nativeShapeId).not.toBe(25);
expect(pptxGenMasterFlag).toBe(false);
expect(nativeMasterFlag).toBe(true);
expect(pptxGenLayoutCache).toMatch(/^10\d{2}$/);
expect(nativeLayoutCache).toBe('‹#›');
expect(pptxGenMasterCache).toBe('null');
expect(nativeMasterCache).toBe('‹#›');
```

Also lock zero width/height fallback, justify→left fallback, ignored italic/transparency/RTL, and hidden ancestor mutation.
Require native explicit zero preservation where legal, strict positive extent rejection, working justify/italic/transparency/RTL,
unique ids, legal caches, enabled master flag, and no cross-owner write.

- [ ] **Step 4: Implement and surface compatibility diagnostics**

Inspect the unique direct placeholder without mutating it and emit stable codes for actual shape-id collision, disabled
master shape, and noncanonical cached text:

```ts
return [
  ...(target.shapeIdCollision ? [{
    severity: 'error' as const,
    code: 'SLIDE_NUMBER_SHAPE_ID_COLLISION',
    message: 'Slide-number shape id collides with another shape id',
    partUri: ownerPartUri,
    compatibility,
  }] : []),
  ...(ownerKind === 'master' && target.masterFlag === false ? [{
    severity: 'warning' as const,
    code: 'SLIDE_NUMBER_MASTER_DISABLED',
    message: 'The master slide-number placeholder is disabled by p:hf',
    partUri: ownerPartUri,
    compatibility,
  }] : []),
  ...(target.cachedText !== expectedCachedText ? [{
    severity: 'warning' as const,
    code: 'SLIDE_NUMBER_CACHE_NONCANONICAL',
    message: `Expected cached slide-number text ${expectedCachedText}`,
    partUri: ownerPartUri,
    compatibility,
  }] : []),
];
```

Export the helper from `@pptx/codecs`. In `PptxDocument.write()`, append diagnostics for every direct slide with its
effective ordinal and every layout/master with `‹#›`. Native canonical output must remain 0 errors/0 warnings; imported
PptxGenJS master/cache mismatches are warnings, while an actual duplicate shape id is an error.

- [ ] **Step 5: Edit imported output and validate every supported profile**

Edit/reopen PptxGenJS slide and layout fields through native APIs; explicitly canonicalize the disabled master; reorder direct
slides; preserve neighboring opaque XML/relationships; and require no unexpected errors under PowerPoint 2010/current,
Keynote current, Google Slides import, and LibreOffice current.

```ts
for (const compatibility of [
  'powerpoint-2010', 'powerpoint-current', 'keynote-current',
  'google-slides-import', 'libreoffice-current',
] as const) {
  await imported.write({ mode: 'permissive', compatibility });
  expect(imported.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
}
```

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts packages/validator/src/index.test.ts packages/pptx/src/index.test.ts
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
git add packages/codecs/src/slide-number.internal.ts packages/codecs/src/slide-number.internal.test.ts \
  packages/codecs/src/index.ts packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.ts packages/sdk/src/index.test.ts packages/validator/src/index.test.ts \
  packages/pptx/src/index.test.ts
git diff --cached --check
git commit -m "feat: diagnose slide number compatibility"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review public-API-only fixture generation, exact owner counts, all observed differences, diagnostic severities/codes,
imported local-patch preservation, validator expectations, and the absence of a PptxGenJS production dependency before
committing.

---

### Task 6: Packed Node, browser, CLI, gallery, and real-client evidence

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Temporary only: clean-build manifests, installed consumer, type fixture, native/PptxGenJS galleries, rendered PNGs,
  validation JSON, LibreOffice round trip, and comparison ledger under `/tmp/pptx-slide-number-*`.

**Interfaces:**

- Consumes: actual packed `@jiayunxie/pptx` artifact, browser conditional export, generated declarations, installed CLI,
  public slide-number APIs, and Task 5 conformance behavior.
- Produces: permanent `slideNumbers: true` Node/browser smoke and reproducible release/client evidence.

- [ ] **Step 1: Extend the permanent packed-package smoke**

Using only installed public ESM/CJS exports, create a deck with direct slide/layout/master numbers, all styles, custom start,
duplicate/move/delete, clear, and reopen. Compile a TypeScript consumer importing every new type:

```ts
import type {
  SlideNumber, SlideNumberColor, SlideNumberMargins,
  SlideNumberMarginInput, SlideNumberOptions,
  SlideNumberTextStyle, SlideNumberTextStyleOptions,
} from '@jiayunxie/pptx';

const options: SlideNumberOptions = {
  align: 'justify', rtl: true, margin: [1, 2, 3, 4],
  style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 20 },
};
```

Inspect the installed CLI JSON for title/slide count and validation. Return `slideNumbers: true` only if public values,
cached texts, owner count, master flag, declarations, inspect, and validate all succeed.

- [ ] **Step 2: Extend the permanent real-Chrome smoke**

From `dist/browser.js`, create with `firstSlideNumber: -2`, add styled direct fields, reorder, write Blob, reopen bytes, and
return geometry/style/cache/diagnostics. Require no console, page, network, or Node-builtin bundle errors:

```js
const document = PptxDocument.create({ firstSlideNumber: -2 });
const first = document.addSlide();
first.slideNumber = { align: 'center', style: { italic: true, transparency: 25 } };
const second = document.duplicateSlide(0);
document.moveSlide(document.slides.indexOf(second), 0);
const reopened = await PptxDocument.open(await document.writeBlob());
```

- [ ] **Step 3: Build twice, hash, pack, install, and run smoke**

Create two clean builds, compute sorted SHA-256 manifests, and require byte-identical dist output. Pack the second build,
install that exact tarball into a fresh consumer without workspace links, then run Node ESM/CJS, declarations, browser,
and installed CLI checks.

```bash
pnpm --config.verify-deps-before-run=false clean
pnpm --config.verify-deps-before-run=false build
pnpm --config.verify-deps-before-run=false test:performance
node scripts/smoke-npm-package.mjs /tmp/pptx-slide-number-pack/package.tgz
node scripts/playwright-browser-smoke.js
```

Record the exact manifest file count and aggregate hash in the final evidence ledger.

- [ ] **Step 4: Generate and structurally inspect native and control galleries**

Generate native slides for default/custom start, left/center/right/justify, top/middle/bottom, uniform/four-side margins,
sRGB/theme/transparency, bold/italic/font/lang/RTL, layout/master owners, duplicate/move/delete final state, and clear. Generate
an independent PptxGenJS public-output control. Use `pptx-inspect` and CLI to record slide titles/order, direct owner count,
field type/cache/style, shape id/index, master flag, relationships, parts, and validator results.

- [ ] **Step 5: Render, inspect, and round-trip with clients**

Render native and PptxGenJS galleries page-by-page at 180 DPI, run overflow checks, and inspect every page. Save the native
gallery through LibreOffice 26.8, reopen, and compare slide count/order, field type/cache/style, direct visibility, owner
parts, and validation. Run PowerPoint automation/control and report only actual open/render/save/reopen outcomes.

- [ ] **Step 6: Run full gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false test
pnpm --config.verify-deps-before-run=false test:performance
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false build
node scripts/playwright-browser-smoke.js
git diff --check
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed slide numbers"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review clean-build manifests, tarball dependency tree, ESM/CJS/types exports, browser diagnostics, CLI JSON, package graph,
gallery counts, rendered pages, LibreOffice ledger, and PowerPoint result. Confirm no `/tmp` artifact is staged.

---

### Task 7: Public documentation and final closure

**Files:**

- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**

- Consumes: final Task 1–6 API, behavior, exact tests, build hashes, package inspection, render, and client evidence.
- Produces: runnable API examples, explicit ownership/support matrix, intentional-difference ledger, and next-work ordering.

- [ ] **Step 1: Document the model, codec, and SDK APIs**

Add runnable examples using public units and all three owners:

```ts
const document = PptxDocument.create({ firstSlideNumber: 5 });
const slide = document.addSlide();
slide.slideNumber = {
  x: inches(11),
  y: inches(6.6),
  width: inches(1),
  height: inches(0.3),
  align: 'right',
  valign: 'middle',
  margin: 0,
  style: {
    fontFamily: 'Aptos',
    fontSize: 12,
    color: { kind: 'scheme', value: 'tx1' },
  },
};
document.layouts[0]!.slideNumber = { align: 'center' };
document.masters[0]!.slideNumber = { align: 'right' };
slide.slideNumber = undefined;
```

Explain direct ownership, detached/frozen getters, defaults, units, strict input, master flag, cache lifecycle, and the
LibreOffice portability reason for direct slide fields.

- [ ] **Step 2: Update compatibility and implementation progress**

Record PptxGenJS 4.0.1 supported output and native intentional corrections: fixed-id avoidance, enabled master flag, legal
layout/master caches, explicit zero/strict extents, working justify/italic/transparency/RTL, and no hidden ancestor writes.
Move slide number from pending to complete. Set the next item to default color, followed by master/layout/placeholder,
advanced text, advanced table/`tableToSlides`, output/runtime helpers, and peer-range full-suite audit.

- [ ] **Step 3: Record exact verification evidence**

Write exact focused/full test counts, performance result, deterministic build manifest hash/count, tarball ESM/CJS/types/
browser/CLI results, gallery slide/part/relationship/field counts, validator errors/warnings, overflow totals, visual inspection,
LibreOffice normalization, and PowerPoint outcome. Do not convert an automation failure into a client round-trip pass.

- [ ] **Step 4: Run final release gates and inspect the complete feature history**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/codecs/src/slide-number.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/presentation-slide-number.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  packages/validator/src/index.test.ts packages/pptx/src/index.test.ts
pnpm --config.verify-deps-before-run=false test
pnpm --config.verify-deps-before-run=false test:performance
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false build
git diff --check
git log --oneline c525f8c..HEAD
git status --short --branch
```

Review public names/declarations, strict-reader semantics, local patching, owner isolation, lifecycle arithmetic/rollback,
six-format output, compatibility claims, client evidence, and the absence of generated artifacts.

- [ ] **Step 5: Commit, push, and verify remote parity**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --cached --check
git commit -m "docs: complete slide number support"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected final divergence is `0 0`; the only remaining untracked entry is the existing `.pnpm-store/` directory.

---

## Self-Review Record

- Spec coverage: Tasks 1–2 cover every public value, strict normalization/reader, PptxGenJS style placement, canonical XML,
  local patch, opaque canonicalization, master flag, no-op, and rollback requirement. Tasks 3–4 cover all three owners,
  `firstSlideNumber`, cache lifecycle, sections, copy/relink, six formats, SDK creation, and live models. Tasks 5–6 cover
  PptxGenJS public-output parity, intentional corrections, validator profiles, packed Node/browser/declarations/CLI, galleries,
  rendering, and client round trips. Task 7 covers all named documentation and next-work ordering.
- Placeholder scan: the plan contains no deferred marker, unspecified implementation step, or cross-task shorthand; every code
  action names its exact interfaces, cases, commands, and expected result.
- Type consistency: every task uses the exact `SlideNumberColor`, `SlideNumberMargins`, `SlideNumberMarginInput`,
  `SlideNumberTextStyleOptions`, `SlideNumberTextStyle`, `SlideNumberOptions`, `SlideNumber`, `SlideNumberOwnerKind`,
  `normalizeSlideNumberOptions()`, `readSlideNumber()`, `replaceSlideNumber()`, `replaceSlideNumberCachedText()`,
  `normalizeFirstSlideNumber()`, `synchronizeSlideNumberCaches()`, and `slideNumberDiagnostics()` names defined above.
