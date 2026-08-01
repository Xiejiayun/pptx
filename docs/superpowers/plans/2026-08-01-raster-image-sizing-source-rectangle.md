# Raster Image Sizing and Source Rectangle Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add bidirectional direct `a:srcRect` support plus intrinsic-aware contain, cover, and pixel-region crop sizing
for PNG/JPEG/GIF images.

**Architecture:** Keep source-rectangle normalization and OOXML mutation in `@pptx/model`, expose a pure sizing
calculator in `@pptx/sdk`, and let `PptxDocument.addImage()` translate explicit sizing into a target transform plus
direct source rectangle only after source inspection. Preserve the existing synchronous model transaction and all
non-sizing defaults.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span edits, OPC transactions, PNG/JPEG/GIF
intrinsic inspection, PptxGenJS 4.0.1 public output, actual npm tarball Node/browser/declaration/CLI smoke,
LibreOffice, Poppler, and PowerPoint 2010 validation.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-01-raster-image-sizing-source-rectangle-design.md` exactly.
- Preserve omitted-sizing output and the existing 0/0/1-inch/1-inch defaults byte-for-byte.
- Model direct values use percent units where `1` means 1%, quantized to 0.001% DrawingML integer percentages.
- Allow negative source-rectangle edges for contain; require every edge `< 100`, pair sums `< 100`, and raw values
  inside signed Int32.
- Keep source pixels, target EMU, and source-rectangle percentages as distinct units.
- Finish options/sizing normalization before source I/O and finish sizing calculation before package mutation.
- Keep `SlideModel.addImage()` synchronous and keep `PptxDocument.addImage()` as the async source boundary.
- Reject high-level direct `sourceRectangle`; callers use `sizing` or the low-level model API.
- Do not add SVG, rounding, transparency, alt-text editing, image hyperlink/shadow/placeholder, or public deletion/GC.
- Execute all tasks inline in the current task; no subagent dispatch is required.
- Never stage, delete, or modify `.pnpm-store/`.
- End every task with focused tests, self-review, an isolated commit, push to `main`, fetch, and
  `HEAD...origin/main == 0 0`.

---

### Task 1: Model source-rectangle normalization and image creation

**Files:**

- Modify: `packages/model/src/image.ts`
- Create: `packages/model/src/image-source-rectangle.internal.ts`
- Create: `packages/model/src/image-source-rectangle.internal.test.ts`
- Modify: `packages/model/src/image-create.internal.ts`
- Modify: `packages/model/src/image-create.internal.test.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: existing `AddImageOptions`, `normalizeEmbeddedRasterImage()`, and canonical `p:pic` renderer.
- Produces: public `ImageSourceRectangle`, internal `NormalizedImageSourceRectangle`,
  `normalizeImageSourceRectangle()`, `renderImageSourceRectangle()`, and creation support through
  `AddImageOptions.sourceRectangle`.

- [ ] **Step 1: Write failing normalization tests**

Add direct tests for explicit zero, positive crop, negative contain, null-prototype input, detachment, deep freeze,
and 0.001% quantization:

```ts
expect(normalizeImageSourceRectangle({
  left: 12.3456,
  top: -20.0004,
  right: 1,
  bottom: 0,
}, 'Image source rectangle')).toEqual({
  left: 12.346,
  top: -20,
  right: 1,
  bottom: 0,
});
```

Reject null/array/class/inherited/accessor/symbol input, missing and extra keys, non-number/NaN/infinity, signed
Int32 overflow after scaling, any edge `>= 100`, `left + right >= 100`, and `top + bottom >= 100`.

Run:

```bash
pnpm exec vitest run packages/model/src/image-source-rectangle.internal.test.ts
```

Expected: fail because the module does not exist.

- [ ] **Step 2: Implement strict normalization and rendering**

Create:

```ts
export interface NormalizedImageSourceRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export function normalizeImageSourceRectangle(
  value: unknown,
  context: string,
): Readonly<NormalizedImageSourceRectangle>;

export function renderImageSourceRectangle(
  value: Readonly<NormalizedImageSourceRectangle>,
  prefix?: string,
): string;
```

Use a closed four-key descriptor-safe reader. Convert each edge with `Math.round(value * 1_000)`, validate signed
Int32 and positive remaining source area, then freeze the normalized object. Render all four attributes in stable
`l/t/r/b` order using integer thousandths.

- [ ] **Step 3: Add the public creation type and failing renderer tests**

Extend `packages/model/src/image.ts`:

```ts
export interface ImageSourceRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface AddImageOptions extends Partial<Transform> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
  readonly sourceRectangle?: ImageSourceRectangle;
}
```

Require creation with `{ left: 25, top: -10, right: 5, bottom: 0 }` to render direct
`<a:srcRect l="25000" t="-10000" r="5000" b="0"/>` between `a:blip` and `a:stretch`. Require omitted
sourceRectangle output to remain identical to the current picture XML.

- [ ] **Step 4: Integrate normalization into atomic image creation**

Add `sourceRectangle` to the model option key set and to `NormalizedEmbeddedRasterImage`. Normalize it before the
transaction and render it with:

```ts
const sourceRectangle = definition.sourceRectangle
  ? renderImageSourceRectangle(definition.sourceRectangle)
  : '';
```

Keep canonical `<a:stretch><a:fillRect/></a:stretch>` after the optional direct rectangle. Add model integration
coverage for returned live `ImageModel`, media/relationship allocation, invalid rectangle zero-mutation, outer
transaction rollback, write/reopen, and all six presentation formats.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/image-source-rectangle.internal.test.ts \
  packages/model/src/image-create.internal.test.ts packages/model/src/model.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/image.ts packages/model/src/image-source-rectangle.internal.ts \
  packages/model/src/image-source-rectangle.internal.test.ts packages/model/src/image-create.internal.ts \
  packages/model/src/image-create.internal.test.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: create images with source rectangles"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected: focused tests and typecheck pass; omitted source-rectangle fixtures do not change; remote divergence is
`0 0`.

---

### Task 2: Existing `ImageModel.sourceRectangle` lifecycle

**Files:**

- Modify: `packages/model/src/image-source-rectangle.internal.ts`
- Modify: `packages/model/src/image-source-rectangle.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: Task 1 normalization/rendering and `SlideModel.resolveShape()`.
- Produces: `readImageSourceRectangle()`, `replaceImageSourceRectangle()`, slide bridge methods, and public
  `ImageModel.sourceRectangle` getter/setter.

- [ ] **Step 1: Write failing direct-read tests**

Parse focused `p:pic` fixtures and require:

```ts
expect(readImageSourceRectangle(xml, picture)).toEqual({
  left: 25,
  top: -10,
  right: 5,
  bottom: 0,
});
```

Cover absent state, explicit zero, missing attributes defaulting to zero, alternate DrawingML prefix, and detached
frozen snapshots. Return `undefined` for duplicate direct rectangles, wrong namespace, noncanonical lexical values,
children/non-whitespace text, Int32 overflow, and invalid remaining source area without mutating input bytes.

- [ ] **Step 2: Write failing replace/repair/clear/no-op tests**

Require whole replacement, absent-to-value insertion after `a:blip` and before `a:tile`/`a:stretch`, repair of one
malformed direct rectangle, and clear of valid/malformed single ownership. Require same normalized assignment and
absent clear to return `false` and preserve exact bytes. Duplicate direct rectangles must throw `ModelParseError`.

Lock preservation of blip attributes/extensions, tile/stretch/fillRect, effects, unknown siblings, namespace
bindings, relationships, and picture transform.

- [ ] **Step 3: Implement namespace-aware direct ownership**

Add:

```ts
export function readImageSourceRectangle(
  xml: LosslessXmlDocument,
  picture: XmlElement,
): Readonly<NormalizedImageSourceRectangle> | undefined;

export function replaceImageSourceRectangle(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  value: Readonly<NormalizedImageSourceRectangle> | undefined,
  partUri: string,
): boolean;
```

Inspect exactly one presentation `p:pic`, exactly one direct presentation `p:blipFill`, and zero/one direct
DrawingML `a:srcRect`. Reuse the in-scope DrawingML prefix from direct siblings; for insertion with no usable prefix,
write canonical `a:srcRect` plus a local DrawingML namespace binding. Do not own `a:tile` or `a:stretch`.

- [ ] **Step 4: Add slide and model bridges**

In `SlideModel` add:

```ts
getImageSourceRectangle(id: number): Readonly<ImageSourceRectangle> | undefined;
setImageSourceRectangle(id: number, value: ImageSourceRectangle | undefined): void;
```

Normalize before starting the package transaction. Resolve the image inside the transaction, require `kind ===
'image'`, apply the focused mutation, and call `setXml()` only when changed.

In `ImageModel` add:

```ts
get sourceRectangle(): Readonly<ImageSourceRectangle> | undefined;
set sourceRectangle(value: ImageSourceRectangle | undefined);
```

- [ ] **Step 5: Add lifecycle integration tests**

Cover create/read/edit/repair/clear, exact no-op journal state, rollback, duplicate slide isolation, move/delete,
write/reopen, malformed existing packages, stable model identity, and all six formats. Confirm source-rectangle edits do
not clone or rewrite media parts/relationships.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/image-source-rectangle.internal.test.ts packages/model/src/model.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/image-source-rectangle.internal.ts \
  packages/model/src/image-source-rectangle.internal.test.ts packages/model/src/slide.ts \
  packages/model/src/shapes.ts packages/model/src/model.test.ts
git commit -m "feat: edit image source rectangles"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 3: Pure contain, cover, and crop calculator

**Files:**

- Create: `packages/sdk/src/raster-image-sizing.ts`
- Create: `packages/sdk/src/raster-image-sizing.test.ts`
- Modify: `packages/sdk/src/index.ts`

**Interfaces:**

- Consumes: public `RasterImageInfo`, `ImageSourceRectangle`, and native positive-safe-integer EMU contract.
- Produces: `RasterImageCropRegion`, `RasterImageSizing`, `RasterImageSizingResult`,
  `normalizeRasterImageSizing()`, and public `calculateRasterImageSizing()`.

- [ ] **Step 1: Write failing public-type and contain/cover branch tests**

Add compile/runtime coverage for landscape, portrait, and square images against landscape, portrait, and square
frames. Lock exact 0.001% results, including:

```ts
expect(calculateRasterImageSizing(
  { contentType: 'image/png', width: 1600, height: 900 },
  { type: 'cover', width: inches(4), height: inches(3) },
)).toEqual({
  width: inches(4),
  height: inches(3),
  sourceRectangle: { left: 12.5, top: 0, right: 12.5, bottom: 0 },
});
```

Require contain to produce symmetric negative edges and equal ratios to produce an explicit four-zero rectangle.

- [ ] **Step 2: Write failing crop tests**

Cover a full-image region, centered subset, fractional pixel boundaries, each source edge, and target aspect ratios
that differ from the crop region. Use:

```ts
{
  type: 'crop',
  width: inches(3),
  height: inches(2),
  source: { x: 400, y: 225, width: 800, height: 450 },
}
```

For a 1600×900 image require `{ left: 25, top: 25, right: 25, bottom: 25 }`.

- [ ] **Step 3: Write failing structural and numeric rejection tests**

Reject unsafe info/sizing/source containers, missing/unknown/accessor/symbol/inherited fields, unknown type, non-safe
or nonpositive target EMU, nonpositive intrinsic dimensions, nonfinite crop coordinates, negative x/y, nonpositive
crop width/height, out-of-bounds region, floating overflow, and contain ratios whose normalized percentage exceeds
signed Int32. Require normalization before caller mutation and frozen nested/result objects.

- [ ] **Step 4: Implement the calculator**

Add the public types from the design and:

```ts
export function normalizeRasterImageSizing(
  value: unknown,
): Readonly<RasterImageSizing>;

export function calculateRasterImageSizing(
  info: RasterImageInfo,
  sizing: RasterImageSizing,
): Readonly<RasterImageSizingResult>;
```

Keep `normalizeRasterImageSizing()` out of the SDK root export; it exists for option integration/tests. Use the exact
cover/contain/crop formulas from the spec, quantize every edge once to 0.001%, validate the source-area contract, and
deep-freeze results. Do not read image bytes or mutate a package.

- [ ] **Step 5: Export public types/functions and run gates**

```bash
pnpm exec vitest run packages/sdk/src/raster-image-sizing.test.ts
pnpm typecheck
git diff --check
git add packages/sdk/src/raster-image-sizing.ts packages/sdk/src/raster-image-sizing.test.ts \
  packages/sdk/src/index.ts
git commit -m "feat: calculate raster image sizing"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 4: Atomic document sizing integration

**Files:**

- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: Task 3 sizing normalizer/calculator, existing source resolver, and Task 1 model creation.
- Produces: strict union `AddImageSourceOptions` and sizing-aware `PptxDocument.addImage()`.

- [ ] **Step 1: Write failing public union tests**

Define the intended type shape:

```ts
type AddImageSourceBaseOptions = Omit<
  AddImageOptions,
  'contentType' | 'sourceRectangle' | 'width' | 'height'
> & {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
};

export type AddImageSourceOptions = AddImageSourceBaseOptions & (
  | {
      readonly sizing?: undefined;
      readonly width?: number;
      readonly height?: number;
    }
  | {
      readonly sizing: RasterImageSizing;
      readonly width?: never;
      readonly height?: never;
    }
);
```

Add `@ts-expect-error` cases for sizing + width/height, any high-level `sourceRectangle`, malformed crop union, SVG
MIME, and low-level missing content type. Compile successful contain/cover/crop examples returning
`Promise<ImageModel>`.

- [ ] **Step 2: Write failing pre-I/O option tests**

Require invalid sizing container/type/frame/source descriptors and width/height conflicts to fail before consuming a
counted async iterable. Mutate caller sizing and nested crop objects immediately after `document.addImage()` and
require the pending result to use the detached pre-await snapshot.

- [ ] **Step 3: Write failing successful integration tests**

For path/data/bytes/Blob/Web stream/async iterable sources, require sizing width/height to become exact transform
extents and calculated source rectangle to be direct in `p:blipFill`. Cover contain, cover, crop, placement x/y,
rotation, flips, name/alt text, returned live identity, write/reopen, and all six formats. Omitted sizing must preserve
the current default output.

- [ ] **Step 4: Write failing zero-mutation tests**

Snapshot parts, relationships, graph, slide XML, shape IDs, model identity, ZIP state, and mutation journal. Require
source failure, MIME assertion failure, extreme contain result, and out-of-bounds crop to leave the snapshot exact.
Invalid slide/options must still fail before source I/O.

- [ ] **Step 5: Implement descriptor-safe option splitting**

Add `sizing` to the high-level closed key set, never add `sourceRectangle`, and normalize/detach sizing before the
first await. Return:

```ts
export interface NormalizedAddImageSourceOptions {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
  readonly imageOptions: Readonly<Omit<AddImageOptions, 'contentType' | 'sourceRectangle'>>;
  readonly sizing?: Readonly<RasterImageSizing>;
}
```

When sizing exists, reject own `width` or `height` before I/O. Without sizing, keep the existing model option flow.

- [ ] **Step 6: Integrate calculation before model mutation**

Update the document method to:

```ts
const resolved = await resolveRasterImageSource(source, normalized.signal);
assertRasterImageContentType(normalized.contentType, resolved);
const placement = normalized.sizing
  ? calculateRasterImageSizing(resolved.info, normalized.sizing)
  : undefined;
return slide.addImage(resolved.bytes, {
  ...normalized.imageOptions,
  ...(placement ?? {}),
  contentType: resolved.info.contentType,
});
```

`placement` supplies only `width`, `height`, and `sourceRectangle`; top-level x/y/rotation/flips/name/alt text stay
in `imageOptions`.

- [ ] **Step 7: Run integration/full gates, review, commit, and push**

```bash
pnpm exec vitest run packages/sdk/src/raster-image-sizing.test.ts \
  packages/sdk/src/raster-image-source.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git commit -m "feat: size images from intrinsic dimensions"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 5: PptxGenJS conformance and packed artifact verification

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Temporary only: packed consumer, source fixtures, declaration fixtures, browser smoke, and comparison output under a
  new `/tmp/pptx-raster-sizing-*` directory

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 public `ImageProps.sizing` and native high-level sizing.
- Produces: locked final-state conformance tests and actual public-package evidence.

- [ ] **Step 1: Add failing contain/cover/crop conformance cases**

Generate public PptxGenJS images with outer `w/h` proportional to fixture intrinsic dimensions and nested sizing for:

- contain horizontal and vertical letterbox;
- cover horizontal and vertical crop;
- equal ratio explicit zero rectangle;
- crop centered source window.

Generate native equivalents from the same path/data sources. Compare final transform, direct `srcRect` integer
percentages, payload bytes, content type, internal relationship, order, name, alt text, rotation, and flips. Normalize
only syntactic `a:stretch` differences; do not normalize source rectangle values.

- [ ] **Step 2: Lock strict runtime divergences**

Cover PptxGenJS truthy fallback and ambiguous outer/nested dimensions without treating malformed output as supported
native state. Require native to reject conflicting dimensions, high-level direct source rectangle, unsafe frame EMU,
out-of-bounds crop, and extreme result before package mutation.

- [ ] **Step 3: Build the actual tarball twice and compare dist hashes**

Use a fresh temporary directory. Build `@jiayunxie/pptx` twice from clean dist state, hash every emitted file in stable
path order, and require identical SHA-256 manifests. Pack the second build and install that exact tarball into a
temporary consumer without workspace links.

- [ ] **Step 4: Run packed Node/browser/declaration/CLI smoke**

Verify from the installed tarball:

- Node ESM contain/cover/crop creation, write, reopen, and direct `sourceRectangle` editing;
- browser bundle import and image sizing with no static Node import;
- declarations for all source-rectangle/sizing types, pure calculator, strict option union, and live getter/setter;
- existing CLI inspect/validate on the generated file;
- no runtime dependency on PptxGenJS or workspace-only paths.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/pptxgenjs-adapter/src/index.test.ts
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
git diff --check
git add packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "test: match PptxGenJS image sizing"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Record the tarball path, two dist manifests, smoke outputs, and exact conformance count for Task 6 documentation.

---

### Task 6: Gallery, client round-trip, and public documentation

**Files:**

- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`
- Temporary only: gallery, renders, validation output, overflow results, and client comparison under the Task 5
  `/tmp/pptx-raster-sizing-*` directory

**Interfaces:**

- Consumes: installed tarball from Task 5 and every public API/evidence result from Tasks 1–5.
- Produces: strict multi-client evidence, current compatibility accounting, and complete public sizing documentation.

- [ ] **Step 1: Generate a multi-ratio sizing gallery through the tarball**

Create at least four slides containing landscape/portrait/square PNG/JPEG/GIF fixtures and visible target-frame guides.
Cover contain/cover horizontal/vertical branches, equal ratio, full/center/edge/fractional crop, direct existing-image
source-rectangle edit/clear, rotation/flips, special-character names, and non-empty alt text.

Require exact slide/shape/picture/part/relationship counts, payload SHA-256, sourceRectangle snapshots, target
transforms, strict reopen, and PowerPoint 2010 validation with zero errors/warnings.

- [ ] **Step 2: Render and visually inspect every page**

Render source gallery at 180 DPI, require every page to be 2400×1350, run the presentation overflow checker, and
inspect every page at original resolution. Confirm contain shows the complete source, cover fills the guide, and crop
matches the labeled source region.

- [ ] **Step 3: LibreOffice round-trip audit**

Open/save through LibreOffice headless, strictly reopen and validate, render every page, and compare page count/size,
picture order/name/non-empty alt text, payload SHA, content type, relationship mode, transform, and source rectangle.
Record each direct-state normalization and maximum numeric delta; never claim byte identity after picture markup
rewrite.

- [ ] **Step 4: Update public docs and compatibility status**

Document:

- low-level `ImageSourceRectangle` units, negative contain values, getter/setter/clear, and atomic behavior;
- pure `calculateRasterImageSizing()` with contain/cover/crop examples;
- high-level `PptxDocument.addImage(..., { sizing })`, target EMU, crop pixel region, and conflict policy;
- exact PptxGenJS parity boundary and strict divergences;
- actual tarball/browser/declaration/CLI evidence and gallery/client results.

Remove contain/cover/crop and `srcRect` from pending status. Make SVG the next image item; keep rounding,
transparency, alt-text editing, image hyperlink/shadow/placeholder, and deletion/GC pending.

- [ ] **Step 5: Run the final release gate**

```bash
pnpm exec vitest run packages/model/src/image-source-rectangle.internal.test.ts \
  packages/model/src/image-create.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/raster-image-sizing.test.ts packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --check
```

Also require clean-build dist reproducibility, actual-tarball Node/browser/declaration/CLI smoke, original and
LibreOffice gallery validation/render/overflow, and direct-state comparison to pass.

- [ ] **Step 6: Review, commit, push, and confirm clean state**

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md
git diff --cached --check
git commit -m "docs: document raster image sizing"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

## Completion Definition

This plan is complete only when all six tasks are checked, six isolated implementation commits are present remotely,
all release gates pass, public docs no longer list contain/cover/crop or direct `srcRect` as unsupported, and the
validated roadmap advances to SVG image creation/editing.
