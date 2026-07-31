# Embedded Raster Image Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict synchronous `SlideModel.addImage(bytes, options)` support for embedded PNG, JPEG, and GIF pictures with live `ImageModel` return values, transactional package ownership, PptxGenJS 4.0.1 semantic conformance, packed-runtime coverage, and real-file validation.

**Architecture:** Keep source I/O and image inspection outside the model. Add public raster image types in a focused module, put descriptor-safe normalization and canonical `p:pic` rendering in an internal module, and let `SlideModel` own only the atomic media-part/relationship/slide mutation. Reuse the existing `ImageModel`, shape decoder, transform editor, slide dependency lifecycle, and clone-on-write `replaceData()`.

**Tech Stack:** TypeScript 5.9, Vitest, `@pptx/model`, `@pptx/sdk`, OPC relationships/content types, PptxGenJS 4.0.1 public output, Node/browser package smoke tests, LibreOffice headless, Poppler, `pptx-inspect`, PowerPoint 2010 Open XML validation.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-01-embedded-raster-image-creation-design.md` exactly.
- Public names are `RasterImageContentType`, `AddImageOptions`, and `SlideModel.addImage(bytes, options)`.
- Support exactly `image/png`, `image/jpeg`, and `image/gif` with `.png`, `.jpeg`, and `.gif` parts.
- Require a non-empty direct `Uint8Array`; copy it before mutation and never retain caller-owned bytes.
- Require explicit content type; do not inspect signatures or intrinsic dimensions in this item.
- Use direct EMU/OOXML-angle transforms with 0/0/1in/1in defaults, zero rotation, and false flips.
- Default names use `Image ${zeroBasedImageCountBeforeInsert}`; omitted alt text materializes `preencoded.png`, while explicit empty remains empty.
- Add one new media part and one internal image relationship per call; do not hash-deduplicate data input.
- Insert canonical `p:pic` before direct `p:extLst`, preserve unrelated bytes, and return a live `ImageModel`.
- Keep media part, content type, relationship, slide XML, shape identity, and mutation journal in one transaction.
- Do not add path/URL/data-URI/Blob/stream loading, image signature/dimension detection, sizing, SVG, rounding, transparency, hyperlink, shadow, placeholder, alt-text editing, or public image deletion.
- Execute inline in the root task; do not dispatch subagents.
- Never stage, delete, or modify `.pnpm-store/`.
- End every task with source review, isolated commit, push to `main`, fetch, and `HEAD...origin/main == 0 0`.

---

### Task 1: Public raster types, strict normalization, and canonical picture renderer

**Files:**
- Create: `packages/model/src/image.ts`
- Create: `packages/model/src/image-create.internal.ts`
- Create: `packages/model/src/image-create.internal.test.ts`
- Modify: `packages/model/src/index.ts`

**Interfaces:**
- Produces: public `RasterImageContentType` and `AddImageOptions`.
- Produces internally: `NormalizedEmbeddedRasterImage`.
- Produces internally: `normalizeEmbeddedRasterImage(bytes, options)`.
- Produces internally: `renderEmbeddedRasterImageXml(id, definition, relationshipId, defaultName)`.
- Consumed by: Task 2 `SlideModel.addImage()`.

- [ ] **Step 1: Write failing public-type and default-normalization tests**

Create `image-create.internal.test.ts`. Import the non-existent public types and internal helpers, then lock the
default result:

```ts
const bytes = new Uint8Array([137, 80, 78, 71]);
const options: AddImageOptions = { contentType: 'image/png' };
const normalized = normalizeEmbeddedRasterImage(bytes, options);

expect(normalized).toEqual({
  bytes,
  contentType: 'image/png',
  extension: '.png',
  name: undefined,
  altText: 'preencoded.png',
  x: 0,
  y: 0,
  width: 914_400,
  height: 914_400,
  rotation: 0,
  flipHorizontal: false,
  flipVertical: false,
});
expect(normalized.bytes).not.toBe(bytes);
expect(Object.isFrozen(normalized)).toBe(true);
```

Add table cases for `image/jpeg -> .jpeg` and `image/gif -> .gif`. Mutate the caller bytes and options after
normalization and require the snapshot to remain unchanged.

Run:

```bash
pnpm vitest run packages/model/src/image-create.internal.test.ts
```

Expected: FAIL because the public and internal image modules do not exist.

- [ ] **Step 2: Add failing strict-input tests**

Cover null/array/class/inherited/accessor/symbol/unknown options; missing/runtime-undefined/unknown content type;
non-`Uint8Array` and empty bytes; non-string/invalid-XML name and alt text; non-finite/fractional/unsafe coordinates;
zero/negative width or height; rotation outside `-21600000..21600000`; and non-boolean flips.

Use an accessor counter and require zero invocations:

```ts
let reads = 0;
const accessor = Object.defineProperty({ contentType: 'image/png' }, 'name', {
  get: () => { reads += 1; return 'unsafe'; },
});
expect(() => normalizeEmbeddedRasterImage(bytes, accessor)).toThrow(TypeError);
expect(reads).toBe(0);
```

Also require a null-prototype options object with explicit empty name/alt text and signed x/y to succeed.

- [ ] **Step 3: Add failing renderer tests**

Render ID 7, relationship `rId4`, explicit special-character name/alt text, rotation, and both flips. Parse the
result with `LosslessXmlDocument` and assert exactly one direct `pic`, `nvPicPr/cNvPr/cNvPicPr/nvPr`,
`blipFill/blip/stretch/fillRect`, and `spPr/xfrm/prstGeom/avLst` chain. Lock exact attribute values and escaping.

```ts
expect(source).toContain('id="7"');
expect(source).toContain('r:embed="rId4"');
expect(source).toContain('rot="2700000" flipH="1" flipV="1"');
expect(source).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');
```

Verify omitted name uses the supplied `defaultName`, explicit empty name remains empty, and zero transform flags
do not emit redundant xfrm attributes.

- [ ] **Step 4: Implement public types and strict normalization**

Create `image.ts`:

```ts
import type { Transform } from './units.js';

export type RasterImageContentType = 'image/png' | 'image/jpeg' | 'image/gif';

export interface AddImageOptions extends Partial<Transform> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
}
```

Export it from `packages/model/src/index.ts`. In `image-create.internal.ts`, use a closed option-key set and
`Reflect.ownKeys()`/property descriptors. Implement:

```ts
export interface NormalizedEmbeddedRasterImage {
  readonly bytes: Uint8Array;
  readonly contentType: RasterImageContentType;
  readonly extension: '.png' | '.jpeg' | '.gif';
  readonly name?: string;
  readonly altText: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

export function normalizeEmbeddedRasterImage(
  bytes: unknown,
  options: unknown,
): NormalizedEmbeddedRasterImage;
```

Round nothing: require exact finite safe integers. Copy bytes before returning and freeze the definition. Preserve
own explicit empty strings with `Object.hasOwn()` semantics.

- [ ] **Step 5: Implement the canonical renderer**

Use `escapeXmlAttribute()` for name, alt text, and relationship ID. Emit the exact tree from the design. Add
rotation/flip attributes only when non-default, always emit `a:off`, `a:ext`, rect geometry, no-change-aspect lock,
and stretch/fillRect.

```ts
export function renderEmbeddedRasterImageXml(
  id: number,
  definition: Readonly<NormalizedEmbeddedRasterImage>,
  relationshipId: string,
  defaultName: string,
): string;
```

- [ ] **Step 6: Run focused/type gates and commit Task 1**

```bash
pnpm vitest run packages/model/src/image-create.internal.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/image.ts packages/model/src/image-create.internal.ts \
  packages/model/src/image-create.internal.test.ts packages/model/src/index.ts
git diff --cached --check
git commit -m "feat: define embedded raster images"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: focused tests and typecheck pass; divergence is `0 0`; only `.pnpm-store/` remains untracked.

### Task 2: Atomic `SlideModel.addImage()` and immediate live image state

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 normalizer/renderer and public `AddImageOptions`.
- Produces: `SlideModel.addImage(bytes, options): ImageModel`.
- Produces package state: one raster part, one internal image relationship, and one canonical direct `p:pic`.

- [ ] **Step 1: Add a failing blank-presentation creation test**

Create a zero-input document, add one slide, and call:

```ts
const source = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const image = slide.addImage(source, {
  contentType: 'image/png',
  name: 'Revenue & <logo>',
  altText: 'Quarterly & annual',
  x: inches(1),
  y: inches(2),
  width: inches(3),
  height: inches(2),
  rotation: degrees(45),
  flipHorizontal: true,
});
```

Require `image instanceof ImageModel`, `image.kind === 'image'`, immediate stable identity in `slide.shapes`, exact
name/transform, `/ppt/media/image1.png`, copied bytes/content type, one internal image relationship, canonical
picture XML, and unchanged source after mutating the caller buffer.

Run:

```bash
pnpm vitest run packages/model/src/model.test.ts -t "creates embedded raster images"
```

Expected: FAIL because `SlideModel.addImage()` does not exist.

- [ ] **Step 2: Add failing defaults/order/multiple-image tests**

Start from a slide containing a preset shape and a direct `p:extLst`. Add PNG, JPEG, and GIF. Require:

- shape order remains preset, PNG, JPEG, GIF, extLst;
- IDs are unique and monotonic across all shapes;
- default names are `Image 0`, `Image 1`, `Image 2`;
- default transforms are 0/0/914400/914400 with no rotation/flips;
- parts are `image1.png`, `image1.jpeg`, `image1.gif` when each extension namespace is free;
- every relationship resolves to its own part and no unrelated part/relationship changes.

- [ ] **Step 3: Implement the transactional creation method**

Add public imports for `ImageModel`, `AddImageOptions`, and the Task 1 helpers. Define the relationship type once:

```ts
const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
```

Implement this sequence:

```ts
addImage(bytes: Uint8Array, options: AddImageOptions): ImageModel {
  const definition = normalizeEmbeddedRasterImage(bytes, options);
  return this.presentation.opcPackage.transaction(() => {
    const pkg = this.presentation.opcPackage;
    const mediaPartUri = pkg.allocatePartUri('/ppt/media', 'image', definition.extension);
    pkg.setPart(mediaPartUri, definition.bytes, definition.contentType);
    const relationship = pkg.addRelationship(this.partUri, {
      type: IMAGE_RELATIONSHIP_TYPE,
      target: relativeRelationshipTarget(this.partUri, mediaPartUri),
      targetMode: 'Internal',
    });
    const { xml } = this.parse();
    const shapeTree = requirePresetShapeTree(xml, this.partUri);
    const nextId = allocatePresetShapeId(xml, shapeTree, this.partUri);
    const imageCount = this.shapes.filter(({ kind }) => kind === 'image').length;
    const pictureXml = renderEmbeddedRasterImageXml(
      nextId,
      definition,
      relationship.id,
      `Image ${imageCount}`,
    );
    const extensionList = directElementChildren(shapeTree, 'extLst')[0];
    if (extensionList) xml.replace(extensionList.start, extensionList.start, pictureXml);
    else xml.appendChildXml(shapeTree, pictureXml);
    this.setXml(xml.serialize());
    const image = this.shapes.find(({ id }) => id === nextId);
    if (!(image instanceof ImageModel) || image.kind !== 'image') {
      throw new ModelParseError(`Created image ${nextId} could not be resolved`, this.partUri);
    }
    return image;
  });
}
```

Do not weaken the strict shape-tree/ID audit. Resource allocation intentionally precedes XML mutation so the
existing package transaction proves orphan cleanup on later failure.

- [ ] **Step 4: Add and pass zero-mutation invalid-state tests**

For every Task 1 invalid input category, snapshot parts, relationships, slide bytes, mutation journal, and shape
identity before the call; require them unchanged after rejection. Inject malformed slide root, ambiguous shape
tree, repeated extLst, duplicate/unsafe IDs, and exhausted max ID. Require allocated media parts and relationships
to roll back when the failure occurs after their creation.

- [ ] **Step 5: Run focused regression gates and commit Task 2**

```bash
pnpm vitest run packages/model/src/image-create.internal.test.ts
pnpm vitest run packages/model/src/model.test.ts -t "embedded raster image"
pnpm vitest run packages/model/src/preset-shape.internal.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/slide.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: create embedded raster images"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: creation and failure isolation pass; divergence is `0 0`; only `.pnpm-store/` remains untracked.

### Task 3: Created-image lifecycle, six formats, and SDK surface

**Files:**
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 2 live `ImageModel` and existing transform/replace/dependency lifecycle.
- Guarantees: created images behave like imported images across edit, duplicate, move, rollback, all formats, and reopen.
- Verifies: `@pptx/sdk` and therefore `@jiayunxie/pptx` export the new public types and method.

- [ ] **Step 1: Add failing edit/duplicate/clone-on-write tests**

Create one PNG, edit its transform, and call `replaceData()` while exclusive; require the part URI to remain stable.
Duplicate the slide, require both images to share the target initially, then replace the duplicate bytes and require
a private target with source bytes unchanged. Reopen and assert both images retain name, transform, content type,
bytes, relationships, and stable per-document model identity.

- [ ] **Step 2: Add failing rollback, move, and delete-policy tests**

Inside an outer transaction create an image, edit it, then throw. Require exact restoration of all package parts,
content types, relationships, slide XML, mutation journal, and shape list. Move the created-image slide and require
the target URI to remain stable. Delete it and assert the existing shared-image dependency policy explicitly:
the slide part/relationship disappear while the media target remains preserved until the later dedicated image-GC
item; do not silently change global dependency policy here.

- [ ] **Step 3: Add all-six-format lifecycle tests**

For `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, and `.potm`, create PNG/JPEG/GIF images, write, reopen, edit
transform/data, write again, and reopen again. Require detected format, image order, part content types, bytes,
relationships, names, transforms, and package diagnostics to remain valid.

- [ ] **Step 4: Add public SDK/type coverage**

In `packages/sdk/src/index.test.ts`, import `type AddImageOptions`, `type RasterImageContentType`, and `ImageModel`
from the SDK root. Compile this exact public usage and assert runtime behavior:

```ts
const contentType: RasterImageContentType = 'image/png';
const options: AddImageOptions = { contentType, width: inches(2), height: inches(1) };
const image: ImageModel = document.addSlide().addImage(new Uint8Array([1]), options);
expect(image.sourcePartUri).toMatch(/\/ppt\/media\/image\d+\.png$/);
```

Add `@ts-expect-error` checks for `image/svg+xml`, missing content type, and a path/data options shape.

- [ ] **Step 5: Run lifecycle/public gates and commit Task 3**

```bash
pnpm vitest run packages/model/src/model.test.ts -t "embedded raster image"
pnpm vitest run packages/sdk/src/index.test.ts -t "embedded raster image"
pnpm typecheck
git diff --check
git add packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "test: verify embedded raster image lifecycle"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: all lifecycle/type tests pass; divergence is `0 0`; only `.pnpm-store/` remains untracked.

### Task 4: PptxGenJS 4.0.1 conformance and actual packed package

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves: public PptxGenJS data-input raster output imports into the same supported final semantic state.
- Proves: the actual packed Node/browser/types/CLI surface contains raster creation without workspace dependencies.

- [ ] **Step 1: Add public-output PNG/JPEG/GIF conformance fixtures**

Use PptxGenJS only through its public constructor, `addSlide()`, `slide.addImage()`, and `write()`. Use valid tiny
PNG/JPEG/GIF base64 data URIs and explicit paired options:

```ts
slide.addImage({
  data: PNG_DATA_URI,
  x: 1,
  y: 2,
  w: 3,
  h: 2,
  rotate: 45,
  flipH: true,
  flipV: false,
  objectName: 'Raster image',
  altText: 'Raster alt',
});
```

Import through `importPptxGenJS()`. Compare it with a native image using paired EMU/OOXML-angle values. Assert
shape kind/order, embedded bytes/content type, internal relationship, transform/flips, name/alt text via direct
OOXML, rect geometry, no-change-aspect lock, and stretch/fillRect. Compare semantic state rather than part URI/rId.

- [ ] **Step 2: Lock strict native divergences**

Cover PptxGenJS zero/falsy transform fallback, invalid base64 header handling, extension/MIME quirks, both data and
path presence, invalid runtime types, and console-only rejection. Require the adapter to preserve whatever valid
package it produces, while native rejects corresponding unsupported inputs before any mutation. Do not expose
PptxGenJS path/data input on `SlideModel`.

- [ ] **Step 3: Extend packed Node smoke**

Add `ImageModel` and raster types to the generated smoke program. Create all three content types, verify immediate
identity, parts/bytes/content types, transform editing, exclusive `replaceData()`, write/reopen, and duplicate
clone-on-write. Add `embeddedRasterImages: true` to both the nested API result and top-level result.

- [ ] **Step 4: Extend browser/type smoke**

In the browser bundle script, create a PNG from `Uint8Array`, write a Blob, reopen it, and verify the image. In the
generated TypeScript file, compile the exact `RasterImageContentType`, `AddImageOptions`, and `ImageModel` usage;
lock invalid SVG/missing-content-type/path options with `@ts-expect-error`.

- [ ] **Step 5: Build and test an actual tarball**

```bash
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "embedded raster image"
pnpm typecheck
pnpm build
git diff --exit-code -- packages/pptx/dist
raster_pack_dir=$(mktemp -d /tmp/pptx-raster-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$raster_pack_dir"
cp "$raster_pack_dir/jiayunxie-pptx-0.1.0.tgz" /tmp/jiayunxie-pptx-embedded-raster-0.1.0.tgz
node scripts/smoke-npm-package.mjs /tmp/jiayunxie-pptx-embedded-raster-0.1.0.tgz
git diff --check
```

Expected: adapter conformance, Node/browser/types/CLI smoke, and dist reproducibility pass from the actual tarball.

- [ ] **Step 6: Review, commit, push, and verify Task 4**

```bash
git add packages/pptxgenjs-adapter/src/index.test.ts scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify packaged raster image creation"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence is `0 0`; only `.pnpm-store/` remains untracked.

### Task 5: Real raster gallery, documentation, and release gate

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Temporary only: scripts/output under `raster_gallery_dir=$(mktemp -d /tmp/pptx-embedded-raster.XXXXXX)`

**Interfaces:**
- Documents: public raster creation, strict bytes/content types, defaults, package lifecycle, PptxGenJS boundary, and remaining image/SVG work.
- Proves: actual-tarball PNG/JPEG/GIF output validates, renders, round-trips, and strict-reopens.

- [ ] **Step 1: Load presentation tooling and verify the offline stack**

Read the current Presentations skill before artifact work, load bundled workspace dependencies, then run:

```bash
command -v pptx-inspect
pptx-inspect --json doctor
```

Use the returned LibreOffice/Python/Poppler paths; do not assume system binaries.

- [ ] **Step 2: Generate an actual-tarball four-slide raster gallery**

Create four slides:

1. PNG defaults plus explicit empty alt text;
2. JPEG explicit transform, rotation, and horizontal/vertical flips;
3. GIF plus multiple image ordering/default names;
4. duplicate/share/clone-on-write result and mixed raster content types.

Save the source PPTX, a JSON inventory of shape IDs/names/transforms/relationships/part URIs/content types/SHA-256,
and the gallery SHA-256. Strict-reopen the actual tarball and require every named target to be an `ImageModel` with
exact bytes/content type/transform.

- [ ] **Step 3: Render, inspect, and check overflow**

Render every slide at 180 DPI, run the bundled overflow checker, build a montage, and inspect every page at original
resolution. Require visible non-broken PNG/JPEG/GIF pictures, correct rotations/flips, readable labels, no clipping,
no overlap, and no off-slide objects. Iterate only temporary artifacts until all checks pass.

- [ ] **Step 4: LibreOffice round-trip and direct-state comparison**

Save with an isolated LibreOffice profile, reopen source and round-trip through the actual tarball, and write
`roundtrip-comparison.json`. Compare slide/shape counts and order, names, transforms, picture geometry, alt text,
relationship modes, content types, and payload SHA-256. Record any raster re-encoding, format conversion, metadata
rewrite, or transform normalization explicitly; visual similarity is not a substitute for direct-state equality.

- [ ] **Step 5: Validate both packages**

```bash
pptx-inspect --json package validate "$raster_gallery_dir/embedded-raster-gallery.pptx" --profile powerpoint-2010
pptx-inspect --json package validate "$raster_gallery_dir/roundtrip/embedded-raster-gallery.pptx" --profile powerpoint-2010
pptx-inspect --json package diff "$raster_gallery_dir/embedded-raster-gallery.pptx" "$raster_gallery_dir/roundtrip/embedded-raster-gallery.pptx"
```

Required: original and round-trip have 0 validation errors. Warnings must be investigated and documented rather
than suppressed; a zero-diff round-trip is not required.

- [ ] **Step 6: Update public documentation**

Document `RasterImageContentType`, `AddImageOptions`, `SlideModel.addImage()`, default transform/name/alt text,
detachment, part/relationship ownership, atomic failure, lifecycle, six formats, PptxGenJS public data-output
conformance, packed tests, and LibreOffice evidence. Update the compatibility row from “image creation unsupported”
to partial embedded-raster support. Keep path/URL/data-URI loader, automatic detection/dimensions, sizing, SVG,
rounding/transparency, alt-text editing, hyperlink/shadow/placeholder, and public delete/GC explicitly remaining.

- [ ] **Step 7: Run the complete release gate**

```bash
rg -n -i "addImage.*(unsupported|not supported|未支持创建)|image creation.*(pending|not implemented|尚未)" \
  CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --exit-code -- packages/pptx/dist
node scripts/smoke-npm-package.mjs /tmp/jiayunxie-pptx-embedded-raster-0.1.0.tgz
git diff --check
```

The stale-text search must return no claim that all image creation is absent; scoped remaining-feature statements
are allowed. Expected: all tests pass with only the established normal-suite performance skip, the explicit
1,000-part performance test passes, dist is reproducible, and the actual tarball smoke succeeds.

- [ ] **Step 8: Review, commit, push, and verify Task 5**

Stage only the five documentation files:

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document embedded raster image creation"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence is `0 0`; only `.pnpm-store/` remains untracked.

## Plan Self-Review

- Spec coverage: Task 1 defines strict public values and canonical XML; Task 2 owns atomic creation; Task 3 proves existing lifecycle/six formats/SDK; Task 4 proves PptxGenJS and packed runtimes/types; Task 5 proves real-file compatibility and documentation.
- Scope: source loading, signature/dimension detection, sizing, SVG, picture styling, alt-text editing, hyperlink/shadow/placeholder, and public deletion/GC remain explicitly excluded.
- Type consistency: every task uses the exact `RasterImageContentType`, `AddImageOptions`, `NormalizedEmbeddedRasterImage`, and `SlideModel.addImage()` names defined in Tasks 1–2.
- Unit consistency: native x/y/width/height remain EMU and rotation remains OOXML angle; conformance fixtures pair them with PptxGenJS inch/degree input.
- Package consistency: every create call owns one unique media part plus one internal image relationship and rolls both back with slide XML on failure.
- Lifecycle consistency: duplicate sharing and `replaceData()` isolation reuse existing policy; slide deletion intentionally does not broaden shared-image GC in this item.
- Existing compatibility: imported embedded/external images and `ImageModel.replaceData()` remain unchanged; no I/O enters the model package.
- Placeholder scan: every task has concrete files, signatures, fixtures, commands, expected results, review scope, commit, push, fetch, and divergence checks.
- Execution: standing user direction selects inline execution and excludes subagent dispatch.
