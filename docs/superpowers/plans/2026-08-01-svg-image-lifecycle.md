# SVG Image Lifecycle Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add complete SVG image create/load/read/edit/fallback/sizing support with PptxGenJS 4.0.1 public-function parity and verified PowerPoint/LibreOffice package behavior.

**Architecture:** Keep the synchronous model responsible for a strict SVG + PNG fallback pair and its atomic OOXML lifecycle. Keep source I/O, SVG inspection, intrinsic sizing, browser Canvas fallback, and Node placeholder fallback in the asynchronous SDK. Recognize existing SVG pictures by namespace-aware extension structure rather than fixed prefixes or MIME alone.

**Tech Stack:** TypeScript 5.8, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, JSZip, PptxGenJS 4.0.1, tsup, Playwright browser smoke, `pptx-inspect`, LibreOffice/Poppler.

## Global Constraints

- Preserve existing `SlideModel.addImage()` raster behavior and every public raster type/function.
- Do not add native, WASM, optional, or runtime SVG-rasterizer dependencies.
- Every generated `.png` part must contain a valid PNG signature; never reproduce PptxGenJS's SVG-bytes-in-PNG-part defect.
- The model accepts only non-empty detached `Uint8Array` SVG and PNG fallback payloads.
- Canonical SVG content type is `image/svg+xml`; the reader must accept relationship-identified LibreOffice `image/svg` parts.
- Canonical SVG extension URI is `{96DAC541-7B7A-43D3-8B79-37D633B846F1}` and namespace is `http://schemas.microsoft.com/office/drawing/2016/SVG/main`.
- High-level SVG sources cover path, HTTP/HTTPS, browser-relative URL, strict base64 data URI, `Uint8Array`, `ArrayBuffer`, Blob/File, Web stream, and async iterable.
- Explicit fallback must inspect as PNG; browser Canvas is second priority; a detached built-in valid transparent PNG is final fallback.
- With no explicit sizing, x/y remain 0 and width/height remain one inch; SVG intrinsic size never silently changes layout.
- All failed operations leave parts, relationships, content types, XML, graph, shape identity/cache, ZIP state, and mutation journal unchanged.
- Every task ends with focused tests, typecheck, diff review, an isolated commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage or commit `.pnpm-store/` or temporary client/gallery artifacts.

---

### Task 1: SVG definition normalization and canonical picture XML

**Files:**

- Modify: `packages/model/src/image.ts`
- Modify: `packages/model/src/image-create.internal.ts`
- Create: `packages/model/src/svg-image-create.internal.ts`
- Create: `packages/model/src/svg-image-create.internal.test.ts`

**Interfaces:**

- Consumes: existing transform/source-rectangle validation and raster picture rendering conventions.
- Produces: `SvgImageContentType`, `AddSvgImageOptions`, `NormalizedEmbeddedSvgImage`, `normalizeEmbeddedSvgImage()`, and `renderEmbeddedSvgImageXml()`.

- [ ] **Step 1: Write failing public type and normalization tests**

Add compile/runtime coverage for:

```ts
const options: AddSvgImageOptions = {
  name: 'Vector & mark',
  altText: 'SVG description',
  x: inches(1),
  y: inches(2),
  width: inches(3),
  height: inches(2),
  rotation: degrees(15),
  flipHorizontal: true,
  sourceRectangle: { left: 12.5, top: 0, right: 12.5, bottom: 0 },
};
const normalized = normalizeEmbeddedSvgImage(svgBytes, pngBytes, options);
```

Require detached copies, frozen definition/source rectangle, default `preencoded.svg` alt text, one-inch defaults,
and the same transform bounds as raster. Reject non-byte/empty payloads, null/array/class/inherited/accessor/symbol/
unknown options, invalid XML strings, invalid transforms, and invalid source rectangles without reading accessors.

- [ ] **Step 2: Write failing canonical XML tests**

Require direct child order and exact relation direction:

```xml
<a:blip r:embed="rIdFallback"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/></a:ext></a:extLst></a:blip>
```

Check escaped name/alt text, direct `a:srcRect` placement, transform/flip/rotation, `a:picLocks`, one direct
extension list, fixed URI/namespace, and no external relationship syntax.

- [ ] **Step 3: Implement the public types and strict definition**

Add:

```ts
export type SvgImageContentType = 'image/svg+xml';
export interface AddSvgImageOptions extends Partial<Transform> {
  readonly name?: string;
  readonly altText?: string;
  readonly sourceRectangle?: ImageSourceRectangle;
}
```

Use a closed-key descriptor-safe read before delegating shared appearance fields to the existing raster normalizer;
never mutate or retain caller input. Store `svgBytes`, `fallbackPngBytes`, and canonical appearance fields in one
frozen definition.

- [ ] **Step 4: Implement canonical renderer with one shared picture skeleton**

Refactor `image-create.internal.ts` only enough to let raster and SVG renderers share the non-visual properties,
source rectangle, and transform skeleton. Keep the current raster XML byte-for-byte unchanged. Render SVG blip XML
from the two relationship ids and constants exported only from the internal SVG module.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/image-create.internal.test.ts \
  packages/model/src/svg-image-create.internal.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/image.ts packages/model/src/image-create.internal.ts \
  packages/model/src/svg-image-create.internal.ts \
  packages/model/src/svg-image-create.internal.test.ts
git diff --cached --check
git commit -m "feat: define canonical SVG pictures"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review that raster render snapshots are unchanged and no source payload/options are retained.

---

### Task 2: Atomic model creation and six-format round-trip

**Files:**

- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: Task 1 `normalizeEmbeddedSvgImage()` and `renderEmbeddedSvgImageXml()`.
- Produces: synchronous `SlideModel.addSvgImage(svgBytes, fallbackPngBytes, options?)` returning a live `ImageModel`.

- [ ] **Step 1: Write failing zero-input creation tests**

Create one SVG picture after a neighbor shape. Require exact shape id/order/live identity, `/ppt/media/imageN.png`,
`/ppt/media/imageN.svg`, effective content types, copied bytes, two internal image relationships, fallback rid on direct
blip, SVG rid on `svgBlip`, and insertion before shape-tree `p:extLst`.

- [ ] **Step 2: Write failing atomic rollback tests**

Reuse the raster malformed-slide matrix and add outer-transaction rollback. Snapshot package parts, relationships,
content types, graph, slide XML, shape cache, ZIP output, and journal before each invalid call. Require exact equality
after payload/options/shape-tree/allocation failure.

- [ ] **Step 3: Implement `SlideModel.addSvgImage()`**

Inside one existing package transaction:

```ts
const definition = normalizeEmbeddedSvgImage(svgBytes, fallbackPngBytes, options);
const fallbackUri = pkg.allocatePartUri('/ppt/media', 'image', '.png');
const svgUri = pkg.allocatePartUri('/ppt/media', 'image', '.svg');
pkg.setPart(fallbackUri, definition.fallbackPngBytes, 'image/png');
pkg.setPart(svgUri, definition.svgBytes, 'image/svg+xml');
const fallbackRelationship = pkg.addRelationship(this.partUri, {
  type: IMAGE_RELATIONSHIP_TYPE,
  target: relativeRelationshipTarget(this.partUri, fallbackUri),
  targetMode: 'Internal',
});
const svgRelationship = pkg.addRelationship(this.partUri, {
  type: IMAGE_RELATIONSHIP_TYPE,
  target: relativeRelationshipTarget(this.partUri, svgUri),
  targetMode: 'Internal',
});
```

Then allocate the shape id, insert rendered XML, update the slide, resolve the same live `ImageModel`, and return it.
Do not add content-type or relationship handling outside `OpcPackage`.

- [ ] **Step 4: Add write/reopen and six-format tests**

For pptx/pptm/potx/potm/ppsx/ppsm require preserved format, picture metadata/transform/source rectangle, two payloads,
two internal relationships, fixed extension URI, and validator errors `[]` after two write/open cycles.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/svg-image-create.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/slide.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create embedded SVG images"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 3: Namespace-aware SVG state and paired clone-on-write editing

**Files:**

- Create: `packages/model/src/svg-image-state.internal.ts`
- Create: `packages/model/src/svg-image-state.internal.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: existing `ImageModel.resolve()`, relationship graph, sibling part allocation, and relationship retargeting.
- Produces: `ImageModel.isSvg`, `svgPartUri`, `fallbackPartUri`, and `replaceSvgData(svgBytes, fallbackPngBytes)`.

- [ ] **Step 1: Write failing namespace/state reader tests**

Cover canonical prefixes, arbitrary alternate prefixes, inherited namespace declarations, PptxGenJS XML, and
LibreOffice `image/svg` parts. Require the direct chain, fixed extension URI, Office SVG namespace, relationship
attribute namespace, internal image relationship type, and existing targets.

Reject or return unsafe state for duplicate extension candidates, wrong namespace with the same local names, wrong
URI/type, unprefixed relationship attribute, external/dangling target, missing fallback, nested non-direct markers,
and ambiguous direct blips. Reads must preserve bytes/journal exactly.

- [ ] **Step 2: Implement namespace resolution and state extraction**

Implement helpers equivalent to:

```ts
function elementNamespaceUri(element: XmlElement): string | undefined;
function attributeNamespaceUri(element: XmlElement, attribute: XmlAttribute): string | undefined;
function readSvgImageState(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  relationships: readonly Relationship[],
  pkg: OpcPackage,
): SvgImageState | undefined;
```

Resolve prefix bindings by walking the element and ancestors. Default namespaces apply only to elements. Return
both reference attributes, relationships, and target URIs in a detached internal state.

- [ ] **Step 3: Add public getter tests and implementation**

Require raster `isSvg === false` and both new URIs `undefined`. For safe SVG require `sourcePartUri ===
fallbackPartUri`, `isSvg === true`, and `svgPartUri` pointing to the vector target before and after reopen. Existing
`sourcePartUri` and `externalUrl` behavior must remain unchanged.

- [ ] **Step 4: Write failing paired replacement lifecycle tests**

Cover exclusive in-place update, duplicate-slide shared targets, same-slide shared relationship ids, independently
shared SVG/fallback targets, noncanonical extensions/content types, outer rollback, write/reopen/edit-again, and
malformed state. Require both new payloads/content types/extensions and no changes to transform, source rectangle,
name, alt text, other shapes, or unknown extension siblings.

- [ ] **Step 5: Implement atomic paired replacement**

Validate/copy both payloads before starting mutation. For each side, overwrite only when target and relationship
reference are exclusive and extension canonical; otherwise allocate `.svg` or `.png`, set the canonical content type,
and retarget only the current reference. Apply both sides in one transaction. If any Office SVG extension candidate is
present, make `replaceData()` reject with `Use replaceSvgData() for SVG images` before changing the fallback.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/svg-image-state.internal.test.ts \
  packages/model/src/model.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/svg-image-state.internal.ts \
  packages/model/src/svg-image-state.internal.test.ts packages/model/src/shapes.ts \
  packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: edit paired SVG image data"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 4: SVG inspection and every source form

**Files:**

- Create: `packages/sdk/src/svg-image-source.ts`
- Create: `packages/sdk/src/svg-image-source.test.ts`
- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`
- Modify: `packages/sdk/src/index.ts`

**Interfaces:**

- Consumes: existing source dispatch, canonical base64 decoder, stream/Blob/path/fetch/abort handling, and raster inspector.
- Produces: `SvgImageInfo`, `ImageInfo`, `ImageSource` aliases, `inspectSvgImage()`, `inspectImage()`, and an internal generic resolved source.

- [ ] **Step 1: Write failing SVG inspector acceptance tests**

Use UTF-8 SVG fixtures for viewBox, unitless/px/in/cm/mm/pt/pc/Q dimensions, one-side derivation, prefixed root
namespace, declaration/comments, and 300×150 default. Require exact positive finite CSS-pixel dimensions and
`contentType: 'image/svg+xml'`.

- [ ] **Step 2: Write failing SVG inspector rejection tests**

Reject non-`Uint8Array`, empty/truncated/non-UTF-8, DTD/ENTITY, multiple roots, HTML, wrong/no SVG namespace,
malformed XML, negative/zero viewBox, and XML text that only contains a nested `<svg>`. Require error messages not to
include source payloads.

- [ ] **Step 3: Implement strict SVG and generic inspection**

Parse with fatal `TextDecoder` and `LosslessXmlDocument`. Resolve the root namespace, parse viewBox/dimensions with
anchored numeric/unit grammar, convert physical units to 96 CSS px, derive missing dimensions, and freeze the result.
`inspectImage()` dispatches raster signatures/truncated raster prefixes to `inspectRasterImage()` and all other bytes
to `inspectSvgImage()`; it never treats arbitrary XML as an image.

- [ ] **Step 4: Write failing generic source tests**

Cover all portable memory forms, local relative/absolute path, HTTP redirect/query URL, browser-relative Fetch,
strict SVG data URI, abort before/during I/O, caller mutation after call, misleading extension/MIME/name, HTTP errors,
malformed SVG, and MIME assertion mismatch. Existing raster resolver must continue rejecting SVG.

- [ ] **Step 5: Refactor source byte loading once and add generic resolution**

Keep current public raster exports. Extract the existing byte-copy and string-I/O path inside
`raster-image-source.ts` so both resolvers call one private loader. Extend strict data-URI media grammar to the four
canonical types while each public resolver enforces its own accepted result. Export generic types from SDK root:

```ts
export type ImageContentType = RasterImageContentType | SvgImageContentType;
export type ImageByteChunk = RasterImageByteChunk;
export type ImageByteStream = RasterImageByteStream;
export type ImageSource = RasterImageSource;
export type ImageInfo = RasterImageInfo | SvgImageInfo;

export interface ResolvedImageSource {
  readonly bytes: Uint8Array;
  readonly info: ImageInfo;
  readonly assertedContentType?: ImageContentType;
}

export function resolveImageSource(
  source: ImageSource,
  signal?: AbortSignal,
): Promise<ResolvedImageSource>;

export function assertImageContentType(
  expected: ImageContentType | undefined,
  resolved: Readonly<ResolvedImageSource>,
): void;
```

- [ ] **Step 6: Run Node/browser gates, review, commit, and push**

```bash
pnpm exec vitest run packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/svg-image-source.test.ts
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
rg -n 'from "node:|from '\''node:|require\("node:' packages/pptx/dist/browser.js
git diff --check
git add packages/sdk/src/svg-image-source.ts packages/sdk/src/svg-image-source.test.ts \
  packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/index.ts
git diff --cached --check
git commit -m "feat: load and inspect SVG image sources"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

The static Node import search must return no matches.

---

### Task 5: Generic sizing, fallback generation, and document integration

**Files:**

- Create: `packages/sdk/src/svg-image-fallback.ts`
- Create: `packages/sdk/src/svg-image-fallback.test.ts`
- Modify: `packages/sdk/src/raster-image-sizing.ts`
- Modify: `packages/sdk/src/raster-image-sizing.test.ts`
- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: Task 2 model creation, Task 4 generic source info, and existing direct source-rectangle sizing.
- Produces: `ImageSizing` aliases, `calculateImageSizing()`, SVG fallback selection, and unified `PptxDocument.addImage()`.

- [ ] **Step 1: Write failing generic sizing tests**

Require SVG contain/cover/equal-ratio/crop to produce the same target transform and quantized source rectangle as
equivalent raster dimensions, including fractional SVG dimensions. Keep existing raster safe-integer validation and
every old output unchanged. Export format-neutral aliases while retaining old names.

- [ ] **Step 2: Implement format-neutral sizing core**

Move ratio/crop math behind:

```ts
export function calculateImageSizing(
  info: ImageInfo,
  sizing: ImageSizing,
): Readonly<ImageSizingResult>;
```

Let `calculateRasterImageSizing()` perform its old strict raster-info normalization before calling the shared core.
SVG info accepts only the exact closed object and positive finite dimensions returned by the inspector.

- [ ] **Step 3: Write failing fallback tests**

Require explicit valid PNG preservation and reject JPEG/GIF/SVG/truncated fallback. In a Node runtime require a
detached, valid, deterministic built-in PNG. With stubbed browser `Image`, object URL, canvas, context, and `toBlob`,
require generated PNG; for decode/context/taint/budget/export failures require the built-in PNG and object URL cleanup.

- [ ] **Step 4: Implement deterministic fallback selection**

Store one valid transparent PNG byte constant, verify it with `inspectRasterImage()` in tests, and always return a
copy. Browser rasterization uses object URL + image decode + Canvas with a maximum 8192 pixels per side and 16,777,216
total pixels. Explicit fallback is resolved with the caller signal and must inspect as PNG.

- [ ] **Step 5: Extend option normalization and type tests**

Add `image/svg+xml` and `fallback` to the closed high-level option surface. Preserve the no-sizing versus sizing
width/height union. Validate the options container, own data properties, signal, content type, and detached sizing
before primary I/O. Do not read a fallback source when the primary resolves to raster. Normalize to:

```ts
export interface NormalizedAddImageSourceOptions {
  readonly contentType?: ImageContentType;
  readonly fallback?: ImageSource;
  readonly signal?: AbortSignal;
  readonly imageOptions: Readonly<Omit<AddSvgImageOptions, 'sourceRectangle'>>;
  readonly sizing?: Readonly<ImageSizing>;
}
```

- [ ] **Step 6: Integrate one-mutation high-level dispatch**

Implement:

```ts
const resolved = await resolveImageSource(source, normalized.signal);
assertImageContentType(normalized.contentType, resolved);
const placement = normalized.sizing
  ? calculateImageSizing(resolved.info, normalized.sizing)
  : undefined;
if (resolved.info.contentType !== 'image/svg+xml') {
  if (normalized.fallback !== undefined) throw new TypeError('fallback is only valid for SVG images');
  return slide.addImage(resolved.bytes, { ...normalized.imageOptions, ...(placement ?? {}), contentType: resolved.info.contentType });
}
const fallback = await resolveSvgFallback(resolved, normalized.fallback, normalized.signal);
return slide.addSvgImage(resolved.bytes, fallback, { ...normalized.imageOptions, ...(placement ?? {}) });
```

Require zero package mutation on primary/fallback/Canvas/abort/sizing/model failure and immediate live identity on success.

- [ ] **Step 7: Run focused/full gates, review, commit, and push**

```bash
pnpm exec vitest run packages/sdk/src/raster-image-sizing.test.ts \
  packages/sdk/src/svg-image-fallback.test.ts packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/svg-image-source.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/sdk/src/svg-image-fallback.ts packages/sdk/src/svg-image-fallback.test.ts \
  packages/sdk/src/raster-image-sizing.ts packages/sdk/src/raster-image-sizing.test.ts \
  packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: add SVG images through the SDK"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 6: PptxGenJS conformance and client-normalized import

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 public SVG data/path output and Tasks 1–5 native APIs.
- Produces: locked final OOXML/function parity plus LibreOffice-normalized SVG import coverage.

- [ ] **Step 1: Add PptxGenJS data/path/sizing conformance tests**

Generate SVG through public `addImage({ data })` and `addImage({ path })`, including name, alt text, transform,
rotation, flips, contain, cover, and crop. Compare native final picture order, transform, direct `srcRect`, fixed
extension URI/namespace, relationship roles, SVG payload, content types, and metadata. Require native fallback to be a
real PNG while recording PptxGenJS path fallback as the one intentional defect divergence.

- [ ] **Step 2: Add PptxGenJS import/edit tests**

Open the generated PptxGenJS file, require all SVG pictures to expose `isSvg`, `fallbackPartUri`, and `svgPartUri`,
replace one pair, write/reopen, and verify neighbors. Exercise arbitrary prefixes without changing unrelated bytes.

- [ ] **Step 3: Add LibreOffice-normalized fixture tests**

Programmatically construct the observed `image/svg` + shared SVG/fallback target pattern for three pictures. Require
all three readable, replace only one pair through clone-on-write, and preserve the other two shared targets.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "test: match PptxGenJS SVG images"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 7: Actual package, browser, gallery, and LibreOffice release evidence

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Temporary only: two clean build manifests, packed consumer, type fixture, source SVG/PNG fixtures, gallery, renders,
  validation JSON, overflow output, round-trip package, and comparison ledger under a new `/tmp/pptx-svg-*` directory.

**Interfaces:**

- Consumes: all public SVG APIs and actual `@jiayunxie/pptx` tarball.
- Produces: Node/browser/types/CLI and multi-client release evidence.

- [ ] **Step 1: Extend permanent actual-package smoke**

In the packed Node consumer, create SVG with explicit and default fallback, inspect both parts, duplicate the slide,
replace one pair, write/reopen, and validate clone-on-write. Add declaration checks for every new type/function and
negative checks for raster low-level SVG misuse and invalid fallback. Include `svgImages: true` in smoke JSON.

- [ ] **Step 2: Extend permanent browser smoke**

From `dist/browser.js`, add SVG by Blob/data URI with a real browser Canvas fallback, write Blob, reopen, and return
`isSvg`, two content types, two internal relationship targets, and fallback PNG signature. Keep the static Node import
search empty and existing download smoke passing.

- [ ] **Step 3: Build twice, hash, pack, install, and run all smoke**

Use a fresh temp directory. Clean-build `@jiayunxie/pptx` twice, hash every dist file in stable path order, require
identical manifests, pack the second build, install that exact tarball without workspace links, and run Node, types,
browser, and installed `pptx-inspect doctor/inspect/validate` against the generated SVG deck.

- [ ] **Step 4: Generate and structurally audit the SVG gallery**

Through the installed tarball create at least four slides covering data/path/bytes/Blob/stream/HTTP URL, explicit and
default fallback, contain/cover/crop, source-rectangle edit, rotation/flips, duplicate + replacement, special-character
name, and non-empty alt text. Require exact slide/shape/picture/part/relationship counts, payload hashes, every fallback
PNG signature, every SVG relationship, strict reopen, and PowerPoint 2010 validation 0/0.

- [ ] **Step 5: Render, inspect, round-trip, and compare**

Render every source page at 180 DPI, run overflow checking, and inspect every page at full size. Save through
LibreOffice, strict reopen/validate/render again, and compare page/shape order, names, alt text, SVG hashes, relationship
roles, transforms, and source rectangles. Record target deduplication, `image/svg+xml` to `image/svg` normalization,
maximum numeric deltas, and any extension rewrite without claiming byte identity.

- [ ] **Step 6: Run release gates, review permanent changes, commit, and push**

```bash
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
pnpm --filter @jiayunxie/pptx build
git diff --check
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed SVG image support"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Confirm only `.pnpm-store/` is untracked and no temp/gallery/dist artifact is staged.

---

### Task 8: Public documentation and compatibility status

**Files:**

- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**

- Consumes: exact API and evidence from Tasks 1–7.
- Produces: current SVG usage, fallback contract, compatibility accounting, and next roadmap item.

- [ ] **Step 1: Document low-level create/read/edit APIs**

Add examples for `slide.addSvgImage(svgBytes, pngBytes, options)`, `isSvg`, `sourcePartUri`/`fallbackPartUri`,
`svgPartUri`, `replaceSvgData()`, duplicate sharing/clone-on-write, and the strict paired-state failure contract.

- [ ] **Step 2: Document high-level sources, sizing, and fallback priority**

Add path/data/bytes/Blob/stream/URL examples, optional `contentType: 'image/svg+xml'`, explicit PNG fallback,
contain/cover/crop, browser Canvas, Node placeholder, and the rule that callers needing old-client fidelity supply PNG.
Keep raster examples and compatibility aliases intact.

- [ ] **Step 3: Publish conformance and client evidence**

Record PptxGenJS 4.0.1 structure, intentional invalid-path-fallback divergence, conformance case count, actual tarball
Node/browser/types/CLI results, deterministic build manifest result, gallery counts, PowerPoint 2010 0/0 validation,
LibreOffice render/save/reopen result, deduplication, MIME normalization, and measured numeric deltas.

- [ ] **Step 4: Update support accounting**

Move SVG create/load/read/replace/sizing from partial/unsupported to supported. Keep external SVG relationships, SVG
DOM editing, script execution, external-resource fetching, arbitrary rasterization fidelity, image rounding,
transparency, hyperlink/shadow/placeholder, and deletion/GC listed accurately for their own future items.

- [ ] **Step 5: Run final gate, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/image-create.internal.test.ts \
  packages/model/src/svg-image-create.internal.test.ts \
  packages/model/src/svg-image-state.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/raster-image-source.test.ts packages/sdk/src/svg-image-source.test.ts \
  packages/sdk/src/raster-image-sizing.test.ts packages/sdk/src/svg-image-fallback.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --check
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md
git diff --cached --check
git commit -m "docs: document SVG image support"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected final state: divergence `0 0`; only `.pnpm-store/` remains untracked.

## Completion Definition

This plan is complete only when all eight tasks are checked, eight isolated implementation/evidence commits are
present on `origin/main`, every release gate passes, public docs mark SVG create/load/read/edit/sizing as supported,
every generated fallback part has a real PNG signature, and original plus LibreOffice-round-tripped galleries have
been fully rendered and visually reviewed without a broken image.
