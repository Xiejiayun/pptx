# Embedded Media Creation Conformance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make strict embedded audio/video creation through `PptxDocument.addAudio()` and `addVideo()` accept every documented source, emit canonical OOXML, and match PptxGenJS 4.0.1 public media output where its output is valid.

**Architecture:** Keep public media types and package mutation in `MediaCodec`, but split descriptor-safe request normalization, asynchronous source/MIME/extension resolution, and pure final-definition/XML rendering into focused internal modules. Resolve and detach all inputs before one synchronous `OpcPackage.transaction()`, then verify the resulting public SDK through PptxGenJS conformance, actual packed Node/browser/type/CLI consumers, and multi-client package evidence.

**Tech Stack:** TypeScript 5.8, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, JSZip, PptxGenJS 4.0.1, tsup, Playwright browser smoke, `pptx-inspect`, ffmpeg, LibreOffice/Poppler.

## Global Constraints

- Keep `PptxDocument.addAudio(slideIndex, source, options)` and `addVideo(slideIndex, source, options)` as the canonical high-level APIs.
- Preserve path, HTTP/HTTPS external URL, `Uint8Array`, `ArrayBuffer`, Blob/File, Web Stream, and async-iterable sources; add strict base64 data URIs.
- `AddMediaOptions` adds only `name` and `altText`; all existing fields and the current transcoder signature remain source-compatible.
- Supported audio is `audio/mpeg` `.mp3`, `audio/mp4` `.m4a`, `audio/wav` `.wav`, and `audio/ogg` `.ogg`.
- Supported video is `video/mp4` `.mp4|.m4v`, `video/quicktime` `.mov`, and `video/webm` `.webm`.
- Supported posters are `image/png` `.png`, `image/jpeg` `.jpg|.jpeg`, and `image/gif` `.gif`; an omitted poster uses the built-in PNG and an HTTP/HTTPS poster is rejected.
- Native audio keeps canonical `a:audioFile` and `audio/mpeg`; do not copy PptxGenJS 4.0.1's `a:videoFile` and `audio/mp3` defects.
- Keep the private playback extension byte-compatible; native timing-tree playback remains a separate slice.
- All request validation, source I/O, transcode work, poster resolution, MIME/extension decisions, hashing, and XML definition work finish before package mutation.
- Failed calls leave parts, content types, relationships, graph, slide XML, ZIP output, shape ids, and the mutation journal unchanged.
- Deduplication remains SHA-256 plus exact content type; same bytes with a different content type never share a target.
- Every task ends with focused tests, typecheck, diff review, an isolated commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage or commit `.pnpm-store/`, probe decks, generated media, packed consumers, gallery files, renders, or validation output.

---

### Task 1: Descriptor-safe request normalization

**Files:**

- Modify: `packages/codecs/src/media.ts`
- Create: `packages/codecs/src/media-create.internal.ts`
- Create: `packages/codecs/src/media-create.internal.test.ts`

**Interfaces:**

- Consumes: current `MediaKind`, `MediaSource`, `MediaPlaybackSettings`, and `AddMediaOptions`.
- Produces: `AddMediaOptions.name`, `AddMediaOptions.altText`, `NormalizedMediaSourceReference`, `NormalizedMediaCreateRequest`, and `normalizeMediaCreateRequest()`.

- [ ] **Step 1: Write failing public type and default tests**

Add compile/runtime coverage equivalent to:

```ts
const options: AddMediaOptions = {
  name: 'Audio & narration',
  altText: 'Spoken overview',
  contentType: 'audio/mpeg',
  fileName: 'overview.mp3',
  poster: new Uint8Array([1, 2, 3]),
  posterContentType: 'image/png',
  x: 0,
  y: -1,
  width: 914_400,
  height: 914_400,
  play: 'click',
  loop: false,
  hideWhenStopped: false,
  volume: 1,
};
const request = normalizeMediaCreateRequest(
  'audio',
  new Uint8Array([4, 5, 6]),
  options,
);
```

Require default x/y = 914400, default audio size = 914400×914400, default video size = 4572000×2571750, click/false/false/1 playback defaults, omitted name/alt text preserved as `undefined`, frozen plain containers, and copied byte/ArrayBuffer inputs.

- [ ] **Step 2: Write failing descriptor and scalar validation tests**

Reject null, arrays, primitives, class instances, custom prototypes, symbol keys, accessors, and unknown keys without invoking a getter or consuming a stream. Reject invalid kind/source, XML-invalid name/alt text, empty file name, non-string/empty MIME assertions, non-function transcode, non-safe x/y, non-positive width/height, wrong play/booleans, and non-finite or out-of-range volume. Require explicit empty name and alt text to survive unchanged.

- [ ] **Step 3: Add the public creation fields and normalized request types**

Extend only the public options:

```ts
export interface AddMediaOptions extends MediaPlaybackSettings {
  readonly name?: string;
  readonly altText?: string;
  readonly contentType?: string;
  readonly fileName?: string;
  readonly poster?: MediaSource;
  readonly posterContentType?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly transcode?: (
    bytes: Uint8Array,
    contentType: string,
    kind: MediaKind,
  ) => Promise<{ bytes: Uint8Array; contentType: string; extension?: string }>;
}
```

Define the internal detached surface:

```ts
export type NormalizedMediaSourceReference =
  | Readonly<{ type: 'string'; value: string }>
  | Readonly<{ type: 'bytes'; bytes: Uint8Array }>
  | Readonly<{ type: 'blob'; value: Blob }>
  | Readonly<{ type: 'stream'; value: MediaByteStream }>;

export interface NormalizedMediaCreateRequest {
  readonly kind: MediaKind;
  readonly source: NormalizedMediaSourceReference;
  readonly poster?: NormalizedMediaSourceReference;
  readonly name?: string;
  readonly altText?: string;
  readonly contentType?: string;
  readonly posterContentType?: string;
  readonly fileName?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
  readonly transcode?: NonNullable<AddMediaOptions['transcode']>;
}
```

- [ ] **Step 4: Implement closed-key normalization before I/O**

Use `Reflect.ownKeys()` plus `Object.getOwnPropertyDescriptor()` to read only own data properties. Accept only `Object.prototype` or null prototypes. Validate XML strings by Unicode code point, preserve negative zero as zero, copy direct byte buffers immediately, treat Blob as immutable, and retain a stream only as the one-shot source to be consumed by Task 2. Freeze every plain container, but do not call `Object.freeze()` on non-empty typed arrays.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-create.internal.test.ts
pnpm typecheck
git diff --check
git add packages/codecs/src/media.ts packages/codecs/src/media-create.internal.ts \
  packages/codecs/src/media-create.internal.test.ts
git diff --cached --check
git commit -m "feat: normalize embedded media requests"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review that no caller options object or caller-owned byte buffer is retained and unsafe getters were never executed.

---

### Task 2: Strict media, poster, MIME, and extension resolution

**Files:**

- Create: `packages/codecs/src/media-source.internal.ts`
- Create: `packages/codecs/src/media-source.internal.test.ts`
- Modify: `packages/codecs/src/media-create.internal.ts`
- Modify: `packages/codecs/src/media-create.internal.test.ts`

**Interfaces:**

- Consumes: Task 1 `NormalizedMediaCreateRequest` and normalized source references.
- Produces: `ResolvedEmbeddedMedia`, `ResolvedExternalMedia`, `ResolvedMediaCreationInputs`, `resolveMediaCreationInputs()`, and strict shared base64 parsing.

- [ ] **Step 1: Write failing valid source-resolution tests**

Cover primary and poster data URIs, direct bytes, ArrayBuffer, Blob/File name inference, Web Stream, async iterable, local path, default built-in poster, and HTTP/HTTPS external media. Use canonical examples such as:

```ts
const request = normalizeMediaCreateRequest(
  'audio',
  'data:audio/mpeg;base64,AQIDBA==',
  {
    contentType: 'audio/mpeg',
    fileName: 'voice.mp3',
    poster: 'data:image/png;base64,iVBORw0KGgo=',
    posterContentType: 'image/png',
  },
);
const resolved = await resolveMediaCreationInputs(request);
```

Require non-empty copied bytes, exact asserted content types, lowercase dotted extensions, `.m4v` preservation for `video/mp4`, `.jpeg` preservation for JPEG posters, canonical defaults for unknown/no extension, and external media with no allocated part data.

- [ ] **Step 2: Write failing strict data-URI and source tests**

Reject missing/extra comma, missing `;base64`, empty payload, whitespace, percent encoding, URL-safe alphabet, wrong padding length, non-zero padding bits, trailing data, unsupported MIME, media kind mismatch, media/poster assertion mismatch, non-HTTP URI schemes, empty local path, empty resolved payload, unsupported stream chunks, and HTTP/HTTPS poster URLs. Require source errors to occur with zero package access.

- [ ] **Step 3: Implement canonical MIME/extension tables and base64 decoder**

Keep the tables private to `media-source.internal.ts` and expose only resolved types. Decode base64 manually with the same canonical length, alphabet, padding, and padding-bit checks already proven by the SDK image loader; do not introduce `Buffer`, `atob`, or a static Node import into the browser path. Treat a recognized extension that maps to another MIME as a mismatch; ignore an unknown path/file extension and choose the final MIME's canonical extension.

- [ ] **Step 4: Implement portable source loading and deterministic precedence**

Preflight both normalized sources, exact supported lowercase MIME assertions, recognized extension compatibility, data-URI headers, poster URL policy, and external/transcode incompatibility before reading a path, Blob, or stream. Then resolve in this order: explicit MIME assertion, data-URI declaration, safe `fileName`/File/path extension, then `audio/mpeg` or `video/mp4`. Read a native File name through the platform `File.prototype.name` getter so a subclass override is not invoked. Use dynamic Node `fs/promises` import only for local paths. Do not fetch HTTP/HTTPS media. For poster omission return a new copy of the built-in PNG; reject an external poster instead of replacing it.

- [ ] **Step 5: Write failing transcode contract tests**

Require an embedded transcoder to receive detached bytes plus final pre-transcode MIME and kind. Accept a promise result with non-empty `Uint8Array`, supported compatible MIME, and optional bare or dotted supported extension. Copy output immediately after await. Reject external-media transcoding, thrown/rejected transcoders, null/array/class/accessor result objects, empty/non-byte results, wrong-kind MIME, MIME/extension mismatch, and caller mutation after resolution.

- [ ] **Step 6: Implement transcode resolution and final detached inputs**

Return this discriminated surface:

```ts
export interface ResolvedEmbeddedMedia {
  readonly type: 'embedded';
  readonly bytes: Uint8Array;
  readonly contentType: string;
  readonly extension: string;
}

export interface ResolvedExternalMedia {
  readonly type: 'external';
  readonly url: string;
}

export interface ResolvedMediaCreationInputs {
  readonly media: ResolvedEmbeddedMedia | ResolvedExternalMedia;
  readonly poster: ResolvedEmbeddedMedia;
}

export function resolveMediaCreationInputs(
  request: Readonly<NormalizedMediaCreateRequest>,
): Promise<Readonly<ResolvedMediaCreationInputs>>;
```

Freeze the containers, keep their copied typed arrays private, and never retain the transcoder result object.

- [ ] **Step 7: Run Node/browser gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-create.internal.test.ts \
  packages/codecs/src/media-source.internal.test.ts
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
rg -n 'from "node:|from '\''node:|require\("node:' packages/pptx/dist/browser.js
git diff --check
git add packages/codecs/src/media-create.internal.ts \
  packages/codecs/src/media-create.internal.test.ts \
  packages/codecs/src/media-source.internal.ts \
  packages/codecs/src/media-source.internal.test.ts
git diff --cached --check
git commit -m "feat: resolve strict embedded media sources"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

The static Node import search must return no matches.

---

### Task 3: Canonical final definition and media picture XML

**Files:**

- Modify: `packages/codecs/src/media-create.internal.ts`
- Modify: `packages/codecs/src/media-create.internal.test.ts`

**Interfaces:**

- Consumes: normalized request plus Task 2 `ResolvedMediaCreationInputs`.
- Produces: `NormalizedMediaCreationDefinition`, `finalizeMediaCreationDefinition()`, and `renderMediaPictureXml()`.

- [ ] **Step 1: Write failing final-definition detachment tests**

Finalize audio/video definitions using an explicit name/alt text and default `Media N`. Require copied media/poster bytes, frozen containers, exact transforms/settings, external-vs-embedded state, and `undefined` alt text remaining absent. Mutate every caller/resolver object after finalization and require the definition to stay unchanged.

- [ ] **Step 2: Write failing exact embedded audio XML tests**

Require this ownership and order with escaped ids/name/description:

```xml
<p:pic><p:nvPicPr><p:cNvPr id="2" name="Audio &amp; narration" descr="Spoken &quot;overview&quot;"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="rId2"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId3"/></p:ext><p:ext uri="{C13D3E4A-5148-4B6D-A7E7-505054582D4F}"><px:playback xmlns:px="urn:pptx-ooxml:media" play="click" loop="0" hideWhenStopped="0" volume="100000"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>
```

Then require the poster blip, stretch fill, normalized transform, rectangular geometry, and closing picture elements in exact order.

- [ ] **Step 3: Write failing video and external XML tests**

Require video to use `a:videoFile`; embedded video includes the Microsoft relationship extension, while an external HTTP/HTTPS media definition omits `p14:media` but retains the kind relationship, poster, click action, private playback extension, aspect lock, transform, and escaped metadata. Require omitted alt text to omit `descr` and explicit empty alt text to emit `descr=""`.

- [ ] **Step 4: Implement pure definition finalization and rendering**

Use `escapeXmlAttribute()` for every caller/relationship string. `finalizeMediaCreationDefinition()` accepts the already-computed default name so it does not inspect a package. Define the complete final surface:

```ts
export interface NormalizedMediaCreationDefinition {
  readonly kind: MediaKind;
  readonly name: string;
  readonly altText?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
  readonly media: ResolvedEmbeddedMedia | ResolvedExternalMedia;
  readonly poster: ResolvedEmbeddedMedia;
}

export function finalizeMediaCreationDefinition(
  request: Readonly<NormalizedMediaCreateRequest>,
  resolved: Readonly<ResolvedMediaCreationInputs>,
  defaultName: string,
): Readonly<NormalizedMediaCreationDefinition>;
```

`renderMediaPictureXml()` accepts:

```ts
export interface MediaRelationshipIds {
  readonly kind: string;
  readonly media?: string;
  readonly poster: string;
}

export function renderMediaPictureXml(
  shapeId: number,
  definition: Readonly<NormalizedMediaCreationDefinition>,
  relationships: Readonly<MediaRelationshipIds>,
): string;
```

Preserve the current playback extension URI, attributes, and rounding exactly; volume was already validated and is rendered as `Math.round(volume * 100000)`.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-create.internal.test.ts \
  packages/codecs/src/media-source.internal.test.ts
pnpm typecheck
git diff --check
git add packages/codecs/src/media-create.internal.ts \
  packages/codecs/src/media-create.internal.test.ts
git diff --cached --check
git commit -m "feat: render canonical embedded media pictures"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review exact child order, `a:audioFile`/`a:videoFile`, `descr` presence semantics, and absence of unescaped interpolation.

---

### Task 4: Atomic `MediaCodec` creation and lifecycle regression

**Files:**

- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/codecs.test.ts`

**Interfaces:**

- Consumes: Tasks 1–3 request normalization, source resolution, final definition, and renderer.
- Produces: strict `MediaCodec.addAudio()` / `addVideo()` package mutation with existing `list()`, `delete()`, `diagnostics()`, and dedup behavior preserved.

- [ ] **Step 1: Write failing canonical package tests**

Create embedded audio and video after an existing media picture. Require default names `Media 0` / `Media 1`, explicit name/alt text, safe shape ids/order, exact media/poster part extensions and content types, kind plus Microsoft media plus poster relationships, click action, fixed Office media extension, aspect lock, transform, copied payloads, and `MediaCodec.list()` returning the correct kind/part/settings.

- [ ] **Step 2: Write failing deduplication and deletion regression tests**

Require same bytes plus same content type to share media/poster targets, same bytes plus different content type to allocate distinct targets, `.mp4` and `.m4v` MIME-compatible extension choices not to defeat content-hash reuse, deletion of one shared picture to preserve targets, deletion of the final reference to collect media/poster targets, and external media deletion to remove only its relationships/poster when unreferenced.

- [ ] **Step 3: Write failing full rollback matrix**

Before each operation snapshot part URIs/content types/bytes, relationship records, graph, slide XML, `pkg.write()`, and journal. Require exact equality after malformed options, malformed media/poster data URI, empty payload, MIME/extension mismatch, source read failure, stream failure, transcode failure, poster failure, missing shape tree, invalid XML append, part-allocation failure, relationship target failure, and an outer transaction rollback.

- [ ] **Step 4: Refactor `MediaCodec.add()` around one resolved definition**

Run `normalizeMediaCreateRequest()`, `resolveMediaCreationInputs()`, SHA-256 lookup, default-name counting, and `finalizeMediaCreationDefinition()` before mutation. Inside one synchronous transaction allocate any missing part URIs, insert missing parts, add relationships, allocate the shape id, append the rendered picture before a shape-tree `p:extLst` when present, update the slide, and return the existing snapshot `MediaModel`.

Count default names only from direct slide `p:pic` objects whose `p:nvPr` contains `a:audioFile` or `a:videoFile`. Do not use shape id or count unrelated pictures/shapes.

- [ ] **Step 5: Remove superseded permissive helpers**

Delete the old `resolveSource()`, `resolvePoster()`, MIME fallback maps, `clamp()`, and duplicated byte-stream logic from `media.ts`. Retain hashing and relationship helpers only when they are still owned by the codec. Keep `list()`, diagnostics, and deletion behavior byte-compatible except for canonical markup now emitted by creation.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-create.internal.test.ts \
  packages/codecs/src/media-source.internal.test.ts packages/codecs/src/codecs.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/codecs/src/media.ts packages/codecs/src/codecs.test.ts
git diff --cached --check
git commit -m "feat: create strict embedded audio and video"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 5: SDK zero-input, source matrix, and six-format evidence

**Files:**

- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: Task 4 through existing `PptxDocument.addAudio()`, `addVideo()`, `media()`, `write()`, and `open()`.
- Produces: public zero-input and round-trip evidence for every supported source/MIME/format without a new SDK facade.

- [ ] **Step 1: Add public type-surface tests**

Compile valid `name`, `altText`, poster, playback, transform, and transcoder calls from `@pptx/sdk`. Add `@ts-expect-error` coverage for wrong name/alt types, invalid play literal, non-boolean flags, and wrong transcoder result bytes. Verify `@pptx/sdk` continues to re-export `MediaKind`, `MediaSource`, `MediaPlaybackSettings`, `AddMediaOptions`, and `MediaModel` through its imported public method signatures.

- [ ] **Step 2: Add every source and supported MIME test**

From a newly created presentation, add audio/video using local path, strict data URI, `Uint8Array`, ArrayBuffer, Blob/File, Web Stream, async iterable, and external HTTP/HTTPS URL. Cover all four audio and all three video MIME families plus `.m4v`, three poster formats, special-character names, explicit/omitted/empty alt text, and one valid transcoder. Require immediate `document.media(0)` state and exact part/relationship metadata.

- [ ] **Step 3: Add asynchronous detachment and zero-mutation tests**

Pause stream and transcoder promises, mutate caller byte buffers/options/result objects, then resume and require original normalized state. For every invalid source/options combination, compare package parts, relationships, slide XML, graph, ZIP output, and journal before/after. Include invalid slide index so no source is consumed.

- [ ] **Step 4: Add write/reopen and six-format tests**

For every key of `PRESENTATION_FORMAT_PROFILES`, create a blank document, add MP3 data-URI audio with PNG poster and MP4 byte video with JPEG poster, write, reopen twice, and require the original format, picture order/names/alt text/transforms, media/poster bytes/content types/extensions, relationship roles, Office media extension, and validator error list `[]`.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-create.internal.test.ts \
  packages/codecs/src/media-source.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "test: verify public embedded media creation"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 6: PptxGenJS public media conformance

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 public `slide.addMedia()` and native Tasks 1–5.
- Produces: locked comparison cases and explicit defect divergences; no adapter runtime behavior changes.

- [ ] **Step 1: Extend the local public PptxGenJS slide type**

Derive the real public method argument from `pptxgenjs` declarations and add it to the existing local test interface:

```ts
type PptxGenJSMediaOptions = Parameters<
  ReturnType<PptxGenJSPublicInstance['addSlide']>['addMedia']
>[0];

interface PptxGenJSSlide {
  addMedia(options: PptxGenJSMediaOptions): void;
}
```

Keep generation entirely through `addMedia()` plus public `writeFile()`/`write()`; do not call PptxGenJS internals.

- [ ] **Step 2: Add valid public output comparisons**

Generate and open these PptxGenJS cases: audio data URI + cover + extn + objectName + transform; video data URI + cover + extn + objectName + transform; local-path audio with inferred extension; and duplicate payloads. Generate native counterparts. Compare picture order/name/transform, click action, poster/media bytes, part extensions, relationship roles/modes/targets, Office media extension, strict reopen, and package validation.

- [ ] **Step 3: Lock intentional divergence assertions**

Assert that PptxGenJS audio emits `a:videoFile` while native emits `a:audioFile`, and that PptxGenJS MP3 uses `audio/mp3` while native uses `audio/mpeg`. Require native strict rejection for missing/empty source, malformed media/cover data URI, wrong kind, MIME/extension mismatch, and invalid runtime options; record PptxGenJS logging/coercion only as comparison evidence, never as behavior to copy.

- [ ] **Step 4: Add native reopen-and-continue coverage**

Open valid PptxGenJS audio/video output with `PptxDocument`, confirm `media()` discovers both objects, add one strict native object after them, and write/reopen. Require existing PptxGenJS parts/pictures to remain unchanged and the new object to use canonical native relationships/content types.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  packages/codecs/src/codecs.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: match PptxGenJS embedded media output"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 7: Actual package, browser, CLI, gallery, and client evidence

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Temporary only: clean-build manifests, packed consumer, declaration fixture, generated WAV/MP3/MP4/WebM/poster fixtures, gallery decks, renders, validation JSON, round-trip package, and comparison ledger under a new `/tmp/pptx-media-*` directory.

**Interfaces:**

- Consumes: the complete public embedded-media slice and actual `@jiayunxie/pptx` tarball.
- Produces: Node/browser/declaration/CLI and PowerPoint-validator/LibreOffice release evidence.

- [ ] **Step 1: Extend permanent packed Node and declaration smoke**

In the installed-tarball consumer create audio from data URI/bytes/path/Blob/stream and video from data URI/bytes, with explicit PNG/JPEG posters, names, alt text, transforms, playback settings, deduplication, delete-one/shared-target preservation, write/reopen, and relationship inspection. Add compile-valid declarations for every source/options field and negative declarations for invalid media kinds, names, poster types, playback, and transcode results. Include `embeddedMedia: true` in smoke JSON.

- [ ] **Step 2: Extend permanent browser smoke**

From `dist/browser.js`, add MP3 data-URI audio plus Blob video and a Web Stream poster, write Blob, reopen, and return names, content types, extensions, three relationship roles per embedded item, `a:audioFile`/`a:videoFile` counts, poster byte signatures, and validation errors. Keep the static Node import search empty and the download smoke passing.

- [ ] **Step 3: Build twice, hash, pack, install, and run all smoke**

Use a fresh temp directory. Clean-build `@jiayunxie/pptx` twice, hash every dist file in stable path order, require identical manifests, pack the second build, install that exact tarball without workspace links, and run Node, declarations, browser, plus installed `pptx-inspect doctor`, `package inspect`, and `package validate --profile powerpoint-2010` against the generated media deck.

- [ ] **Step 4: Generate playable fixtures and a media gallery**

Use ffmpeg's lavfi sources to generate a one-second silent WAV/MP3/OGG/M4A and color-frame MP4/MOV/WebM under the temp directory. Through the installed tarball create at least five slides covering all supported MIME/extensions, every portable memory source, Node path, explicit/default poster, special-character names, alt text, transforms, duplicate deduplication, delete-one preservation, and external URL diagnostics on a separate non-zero-warning deck.

- [ ] **Step 5: Structurally audit and render the gallery**

Require exact slide/picture/media/poster/relationship counts, payload hashes, content types/extensions, no orphan media targets, correct audio/video element kind, strict reopen, and PowerPoint 2010 validation 0 errors/0 warnings for the embedded gallery. Render every source slide at 180 DPI, run overflow checking, and inspect every slide at full size for poster placement, clipping, and unintended blank output.

- [ ] **Step 6: LibreOffice save/reopen comparison**

Save the embedded gallery through LibreOffice, strict reopen/validate/render it again, and compare slide/picture order, names, non-empty alt text, media/poster hashes, relationship roles, transforms, extensions, and content types. Record any client deduplication, MIME normalization, maximum numeric deltas, removed playback extension, or poster rewrite as observed normalization rather than byte equality.

- [ ] **Step 7: Run release gates, review permanent changes, commit, and push**

```bash
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
pnpm --filter @jiayunxie/pptx build
git diff --check
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed embedded media support"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Review that only the two permanent smoke files are staged and no temp, generated-media, gallery, render, tarball, dist, or `.pnpm-store/` artifact is included.

---

### Task 8: Public documentation and compatibility accounting

**Files:**

- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**

- Consumes: exact APIs, conformance cases, and release evidence from Tasks 1–7.
- Produces: accurate media creation usage, compatibility status, intentional divergences, and the next roadmap slice.

- [ ] **Step 1: Document strict high-level creation**

Add audio/video examples for strict data URI, path, bytes, Blob/stream, explicit poster, `name`, `altText`, `fileName`, `contentType`, transform, and playback settings. Explain that HTTP/HTTPS media remains external and is never fetched, poster URLs are rejected, and callers use `inches()` for EMU placement.

- [ ] **Step 2: Document supported MIME, extension, and transcode rules**

Publish the exact four audio, three video, and three poster MIME families; assertion/inference priority; data-URI grammar; extension mismatch behavior; default poster; hash+content-type deduplication; async detachment; and atomic rollback contract. State that playback preferences remain in the private extension until the timing slice.

- [ ] **Step 3: Publish conformance and client evidence**

Record PptxGenJS 4.0.1 public valid-case count, canonical native `a:audioFile`/`audio/mpeg` divergences, actual tarball Node/browser/declaration/CLI results, deterministic build manifest result, embedded gallery counts, PowerPoint 2010 validation, and LibreOffice render/save/reopen observations with measured normalization values.

- [ ] **Step 4: Update support accounting and remaining roadmap**

Move data URI, cover/poster, extension mapping, object name, alt text creation, strict embedded audio/video creation, and public conformance into supported. Keep online video, remote-fetch embedding, stable live media identity/editing, complete duplicate/move/delete isolation, native timing-tree playback, captions/subtitles, crop/rounding/shadow/hyperlink/placeholder styles, and broad client certification explicitly pending. Set the next media item to the stable live media lifecycle.

- [ ] **Step 5: Run final gate, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-create.internal.test.ts \
  packages/codecs/src/media-source.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
pnpm --filter @jiayunxie/pptx build
git diff --check
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md
git diff --cached --check
git commit -m "docs: document embedded media creation"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected final state: divergence `0 0`; only `.pnpm-store/` remains untracked.

## Completion Definition

This plan is complete only when all eight tasks are checked, eight isolated implementation/evidence commits are present on `origin/main`, every release gate passes, public docs accurately mark strict embedded audio/video creation as supported, the actual packed package passes Node/browser/declaration/CLI smoke, and the source plus LibreOffice-round-tripped playable media galleries have been structurally and visually reviewed.
