# Raster Image Source Loading and Inspection Implementation Plan

**Goal:** Add strict PNG/JPEG/GIF signature and intrinsic-size inspection plus a high-level async
`PptxDocument.addImage()` that accepts local paths, URLs, data URIs, bytes, Blob/File, and streams while preserving
the existing atomic `SlideModel.addImage()` core.

**Architecture:** Create one SDK raster-source module containing public source/info types, a pure byte inspector,
and an internal async resolver. Resolve and inspect every source before package mutation, then pass detached bytes
and the detected content type into the existing synchronous model method. Keep Node file access behind dynamic import
and use Fetch for HTTP and browser-relative URLs.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, Node 20+ Fetch/file APIs, browser Blob/Web streams, OPC/model
transactions, PptxGenJS 4.0.1 public output, actual npm tarball Node/browser/declaration/CLI smoke, LibreOffice,
Poppler, and PowerPoint 2010 validation.

## Global constraints

- Implement `docs/superpowers/specs/2026-08-01-raster-image-source-loading-design.md` exactly.
- Preserve `SlideModel.addImage(bytes, options)` as the synchronous strict package-mutation primitive.
- Detect content type only from signature; extensions, `File.name`, `Blob.type`, and HTTP headers are not truth.
- Support exactly PNG, JPEG, and GIF, with positive intrinsic pixel dimensions.
- Keep existing 0/0/1-inch/1-inch transform defaults; do not apply natural-size layout in this stage.
- Finish all source I/O, detachment, signature inspection, and MIME assertions before package mutation.
- Keep the public browser bundle free of static Node imports and `Buffer` dependencies.
- Do not add sizing, crop/srcRect editing, SVG, transparency, rounding, image hyperlink/shadow/placeholder, or image
  metadata editing in this stage.
- Execute all tasks inline in this task.
- Never stage, delete, or modify `.pnpm-store/`.
- End every task with focused tests, self-review, an isolated commit, push to `main`, fetch, and
  `HEAD...origin/main == 0 0`.

---

### Task 1: Pure raster signature and intrinsic-dimension inspector

**Files:**

- Create: `packages/sdk/src/raster-image-source.ts`
- Create: `packages/sdk/src/raster-image-source.test.ts`
- Modify: `packages/sdk/src/index.ts`

**Public output:** `RasterImageInfo` and `inspectRasterImage(bytes)`.

- [ ] **Step 1: Write failing PNG and GIF inspection tests**

Add minimal header fixtures with non-square dimensions. Cover PNG signature, direct first `IHDR`, big-endian
width/height, GIF87a/GIF89a, and little-endian logical-screen dimensions.

```ts
expect(inspectRasterImage(pngHeader(640, 360))).toEqual({
  contentType: 'image/png',
  width: 640,
  height: 360,
});
expect(inspectRasterImage(gifHeader('GIF89a', 320, 200))).toEqual({
  contentType: 'image/gif',
  width: 320,
  height: 200,
});
```

Require repeated calls to remain stable and the input bytes to remain unchanged.

Run:

```bash
pnpm vitest run packages/sdk/src/raster-image-source.test.ts -t "inspects PNG and GIF"
```

Expected: fail because the module does not exist.

- [ ] **Step 2: Write failing JPEG marker traversal tests**

Build synthetic SOI + segment + SOF fixtures for C0-C3, C5-C7, C9-CB, and CD-CF. Cover APP segments, repeated
`FF` fill bytes, TEM and RST standalone markers, and exact big-endian dimensions. Require DHT C4, JPG C8, DAC CC,
SOS, and EOI not to be misread as SOF.

- [ ] **Step 3: Write failing malformed-header tests**

Cover non-`Uint8Array`, empty/unknown signature, truncated PNG signature/IHDR, PNG non-IHDR first chunk, truncated
GIF descriptor, zero PNG/GIF dimensions, truncated JPEG marker/length/payload, segment length below 2, short SOF,
zero JPEG dimensions, SOS/EOI before SOF, and no SOF. Lock deterministic `TypeError` categories without embedding
source payload in messages.

- [ ] **Step 4: Implement the public types and inspector**

In `raster-image-source.ts` add:

```ts
export interface RasterImageInfo {
  readonly contentType: RasterImageContentType;
  readonly width: number;
  readonly height: number;
}

export function inspectRasterImage(bytes: Uint8Array): RasterImageInfo;
```

Use direct byte reads, not `DataView` over an unchecked buffer range. PNG requires the 8-byte signature and direct
13-byte IHDR. GIF requires exactly GIF87a/GIF89a. JPEG walks marker lengths safely and stops at the first supported
SOF. Freeze the returned info object so callers cannot mutate the inspection snapshot.

Export only public inspector types/functions from `packages/sdk/src/index.ts`.

- [ ] **Step 5: Run gates, review, commit, and push**

```bash
pnpm vitest run packages/sdk/src/raster-image-source.test.ts
pnpm typecheck
git diff --check
git add packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts packages/sdk/src/index.ts
git diff --cached --check
git commit -m "feat: inspect raster image dimensions"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected: focused suite and typecheck pass; remote divergence is `0 0`.

---

### Task 2: In-memory, Blob, and stream source resolution

**Files:**

- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`

**Internal output:** detached `{ bytes, info, assertedContentType? }` resolution result and public source types.

- [ ] **Step 1: Add failing public-type and byte-source tests**

Add:

```ts
export type RasterImageByteChunk = number | Uint8Array | ArrayBuffer | ArrayBufferView;
export type RasterImageByteStream =
  | ReadableStream<RasterImageByteChunk>
  | AsyncIterable<RasterImageByteChunk>;
export type RasterImageSource =
  | string
  | Uint8Array
  | ArrayBuffer
  | Blob
  | RasterImageByteStream;
```

Test `Uint8Array` and `ArrayBuffer` resolution for PNG/JPEG/GIF, then mutate the caller source and require resolved
bytes/info to remain unchanged. Reject empty buffers through the inspector.

- [ ] **Step 2: Add failing Blob/File tests**

Cover valid `Blob`, named `File` when available, empty Blob, Blob with wrong/unknown MIME, and Blob arrayBuffer
failure. Require `Blob.type` and `File.name` not to change signature detection. Check abort before and after async
Blob reading.

- [ ] **Step 3: Add failing Web stream and async iterable tests**

Cover byte numbers, `Uint8Array`, `ArrayBuffer`, `DataView`, and typed-array views with non-zero offsets. Assert exact
chunk ordering, copying, reader lock release, iterator closing on failure, error propagation, and abort before/mid-read.
Reject fractions, negative/above-255 numbers, strings, sync iterables, plain objects, and invalid chunks.

- [ ] **Step 4: Implement resolver primitives**

Add an internal exported-for-test helper:

```ts
export interface ResolvedRasterImageSource {
  readonly bytes: Uint8Array;
  readonly info: RasterImageInfo;
  readonly assertedContentType?: RasterImageContentType;
}

export async function resolveRasterImageSource(
  source: RasterImageSource,
  signal?: AbortSignal,
): Promise<ResolvedRasterImageSource>;
```

Do not export `ResolvedRasterImageSource` or the resolver from package root. Implement cross-platform Blob and stream
guards, strict chunk normalization, concatenation, `throwIfAborted()`, and final inspection. Copy every accepted
source/chunk before any await boundary that could expose caller mutation.

- [ ] **Step 5: Export public source types and run gates**

Export `RasterImageByteChunk`, `RasterImageByteStream`, and `RasterImageSource` from SDK root. Run:

```bash
pnpm vitest run packages/sdk/src/raster-image-source.test.ts
pnpm typecheck
git diff --check
git add packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts packages/sdk/src/index.ts
git commit -m "feat: load raster image byte sources"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review that no caller-owned buffer, view, Blob, stream, reader, or iterator is retained.

---

### Task 3: Strict base64 image data URI loading

**Files:**

- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`

- [ ] **Step 1: Add failing valid data URI tests**

Cover PNG/JPEG/GIF, case-insensitive media type/`base64` token, exact payload preservation, detected dimensions, and
declared MIME returned as `assertedContentType`.

- [ ] **Step 2: Add failing strict rejection tests**

Reject missing/empty MIME, unsupported `image/jpg`/SVG, missing `;base64`, extra parameters, percent-encoded raw
payload, empty payload, whitespace, URL-safe alphabet, invalid character/count/padding, non-zero unused padding bits,
trailing comma/data, and declared MIME/signature mismatch. Require every message to omit the payload.

- [ ] **Step 3: Implement a browser-safe canonical base64 decoder**

Validate the exact four-character grammar, `=` position, and unused padding bits. Decode directly to `Uint8Array`
without `Buffer`. Normalize the three accepted MIME values and inspect decoded bytes. Compare the declared MIME with
the detected type before returning.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm vitest run packages/sdk/src/raster-image-source.test.ts -t "data URI"
pnpm vitest run packages/sdk/src/raster-image-source.test.ts
pnpm typecheck
git diff --check
git add packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts
git commit -m "feat: load raster image data URIs"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 4: Node local paths and Fetch URL loading

**Files:**

- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`

- [ ] **Step 1: Add failing Node local-path tests**

Use a temporary directory and valid PNG/JPEG/GIF fixtures. Cover relative and absolute paths, Windows-drive path
classification, missing file, directory/read failure, empty/truncated file, and abort. Require the module to have no
top-level Node import.

- [ ] **Step 2: Add failing Fetch tests**

Use a loopback HTTP server for success, redirect, 404, response with incorrect/missing Content-Type, query-string
URLs, truncated image, delayed response abort, and network failure. Require HTTP/HTTPS strings to download and embed
bytes rather than become external relationships.

Stub browser runtime state and Fetch to cover relative URL, root-relative URL, URL-encoded path, response failure,
and abort. Reject empty string and non-data/non-http explicit schemes. Preserve Windows drive paths as Node files,
not URI schemes.

- [ ] **Step 3: Implement deterministic string dispatch**

Dispatch in order: strict data URI, absolute HTTP/HTTPS, unsupported explicit scheme, then Node local path or browser
Fetch. Dynamically load `node:fs/promises` only inside the local-path branch. Use global Fetch with `signal`, require
`response.ok`, copy `arrayBuffer()`, and inspect bytes. Ignore file extension, Blob/File name, and response MIME for
format truth.

- [ ] **Step 4: Run Node/browser gates, review, commit, and push**

```bash
pnpm vitest run packages/sdk/src/raster-image-source.test.ts
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
rg -n 'from "node:|from '\''node:|require\("node:' packages/pptx/dist/browser.js
git diff --check
git add packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts
git commit -m "feat: load raster images from paths and URLs"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

The browser static-Node-import search must return no matches.

---

### Task 5: Atomic `PptxDocument.addImage()` integration

**Files:**

- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Public output:** `AddImageSourceOptions` and async `PptxDocument.addImage()`.

- [ ] **Step 1: Add failing API/type tests**

Add:

```ts
export interface AddImageSourceOptions extends Omit<AddImageOptions, 'contentType'> {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
}
```

Compile examples for every source kind, optional content type, transform/name/alt text, and returned
`Promise<ImageModel>`. Add `@ts-expect-error` cases for unsupported MIME/source and SlideModel's still-required
content type.

- [ ] **Step 2: Add failing successful-integration tests**

On blank and existing decks, load all source forms and require the returned object to be the exact live image in
`slide.shapes`. Check default and explicit transforms, names, alt text, part extensions/content types, payload bytes,
internal image relationships, shape order, write/reopen, and all six presentation formats.

- [ ] **Step 3: Add failing pre-I/O and zero-mutation tests**

Require an invalid slide index plus structurally invalid or source-only options to fail before reading a getter,
opening a file, starting Fetch, reading Blob, or consuming a stream. Cover null/array/class/inherited/accessor/symbol/
unknown options, invalid signal, and invalid content type. Cover transform/name/alt-text semantic errors plus every
source/parser/assertion failure as zero-package-mutation cases; semantic image fields continue to use the model core's
single validation contract and may be rejected after source resolution. Snapshot parts, relationships, slide XML,
shapes, graph, and mutation journal and require exact equality after each rejection.

- [ ] **Step 4: Implement descriptor-safe option splitting**

Normalize only ordinary/null-prototype objects with own data properties from the closed key set. Copy values without
invoking accessors. Validate `contentType`, `signal`, and source-only fields before I/O; copy image fields into a fresh
model options object. Let the existing model normalizer remain the only transform/name/alt-text semantic validator when
the detached bytes reach `slide.addImage()`. Do not duplicate those semantic rules in SDK code.

- [ ] **Step 5: Implement the document method**

```ts
async addImage(
  slideIndex: number,
  source: RasterImageSource,
  options: AddImageSourceOptions = {},
): Promise<ImageModel> {
  const slide = this.slides[slideIndex];
  if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
  const normalized = normalizeAddImageSourceOptions(options);
  const resolved = await resolveRasterImageSource(source, normalized.signal);
  assertRasterImageContentType(normalized.contentType, resolved);
  return slide.addImage(resolved.bytes, {
    ...normalized.imageOptions,
    contentType: resolved.info.contentType,
  });
}
```

Ensure options/data URI MIME assertions are both checked. Do not use intrinsic dimensions to alter transform.

- [ ] **Step 6: Run integration/full gates, review, commit, and push**

```bash
pnpm vitest run packages/sdk/src/raster-image-source.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/sdk/src/raster-image-source.ts packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git commit -m "feat: add images from portable sources"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 6: PptxGenJS conformance, packed runtime, gallery, and release evidence

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`
- Temporary only: source fixtures, local HTTP server, packed-runtime scripts, gallery, renders, and comparison output
  under a new `/tmp/pptx-raster-source-*` directory

- [ ] **Step 1: Add PptxGenJS 4.0.1 public conformance**

Generate PNG/JPEG/GIF pictures through PptxGenJS public `path` and `data`, and through native local path, URL, and
data URI. Compare picture kind/order, payload, content type, internal relationship, name/alt text, and transform after
normalizing public inch/degree inputs to native EMU/angles. Record strict native divergences for malformed data URI,
signature/MIME mismatch, and unknown image data.

- [ ] **Step 2: Build and test the actual npm tarball**

Clean/rebuild twice and compare `packages/pptx/dist` hashes. Pack the public package, install the emitted tarball into
a temporary consumer, and verify:

- Node ESM path/data/bytes/Blob/stream input and write/reopen;
- browser bundle import with no static Node modules and mocked relative Fetch;
- declarations for all source/info/options types and `Promise<ImageModel>`;
- existing CLI inspect/validate behavior;
- no runtime dependency on PptxGenJS or workspace-only paths.

- [ ] **Step 3: Generate and inspect a multi-source gallery**

Through the installed tarball, create a wide deck containing at least PNG/JPEG/GIF from bytes, Blob, data URI, local
path, URL, Web stream, and async iterable. Include default and explicit transforms, special-character metadata, and a
source-dimension legend. Require strict reopen, expected shape/part/relationship counts, exact payload SHA-256 values,
and PowerPoint 2010 validation with zero errors/warnings.

Render every slide at 180 DPI, run the presentation overflow checker, and visually inspect at original resolution.
Use `pptx-inspect` package/slide summaries as stable evidence.

- [ ] **Step 4: LibreOffice round-trip audit**

Open/save the gallery through LibreOffice headless, strictly reopen and validate the result, render every page, and
compare page count/size, picture order/name/content type, relationship targets, and payload SHA-256 values. Record every
direct-state rewrite explicitly; do not claim byte identity when the client normalizes transforms or picture markup.

- [ ] **Step 5: Update public documentation and progress accounting**

Document both layers:

- synchronous `SlideModel.addImage(bytes, { contentType, ... })` for strict atomic core use;
- async `PptxDocument.addImage(slideIndex, source, options?)` for portable source loading;
- `inspectRasterImage()` and its raw pixel-dimension semantics;
- exact source matrix, string dispatch, signature-first MIME policy, abort behavior, default sizing, strict errors,
  PptxGenJS parity boundary, and client evidence.

Update the image/SVG roadmap: source loading and detection complete; contain/cover/crop becomes the next stage.

- [ ] **Step 6: Run final release gate**

```bash
pnpm vitest run packages/sdk/src/raster-image-source.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --check
```

Also require clean-build dist reproducibility, actual-tarball Node/browser/types/CLI checks, strict gallery validation,
LibreOffice reopen/render, and overflow checks to pass. Review all changed code/tests/docs, remove temporary repository
artifacts, then:

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "docs: document raster image source loading"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

## Completion definition

This plan is complete only when every checkbox is satisfied, six isolated commits are present remotely, all release
gates pass, and documentation no longer lists raster path/data/source loading or PNG/JPEG/GIF detection as partial or
unsupported. The next independent stage is contain/cover/crop sizing and direct `srcRect` editing.
