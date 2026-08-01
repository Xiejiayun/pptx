# Slide Background Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add complete direct slide-background create/read/edit/clear support for no-fill, solid, gradient, and raster image backgrounds with PptxGenJS 4.0.1 parity and safe media lifecycle behavior.

**Architecture:** Keep a synchronous, model-owned strict background module responsible for normalization, direct OOXML projection, local patches, image relationships, clone-on-write, and garbage collection. Reuse the existing gradient codec and simple-fill helpers, while the asynchronous SDK reuses the raster source loader before assigning the same model property.

**Tech Stack:** TypeScript 5.8, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, JSZip, PptxGenJS 4.0.1, tsup, Playwright browser smoke, `pptx-inspect`, LibreOffice/Poppler.

## Global Constraints

- Execute inline in the current task; do not delegate implementation or review.
- Support only direct slide `p:cSld/p:bg/p:bgPr`; layout/master background editing, pattern/group fill, `p:bgRef` semantic editing, image crop/tile/effects, and default color remain outside this slice.
- Public background values are `undefined`, `none`, `solid`, existing linear/path gradient, or PNG/JPEG/GIF image bytes.
- `undefined` removes direct `p:bg` and restores inheritance; `{ kind: 'none' }` writes legal `a:noFill`.
- Never reproduce PptxGenJS 4.0.1's empty-`p:bgPr` defect for `{ color, type: 'none' }`.
- Model input is descriptor-safe, closed, detached, and validated before mutation; SDK source input also receives signature and asserted-MIME validation.
- Reads are strict and non-mutating; unsupported or ambiguous direct backgrounds project to `undefined` without repair.
- Equal supported values are exact no-ops with unchanged XML, relationships, parts, target URIs, and mutation journal.
- Image replacement isolates duplicated/shared targets and shared relationship ids; targets are deleted only after graph incoming reaches zero.
- Apply identical behavior to `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`.
- Production packages do not import or call PptxGenJS; the locked 4.0.1 package remains adapter/test evidence only.
- Use fresh `/tmp/pptx-slide-background-*` directories for tarballs, fixtures, galleries, renders, validation output, and client round trips.
- Every task ends with focused tests, typecheck where applicable, diff review, one isolated commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage or commit `.pnpm-store/`, generated decks, renders, packed tarballs, or temporary client artifacts.

---

## File Map

- Create `packages/model/src/slide-background.ts`: public `SimpleFill`, `SlideBackgroundImage`, and `SlideBackground` types.
- Create `packages/model/src/slide-background.internal.ts`: strict normalization, direct reader, equality, renderer, mutation, image relationship isolation, and media target collection.
- Create `packages/model/src/slide-background.internal.test.ts`: focused structural, validation, no-op, mutation, clone-on-write, and GC tests.
- Modify `packages/model/src/simple-fill.internal.ts`: consume the promoted public `SimpleFill` type without changing existing fill behavior.
- Modify `packages/model/src/slide.ts`: route the live `background` property through the new model module.
- Modify `packages/model/src/dependency.internal.ts`: include strict background image targets in slide-deletion media collection.
- Modify `packages/model/src/index.ts`: export the new public types.
- Create `packages/sdk/src/slide-background-source.ts`: strict high-level image-background options and source resolution.
- Create `packages/sdk/src/slide-background-source.test.ts`: options, source, abort, MIME, and zero-mutation tests.
- Modify `packages/sdk/src/index.ts`: add `PptxDocument.setSlideBackgroundImage()` and use the strict background projection for gradient diagnostics.
- Extend model, SDK, adapter, root-package, packed Node, real-browser, validator, gallery, and public documentation evidence.

---

### Task 1: Public types, normalization, and strict reader

**Files:**

- Create: `packages/model/src/slide-background.ts`
- Create: `packages/model/src/slide-background.internal.ts`
- Create: `packages/model/src/slide-background.internal.test.ts`
- Modify: `packages/model/src/simple-fill.internal.ts`
- Modify: `packages/model/src/index.ts`

**Interfaces:**

- Consumes: `RichTextColor`, `RasterImageContentType`, `GradientFill`, `GradientCodec.decode()`, `readSimpleFillChoice()`, and OPC relationship/part lookup.
- Produces:

```ts
export type SimpleFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export interface SlideBackgroundImage {
  readonly kind: 'image';
  readonly contentType: RasterImageContentType;
  readonly bytes: Uint8Array;
}

export type SlideBackground = SimpleFill | GradientFill | SlideBackgroundImage;

export function normalizeSlideBackground(
  value: unknown,
): SlideBackground | undefined;

export function readSlideBackground(
  pkg: OpcPackage,
  slidePartUri: string,
): SlideBackground | undefined;
```

- [ ] **Step 1: Add failing public-type and normalization tests**

Compile and execute representative states:

```ts
const values: readonly (SlideBackground | undefined)[] = [
  undefined,
  { kind: 'none' },
  { kind: 'solid', color: { kind: 'srgb', value: 'ff3399' }, transparency: 50 },
  {
    kind: 'linear-gradient', angle: 45, stops: [
      { offset: 0, color: 'FF0000' },
      { offset: 1, color: '0000FF', alpha: 0.5 },
    ],
  },
  { kind: 'image', contentType: 'image/png', bytes: Uint8Array.of(1, 2, 3) },
];
expect(values.map(normalizeSlideBackground)).toHaveLength(5);
```

Require uppercase normalized sRGB, transparency quantized to 0.001%, deep-detached/frozen gradient arrays, rectangles,
color transforms, and detached image bytes. Reject null/array/class/inherited/accessor/symbol/unknown fields, invalid
kind/color/content type, empty/non-byte image data, non-finite/out-of-range transparency, invalid gradient enums,
fewer than two stops, sparse arrays, and non-finite stop/angle/rectangle/alpha values without invoking accessors.

- [ ] **Step 2: Add failing strict-reader fixtures**

Build compact real-namespace slide packages and require:

```ts
expect(readSlideBackground(pkg, '/ppt/slides/slide1.xml')).toEqual({
  kind: 'solid',
  color: { kind: 'srgb', value: 'FF3399' },
  transparency: 50,
});
```

Cover absent direct background, `a:noFill`, sRGB/scheme solid, transparency, linear/path gradient, and PNG/JPEG/GIF
internal image relationships. For image, assert bytes are copied on every read. Snapshot parts, relationships, XML,
graph, and journal before reads and require exact equality afterward.

- [ ] **Step 3: Add failing unsafe-reader fixtures**

Require `undefined` for wrong namespaces, descendant traps, duplicate `p:cSld/p:bg/p:bgPr`, multiple fill choices,
empty `p:bgPr`, `p:bgRef`, pattern/group fill, malformed solid/gradient, missing/duplicate `a:blip`, unqualified or wrong-
namespace embed, external/wrong-type/duplicate/dangling relationships, missing target, and non-raster target MIME.

- [ ] **Step 4: Implement public types and reuse `SimpleFill`**

Move only the exported `SimpleFill` type ownership into `slide-background.ts`; make `simple-fill.internal.ts` import it.
Keep all existing simple-fill runtime helpers and tests unchanged. Export `slide-background.ts` from model index and
import `GradientFill` as a public type from `@pptx/codecs`.

- [ ] **Step 5: Implement descriptor-safe normalization**

Normalize each union branch before package access. Copy and freeze arrays/objects recursively, preserve valid
`OoxmlColor` source/transforms, canonicalize sRGB hex, and return new image bytes. Use exact closed-key tables per kind,
not a permissive shared object cast.

- [ ] **Step 6: Implement the namespace-aware direct reader**

Walk only direct children whose resolved namespaces equal PresentationML/DrawingML. For image, resolve exactly one
relationship-namespace `embed` attribute to exactly one internal standard image relationship and an existing part with
PNG/JPEG/GIF content type. Use `GradientCodec.decode()` only after selecting a unique direct `a:gradFill` candidate.

- [ ] **Step 7: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/simple-fill.internal.test.ts \
  packages/model/src/slide-background.internal.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/slide-background.ts \
  packages/model/src/slide-background.internal.ts \
  packages/model/src/slide-background.internal.test.ts \
  packages/model/src/simple-fill.internal.ts packages/model/src/index.ts
git diff --cached --check
git commit -m "feat: read strict slide backgrounds"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review the public declaration surface, strict direct-chain behavior, input detachment, and proof that no read mutates
the package.

---

### Task 2: Atomic non-image background editing

**Files:**

- Modify: `packages/model/src/slide-background.internal.ts`
- Modify: `packages/model/src/slide-background.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.ts`

**Interfaces:**

- Consumes: Task 1 normalized values/reader, simple-fill renderer/equality, gradient encoder/diagnostics, slide XML transaction access.
- Produces:

```ts
export function replaceSlideBackground(
  pkg: OpcPackage,
  slidePartUri: string,
  value: unknown,
): void;

export class SlideModel {
  get background(): SlideBackground | undefined;
  set background(value: SlideBackground | undefined);
}
```

- [ ] **Step 1: Write failing create/edit/clear tests**

From a zero-input presentation, execute:

```ts
const slide = document.addSlide();
slide.background = { kind: 'none' };
slide.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
};
slide.background = {
  kind: 'path-gradient',
  path: 'circle',
  stops: [
    { offset: 0, color: 'FFFFFF' },
    { offset: 1, color: '000000' },
  ],
};
slide.background = undefined;
```

Require legal `p:bg/p:bgPr`, exactly one fill choice, `p:bg` before `p:spTree`, explicit `a:noFill`, correct alpha,
canonical gradient, and complete direct `p:bg` removal on clear.

- [ ] **Step 2: Write failing lossless patch and no-op tests**

Use imported `p:bgPr` attributes, `a:effectLst`, unknown siblings/comments/whitespace, and unrelated shape XML. Require
supported fill replacement to preserve every unowned byte span. Assign each normalized value twice and assert the
second call leaves slide bytes, parts, relationships, graph, and journal unchanged. Clearing an absent background is
also an exact no-op.

- [ ] **Step 3: Write failing opaque replacement and rollback tests**

Replace `p:bgRef`, pattern fill, empty/ambiguous `p:bgPr`, and wrong-namespace direct background with canonical supported
background. Require exactly one direct canonical `p:bg`, with no descendant trap modified. Exercise invalid setter input,
malformed/missing `p:cSld`, and an outer transaction that throws after a successful edit; all failure snapshots must
match byte-for-byte.

- [ ] **Step 4: Implement canonical rendering and local patching**

Render:

```xml
<p:bg><p:bgPr>{fill}<a:effectLst/></p:bgPr></p:bg>
```

For a safe existing `p:bgPr`, replace only its unique direct fill choice. For absent background, insert before the first
direct `p:cSld` child. For `p:bgRef` or unsafe background structure, replace the direct `p:bg` element. On clear, remove
only the direct `p:bg`. Set the slide part once per changed transaction.

- [ ] **Step 5: Route `SlideModel.background` and write diagnostics**

Replace the current `GradientCodec.get/setSlideBackground()` calls with `readSlideBackground()` and
`replaceSlideBackground()`. In `PptxDocument.write()`, read `slide.background`; only gradient kinds are passed to
`GradientCodec.diagnostics()`. Keep the codec's existing public gradient methods source-compatible.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/slide-background.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/codecs/src/codecs.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/slide-background.internal.ts \
  packages/model/src/slide-background.internal.test.ts packages/model/src/slide.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.ts
git diff --cached --check
git commit -m "feat: edit direct slide backgrounds"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review schema order, lossless spans, no-op behavior, rollback, and compatibility diagnostics before committing.

---

### Task 3: Raster background relationships and lifecycle

**Files:**

- Modify: `packages/model/src/slide-background.internal.ts`
- Modify: `packages/model/src/slide-background.internal.test.ts`
- Modify: `packages/model/src/dependency.internal.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: Task 2 mutation transaction, image relationship type, OPC graph incoming, part allocation/update/removal, relationship-reference counting.
- Produces:

```ts
export function slideBackgroundMediaTargets(
  pkg: OpcPackage,
  slidePartUri: string,
): readonly string[];
```

`replaceSlideBackground()` gains complete image create/replace/remove and cleanup behavior.

- [ ] **Step 1: Write failing canonical image creation tests**

Assign PNG, JPEG, and GIF states. Require copied bytes, `/ppt/media/backgroundN.png|jpeg|gif`, exact content type,
one internal standard image relationship, and canonical:

```xml
<a:blipFill><a:blip r:embed="rIdN"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
```

Read the result immediately and after write/reopen. Mutate both caller bytes and getter bytes and require stored payload
to remain unchanged.

- [ ] **Step 2: Write failing replacement and exact no-op tests**

Cover identical bytes/MIME no-op, exclusive same-MIME in-place target update with stable relationship id/URI, MIME
change with stable exclusive relationship id and new target extension, image-to-none/solid/gradient/clear, and non-image-
to-image. Require superseded unreferenced relationships/parts to disappear and unrelated media to remain.

- [ ] **Step 3: Write failing shared-id and clone-on-write tests**

Create fixtures where:

- background and a `p:pic` use the same relationship id;
- two slide relationships use one target;
- a duplicated slide shares the original target;
- multiple slide relationships target the same media part.

On background edit, require a new relationship when the id has another XML reference, a new target whenever graph
incoming exceeds the current exclusive role, and no changes to the original picture/slide payload. Reopen both slides
and repeat replacement in the opposite direction.

- [ ] **Step 4: Write failing clear/delete GC and rollback tests**

Clear or replace image backgrounds and verify target deletion only at incoming zero. Delete source then duplicate and
duplicate then source; shared target must survive the first deletion and disappear after the last. Include a target also
used by a regular picture/media poster. Force outer rollback after target/relationship/XML changes and require an exact
package snapshot restore.

- [ ] **Step 5: Implement relationship-aware image mutation**

Before patching, collect relationship-namespace references inside the old direct background and count every relationship
reference across the slide XML. Reuse an exclusive same-extension target in place; otherwise allocate a detached
`backgroundN` part. Reuse an exclusive relationship id by retargeting; otherwise add a new relationship and patch only
the direct background embed. After XML update, remove old relationships with zero XML references, then remove old media
targets with graph incoming zero.

- [ ] **Step 6: Extend slide-deletion media target collection**

Make `mediaSlideDependencyTargets()` union its existing `p:pic` media targets with `slideBackgroundMediaTargets()`.
Keep `garbageCollectMediaDependencies()` unchanged so deletion still occurs only after the slide part/relationships are
gone and graph incoming is zero.

- [ ] **Step 7: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/model/src/slide-background.internal.test.ts \
  packages/model/src/model.test.ts
pnpm typecheck
pnpm test:performance
git diff --check
git add packages/model/src/slide-background.internal.ts \
  packages/model/src/slide-background.internal.test.ts \
  packages/model/src/dependency.internal.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: manage slide background images"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review shared relationship-id isolation separately from shared target isolation, and inspect every deletion path for
incoming-zero protection.

---

### Task 4: SDK raster source API and six-format round-trip

**Files:**

- Create: `packages/sdk/src/slide-background-source.ts`
- Create: `packages/sdk/src/slide-background-source.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: `RasterImageSource`, `resolveRasterImageSource()`, `assertRasterImageContentType()`, and Task 3 model image state.
- Produces:

```ts
export interface SetSlideBackgroundImageOptions {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
}

export async function resolveSlideBackgroundImage(
  source: RasterImageSource,
  options?: SetSlideBackgroundImageOptions,
): Promise<SlideBackgroundImage>;

export class PptxDocument {
  setSlideBackgroundImage(
    slideIndex: number,
    source: RasterImageSource,
    options?: SetSlideBackgroundImageOptions,
  ): Promise<void>;
}
```

- [ ] **Step 1: Write failing options and source-resolution tests**

Require a closed ordinary/null-prototype own-data options object, detached result bytes, optional exact content type,
and valid AbortSignal. Reject null/array/class/inherited/accessor/symbol/unknown options, SVG, malformed/truncated
PNG/JPEG/GIF, asserted MIME mismatch, explicit MIME mismatch, aborted source, oversized stream, and invalid chunks.

- [ ] **Step 2: Write failing portable-source integration tests**

Set backgrounds from relative/absolute path, HTTP redirect/query URL, browser-relative fetch, strict data URI,
`Uint8Array`, `ArrayBuffer`, `Blob`/`File`, Web stream, and async iterable. Reuse existing raster source fixtures; assert
detected MIME, exact payload, one background relationship, and zero picture shapes.

- [ ] **Step 3: Write failing document API and zero-mutation tests**

Require invalid `slideIndex` to throw before source consumption. Require source/options/signature/MIME/abort/model failure
to leave package snapshots unchanged. On success, require `Promise<void>`, immediate live `slide.background`, and
subsequent synchronous none/solid/gradient/clear edits through the same property.

- [ ] **Step 4: Implement the focused SDK module and facade**

Resolve and inspect the complete source before assigning:

```ts
const resolved = await resolveRasterImageSource(source, normalized.signal);
assertRasterImageContentType(normalized.contentType, resolved);
slide.background = {
  kind: 'image',
  contentType: resolved.info.contentType,
  bytes: resolved.bytes,
};
```

Export only the public options type from the SDK root; keep normalization helper internal to the SDK package.

- [ ] **Step 5: Add all-six-format write/reopen tests**

For pptx/pptm/potx/potm/ppsx/ppsm, create a zero-input document, add a slide, set an image background, write/open twice,
and require preserved format profile, payload, relationship, direct XML, and validator errors `[]`. Also round-trip
none, solid transparency, and gradient in the matrix without duplicating source-loader cases.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/slide-background-source.test.ts packages/sdk/src/index.test.ts \
  packages/model/src/slide-background.internal.test.ts
pnpm typecheck
pnpm build
git diff --check
git add packages/sdk/src/slide-background-source.ts \
  packages/sdk/src/slide-background-source.test.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: load raster slide backgrounds"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review that no Node-only import enters browser code and that the SDK never mutates a document until source resolution
has fully succeeded.

---

### Task 5: PptxGenJS conformance and package validation

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/validator/src/index.test.ts`

**Interfaces:**

- Consumes: locked `pptxgenjs@4.0.1`, adapter `importPptxGenJS()`, public SDK/model background API, validator profiles.
- Produces: permanent imported-output and public-output conformance evidence.

- [ ] **Step 1: Generate real PptxGenJS fixtures in tests**

Create slides with default, `{ type: 'none' }`, `{ color: 'FF3399' }`, `{ color: 'FF3399', transparency: 50 }`,
`{ data: PNG_DATA_URI }`, and deprecated `{ fill: '00FF00' }`. Write through public PptxGenJS, import bytes, and assert
native projection for every supported output.

- [ ] **Step 2: Compare native public output structurally**

Create equivalent native slides and compare fill kind, normalized color, alpha, direct chain, image relationship type/
mode, target content type/extension, payload SHA-256, and stretch/fillRect. Assert native noFill is legal and record the
PptxGenJS empty-`bgPr` combination only as an intentional divergence fixture.

- [ ] **Step 3: Exercise imported edit/reopen and validators**

Edit PptxGenJS solid → image → gradient → clear and PptxGenJS image → solid. Reopen after each write. Require unrelated
PptxGenJS XML/relationships to remain, superseded media to be collected, and validation under PowerPoint 2010/current,
Keynote current, Google Slides import, and LibreOffice current to contain no unexpected errors.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/validator/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts packages/validator/src/index.test.ts
git diff --cached --check
git commit -m "test: compare slide backgrounds with pptxgenjs"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review that all comparisons use public PptxGenJS APIs and that no production manifest gains a PptxGenJS dependency.

---

### Task 6: Packed Node, browser, CLI, gallery, and client evidence

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Temporary only: clean-build manifests, packed consumer, type fixture, native/PptxGenJS galleries, rendered PNGs,
  validation JSON, LibreOffice round trip, and comparison ledger under `/tmp/pptx-slide-background-*`.

**Interfaces:**

- Consumes: actual packed `@jiayunxie/pptx` artifact, browser conditional export, generated declarations, CLI, and Task 5 conformance behavior.
- Produces: permanent smoke coverage plus reproducible real-client release evidence.

- [ ] **Step 1: Extend the permanent packed-package smoke**

Using only public imports from the installed tarball, create none/solid/gradient/PNG background slides, duplicate and
replace the image background, clear another background, write/reopen, and inspect isolation/cleanup. Compile a TS
consumer that imports `SlideBackground`, `SlideBackgroundImage`, `SimpleFill`, and
`SetSlideBackgroundImageOptions`.

- [ ] **Step 2: Extend the permanent real-browser smoke**

From `dist/browser.js`, set an image background from Blob and data URI, assign solid/gradient directly, write Blob,
reopen bytes, and return fill kinds, payload hashes, relationship counts, and validator diagnostics. Assert no Node
builtin import appears in the browser bundle.

- [ ] **Step 3: Build twice, hash, pack, install, and run smoke**

Use two clean build directories and SHA-256 manifests to require identical dist output. Pack the second build, install
that exact tarball into a fresh consumer without workspace links, then run Node, declarations, browser, and installed
CLI `doctor/inspect/validate` against the generated deck.

- [ ] **Step 4: Generate and structurally inspect the background gallery**

Create slides for inherited, noFill, sRGB, scheme+transparency, linear gradient, path gradient, PNG, JPEG, GIF,
duplicate-before-replacement, and cleared inheritance. Use `pptx-inspect` to record slide titles, direct background
choices, relationships, targets, content types, and validator results. Keep all gallery files under `/tmp`.

- [ ] **Step 5: Render and round-trip with clients**

Open/render the native and PptxGenJS galleries in PowerPoint and LibreOffice. Save the native gallery through
LibreOffice, reopen with the SDK, rerun strict package validation, compare slide count/order/background kind/payload
hash, and record any client normalization without weakening native correctness assertions.

- [ ] **Step 6: Run full gates, review, commit, and push**

```bash
pnpm test
pnpm test:performance
pnpm typecheck
pnpm build
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
git diff --check
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed slide backgrounds"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review the clean build manifests, tarball dependency tree, browser console/network errors, CLI JSON, package graph,
gallery counts, render output, and LibreOffice round-trip ledger. Confirm no `/tmp` artifact is staged.

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

- Consumes: final Task 1–6 API, behavior, test counts, build hashes, package inspection, render, and client round-trip evidence.
- Produces: accurate user-facing API examples, support matrix, intentional-difference record, and next-work ordering.

- [ ] **Step 1: Document the model and SDK APIs**

Add runnable examples for:

```ts
slide.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
};
await document.setSlideBackgroundImage(0, './background.png');
slide.background = { kind: 'none' };
slide.background = undefined;
```

Explain detached image bytes, direct-only scope, noFill versus inherited clear, image source forms, and clone-on-write.

- [ ] **Step 2: Update compatibility and implementation progress**

Record PptxGenJS 4.0.1 equivalent solid/transparency/image behavior and the intentional correction for empty `bgPr`.
Move slide background from pending to completed. Set the next item to slide number, followed by default color, then
master/layout/placeholder and the remaining advanced/output/helper audit groups.

- [ ] **Step 3: Record exact verification evidence**

Write the final focused/full test counts, performance result, deterministic build manifest hash/count, tarball/browser/
types/CLI results, gallery slide/relationship/media counts, validator profile counts, and LibreOffice/PowerPoint outcomes.
Do not use approximate claims when an exact result is available.

- [ ] **Step 4: Run final release gates and inspect the entire diff history**

```bash
pnpm exec vitest run packages/model/src/slide-background.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/slide-background-source.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  packages/validator/src/index.test.ts
pnpm test
pnpm test:performance
pnpm typecheck
pnpm build
git diff --check
git log --oneline 49e7d23..HEAD
git status --short --branch
```

Review public API names, declarations, strict-read semantics, all image isolation/GC cases, docs against observed
evidence, and the absence of generated artifacts.

- [ ] **Step 5: Commit, push, and verify remote parity**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --cached --check
git commit -m "docs: complete slide background support"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected final divergence is `0 0`; the only remaining untracked entry is the existing `.pnpm-store/` directory.
