# Stable Media Lifecycle Implementation Plan

**Goal:** Make audio and video pictures stable live shape models that support metadata, transform, playback-preference,
source, poster, and deletion editing with PptxGenJS-compatible import, relationship-aware isolation, and exact rollback.

**Architecture:** Keep strict source resolution and OOXML/relationship mutation in `@pptx/codecs`, but decode media
pictures into a class-based `MediaModel` owned by the existing `SlideModel` shape registry. Every live getter resolves
current XML and relationships. Source/poster replacement uses content-aware deduplication and clone-on-write; object
and slide deletion collect only unreferenced `/ppt/media` payloads.

**Tech Stack:** TypeScript 5.8, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/codecs`, `@pptx/model`, JSZip,
PptxGenJS 4.0.1, tsup, Playwright browser smoke, `pptx-inspect`, ffmpeg, LibreOffice/Poppler.

## Global Constraints

- Preserve `PptxDocument.addAudio(slideIndex, source, options)`, `addVideo()`, and `media()` call signatures.
- Add one runtime `MediaModel` class in `@pptx/model`; `slide.shapes`, `slide.media`, document listing, and create return
  the same per-slide/per-shape object.
- Keep `shapeId`, `slidePartUri`, `kind`, `settings`, `mediaPartUri`, `externalUrl`, and `posterPartUri` source-compatible.
- Infer kind from standard relationship first, target MIME second, and marker element last; never rewrite during read.
- Read PptxGenJS 4.0.1 `a:videoFile` audio, `audio/mp3`, and duplicated-audio Microsoft media relationship defects.
- Preserve strict creation source, MIME, extension, byte-copy, data-URI, stream, poster, and transcoder behavior.
- Source replacement supports embedded↔external transitions without changing an object's audio/video kind.
- Poster replacement supports PNG/JPEG/GIF and omitted-source reset to the built-in default PNG.
- Deduplicate embedded media/posters by SHA-256 plus exact MIME; never share equal bytes under different MIME.
- In-place mutation is allowed only when every incoming relationship belongs exclusively to the edited picture's roles.
- If a relationship id is shared by multiple XML nodes, allocate a new relationship and patch only the edited picture.
- Remove superseded relationships only after their XML reference count reaches zero and delete `/ppt/media` targets
  only after package-graph incoming reaches zero.
- Metadata, transform, and playback edits never canonicalize imported nonconforming primary media roles.
- Playback settings remain the private preference extension; native `p:timing` is preserved and not edited here.
- Resolve and detach every asynchronous input before one synchronous package transaction.
- A failed operation leaves parts, content types, relationships, graph, XML, ZIP bytes, identity, and journal unchanged.
- Every task ends with focused tests, typecheck, diff review, an isolated commit, push, fetch, and divergence `0 0`.
- Never stage or commit `.pnpm-store/`, generated fixtures, probe decks, tarballs, renders, galleries, or validation output.

---

### Task 1: Strict media state and stable shape identity

**Files:**

- Create: `packages/codecs/src/media-state.internal.ts`
- Create: `packages/codecs/src/media-state.internal.test.ts`
- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Produces: `MediaState`, `readMediaState()`, codec `MediaDescriptor`, runtime model `MediaModel`, `ShapeKind` audio/video,
  `SlideModel.media`, `addAudio()`, and `addVideo()`.
- Preserves: deprecated `@pptx/codecs` `MediaModel` type alias and all current SDK method signatures.

- [ ] **Step 1: Write failing strict-state reader tests**

Build direct `p:pic` fixtures for canonical embedded/external audio and video, absent optional poster/media extension,
multiple unrelated pictures, dangling optional relationships, and unknown content types. Require detached state containing
kind, shape id, slide URI, primary internal/external target, poster target, metadata, and playback preferences without
package mutation.

- [ ] **Step 2: Lock PptxGenJS legacy import precedence**

Add fixtures where audio uses `a:videoFile`, MP3 is `audio/mp3`, and duplicate audio exposes only Microsoft media role
plus an audio-typed target. Require relationship type to win over marker, MIME to win over marker when the standard
relationship is absent, marker to remain the final fallback, and ordinary pictures to return no media state.

- [ ] **Step 3: Implement one pure state reader**

In `media-state.internal.ts`, inspect only the supplied direct picture and current relationships. Resolve the standard
kind role, Microsoft media role, primary source, poster role, `cNvPr` metadata, and private playback extension
independently. Accept `audio/mp3` only as an import alias. Return frozen plain containers and do not repair malformed
optional roles during a read.

Export the internal reader/types through `media.ts` with internal documentation so `@pptx/model` can consume them
without a dependency cycle. Rename the codec result interface to `MediaDescriptor` and retain:

```ts
/** @deprecated Use the runtime MediaModel from @pptx/model or @pptx/sdk. */
export type MediaModel = MediaDescriptor;
```

- [ ] **Step 4: Write failing identity and semantic-shape tests**

Require media pictures to decode as kind `audio`/`video`, never `ImageModel`; repeated `slide.shapes`, `slide.media`,
and `document.media()` reads must return `===` objects. Require create to return that same object, distinct slides to
own distinct handles after duplicate, move to preserve handles, rollback to preserve handles, and raw media↔image or
audio↔video semantic changes to replace the registry entry.

- [ ] **Step 5: Add the class-based model and slide entry points**

Add `MediaModel extends BaseShapeModel` with live readonly getters for compatibility state, `shapeId` aliasing `id`, and
`slidePartUri` exposing the owner. Extend `ShapeKind` and `SemanticShape`; call `readMediaState()` before the generic
picture branch in `decodeShape()`.

Add `SlideModel.media` as a filtered view of `shapes`. `SlideModel.addAudio()` / `addVideo()` call the codec, reconcile
the returned shape id through the existing registry, and return the runtime class. Change SDK methods to delegate to
the selected slide. Use an explicit SDK export for the runtime `MediaModel` so codec/model star exports are unambiguous.

- [ ] **Step 6: Prove live reads and no-op preservation**

After direct relationship/XML mutation, require an existing handle to report current URIs, URL, settings, name, and alt
text. Require no-mutation listing to leave package bytes, journal, relationship records, and graph unchanged. Require a
deleted/missing picture to use the existing shape-resolution error path rather than return cached state.

- [ ] **Step 7: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/codecs/src/media-state.internal.ts packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media.ts packages/codecs/src/codecs.test.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: add stable media models"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Review that there is one identity registry, collection order follows the current shape tree, and the codec descriptor is
never returned from the high-level API.

---

### Task 2: Live metadata, transform, and playback-preference editing

**Files:**

- Create: `packages/codecs/src/media-edit.internal.ts`
- Create: `packages/codecs/src/media-edit.internal.test.ts`
- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Adds: `MediaModel.name`, `altText`, and `settings` setters while retaining inherited `transform` / `setTransform()`.
- Produces: strict metadata/playback normalizers and narrow picture patch helpers.

- [ ] **Step 1: Write failing public metadata and transform tests**

Through a created model and an imported PptxGenJS model, edit special-character, empty, and Unicode names; set, empty,
and clear alt text; and change x/y/width/height/rotation/flips through inherited transform. Require the same model
identity immediately and after reopen, with only direct `p:cNvPr` attributes and transform nodes changed.

- [ ] **Step 2: Write failing strict metadata validation tests**

Reject non-strings and illegal XML control characters. Verify no getter invocation through unsafe runtime values where
applicable, no coercion, no partial package change, and usable same-handle retry. Same-value metadata writes and clearing
an already absent alt text must be exact no-ops.

- [ ] **Step 3: Write failing playback preference tests**

Assign complete strict objects for click/auto, loop, hide-when-stopped, and finite `0..1` volume. Require detached frozen
getter state, creation defaults for omitted fields, exact private extension replacement, and `settings = undefined` to
remove only the owned extension. Reject arrays, class/custom-prototype objects, inherited/enumerable/unknown/symbol
keys, accessors, invalid literals/booleans, and non-finite/out-of-range volume without invoking getters.

- [ ] **Step 4: Implement narrow metadata and playback patching**

Add descriptor-safe normalizers and helpers that locate the direct media picture by shape id. Patch `cNvPr@name` and
optional `@descr` through the lossless attribute API. Replace or remove only the fixed library playback extension URI,
preserving the media extension, unknown siblings, native timing, actions, relationships, and picture geometry.

Expose synchronous codec methods used only by `SlideModel`; add live model accessors that delegate by id and return
current state through `readMediaState()`. Do not cache relationship ids or the settings object.

- [ ] **Step 5: Add rollback and isolation tests**

Inject XML serialization/setPart failures and wrap successful setters in an outer failing transaction. Require exact
part/relationship/graph/ZIP/journal restoration, preserved handle identity, and untouched sibling pictures/extensions.
Duplicate a slide, edit each metadata/settings/transform surface on one copy, and require the source XML unchanged.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media-edit.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/codecs/src/media-edit.internal.ts packages/codecs/src/media-edit.internal.test.ts \
  packages/codecs/src/media.ts packages/codecs/src/codecs.test.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit live media metadata"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 3: Clone-on-write primary media source replacement

**Files:**

- Create: `packages/codecs/src/media-replace.internal.ts`
- Create: `packages/codecs/src/media-replace.internal.test.ts`
- Modify: `packages/codecs/src/media-source.internal.ts`
- Modify: `packages/codecs/src/media-source.internal.test.ts`
- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Adds: `ReplaceMediaSourceOptions` and `MediaModel.replaceSource(source, options?): Promise<this>`.
- Produces: reusable strict single-source resolution, relationship reference counting, primary-role canonicalization,
  and content-aware clone-on-write.

- [ ] **Step 1: Write failing option and source-resolution tests**

Accept only own data properties `contentType`, `fileName`, and `transcode` on an ordinary/null-prototype object. Reuse
all creation source forms, seven audio/video MIME families, `.m4v`, strict data URIs, File/path inference, copied bytes,
stream chunks, and transcode output rules. Reject placement/poster/metadata keys, kind-mismatched MIME, empty sources,
invalid extensions, unsafe descriptors, and external URL plus transcode before package access.

- [ ] **Step 2: Extract strict single-source resolution without weakening creation**

Refactor `media-source.internal.ts` so creation and replacement call the same normalized embedded/external resolver.
Keep poster rules separate. Preserve all existing creation tests byte-for-byte and ensure replacement resolves/detaches
before mutation. Do not add static Node imports to the browser path.

- [ ] **Step 3: Write failing transition and canonical-role tests**

Cover embedded→embedded with same/different MIME and extension, embedded→external, external→embedded, and
external→external for audio and video. Require the model kind/identity, poster, transform, metadata, settings, actions,
and unknown extensions to remain unchanged. Replacing PptxGenJS legacy audio must canonicalize only the current primary
roles to standard audio + Microsoft media + `a:audioFile`; metadata-only edits must continue preserving legacy bytes.

- [ ] **Step 4: Write failing deduplication and clone-on-write tests**

Require equal payload+MIME to reuse an existing target, equal bytes under another MIME to remain separate, and an
exclusive target to retain its URI only when extension/content type remain compatible. Cover duplicated slides where
two role relationships per picture point to one shared target, unrelated incoming relationships, same-slide pictures
sharing a target, and two pictures sharing the same relationship id. Require changes to remain limited to the edited
picture.

- [ ] **Step 5: Implement atomic primary replacement**

Normalize and resolve before mutation. Inside one transaction, locate current roles, choose existing/dedup/new target,
write or reuse the part, update or allocate kind and Microsoft relationships based on XML reference ownership, patch
only current marker/extension ids, remove now-unreferenced superseded relationships, then collect zero-incoming old
`/ppt/media` targets. External replacements retain only the standard external kind relationship and remove the owned
Office media role from the current picture.

Add `MediaCodec.replaceSource()`, `SlideModel.replaceMediaSource()`, and the live async model method returning `this`.
Do not allow replacement to change audio to video or video to audio.

- [ ] **Step 6: Add failure-injection and async-detachment tests**

Pause streams/transcoders, mutate caller input, resume, and require detached bytes/options. Inject allocation, part,
relationship, XML patch, serialization, setPart, and old-target-GC failures for each transition. Require exact package
snapshot and same-handle recovery after every rejection and after an outer transaction rollback.

- [ ] **Step 7: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-source.internal.test.ts \
  packages/codecs/src/media-replace.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
pnpm --filter @jiayunxie/pptx build
rg -n 'from "node:|from '\''node:|require\("node:' packages/pptx/dist/browser.js
git diff --check
git add packages/codecs/src/media-replace.internal.ts packages/codecs/src/media-replace.internal.test.ts \
  packages/codecs/src/media-source.internal.ts packages/codecs/src/media-source.internal.test.ts \
  packages/codecs/src/media.ts packages/codecs/src/codecs.test.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: replace live media sources"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

The static Node import search must return no matches.

---

### Task 4: Clone-on-write poster replacement and default reset

**Files:**

- Modify: `packages/codecs/src/media-replace.internal.ts`
- Modify: `packages/codecs/src/media-replace.internal.test.ts`
- Modify: `packages/codecs/src/media-source.internal.ts`
- Modify: `packages/codecs/src/media-source.internal.test.ts`
- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Adds: `ReplaceMediaPosterOptions` and `MediaModel.replacePoster(source?, options?): Promise<this>`.
- Preserves: current primary source roles and every non-poster picture field.

- [ ] **Step 1: Write failing strict poster input tests**

Accept PNG/JPEG/GIF bytes, data URI, local path, Blob/File, Web Stream, and async iterable with exact MIME/extension
assertion/inference. `replacePoster()` with no source restores a copied built-in PNG. Reject HTTP/HTTPS URLs, empty
payloads, media MIME, mismatch, unsupported extension, unsafe/unknown option keys, accessors, and transcode.

- [ ] **Step 2: Write failing direct replacement and reset tests**

Replace each poster format, change extensions safely, reopen, and verify exact bytes/content type/part URI plus the live
getter. Reset explicit poster to default twice and require stable dedup/no-op behavior. Preserve primary media/internal
or external relationships, marker, metadata, settings, transform, actions, crop, and unknown picture extensions.

- [ ] **Step 3: Write failing poster COW tests**

Cover exclusive poster target/id, duplicate-slide shared target, same-slide target dedup, unrelated incoming image
relationship, and two pictures sharing one poster relationship id. Require in-place update only for a truly exclusive
compatible target, otherwise reuse/allocate and patch only the current blip. Delete superseded relationship/target only
when XML references/incoming reach zero.

- [ ] **Step 4: Implement atomic poster replacement**

Resolve strict poster state before mutation. Reuse the source resolver's poster branch and built-in PNG copy. Inside one
transaction choose dedup/new/in-place part, update or allocate an image relationship, patch only direct
`a:blip/@r:embed`, remove an unreferenced old relationship, and collect only zero-incoming old `/ppt/media` target.

Add codec/slide delegation and the live model async method returning `this`. Construct a missing imported poster role
canonically when the direct picture target is unambiguous; reject multiple direct blips before mutation.

- [ ] **Step 5: Add rollback and deletion-regression tests**

Inject source, allocation, relationship, blip patch, part write, and GC failures plus an outer rollback. Require exact
snapshot restoration and handle reuse. Re-run primary source replacement and existing codec deletion suites to prove
poster changes did not alter their ownership rules.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-source.internal.test.ts \
  packages/codecs/src/media-replace.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/codecs/src/media-replace.internal.ts packages/codecs/src/media-replace.internal.test.ts \
  packages/codecs/src/media-source.internal.ts packages/codecs/src/media-source.internal.test.ts \
  packages/codecs/src/media.ts packages/codecs/src/codecs.test.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: replace live media posters"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 5: Reference-safe object and slide deletion lifecycle

**Files:**

- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/model/src/dependency.internal.ts`
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Adds: `MediaModel.remove()` and `SlideModel.deleteMedia(shapeId)`.
- Strengthens: `MediaCodec.delete()` XML-reference safety and `PresentationModel.deleteSlide()` narrow media/poster GC.

- [ ] **Step 1: Write failing object removal tests**

Remove canonical embedded/external and PptxGenJS legacy objects through `MediaModel.remove()`. Require the picture to
disappear, only exclusively referenced relationships to be removed, external targets never fetched/deleted, and only
zero-incoming `/ppt/media` primary/poster parts to be collected. The removed handle must throw on every live getter or
mutator and must never re-create the shape.

- [ ] **Step 2: Lock shared relationship and payload behavior**

Cover same target across two pictures, paired kind/media relationships, a relationship id shared by two pictures,
duplicate slides, an unrelated relationship into the target, and a target outside `/ppt/media`. Removing one object
must leave every survivor valid; removing the final owner collects only the eligible conventional media part.

- [ ] **Step 3: Implement reference-safe object deletion**

Resolve the current picture's referenced role ids and targets, remove the picture, recompute slide-wide XML relationship
references, remove only zero-reference ids, then collect captured zero-incoming `/ppt/media` targets in one transaction.
Add slide delegation and `MediaModel.remove()`; keep the old codec entry point source-compatible.

- [ ] **Step 4: Write failing duplicate/move/slide-delete tests**

Duplicate a slide containing shared audio/video/posters, move both slides, replace one source/poster, delete either slide,
then the final slide. Require identity rules, source/poster isolation, retained shared targets after the first deletion,
and final GC after the last deletion. Include external media, ordinary image targets, owned chart dependencies, opaque
relationships, and a `/ppt/media` part still referenced from another non-slide part.

- [ ] **Step 5: Add narrow slide media target capture and GC**

Add dependency helpers that inspect only media pictures on the soon-to-be-deleted slide and capture their internal
primary/poster targets under `/ppt/media`. After slide part/relationship deletion and existing owned-subgraph GC, delete
only captured targets whose current graph incoming is zero. Do not change shared relationship lifecycle classification,
recursively traverse media, or collect arbitrary images.

- [ ] **Step 6: Add rollback and fault-injection coverage**

Inject picture removal, relationship removal, slide deletion, owned-GC, and media-GC failures and outer transaction
rollback. Require exact parts/content types/relationships/XML/graph/ZIP/journal restoration. Original media handles must
remain `===` and live after rollback.

- [ ] **Step 7: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media-replace.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/codecs/src/media.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/dependency.internal.ts packages/model/src/presentation.ts packages/model/src/shapes.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: manage media deletion lifecycle"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 6: Public SDK, PptxGenJS, and six-format conformance

**Files:**

- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Create: `packages/pptx/src/index.test.ts`

**Interfaces:**

- Consumes: Tasks 1–5 only through public APIs and packed-facing exports.
- Produces: declaration/runtime compatibility, PptxGenJS legacy import/edit evidence, and six-format write/reopen proof.

- [ ] **Step 1: Lock public type and runtime exports**

Compile one `MediaModel` value from add/list/shapes and every setter/replacement/removal method through `@pptx/sdk` and
`@jiayunxie/pptx`. Require `instanceof MediaModel`, stable `===`, `shapeId`/`slidePartUri` compatibility, fluent async
replacement returns, and compile-time rejection of invalid option keys, source values, kinds, metadata, and settings.

- [ ] **Step 2: Add complete public lifecycle scenarios**

From zero input create audio/video, edit metadata/transform/settings, replace embedded↔external sources, replace/reset
posters, duplicate/move, isolate shared resources, remove objects, write/reopen, and continue editing through old live
handles where valid. Assert strict validation and complete no-mutation snapshots for public failures.

- [ ] **Step 3: Add all MIME and six-format coverage**

Across focused cases cover all four audio, three video, and three poster MIME families. For each
`pptx|pptm|potx|potm|ppsx|ppsm`, create, mutate, duplicate, remove one shared object, write/reopen twice, retain the format
profile, and require exact shape order/kind/metadata/settings/transforms/part metadata/relationships plus no validation
errors.

- [ ] **Step 4: Add PptxGenJS 4.0.1 legacy lifecycle fixtures**

Generate fixtures only through public `addMedia()` and public write APIs. Cover audio data/cover, video path/cover, and
duplicate audio. Require stable import of `a:videoFile` audio, `audio/mp3`, and Microsoft-media duplicate role; narrow
metadata edits preserve original defects, source replacement canonicalizes only the edited picture, poster replacement
isolates it, duplicate lifecycle remains safe, and removal leaves valid survivors.

- [ ] **Step 5: Compare final supported semantics and divergences**

Compare picture order/kind/name/alt text/transform, media/poster bytes and MIME/extensions, relationships, click action,
Office media extension, strict reopen, and validation. Record canonical native output and strict-input differences as
intentional; never introduce coercion or reproduce malformed PptxGenJS output.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media-edit.internal.test.ts packages/codecs/src/media-replace.internal.test.ts \
  packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts packages/pptx/src/index.test.ts
pnpm typecheck
pnpm test
git diff --check
git add packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  packages/pptx/src/index.test.ts
git diff --cached --check
git commit -m "test: verify live media conformance"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 7: Packed/browser/client evidence and public documentation

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`
- Temporary only: packed consumer, declaration fixture, media fixtures, galleries, renders, validation JSON, and
  LibreOffice round-trip package under a fresh `/tmp/pptx-media-lifecycle-*` directory.

**Interfaces:**

- Consumes: the complete stable media lifecycle through the actual `@jiayunxie/pptx` tarball.
- Produces: permanent Node/browser/declaration/CLI coverage, deterministic build/client evidence, and accurate support
  accounting with the next media gap identified.

- [ ] **Step 1: Extend permanent packed Node and declaration smoke**

Use the installed tarball to create audio/video, prove stable class identity across shapes/media/document, edit all
metadata/settings/transform surfaces, replace embedded/external sources and posters with Blob/stream/data/path inputs,
duplicate and isolate shared targets, remove one object, write/reopen, and inspect relationships/GC. Add valid and
negative declaration fixtures and emit `stableMediaLifecycle: true` in the smoke JSON.

- [ ] **Step 2: Extend permanent browser smoke**

Through `dist/browser.js`, create data-URI/Blob/stream media, edit metadata/settings/transform, replace source/poster,
duplicate, remove, write Blob, reopen, and return identity, relationship, MIME/extension, payload signature, target
isolation, orphan count, and validator results. Keep the static Node-import scan empty and download smoke passing.

- [ ] **Step 3: Clean-build twice, pack, install, and run all public smoke**

Use a fresh temp directory. Build `@jiayunxie/pptx` twice from clean state, hash every dist file in stable path order,
require identical manifests, pack the second build, install only that tarball without workspace links, and run Node,
types, browser, plus installed `pptx-inspect doctor`, `package inspect`, and PowerPoint 2010 validation.

- [ ] **Step 4: Generate a complete lifecycle gallery**

Use ffmpeg to generate playable MP3/M4A/WAV/OGG and MP4/MOV/WebM inputs plus PNG/JPEG/GIF posters. Through the installed
tarball create an embedded zero-warning gallery and a separate external-warning deck. Cover all MIME/extensions,
metadata/settings/transforms, explicit/default posters, source/poster replacement, dedup, duplicate COW, shared rIds,
object deletion, slide deletion, and no orphan targets.

- [ ] **Step 5: Validate, render, inspect, and LibreOffice round-trip**

Require exact slide/picture/media/poster/relationship counts, hashes, MIME/extensions, role ownership, model identity
within each open session, strict reopen, zero orphan media, and PowerPoint 2010 validation 0 errors/0 warnings for the
embedded gallery. Render every slide at 180 DPI, run overflow checking, and inspect all poster placements. Save through
LibreOffice, reopen/render/validate, and record exact client normalization of roles, MIME, settings, posters, or numeric
transform values.

- [ ] **Step 6: Update public API and compatibility documentation**

Document live class identity, slide/document entry points, metadata/settings/transform editing, strict replacement
options, embedded/external transitions, poster reset, dedup/COW, object/slide GC, rollback, PptxGenJS legacy import and
canonicalization boundary, six-format results, packed/browser/CLI evidence, deterministic builds, and client results.

Move stable live media identity/editing and complete duplicate/move/delete isolation to supported. Keep native timing,
online video/remote fetch, captions/subtitles, crop/rounding/shadow/hyperlink/placeholder styles, transcoding engines,
and broad client certification pending. Set native media timing as the next media slice.

- [ ] **Step 7: Run final gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media-edit.internal.test.ts packages/codecs/src/media-replace.internal.test.ts \
  packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts packages/pptx/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
pnpm --filter @jiayunxie/pptx build
git diff --check
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js CHANGELOG.md README.md \
  packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md \
  docs/implementation-progress.md
git diff --cached --check
git commit -m "docs: document stable media lifecycle"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected final state: divergence `0 0`; only `.pnpm-store/` remains untracked.
