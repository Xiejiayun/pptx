# Stable Media Lifecycle Design

## Goal

Complete one independently reviewable editing slice: promote audio and video pictures from transient codec snapshots to
stable live shape models, then support metadata, transform, playback-preference, source, poster, and object-lifecycle
editing through those models.

The result must preserve model identity across collection reads, create, duplicate, move, successful mutation, and
transaction rollback. Media and poster replacement must isolate shared targets, object and slide deletion must reclaim
only unreferenced owned payloads, and imported PptxGenJS 4.0.1 output must remain readable despite its known media
relationship and MIME defects.

This slice builds on strict embedded media creation. It does not claim native timing-tree playback, online-video
creation, transcoding implementations, captions, or the final multi-client release matrix.

## Current-State Audit

The package already has most of the low-level mechanics required for a safe implementation:

- `MediaCodec` can create embedded or external audio/video pictures, resolve strict portable sources, deduplicate
  embedded media and posters by SHA-256 plus exact MIME, and perform transactional deletion;
- `SlideModel` owns a stable shape registry keyed by `p:cNvPr/@id` and reconciles semantic class changes;
- slide duplication intentionally shares image, audio, video, and Microsoft media targets;
- image and chart mutation already implement relationship-aware clone-on-write;
- the OPC graph reports incoming relationships and transactions roll back parts, content types, relationships, XML,
  and the mutation journal together.

The public media lifecycle is nevertheless incomplete:

- `MediaCodec.list()` creates a new interface snapshot on every call, so media has no stable JavaScript identity;
- `PptxDocument.addAudio()`, `addVideo()`, and `media()` construct fresh codec instances and bypass the shape registry;
- `decodeShape()` treats every `p:pic` as `ImageModel`, including media pictures;
- names, alt text, transform, playback preferences, source, and poster cannot be edited through a live media object;
- the main package has no object-level media removal API;
- slide deletion regards media as shared and therefore never reclaims media or poster parts whose final incoming
  relationships disappeared with the slide;
- media reading trusts `a:audioFile` / `a:videoFile` too early and cannot classify all known PptxGenJS output robustly.

## Approaches Considered

### A codec-owned media registry

Keep media outside the shape model and cache snapshots inside one long-lived `MediaCodec`. This is the smallest change,
but creates a second identity domain beside `SlideModel.#shapeModels`. Raw edits that change a picture's semantic class
would require cross-registry invalidation, and `slide.shapes` would still misrepresent media as images.

### Media in the slide shape registry — selected

Add a class-based `MediaModel` extending `BaseShapeModel`; classify media before generic images in `decodeShape()`; and
route creation, listing, editing, and removal through `SlideModel`. Identity, reconciliation, missing-shape behavior,
and transform handling then reuse the established shape architecture. The codec remains responsible for strict source
resolution and OOXML/relationship mutation, while the model provides the stable public handle.

### One media, timing, and online-video rewrite

Implement the entire PptxGenJS media surface together, including native timing trees and YouTube/Vimeo semantics. This
would mix three independently risky OOXML protocols, make rollback and compatibility failures harder to localize, and
prevent an independently shippable live editing slice.

## Public Model and Entry Points

`@pptx/model` adds a runtime `MediaModel` class and `ShapeKind` adds `audio | video`. `SemanticShape` includes
`MediaModel`, so `slide.shapes` reports the true semantic type of a media picture.

```ts
class MediaModel extends BaseShapeModel {
  readonly kind: 'audio' | 'video';

  get shapeId(): number;
  get slidePartUri(): string;

  get name(): string;
  set name(value: string);

  get altText(): string | undefined;
  set altText(value: string | undefined);

  get settings(): Readonly<MediaPlaybackSettings>;
  set settings(value: MediaPlaybackSettings | undefined);

  get mediaPartUri(): string | undefined;
  get externalUrl(): string | undefined;
  get posterPartUri(): string | undefined;

  replaceSource(source: MediaSource, options?: ReplaceMediaSourceOptions): Promise<this>;
  replacePoster(source?: MediaSource, options?: ReplaceMediaPosterOptions): Promise<this>;
  remove(): void;
}
```

`transform` and `setTransform()` remain inherited. All getters resolve current slide XML and relationships on every
call; no relationship id, URI, URL, settings object, name, or alt text is cached in the handle. Returned settings are
detached and frozen.

The collection and creation entry points become:

```ts
slide.media: readonly MediaModel[];
await slide.addAudio(source, options): Promise<MediaModel>;
await slide.addVideo(source, options): Promise<MediaModel>;
slide.deleteMedia(shapeId): void;

document.media(slideIndex): readonly MediaModel[];
await document.addAudio(slideIndex, source, options): Promise<MediaModel>;
await document.addVideo(slideIndex, source, options): Promise<MediaModel>;
```

The document methods remain source-compatible and delegate to the selected `SlideModel`. `slide.media` is a filtered
view of the same objects returned by `slide.shapes`; it does not own another registry. A created model is exactly the
object returned by the next `slide.media`, `slide.shapes`, or `document.media()` read.

The codec's existing structural result becomes `MediaDescriptor`; its legacy `MediaModel` type name remains as a
deprecated type alias inside `@pptx/codecs`. The SDK explicitly exports the runtime model class from `@pptx/model`, so
the main package has one unambiguous `MediaModel` value and type. `shapeId` aliases inherited `id`, while
`slidePartUri` exposes the owning slide URI, preserving the previous document-level result fields.

### Replacement option contracts

```ts
interface ReplaceMediaSourceOptions {
  readonly contentType?: string;
  readonly fileName?: string;
  readonly transcode?: AddMediaOptions['transcode'];
}

interface ReplaceMediaPosterOptions {
  readonly contentType?: string;
  readonly fileName?: string;
}
```

Primary-source replacement accepts the same local bytes, data URI, path, Blob/File, stream, async iterable, and
HTTP/HTTPS external URL forms as creation. Poster replacement accepts the same embedded poster forms as creation;
HTTP/HTTPS poster URLs remain rejected. Calling `replacePoster()` without a source restores the built-in detached
default PNG. Replacement options are strict closed data objects and reuse the creation MIME, extension, byte-copy,
data-URI, stream, and transcode rules. Placement and metadata are deliberately not replacement options because those
already have live model setters.

## Strict Media-State Reader

Add one internal reader that receives a slide XML document, a direct `p:pic`, and current slide relationships. It
returns a detached `MediaState` only when the picture has a valid numeric `p:cNvPr/@id` and credible media evidence.
It never mutates or repairs imported XML during a read.

Classification priority is:

1. a direct relationship referenced from the picture whose standard type ends in `/audio` or `/video`;
2. an internal referenced target whose package content type has top-level `audio/` or `video/`;
3. a direct `a:audioFile` or `a:videoFile` marker;
4. no media classification, leaving the picture available to `ImageModel`.

Standard relationship evidence wins over a contradictory marker. Exact MIME lookup recognizes the strict creation
table plus imported aliases required for reading, including PptxGenJS `audio/mp3`. Unknown audio/video subtypes still
identify the kind from their top-level type but remain subject to replacement validation.

The reader resolves roles independently:

- kind relationship: the standard audio/video relationship when present, otherwise the marker's `r:link` target;
- Microsoft media relationship: `p14:media/@r:embed` when present;
- primary source: external standard relationship URL, otherwise an internal kind or Microsoft media target;
- poster: direct `a:blip/@r:embed` image relationship;
- playback preferences: only the library-owned private playback extension;
- name and alt text: direct `p:cNvPr@name` and optional `@descr`.

Malformed optional roles become absent state rather than causing the media picture to disappear. Mutators that need a
missing role either construct it canonically or fail before mutation when intent is ambiguous. Diagnostics may report
contradictory or dangling roles separately; ordinary reads preserve them losslessly.

This order deliberately reads three known PptxGenJS 4.0.1 defects:

- audio pictures rendered with `a:videoFile`;
- MP3 parts declared as `audio/mp3`;
- duplicated audio whose kind relationship was changed to the Microsoft media relationship type.

## Stable Identity and Reconciliation

`decodeShape()` invokes the media-state reader before its generic `p:pic` branch. It constructs `MediaModel` with the
shape id and inferred `audio | video` kind. `SlideModel.shapes` reuses the existing registry entry when id, constructor,
and kind match, exactly as for other semantic shapes.

The registry rules are:

- collection order always follows the current shape tree;
- duplicate slides receive new `SlideModel` owners and therefore distinct media handles even when payload targets are
  initially shared;
- moving a slide keeps its `SlideModel` and media handles;
- metadata, transform, playback, source, and poster edits keep the media handle;
- rollback keeps the handle and subsequent getters observe the restored package state;
- if raw mutation changes the same shape id from image to media, media to image, or audio to video, reconciliation
  replaces the semantic object because its class or kind changed;
- after deletion, an existing handle resolves through the normal missing-shape error path and never serves stale data.

Creation returns a structural descriptor from the codec only long enough to obtain `shapeId`; `SlideModel` immediately
reconciles that id through its own registry and returns the live `MediaModel`. It never returns the descriptor itself.

## Metadata, Transform, and Playback Editing

Name and alt-text setters validate strings with the same XML-control rules as creation, escape only during
serialization, and patch only the direct `p:cNvPr` attributes of the current picture. `altText = undefined` removes
`descr`; an empty string writes an explicit empty attribute. Both edits are synchronous transactions and preserve all
unknown children and attribute order outside the targeted attribute.

Transform reads and writes continue through `BaseShapeModel` and `SlideModel.setShapeTransform()`, so media inherits
the same EMU, rotation, flip, live-read, and rollback behavior as other shapes.

The `settings` setter validates a closed ordinary object using the creation rules: `play` is `click | auto`, booleans
remain strict, and volume is finite `0..1`. Supplied fields form the complete private preference state rather than a
partial merge; omitted fields take creation defaults when rendering. Assigning `undefined` removes only the
library-owned playback extension. It must not create, edit, or delete a native `p:timing` tree in this slice.

Imported native timing remains losslessly preserved and is not projected into `settings`. Public documentation and
diagnostics continue to distinguish stored preferences from native client playback commands.

## Primary Source Replacement

Source resolution, optional transcoding, MIME/extension decisions, and complete replacement definition finish before
the package transaction. The algorithm then operates on the two primary media roles as one logical unit:

1. resolve the current kind relationship, Microsoft media relationship, marker, and all primary targets;
2. allocate or reuse an embedded target by SHA-256 plus exact MIME, or prepare an external standard relationship;
3. ensure the picture has a canonical kind marker and standard kind relationship for its current model kind;
4. for embedded media, ensure `p14:media` and a Microsoft media relationship target the same new part;
5. for external media, remove the library-owned `p14:media` role and keep only the external standard relationship;
6. patch only the current picture's relationship references;
7. remove superseded relationships only when no slide XML node still references their ids;
8. garbage-collect superseded `/ppt/media` targets only after graph incoming reaches zero.

Embedded-to-embedded, embedded-to-external, external-to-embedded, and external-to-external transitions are supported.
Changing an audio model to video or vice versa is not a source replacement; callers create the new kind and remove the
old object, preserving the stable-kind identity contract.

### Clone-on-write and relationship isolation

An internal media target may be mutated in place only when all of the following hold:

- every package relationship pointing to the target is one of the current picture's primary roles;
- every primary relationship id being retained is referenced only by the current picture;
- final MIME and canonical extension are compatible with the existing part URI;
- the replacement is embedded and no content-hash dedup target already provides the result.

Otherwise replacement reuses a matching deduplicated target or allocates a sibling part. If a relationship id is
referenced by multiple XML nodes, a new relationship is allocated and only the current picture is patched. A
relationship id used solely by the current picture may be retargeted. When kind and Microsoft roles have separate ids,
both are treated independently even if they resolve to the same part.

This guarantees that duplicated slides and imported presentations with shared targets or shared relationship ids are
isolated without producing unnecessary parts for exclusive sources.

When replacement owns the primary roles, it canonicalizes those roles to standard `/audio` or `/video`, Microsoft
`/media`, and `a:audioFile` or `a:videoFile`. Unrelated XML and relationships, including contradictory opaque nodes not
referenced by the current picture, remain untouched. Merely reading or editing metadata never canonicalizes imported
PptxGenJS defects.

## Poster Replacement

Poster resolution completes before mutation and uses the strict PNG/JPEG/GIF table. The algorithm mirrors image
clone-on-write:

1. resolve the current direct `a:blip/@r:embed` relationship;
2. reuse a SHA-256 plus exact-MIME target or allocate a sibling `/ppt/media/poster*` part;
3. update the existing relationship only when its id is exclusive to this picture;
4. otherwise allocate a new image relationship and patch only this picture's blip;
5. remove a superseded relationship only when its id has no remaining XML references;
6. delete a superseded `/ppt/media` poster target only when graph incoming is zero.

Calling `replacePoster()` without a source follows the same path with the built-in default PNG; it does not assume the
existing default target is exclusive. Poster replacement never changes primary media roles, transform, metadata,
playback state, crop, or unrelated picture extensions.

## Object, Duplicate, Move, and Slide Lifecycle

`MediaModel.remove()` delegates to `SlideModel.deleteMedia(id)`. The codec removes the current picture, all
relationships referenced exclusively by that picture, and then any unreferenced `/ppt/media` primary or poster parts
in one transaction. A relationship id still referenced elsewhere is retained. Shared payloads remain until their last
incoming relationship disappears.

Slide duplication retains its current policy: XML relationship ids are preserved and audio, video, Microsoft media,
and image targets are shared. Source or poster replacement on either duplicate then separates only the edited object
through clone-on-write. Media models on source and duplicate are never the same JavaScript object.

Slide move changes only presentation order and preserves all media model identities and targets.

Before deleting a slide, `PresentationModel.deleteSlide()` captures internal `/ppt/media` targets referenced by that
slide's audio, video, Microsoft media, and poster image relationships. After removing the slide part and relationships,
it deletes each captured target only when current graph incoming is zero. It does not recursively traverse arbitrary
image targets, delete media outside `/ppt/media`, or treat media as an owned dependency subtree. This narrow sweep
reclaims library-created and conventional media payloads while preserving opaque external package structure.

## Atomicity and Error Handling

All async source reading and normalization happens before package mutation. Every synchronous metadata, playback,
relationship, payload, picture, delete, and garbage-collection change is contained in one package transaction.

Failures in source streams, data URIs, transcoders, option validation, MIME/extension resolution, XML parsing,
relationship allocation, part allocation, content-type updates, shape patching, or garbage collection leave parts,
relationships, content types, slide XML, graph state, model identity, and the mutation journal unchanged.

Replacement copies caller-owned byte containers before an asynchronous boundary and copies transcode output
immediately. It never fetches HTTP/HTTPS URLs. Aborted or rejected input cannot allocate a part or relationship.

Missing shape, missing slide, invalid kind, or ambiguous required relationship errors include the slide part URI and
shape id. A deleted model never silently re-creates its picture. Failed mutations do not poison the model handle; a
later valid operation may use the same handle after rollback.

## Compatibility and Preservation

- No-mutation open/write remains byte-stable under the existing package preservation contract.
- Metadata, transform, and playback edits patch only their owned XML regions.
- Replacement canonicalizes only the current picture's primary or poster roles that it must own to complete the edit.
- External URLs remain external and are never downloaded implicitly.
- Unknown picture extensions, native timing, hyperlinks, actions, and unrecognized relationships remain preserved.
- Existing document-level creation and listing signatures remain available; their return value becomes a richer live
  class without removing the prior readonly properties.
- Six supported presentation formats (`pptx`, `pptm`, `potx`, `potm`, `ppsx`, `ppsm`) receive identical media model
  behavior while retaining their macro/template/show package profile.

## Verification

Each implementation task ends with focused tests, typecheck, diff review, commit, push, fetch, and zero local/remote
divergence. This slice is complete only when all of the following pass:

1. strict reader tests cover native audio/video, external media, missing optional roles, contradictory markers,
   `audio/mp3`, and all three known PptxGenJS defects;
2. identity tests prove `===` across `slide.shapes`, `slide.media`, document listing, create, metadata/transform/playback
   edits, duplicate isolation, move, and transaction rollback;
3. metadata and playback tests cover strict validation, empty/undefined semantics, escaping, narrow XML mutation, and
   native timing preservation;
4. source replacement tests cover every embedded/external transition, seven creation MIME formats, extension changes,
   transcode, deduplication, exclusive in-place mutation, graph-shared targets, shared rIds, and rollback;
5. poster tests cover PNG/JPEG/GIF, default reset, deduplication, shared targets, shared rIds, unrelated-role isolation,
   and rollback;
6. object and slide deletion tests prove last-reference GC, duplicate retention, shared poster retention, external
   preservation, missing-handle errors, and complete transaction rollback;
7. public SDK tests cover zero-input creation, edit/reopen, stable return types, and all six package formats;
8. public PptxGenJS 4.0.1 fixtures prove import, noncanonical no-op preservation, narrow metadata edit, canonical source
   replacement, duplicate isolation, and delete behavior;
9. declarations, Node, browser, CLI, actual packed tarball, deterministic double-build, and gallery validation pass;
10. PowerPoint 2010 and LibreOffice open/render/save/reopen evidence is recorded, followed by README, API,
    compatibility, progress, and changelog updates.

## Deferred Follow-Up Designs

The following remain explicit later slices:

- native timing-tree create/read/edit for autoplay, click triggers, loop, hide-when-stopped, volume, trim, bookmarks,
  and media-command sequences;
- online/YouTube/Vimeo video creation and editing, plus explicit remote-fetch embedding;
- captions, subtitles, accessibility tracks, crop/rounding/shadow/hyperlink/placeholder media styles;
- built-in transcoding engines and runtime codec-availability promises;
- final Keynote, Google Slides, Windows PowerPoint, and broad real-world corpus certification.

These items stay visible in the compatibility matrix so a stable live model cannot be mistaken for full media or
PptxGenJS parity.
