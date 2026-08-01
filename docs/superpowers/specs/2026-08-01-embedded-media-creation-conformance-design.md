# Embedded Media Creation Conformance Design

## Goal

Complete one independently reviewable media slice: strict embedded audio/video creation through the existing
`PptxDocument.addAudio()` / `addVideo()` APIs, including canonical data-URI sources, caller-visible names and alt text,
poster/cover input, extension hints, atomic failure, and PptxGenJS 4.0.1 public-output conformance.

This slice strengthens zero-input creation without claiming the later media lifecycle is complete. Online video,
stable live media editing, native timing-tree playback commands, full duplicate/delete isolation, and broader release
evidence remain separate follow-up designs.

## Current-State Audit

The existing `MediaCodec` already provides useful infrastructure:

- embedded audio/video parts and the paired Office `audio|video` plus Microsoft media relationships;
- external HTTP/HTTPS relationships;
- poster parts and picture markup;
- path, `Uint8Array`, `ArrayBuffer`, Blob/File, Web Stream, and async-iterable sources;
- content-hash reuse, deletion with reference-aware media garbage collection, diagnostics, and transaction rollback;
- SDK `addAudio()`, `addVideo()`, and `media()` entry points.

The current creation contract is not yet ready to count as PptxGenJS parity:

- `AddMediaOptions` is not closed or descriptor-safe; invalid coordinates, playback values, MIME values, extensions,
  poster inputs, accessors, inherited keys, and unknown fields are not normalized before source I/O;
- base64 media and cover data URIs are treated as local paths;
- empty payloads are accepted;
- object names are hard-coded from kind and shape id, cannot be supplied by the caller, and are not escaped through
  the shared XML attribute helper;
- alt text cannot be created;
- poster MIME/extension decisions and transcode results are not validated as one detached definition;
- there is no public PptxGenJS `data`/`path`/`cover`/`extn`/`objectName` conformance evidence;
- the test surface is two codec tests and does not cover SDK, browser, declarations, six formats, or a packed consumer.

The public PptxGenJS 4.0.1 probe establishes the comparison baseline:

- `MediaProps.type` is `audio | video | online`;
- embedded audio/video accepts `data` or `path`, `cover`, `extn`, `objectName`, and `x/y/w/h`;
- every embedded item creates one media part, one poster PNG, a kind relationship, a Microsoft media relationship,
  click-to-play action, and the Office 2010 media extension;
- `objectName` is escaped into `p:cNvPr@name`;
- `extn` chooses the media part extension;
- PptxGenJS writes `audio/mp3` for an MP3 data URI and incorrectly renders `a:videoFile` even for audio. Native will
  retain canonical `audio/mpeg` and `a:audioFile` rather than reproduce those defects.

## Approaches Considered

### Thin option-alias layer

Add `objectName`, `cover`, and `extn` aliases directly to the current codec and call the existing implementation.
This is quick but preserves unsafe input handling, data-URI failure, ambiguous MIME/extension state, and weak rollback
coverage. It would make later hardening a breaking rewrite.

### One-step full media model rewrite

Replace `MediaCodec` with a stable shape-backed model covering create/read/edit/delete/duplicate, timing, online media,
and all sources at once. This reaches the eventual architecture sooner but produces an oversized review surface and
mixes independent compatibility risks.

### Incremental vertical slices — selected

Harden embedded creation first, then build stable lifecycle, timing, online video, and release evidence as separate
plans. This preserves the useful codec, creates a strict boundary before adding more behavior, and keeps every commit
independently testable and revertible.

## Public API

The existing high-level functions remain canonical:

```ts
await document.addAudio(slideIndex, source, options);
await document.addVideo(slideIndex, source, options);
```

`MediaSource` continues to accept path, HTTP/HTTPS URL, `Uint8Array`, `ArrayBuffer`, Blob/File, Web Stream, and async
iterable. This slice additionally accepts strict base64 data URIs whose declared top-level type matches the requested
kind. HTTP/HTTPS remains an explicit external relationship; it is not fetched. Remote embedding and online-video
semantics are later slices.

`AddMediaOptions` gains two native object fields while retaining all current fields:

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

PptxGenJS field mapping is semantic rather than an alias facade:

| PptxGenJS 4.0.1 | Native API |
| --- | --- |
| `type: 'audio' | 'video'` | `addAudio()` / `addVideo()` |
| `data` / `path` | `MediaSource` |
| `cover` | `poster` |
| `extn` | `fileName` extension hint and/or canonical `contentType` |
| `objectName` | `name` |
| `x/y/w/h` inches | native `x/y/width/height` EMU, normally via `inches()` |

Omitted names use `Media N`, where `N` is the count of existing `p:pic` elements whose `p:nvPr` contains an
`a:audioFile` or `a:videoFile` on the target slide. This matches PptxGenJS media-count semantics without depending on
shape ids. Omitted alt text stays absent; explicit empty name or alt text remains a direct empty attribute.

## Strict Normalization

Create a focused internal normalizer in `packages/codecs/src/media-create.internal.ts`. It receives `kind`, raw
options, resolved media bytes/relationship state, and resolved poster state, then returns a deeply detached frozen
definition. Plain containers are recursively frozen; non-empty typed arrays are copied into private storage and
treated as immutable because JavaScript cannot freeze their indexed elements. Source I/O stays asynchronous; package
mutation starts only after the complete definition exists.

Options must be an ordinary or null-prototype object containing only own data properties from the public interface.
Custom-prototype objects, enumerable inherited fields, accessors, symbol keys, unknown fields, arrays, class instances,
and primitives are rejected without reading unsafe getters or consuming a source.

Field rules:

- `name` and `altText` are strings without illegal XML control characters; attribute metacharacters are escaped only
  at render time;
- x/y are signed safe-integer EMU; width/height are positive safe-integer EMU;
- `play` is `click | auto`; loop/hideWhenStopped are booleans; volume is finite `0..1` and is not clamped;
- `contentType` and `posterContentType` must be exact lowercase members of the supported MIME table below, without
  parameters; the media MIME top-level type must match audio or video, while poster MIME must be `image/png`,
  `image/jpeg`, or `image/gif`;
- `fileName` is a non-empty string used only as an extension hint, never serialized into XML;
- `transcode` is a function. Its result must be an ordinary object with non-empty `Uint8Array` bytes, a compatible
  MIME string, and an optional safe extension;
- every supplied byte payload is copied before any asynchronous boundary can expose it to mutation; transcoder output
  bytes are copied immediately after its promise resolves.

## Source and Extension Resolution

Add one strict base64 data-URI parser shared by primary media and poster resolution. It accepts exactly one comma,
a supported MIME declaration, the `;base64` flag, canonical base64 alphabet/padding, and a non-empty decoded payload.
Whitespace, percent encoding, malformed padding, mismatched kind, and trailing garbage are rejected.

Resolution priority is deterministic:

1. explicit `contentType` asserts or supplies the media MIME;
2. a data URI supplies a declared MIME that must agree with the assertion;
3. a safe extension from `fileName`, File name, or path maps to a known MIME;
4. otherwise the current kind default remains `audio/mpeg` or `video/mp4`.

The part extension comes from `fileName`/source name when it safely matches the final MIME; otherwise it comes from the
canonical MIME map. A mismatch is rejected rather than producing a misleading extension. Supported creation MIME and
extensions in this slice are:

- audio: `audio/mpeg` `.mp3`, `audio/mp4` `.m4a`, `audio/wav` `.wav`, `audio/ogg` `.ogg`;
- video: `video/mp4` `.mp4` or `.m4v`, `video/quicktime` `.mov`, `video/webm` `.webm`;
- poster: PNG `.png`, JPEG `.jpg|.jpeg`, GIF `.gif`.

HTTP/HTTPS sources stay external and do not allocate a media part. A poster URL remains rejected in this slice rather
than silently replaced. Omitted poster uses the existing detached built-in PNG.

## Canonical OOXML

Embedded audio and video keep the proven three-relationship picture:

- direct `a:audioFile` or `a:videoFile` points through the kind relationship;
- `p14:media` in fixed extension URI `{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}` points through the Microsoft media
  relationship;
- direct blip points through the poster image relationship;
- `p:cNvPr` owns escaped name, optional escaped `descr`, and the media click action;
- `p:cNvPicPr/a:picLocks@noChangeAspect="1"`, rectangular geometry, and stretch fill remain canonical;
- transform uses the fully normalized EMU definition.

The current private playback extension remains byte-compatible in this slice; native timing is designed separately.
No code in this slice claims that autoplay/loop/hide/volume is already a native timing-tree command.

## Atomicity and Deduplication

Media source, optional transcode output, poster source, MIME/extension resolution, options, and complete XML definition
must all finish before the package transaction. The transaction owns media/poster part insertion, all relationships,
shape id allocation, shape-tree insertion, and the mutation journal.

Any source, data URI, poster, transcode, option, MIME, extension, XML, relationship, allocation, or shape-tree failure
must leave parts, content types, relationships, graph, slide XML, ZIP output, shape ids, and mutation journal unchanged.

Hash deduplication remains content-type-aware. Same payload plus same content type may share a media or poster target;
same bytes with a different content type must not. Creation never deletes a pre-existing shared target.

## PptxGenJS Conformance and Intentional Divergences

Generate comparison files only through public PptxGenJS 4.0.1 APIs. Cover at least:

1. audio data URI + explicit cover + extn + objectName + transform;
2. video data URI + explicit cover + extn + objectName + transform;
3. audio local path with inferred extension;
4. duplicate embedded payload reuse;
5. invalid missing source, malformed media data URI, malformed cover, mismatched type, and invalid runtime options.

For supported valid cases compare final picture order, name, transform, click action, poster bytes, media bytes, part
extensions, relationship roles/modes/targets, content types, Office media extension, strict reopen, and validation.

Record these strict native divergences:

- native audio uses `a:audioFile`; PptxGenJS 4.0.1 writes `a:videoFile` even when the relationship is audio;
- native MP3 content type is `audio/mpeg`; PptxGenJS writes noncanonical `audio/mp3`;
- native rejects unsafe/falsy/coerced values before mutation instead of logging and continuing;
- native accepts portable byte/Blob/stream sources directly and does not require callers to construct PptxGenJS data
  strings.

## Verification

Each implementation task ends with focused tests, typecheck, diff review, commit, push, fetch, and zero divergence.
The slice is complete only when all of the following pass:

- codec normalization, data-URI, renderer, rollback, deduplication, and deletion-regression tests;
- SDK zero-input and six-format write/reopen tests;
- PptxGenJS public-output conformance;
- packed Node/browser/declaration/CLI smoke for data URI, Blob/stream, poster, names, and relationship roles;
- deterministic double build and an actual-tarball gallery validated against PowerPoint 2010;
- LibreOffice open/render/save/reopen comparison with any relationship/content-type normalization recorded;
- public README/API/compatibility/progress/changelog updates.

## Deferred Follow-Up Designs

The following are deliberately outside this first media slice:

- online/YouTube video and explicit remote-fetch embedding;
- stable live `MediaModel` identity, name/alt-text/transform/settings/poster/media replacement, and model-level entry;
- complete duplicate/move/delete editing isolation across shared media/poster relationships;
- native timing-tree create/read/edit for play mode, loop, hide-when-stopped, volume, trim, bookmark, and triggers;
- transcoding implementations or codec-availability promises;
- captions, subtitles, accessibility tracks, crop/rounding/shadow/hyperlink/placeholder styles;
- Keynote, Google Slides, and Windows PowerPoint corpus certification beyond the existing validation profiles.

These deferrals remain visible in the compatibility matrix so this slice cannot be mistaken for complete media parity.
