# PptxGenJS Media Core Surface Family Design

Date: 2026-08-04

Status: confirmed under the user's autonomous-decision authorization

## Context

The declaration audit has fourteen unverified media atoms covering public media
kind selection, embedded and external sources, poster selection, explicit
extension hints, metadata, and frame geometry. Native already creates, opens,
edits, replaces, duplicates, removes, and reopens audio and video media through
typed live models in Node, installed-package, browser, and six-format OOXML
flows. This family binds that implementation to the declaration matrix; it does
not add a parallel media abstraction.

## Batch Boundary

The family contains exactly:

- `MediaProps.cover`, `data`, `extn`, `link`, `objectName`, `path`, and `type`;
- `MediaProps.x`, `y`, `w`, and `h`;
- the `audio`, `video`, and `online` members of `MediaType`.

It excludes inherited `DataOrPathProps.data/path`; playback/timing fields owned
by native `MediaSettings`; hyperlinks; placeholders; and image-only source,
sizing, and appearance atoms.

## Approach Decision

Three approaches were considered:

1. Copy PptxGenJS's permissive `MediaProps` bag and extension-driven MIME
   behavior into native. This would weaken source validation and editable
   package invariants.
2. Mark legal audio/video output as fully supported because representative
   packages match. This would hide API splitting, strict source detection,
   online-link handling, unit, error-timing, and legacy OOXML differences.
3. Reuse the existing typed media vertical slice and classify all fourteen
   atoms as deliberate differences. Runtime and OOXML probes confirm that
   PptxGenJS 4.0.1 represents `online/link` as an external video relationship
   plus an embedded poster. Native `addVideo(HTTP(S) URL)` creates the same core
   relationship and reopened media state, with stricter URL validation and
   richer canonical playback/timing state.

## Locked Classification

All fourteen atoms are deliberate differences:

- `data/path`: PptxGenJS selects between optional embedded source fields;
  native accepts one `MediaSource` and resolves or rejects it before mutation.
- `cover`: PptxGenJS requires or synthesizes a permissive cover string; native
  uses a typed `poster` source with content detection and a canonical default.
- `extn`: PptxGenJS lets the caller drive the emitted extension and can retain
  legacy MIME aliases; native derives or validates content type, file name, and
  extension as one consistent descriptor.
- `type` and the `audio/video/online` members: PptxGenJS uses a discriminator
  inside one `addMedia` call; native deliberately exposes typed `addAudio` and
  `addVideo` operations. The online branch maps to strict external HTTP(S)
  video input rather than a separate or nonexistent `addOnlineVideo` method.
- `link`: PptxGenJS accepts loosely coerced external targets and writes no
  playback state. Native accepts a strict HTTP(S) `MediaSource`, preserves the
  same external relationship and poster semantics, and adds editable native
  playback/timing state.
- `objectName`: native names the field `name`, validates XML-safe text, and
  exposes it on the live `MediaModel`.
- `x/y/w/h`: native uses `x/y/width/height` with explicit EMU or `inches()`
  units, strict transform validation, and live post-create editing.

## Evidence Architecture

- Runtime/control: reuse the valid embedded audio/video comparison, strict
  defect, append-without-rewrite, and legacy lifecycle controls; add one small
  aggregate control that proves the public `online/link` branch is structurally
  different from native external-file video.
- Native: map to `MediaSource`, `AddMediaOptions`, `MediaModel`,
  `SlideModel.addAudio/addVideo`, `PptxDocument.addAudio/addVideo`, and media
  source/descriptor normalization.
- Package/client: extend the actual npm tarball and persistent-browser media
  decks with one retained external online-video case. Each batch then proves
  its target, XML escaping, poster, transform, playback/timing, write/reopen,
  and expected portability diagnostics without a per-atom rebuild.
- OOXML: reuse the all-source/all-MIME/all-poster/external-mode SDK test,
  relationship-role checks, media timing checks, six-format reopen, and
  PowerPoint 2010 validation.
- Matrix: generate exactly fourteen frozen entries, correct the prior
  `addOnlineVideo` evidence label, and assert scope, exclusions, evidence, and
  totals.

Evidence agents remain read-only. Only the main agent edits shared files,
reviews, commits, and pushes.

## Acceptance Criteria

- Exactly fourteen atoms move from `unverified` to
  `deliberate-difference`; no other atom changes status.
- Existing classified media method entries may only change the inaccurate
  native API label from `addOnlineVideo` to the real external-source
  `addVideo` mapping.
- Global totals become 620 supported, 339 deliberate-difference, 91
  deprecated-alias, 359 defect-excluded, and 365 unverified: 1,409 of 1,774
  atoms classified (79.43%).
- Focused media tests run once; audit tests, TypeScript, deterministic
  generation, exact migration review, and `git diff --check` pass.
- One actual npm tarball run, one persistent Chrome session, one OOXML
  inspection batch, and one full Vitest run pass at this family boundary.
- Review produces one capability-family commit and push with no unrelated
  tracked or staged files.
