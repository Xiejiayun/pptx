# Native Media Timing Design

## Goal

Complete one independently reviewable media slice: make `MediaModel.settings` create, read, edit, and clear native
PowerPoint timing-tree playback for embedded and external audio/video while keeping the existing private playback
extension as a stable library preference and ownership record.

The result must make presentations created from zero behave in PowerPoint without installing the animations plugin,
must edit recognized PowerPoint media timing without damaging ordinary animations, and must keep create, duplicate,
move, delete, source replacement, no-op, and rollback behavior reference-safe.

This slice covers the current four public playback controls: click/automatic start, loop, hide while stopped, and
volume. It does not claim trimming, bookmarks, captions, narration, cross-slide audio, arbitrary media commands, or a
general animation rewrite.

## Evidence and Current-State Audit

The current implementation has two disconnected representations:

- `MediaCodec` always writes the complete private `px:playback` preference with `play`, `loop`,
  `hideWhenStopped`, and `volume`;
- `MediaModel.settings` reads and edits only that private extension and deliberately preserves `p:timing`;
- the optional animations plugin scans private preferences once during installation and appends a simplified
  `p:audio` or `p:video` node;
- later media setting edits do not update the materialized node, clearing settings does not remove it, and media
  deletion leaves its shape targets dangling;
- the plugin neither represents `hideWhenStopped` nor owns a complete PowerPoint media-command graph;
- its duplicate check treats any `cMediaNode` for the same shape as equivalent, even when kind, trigger, loop,
  volume, or structure differs.

PptxGenJS 4.0.1 exposes no autoplay, loop, hide, or volume fields on `MediaProps` and emits no timing tree. This work
therefore exceeds its current public media input surface while retaining PptxGenJS import compatibility.

The ISO/IEC 29500 model establishes that `p:cMediaNode` owns `vol`, `mute`, `numSld`, and
`showWhenStopped`, plus one `p:cTn` and one target. `showWhenStopped` is the inverse of the public
`hideWhenStopped` setting. Looping is a time-node repeat, not a media-node attribute.

Public Apache POI `EmbeddedAudio.pptx` and `EmbeddedVideo.pptx` fixtures produced by PowerPoint provide the concrete
client grammar:

- the slide has one `tmRoot` and a `mainSeq` under its direct `childTnLst`;
- the actual media state is a direct root child `p:audio` or `p:video` containing `p:cMediaNode` and a shape target;
- playback is activated by a `presetClass="mediacall"` command node containing `cmd="playFrom(0.0)"` and a
  behavior target for the same shape;
- video also has an `interactiveSeq` with a shape-click condition and `cmd="togglePause"`;
- the media time node starts from `delay="indefinite"`, while its command placement and start condition determine
  when playback begins;
- the audio node has the PowerPoint `onStopAudio` end condition and the sample volume is `80000`.

The plugin's current two-`cTn` encoding is schema-shaped but not client-canonical and is not a safe foundation for
editable PowerPoint playback.

## Approaches Considered

### Keep one-time plugin materialization

Continue storing preferences in the private extension and require callers to install the animations plugin before
writing. This keeps the built-in codec small, but output behavior depends on plugin installation order and becomes
stale after the first edit. It also cannot make `MediaModel.settings` a truthful live editing surface.

### Make the animations plugin the live owner

Route every media mutation through the optional plugin. This creates an optional runtime dependency in the core model,
leaves zero-plugin output incomplete, and gives two packages overlapping responsibility for one transaction.

### Shared built-in media timing codec — selected

Add a media-specific timing module to `@pptx/codecs`. `MediaCodec` calls it inside the same package transaction for
creation, settings edits, clearing, and deletion. The animations plugin delegates its legacy materialization entry
point to this module rather than maintaining another encoder. General effect APIs remain in the plugin.

This produces complete core output, a single ownership policy, one ID allocator, and one preservation strategy without
turning `MediaCodec` into a general animation API.

## Two-Layer State Model

Playback has two coordinated layers:

1. the private `px:playback` extension is the exact, stable library preference and records owned native timing IDs;
2. native `p:timing` is the PowerPoint client behavior graph.

For library-created or adopted media, the private layer is authoritative. The native layer is regenerated only from a
strict normalized preference. This avoids inferring public intent from every possible animation grammar on each read.

For imported media without a private extension, the reader may project settings from one strictly recognized native
PowerPoint graph. It never invents settings from a partial or ambiguous graph.

The private extension remains at its existing fixed URI and preserves the four current attributes byte-for-byte when
no native ownership has been assigned. Once synchronized, it also records a versioned native ownership tuple:

```xml
<px:playback xmlns:px="urn:pptx-ooxml:media"
  play="click" loop="0" hideWhenStopped="0" volume="100000"
  nativeVersion="1" mediaTnId="7" playTnId="5" pauseTnId="11"/>
```

`pauseTnId` is optional for audio or imported graphs without a recognized pause branch. The attributes point to
`p:cTn/@id` values, not element offsets. Unknown attributes and children on an imported private extension are not an
ownership signal.

Ownership recovery is allowed only when the extension has `nativeVersion="1"`, all recorded nodes still exist, and
the complete branch grammar and shape targets match. If a client renumbers timing IDs, recovery may replace stale IDs
only when exactly one complete canonical graph for that media shape matches the stored preference. Zero or multiple
matches are not adopted silently.

## Public Semantics

The public types do not change:

```ts
interface MediaPlaybackSettings {
  readonly play?: 'click' | 'auto';
  readonly loop?: boolean;
  readonly hideWhenStopped?: boolean;
  readonly volume?: number;
}
```

Creation continues to normalize omitted fields to `click`, `false`, `false`, and `1`, but now emits both layers before
returning the live model. A created presentation therefore has native click playback even when the caller supplies no
playback options.

`media.settings` resolves in this order:

1. a valid private preference, whether or not its owned graph is currently healthy;
2. one complete recognized imported native graph;
3. an empty frozen object.

The getter never mutates, repairs, or adopts. Returned values remain detached and frozen.

Assigning a settings object is replacement, not merge. It normalizes the complete four-field preference, then:

- updates an existing healthy owned graph in place where narrow attribute patches are sufficient;
- otherwise replaces the complete owned graph and records its new IDs;
- adopts and replaces exactly one recognized imported graph when no private preference exists;
- rejects an ambiguous or malformed media graph before mutation instead of adding competing playback commands.

Assigning `undefined` removes the private preference and all branches proven to be owned. For imported media without a
private extension, it removes exactly one complete recognized graph. Unknown, partial, or ambiguous native media
timing is preserved and causes a narrow ownership error. After a successful clear, `settings` returns `{}`.

No-op assignment compares the normalized preference and verified native graph. It writes nothing only when both layers
already agree. Matching private XML with missing, stale, or contradictory native timing is a repair, not a no-op.

## Recognized Native Graph

The media timing reader operates only below the slide's direct `p:timing/p:tnLst` root. It requires exactly one
`p:par/p:cTn[@nodeType="tmRoot"]` and its direct `p:childTnLst`. Descendant-wide searches are never used to claim
ownership.

A media-state branch is recognized when all of the following hold:

- one direct root child is `p:audio` or `p:video` matching the media picture kind;
- it contains exactly one direct `p:cMediaNode`, one direct `p:cTn`, and one direct target;
- the target is exactly one `p:spTgt/@spid` matching the media picture;
- the media time-node ID is a unique positive unsigned integer in the slide timing tree;
- `vol`, `mute`, `showWhenStopped`, and repeat tokens are valid when present;
- it does not use `numSld`, sound targets, slide targets, nested media, unsupported repeat duration, trimming, or
  unknown media-node children;
- exactly one compatible `playFrom(0.0)` command branch targets the same shape;
- an optional `togglePause` branch, when present, has one shape-click trigger and the same target.

Click versus automatic start is decoded from the recognized play-command container and its condition grammar, not
from the presence of the picture hyperlink action. The implementation accepts the exact PowerPoint main-sequence
forms captured in fixtures and the library's canonical form. It does not guess from arbitrary `par` nesting.

`loop` is true only for `repeatCount="indefinite"` on the media time node. A finite repeat is outside the public
contract and prevents adoption. `volume` is `0` for `mute="1"`; otherwise valid `vol` is divided by `100000`, using
`100000` when absent. `hideWhenStopped` is the inverse of valid `showWhenStopped`; absence uses PowerPoint's visible
when stopped behavior.

The reader returns separate states: `absent`, `recognized-imported`, `owned-healthy`, `owned-stale`,
`unsupported`, or `ambiguous`. Diagnostics and mutators use this state directly so a target-shaped coincidence cannot
be mistaken for ownership.

## Canonical Encoding

When a slide has no timing tree, the codec creates the minimal PowerPoint root with `tmRoot`, `mainSeq`, previous/next
slide conditions, and direct child lists in schema order. When a valid root already exists, it appends only inside the
correct direct lists. It never rebuilds an imported root to improve formatting.

Each owned media graph contains:

- a direct root `p:audio` or `p:video` state node with `p:cMediaNode`;
- a media `p:cTn` with `fill="hold"`, `display="0"`, a canonical start condition, and the media shape target;
- canonical `vol` and `showWhenStopped`; loop uses `repeatCount="indefinite"` only when enabled;
- the audio `onStopAudio` end condition required by the PowerPoint grammar;
- one `presetClass="mediacall"`, `cmd="playFrom(0.0)"` command group placed in the recognized click or automatic
  sequence form;
- for video, one canonical `interactiveSeq` containing a `togglePause` command for clicks on the video shape.

The existing `a:hlinkClick action="ppaction://media"` remains on the picture. It is a media action affordance and is
not used as the ownership marker.

All new `p:cTn/@id` values come from a slide-wide allocator over valid existing IDs. The allocator never reuses an ID
within a mutation, never changes imported IDs, rejects exhaustion, and writes IDs only after the complete mutation has
been planned. Root creation reserves its IDs first, then media graph IDs in deterministic order.

XML insertion follows slide schema order: `p:timing` is placed after `p:transition` when present and before slide-level
`p:extLst`; root children are added without moving opaque siblings. Removing the last owned graph removes empty owned
command wrappers. It removes the entire `p:timing` element only when the codec created the root and no imported or
general animation content remains. Imported empty containers are preserved.

## Mutation Isolation

### Creation and settings edits

Media picture insertion, relationships, private preference, native graph, and ownership IDs are one transaction. A
failure in timing parsing, ID allocation, branch insertion, or extension update rolls back the entire media creation.

Settings edits plan and validate both layers before the first patch. Updating volume or hide may patch only owned
attributes when the graph is healthy; trigger or kind-sensitive changes may replace the owned command group. Unknown
siblings, whitespace, attributes, animations, transitions, and slide extensions remain byte-preserved outside the
owned spans.

### Source replacement

Embedded/external replacement does not change timing IDs or command branches. Kind is stable for a live
`MediaModel`, so audio remains `p:audio` and video remains `p:video`. If replacement discovers a stale owned graph, it
does not opportunistically repair it because source replacement has no playback intent.

### Duplicate and move

Slide duplication copies slide XML, so timing shape targets and timing IDs remain valid inside the duplicate. Its
private ownership tuple remains local to the new slide and subsequent settings edits isolate naturally because slide
parts are distinct. No timing ID needs global uniqueness across slides.

Slide move changes presentation order only and preserves timing unchanged. Stable media handles continue to resolve
through their original slide part.

### Object and slide deletion

Deleting a media object first resolves its timing state. Healthy owned branches are removed with the picture. A
recognized imported graph may also be removed because its only target is the deleted shape. Unsupported or ambiguous
branches are not guessed; deletion removes target branches only when the complete direct graph is structurally proven,
otherwise it rejects before mutating so it cannot leave a knowingly dangling target.

Deleting a whole slide deletes its timing with the slide part and requires no separate retargeting. Media payload
garbage collection remains relationship-based and independent of timing.

## Animations Plugin Integration

`AnimationTimingCodec.materializeMediaPlayback()` remains for backward compatibility with presentations written by
older library versions. It delegates to the shared media timing module for each direct media picture and returns the
number of graphs created or repaired.

Plugin installation becomes idempotent:

- current library output is already healthy and produces zero changes;
- legacy private-only output receives a canonical native graph;
- recognized imported native output without a private preference is preserved;
- unsupported or ambiguous imported timing is preserved and reported, not shadowed by a second graph.

General `tree()`, `add()`, `remove()`, `retargetShape()`, and animation validation remain plugin features. Their
existing broad behavior is not expanded in this slice, but the plugin must use the shared slide-wide timing ID
allocator so general effects and media nodes cannot collide.

## Diagnostics

The current `MEDIA_PLAYBACK_TIMING_EXTENSION` information diagnostic is retired for healthy native graphs. New
diagnostics distinguish actionable states:

- `MEDIA_TIMING_MISSING`: valid private preference has no native graph;
- `MEDIA_TIMING_STALE`: recorded ownership IDs or graph values disagree with the preference;
- `MEDIA_TIMING_UNSUPPORTED`: a target graph exists but uses unsupported media semantics;
- `MEDIA_TIMING_AMBIGUOUS`: multiple candidate graphs or duplicate IDs prevent safe ownership;
- `MEDIA_TIMING_DANGLING_TARGET`: a media timing target has no matching media picture;
- `MEDIA_TIMING_KIND_MISMATCH`: audio/video timing kind contradicts the picture kind.

Diagnostics never repair. Profile-specific autoplay and codec portability warnings remain unchanged.

## Atomicity and Error Handling

Every media timing mutation runs inside the surrounding `OpcPackage.transaction()`. The codec validates:

- the direct slide/timing/root/container structure;
- media picture identity and kind;
- strict playback preference data;
- every timing ID and target;
- all branches it intends to own or remove;
- final ID capacity and deterministic insertion points.

Only then does it patch XML. Any failure leaves slide XML, relationships, media parts, private preferences, timing IDs,
model identity, graph state, and the mutation journal unchanged.

Malformed imported timing is never normalized as a side effect of reading, metadata edits, transform edits, source or
poster replacement, or package write. Error messages include slide URI and media shape ID without serializing caller
data or media URLs.

## Compatibility and Preservation

- No-mutation open/write remains byte-stable.
- Existing private-only files remain readable and can be materialized safely.
- PowerPoint-style imported media graphs are readable and editable when they match the strict grammar.
- General animations and unknown timing branches remain untouched.
- PptxGenJS media without timing remains readable; a settings assignment or legacy materialization adds canonical
  native playback without changing its media payload.
- All six presentation package formats receive the same slide timing behavior while retaining macro/template/show
  profiles.
- The codec does not fetch external media, inspect media duration, or invent finite command duration from payload
  bytes; canonical commands use timing values that do not require a decoder.

## Verification

Each implementation task ends with focused tests, review, commit, push, fetch, and zero local/remote divergence. This
slice is complete only when all of the following pass:

1. unit fixtures lock exact PowerPoint audio/video graph recognition, schema ordering, ID allocation, click/auto,
   loop, hide, volume, mute, and imported defaults;
2. creation tests prove default and customized audio/video emit coherent private and native layers atomically;
3. live settings tests cover no-op, narrow updates, trigger replacement, clearing, repair, strict validation, reopen,
   and rollback;
4. malformed tests cover duplicate IDs, repeated roots/lists/nodes/targets, dangling targets, kind mismatch,
   unsupported finite repeats, cross-slide audio, sound targets, and ambiguous command graphs;
5. preservation tests combine ordinary animations, transitions, opaque timing nodes, whitespace, and slide extensions
   and prove mutation isolation;
6. lifecycle tests cover duplicate-slide isolation, move stability, source/poster replacement, media deletion, whole
   slide deletion, and failed deletion rollback;
7. plugin tests prove legacy materialization, healthy zero-change installation, shared ID allocation, and no duplicate
   graphs;
8. public SDK, declarations, Node, browser, CLI, installed tarball, deterministic build, and all six package formats
   cover native timing output;
9. the complete suite remains green and compatibility/progress/API/plugin documentation is updated;
10. generated galleries are validated, opened, rendered, saved, and reopened through PowerPoint and LibreOffice, with
    exact timing normalization recorded before the slice is declared complete.

## Deferred Follow-Up Designs

The following remain visible later slices:

- trim start/end, bookmarks, seek/pause/stop commands, finite repeats, narration, and cross-slide audio;
- captions, subtitles, accessibility tracks, and online-video timing;
- arbitrary animation sequence editing beyond the existing plugin surface;
- media crop, rounding, shadow, hyperlink, and placeholder styles;
- built-in transcoding and broad Windows PowerPoint, Keynote, and Google Slides certification.

These items are not implied by supporting the four current `MediaPlaybackSettings` fields.
