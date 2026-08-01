# Native Media Timing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the built-in media API create, read, edit, clear, diagnose, and lifecycle-manage native PowerPoint
timing for click/automatic start, loop, hide-when-stopped, and volume.

**Architecture:** Add a strict native timing reader and canonical editor inside `@pptx/codecs`, then make
`MediaCodec` synchronize the private playback preference and native timing graph in one transaction. The animations
plugin reuses the shared codec for legacy materialization while retaining general animation APIs.

**Tech Stack:** TypeScript 5.8, Vitest, lossless OOXML patching, OPC package transactions, JSZip fixtures,
PptxGenJS 4.0.1 conformance, Playwright, pptx-inspect, PowerPoint, and LibreOffice.

## Global Constraints

- Execute inline in the current task; do not delegate implementation or review.
- Preserve no-mutation byte stability and never canonicalize imported timing during reads.
- Keep the private `px:playback` extension as the exact preference and versioned native ownership record.
- Recognize and mutate only direct, complete media timing graphs; preserve ordinary and unknown animation branches.
- Keep `MediaPlaybackSettings` unchanged: `play`, `loop`, `hideWhenStopped`, and `volume` only.
- Do not fetch external media, decode duration, add a runtime dependency, or claim trim/bookmark/caption support.
- Apply identical slide timing behavior to `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`.
- Keep every create/edit/clear/delete mutation atomic under `OpcPackage.transaction()`.
- Use fresh `/tmp/pptx-native-media-timing-*` directories for downloaded samples, tarballs, galleries, renders, and
  client round trips; never commit those artifacts or `.pnpm-store/`.
- End every task with focused tests, self-review, one commit, push, remote verification, and divergence `0 0`.

---

## File Map

- Create `packages/codecs/src/media-timing-state.internal.ts`: strict direct-structure native timing reader,
  ownership/status model, PowerPoint setting projection, and slide-wide ID allocation.
- Create `packages/codecs/src/media-timing-state.internal.test.ts`: exact PowerPoint audio/video fixtures and malformed
  state coverage.
- Create `packages/codecs/src/media-timing-edit.internal.ts`: canonical root/command/media encoding plus safe
  synchronize/clear operations.
- Create `packages/codecs/src/media-timing-edit.internal.test.ts`: exact XML, insertion order, no-op, preservation,
  recovery, and removal coverage.
- Modify `packages/codecs/src/media-edit.internal.ts`: read and render native ownership attributes beside the existing
  four playback fields.
- Modify `packages/codecs/src/media-state.internal.ts`: project one recognized imported native graph only when no
  valid private preference exists.
- Modify `packages/codecs/src/media.ts`: synchronize native timing during creation/settings/deletion and emit precise
  diagnostics.
- Modify `plugins/animations/src/index.ts`: delegate legacy media materialization and share timing ID allocation.
- Extend codec, model, SDK, adapter, package, packed Node, browser, and client verification surfaces without adding a
  second public timing API.

---

### Task 1: Strict native timing state and ID allocation

**Files:**

- Create: `packages/codecs/src/media-timing-state.internal.ts`
- Create: `packages/codecs/src/media-timing-state.internal.test.ts`

**Interfaces:**

- Consumes: `LosslessXmlDocument`, direct `XmlElement` traversal, media kind, shape ID, and optional recorded
  ownership.
- Produces:

```ts
export interface NativeMediaTimingOwnership {
  readonly version: 1;
  readonly mediaTnId: number;
  readonly playTnId: number;
  readonly pauseTnId?: number;
}

export interface NativeMediaTimingSettings {
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
}

export type NativeMediaTimingStatus =
  | 'absent'
  | 'recognized-imported'
  | 'owned-healthy'
  | 'owned-stale'
  | 'unsupported'
  | 'ambiguous';

export interface NativeMediaTimingState {
  readonly status: NativeMediaTimingStatus;
  readonly settings?: Readonly<NativeMediaTimingSettings>;
  readonly ownership?: Readonly<NativeMediaTimingOwnership>;
  readonly reason?: string;
}

export function readNativeMediaTiming(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: 'audio' | 'video',
  recorded?: Readonly<NativeMediaTimingOwnership>,
): Readonly<NativeMediaTimingState>;

export function allocateNativeTimingIds(
  xml: LosslessXmlDocument,
  count: number,
): readonly number[];
```

- [ ] **Step 1: Add exact compact PowerPoint fixtures**

Copy the timing fragments from the inspected Apache POI `EmbeddedAudio.pptx` and `EmbeddedVideo.pptx` samples into
test constants. Keep the `tmRoot`, `mainSeq`, `playFrom(0.0)`, `cMediaNode`, audio `onStopAudio`, video
`interactiveSeq`, and `togglePause` branches intact, with shape IDs and timing IDs unchanged.

```ts
const POWERPOINT_AUDIO = [
  '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
  '<p:cSld><p:spTree>', mediaPicture(4, 'audio'), '</p:spTree></p:cSld>',
  '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">',
  '<p:childTnLst>', powerPointMainSequence(4, 5, 6),
  '<p:audio><p:cMediaNode vol="80000"><p:cTn id="7" fill="hold" display="0">',
  '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:endCondLst>',
  '<p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>',
  '</p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="4"/></p:tgtEl>',
  '</p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>',
].join('');
```

Define `mediaPicture()` and `powerPointMainSequence()` in the test file as deterministic fixture renderers containing
the exact inspected picture and command XML; neither helper calls production code.

- [ ] **Step 2: Write failing reader and allocator tests**

Require PowerPoint audio/video to decode as `recognized-imported`, with click start, loop false, hide false, volume
`0.8`, exact media/play/pause IDs, and no XML mutation. Require matching recorded ownership to become
`owned-healthy`; wrong IDs become `owned-stale`.

```ts
const xml = LosslessXmlDocument.parse(POWERPOINT_VIDEO);
const before = xml.serialize();
expect(readNativeMediaTiming(xml, 2, 'video')).toMatchObject({
  status: 'recognized-imported',
  settings: { play: 'click', loop: false, hideWhenStopped: false, volume: 0.8 },
  ownership: { version: 1, mediaTnId: 7, playTnId: 5, pauseTnId: 11 },
});
expect(xml.serialize()).toBe(before);
```

Cover absent timing, automatic canonical conditions, indefinite loop, mute, `showWhenStopped`, kind mismatch,
`numSld`, finite repeat, repeated roots/lists/media nodes/targets, sound targets, duplicate timing IDs, dangling shape
targets, multiple compatible graphs, and unrelated ordinary animations.

For allocation, require deterministic IDs above the maximum valid `p:cTn/@id`, rejection of zero/negative/decimal/
overflow existing IDs, count validation, and exhaustion before mutation.

- [ ] **Step 3: Run the focused test and confirm red state**

```bash
pnpm exec vitest run packages/codecs/src/media-timing-state.internal.test.ts
```

Expected: FAIL because the module and exported reader do not exist.

- [ ] **Step 4: Implement direct-structure recognition**

Use direct-child helpers for `timing`, `tnLst`, `tmRoot`, root `childTnLst`, `audio`/`video`, `cMediaNode`, `cTn`,
`tgtEl`, command groups, and interactive sequences. Use descendant traversal only inside an already claimed canonical
branch, and reject repeated required nodes before reading attributes.

```ts
function directChildren(parent: XmlElement, localName: string): readonly XmlElement[] {
  return parent.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function positiveTimingId(xml: LosslessXmlDocument, node: XmlElement): number | undefined {
  const value = xml.attribute(node, 'id')?.value;
  if (!value || !/^[1-9]\d*$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed <= 0xFFFF_FFFF ? parsed : undefined;
}
```

Decode absence defaults as volume `1`, hide false, loop false. Treat `mute="1"` as volume `0`; accept only
`repeatCount="indefinite"`; classify every unsupported but shape-targeting graph separately from absence.

- [ ] **Step 5: Implement slide-wide deterministic allocation**

Scan every `p:cTn` once, require unique valid IDs, reserve `max + 1` through `max + count`, and throw before returning
when the range exceeds unsigned 32-bit capacity.

```ts
export function allocateNativeTimingIds(xml: LosslessXmlDocument, count: number): readonly number[] {
  if (!Number.isSafeInteger(count) || count < 0) throw new RangeError('Timing ID count must be a non-negative integer');
  const values = xml.elements('cTn').map((node) => xml.attribute(node, 'id')?.value);
  if (values.some((value) => !value || !/^[1-9]\d*$/.test(value))) {
    throw new Error('Timing tree contains an invalid time-node ID');
  }
  const ids = values.map(Number);
  if (ids.some((id) => !Number.isSafeInteger(id) || id > 0xFFFF_FFFF)
    || new Set(ids).size !== ids.length) {
    throw new Error('Timing tree contains invalid or duplicate time-node IDs');
  }
  const maximum = ids.length === 0 ? 0 : Math.max(...ids);
  if (maximum + count > 0xFFFF_FFFF) throw new RangeError('Native timing ID space is exhausted');
  return Object.freeze(Array.from({ length: count }, (_, index) => maximum + index + 1));
}
```

- [ ] **Step 6: Run focused tests and typecheck**

```bash
pnpm exec vitest run packages/codecs/src/media-timing-state.internal.test.ts
pnpm typecheck
```

Expected: all state/allocator tests pass and declarations compile.

- [ ] **Step 7: Review, commit, and push**

```bash
git diff --check
git add packages/codecs/src/media-timing-state.internal.ts \
  packages/codecs/src/media-timing-state.internal.test.ts
git diff --cached --check
git commit -m "feat: read native media timing"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 2: Versioned playback ownership and canonical timing editor

**Files:**

- Create: `packages/codecs/src/media-timing-edit.internal.ts`
- Create: `packages/codecs/src/media-timing-edit.internal.test.ts`
- Modify: `packages/codecs/src/media-edit.internal.ts`
- Modify: `packages/codecs/src/media-edit.internal.test.ts`

**Interfaces:**

- Consumes: normalized complete playback settings, strict timing state, ID allocator, a direct media picture, and its
  slide XML.
- Produces:

```ts
export interface MediaPlaybackExtensionRecord {
  readonly settings?: Readonly<NormalizedMediaPlaybackSettings>;
  readonly ownership?: Readonly<NativeMediaTimingOwnership>;
  readonly malformed: boolean;
}

export function readMediaPlaybackExtension(
  xml: LosslessXmlDocument,
  picture: XmlElement,
): Readonly<MediaPlaybackExtensionRecord>;

export interface NativeMediaTimingSyncResult {
  readonly changed: boolean;
  readonly ownership: Readonly<NativeMediaTimingOwnership>;
}

export function syncNativeMediaTiming(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: 'audio' | 'video',
  settings: Readonly<NormalizedMediaPlaybackSettings>,
  recorded?: Readonly<NativeMediaTimingOwnership>,
): Readonly<NativeMediaTimingSyncResult>;

export function clearNativeMediaTiming(
  xml: LosslessXmlDocument,
  shapeId: number,
  kind: 'audio' | 'video',
  recorded?: Readonly<NativeMediaTimingOwnership>,
): boolean;
```

Extend `replaceMediaPlaybackExtension(xml, picture, value, ownership?)` to render ownership attributes only when an
ownership tuple is supplied.

- [ ] **Step 1: Write failing private-extension ownership tests**

Require strict parsing of the four preference fields and the `nativeVersion="1"`, `mediaTnId`, `playTnId`, optional
`pauseTnId` tuple. Repeated extensions, repeated playback children, partial tuples, unsupported versions, invalid IDs,
or repeated attributes must set `malformed: true` and grant no ownership.

```ts
expect(readMediaPlaybackExtension(xml, picture)).toEqual({
  settings: { play: 'auto', loop: true, hideWhenStopped: true, volume: 0.5 },
  ownership: { version: 1, mediaTnId: 7, playTnId: 5, pauseTnId: 11 },
  malformed: false,
});
```

- [ ] **Step 2: Write failing canonical editor tests**

Require exact deterministic audio/video XML for click and automatic start, loop, hidden stopped state, zero/fractional/
full volume, audio stop condition, and video toggle-pause interaction. Cover insertion after transition and before
slide extensions, append into a valid imported root, unchanged ordinary animation bytes, healthy no-op, stale repair,
recognized import adoption, clear, last-owned-root cleanup, imported empty-container preservation, and rejection of
unsupported/ambiguous state without mutation.

- [ ] **Step 3: Confirm both focused suites fail**

```bash
pnpm exec vitest run packages/codecs/src/media-edit.internal.test.ts \
  packages/codecs/src/media-timing-edit.internal.test.ts
```

Expected: FAIL on the new ownership reader and timing editor interfaces.

- [ ] **Step 4: Implement strict ownership read/render**

Render the existing four attributes in their current order, then append deterministic ownership attributes. Preserve
the old exact string when ownership is absent.

```ts
const native = ownership
  ? ` nativeVersion="1" mediaTnId="${ownership.mediaTnId}" playTnId="${ownership.playTnId}"`
    + (ownership.pauseTnId === undefined ? '' : ` pauseTnId="${ownership.pauseTnId}"`)
  : '';
return `<p:ext uri="${PLAYBACK_EXTENSION_URI}"><px:playback xmlns:px="urn:pptx-ooxml:media" `
  + `play="${value.play}" loop="${value.loop ? 1 : 0}" `
  + `hideWhenStopped="${value.hideWhenStopped ? 1 : 0}" `
  + `volume="${Math.round(value.volume * 100_000)}"${native}/></p:ext>`;
```

- [ ] **Step 5: Implement canonical root and media graph encoding**

Create a PowerPoint-shaped root/main-sequence only when absent. Allocate all IDs before patching. Encode media state,
`playFrom(0.0)`, click/automatic condition grammar, audio `onStopAudio`, and video `togglePause` as focused rendering
functions that receive explicit IDs and escaped numeric values.

```ts
function renderMediaStartAndEnd(kind: 'audio' | 'video'): string {
  const end = kind === 'audio'
    ? '<p:endCondLst><p:cond evt="onStopAudio" delay="0">'
      + '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>'
    : '';
  return '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' + end;
}

function renderMediaNode(
  kind: 'audio' | 'video',
  shapeId: number,
  id: number,
  settings: Readonly<NormalizedMediaPlaybackSettings>,
): string {
  const repeat = settings.loop ? ' repeatCount="indefinite"' : '';
  const shown = settings.hideWhenStopped ? 0 : 1;
  return `<p:${kind}><p:cMediaNode vol="${Math.round(settings.volume * 100_000)}" `
    + `showWhenStopped="${shown}"><p:cTn id="${id}" fill="hold" display="0"${repeat}>`
    + renderMediaStartAndEnd(kind) + '</p:cTn><p:tgtEl>'
    + `<p:spTgt spid="${shapeId}"/></p:tgtEl></p:cMediaNode></p:${kind}>`;
}
```

- [ ] **Step 6: Implement sync, adoption, repair, and clear**

Call `readNativeMediaTiming()` first. Return unchanged only for `owned-healthy` with equal settings. Replace only the
recognized play/media/pause branches for healthy, stale-recoverable, or uniquely recognized imported graphs. Throw on
unsupported or ambiguous shape-targeting timing. Remove empty library-created wrappers but preserve imported timing
containers and opaque siblings.

- [ ] **Step 7: Run focused tests and typecheck**

```bash
pnpm exec vitest run packages/codecs/src/media-edit.internal.test.ts \
  packages/codecs/src/media-timing-state.internal.test.ts \
  packages/codecs/src/media-timing-edit.internal.test.ts
pnpm typecheck
```

- [ ] **Step 8: Review, commit, and push**

```bash
git diff --check
git add packages/codecs/src/media-edit.internal.ts packages/codecs/src/media-edit.internal.test.ts \
  packages/codecs/src/media-timing-edit.internal.ts \
  packages/codecs/src/media-timing-edit.internal.test.ts
git diff --cached --check
git commit -m "feat: edit native media timing"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 3: Atomic MediaCodec create/read/settings synchronization

**Files:**

- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/media-state.internal.ts`
- Modify: `packages/codecs/src/media-state.internal.test.ts`
- Modify: `packages/codecs/src/media-create.internal.test.ts`
- Modify: `packages/codecs/src/media-edit.internal.test.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: strict playback extension record, native state reader, sync/clear editor, and existing live
  `MediaModel.settings` delegation.
- Produces: zero-plugin native timing on creation, imported native setting projection, atomic live settings edits,
  repair/no-op semantics, and exact rollback.

- [ ] **Step 1: Write failing creation integration tests**

Create default/custom audio and video through `MediaCodec` and require one coherent native graph per picture, unique
slide timing IDs, complete ownership attributes, correct shape targets, private/native value equality, schema order,
and strict reopen.

```ts
const descriptor = await codec.addAudio(slideUri, Uint8Array.of(1), {
  contentType: 'audio/mpeg',
  play: 'auto',
  loop: true,
  hideWhenStopped: true,
  volume: 0.25,
});
const xml = LosslessXmlDocument.parse(pkg.requirePart(slideUri).bytes);
const state = readNativeMediaTiming(xml, descriptor.shapeId, 'audio');
expect(state.settings).toEqual({
  play: 'auto', loop: true, hideWhenStopped: true, volume: 0.25,
});
```

Force timing parse and ID allocation failures after media input resolution and prove package snapshots are unchanged.

- [ ] **Step 2: Write failing imported-read and live-edit tests**

Require private settings to win over native values, recognized PowerPoint native-only graphs to project settings, and
unknown native graphs to produce `{}` without mutation. Through the live model cover equal healthy no-op, equal stale
repair, click↔auto, loop/hide/volume edits, clear, recognized-import adoption, malformed rejection, reopen, and nested
transaction rollback.

- [ ] **Step 3: Confirm focused integration failures**

```bash
pnpm exec vitest run packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media-create.internal.test.ts packages/codecs/src/media-edit.internal.test.ts \
  packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts
```

- [ ] **Step 4: Project effective settings without mutation**

Replace the old descendant-first playback helper with `readMediaPlaybackExtension()`. Return its valid settings first;
otherwise call `readNativeMediaTiming(xml, shapeId, kind)` and return settings only for
`recognized-imported`/`owned-healthy`.

```ts
const preference = readMediaPlaybackExtension(xml, picture);
const native = preference.settings
  ? undefined
  : readNativeMediaTiming(xml, shapeId, kind, preference.ownership);
const settings = preference.settings
  ?? (native?.status === 'recognized-imported' || native?.status === 'owned-healthy'
    ? native.settings
    : undefined)
  ?? Object.freeze({});
```

- [ ] **Step 5: Synchronize both layers during creation**

After picture insertion and before `setPart`, find the inserted direct picture, call `syncNativeMediaTiming()`, then
rewrite the private extension with returned ownership. Keep all part, relationship, picture, and timing mutations in
the existing transaction.

- [ ] **Step 6: Synchronize settings edits and clearing**

In `setSettings`, read preference/native state before no-op comparison. For defined settings, sync native first and
write the preference with returned ownership. For `undefined`, clear native first and then remove the private
extension. Call `setPart` once when either layer changed.

```ts
if (normalized) {
  const sync = syncNativeMediaTiming(
    xml, shapeId, state.kind, normalized, preference.ownership,
  );
  const preferenceChanged = replaceMediaPlaybackExtension(
    xml, picture, normalized, sync.ownership,
  );
  return sync.changed || preferenceChanged;
}
const timingChanged = clearNativeMediaTiming(
  xml, shapeId, state.kind, preference.ownership,
);
return replaceMediaPlaybackExtension(xml, picture, undefined) || timingChanged;
```

- [ ] **Step 7: Run integration tests, typecheck, and full codec/model suites**

```bash
pnpm exec vitest run packages/codecs/src/media-timing-state.internal.test.ts \
  packages/codecs/src/media-timing-edit.internal.test.ts packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media-create.internal.test.ts packages/codecs/src/media-edit.internal.test.ts \
  packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts
pnpm typecheck
```

- [ ] **Step 8: Review, commit, and push**

```bash
git diff --check
git add packages/codecs/src/media.ts packages/codecs/src/media-state.internal.ts \
  packages/codecs/src/media-state.internal.test.ts packages/codecs/src/media-create.internal.test.ts \
  packages/codecs/src/media-edit.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: sync native media playback"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 4: Timing lifecycle and diagnostics

**Files:**

- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/codecs.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: synchronized live playback and strict clear operation.
- Produces: delete/duplicate/move/source/poster isolation plus precise native timing diagnostics.

- [ ] **Step 1: Write failing deletion and lifecycle tests**

Cover deleting healthy owned audio/video, deleting one recognized imported graph, rejecting unsupported/ambiguous
target graphs without removing picture or relationships, removing the last owned timing root, preserving ordinary
animations, and rolling back after a forced relationship/part failure.

Duplicate a timed slide and require independent slide-local ownership/IDs with initially identical XML. Edit and
delete only the duplicate; prove source timing and media payloads remain unchanged. Move slides and prove stable model
identity plus byte-identical timing. Replace source/poster and prove timing IDs/branches stay unchanged.

- [ ] **Step 2: Write failing diagnostic tests**

Construct one case for each code and assert slide URI, shape ID in the message, and no package mutation:

```ts
expect(codec.diagnostics(model, 'powerpoint-2010')).toEqual(expect.arrayContaining([
  expect.objectContaining({ code: 'MEDIA_TIMING_STALE', partUri: slideUri }),
]));
```

Cover `MEDIA_TIMING_MISSING`, `MEDIA_TIMING_STALE`, `MEDIA_TIMING_UNSUPPORTED`, `MEDIA_TIMING_AMBIGUOUS`,
`MEDIA_TIMING_DANGLING_TARGET`, and `MEDIA_TIMING_KIND_MISMATCH`. Retire
`MEDIA_PLAYBACK_TIMING_EXTENSION` only for healthy native timing.

- [ ] **Step 3: Confirm lifecycle and diagnostic failures**

```bash
pnpm exec vitest run packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
```

- [ ] **Step 4: Remove timing before media picture deletion**

Inside the existing delete transaction, resolve the direct picture and preference, call `clearNativeMediaTiming()`
before `xml.removeElement(picture)`, and abort the transaction on unsupported/ambiguous ownership. Keep relationship
reference counting and media target GC unchanged after the XML update.

- [ ] **Step 5: Replace the broad timing-presence diagnostic**

Read native timing status and map each non-healthy state to one deterministic diagnostic. Scan media timing targets
once per slide for dangling/kind mismatch reporting; do not modify the graph or conflate general animations targeting
the media picture.

- [ ] **Step 6: Run lifecycle tests, typecheck, and full suite**

```bash
pnpm exec vitest run packages/codecs/src/media-timing-state.internal.test.ts \
  packages/codecs/src/media-timing-edit.internal.test.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
pnpm test
```

- [ ] **Step 7: Review, commit, and push**

```bash
git diff --check
git add packages/codecs/src/media.ts packages/codecs/src/codecs.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: manage native media timing lifecycle"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 5: Animations plugin reuse and shared ID safety

**Files:**

- Modify: `packages/codecs/src/media.ts`
- Modify: `packages/codecs/src/index.ts`
- Modify: `plugins/animations/src/index.ts`
- Modify: `plugins/animations/src/index.test.ts`

**Interfaces:**

- Consumes: `MediaCodec.materializePlayback(slidePartUri, shapeId?)` and `allocateNativeTimingIds()` from the built-in
  codec package.
- Produces: idempotent legacy private-only materialization and one timing ID allocator for media and general effects.

- [ ] **Step 1: Write failing plugin compatibility tests**

Require plugin installation to make zero changes to current healthy output, materialize a legacy private-only audio
and video exactly once, repair a stale owned graph, preserve recognized native-only PowerPoint graphs, preserve and
report unsupported timing, and remain byte-identical on a second installation attempt.

Add an existing animation with high sparse IDs, then materialize media and add another effect. Require all `p:cTn`
IDs to be unique and monotonically allocated above the imported maximum.

- [ ] **Step 2: Confirm plugin tests fail**

```bash
pnpm exec vitest run plugins/animations/src/index.test.ts
```

- [ ] **Step 3: Add the low-level materialization entry point**

Expose a narrow `MediaCodec.materializePlayback(slidePartUri, shapeId?)` that lists direct media pictures, normalizes
only valid private preferences, calls shared synchronization inside one transaction, writes changed parts once, and
returns the number created or repaired.

```ts
materializePlayback(slidePartUri: string, shapeId?: number): number {
  return this.pkg.transaction(() => {
    const part = this.pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    let changed = 0;
    for (const picture of directMediaPictures(xml, shapeId)) {
      const state = readMediaState(this.pkg, slidePartUri, xml, picture);
      const preference = readMediaPlaybackExtension(xml, picture);
      if (!state || !preference.settings || preference.malformed) continue;
      const sync = syncNativeMediaTiming(
        xml, state.shapeId, state.kind, preference.settings, preference.ownership,
      );
      const preferenceChanged = replaceMediaPlaybackExtension(
        xml, picture, preference.settings, sync.ownership,
      );
      if (sync.changed || preferenceChanged) changed += 1;
    }
    if (changed > 0) this.pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
    return changed;
  });
}
```

Define `directMediaPictures(xml, shapeId?)` beside the method using direct `cSld/spTree/pic` ownership and numeric
`cNvPr/@id`; it must not use a descendant-wide picture scan.

- [ ] **Step 4: Delegate plugin media materialization**

Replace the plugin's descendant scan and `encodeMediaTiming()` with `new MediaCodec(this.pkg).materializePlayback()`.
Delete the duplicate media encoder. Change `nextTimingId()` to call the shared allocator for general animation IDs.

- [ ] **Step 5: Run plugin, codec, and type gates**

```bash
pnpm exec vitest run plugins/animations/src/index.test.ts \
  packages/codecs/src/media-timing-state.internal.test.ts \
  packages/codecs/src/media-timing-edit.internal.test.ts packages/codecs/src/codecs.test.ts
pnpm typecheck
pnpm test
```

- [ ] **Step 6: Review, commit, and push**

```bash
git diff --check
git add packages/codecs/src/media.ts packages/codecs/src/index.ts \
  plugins/animations/src/index.ts plugins/animations/src/index.test.ts
git diff --cached --check
git commit -m "refactor: share native media timing codec"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 6: Public SDK, format, and PptxGenJS conformance

**Files:**

- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/codecs/src/codecs.test.ts`

**Interfaces:**

- Consumes: the complete built-in native timing lifecycle through public `PptxDocument`/`MediaModel` APIs.
- Produces: declaration-safe current API behavior, six-format coverage, PptxGenJS legacy adoption, and public package
  conformance evidence.

- [ ] **Step 1: Extend public SDK tests**

Through `PptxDocument.create()`, add default/custom audio and video, assert live identity, inspect native timing, edit
all four settings, duplicate/move/remove, write/reopen, and verify package diagnostics. Compile valid partial settings
and existing negative declarations to prove no public type widening.

- [ ] **Step 2: Add six-format native timing matrix**

For each `pptx|pptm|potx|potm|ppsx|ppsm`, create timed audio/video, edit after reopen, duplicate, delete one media,
write twice, and require retained package profile, exact native settings/targets, unique IDs, and zero timing errors.

```ts
for (const format of ['pptx', 'pptm', 'potx', 'potm', 'ppsx', 'ppsm'] as const) {
  const document = PptxDocument.create({ format });
  const slide = document.addSlide();
  await slide.addAudio(Uint8Array.of(1), { contentType: 'audio/mpeg', play: 'auto' });
  const reopened = await PptxDocument.open(await document.write());
  expect(reopened.format).toBe(format);
  expect(reopened.media(0)[0]?.settings.play).toBe('auto');
}
```

- [ ] **Step 3: Add PptxGenJS 4.0.1 native-timing adoption fixtures**

Generate public audio/video/duplicate-media output with `addMedia()` and public write APIs. Prove import remains
private/timing-free and byte-preserved on no-op. Assign settings to add one canonical native graph, reopen, edit again,
duplicate, and remove while retaining the known `a:videoFile` audio, `audio/mp3`, and Microsoft-media duplicate import
compatibility outside the timing-owned spans.

- [ ] **Step 4: Run public and complete gates**

```bash
pnpm exec vitest run packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  packages/pptx/src/index.test.ts plugins/animations/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
```

- [ ] **Step 5: Review, commit, and push**

```bash
git diff --check
git add packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  packages/pptx/src/index.test.ts packages/codecs/src/codecs.test.ts
git diff --cached --check
git commit -m "test: verify native media timing conformance"
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
- Modify: `docs/plugins.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/compatibility/cross-client-testing.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`
- Temporary only: installed-tarball consumer, declaration fixture, generated media/timing galleries, rendered slides,
  validation JSON, and PowerPoint/LibreOffice round trips under `/tmp/pptx-native-media-timing-*`.

**Interfaces:**

- Consumes: complete native timing behavior through the actual packed `@jiayunxie/pptx` package.
- Produces: permanent Node/browser smoke, deterministic build evidence, PowerPoint/LibreOffice normalization records,
  and accurate support/remaining-work documentation.

- [ ] **Step 1: Extend packed Node and declaration smoke**

Use only the installed tarball to create default/custom audio/video, inspect native XML/IDs/targets, edit and clear live
settings, duplicate and isolate slide-local timing, remove one media, write/reopen, and validate. Add
`nativeMediaTiming: true` to smoke JSON only when private/native equality, unique IDs, no dangling targets, and zero
timing diagnostics all hold.

- [ ] **Step 2: Extend permanent browser smoke**

Through `dist/browser.js`, repeat audio/video create/edit/clear/duplicate/remove using data URI, Blob, and stream
inputs.
Return timing element/command counts, unique ID count, settings after reopen, target isolation, diagnostics, and
`nativeMediaTiming: true`. Keep the static Node-import scan and download smoke passing.

- [ ] **Step 3: Clean-build twice, pack, install, and run public gates**

Build the public package twice from clean state, hash every dist file in stable order, require identical manifests,
pack the second build, install only that tarball without workspace links, and run Node, declaration, browser, CLI, plus
installed `pptx-inspect doctor`, package inspection, and PowerPoint 2010 validation.

```bash
pnpm build
pnpm --filter @jiayunxie/pptx build
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-native-media-timing-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-native-media-timing-pack/jiayunxie-pptx-0.1.0.tgz
node scripts/playwright-browser-smoke.js
```

- [ ] **Step 4: Generate and inspect native timing galleries**

Use ffmpeg to create playable MP3/M4A/WAV/OGG and MP4/MOV/WebM inputs with PNG/JPEG/GIF posters. Through the installed
tarball create slides covering click/auto, loop, visible/hidden stopped state, volume 0/0.25/0.5/1, mixed ordinary
animations, two media on one slide, legacy materialization, recognized PowerPoint import, duplicate edits, clear, and
deletion. Keep external-media warnings in a separate deck.

Use `pptx-inspect` to inspect/validate exact package parts and record slide/shape/timing IDs, commands, targets,
settings, diagnostics, relationship roles, payload hashes, and orphan counts.

- [ ] **Step 5: PowerPoint and LibreOffice open/save/reopen validation**

Open the embedded gallery in PowerPoint, save a copy, reopen with the library, and compare semantic playback settings,
shape targets, unique timing IDs, media/poster hashes, and non-owned animations. Repeat through LibreOffice. Render all
slides at 180 DPI, check overflow, and visually inspect poster visibility. Record client rewrites as normalization;
require no dangling target, lost media, duplicate playback graph, package error, or unexpected ordinary-animation
change.

- [ ] **Step 6: Update public documentation and support accounting**

Document zero-plugin native playback, effective read order, strict import/adoption boundary, settings clear semantics,
owned timing lifecycle, plugin legacy behavior, diagnostics, PptxGenJS differences, six-format results, packed/browser/
CLI evidence, deterministic build, and client normalization.

Move click/auto, loop, hide-when-stopped, and volume native timing into supported. Keep trim/bookmarks, finite repeats,
narration/cross-slide audio, captions/subtitles, online video, media styling, transcoding engines, and broad client
certification pending. Set the next highest-value PptxGenJS parity slice from the refreshed compatibility matrix.

- [ ] **Step 7: Run final gates, review, commit, and push**

```bash
pnpm exec vitest run packages/codecs/src/media-timing-state.internal.test.ts \
  packages/codecs/src/media-timing-edit.internal.test.ts packages/codecs/src/media-state.internal.test.ts \
  packages/codecs/src/media-create.internal.test.ts packages/codecs/src/media-edit.internal.test.ts \
  packages/codecs/src/codecs.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts packages/pptx/src/index.test.ts \
  plugins/animations/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
pnpm --filter @jiayunxie/pptx build
git diff --check
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js README.md \
  packages/pptx/README.md docs/api/README.md docs/plugins.md \
  docs/compatibility/pptxgenjs-baseline.md docs/compatibility/cross-client-testing.md \
  docs/implementation-progress.md CHANGELOG.md
git diff --cached --check
git commit -m "docs: document native media timing"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected final state: divergence `0 0`; only `.pnpm-store/` remains untracked.
