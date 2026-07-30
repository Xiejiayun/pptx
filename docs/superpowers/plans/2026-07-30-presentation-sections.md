# Presentation Sections Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add PptxGenJS-compatible presentation section creation plus strict direct-state reading, surgical existing-deck editing, slide-lifecycle synchronization, adapter conformance, packed-surface proof, and real-deck QA.

**Architecture:** Put namespace-aware PowerPoint 2010 section parsing and source-span mutation in one model-internal helper. Expose detached snapshots and ID-addressed atomic commands from `PresentationModel`, then make add/duplicate/delete/move slide operations maintain section membership inside their existing OPC transactions.

**Tech Stack:** TypeScript strict mode, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/model`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, `pptx-inspect`, LibreOffice headless.

## Global Constraints

- Sections own only the presentation root's unique PowerPoint 2010 section extension, its section order/IDs/names, and direct slide-ID memberships.
- Read and mutation discovery is direct-child and expanded-namespace correct; lexical prefix, part URI, descendant order, and first-match guessing are not contracts.
- Absence and a valid empty list read as `[]`; structurally unsafe or ambiguous section state reads as `undefined` and rejects mutation before any package change.
- Empty sections and loose slides are valid; section membership may omit presentation slides but may not reference unknown IDs or contain one slide more than once.
- Public snapshots are deeply detached. Section commands address exact brace-wrapped GUIDs; duplicate titles remain legal.
- New titles and option objects use descriptor-safe strict normalization. No coercion, trimming, getter invocation, inherited keys, retained caller references, or invalid XML characters are allowed.
- `addSlide({ sectionTitle })` selects the first exact title; unknown titles throw before slide creation. Omitted titles under active sections create/continue canonical `Default-N` final state.
- Add/duplicate/delete/move slide operations keep section references valid and execute atomically with existing dependency lifecycle work.
- Same-value rename, same-position move, same-membership assignment, and absent unassign are exact part/journal no-ops.
- Foreign extension siblings, PowerPoint 2012 slide guides, alternate prefixes, comments, whitespace, unknown section-container content, slide parts, and unrelated relationships/parts remain preserved.
- Adapter tests use only PptxGenJS public constructor, `addSection()`, `addSlide()`, and `write()`.
- Each repository-changing task ends with focused review, explicit staging, commit, SSH port 443 push, fetch, and `origin/main...HEAD` equal to `0 0`.
- Execute inline because the user authorized autonomous continuation and repository instructions prohibit subagent dispatch.
- Never modify, delete, stage, or commit `.pnpm-store/`.

---

### Task 1: Add the strict section XML helper

**Files:**
- Create: `packages/model/src/presentation-sections.internal.ts`
- Create: `packages/model/src/presentation-sections.internal.test.ts`

**Interfaces:**
- Produces `SectionSnapshotData { id: string; title: string; slideIds: readonly number[] }`.
- Produces `readPresentationSections(xml, validSlideIds): readonly SectionSnapshotData[] | undefined`.
- Produces `insertPresentationSection(xml, validSlideIds, title, order, id): SectionSnapshotData`.
- Produces `renamePresentationSection(xml, validSlideIds, sectionId, title): boolean`.
- Produces `movePresentationSection(xml, validSlideIds, sectionId, toIndex): boolean`.
- Produces `deletePresentationSection(xml, validSlideIds, sectionId): boolean`.
- Produces `assignPresentationSlideToSection(xml, validSlideIds, slideId, sectionId: string | undefined): boolean`.
- Produces `removePresentationSlideFromSections(xml, validSlideIds, slideId): boolean`.
- Produces `copyPresentationSlideSection(xml, validSlideIds, sourceSlideId, targetSlideId): boolean`.
- Produces `sortPresentationSectionSlides(xml, validSlideIds, orderedSlideIds): boolean`.
- Produces strict normalization helpers for title, ID, indices, add-section options, and add-slide options.

- [ ] **Step 1: Write RED fixtures and read tests**

Create a compact `presentationXml()` fixture with canonical presentation slide IDs and injectable direct `extLst` content. Require:

```ts
const sections = readPresentationSections(
  LosslessXmlDocument.parse(presentationXml(sectionList([
    section('{00000000-0000-0000-0000-000000000001}', 'Intro', [256]),
    section('{00000000-0000-0000-0000-000000000002}', 'Data', [257, 258]),
  ]))),
  new Set([256, 257, 258]),
);

expect(sections).toEqual([
  { id: '{00000000-0000-0000-0000-000000000001}', title: 'Intro', slideIds: [256] },
  { id: '{00000000-0000-0000-0000-000000000002}', title: 'Data', slideIds: [257, 258] },
]);
expect(readPresentationSections(
  LosslessXmlDocument.parse(presentationXml()),
  new Set([256]),
)).toEqual([]);
```

Also require valid empty sections, loose slides, duplicate names, lower/upper GUID hex, XML-decoded names, alternate presentation and p14 prefixes, namespace declarations on `sectionLst`, foreign extension siblings, slide-guide extension, comments, whitespace, and unknown non-owned children. Mutate returned arrays and prove a second read is unchanged.

- [ ] **Step 2: Write RED unsafe-state tests**

Require `undefined` for duplicate direct `extLst`; duplicate owned `p:ext`; wrong/missing/repeated unqualified extension `uri`; missing/repeated/wrong-namespace `sectionLst`; descendant impostors; missing/repeated `name` or `id`; invalid/unbraced/repeated section GUIDs; missing/repeated/wrong-namespace `sldIdLst`; missing/repeated/non-decimal/negative/unsafe/unknown member ID; duplicate member within one section or across sections; wrong-namespace member; and multiple presentation roots.

Assert malformed XML throws `LosslessXmlError`, and every read leaves `xml.changed === false`.

- [ ] **Step 3: Write RED source-span mutation tests**

Cover absent extension creation, reuse of an existing `extLst`, append and insertion at indices `0`, middle, and end, empty/self-closing list expansion, alternate prefixes, title escaping, stable supplied GUID, rename, move, delete, assignment/reassignment/unassignment, member removal/copy/sort, same-state no-ops, and foreign-state preservation.

Use exact isolation assertions:

```ts
const xml = LosslessXmlDocument.parse(sourceWithGuidesAndForeignChildren);
expect(renamePresentationSection(
  xml,
  new Set([256, 257]),
  '{00000000-0000-0000-0000-000000000001}',
  'A&B <One> "Two"',
)).toBe(true);
expect(xml.serialize()).toContain('name="A&amp;B &lt;One&gt; &quot;Two&quot;"');
expect(xml.serialize()).toContain('<p15:sldGuideLst');
expect(xml.serialize()).toContain('<x:keep value="FOREIGN"/>');

const same = LosslessXmlDocument.parse(sourceWithGuidesAndForeignChildren);
expect(renamePresentationSection(
  same,
  new Set([256, 257]),
  '{00000000-0000-0000-0000-000000000001}',
  'Intro',
)).toBe(false);
expect(same.changed).toBe(false);
expect(same.serialize()).toBe(sourceWithGuidesAndForeignChildren);
```

Deleting the last canonical-only section must remove the owned extension and remove `extLst` only when no other children remain. When the section extension has foreign state, retain its empty section list.

- [ ] **Step 4: Write RED strict-normalization tests**

Require valid frozen and null-prototype option objects. Reject empty/whitespace/XML-invalid titles; null/array/custom-prototype/accessor/inherited/symbol/extra-key options; non-string section IDs/titles; invalid GUIDs; non-finite/fractional/unsafe/out-of-range indices; unknown section IDs; and unknown slide IDs. Capture XML source and `changed` before every failure and require exact zero mutation.

- [ ] **Step 5: Run RED**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-sections.internal.test.ts --reporter=dot
```

Expected: FAIL because `presentation-sections.internal.ts` does not exist.

- [ ] **Step 6: Implement namespace-aware discovery and normalization**

Start the helper with these constants and shapes:

```ts
import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const SECTION_NAMESPACE =
  'http://schemas.microsoft.com/office/powerpoint/2010/main';
const SECTION_EXTENSION_URI =
  '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const SECTION_ID = /^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/i;

export interface SectionSnapshotData {
  readonly id: string;
  readonly title: string;
  readonly slideIds: readonly number[];
}
```

Implement `directChildren()`, exact unqualified attribute collection, lexical prefix extraction, ancestor-walking in-scope namespace resolution, and section-root discovery. Attribute values come from the lossless parser's decoded `value`; mutation uses `escapeXmlAttribute()` or `replaceAttribute()`.

Normalize objects with `Reflect.ownKeys()` and `Object.getOwnPropertyDescriptor()`. Only ordinary and null-prototype objects, supported own string keys, and data descriptors are accepted. Validate titles with `/\S/u` plus the repository XML 1.0 character contract. Validate indices without flooring/clamping.

- [ ] **Step 7: Implement strict snapshots**

`readPresentationSections()` must return a freshly allocated array and member arrays. It may return `[]` only for true absence or a valid empty list. Build a set of seen section GUIDs and a set of seen member IDs; require every member in `validSlideIds`. Preserve lexical GUID case in the snapshot while comparing IDs exactly for command lookup.

Malformed or ambiguous owned structure returns `undefined`; XML parse errors remain thrown by the caller. Do not set patches during read.

- [ ] **Step 8: Implement minimal mutations**

Each mutation first calls one strict structural resolver and throws `ModelParseError('Presentation sections are not safely editable')` when it cannot establish ownership. No patch is added before complete validation.

Use these canonical fragments for new state, substituting the presentation lexical prefix and escaping values:

```ts
const list = `<p14:sectionLst xmlns:p14="${SECTION_NAMESPACE}">${sectionXml}</p14:sectionLst>`;
const extension = `<${p}ext uri="${SECTION_EXTENSION_URI}">${list}</${p}ext>`;
const sectionXml = `<p14:section name="${escapeXmlAttribute(title)}" id="${id}"><p14:sldIdLst/></p14:section>`;
const memberXml = `<${sectionPrefix}sldId id="${slideId}"/>`;
```

Move and sort operations replace only the owned list interior with original element byte slices in the desired order while retaining foreign direct children in their original relative order. Rename uses `replaceAttribute()`. Add/remove member operations touch only direct member spans. Return `false` without patches for exact no-ops.

- [ ] **Step 9: Run focused and model gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-sections.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
git diff --check
```

- [ ] **Step 10: Review, commit, push, and verify**

Review expanded namespaces, direct-child cardinality, valid empty/loose states, duplicate membership rejection, descriptor safety, no patch-before-validation, self-closing expansion, foreign XML preservation, no-op patches, and exact helper signatures. Then:

```sh
git add -- packages/model/src/presentation-sections.internal.ts \
  packages/model/src/presentation-sections.internal.test.ts
git diff --cached --check
git commit -m "feat: add presentation section codec"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Expose section snapshots and atomic editing commands

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces public `PresentationSection`, `AddSectionOptions`, and `AddSlideOptions` types from `@pptx/model` and the aggregate package.
- Produces `PresentationModel.sections`, `addSection()`, `renameSection()`, `moveSection()`, `deleteSection()`, and `assignSlideToSection()`.
- Does not yet change add/duplicate/delete/move slide membership; Task 3 owns lifecycle integration.

- [ ] **Step 1: Write RED public snapshot tests**

Add one SDK test near the zero-slide creation test:

```ts
const document = PptxDocument.create();
expect(document.sections).toEqual([]);
document.addSlide();
const intro = document.addSection({ title: 'Intro & <Start>' });
expect(intro).toMatchObject({ title: 'Intro & <Start>', slideIds: [] });
expect(intro.id).toMatch(/^\{[0-9A-F-]{36}\}$/);

const first = document.sections!;
(first as { title: string }[])[0]!.title = 'detached';
(first[0]!.slideIds as number[]).push(999);
expect(document.sections).toEqual([{ ...intro, slideIds: [] }]);
```

Write/reopen and require the same ID/title/order/membership state. Repeat for all six `PRESENTATION_FORMAT_PROFILES`.

- [ ] **Step 2: Write RED command tests**

Build sections A/C, insert B at order 1, rename with escaped Unicode, move C to 0, assign/reassign/unassign three slides, delete B, and reopen. Require exact snapshots after each operation and no slide order/part URI change.

Capture presentation bytes around same-title rename, same-index move, repeated assignment, absent unassign, and require byte equality plus no new journal entry.

- [ ] **Step 3: Write RED invalid and rollback tests**

Cover invalid titles/options/indices/IDs, unknown section/slide, accessor/symbol/extra keys, duplicate title editing by ID, unsafe injected extension, and outer transaction rollback:

```ts
const before = await document.write();
expect(() => document.transaction((draft) => {
  const section = draft.addSection({ title: 'Temporary' });
  draft.assignSlideToSection(0, section.id);
  throw new Error('rollback sections');
})).toThrow('rollback sections');
expect(await document.write()).toEqual(before);
```

Assert all invalid calls preserve `document.slides[0]` identity and every package part hash.

- [ ] **Step 4: Run RED**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts -t 'presentation section' --reporter=dot
```

Expected: FAIL because public types and methods do not exist.

- [ ] **Step 5: Add public types and snapshot getter**

Add to `presentation.ts`:

```ts
export interface PresentationSection {
  readonly id: string;
  readonly title: string;
  readonly slideIds: readonly number[];
}

export interface AddSectionOptions {
  readonly title: string;
  readonly order?: number;
}

export interface AddSlideOptions {
  readonly sectionTitle?: string;
}
```

`packages/model/src/index.ts` already uses `export * from './presentation.js'`, so do not add a duplicate export. Implement:

```ts
get sections(): readonly PresentationSection[] | undefined {
  const { xml } = this.parsePresentation();
  return readPresentationSections(xml, new Set(this.slides.map(({ slideId }) => slideId)));
}
```

The helper's structural result is assignable to the public readonly shape; never return cached input arrays.

- [ ] **Step 6: Implement atomic commands**

Normalize option/argument values before entering the package transaction where possible. Inside each transaction, parse the current presentation, pass the current slide-ID set, call exactly one helper, and save only when `xml.changed`.

Use this pattern:

```ts
renameSection(sectionId: string, title: string): void {
  const normalizedId = normalizePresentationSectionId(sectionId);
  const normalizedTitle = normalizePresentationSectionTitle(title);
  this.opcPackage.transaction(() => {
    const { xml } = this.parsePresentation();
    const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
    if (renamePresentationSection(xml, slideIds, normalizedId, normalizedTitle)) {
      this.setXmlPart(this.presentationPartUri, xml.serialize());
    }
  });
}
```

Generate new IDs with `globalThis.crypto.randomUUID().toUpperCase()` inside braces before mutation; if unavailable, throw before patching. `addSection()` returns a newly allocated copy of the helper result.

- [ ] **Step 7: Run public, model, and type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-sections.internal.test.ts \
  packages/sdk/src/index.test.ts -t 'presentation section' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src packages/sdk/src --reporter=dot
pnpm typecheck
git diff --check
```

- [ ] **Step 8: Review, commit, push, and verify**

Review public type exports, detached values, exact ID addressing, strict indices, duplicate title behavior, six-format reopen, part isolation, no-op journal behavior, and outer rollback. Then:

```sh
git add -- packages/model/src/presentation.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit presentation sections"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Synchronize sections with the slide lifecycle

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Changes `addSlide()` to `addSlide(options?: AddSlideOptions): SlideModel`.
- Keeps current no-argument source compatibility.
- Makes add/duplicate/delete/move operations preserve valid section final state in the same transaction.

- [ ] **Step 1: Write RED add-slide parity tests**

Require explicit first-title matching under duplicate titles, escaped titles, unknown-title zero mutation, loose slides before the first section, and automatic defaults:

```ts
const document = PptxDocument.create();
const intro = document.addSection({ title: 'Intro' });
const explicit = document.addSlide({ sectionTitle: 'Intro' });
const automaticOne = document.addSlide();
const automaticTwo = document.addSlide();
expect(document.sections).toEqual([
  { ...intro, slideIds: [explicit.slideId] },
  expect.objectContaining({ title: 'Default-1', slideIds: [automaticOne.slideId, automaticTwo.slideId] }),
]);
```

Add another user section after Default-1, then call `addSlide()` and require a new Default-2. Reopen and add one more slide; require continuation of the last canonical default.

- [ ] **Step 2: Write RED duplicate/delete/move tests**

Create three sections with loose and assigned slides. Require duplicate of an assigned slide to join the same section, duplicate of loose to remain loose, delete to remove every reference while retaining an empty section, and move to preserve membership while sorting each section's member IDs by global presentation order.

Assert slide dependency clone/GC behavior from existing tests remains unchanged and only presentation XML plus the already-expected slide/dependency parts change.

- [ ] **Step 3: Write RED unsafe-state and rollback tests**

Inject duplicate membership and require add/duplicate/delete/move to throw before changing bytes, relationships, content types, slide cache, or journal. In an outer transaction, combine `addSlide({ sectionTitle })`, `duplicateSlide()`, `moveSlide()`, and `deleteSlide()` and throw; require exact restoration.

- [ ] **Step 4: Run RED**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts -t 'section.*slide|slide.*section' --reporter=dot
```

Expected: FAIL because slide lifecycle operations do not yet update memberships.

- [ ] **Step 5: Extend addSlide atomically**

Normalize `AddSlideOptions` before transaction. At transaction start, read the strict section state. If `sectionTitle` is defined, find the first exact title or throw before `setPart()`/`addRelationship()`.

For omitted title under active sections, choose the last section when its title matches `^Default-[1-9]\d*$`; otherwise insert a new `Default-N` using the lowest positive N not already present. Then call `attachSlide()` and assign the new `slideId` in the same outer transaction.

Do not create a section when the current section list is absent/empty. A slide added before `addSection()` stays loose.

- [ ] **Step 6: Extend duplicate, delete, and move**

Before duplicate, read the source section ID; after `attachSlide()`, call `copyPresentationSlideSection()` on a fresh presentation parse.

During delete, call `removePresentationSlideFromSections()` on the same parsed presentation before removing the direct `p:sldId`; both patches are non-overlapping and serialize once.

During move, compute the final slide ID order from the reordered direct `sldId` elements and call `sortPresentationSectionSlides()` before serialization. Do not infer a new section from visual adjacency.

- [ ] **Step 7: Run focused and regression gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-sections.internal.test.ts \
  packages/sdk/src/index.test.ts -t 'presentation section|section.*slide|slide.*section|dependency|duplicate slide|delete slide|move slide' \
  --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src --reporter=dot
pnpm typecheck
git diff --check
```

- [ ] **Step 8: Review, commit, push, and verify**

Review pre-mutation validation, explicit/automatic assignment, Default-N persistence, duplicate title first match, loose-slide behavior, dependency isolation, dangling-reference prevention, section-member sorting, model identity, and rollback. Then:

```sh
git add -- packages/model/src/presentation.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: synchronize slides with presentation sections"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Compare public PptxGenJS section output

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Proves PptxGenJS 4.0.1 none, explicit, escaped, empty, order, default, loose-before, and unknown-title final states import through `document.sections`.
- Compares native and PptxGenJS semantic final state without depending on random GUID equality.

- [ ] **Step 1: Extend the public-only test interface**

Extract the current anonymous slide return type and add only documented APIs:

```ts
interface PptxGenJSSlide {
  addText(
    text: string | readonly { readonly text: string; readonly options?: Record<string, unknown> }[],
    options: Record<string, unknown>,
  ): void;
  addTable(
    rows: readonly (readonly {
      readonly text?: string;
      readonly options?: Record<string, unknown>;
    }[])[],
    options: Record<string, unknown>,
  ): void;
}

interface PptxGenJSInstance {
  addSection(options: { readonly title: string; readonly order?: number }): void;
  addSlide(options?: { readonly sectionTitle?: string }): PptxGenJSSlide;
}
```

Do not expose `_sections`, `_type`, `_slides`, UUID helpers, or generated ZIP internals.

- [ ] **Step 2: Add semantic normalization**

Normalize random section IDs away while retaining titles/order/membership:

```ts
const sectionState = (document: PptxDocument) =>
  document.sections?.map(({ title, slideIds }) => ({ title, slideIds }));
```

Because both writers start presentation slide IDs at 256, exact member arrays are comparable.

- [ ] **Step 3: Add public-output conformance cases**

For each case create a fresh PptxGenJS instance, call only public methods, `importPptxGenJS()`, then write/reopen through native:

- none with two slides → `[]`;
- explicit escaped title with two assigned slides;
- empty section and loose slide;
- A/C plus B at positive `order: 1` → A/B/C;
- user Intro plus two omitted-title slides → Intro and Default-1;
- loose slide before Later section → only second slide assigned;
- unknown title → warning final state with empty Known section and one loose slide.

Spy on `console.warn` only around the unknown case and restore it in `finally`. Require all imported IDs to be valid but do not compare UUID values. Require native write/reopen state to match before reopen.

- [ ] **Step 4: Compare supported native final state**

Build native explicit/order/default/loose cases through `addSection()` and `addSlide()`. Compare `sectionState()` with PptxGenJS imports. Document in assertions that native rejects unknown title and repairs `order: 0` to first insertion rather than copying PptxGenJS's append bug.

- [ ] **Step 5: Run focused and full adapter gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts -t 'presentation section' --reporter=dot
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review public-only API usage, UUID-independent comparison, exact slide IDs, escaped names, empty/loose state, order behavior, automatic defaults, warning restoration, native strict repairs, and reopen stability. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare presentation sections with pptxgenjs"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Prove packed Node, browser, and declaration sections

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves aggregate package runtime and types expose section snapshots, commands, and `addSlide({ sectionTitle })`.
- Adds the JSON check `presentationSections: true`.

- [ ] **Step 1: Add packed Node runtime coverage**

In the generated Node consumer, create three slides and sections, then rename/move/assign/delete/reopen:

```ts
const sectioned = PptxDocument.create();
const firstSection = sectioned.addSection({ title: 'Packed & <Intro>' });
const assignedSlide = sectioned.addSlide({ sectionTitle: 'Packed & <Intro>' });
const automaticSlide = sectioned.addSlide();
const dataSection = sectioned.addSection({ title: 'Data', order: 1 });
sectioned.assignSlideToSection(1, dataSection.id);
sectioned.renameSection(firstSection.id, 'Edited intro');
sectioned.moveSection(dataSection.id, 0);
const reopenedSections = (await PptxDocument.open(await sectioned.write())).sections;
```

Require order Data/Edited intro, assigned membership, valid IDs, detached snapshots, and escaped XML. Add `presentationSections` to the result object.

- [ ] **Step 2: Add packed browser runtime coverage**

Inside the existing browser bundle consumer, repeat a short add/assign/rename/write/reopen path using only globals available in the bundle. Throw `Error('Browser presentation sections failed')` unless title/order/member IDs survive.

- [ ] **Step 3: Add declaration coverage**

Import `AddSectionOptions`, `AddSlideOptions`, and `PresentationSection`. Add:

```ts
const addSectionOptions: AddSectionOptions = { title: 'Typed', order: 0 };
const typedSection: PresentationSection = createdDocument.addSection(addSectionOptions);
const addSlideOptions: AddSlideOptions = { sectionTitle: typedSection.title };
createdDocument.addSlide(addSlideOptions);
const sectionSnapshot: readonly PresentationSection[] | undefined = createdDocument.sections;
createdDocument.renameSection(typedSection.id, 'Renamed');
createdDocument.moveSection(typedSection.id, 0);
createdDocument.assignSlideToSection(0, typedSection.id);
createdDocument.deleteSection(typedSection.id);
```

Include all new variables in the final `void [...]` expression.

- [ ] **Step 4: Build, pack, and run the actual tarball**

```sh
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
pptx_sections_pack_dir=$(mktemp -d /tmp/pptx-sections-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_sections_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_sections_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require JSON containing `"presentationSections": true`, `"types": true`, and `"cli": "0.1.0"`.

- [ ] **Step 5: Review, commit, push, and verify**

Review root export, generated declarations, browser bundle, detached snapshots, packed edit/reopen behavior, and absence of workspace-only imports. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed presentation sections"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Document section creation, editing, and parity boundaries

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Marks PptxGenJS presentation sections supported for valid final states.
- Documents strict snapshots, ID-addressed commands, empty/loose sections, automatic defaults, slide lifecycle synchronization, and intentional repairs.

- [ ] **Step 1: Update the compatibility matrix**

Add beside the other presentation rows:

```md
| presentation `addSection({ title, order? })` / `addSlide({ sectionTitle })` | `document.sections` / `addSection()` / section editing commands / `addSlide({ sectionTitle })` | 已支持 strict section extension 创建、detached读取、ID寻址编辑与slide lifecycle同步；native修复unknown title和order 0，保留empty section与loose slide |
```

Replace sections in the pending sentence after theme. Add one focused paragraph covering p14 ownership, member use of presentation slide IDs, duplicate-title first match, Default-N behavior, PptxGenJS random-ID normalization, no slide-guide synthesis, and native existing-deck extensions. Keep custom properties, masters, layouts, and placeholders pending.

- [ ] **Step 2: Update package README and changelog**

Add a compact API example:

```ts
const intro = document.addSection({ title: 'Intro' });
document.addSlide({ sectionTitle: 'Intro' });
document.renameSection(intro.id, 'Overview');
document.assignSlideToSection(0, intro.id);
```

Document `sections === []` for absence, `undefined` for unsafe state, detached `slideIds`, delete leaving slides loose, automatic Default-N, and section-aware duplicate/delete/move. Add one changelog bullet describing the complete public slice without claiming master/layout/placeholder completion.

- [ ] **Step 3: Run docs, type, and search gates**

```sh
rg -n 'sections|addSection|sectionTitle|assignSlideToSection|Default-' \
  CHANGELOG.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
rg -n 'sections.*(pending|待)|presentation sections.*(unsupported|未支持)' \
  CHANGELOG.md docs/compatibility packages/pptx/README.md || true
pnpm typecheck
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

Review that docs do not claim named masters, layouts, placeholders, slide guides, hidden-slide state, or byte-identical random UUIDs. Then:

```sh
git add -- CHANGELOG.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "docs: document presentation sections"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run full release and real-deck QA

**Files:**
- No repository changes expected.
- Generate artifacts only under `/tmp/pptx-sections-qa-20260730` and a `mktemp` pack directory.

**Interfaces:**
- Proves full repository compatibility, performance, packed semantics, PowerPoint validation, mutation isolation, save stability, slide rendering, and remote synchronization.

- [ ] **Step 1: Run complete repository gates**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
pnpm --filter @jiayunxie/pptx build
pptx_sections_qa_pack_dir=$(mktemp -d /tmp/pptx-sections-qa-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_sections_qa_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_sections_qa_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require all tests, the 5-second performance budget, actual Node/browser/types tarball smoke, and CLI smoke to pass.

- [ ] **Step 2: Generate representative native and PptxGenJS decks**

Using only built public APIs, generate:

```text
native-none.pptx
native-sections.pptx
native-edited.pptx
native-reopened.pptx
native-second-write.pptx
pptxgenjs-explicit.pptx
pptxgenjs-default.pptx
pptxgenjs-loose.pptx
```

Every deck uses wide layout and visible numbered slide labels. Native sections must include escaped Unicode title, empty section, duplicate title, loose slide, reassignment, rename, move, delete, duplicate slide, and automatic Default-N. Keep generator/verifier source under `/tmp` only and create it with `apply_patch`.

- [ ] **Step 3: Validate exact section state**

Run `pptx-inspect package validate --profile powerpoint-2010` on all eight files and require zero errors/warnings. Read `/ppt/presentation.xml` and verify exact section title order, valid unique GUIDs, exact member slide IDs, no unknown member, no duplicate membership, preserved loose slides, escaped names, native absence of synthetic p15 guide list, and PptxGenJS presence of its guide extension.

- [ ] **Step 4: Verify mutation isolation and save stability**

Run package diff native-sections→native-edited and require only `/ppt/presentation.xml` plus explicitly duplicated/deleted slide dependency parts to change. Run native-edited→native-reopened and native-edited→native-second-write and require zero part changes. Same-title rename and same-position move must produce byte-identical output.

- [ ] **Step 5: Render and inspect every slide**

Use the bundled LibreOffice, `pdfinfo`, and `pdftoppm` on native sections/edited/reopened plus all PptxGenJS baselines. Require expected page counts, identical wide page size, and labels in presentation order. Run Presentations `render_slides.py` and `slides_test.py`; inspect every PNG for missing, duplicated, clipped, or reordered visible slides. Section metadata itself has no visible rendering requirement.

- [ ] **Step 6: Final repository and remote audit**

```sh
git diff --check
git diff --cached --check
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git status --short --branch
git rev-list --left-right --count origin/main...HEAD
```

Expected: only user-owned `?? .pnpm-store/`, branch `main...origin/main`, and divergence `0 0`. Do not create an empty QA commit.
