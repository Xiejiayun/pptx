# Slide Speaker Notes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add PptxGenJS-compatible plain speaker notes creation plus strict existing-deck reading, editing, empty/clear semantics, notes-master repair, lifecycle preservation, packed-surface proof, documentation, and real-deck QA.

**Architecture:** Put notes-part ownership, namespace-aware body-placeholder text handling, canonical notes slide/master creation, and relationship mutation in one model-internal helper. Expose a live `SlideModel.notes` property plus `addNotes(string)`, reuse the existing owned dependency lifecycle for duplicate/delete, and prove parity only through PptxGenJS public output.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML source spans, OPC relationships/transactions, PptxGenJS 4.0.1 public output, tsup, pnpm pack, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- The authoritative behavior baseline is PptxGenJS 4.0.1 public `slide.addNotes(string)` output and standard PresentationML notesSlide/notesMaster relationships.
- Native absent notes remain lazy `undefined`; explicit empty is a materialized empty notes body `''`; non-empty string is plain whole-value content.
- String inputs normalize CRLF/CR to LF, preserve leading/trailing whitespace, escape XML metacharacters, and reject XML 1.0 illegal controls before mutation.
- `undefined` removes only the selected slide's notesSlide relationship and owned per-slide notes part; shared notes master/theme and other slides remain.
- Own only the unique internal notesSlide relationship, correct notes content type/root, unique direct body placeholder, its direct text body paragraphs, and a canonical notes-master chain created only from fully absent safe state.
- Preserve unknown notes shapes, body properties, placeholder metadata, extensions, relationships, notes master/theme bytes, slides, sections, hidden state, and all unrelated parts.
- Do not implement rich notes, notes page layout editing, comments, header/footer/date editing, slide-number controls, custom shows, transitions, or show ranges in this slice.
- All validation precedes the first package change; failures and outer rollback preserve exact bytes, relationships, content types, journal, and stable model identity.
- Node 20+ and the browser bundle expose the same behavior from the single `@jiayunxie/pptx` package.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed over SSH port 443, fetched, and verified at divergence `0 0`.
- User has authorized inline autonomous execution; do not dispatch subagents or pause for routine choices.

---

### Task 1: Add the strict notes body XML codec

**Files:**
- Create: `packages/model/src/slide-notes.internal.ts`
- Create: `packages/model/src/slide-notes.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, XML escaping utilities, and `ModelParseError`.
- Produces: `normalizeSlideNotes(value, context)`, `readNotesBody(xml)`, `replaceNotesBody(xml, value, partUri)`, and `createNotesSlideXml(value)` for package integration in Task 2.

- [ ] **Step 1: Write normalization and read tests**

Create strict fixtures with real PresentationML/DrawingML namespaces and a direct notes body placeholder:

```ts
expect(normalizeSlideNotes('A\r\nB\rC', 'Slide notes')).toBe('A\nB\nC');
expect(() => normalizeSlideNotes(7, 'Slide notes')).toThrow(TypeError);
expect(() => normalizeSlideNotes('A\u0000B', 'Slide notes')).toThrow(TypeError);

expect(readNotesBody(parse(notesXml('')))).toBe('');
expect(readNotesBody(parse(notesXml('Speaker &amp; notes')))).toBe('Speaker & notes');
expect(readNotesBody(parse(notesXmlWithParagraphs()))).toBe('First\nsoft\nbreak\nSecond');
```

Cover alternate valid prefixes, field/rich runs, `a:br`, empty/self-closing text, `xml:space`, leading/trailing whitespace, and unrelated slide-image/slide-number/unknown shapes. Capture `xml.source` and `xml.changed` and require reads to be exact no-ops.

- [ ] **Step 2: Write unsafe ownership tests**

Require `undefined` for wrong root namespace/name, multiple roots, missing/repeated direct `cSld` or `spTree`, duplicate body placeholders, qualified `x:type="body"` lookalikes, duplicate direct `txBody`, and a body shape whose placeholder chain is ambiguous. Malformed XML remains a parser error.

Require mutation to throw `ModelParseError` before patching for the same unsafe states. A safe shape tree without a body placeholder is readable as `undefined` but is repairable by string mutation.

- [ ] **Step 3: Write replace, insertion, no-op, and canonical creation tests**

```ts
const existing = parse(notesXml('Before', '<p:extLst><x:keep/></p:extLst>'));
expect(replaceNotesBody(existing, 'After\nLine 2', '/ppt/notesSlides/notesSlide1.xml')).toBe(true);
expect(readNotesBody(existing)).toBe('After\nLine 2');
expect(existing.serialize()).toContain('<x:keep/>');

const same = parse(notesXml('Same'));
expect(replaceNotesBody(same, 'Same', '/ppt/notesSlides/notesSlide1.xml')).toBe(false);
expect(same.changed).toBe(false);

const created = LosslessXmlDocument.parse(createNotesSlideXml('A & <B>'));
expect(readNotesBody(created)).toBe('A & <B>');
```

Require whole-value replacement to preserve direct `bodyPr`, `lstStyle`, placeholder metadata, shape properties, unknown sibling shapes/comments/extensions, and root bytes outside body paragraphs. Missing-body insertion must allocate the next safe `cNvPr@id` and preserve all existing IDs. Duplicate, non-integer, negative, unsafe, or exhausted IDs reject before mutation.

- [ ] **Step 4: Run the focused test and verify expected failure**

```sh
pnpm vitest run packages/model/src/slide-notes.internal.test.ts --reporter=dot
```

Expected: FAIL because `slide-notes.internal.ts` does not exist.

- [ ] **Step 5: Implement normalization and strict body resolution**

Implement namespace resolution through the lexical root/body prefixes and exact direct-child navigation. Plain text flattening must use direct paragraph order, recursively concatenate text under direct runs/fields, emit LF for `br`, and emit one LF between paragraphs.

```ts
export function normalizeSlideNotes(value: unknown, context: string): string {
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value.replace(/\r\n?/g, '\n');
}
```

Do not locate body by shape name, index, text, or position.

- [ ] **Step 6: Implement minimal body replacement and canonical XML**

Render a deterministic single paragraph/run/text with `xml:space="preserve"`; replace only the direct paragraph span inside `txBody`. For a missing body, insert a canonical `p:sp` immediately before the shape-tree end tag. `createNotesSlideXml()` writes the unique `p:notes/p:cSld/p:spTree` chain, group properties, and body placeholder without random IDs or timestamps.

- [ ] **Step 7: Run focused and model regression gates**

```sh
pnpm vitest run \
  packages/model/src/slide-notes.internal.test.ts \
  packages/model/src/model.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 8: Review, commit, push, and verify**

Review namespace ownership, exact direct body selection, plain flattening, line-ending normalization, invalid XML rejection, missing-body insertion, same-value no-op, unknown-byte preservation, and absence of public exports. Then:

```sh
git add -- packages/model/src/slide-notes.internal.ts \
  packages/model/src/slide-notes.internal.test.ts
git diff --cached --check
git commit -m "feat: add speaker notes text codec"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Manage notes parts and expose the live slide API

**Files:**
- Modify: `packages/model/src/slide-notes.internal.ts`
- Modify: `packages/model/src/slide-notes.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 text codec; `OpcPackage`, `relativeRelationshipTarget()`, and existing owned dependency garbage collection.
- Produces: `readSlideNotes(pkg, presentationPartUri, slidePartUri)`, `replaceSlideNotes(...)`, public `SlideModel.notes`, and `SlideModel.addNotes(string)` for existing-master decks.

- [ ] **Step 1: Add package-state read and failure tests**

Build fixtures for absent, valid internal, duplicate, external, unresolved, missing-part, wrong-content-type, wrong-root, and shared-target notesSlide relationships. Require:

```ts
expect(readSlideNotes(pkg, '/ppt/presentation.xml', '/ppt/slides/slide1.xml')).toBe('Speaker note');
expect(readSlideNotes(absentPkg, presentationUri, slideUri)).toBeUndefined();
expect(readSlideNotes(duplicatePkg, presentationUri, slideUri)).toBeUndefined();
```

Every read preserves all bytes, relationships, content types, and journal.

- [ ] **Step 2: Add native create/edit/empty/clear tests**

In `model.test.ts`:

```ts
const slide = model.addSlide();
expect(slide.notes).toBeUndefined();
expect(slide.addNotes('First\r\nSecond')).toBe(slide);
expect(slide.notes).toBe('First\nSecond');
slide.notes = 'Edited';
expect(slide.notes).toBe('Edited');
slide.notes = '';
expect(slide.notes).toBe('');
slide.notes = undefined;
expect(slide.notes).toBeUndefined();
```

Verify the notes part content type, slide→notesSlide, notesSlide→slide, and notesSlide→notesMaster relationships. Require same normalized string and absent clear to preserve exact parts/journal. Invalid runtime input to the property and `addNotes()` must throw before mutation.

- [ ] **Step 3: Add lifecycle, rollback, sections/hidden, and six-format tests**

Create visible/hidden slides in sections, add notes, duplicate a noted slide, move it, edit duplicate notes, and delete the source. Require the duplicate notes part to differ from source, point back to duplicate slide, and share the same notes master. Other slides and section membership remain correct.

In an outer `document.transaction()`, edit/create/clear notes, duplicate/move/delete, then throw; require exact package/journal/identity restoration. For each `pptx/pptm/ppsx/ppsm/potx/potm`, write/reopen `[undefined, '', 'Notes']` and require exact snapshots.

- [ ] **Step 4: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/model.test.ts -t 'speaker notes' \
  packages/sdk/src/index.test.ts -t 'speaker notes' --reporter=dot
```

Expected: FAIL because the package APIs and `SlideModel.notes`/`addNotes()` are absent.

- [ ] **Step 5: Implement relationship resolution and existing-master creation**

Add exact relationship constants and safe resolution. For string create, require exactly one valid internal presentation notesMaster relationship, allocate `/ppt/notesSlides/notesSlideN.xml`, set its content type, and add the three relationships using relative targets. For clear, preflight all ownership/incoming references, remove the slide relationship, and garbage-collect the owned notes root. Return `false` for absent clear and same text.

```ts
export function readSlideNotes(
  pkg: OpcPackage,
  presentationPartUri: string,
  slidePartUri: string,
): string | undefined;

export function replaceSlideNotes(
  pkg: OpcPackage,
  presentationPartUri: string,
  slidePartUri: string,
  value: string | undefined,
): boolean;
```

- [ ] **Step 6: Expose the live API**

In `SlideModel`:

```ts
get notes(): string | undefined {
  return readSlideNotes(this.presentation.opcPackage, this.presentation.presentationPartUri, this.partUri);
}

set notes(value: string | undefined) {
  replaceSlideNotes(this.presentation.opcPackage, this.presentation.presentationPartUri, this.partUri, value);
}

addNotes(value: string): this {
  this.notes = normalizeSlideNotes(value, 'Slide notes');
  return this;
}
```

`replaceSlideNotes()` performs its own normalization so runtime misuse of the property remains strict. Avoid nested mutations before validation; rely on package transactions for rollback.

- [ ] **Step 7: Run focused, dependency, and full type gates**

```sh
pnpm vitest run \
  packages/model/src/slide-notes.internal.test.ts \
  packages/model/src/model.test.ts -t 'speaker notes|dependency|duplicate slide|delete slide' \
  packages/sdk/src/index.test.ts -t 'speaker notes|section|hidden|rollback' --reporter=dot
pnpm typecheck
pnpm test:dependency-boundaries
git diff --check
```

Expected: all pass.

- [ ] **Step 8: Review, commit, push, and verify**

Review lazy absent state, explicit empty state, relationship directions/content types, clear isolation, duplicate retarget, shared master preservation, move/delete, six formats, rollback, stable identity, and invalid zero mutation. Then:

```sh
git add -- packages/model/src/slide-notes.internal.ts \
  packages/model/src/slide-notes.internal.test.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit slide speaker notes"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Create a missing notes master chain safely

**Files:**
- Modify: `packages/model/src/slide-notes.internal.ts`
- Modify: `packages/model/src/slide-notes.internal.test.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 2 package mutation and presentation/slide-master/theme relationships.
- Produces: string assignment on a deck with fully absent notes-master state creates one canonical notes master/list/relationship chain.

- [ ] **Step 1: Add fully absent and theme-selection tests**

Start from a native package, remove `notesMasterIdLst`, the presentation notesMaster relationship, notes master part, and its relationship part. Then set notes and require:

- one new notes master part/content type;
- one presentation notesMaster relationship;
- one direct `notesMasterIdLst` with one `notesMasterId@r:id` matching the relationship;
- notes master→theme relationship using the unique direct presentation theme;
- the new notes slide referencing that master;
- all pre-existing slide/master/layout/theme bytes unchanged except presentation topology/content-types/rels.

Repeat after removing the direct presentation theme relationship and require deterministic fallback to the first presentation slide master's unique internal theme.

- [ ] **Step 2: Add unsafe and rollback tests**

Cover duplicate direct notesMaster relationships/lists/IDs, external/missing/wrong-content master targets, duplicate direct presentation themes, a first slide master with duplicate themes, and no safely reusable theme. Each must throw before mutation. Outer rollback after successful master and notes creation restores the exact absent package and journal.

- [ ] **Step 3: Run focused tests and verify expected failure**

```sh
pnpm vitest run \
  packages/model/src/slide-notes.internal.test.ts -t 'notes master' \
  packages/model/src/model.test.ts -t 'speaker notes master' \
  packages/sdk/src/index.test.ts -t 'speaker notes master' --reporter=dot
```

Expected: FAIL because Task 2 requires an existing notes master.

- [ ] **Step 4: Implement strict absent-state resolution**

Resolve the presentation root and exact direct notesMaster list. Existing valid state must have one list entry whose `r:id` resolves to the unique relationship and correct part. Any partial/ambiguous state is unsafe; only zero list + zero relationship + zero candidate notes-master ownership is creatable.

- [ ] **Step 5: Implement theme selection and canonical master creation**

Choose a theme in this exact order:

1. exactly one valid internal direct presentation theme relationship;
2. the first presentation-ordered valid slide master and its exactly one valid internal theme relationship.

Reject zero or ambiguous candidates. Create canonical notes master XML with group shape tree, color map, `hf`, and `notesStyle`; add its theme relationship. Insert `notesMasterIdLst` immediately after `sldIdLst` when present, otherwise before `sldSz`, using the actual allocated relationship ID and the presentation root prefix.

- [ ] **Step 6: Run focused and full topology gates**

```sh
pnpm vitest run \
  packages/model/src/slide-notes.internal.test.ts \
  packages/model/src/model.test.ts -t 'speaker notes|slide size|dependency' \
  packages/sdk/src/index.test.ts -t 'speaker notes|create|format' --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 7: Review, commit, push, and verify**

Review exact absent-vs-partial distinction, theme selection order, prefix/relationship ID correctness, insertion order, mutation isolation, content types, rollback, and no random data. Then:

```sh
git add -- packages/model/src/slide-notes.internal.ts \
  packages/model/src/slide-notes.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create missing notes master"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Compare only public PptxGenJS speaker-notes output

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `new PptxGenJS()`, `addSlide()`, returned slide `addNotes(string)`, and `write()`; native notes API from Tasks 2–3.
- Produces: conformance evidence for omitted, empty, plain, multiline, XML-safe, duplicate, and reopen final states.

- [ ] **Step 1: Extend the local public test interface**

Add only:

```ts
interface PptxGenJSSlide {
  addNotes(notes: string): PptxGenJSSlide;
}
```

Do not expose `_notes`, `_slideObjects`, `_slides`, ZIP helpers, or package internals.

- [ ] **Step 2: Add public-output cases**

Generate omitted, explicit empty, plain `Speaker & <notes>`, and multiline `Line 1\nLine 2\r\nLine 3` decks. Import each through `importPptxGenJS()` and require notes snapshots `["", "", "Speaker & <notes>", "Line 1\nLine 2\nLine 3"]`, version `4.0.1`, valid final packages, and native write/reopen stability.

Read relationship/notes XML only from the imported public output and verify one notesSlide/body/master chain. Record PptxGenJS's eager empty notes part without treating it as a native requirement.

- [ ] **Step 3: Compare valid native final state and strict differences**

Create native lazy absent, explicit empty, plain, and multiline notes. Require semantic equality for valid string states, native `undefined` only for lazy absent, XML escaping, and normalized LF. Prove native rejects number/object/runtime invalid controls before part or journal changes.

- [ ] **Step 4: Run the focused conformance test**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'speaker notes' --reporter=dot
```

Expected: PASS with no private access.

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only generation, eager-vs-lazy boundary, empty/omitted interpretation, CR normalization, no private types, and strict invalid handling. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare speaker notes with pptxgenjs"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball public API.
- Produces: runtime/type proof and JSON check `speakerNotes: true`.

- [ ] **Step 1: Add packed Node runtime coverage**

In the generated consumer, create lazy/empty/plain noted slides, use both property and `addNotes()`, duplicate/edit/clear, write/reopen, and inspect only public state. Require final snapshots `[undefined, '', 'Edited', 'Original']`, duplicate independence, exactly the expected notes parts, and notesSlide→slide retargeting.

- [ ] **Step 2: Add browser bundle runtime coverage**

In the Playwright bundle source:

```js
const browserNotesDeck = PptxDocument.create();
const browserNotesSlide = browserNotesDeck.addSlide().addNotes('Browser notes');
if (browserNotesSlide.notes !== 'Browser notes') throw new Error('Browser notes immediate state failed');
const reopenedBrowserNotes = await PptxDocument.open(await browserNotesDeck.writeBlob());
if (reopenedBrowserNotes.slides[0]?.notes !== 'Browser notes') {
  throw new Error('Browser speaker notes failed');
}
```

- [ ] **Step 3: Add declaration coverage**

```ts
const typedNotesSlide = createdDocument.addSlide();
const notesSnapshot: string | undefined = typedNotesSlide.notes;
typedNotesSlide.notes = 'Typed notes';
typedNotesSlide.notes = '';
typedNotesSlide.notes = undefined;
const returnedNotesSlide: SlideModel = typedNotesSlide.addNotes('Returned');
```

Add all symbols to the existing no-unused aggregate.

- [ ] **Step 4: Build, pack, and run the actual tarball**

```sh
pnpm build
pptx_notes_pack_dir=$(mktemp -d /tmp/pptx-notes-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_notes_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_notes_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require JSON containing `"speakerNotes":true`, `"types":true`, and `"cli":"0.1.0"`.

- [ ] **Step 5: Review, commit, push, and verify**

Review actual tarball use, Node/browser parity, no workspace import leakage, declaration strictness, notes relationships, and unchanged CLI behavior. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed speaker notes"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Document speaker-notes creation, editing, and parity boundaries

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

**Interfaces:**
- Consumes: complete notes behavior and conformance evidence.
- Produces: public usage, parity status, and explicit non-goals.

- [ ] **Step 1: Update the parity matrix and focused baseline**

Add:

```md
| `slide.addNotes(string)` | `SlideModel.addNotes()` / `SlideModel.notes` | 已支持 plain string create/read/edit/empty/clear、lazy native state、notes-master repair、duplicate/move/reopen |
```

Explain PptxGenJS eager empty notes parts, native lazy absent, exact `''`/`undefined` distinction, CR normalization, unique body ownership, missing-master creation, strict invalid handling, and adapter public-output boundary. Remove speaker notes from the remaining slide-level gap sentence without removing slide numbers/custom shows/show ranges.

- [ ] **Step 2: Update API README, npm README, and changelog**

Use this public example:

```ts
const slide = document.addSlide();
slide.addNotes('Opening context\nKey talking point');
slide.notes = 'Revised talking point';
slide.notes = '';
slide.notes = undefined;
```

Document empty-vs-clear, duplicate behavior, notes-master preservation, and plain-only scope. Add one changelog bullet without claiming rich notes, notes page layout, comments, slide numbers, or custom shows.

- [ ] **Step 3: Run documentation and type gates**

```sh
rg -n 'addNotes|SlideModel.notes|speaker notes|演讲者备注' \
  CHANGELOG.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
rg -n 'speaker notes.*(pending|unsupported|尚未支持|仍待)' \
  packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
git diff --check
```

The second search may only find explicit rich/layout non-goals, not plain speaker notes.

- [ ] **Step 4: Review, commit, push, and verify**

Review API names, exact semantics, no overclaim, removed stale gap text, migration example, and changelog scope. Then:

```sh
git add -- CHANGELOG.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document slide speaker notes"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run full release and real-deck QA

**Files:**
- No repository changes expected; generate only under `/tmp/pptx-notes-qa-20260730` and a `mktemp` pack directory.

**Interfaces:**
- Consumes: complete speaker-notes slice and installed `pptx-inspect`/LibreOffice/Presentations QA tools.
- Produces: release evidence, package isolation proof, notes extraction proof, client render/open proof, and clean synced repository.

- [ ] **Step 1: Run complete repository and package gates**

```sh
pnpm check
pnpm test:performance
pnpm build
pptx_notes_qa_pack_dir=$(mktemp -d /tmp/pptx-notes-qa-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_notes_qa_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_notes_qa_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require all tests/type gates/builds, the 5-second performance budget, and packed JSON checks to pass.

- [ ] **Step 2: Generate representative public-API decks**

Generate native lazy/empty/plain/multiline, native mixed lifecycle/sections/hidden, edited, reopened, second-write, cleared, missing-master-repaired, and PptxGenJS omitted/empty/plain/multiline decks. Use numbered visible slide labels and notes values that identify their source slide. Keep generator/verifier `.mjs` files under `/tmp` and create them with `apply_patch`.

- [ ] **Step 3: Validate relationships, notes text, and package isolation**

For every deck run `pptx-inspect --json package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and exact reads of presentation, slide, notesSlide, notesMaster, and relationship parts. Require 0 errors/warnings, exact notes snapshots, unique relationship directions, correct content types, shared master identity, duplicate retarget, normalized LF, XML escaping, unchanged sections/hidden state, and no orphan notes parts.

Run package diffs:

```sh
pptx-inspect --json package diff native-mixed.pptx native-edited.pptx
pptx-inspect --json package diff native-edited.pptx native-reopened.pptx
pptx-inspect --json package diff native-edited.pptx native-second-write.pptx
```

The first diff may change only requested notes/topology parts; the latter two require zero added/removed/changed parts. Same-value notes and absent clear must be byte/journal no-ops in the verifier.

- [ ] **Step 4: Open/render and inspect slides and notes**

Use bundled LibreOffice, `pdfinfo`, `pdftoppm`, Presentations `render_slides.py`, and `slides_test.py`. Verify slide visuals/page counts are unchanged by notes operations, every PNG at full size has no clipping/reorder/duplication, and LibreOffice opens native/PptxGenJS noted decks without repair. Extract notes text through a package verifier and, where available, client notes export; do not infer notes correctness only from slide PDF rendering.

- [ ] **Step 5: Final repository and remote audit**

```sh
git diff --check
git diff --cached --check
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git status --short --branch
git rev-list --left-right --count origin/main...HEAD
```

Expected: only untracked `.pnpm-store/`, divergence `0 0`. Do not create an empty QA commit.
