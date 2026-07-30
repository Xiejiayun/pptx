# Hidden Slide State Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add PptxGenJS-compatible hidden-slide creation through the live slide model plus strict existing-deck reading, lossless editing, lifecycle preservation, packed-surface proof, documentation, and real-deck QA.

**Architecture:** Put namespace-aware slide-root `show` parsing and source-span mutation in one model-internal helper. Expose the semantic state as `SlideModel.hidden`, keep add/duplicate/move/delete behavior within the existing slide-part lifecycle, and prove PptxGenJS parity only through public generated OOXML.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML source spans, OPC transactions, PptxGenJS 4.0.1 public output, tsup, pnpm pack, `pptx-inspect`, LibreOffice, Poppler, and the Presentations render/overflow tools.

## Global Constraints

- The authoritative behavior baseline is PptxGenJS 4.0.1 public `slide.hidden` output and standard PresentationML `p:sld@show` semantics.
- Own only the unique slide root's exact unqualified `show` attribute; preserve qualified lookalikes, children, comments, whitespace, relationships, parts, and unknown extensions.
- Absence and recognized true tokens mean visible `false`; recognized false tokens mean hidden `true`; unsafe or unknown read state is `undefined`.
- Setter input is strict boolean. `true` writes canonical `show="0"`; `false` removes the direct override. No coercion is allowed.
- Do not add `hidden` to `AddSlideOptions`, and do not implement slide numbers, notes, custom shows, transitions, section UI state, or presentation show ranges in this slice.
- All mutation validation happens before the first package change; failure and outer transaction rollback preserve exact bytes, journal, relationships, and stable model identity.
- Node 20+ and the browser bundle must expose the same behavior from the single `@jiayunxie/pptx` package.
- Never stage or commit `.pnpm-store/`; every repository-changing task is independently reviewed, committed, pushed over SSH port 443, fetched, and verified at divergence `0 0`.
- User has authorized inline autonomous execution; do not dispatch subagents or pause for routine choices.

---

### Task 1: Add the strict slide visibility XML helper

**Files:**
- Create: `packages/model/src/slide-visibility.internal.ts`
- Create: `packages/model/src/slide-visibility.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlAttribute`, and `XmlElement` from `@pptx/lossless-xml`; `ModelParseError` from `./errors.js`.
- Produces: `readSlideHidden(xml): boolean | undefined` and `replaceSlideHidden(xml, value): boolean` for `SlideModel` integration in Task 2.

- [ ] **Step 1: Write the direct-state read tests**

Create a fixture helper with a real PresentationML namespace and test absence plus all six legal lexical forms:

```ts
const slideXml = (attributes = '', body = '<p:cSld/>') =>
  `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}"${attributes}>${body}</p:sld>`;

expect(readSlideHidden(LosslessXmlDocument.parse(slideXml()))).toBe(false);
for (const token of ['0', 'false', 'off']) {
  expect(readSlideHidden(LosslessXmlDocument.parse(slideXml(` show="${token}"`)))).toBe(true);
}
for (const token of ['1', 'true', 'on']) {
  expect(readSlideHidden(LosslessXmlDocument.parse(slideXml(` show="${token}"`)))).toBe(false);
}
```

Also require an alternate root prefix, other root attributes, comments, whitespace, a direct qualified `x:show`, and descendant `show` lookalikes to remain ignored and byte-stable during reads.

- [ ] **Step 2: Write unsafe ownership tests**

Require `undefined` for:

```ts
[
  '<p:sld xmlns:p="urn:wrong"/>',
  `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" show=""/>`,
  `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" show="maybe"/>`,
  `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" show="0" show="1"/>`,
  `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
  `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}"/><p:sld xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
]
```

Capture `xml.source` and `xml.changed` before each read and require exact no mutation. Malformed XML remains a parser error rather than an `undefined` state.

- [ ] **Step 3: Write mutation, repair, and isolation tests**

Cover the canonical transitions:

```ts
const visible = LosslessXmlDocument.parse(slideXml(' custom="KEEP"'));
expect(replaceSlideHidden(visible, true)).toBe(true);
expect(visible.serialize()).toContain(' custom="KEEP" show="0"');

const hidden = LosslessXmlDocument.parse(slideXml(' custom="KEEP" show="0"'));
expect(replaceSlideHidden(hidden, false)).toBe(true);
expect(hidden.serialize()).toBe(slideXml(' custom="KEEP"'));
```

Require `false`, byte identity, and `xml.changed === false` for absent→false and canonical `0`→true. Require `false`/`off`/unknown→true to replace only the value with `0`, `1`/`true`/`on`/unknown→false to remove only the attribute, and qualified `x:show` plus every non-owned byte to survive. Duplicate attributes, duplicate namespace declarations, wrong namespace, and multiple roots must throw `ModelParseError` before any patch.

- [ ] **Step 4: Run the new test file and verify the expected failure**

```sh
pnpm vitest run packages/model/src/slide-visibility.internal.test.ts --reporter=dot
```

Expected: FAIL because `slide-visibility.internal.ts` does not exist.

- [ ] **Step 5: Implement strict root resolution and reading**

Implement the helper with separate structural and lexical decisions:

```ts
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const TRUE_TOKENS = new Set(['1', 'true', 'on']);
const FALSE_TOKENS = new Set(['0', 'false', 'off']);

interface SlideVisibilityState {
  readonly root: XmlElement;
  readonly show?: XmlAttribute;
}

export function readSlideHidden(xml: LosslessXmlDocument): boolean | undefined {
  const state = resolveSlideVisibilityState(xml);
  if (!state) return undefined;
  if (!state.show) return false;
  if (FALSE_TOKENS.has(state.show.value)) return true;
  if (TRUE_TOKENS.has(state.show.value)) return false;
  return undefined;
}
```

`resolveSlideVisibilityState()` must require one root, exact `sld` local name, namespace resolution through the root's lexical prefix, and at most one exact attribute whose `name === 'show'`. Implement local `lexicalPrefix()` and ancestor-walking `namespaceUriForPrefix()` helpers matching the strict section parser; duplicate matching namespace declarations return unsafe.

- [ ] **Step 6: Implement minimal mutation**

```ts
export function replaceSlideHidden(
  xml: LosslessXmlDocument,
  value: boolean,
): boolean {
  const state = resolveSlideVisibilityState(xml);
  if (!state) throw new ModelParseError('Slide visibility is not safely editable');
  if (value) {
    if (state.show?.value === '0') return false;
    if (state.show) xml.replaceAttribute(state.show, '0');
    else insertRootAttribute(xml, state.root, ' show="0"');
    return true;
  }
  if (!state.show) return false;
  removeRootAttribute(xml, state.root, state.show);
  return true;
}
```

`insertRootAttribute()` inserts immediately before `/>` or `>` without changing other bytes. `removeRootAttribute()` expands left only across horizontal spaces/tabs inside the start tag, then removes exactly that span. It must never remove newlines, neighboring attributes, or qualified lookalikes.

- [ ] **Step 7: Run focused and model regression tests**

```sh
pnpm vitest run \
  packages/model/src/slide-visibility.internal.test.ts \
  packages/model/src/model.test.ts --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 8: Review, commit, push, and verify**

Review namespace ownership, exact unqualified attribute matching, six legal tokens, foreign-state preservation, canonical no-ops, unknown-token repair, unsafe zero mutation, and absence of public exports. Then:

```sh
git add -- packages/model/src/slide-visibility.internal.ts \
  packages/model/src/slide-visibility.internal.test.ts
git diff --cached --check
git commit -m "feat: add hidden slide state codec"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Expose hidden state and preserve it through slide lifecycle operations

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `readSlideHidden()` and `replaceSlideHidden()`.
- Produces: public `SlideModel.hidden: boolean | undefined` getter and boolean setter, automatically visible from SDK and aggregate declarations.

- [ ] **Step 1: Add model-facing creation and editing tests**

In `packages/model/src/model.test.ts`, create a presentation with one visible slide and require:

```ts
const slide = document.addSlide();
expect(slide.hidden).toBe(false);
const sameSlide = document.slides[0];
slide.hidden = true;
expect(slide.hidden).toBe(true);
expect(document.slides[0]).toBe(sameSlide);
slide.hidden = false;
expect(slide.hidden).toBe(false);
```

Read the slide part and require `show="0"` only in the hidden state. Capture part fingerprints and journal to prove visible→visible and canonical hidden→hidden exact no-ops. Inject each legal alternate token and verify the getter; inject one unknown token and verify `undefined`, then explicit true/false repair. Add invalid runtime values and prove zero mutation.

- [ ] **Step 2: Add SDK lifecycle, rollback, and six-format tests**

In `packages/sdk/src/index.test.ts`, combine visibility with sections and slide topology:

```ts
const document = PptxDocument.create();
const visible = document.addSlide();
const hidden = document.addSlide();
hidden.hidden = true;
const section = document.addSection({ title: 'Hidden section' });
document.assignSlideToSection(1, section.id);
const duplicate = document.duplicateSlide(1);
expect(duplicate.hidden).toBe(true);
document.moveSlide(document.slides.indexOf(duplicate), 0);
expect(document.slides.map(({ hidden }) => hidden)).toEqual([true, false, true]);
```

Delete the original hidden slide and require the duplicate state plus section membership to remain valid. In an outer `document.transaction()`, change visible/hidden states, duplicate/move/delete, throw, and require exact restoration of parts, journal, slides, identities, order, and membership. For each `pptx/pptm/ppsx/ppsm/potx/potm`, write/reopen one visible and one hidden slide and require `[false, true]`.

- [ ] **Step 3: Run tests and verify failure before integration**

```sh
pnpm vitest run \
  packages/model/src/model.test.ts -t 'hidden slide' \
  packages/sdk/src/index.test.ts -t 'hidden slide' --reporter=dot
```

Expected: FAIL because `SlideModel.hidden` is not defined.

- [ ] **Step 4: Add the live property to `SlideModel`**

Add imports and the public accessor near the other slide-level state:

```ts
get hidden(): boolean | undefined {
  const { xml } = this.parse();
  return readSlideHidden(xml);
}

set hidden(value: boolean) {
  if (typeof value !== 'boolean') {
    throw new TypeError('Slide hidden state must be a boolean');
  }
  this.presentation.opcPackage.transaction(() => {
    const { xml } = this.parse();
    if (!replaceSlideHidden(xml, value)) return;
    this.setXml(xml.serialize());
  });
}
```

Do not add new SDK wrappers, create options, cache entries, or exports: `SlideModel` is already public and returned by `PptxDocument.slides/addSlide/duplicateSlide`.

- [ ] **Step 5: Run focused, dependency, and full type gates**

```sh
pnpm vitest run \
  packages/model/src/slide-visibility.internal.test.ts \
  packages/model/src/model.test.ts -t 'hidden slide|duplicate slide|delete slide|move slide' \
  packages/sdk/src/index.test.ts -t 'hidden slide|presentation section|dependency|rollback' \
  --reporter=dot
pnpm typecheck
git diff --check
```

Expected: all pass.

- [ ] **Step 6: Review, commit, push, and verify**

Review strict input handling before mutation, invalid-token repair, visible/hidden semantic direction, no-op behavior, stable identity, all six formats, section orthogonality, duplicate preservation, move/delete behavior, and rollback. Then:

```sh
git add -- packages/model/src/slide.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit hidden slide state"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Compare only public PptxGenJS hidden-slide output

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `new PptxGenJS()`, `addSlide()`, returned `slide.hidden`, and `write()`; native `SlideModel.hidden` from Task 2.
- Produces: public conformance proof for omitted, false, true, truthy invalid, native strictness, and write/reopen.

- [ ] **Step 1: Extend the local public test interface**

Add only the public property used by the test:

```ts
interface PptxGenJSSlide {
  hidden: unknown;
  // existing public addText/addTable declarations remain unchanged
}
```

Do not expose `_hidden`, `_slides`, ZIP helpers, or generated package internals.

- [ ] **Step 2: Add public-output cases**

Create four independent presentations using public APIs:

```ts
const omitted = new PptxGenJS();
omitted.addSlide();

const explicitFalse = new PptxGenJS();
explicitFalse.addSlide().hidden = false;

const explicitTrue = new PptxGenJS();
explicitTrue.addSlide().hidden = true;

const truthyInvalid = new PptxGenJS();
truthyInvalid.addSlide().hidden = 'yes';
```

Import each through `importPptxGenJS()` and require hidden snapshots `[false, false, true, true]`, version `4.0.1`, valid final packages, and native write/reopen stability. Read imported slide XML only through `opcPackage` after public output and require `show="0"` only for true/truthy final states.

- [ ] **Step 3: Add native parity and intentional repair assertions**

Create native visible and hidden slides through public APIs, compare their final hidden snapshots and root `show` state to valid PptxGenJS false/true outputs, then prove native rejects `'yes'` before changing bytes or journal:

```ts
const native = PptxDocument.create();
const visible = native.addSlide();
const hidden = native.addSlide();
hidden.hidden = true;
expect([visible.hidden, hidden.hidden]).toEqual([false, true]);
expect(() => { (visible as unknown as { hidden: unknown }).hidden = 'yes'; }).toThrow(TypeError);
```

- [ ] **Step 4: Run the focused conformance test**

```sh
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'hidden slide' --reporter=dot
```

Expected: pass after Task 2; the test must use no private field.

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only construction, exact omitted/false/true states, truthy runtime defect documentation, native strict rejection, adapter import, and reopen. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare hidden slides with pptxgenjs"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: aggregate-package `PptxDocument` and inferred public `SlideModel` declarations.
- Produces: actual tarball runtime/type proof and JSON check `hiddenSlides: true`.

- [ ] **Step 1: Add packed Node runtime coverage**

In the generated Node consumer, create visible and hidden slides, duplicate and edit them, write/reopen, and inspect only public state:

```js
const hiddenDeck = PptxDocument.create();
const packedVisibleSlide = hiddenDeck.addSlide();
const packedHiddenSlide = hiddenDeck.addSlide();
packedHiddenSlide.hidden = true;
const packedHiddenDuplicate = hiddenDeck.duplicateSlide(1);
packedVisibleSlide.hidden = true;
packedHiddenSlide.hidden = false;
const reopenedHiddenStates = (await PptxDocument.open(await hiddenDeck.write()))
  .slides.map(({ hidden }) => hidden);
```

Add `hiddenSlides` to the result object and require final `[true, false, true]`, duplicate preservation, and direct `show="0"` only on the expected parts.

- [ ] **Step 2: Add browser bundle runtime coverage**

Inside the generated browser consumer:

```js
const browserHiddenDeck = PptxDocument.create();
const browserHiddenSlide = browserHiddenDeck.addSlide();
browserHiddenSlide.hidden = true;
const reopenedBrowserHidden = await PptxDocument.open(await browserHiddenDeck.writeBlob());
if (reopenedBrowserHidden.slides[0]?.hidden !== true) {
  throw new Error('Browser hidden slide failed');
}
```

- [ ] **Step 3: Add declaration coverage**

In the TypeScript consumer:

```ts
const typedVisibilitySlide = createdDocument.addSlide();
const hiddenSnapshot: boolean | undefined = typedVisibilitySlide.hidden;
typedVisibilitySlide.hidden = true;
typedVisibilitySlide.hidden = false;
```

Include both variables in the final `void [...]` expression so `tsc --noEmit` consumes them.

- [ ] **Step 4: Build, pack, and run the actual tarball**

```sh
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
pptx_hidden_pack_dir=$(mktemp -d /tmp/pptx-hidden-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_hidden_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_hidden_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require JSON containing `"hiddenSlides":true`, `"types":true`, and `"cli":"0.1.0"`.

- [ ] **Step 5: Review, commit, push, and verify**

Review actual packaged imports, Node/browser parity, inferred declarations, duplicate/edit/reopen semantics, JSON assertion strength, and absence of workspace-only imports. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed hidden slides"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Document hidden-slide creation, editing, and parity boundaries

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: proven public behavior from Tasks 1-4.
- Produces: accurate compatibility status, API example, and remaining slide-level gap list.

- [ ] **Step 1: Update the parity matrix and focused baseline**

Add a row:

```md
| `slide.hidden` | `SlideModel.hidden` | 已支持 strict boolean create/edit、direct root read/repair、duplicate/move/reopen；native拒绝truthy coercion |
```

Add one paragraph covering PptxGenJS omitted/false/true/truthy output, inverted `show` semantics, six accepted lexical tokens, native visible absence, unsafe `undefined`, qualified lookalike preservation, and public-output adapter import. Keep slide numbers, notes, masters/layouts/placeholders, custom shows, shapes/images/charts/media, and advanced tables pending.

- [ ] **Step 2: Update package README and changelog**

Add a compact README example:

```ts
const hiddenSlide = document.addSlide();
hiddenSlide.hidden = true;
hiddenSlide.hidden = false;
```

Explain absence/true-token as visible false, false-token as hidden true, unsafe as undefined, strict boolean setter, duplicate preservation, move independence, and section orthogonality. Add one changelog bullet without claiming slide numbers, speaker notes, or custom shows.

- [ ] **Step 3: Run documentation and type gates**

```sh
rg -n 'hidden|show="0"|SlideModel.hidden|truthy' \
  CHANGELOG.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
rg -n 'hidden.*(pending|unsupported|尚未支持|仍待)' \
  CHANGELOG.md docs/compatibility packages/pptx/README.md || true
pnpm typecheck
git diff --check
```

Classify every match and retain only statements about capabilities genuinely still pending.

- [ ] **Step 4: Review, commit, push, and verify**

Review semantic inversion, strict/unsafe states, lifecycle claims, packed proof, and explicit remaining boundaries. Then:

```sh
git add -- CHANGELOG.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "docs: document hidden slides"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Run full release and real-deck QA

**Files:**
- No repository changes expected.
- Generate only under `/tmp/pptx-hidden-qa-20260730` and a `mktemp` pack directory.

**Interfaces:**
- Consumes: the complete hidden-slide slice and installed `pptx-inspect`/LibreOffice/Presentations QA tools.
- Produces: release-gate, structural, mutation-isolation, save-stability, render, and remote-sync evidence.

- [ ] **Step 1: Run complete repository and package gates**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
pnpm --filter @jiayunxie/pptx build
pptx_hidden_qa_pack_dir=$(mktemp -d /tmp/pptx-hidden-qa-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_hidden_qa_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_hidden_qa_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require all tests, the 5-second performance budget, actual Node/browser/types tarball smoke, and CLI smoke to pass.

- [ ] **Step 2: Generate representative public-API decks**

Using built public APIs only, generate:

```text
native-visible.pptx
native-hidden.pptx
native-mixed.pptx
native-edited.pptx
native-reopened.pptx
native-second-write.pptx
pptxgenjs-hidden.pptx
pptxgenjs-mixed.pptx
```

Use wide layout and numbered visible labels. Native mixed state must combine visible/hidden slides, sections, hidden duplicate, move, visible/hidden edits, and deletion without changing expected content or memberships. Keep generator/verifier `.mjs` files under `/tmp` and create them with `apply_patch`.

- [ ] **Step 3: Validate exact root state and package isolation**

Run `pptx-inspect --json package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and exact `/ppt/slides/slideN.xml` reads. Require 0 errors/warnings, exact slide order, `show="0"` only on hidden slides, no attribute for native visible slides, unchanged section memberships, and PptxGenJS hidden final state importable as true.

Run package diffs:

```text
native-mixed → native-edited
native-edited → native-reopened
native-edited → native-second-write
```

Require the first diff to change only explicitly edited slide parts and requested topology parts; require the latter two to have zero added/removed/changed parts. Same-value visible/hidden assignments must be byte/journal no-ops in the verifier.

- [ ] **Step 4: Render and inspect every slide**

Use bundled LibreOffice, `pdfinfo`, `pdftoppm`, Presentations `render_slides.py`, and `slides_test.py`. Verify wide page size, expected page/slide handling with hidden slides documented per client, and every rendered PNG at full size for missing, duplicated, clipped, reordered, or changed labels. Do not infer OOXML hidden correctness only from PDF page count.

- [ ] **Step 5: Final repository and remote audit**

```sh
git diff --check
git diff --cached --check
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git status --short --branch
git rev-list --left-right --count origin/main...HEAD
```

Expected: only user-owned `?? .pnpm-store/`, branch `main...origin/main`, and divergence `0 0`. Do not create an empty QA commit.
