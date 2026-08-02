# Text Shape Simple Shadow Creation Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict simple-shadow creation to every public text-shape creation surface, with PptxGenJS 4.0.1 conformance and actual packed Node/browser/type/CLI proof.

**Architecture:** Reuse the existing `normalizeShapeShadow()` / `renderSimpleShadow()` value codec and existing `ShapeModel.shadow` editor. Normalize one detached shadow snapshot inside `validateAddTextOptions()`, thread it through the current text creation transaction, and append one canonical effect list after the existing line/endpoints; fill, line, arrows, and shadow ownership remain independent.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, Playwright Chromium smoke, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- `AddTextOptions.shadow` uses only the existing public `ShapeShadow` value: strict `kind: 'outer' | 'inner'`, shared color/opacity/blur/angle/distance, and outer-only `rotateWithShape`.
- Omitted/runtime-`undefined` shadow creates no effect list. There is no direct no-shadow token; PptxGenJS omitted and `{ type: 'none' }` both map to native absence.
- Defaults remain black, opacity `0.75`, blur `8pt`, angle `270°`, distance `4pt`, and outer rotate false. Explicit zero survives normalization.
- Normalize and detach shadow before package mutation. Invalid kind/object shape/accessor/symbol/unknown key/alias/color/range must leave package bytes, relationships, parts, mutation journal, shape order, and live caches unchanged.
- Plain text, rich text, placeholder creation/population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects must share one implementation.
- Placeholder population preserves owner name, geometry, identity, layout source, and sibling objects; only the populated slide owner receives the requested shadow.
- Render supplied shadow after line/endpoints as one canonical DrawingML `a:effectLst`; omitted shadow preserves current text output byte-for-byte.
- Existing `ShapeModel.shadow` read/edit/clear/no-op behavior is unchanged. Fill/line/arrows clears preserve shadow, and shadow clear preserves fill/line/arrows plus legal effect siblings.
- PptxGenJS malformed inner output, zero fallback, ignored rotate true, coercion, warning correction, and permissive values are recorded as intentional differences and never copied.
- Do not add generic effects, preset shadow, reflection, soft edge, effect DAG, custom shadow transforms, text hyperlink/geometry, `rectRadius`, `isTextBox`, or `breakLine` behavior.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, browser artifacts, or package build output.
- Use repository-local binaries. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0`.

---

### Task 1: Add strict text-shadow normalization and canonical rendering

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: public `ShapeShadow`, `normalizeShapeShadow(value, context)`, `renderSimpleShadow(shadow, prefix)`, existing `validateAddTextOptions()`, `addTextShape()`, `textShapeXml()`, and `ShapeModel.shadow`.
- Produces: `AddTextOptions.shadow?: ShapeShadow`, normalized `shadow?: NormalizedShapeShadow`, and ordered direct effect-list output for plain/rich text creation.

- [ ] **Step 1: Write failing public-model tests**

Add focused coverage for omitted, runtime `undefined`, outer/inner defaults, fully custom values, explicit zero, sRGB/theme color, and outer rotate true:

```ts
const source = {
  kind: 'outer',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0.42,
  blur: 7.25,
  angle: 123.4,
  distance: 5.5,
  rotateWithShape: true,
} as const;
const plain = slide.addText('Plain shadow', { shadow: source });
const rich = slide.addRichText([{ runs: [{ text: 'Rich shadow' }] }], {
  shadow: { kind: 'inner', opacity: 0, blur: 0, angle: 0, distance: 0 },
});
```

Require immediate deep-frozen detached snapshots, caller/nested-color detachment, exact `a:ln` → `a:effectLst` order, legal matching inner tag, same-value bytes/journal no-op, whole replacement, clear, and bidirectional fill/line/arrows/shadow ownership. Require omitted output to equal the pre-change canonical text shape bytes.

- [ ] **Step 2: Write failing invalid-input zero-mutation tests**

Before each invalid call snapshot slide bytes, relationships, package part URI list, mutation journal, shape array, and existing object identity. Reject:

```ts
{ kind: 'none' }
{ type: 'outer' }
{ kind: 'outer', offset: 4 }
{ kind: 'inner', rotateWithShape: true }
{ kind: 'outer', opacity: '0.5' }
{ kind: 'outer', blur: -1 }
{ kind: 'outer', angle: 360 }
{ kind: 'outer', distance: 201 }
```

Also reject invalid colors, symbols, unknown keys, custom prototypes, inherited-only fields, and accessor properties without invoking their getters. Run representative invalid values through both `addText()` and `addRichText()`.

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts -t "text with strict direct shadow|invalid text shadow" --reporter=dot
```

Expected: compile/runtime failure because `AddTextOptions.shadow` and text effect-list rendering are absent.

- [ ] **Step 4: Implement the minimal model change**

In `packages/model/src/slide.ts`, import the existing renderer and normalized type, then extend the public and normalized values:

```ts
import {
  normalizeShapeShadow,
  renderSimpleShadow,
  type NormalizedShapeShadow,
} from './simple-shadow.internal.js';

export interface AddTextOptions extends Partial<Transform> {
  readonly shadow?: ShapeShadow;
  // retain every existing field
}

interface NormalizedAddTextOptions {
  readonly shadow?: NormalizedShapeShadow;
  // retain every existing field
}

const shadow = options.shadow === undefined
  ? undefined
  : normalizeShapeShadow(options.shadow, 'Text shape shadow');
```

Thread `shadow` through `NormalizedTextInput`, `addText()`, both placeholder paths, `addRichText()`, `addTextShape()`, and `textShapeXml()`. Render only when present:

```ts
const effectXml = shadow === undefined
  ? ''
  : `<a:effectLst>${renderSimpleShadow(shadow, 'a:')}</a:effectLst>`;
```

Append `effectXml` immediately after `lineXml`. Do not add a second codec or invoke the live setter after creation.

- [ ] **Step 5: Run model regression gates**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/simple-shadow.internal.test.ts packages/model/src/shape-shadow.internal.test.ts packages/model/src/shape-arrows.internal.test.ts packages/model/src/simple-line.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

Expected: all pass; omitted text output, preset-shape shadows, text arrows, and line editing remain unchanged.

- [ ] **Step 6: Review, commit, push, and verify**

Review every changed line against shadow-only scope, validation before mutation, default-byte stability, line/endpoints→effect-list ordering, matching inner tags, and absence of a second mutation. Stage only the two files, commit `feat: create text shape shadows`, push `main:main`, fetch remote `main`, and require `git rev-list --left-right --count HEAD...FETCH_HEAD` to print `0 0`.

---

### Task 2: Prove placeholder, master/layout, root-package, and lifecycle coverage

**Files:**
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddTextOptions.shadow`, `AddPlaceholderOptions`, `SlideMasterObject`, `SlideLayoutModel` / `SlideMasterModel`, and `PptxDocument` lifecycle APIs.
- Produces: strict declarative option acceptance plus end-to-end public coverage for every text owner and presentation format.

- [ ] **Step 1: Add failing SDK surface tests**

Create shadows across slide, layout, master, rich text, placeholders, placeholder population, and declarative named layouts:

```ts
layout.addText('Layout shadow', {
  name: 'layout_shadow',
  shadow: { kind: 'outer', color: { kind: 'scheme', value: 'accent2' } },
});
master.addRichText([{ runs: [{ text: 'Master shadow' }] }], {
  name: 'master_shadow',
  shadow: { kind: 'inner', opacity: 0.5 },
});
layout.addPlaceholder('Prompt', {
  name: 'title_shadow',
  type: 'title',
  shadow: { kind: 'outer', rotateWithShape: true },
});
```

Require layout source isolation after slide population, stable wrapper/shape identity, duplicate independence, outer rollback, asynchronous declarative detachment, write/reopen, and every `PRESENTATION_FORMAT_PROFILES` format.

- [ ] **Step 2: Add strict declarative and root-package tests**

Add `'shadow'` to `TEXT_OPTION_KEYS`. Pass invalid text/placeholder shadows through `defineSlideMaster()` with and without a delayed image source; require rejection before observable package mutation.

In `packages/pptx/src/index.test.ts`, import `AddTextOptions` and `ShapeShadow`, compile valid outer/inner/zero values, create/reopen them from `@jiayunxie/pptx`, and add exact negative cases:

```ts
const typedTextShadow: ShapeShadow = { kind: 'outer', rotateWithShape: true };
const typedTextOptions: AddTextOptions = { shadow: typedTextShadow };
// @ts-expect-error PptxGenJS aliases are intentionally unsupported
const invalidTextShadowAlias: AddTextOptions = { shadow: { type: 'outer' } };
// @ts-expect-error inner shadow cannot rotate with the shape
const invalidInnerRotate: AddTextOptions = { shadow: { kind: 'inner', rotateWithShape: true } };
```

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shadow|text shape shadow" --reporter=dot
```

Expected: declarative definition rejects `shadow` until the closed key set is extended; other public tests exercise Task 1.

- [ ] **Step 4: Implement declarative acceptance only**

Add exactly one key to `TEXT_OPTION_KEYS`:

```ts
'shadow',
```

Do not add custom cloning: `readOptions()` / `cloneDataValue()` already detach nested values before asynchronous source preparation, and Task 1 performs semantic validation before package mutation.

- [ ] **Step 5: Run SDK/root/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review placeholder source isolation, asynchronous detachment, layout/master wrapper identity, compile-time surface, six-format persistence, and no unrelated option expansion. Commit `test: cover public text shape shadows`, push, fetch, and require divergence `0 0`.

---

### Task 3: Lock PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()` / `write()`, adapter import, native `PptxDocument`, and live `ShapeModel.shadow`.
- Produces: exact supported-case and intentional-divergence evidence for text shape simple shadow.

- [ ] **Step 1: Add a public-only PptxGenJS fixture**

Create text shapes through public PptxGenJS only for omitted, `type:none`, outer default, custom outer, zero values, rotate true, inner, missing/invalid type, hash color, coercible angle/opacity, and invalid range behavior:

```ts
generatedSlide.addText('Outer shadow', {
  objectName: 'Outer shadow',
  shadow: {
    type: 'outer',
    color: '112233',
    opacity: 0.4,
    blur: 2.5,
    angle: 45,
    offset: 3,
    rotateWithShape: true,
  },
});
```

Inspect exact `a:effectLst`, shadow tag, attributes, color/alpha, child order, and package well-formedness. Record that PptxGenJS inner uses an outer closing tag rather than importing the malformed slide as a legal semantic fixture.

- [ ] **Step 2: Add native comparison and strict-divergence assertions**

Create supported native equivalents with `AddTextOptions.shadow`; compare name, text, transform, geometry, outer snapshot, and raw order. Require native inner to be well-formed, explicit zeros to remain zero, rotate true to remain true, theme color to render, and aliases/coercions/invalid values to reject with exact zero mutation.

- [ ] **Step 3: Run focused conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text shape shadow" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape shadow|text shadow|strict direct shadow" --reporter=dot
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

Review that generation uses public PptxGenJS only, assertions distinguish supported final semantics from malformed/permissive differences, invalid native inputs stay strict, and no adapter production code changed. Commit `test: compare text shape shadows with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 4: Prove actual packed Node, browser, declaration, and CLI behavior

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball, Node/browser conditional exports, generated declarations, and installed `pptx-inspect` CLI.
- Produces: `textShapeShadows: true` in Node and real-Chrome smoke output plus a validated packed deck.

- [ ] **Step 1: Extend packed Node and declaration smoke**

Using only the installed tarball, create plain/rich/placeholder/layout/master/declarative text with outer/inner/default/custom/zero and combined fill+line+arrows+shadow. Read immediate snapshots, mutate caller input, duplicate, replace/clear through `ShapeModel.shadow`, clear line/arrows independently, write/reopen, and save `text-shape-shadows-smoke.pptx`. Add `textShapeShadows` to final JSON. In the generated TypeScript consumer, import `AddTextOptions` and `ShapeShadow`, compile both valid branches, and add `@ts-expect-error` cases matching Task 2.

- [ ] **Step 2: Extend real-Chrome smoke**

In `page.evaluate`, create/reopen a Blob-backed deck with plain, rich, placeholder, and combined style+shadow cases. Return exact immediate/detached/reopen snapshots plus validation count. Add expected `textShapeShadows` output and require zero console, page, and network errors.

- [ ] **Step 3: Add installed CLI validation**

Run installed `pptx-inspect package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and `part read` on `text-shape-shadows-smoke.pptx`. Require 0 errors/0 warnings, expected slide/shape counts, and raw outer/inner/effect-list/zero/theme/rotate plus line→endpoints→effect order evidence.

- [ ] **Step 4: Build and pack without repository artifacts**

```sh
cd packages/pptx
../../node_modules/.bin/tsup
../../node_modules/.bin/tsup --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
text_shadow_pack_dir="$(mktemp -d /tmp/pptx-text-shadow-pack.XXXXXX)"
npm pack --ignore-scripts --pack-destination "$text_shadow_pack_dir"
cd ../..
node scripts/smoke-npm-package.mjs "$text_shadow_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require exactly 57 tarball files unless a reviewed declaration closure change explains a new count. The installed smoke must reject workspace protocol/internal runtime imports and report `textShapeShadows: true`.

- [ ] **Step 5: Run type/build gates**

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-tarball isolation, Node/browser parity, declaration negatives, CLI profile result, existing smoke fields, and absence of temporary repository artifacts. Commit `test: verify packed text shape shadows`, push, fetch, and require divergence `0 0`.

---

### Task 5: Close documentation and full release gates

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: implementation, conformance, packed, browser, validator, and final-suite evidence from Tasks 1–4.
- Produces: accurate public support language and selection of text-shape hyperlink creation as the next advanced-text subitem.

- [ ] **Step 1: Update public API examples and contracts**

Add outer/inner/placeholder `shadow` examples, exact `ShapeShadow` rules, defaults/explicit-zero semantics, effect ownership, lifecycle behavior, and layout/master/declarative coverage. Do not imply generic effects, text hyperlink/geometry, or non-text-owner shadow support.

- [ ] **Step 2: Update compatibility and progress ledgers**

Move simple text-shape shadow creation from the `slide.addText` gap column into supported behavior. Record PptxGenJS omitted/none, zero fallback, rotate ignored, correction/coercion, and malformed inner differences; record final test counts and packed/browser/CLI results; retain hyperlink/geometry and advanced-effect gaps. Set the next small item to text-shape hyperlink creation.

- [ ] **Step 3: Run stale-language and Markdown checks**

Search all six documents for statements that still list text simple shadow as unsupported or next. Inspect the changed compatibility table and run:

```sh
git diff --check
rg -n -i "text[- ]shape shadow.*(unsupported|pending|未支持)|next.*text[- ]shape shadow|下一.*text[- ]shape shadow" README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

- [ ] **Step 4: Run final release gates**

```sh
node_modules/.bin/vitest run
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
```

Rerun both tsup builds, declaration closure, final actual tarball smoke, real Chromium smoke, and installed CLI profile. Require `textShapeShadows: true`, Chromium zero console/page/network errors, and PowerPoint 2010 validation 0 errors / 0 warnings.

- [ ] **Step 5: Review, commit, push, and verify**

Review every claim against current command output and raw OOXML evidence. Stage only the six documentation files, commit `docs: document text shape shadow support`, push, fetch, require divergence `0 0`, and report completed item, remaining advanced-text gaps, and updated progress.

## Plan Self-Review

- Spec coverage: Task 1 owns public normalization/rendering; Task 2 covers every owner and lifecycle; Task 3 locks PptxGenJS evidence; Task 4 proves the packed package; Task 5 closes documentation and release gates.
- Placeholder scan: every step names exact files, interfaces, test cases, commands, expected outcomes, commit messages, and remote checks; there are no deferred implementation markers.
- Type consistency: all tasks use existing `ShapeShadow`, `NormalizedShapeShadow`, `AddTextOptions.shadow`, `ShapeModel.shadow`, and the new smoke field `textShapeShadows` consistently.
- Scope check: generic effects, advanced shadow transforms, hyperlink, geometry, and non-text owner APIs remain explicitly outside this vertical slice.
