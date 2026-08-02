# Text Shape Arrows Creation Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict begin/end arrow creation to every public text-shape creation surface, with PptxGenJS 4.0.1 conformance and actual packed Node/browser/type/CLI proof.

**Architecture:** Reuse the existing `normalizeShapeArrows()` / `renderShapeArrows()` value codec and existing `ShapeModel.arrows` editor. Normalize one detached endpoint snapshot inside `validateAddTextOptions()`, thread it through the current text creation transaction, and append head/tail to the already normalized text `a:ln` after fill/dash; line paint and endpoint ownership remain independent.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, Playwright Chromium smoke, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- `AddTextOptions.arrows` uses only the existing public `ShapeArrows` value: optional `begin` / `end`, each one of `none | arrow | diamond | oval | stealth | triangle`.
- Omitted/`undefined`/empty arrows create no endpoint children. Explicit endpoint `none` remains direct `type="none"` and is not folded into absence.
- Text line default remains canonical direct no-fill. Arrow-only native creation therefore renders `a:noFill` followed by endpoints; do not copy PptxGenJS's empty-paint line encoding.
- When `line` and `arrows` are both present, render line fill/dash before `headEnd`, and `headEnd` before `tailEnd`.
- Normalize and detach arrows before package mutation. Invalid token/object shape/accessor/symbol/unknown key/alias must leave package bytes, relationships, parts, mutation journal, shape order, and live caches unchanged.
- Plain text, rich text, placeholder creation/population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects must share one implementation.
- Placeholder population preserves owner name, geometry, identity, layout source, and sibling objects; only the populated slide owner receives requested endpoints.
- Existing `ShapeModel.arrows` read/edit/clear/no-op behavior is unchanged. `shape.line = undefined` preserves endpoints, and `shape.arrows = undefined` preserves line width/fill/dash.
- Do not add arrow size, shadow, hyperlink, advanced line fill/custom dash/join/cap/compound/alignment, text geometry, `rectRadius`, `isTextBox`, or `breakLine` behavior.
- Preserve current output byte-for-byte when arrows are omitted.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, browser artifacts, or package build output.
- Use repository-local binaries. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0`.

---

### Task 1: Add strict text-arrow normalization and canonical rendering

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: public `ShapeArrows`, `normalizeShapeArrows(value, context)`, `renderShapeArrows(arrows, prefix)`, existing `validateAddTextOptions()`, `addTextShape()`, `textShapeXml()`, and `ShapeModel.arrows`.
- Produces: `AddTextOptions.arrows?: ShapeArrows`, normalized `arrows?: NormalizedShapeArrows`, and ordered direct endpoints for plain/rich text creation.

- [ ] **Step 1: Write failing public-model tests**

Add focused coverage for omitted, runtime `undefined`, empty, begin-only, end-only, both, explicit none, and all six endpoint tokens:

```ts
const source = { begin: 'triangle', end: 'arrow' } as const;
const plain = slide.addText('Plain arrows', { arrows: source });
const rich = slide.addRichText([{ runs: [{ text: 'Rich arrows' }] }], {
  arrows: { begin: 'none', end: 'stealth' },
});
const combined = slide.addText('Line and arrows', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: { begin: 'diamond', end: 'oval' },
});
```

Require immediate frozen detached snapshots, caller detachment, exact `a:noFill`/`a:solidFill`/`a:prstDash`/`a:headEnd`/`a:tailEnd` order, same-value bytes/journal no-op, whole replacement, clear, and bidirectional line ownership. Require omitted output to equal the pre-change canonical text shape bytes.

- [ ] **Step 2: Write failing invalid-input zero-mutation tests**

Before each invalid call snapshot slide bytes, relationships, package part URI list, mutation journal, shape array, and existing object identity. Reject:

```ts
{ begin: '' }
{ end: 'bogus' }
{ beginArrowType: 'triangle' }
{ lineHead: 'triangle' }
{ begin: 'triangle', extra: true }
[]
new Date()
```

Also reject symbol keys, custom prototypes, inherited-only fields, and accessor properties without invoking their getters. Run representative invalid values through both `addText()` and `addRichText()`.

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts -t "text with strict direct arrows|invalid text arrows" --reporter=dot
```

Expected: compile/runtime failure because `AddTextOptions.arrows` and text endpoint rendering are absent.

- [ ] **Step 4: Implement the minimal model change**

In `packages/model/src/slide.ts`, import the existing renderer and normalized type, then extend the public and normalized values:

```ts
import {
  normalizeShapeArrows,
  readShapeArrows,
  renderShapeArrows,
  replaceShapeArrows,
  type NormalizedShapeArrows,
} from './shape-arrows.internal.js';

export interface AddTextOptions extends Partial<Transform> {
  readonly arrows?: ShapeArrows;
  // retain every existing field
}

interface NormalizedAddTextOptions {
  readonly arrows?: NormalizedShapeArrows;
  // retain every existing field
}

const arrows = normalizeShapeArrows(options.arrows, 'Text shape arrows');
```

Thread `arrows` through `NormalizedTextInput`, `addText()`, both placeholder paths, `addRichText()`, `addTextShape()`, and `textShapeXml()`. Build the existing line contents once:

```ts
const lineContents = renderSimpleLine(line, 'a:') + renderShapeArrows(arrows, 'a:');
const lineXml = line.kind === 'none'
  ? `<a:ln>${lineContents}</a:ln>`
  : `<a:ln w="${points(line.width)}">${lineContents}</a:ln>`;
```

Do not add a second codec or invoke the live setter after creation.

- [ ] **Step 5: Run model regression gates**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/shape-arrows.internal.test.ts packages/model/src/simple-line.internal.test.ts packages/model/src/shape-line.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

Expected: all pass; omitted text output, preset-shape arrows, and text/shape line editing remain unchanged.

- [ ] **Step 6: Review, commit, push, and verify**

Review every changed line against arrows-only scope, validation before mutation, default-byte stability, line→head→tail ordering, and absence of a second mutation. Stage only the two files, commit `feat: create text shape arrows`, push `main:main`, fetch remote `main`, and require `git rev-list --left-right --count HEAD...FETCH_HEAD` to print `0 0`.

---

### Task 2: Prove placeholder, master/layout, root-package, and lifecycle coverage

**Files:**
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddTextOptions.arrows`, `AddPlaceholderOptions`, `SlideMasterObject`, `SlideLayoutModel` / `SlideMasterModel`, and `PptxDocument` lifecycle APIs.
- Produces: strict declarative option acceptance plus end-to-end public coverage for every text owner and presentation format.

- [ ] **Step 1: Add failing SDK surface tests**

Create arrows across slide, layout, master, rich text, placeholders, placeholder population, and declarative named layouts:

```ts
layout.addText('Layout arrows', {
  name: 'layout_arrows',
  arrows: { begin: 'triangle' },
});
master.addRichText([{ runs: [{ text: 'Master arrows' }] }], {
  name: 'master_arrows',
  arrows: { end: 'arrow' },
});
layout.addPlaceholder('Prompt', {
  name: 'title_arrows',
  type: 'title',
  arrows: { begin: 'none', end: 'stealth' },
});
```

Require layout source isolation after slide population, stable wrapper/shape identity, duplicate independence, outer rollback, async declarative detachment, write/reopen, and every `PRESENTATION_FORMAT_PROFILES` format.

- [ ] **Step 2: Add strict declarative and root-package tests**

Add `'arrows'` to `TEXT_OPTION_KEYS`. Pass invalid text/placeholder arrows through `defineSlideMaster()` with and without a delayed image source; require rejection before observable package mutation.

In `packages/pptx/src/index.test.ts`, import `AddTextOptions` and `ShapeArrows`, compile valid begin/end/none values, create/reopen them from `@jiayunxie/pptx`, and add exact negative cases:

```ts
const typedTextArrows: ShapeArrows = { begin: 'triangle', end: 'arrow' };
const typedTextOptions: AddTextOptions = { arrows: typedTextArrows };
// @ts-expect-error PptxGenJS aliases are intentionally unsupported
const invalidTextArrowsAlias: AddTextOptions = { arrows: { beginArrowType: 'arrow' } };
// @ts-expect-error endpoint tokens are a closed union
const invalidTextArrowToken: AddTextOptions = { arrows: { end: 'bogus' } };
```

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text arrows|text shape arrows" --reporter=dot
```

Expected: declarative definition rejects `arrows` until the closed key set is extended; other public tests exercise Task 1.

- [ ] **Step 4: Implement declarative acceptance only**

Add exactly one key:

```ts
const TEXT_OPTION_KEYS = new Set([
  'name',
  'placeholder',
  'align',
  'bullet',
  'fill',
  'line',
  'arrows',
  // retain the remaining existing keys
]);
```

Do not add custom cloning: `readOptions()` / `cloneDataValue()` already detach nested values before asynchronous source preparation, and Task 1 performs semantic validation before package mutation.

- [ ] **Step 5: Run SDK/root/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review placeholder source isolation, asynchronous detachment, layout/master wrapper identity, compile-time surface, and no unrelated option expansion. Commit `test: cover public text shape arrows`, push, fetch, and require divergence `0 0`.

---

### Task 3: Lock PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()` / `write()`, adapter import, native `PptxDocument`, and live `ShapeModel.arrows` / `ShapeModel.line`.
- Produces: exact supported-case and intentional-divergence evidence for text shape endpoints.

- [ ] **Step 1: Add a public-only PptxGenJS fixture**

Create text shapes through public PptxGenJS only for omitted, begin-only/end-only, both, six types on each side, explicit none, arrow-only, `{ type: 'none' }` plus arrow, solid line plus arrows, empty string, nested/top-level deprecated aliases, and invalid token passthrough:

```ts
generatedSlide.addText('Arrow only', {
  objectName: 'Arrow only',
  line: { beginArrowType: 'triangle', endArrowType: 'arrow' },
});
generatedSlide.addText('Solid arrows', {
  objectName: 'Solid arrows',
  line: {
    color: '112233',
    width: 2.5,
    dashType: 'dashDot',
    beginArrowType: 'stealth',
    endArrowType: 'oval',
  },
});
```

Import through the public adapter and inspect exact `a:ln`, `headEnd`, `tailEnd`, child order, paint, width, and dash.

- [ ] **Step 2: Add native comparison and strict-divergence assertions**

Create supported native equivalents with `AddTextOptions.arrows`; compare name, text, transform, geometry, endpoint snapshots, line snapshots, and raw order. Require native arrow-only output to retain canonical direct no-fill, explicit endpoint none to remain distinguishable, and aliases/empty/invalid values to reject with exact zero mutation. Require imported invalid PptxGenJS endpoint state to remain preservation-only.

- [ ] **Step 3: Run focused conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text shape arrows" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape arrows|text arrows|strict direct arrows" --reporter=dot
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

Review that generation uses public PptxGenJS only, assertions distinguish direct-state differences from supported semantic parity, invalid native inputs stay strict, and no adapter production code changed. Commit `test: compare text shape arrows with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 4: Prove actual packed Node, browser, declaration, and CLI behavior

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball, Node/browser conditional exports, generated declarations, and installed `pptx-inspect` CLI.
- Produces: `textShapeArrows: true` in Node and real-Chrome smoke output plus a validated packed deck.

- [ ] **Step 1: Extend packed Node and declaration smoke**

Using only the installed tarball, create plain/rich/placeholder/layout/master/declarative text with begin/end/none and combined line+arrows. Read immediate snapshots, mutate caller input, duplicate, replace/clear through `ShapeModel.arrows`, clear line independently, write/reopen, and save `text-shape-arrows-smoke.pptx`. Add `textShapeArrows` to final JSON. In the generated TypeScript consumer, import `AddTextOptions` and `ShapeArrows`, compile all valid branches, and add `@ts-expect-error` cases matching Task 2.

- [ ] **Step 2: Extend real-Chrome smoke**

In `page.evaluate`, create/reopen a Blob-backed deck with plain, rich, placeholder, and combined line+arrows. Return exact immediate/detached/reopen snapshots plus validation count. Add expected `textShapeArrows` output and require zero console, page, and network errors.

- [ ] **Step 3: Add installed CLI validation**

Run installed `pptx-inspect package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and `part read` on `text-shape-arrows-smoke.pptx`. Require 0 errors/0 warnings, expected slide/shape counts, and raw direct no-fill, solid line, `headEnd`, `tailEnd`, endpoint-none, and child-order evidence.

- [ ] **Step 4: Build and pack without repository artifacts**

Run:

```sh
cd packages/pptx
../../node_modules/.bin/tsup
../../node_modules/.bin/tsup --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
text_arrows_pack_dir="$(mktemp -d /tmp/pptx-text-arrows-pack.XXXXXX)"
npm pack --ignore-scripts --pack-destination "$text_arrows_pack_dir"
cd ../..
node scripts/smoke-npm-package.mjs "$text_arrows_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require exactly 57 tarball files unless a reviewed declaration closure change explains a new count. The installed smoke must reject workspace protocol/internal runtime imports and report `textShapeArrows: true`.

- [ ] **Step 5: Run type/build gates**

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-tarball isolation, Node/browser parity, declaration negatives, CLI profile result, existing smoke fields, and absence of temporary repository artifacts. Commit `test: verify packed text shape arrows`, push, fetch, and require divergence `0 0`.

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
- Produces: accurate public support language and selection of text-shape simple-shadow creation as the next advanced-text subitem.

- [ ] **Step 1: Update public API examples and contracts**

Add plain/rich/placeholder `arrows` examples, exact `ShapeArrows` rules, explicit-none/absence semantics, line ownership, lifecycle behavior, and layout/master/declarative coverage. Do not imply arrow size, text shadow/hyperlink/geometry, or advanced-line support.

- [ ] **Step 2: Update compatibility and progress ledgers**

Move outer text-shape arrows creation from the `slide.addText` gap column into supported behavior. Record PptxGenJS arrow-only/no-line, alias, empty, and invalid-token differences; record final test counts and packed/browser/CLI results; retain shadow/hyperlink/geometry and advanced-line gaps. Set the next small item to text-shape simple-shadow creation.

- [ ] **Step 3: Run stale-language and Markdown checks**

Search all six documents for statements that still list text arrows as unsupported or next. Inspect the changed compatibility table and run:

```sh
git diff --check
rg -n -i "text[- ]shape arrows.*(unsupported|pending|未支持)|next.*text[- ]shape arrows|下一.*text[- ]shape arrows" README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

- [ ] **Step 4: Run final release gates**

```sh
node_modules/.bin/vitest run
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
```

Rerun both tsup builds, declaration closure, final actual tarball smoke, and installed CLI profile. Require `textShapeArrows: true`, Chromium zero console/page/network errors, and PowerPoint 2010 validation 0 errors / 0 warnings.

- [ ] **Step 5: Review, commit, push, and verify**

Review every claim against current command output and raw OOXML evidence. Stage only the six documentation files, commit `docs: document text shape arrow support`, push, fetch, require divergence `0 0`, and report completed item, remaining advanced-text gaps, and updated progress.

## Plan Self-Review

- Spec coverage: Task 1 owns public normalization/rendering; Task 2 covers every owner/lifecycle/root type; Task 3 proves PptxGenJS behavior; Task 4 proves installed package and clients; Task 5 closes docs and release gates.
- Placeholder scan: no unresolved placeholder steps, undefined implementation names, or deferred decisions remain.
- Type consistency: every task uses existing `ShapeArrows`, `NormalizedShapeArrows`, `AddTextOptions`, `ShapeModel.arrows`, `ShapeLine`, `PptxDocument`, and named master/layout APIs; no duplicate public endpoint type is introduced.
- Scope control: arrow size, text shadow/hyperlink/geometry, advanced line styles, `rectRadius`, `isTextBox`, and `breakLine` remain explicitly outside this subitem.
