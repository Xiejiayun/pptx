# Text Shape Fill Creation Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict direct none/solid fill creation to every public text-shape creation surface, with PptxGenJS 4.0.1 conformance and actual packed Node/browser/type/CLI proof.

**Architecture:** Reuse the existing `normalizeSimpleFill()` and `renderSimpleFill()` value codec. Normalize one detached fill inside `validateAddTextOptions()`, thread it through the existing text creation transaction, and replace only the renderer's hard-coded `<a:noFill/>`; existing `ShapeModel.fill` continues to own read/edit/clear semantics.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, Playwright Chromium smoke, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- `AddTextOptions.fill` uses the existing public `ShapeFill` union only: direct none or solid sRGB/scheme color with optional finite `0..100` transparency quantized to `0.001%`.
- Omitted/`undefined`/none creation remains canonical direct `a:noFill`; solid fill is rendered after geometry and before line.
- Explicit transparency zero remains direct alpha `100000`; do not copy PptxGenJS falsy collapse. Native none remains direct no-fill; do not copy PptxGenJS's absent-fill encoding.
- Normalize and detach nested fill before package mutation. Invalid kind/color/range/object shape/accessor/symbol/unknown key must leave package bytes, relationships, parts, mutation journal, shape order, and live caches unchanged.
- Plain text, rich text, placeholder creation/population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects must share one implementation.
- Placeholder population preserves owner name, geometry, identity, layout source, and sibling objects; only the populated slide owner receives the requested direct fill.
- Existing `ShapeModel.fill` read/edit/clear/no-op behavior is unchanged; no new fill editor or duplicate codec is introduced.
- Do not add line, arrows, shadow, hyperlink, preset geometry, `rectRadius`, `isTextBox`, run hyperlink, advanced fill, or `breakLine` behavior in this subitem.
- Preserve current output byte-for-byte when fill is omitted.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, or browser artifacts.
- Use repository-local binaries for validation. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0`.

---

### Task 1: Add strict text-fill normalization and canonical rendering

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: public `ShapeFill`, `normalizeSimpleFill(value, context)`, `renderSimpleFill(fill, prefix)`, existing `validateAddTextOptions()`, `addTextShape()`, and `textShapeXml()`.
- Produces: `AddTextOptions.fill?: ShapeFill`, normalized `fill: ShapeFill`, and direct none/solid fill output for plain/rich text creation.

- [ ] **Step 1: Write failing public-model tests**

Add `creates plain and rich text with strict direct fills` to `model.test.ts`. Create omitted, explicit none, sRGB, scheme, explicit-zero, and 100%-transparent shapes:

```ts
const source = { kind: 'solid', color: { kind: 'srgb', value: '#ab12cd' }, transparency: 25 } as const;
const plain = slide.addText('Plain fill', { fill: source });
const rich = slide.addRichText([{ runs: [{ text: 'Rich fill' }] }], {
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' }, transparency: 0 },
});
const none = slide.addText('None', { fill: { kind: 'none' } });
```

Require immediate snapshots `AB12CD/25`, `accent2/0`, and none; mutate the caller's nested color/transparency after creation and require unchanged XML/snapshot. Assert exact direct `solidFill`, alpha `75000`/`100000`, canonical no-fill, and geometry→fill→line ordering. Require omitted output to equal the pre-change canonical no-fill skeleton.

- [ ] **Step 2: Write failing invalid-input zero-mutation tests**

Before each invalid call snapshot slide bytes, relationships, package part URI list, mutation journal, shape array, and existing object identity. Reject:

```ts
{ kind: 'solid' }
{ kind: 'solid', color: { kind: 'srgb', value: 'GG0000' } }
{ kind: 'solid', color: { kind: 'scheme', value: 'unknown' } }
{ kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: -1 }
{ kind: 'none', transparency: 1 }
{ type: 'none' }
```

Also reject arrays, dates, symbol keys, unknown keys, and accessors without invoking the nested getter. Run the same invalid fill through `addText()` and `addRichText()`.

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts -t "text with strict direct fills|invalid text fill" --reporter=dot
```

Expected: TypeScript/runtime failure because `AddTextOptions.fill` and text fill rendering are absent.

- [ ] **Step 4: Implement the minimal model change**

In `slide.ts`:

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly fill?: ShapeFill;
  // retain all existing fields
}

interface NormalizedAddTextOptions {
  readonly fill: ShapeFill;
  // retain all existing fields
}
```

Import `renderSimpleFill`. In `validateAddTextOptions()` compute:

```ts
const fill = normalizeSimpleFill(options.fill, 'Text shape fill') ?? { kind: 'none' };
```

Return `fill`, pass it from `addText()`, `addRichText()`, and both `addPlaceholder()` branches into `addTextShape()`, then into `textShapeXml()`. Replace only the hard-coded `<a:noFill/>` with `renderSimpleFill(fill, 'a:')`.

- [ ] **Step 5: Run model regression gates**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/simple-fill.internal.test.ts packages/model/src/shape-fill.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

Expected: all pass; omitted text output and existing shape-fill editing tests remain unchanged.

- [ ] **Step 6: Review, commit, push, and verify**

Review every changed line against the fill-only scope, validation before mutation, default-byte stability, renderer ordering, and absence of a second codec. Then stage only the two files, commit `feat: create text shape fills`, push through the working GitHub transport, fetch `origin/main`, and require `git rev-list --left-right --count HEAD...origin/main` to print `0 0`.

---

### Task 2: Prove placeholder, master/layout, root-package, and lifecycle coverage

**Files:**
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddTextOptions.fill`, `AddPlaceholderOptions`, `SlideMasterObject`, `SlideLayoutModel`/`SlideMasterModel`, and `PptxDocument` lifecycle APIs.
- Produces: strict declarative option acceptance plus end-to-end public coverage for all creation owners and formats.

- [ ] **Step 1: Add failing SDK surface tests**

Add `creates text fills across slide layout master and placeholder owners` to `packages/sdk/src/index.test.ts`. Require:

```ts
layout.addText('Layout', { fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } } });
master.addRichText([{ runs: [{ text: 'Master' }] }], {
  fill: { kind: 'solid', color: { kind: 'srgb', value: '112233' }, transparency: 40 },
});
const populated = slide.addText('Populated', {
  placeholder: 'title_fill',
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent3' }, transparency: 25 },
});
```

Define a named layout with text and placeholder objects containing fills, select it, populate the placeholder, and require direct snapshots on layout/master/slide while the layout placeholder remains unchanged. Prove stable wrapper/shape identity, duplicate independence, outer rollback, write/reopen, and every `PRESENTATION_FORMAT_PROFILES` format.

- [ ] **Step 2: Add strict declarative and root-package tests**

Add `fill` to `TEXT_OPTION_KEYS`. In SDK tests, pass invalid text/placeholder fills through `defineSlideMaster()` with and without a delayed image source; require rejection before observable package mutation. In `packages/pptx/src/index.test.ts`, import `AddTextOptions` and `ShapeFill`, compile valid none/sRGB/scheme values, create/reopen them from `@jiayunxie/pptx`, and add `@ts-expect-error` cases for PptxGenJS-shaped `{ color: 'FF0000' }`, invalid kind, missing solid color, bad transparency type, and unknown key.

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text fills|text shape fill" --reporter=dot
```

Expected: declarative definition rejects `fill` until `TEXT_OPTION_KEYS` is extended; new public tests otherwise exercise Task 1.

- [ ] **Step 4: Implement declarative acceptance only**

Add `'fill'` to `TEXT_OPTION_KEYS`. Do not add custom cloning: existing `readOptions()`/`cloneDataValue()` already detaches nested fill before asynchronous source preparation, and Task 1's model normalizer performs semantic validation before the package transaction commits objects.

- [ ] **Step 5: Run SDK/root/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

Expected: all pass, including six-format and named-layout tests.

- [ ] **Step 6: Review, commit, push, and verify**

Review placeholder source isolation, asynchronous detachment, layout/master wrapper identity, compile-time surface, and no unrelated option expansion. Commit `test: cover public text shape fills`, push, fetch, and require divergence `0 0`.

---

### Task 3: Lock PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()`/`write()`, adapter import, native `PptxDocument`, and live `ShapeModel.fill`.
- Produces: exact supported-case and intentional-divergence evidence for text shape fill.

- [ ] **Step 1: Add a public-only PptxGenJS fixture**

Create six text shapes through public PptxGenJS only: omitted, `{ type: 'none' }`, sRGB 25%, scheme omitted alpha, explicit-zero transparency, and invalid/missing-color fallback. Import via `importPptxGenJS()` and require snapshots:

```ts
[
  { kind: 'none' },
  undefined,
  { kind: 'solid', color: { kind: 'srgb', value: 'AB12CD' }, transparency: 25 },
  { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
  { kind: 'solid', color: { kind: 'srgb', value: '00AA00' } },
]
```

Inspect raw slide XML to prove PptxGenJS omitted direct no-fill, none absent choice, alpha `75000`, scheme color, and zero-alpha omission. Reopen imported output and require unchanged snapshots/bytes.

- [ ] **Step 2: Add native comparison and strict-divergence assertions**

Create the supported native equivalents using `AddTextOptions.fill`; compare names, transform, geometry, direct fill type/color/alpha semantics, line, and text. Require native explicit none to remain direct no-fill and explicit zero to remain alpha `100000`. Reject PptxGenJS object shape `{ color: 'AB12CD' }`, `type`, `alpha`, missing color, coercible transparency, and out-of-range values with exact zero mutation.

- [ ] **Step 3: Run focused test and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text shape fill" --reporter=dot
```

Expected before Task 1: native creation cannot compile/render fill. After Tasks 1–2: pass without production changes.

- [ ] **Step 4: Run adapter and full focused regression**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape fill|text fills|strict direct fills" --reporter=dot
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review that generation uses public PptxGenJS only, assertions distinguish direct-state differences from effective conformance, invalid native inputs remain strict, and no adapter production code changed. Commit `test: compare text shape fills with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 4: Prove actual packed Node, browser, declaration, and CLI behavior

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball, Node/browser conditional exports, generated declarations, and installed `pptx-inspect` CLI.
- Produces: `textShapeFills: true` in Node and real-Chrome smoke output plus a validated packed deck.

- [ ] **Step 1: Extend packed Node and declaration smoke**

Using only the installed tarball, create plain/rich/placeholder/layout/master/declarative text with none/sRGB/scheme/zero/100 fills, read immediate snapshots, duplicate, edit/clear through `ShapeModel.fill`, write/reopen, and save `text-shape-fill-smoke.pptx`. Add `textShapeFills` to the final JSON. In the generated TypeScript consumer, import `AddTextOptions` and `ShapeFill`, compile all valid branches, and add `@ts-expect-error` cases matching Task 2.

- [ ] **Step 2: Extend real-Chrome smoke**

In `page.evaluate`, create/reopen a Blob-backed deck with plain, rich, and placeholder fills. Return exact snapshot/color/transparency arrays plus validation error count. Add expected `textShapeFills` output and require zero console, page, and network errors.

- [ ] **Step 3: Add installed CLI validation**

Run installed `pptx-inspect package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and `part read` on `text-shape-fill-smoke.pptx`. Require 0 errors/0 warnings, expected slide/shape counts, and raw `srgbClr`, `schemeClr`, `alpha 75000`, `alpha 100000`, and direct no-fill evidence.

- [ ] **Step 4: Build and pack without persistent artifacts**

From `packages/pptx`, run the two repository-local tsup configs and `node ../../scripts/build-npm-package-types.mjs`; create a fresh temporary directory, run `npm pack --ignore-scripts --pack-destination <temp>`, then execute `node scripts/smoke-npm-package.mjs <tarball>`. Confirm tarball files contain no workspace protocol/internal runtime import and browser smoke reports `textShapeFills: true`.

- [ ] **Step 5: Run type/build gates**

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-tarball isolation, Node/browser parity, declaration negatives, CLI profile result, existing smoke fields, and absence of temporary artifacts. Commit `test: verify packed text shape fills`, push, fetch, and require divergence `0 0`.

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
- Consumes: all implementation, conformance, packed, validator, and browser evidence from Tasks 1–4.
- Produces: accurate public support language and the next advanced-text subitem selection.

- [ ] **Step 1: Update the public API examples and contracts**

Add plain/rich/placeholder `fill` examples, exact `ShapeFill` value rules, default/none/solid/explicit-zero semantics, lifecycle behavior, and layout/master/declarative coverage. Do not imply line/shadow/hyperlink/geometry support.

- [ ] **Step 2: Update compatibility and progress ledgers**

Move outer text-shape fill creation from the `slide.addText` gap column into supported behavior. Record PptxGenJS omitted/none/zero differences, test counts, packed/browser/CLI results, and retain remaining line/arrows/shadow/hyperlink/shape/rectRadius/isTextBox/breakLine gaps. Select text shape simple line creation as the next subitem because it reuses the same renderer boundary and existing line codec.

- [ ] **Step 3: Run stale-language and Markdown checks**

Search all six documents for statements that still list text outer fill as unsupported. Check Markdown table column counts with a local Node script and run `git diff --check`.

- [ ] **Step 4: Run final release gates**

```sh
node_modules/.bin/vitest run
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
```

Rerun the final actual tarball smoke after the documentation edits and require `textShapeFills: true`, Chromium zero errors, and PowerPoint 2010 validation 0/0.

- [ ] **Step 5: Review, commit, push, and verify**

Review claims against current command output and raw OOXML evidence. Stage only the six documentation files, commit `docs: document text shape fill support`, push, fetch, require divergence `0 0`, and report completed item, remaining advanced-text gaps, and updated progress.

## Plan Self-Review

- Spec coverage: Task 1 owns public model normalization/rendering; Task 2 covers every owner/lifecycle/root type; Task 3 proves PptxGenJS behavior; Task 4 proves installed package and clients; Task 5 closes docs and release gates.
- Placeholder scan: no unresolved placeholder steps or undefined implementation names remain.
- Type consistency: every task uses existing `ShapeFill`, `AddTextOptions`, `ShapeModel.fill`, `PptxDocument`, and named master/layout APIs; no duplicate public fill type is introduced.
- Scope control: text line/arrows/shadow/hyperlink/geometry/isTextBox/breakLine remain explicitly outside this subitem.
