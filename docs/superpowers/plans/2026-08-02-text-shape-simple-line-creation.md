# Text Shape Simple Line Creation Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict direct none/solid simple-line creation to every public text-shape creation surface, with PptxGenJS 4.0.1 conformance and actual packed Node/browser/type/CLI proof.

**Architecture:** Reuse the existing `normalizeSimpleLine()` and `renderSimpleLine()` value codec. Normalize one detached line inside `validateAddTextOptions()`, thread it through the existing text creation transaction, and replace only the renderer's hard-coded no-fill line with a normalized `a:ln` container; existing `ShapeModel.line` continues to own read/edit/clear semantics.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, Playwright Chromium smoke, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- `AddTextOptions.line` uses the existing public `ShapeLine` union only: direct none or solid sRGB/scheme line with optional finite `0..100` transparency, `0..1584` point width, and one of 8 preset dash values.
- Omitted/`undefined`/none creation remains canonical `<a:ln><a:noFill/></a:ln>`; solid line is rendered after shape fill with normalized width and dash.
- Explicit transparency zero remains direct alpha `100000`, width zero remains direct `w="0"`, and omitted width/dash materialize as 1pt/solid. Do not copy PptxGenJS falsy collapse or absent-owned-state encoding.
- Normalize and detach nested line/color before package mutation. Invalid kind/color/range/object shape/accessor/symbol/unknown key must leave package bytes, relationships, parts, mutation journal, shape order, and live caches unchanged.
- Plain text, rich text, placeholder creation/population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects must share one implementation.
- Placeholder population preserves owner name, geometry, identity, layout source, and sibling objects; only the populated slide owner receives the requested direct line.
- Existing `ShapeModel.line` read/edit/clear/no-op behavior is unchanged; no new line editor or duplicate codec is introduced.
- Do not add arrows, shadow, hyperlink, advanced line fill/custom dash/join/cap/compound/alignment, preset geometry, `rectRadius`, `isTextBox`, or `breakLine` behavior in this subitem.
- Preserve current output byte-for-byte when line is omitted.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, or browser artifacts.
- Use repository-local binaries for validation. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0`.

---

### Task 1: Add strict text-line normalization and canonical rendering

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: public `ShapeLine`, `normalizeSimpleLine(value, context)`, `renderSimpleLine(line, prefix)`, existing `validateAddTextOptions()`, `addTextShape()`, and `textShapeXml()`.
- Produces: `AddTextOptions.line?: ShapeLine`, normalized `line: NormalizedSimpleLine`, and direct none/solid line output for plain/rich text creation.

- [ ] **Step 1: Write failing public-model tests**

Add focused model coverage for omitted, explicit none, sRGB, scheme, explicit-zero/100%-transparent, width 0/default/positive, and every supported dash:

```ts
const source = {
  kind: 'line',
  color: { kind: 'srgb', value: '#ab12cd' },
  transparency: 25,
  width: 2.5,
  dash: 'dashDot',
} as const;
const plain = slide.addText('Plain line', { line: source });
const rich = slide.addRichText([{ runs: [{ text: 'Rich line' }] }], {
  line: { kind: 'line', color: { kind: 'scheme', value: 'accent2' } },
});
const none = slide.addText('None', { line: { kind: 'none' } });
```

Require immediate normalized snapshots, caller detachment, exact `w`, `solidFill`, alpha and `prstDash`, and geometry→fill→line ordering. Require omitted output to equal the pre-change canonical no-fill line skeleton. Exercise same-value set and clear through the returned live `ShapeModel`.

- [ ] **Step 2: Write failing invalid-input zero-mutation tests**

Before each invalid call snapshot slide bytes, relationships, package part URI list, mutation journal, shape array, and existing object identity. Reject:

```ts
{ kind: 'line' }
{ kind: 'line', color: { kind: 'srgb', value: 'GG0000' } }
{ kind: 'line', color: { kind: 'scheme', value: 'unknown' } }
{ kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, transparency: -1 }
{ kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1585 }
{ kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, dash: 'dot' }
{ kind: 'none', width: 1 }
{ type: 'none' }
```

Also reject arrays, dates, symbol keys, unknown keys, and accessors without invoking the nested getter. Run the same invalid line through `addText()` and `addRichText()`.

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts -t "text with strict direct lines|invalid text line" --reporter=dot
```

Expected: TypeScript/runtime failure because `AddTextOptions.line` and text line rendering are absent.

- [ ] **Step 4: Implement the minimal model change**

In `slide.ts` add `line?: ShapeLine`, import `renderSimpleLine` and `NormalizedSimpleLine`, and normalize once:

```ts
const line = normalizeSimpleLine(options.line, 'Text shape line') ?? { kind: 'none' };
```

Add normalized line to `NormalizedAddTextOptions` and `NormalizedTextInput`; pass it from `addText()`, both `addPlaceholder()` branches, and `addRichText()` into `addTextShape()` and `textShapeXml()`. Render:

```ts
line.kind === 'none'
  ? `<a:ln>${renderSimpleLine(line, 'a:')}</a:ln>`
  : `<a:ln w="${points(line.width)}">${renderSimpleLine(line, 'a:')}</a:ln>`
```

Do not alter fill, transform, paragraph, placeholder, or transaction logic.

- [ ] **Step 5: Run model regression gates**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/simple-line.internal.test.ts packages/model/src/shape-line.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

Expected: all pass; omitted text output and existing shape-line editing tests remain unchanged.

- [ ] **Step 6: Review, commit, push, and verify**

Review every changed line against line-only scope, validation before mutation, default-byte stability, renderer ordering, and absence of a second codec. Stage only the two files, commit `feat: create text shape lines`, push, fetch `origin/main`, and require `git rev-list --left-right --count HEAD...origin/main` to print `0 0`.

---

### Task 2: Prove placeholder, master/layout, root-package, and lifecycle coverage

**Files:**
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddTextOptions.line`, `AddPlaceholderOptions`, `SlideMasterObject`, `SlideLayoutModel`/`SlideMasterModel`, and `PptxDocument` lifecycle APIs.
- Produces: strict declarative option acceptance plus end-to-end public coverage for all creation owners and formats.

- [ ] **Step 1: Add failing SDK surface tests**

Create lines across slide, layout, master, rich text, placeholders, and declarative named layouts. Require direct snapshots while the layout placeholder source remains unchanged. Prove stable wrapper/shape identity, duplicate independence, outer rollback, write/reopen, and every `PRESENTATION_FORMAT_PROFILES` format.

- [ ] **Step 2: Add strict declarative and root-package tests**

Add `line` to `TEXT_OPTION_KEYS`. Pass invalid text/placeholder lines through `defineSlideMaster()` with and without a delayed image source; require rejection before observable package mutation. In `packages/pptx/src/index.test.ts`, import `AddTextOptions` and `ShapeLine`, compile valid none/sRGB/scheme values, create/reopen them from `@jiayunxie/pptx`, and add `@ts-expect-error` cases for PptxGenJS-shaped `{ color: 'FF0000' }`, invalid kind, missing solid color, bad width/transparency/dash types, and unknown keys.

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text lines|text shape line" --reporter=dot
```

Expected: declarative definition rejects `line` until `TEXT_OPTION_KEYS` is extended; new public tests otherwise exercise Task 1.

- [ ] **Step 4: Implement declarative acceptance only**

Add `'line'` to `TEXT_OPTION_KEYS`. Do not add custom cloning: existing `readOptions()`/`cloneDataValue()` already detaches nested line before asynchronous source preparation, and Task 1's model normalizer performs semantic validation before package mutation.

- [ ] **Step 5: Run SDK/root/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review placeholder source isolation, asynchronous detachment, layout/master wrapper identity, compile-time surface, and no unrelated option expansion. Commit `test: cover public text shape lines`, push, fetch, and require divergence `0 0`.

---

### Task 3: Lock PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()`/`write()`, adapter import, native `PptxDocument`, and live `ShapeModel.line`.
- Produces: exact supported-case and intentional-divergence evidence for text shape simple lines.

- [ ] **Step 1: Add a public-only PptxGenJS fixture**

Create text shapes through public PptxGenJS only for omitted, none, empty/missing-color, sRGB, scheme, transparency 0/25/100, width 0/positive, each of the 8 dash values, and deprecated nested aliases. Import via `importPptxGenJS()` and require supported strict snapshots where direct state is fully decodable; inspect raw slide XML for empty `a:ln`, fill type/color/alpha, width, dash, and fill→line ordering.

- [ ] **Step 2: Add native comparison and strict-divergence assertions**

Create supported native equivalents using `AddTextOptions.line`; compare names, transform, geometry, direct color/alpha/width/dash semantics, fill, and text. Require native omitted/none to remain direct no-fill, default width/dash to materialize, explicit zero alpha to remain `100000`, and width zero to remain `w="0"`. Reject PptxGenJS object shapes and aliases with exact zero mutation.

- [ ] **Step 3: Run focused conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text shape line" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape line|text lines|strict direct lines" --reporter=dot
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

Review that generation uses public PptxGenJS only, assertions distinguish direct-state differences from effective conformance, invalid native inputs remain strict, and no adapter production code changed. Commit `test: compare text shape lines with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 4: Prove actual packed Node, browser, declaration, and CLI behavior

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball, Node/browser conditional exports, generated declarations, and installed `pptx-inspect` CLI.
- Produces: `textShapeLines: true` in Node and real-Chrome smoke output plus a validated packed deck.

- [ ] **Step 1: Extend packed Node and declaration smoke**

Using only the installed tarball, create plain/rich/placeholder/layout/master/declarative text with none/sRGB/scheme/zero/100 lines, read immediate snapshots, duplicate, edit/clear through `ShapeModel.line`, write/reopen, and save `text-shape-line-smoke.pptx`. Add `textShapeLines` to final JSON. In the generated TypeScript consumer, import `AddTextOptions` and `ShapeLine`, compile all valid branches, and add `@ts-expect-error` cases matching Task 2.

- [ ] **Step 2: Extend real-Chrome smoke**

In `page.evaluate`, create/reopen a Blob-backed deck with plain, rich, and placeholder lines. Return exact snapshot/color/transparency/width/dash arrays plus validation error count. Add expected `textShapeLines` output and require zero console, page, and network errors.

- [ ] **Step 3: Add installed CLI validation**

Run installed `pptx-inspect package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and `part read` on `text-shape-line-smoke.pptx`. Require 0 errors/0 warnings, expected slide/shape counts, and raw `srgbClr`, `schemeClr`, alpha, width, dash, and direct no-fill evidence.

- [ ] **Step 4: Build and pack without persistent artifacts**

From `packages/pptx`, run the two repository-local tsup configs and `node ../../scripts/build-npm-package-types.mjs`; create a fresh temporary directory, run `npm pack --ignore-scripts --pack-destination <temp>`, then execute `node scripts/smoke-npm-package.mjs <tarball>`. Confirm tarball files contain no workspace protocol/internal runtime import and browser smoke reports `textShapeLines: true`.

- [ ] **Step 5: Run type/build gates**

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-tarball isolation, Node/browser parity, declaration negatives, CLI profile result, existing smoke fields, and absence of temporary artifacts. Commit `test: verify packed text shape lines`, push, fetch, and require divergence `0 0`.

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

- [ ] **Step 1: Update public API examples and contracts**

Add plain/rich/placeholder `line` examples, exact `ShapeLine` rules, default/none/solid/zero semantics, lifecycle behavior, and layout/master/declarative coverage. Do not imply arrows/shadow/hyperlink/geometry/advanced-line support.

- [ ] **Step 2: Update compatibility and progress ledgers**

Move outer text-shape simple-line creation from the `slide.addText` gap column into supported behavior. Record PptxGenJS omitted/none/default/zero differences, test counts, packed/browser/CLI results, and retain remaining arrows/shadow/hyperlink/shape/rectRadius/isTextBox/breakLine gaps. Select the next smallest advanced-text subitem based on the remaining renderer boundary.

- [ ] **Step 3: Run stale-language and Markdown checks**

Search all six documents for statements that still list text outer simple line as unsupported. Check Markdown table column counts with a local Node script and run `git diff --check`.

- [ ] **Step 4: Run final release gates**

```sh
node_modules/.bin/vitest run
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
```

Rerun the final actual tarball smoke and require `textShapeLines: true`, Chromium zero errors, and PowerPoint 2010 validation 0/0.

- [ ] **Step 5: Review, commit, push, and verify**

Review claims against current command output and raw OOXML evidence. Stage only the six documentation files, commit `docs: document text shape line support`, push, fetch, require divergence `0 0`, and report completed item, remaining advanced-text gaps, and updated progress.

## Plan Self-Review

- Spec coverage: Task 1 owns public model normalization/rendering; Task 2 covers every owner/lifecycle/root type; Task 3 proves PptxGenJS behavior; Task 4 proves installed package and clients; Task 5 closes docs and release gates.
- Placeholder scan: no unresolved placeholder steps or undefined implementation names remain.
- Type consistency: every task uses existing `ShapeLine`, `NormalizedSimpleLine`, `AddTextOptions`, `ShapeModel.line`, `PptxDocument`, and named master/layout APIs; no duplicate public line type is introduced.
- Scope control: text arrows/shadow/hyperlink/geometry/isTextBox/breakLine remain explicitly outside this subitem.
