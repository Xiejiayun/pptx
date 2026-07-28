# Text Box Wrapping Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native create/read/edit/clear support for text-box wrapping with PptxGenJS 4.0.1 output conformance.

**Architecture:** Add a focused wrapping codec that maps public booleans to direct `bodyPr@wrap` tokens and reuses the existing body-properties locator/attribute updater. Normalize creation input before mutation, keep the existing explicit wrapped default, and delegate live reads/edits through `ShapeModel.textWrap` without interpreting inheritance or rebuilding text content.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation API is `AddTextOptions.wrap?: boolean`; live editing API is `ShapeModel.textWrap`.
- Public true/false serialize as direct `bodyPr@wrap` tokens `square` / `none`.
- Omitted creation remains explicit true (`wrap="square"`); live setter `undefined` removes the direct wrap override.
- Read only exact direct `square` / `none`; absent, empty, case variants, whitespace, and unknown tokens return `undefined`.
- Preserve unsupported wrap tokens and every unowned body property during getters and unrelated edits.
- Fit/overflow, table cells, placeholders/master inheritance, vertical text, RTL, columns, and word breaking are outside this item.
- PptxGenJS is used only by the adapter conformance test and never becomes a non-adapter dependency.
- Do not stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.
- The user delegated design and execution choices; execute inline without pausing for questions or dispatching subagents.

---

### Task 1: Strict wrapping codec and live model accessors

**Files:**
- Create: `packages/model/src/text-box-wrapping.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `requireTextBodyProperties()`, `updateTextBodyAttribute()`, `LosslessXmlDocument`, `XmlElement`, `SlideModel` delegation, and stable `ShapeModel` identity.
- Produces: `normalizeTextBoxWrap()`, `renderTextBoxWrapAttribute()`, `readTextBoxWrap()`, `replaceTextBoxWrap()`, `SlideModel.getShapeTextWrap()` / `setShapeTextWrap()`, and `ShapeModel.textWrap`.

- [ ] **Step 1: Write a failing strict-read and lossless-edit model test**

Inject expanded direct body properties containing single-quoted `wrap='none'`, margins, vertical anchor, text direction metadata, an autofit child, and an unknown child:

```ts
const bodyProperties = [
  `<a:bodyPr wrap='none' lIns="127000" anchor="b" vert="vert" custom="KEEP">`,
  '<a:normAutofit fontScale="90000"/>',
  '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
  '</a:bodyPr>',
].join('');
```

Assert `shape.textWrap === false`, repeated getter access adds no mutation, and snapshots remain live. Set true and verify only the existing attribute value becomes `wrap='square'`; then set false. Replace plain/rich text, margins, vertical alignment, and transform, asserting wrap and all unowned metadata survive. Roll back a wrapping change and finally set `undefined`, asserting only direct wrap disappears while content, identity, anchor, vert, autofit, and unknown XML remain.

- [ ] **Step 2: Run the model test and confirm the red state**

Run:

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new tests fail because `ShapeModel.textWrap` and the wrapping codec do not exist.

- [ ] **Step 3: Add the focused wrapping codec**

Create `text-box-wrapping.internal.ts` with exact mappings:

```ts
const TO_OOXML = new Map<boolean, string>([
  [true, 'square'],
  [false, 'none'],
]);

const FROM_OOXML = new Map<string, boolean>([
  ['square', true],
  ['none', false],
]);
```

Implement:

```ts
export function normalizeTextBoxWrap(value: unknown, context: string): boolean;
export function renderTextBoxWrapAttribute(value: boolean): string;
export function readTextBoxWrap(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): boolean | undefined;
export function replaceTextBoxWrap(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: boolean | undefined,
  partUri: string,
): void;
```

Normalization accepts only boolean. Read only an exact direct token. Replace only `wrap` through the shared body-properties helper.

- [ ] **Step 4: Delegate live reads and edits through SlideModel/ShapeModel**

Add to `SlideModel`:

```ts
getShapeTextWrap(id: number): boolean | undefined;
setShapeTextWrap(id: number, value: boolean | undefined): void;
```

The setter validates non-undefined input before applying a patch, runs in `opcPackage.transaction()`, resolves the current live shape, and saves with `setXml()`.

Add asymmetric accessors to `ShapeModel`:

```ts
get textWrap(): boolean | undefined;
set textWrap(value: boolean | undefined);
```

- [ ] **Step 5: Complete strict-token and malformed-shape coverage**

Use fresh fixtures for `square`, `none`, absent, empty, `Square`, ` false `, `tight`, and unknown tokens. Assert only `square` / `none` map to public values and getters never mutate. For an unsupported token, run `.text`, `.richText`, `textMargins`, `verticalAlignment`, and transform edits and verify the token remains byte-for-byte. Extend malformed direct `txBody/bodyPr` tests to both getter and setter, checking bytes and journal remain unchanged.

- [ ] **Step 6: Run focused model regressions**

Run the Task 1 model command. Expected: all model tests pass, including text-box margin and vertical-alignment suites that share the same body-properties helper.

### Task 2: Plain/rich creation lifecycle and invalid-input isolation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 codec, `AddTextOptions`, shared plain/rich `addTextShape()` path, duplicate/write/reopen transactions.
- Produces: `AddTextOptions.wrap`, explicit true default, false creation, and native lifecycle evidence.

- [ ] **Step 1: Write failing creation and lifecycle tests**

Create plain/rich shapes for omitted, true, and false while combining wrap with margins, vertical alignment, paragraph alignment, and run style:

```ts
const omitted = slide.addText('Omitted');
const wrapped = slide.addRichText(
  [{ align: 'center', runs: [{ text: 'Wrapped', style: { bold: true } }] }],
  { wrap: true, margin: 8, valign: 'bottom' },
);
const unwrapped = slide.addText('Unwrapped', { wrap: false, valign: 'top' });
```

Assert snapshots are true/true/false and XML contains `square`, `square`, and `none` on the expected body properties. Duplicate before editing; change wrapped to false, clear unwrapped, add true back to a self-closing bodyPr that lacks wrap, and verify edited versus duplicated snapshots after write/reopen. Plain/rich, margin, vertical-alignment, and transform edits must preserve wrapping.

- [ ] **Step 2: Normalize creation options and serialize wrapping**

Extend `AddTextOptions`:

```ts
readonly wrap?: boolean;
```

Extend `NormalizedAddTextOptions` and `NormalizedTextInput` with required `textWrap: boolean`. In `validateAddTextOptions()`, use true when omitted, otherwise call `normalizeTextBoxWrap(options.wrap, 'Text wrap')`. Pass the normalized boolean through `addTextShape()` to `textShapeXml()` and replace the hard-coded `wrap="square"` with `renderTextBoxWrapAttribute(textWrap)`.

- [ ] **Step 3: Add invalid-input mutation-isolation cases**

Test null, numbers, strings including `square` / `none`, arrays, objects, and symbols. For `addText()`, `addRichText()`, and the live setter, assert `TypeError` and exact preservation of slide bytes, mutation journal, text, margins, vertical alignment, wrapping snapshot, and stable identity.

- [ ] **Step 4: Run typecheck and focused model/SDK tests**

Run:

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck exits 0 and both suites pass.

### Task 3: PptxGenJS conformance and public release surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/api/README.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 1/2 public API and existing PptxGenJS public-write fixture.
- Produces: real 4.0.1 wrapping evidence, compatibility matrix, public docs, and Node/browser/declaration smoke.

- [ ] **Step 1: Add PptxGenJS-generated wrapping shapes**

Generate named shapes for omitted, true, false, invalid string fallback, and a `TextProps[]` run with `options.wrap: false` but no text-box wrap. Import public bytes and assert snapshots are true/true/false/true/true. Assert raw direct tokens, then duplicate/write/reopen and verify the same values.

- [ ] **Step 2: Advance compatibility and API documentation**

Add a completed matrix row:

```md
| 文本框 `wrap: boolean` 与 direct wrapping 编辑 | `AddTextOptions.wrap` / `ShapeModel.textWrap` | 已支持 |
```

Remove wrap from the remaining text-body partial row. Document creation default true versus live clear semantics in `docs/api/README.md`, mention wrapping in the npm README, and add one Unreleased changelog bullet.

- [ ] **Step 3: Extend packed-package smoke**

In Node smoke, create false, set true, then clear. In browser smoke, create false and assert the snapshot. In generated TypeScript smoke, compile:

```ts
const wrapped = createdDocument.addSlide().addText('Wrapped', { wrap: true });
const wrapSnapshot: boolean | undefined = wrapped.textWrap;
wrapped.textWrap = false;
wrapped.textWrap = undefined;
```

Ensure packed declarations accept creation, getter, setter, and clear assignment.

- [ ] **Step 4: Run adapter, typecheck, full tests, and performance**

Run the adapter test, full typecheck, full Vitest suite, and isolated `RUN_PERF=1` performance suite. Expected: all functional tests pass, the default performance test remains skipped in the full run, and the isolated performance test passes.

### Task 4: Package, CLI, LibreOffice, review, commit, and push

**Files:**
- Review all files listed in Tasks 1–3; do not stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation/tests/docs.
- Produces: reviewed `feat: support text box wrapping` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the real npm tarball**

From `packages/pptx`, run pinned tsup 8.5.1 Node/browser builds and `scripts/build-npm-package-types.mjs`. Create the destination with:

```sh
package_dir=$(mktemp -d /tmp/pptx-text-wrap-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$package_dir"
node ../../scripts/smoke-npm-package.mjs "$package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: API, browser, types, and CLI checks all succeed.

- [ ] **Step 2: Validate and render native/PptxGenJS comparison decks**

Generate wide-layout native and PptxGenJS decks in a fresh `/tmp/pptx-text-wrap-validation.*` directory. Each deck contains equal-size narrow filled boxes with the same long text, explicit margins, and wrap true versus false. Validate both through the repository CLI using `powerpoint-2010`; expect 0 errors and 0 warnings.

Record source SHA-256 hashes, export both through `soffice --headless --convert-to pdf`, record hashes again, and require no repair output or source mutation. Render PDFs with `pdftoppm -png -r 150` and visually verify wrapped text stays within the narrow box while unwrapped text remains on one line/overflows in the same direction, with native and PptxGenJS behavior matching.

- [ ] **Step 3: Final review**

Run `git diff --check`; inspect the complete diff; verify the new codec owns only `wrap`; confirm margin and vertical-alignment serialization/tests remain unchanged; confirm every changed line traces to wrapping; verify status lists only intended files plus untracked `.pnpm-store/`. Re-run focused tests after any review fix.

- [ ] **Step 4: Commit and push**

Stage only intended files, commit:

```sh
git commit -m "feat: support text box wrapping"
```

Push `main` through SSH-over-443. Verify `origin/main...HEAD` is `0 0` and only `.pnpm-store/` remains untracked, then immediately select the next unsupported PptxGenJS surface and repeat the design/plan/implementation cycle.
