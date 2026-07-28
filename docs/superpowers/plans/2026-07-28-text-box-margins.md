# Text Box Margins Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native create/read/edit/clear support for point-based text-box margins with PptxGenJS 4.0.1 output conformance.

**Architecture:** Add a focused internal text-box-margin codec that normalizes public scalar/tuple/object input, renders the four `a:bodyPr` inset attributes, reads only strict direct Int32 values, and patches only those attributes. Route both plain/rich text creation and live `ShapeModel` editing through it while leaving paragraph `marL/indent`, other body properties, and text content untouched.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation API is `AddTextOptions.margin?: TextBoxMarginInput`; live editing API is `ShapeModel.textMargins`.
- Margins use points. Scalar means all four sides; tuple order is `[top, right, bottom, left]`; object keys are `top/right/bottom/left`.
- Write `raw = Math.round(points * 12700)` to `tIns/rIns/bIns/lIns`; explicit zero must survive.
- Accept negative finite values, but reject any quantized raw value outside signed Int32.
- Read only strict integer direct `bodyPr` attributes within signed Int32 and return `raw / 12700`; malformed sides do not hide valid siblings.
- Setter replaces only the four supported direct attributes; `undefined` or `{}` clears them and preserves all other `bodyPr` attributes/children byte-for-byte.
- Paragraph margins, first-line indent, bullet hanging indent, fit, wrap, vertical alignment, direction, RTL, and columns are outside this item.
- PptxGenJS is used only by the adapter conformance test and never becomes a non-adapter dependency.
- Do not stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.

---

### Task 1: Public types, internal normalization, strict read, and lossless patch

**Files:**
- Modify: `packages/model/src/text.ts`
- Create: `packages/model/src/text-box-margins.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, direct `txBody/bodyPr`, and the existing `SlideModel`/`ShapeModel` delegation pattern.
- Produces: `TextBoxMargins`, `TextBoxMarginInput`, `normalizeTextBoxMargins()`, `renderTextBoxMarginAttributes()`, `readTextBoxMargins()`, `replaceTextBoxMargins()`, and live shape accessors.

- [ ] **Step 1: Write a failing strict-read and lossless-edit model test**

Add a fixture test that injects a direct body-properties element before the first paragraph:

```ts
const marginBody = [
  '<a:bodyPr wrap="square" lIns="127000" tIns="-6350"',
  ' rIns="2147483647" bIns="1e3" custom="KEEP">',
  '<a:normAutofit fontScale="90000"/>',
  '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
  '</a:bodyPr>',
].join('');
```

Assert `shape.textMargins` returns only `{ left: 10, top: -0.5, right: 2147483647 / 12700 }`, omits malformed `bIns`, and does not add a mutation. Set `{ top: 4, left: 8 }`, then assert only `tIns="50800" lIns="101600"` remain while `wrap`, `custom`, `normAutofit`, the unknown child, all paragraphs/runs, and stable shape identity survive. Set `undefined` and assert all four inset attributes disappear without changing other body metadata.

- [ ] **Step 2: Run the model test and confirm it fails before implementation**

Run:

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: TypeScript/Vitest fails because `ShapeModel.textMargins` does not exist.

- [ ] **Step 3: Add public input and snapshot types**

In `packages/model/src/text.ts` add:

```ts
export interface TextBoxMargins {
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly left?: number;
}

export type TextBoxMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | TextBoxMargins;
```

- [ ] **Step 4: Implement the focused internal codec**

Create `packages/model/src/text-box-margins.internal.ts` with a fixed mapping:

```ts
const SIDES = [
  ['left', 'lIns'],
  ['top', 'tIns'],
  ['right', 'rIns'],
  ['bottom', 'bIns'],
] as const;
const EMU_PER_POINT = 12_700;
const MIN_INT32 = -2_147_483_648;
const MAX_INT32 = 2_147_483_647;
```

`normalizeTextBoxMargins(value, context)` must expand a scalar to four sides, map tuple indexes as top/right/bottom/left, reject arrays whose length is not four, reject non-object/non-number inputs, reject object keys outside the four supported names, normalize each provided side to `Math.round(value * 12700) / 12700`, and check the quantized raw signed-Int32 boundary. Preserve explicit zero by testing `!== undefined`.

`readTextBoxMargins(xml, shape, partUri)` must require direct `txBody/bodyPr`, read each direct attribute only when `/^-?\d+$/` matches and the numeric value is a safe signed Int32, and return a fresh object or `undefined` if no side is valid.

`renderTextBoxMarginAttributes(margins)` must emit attributes in `lIns/tIns/rIns/bIns` order using quantized raw integers. `replaceTextBoxMargins()` must extract the exact `bodyPr` template, reparse for each attribute update/removal so source spans stay valid, and replace only that element in the slide document.

- [ ] **Step 5: Delegate live reads and edits through the model**

Add to `SlideModel`:

```ts
getShapeTextMargins(id: number): TextBoxMargins | undefined;
setShapeTextMargins(id: number, value: TextBoxMarginInput | undefined): void;
```

The setter must run in `opcPackage.transaction()`, normalize the complete input before applying any attribute patch, save through `setXml()`, and throw `ModelParseError` for missing direct text-body structure.

Add asymmetric accessors to `ShapeModel`:

```ts
get textMargins(): TextBoxMargins | undefined {
  return this.slide.getShapeTextMargins(this.id);
}

set textMargins(value: TextBoxMarginInput | undefined) {
  this.slide.setShapeTextMargins(this.id, value);
}
```

- [ ] **Step 6: Complete model strictness, preservation, and rollback assertions**

Extend the fixture with decimal, scientific, empty, leading-plus, and out-of-range values distributed across multiple body properties. Assert malformed sides are omitted independently. Assert plain `.text`, rich `.richText`, and `setTransform()` preserve the margin XML; a setter inside an outer transaction restores the exact old bytes/snapshot on rollback; malformed shapes missing `txBody` or `bodyPr` throw without adding mutations.

- [ ] **Step 7: Run the focused model test**

Run the Task 1 Vitest command again. Expected: all model tests pass.

### Task 2: Plain/rich creation lifecycle and validation isolation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `TextBoxMarginInput`, `normalizeTextBoxMargins()`, `renderTextBoxMarginAttributes()`, `SlideModel.addText()`, `SlideModel.addRichText()`, transaction/duplicate/write/reopen paths.
- Produces: `AddTextOptions.margin`, identical plain/rich `bodyPr` serialization, and native lifecycle evidence.

- [ ] **Step 1: Write failing creation and lifecycle coverage**

Add SDK tests that create:

```ts
const uniform = slide.addText('Uniform', { margin: 10 });
const tuple = slide.addRichText([{ runs: [{ text: 'Tuple' }] }], {
  margin: [1, 2, 3, 4],
});
const named = slide.addText('Named', {
  margin: { top: -0.5, right: 0, bottom: 0.125, left: 8 },
});
```

Assert snapshots equal the quantized point values; XML maps tuple top/right/bottom/left to `tIns=12700`, `rIns=25400`, `bIns=38100`, `lIns=50800`; explicit zero is present; and omitted margin creates no inset attributes. Edit, partially replace, clear, duplicate, write, reopen, and verify shape identity plus values at each stage.

- [ ] **Step 2: Add `margin` to normalized creation options**

Extend `AddTextOptions`:

```ts
readonly margin?: TextBoxMarginInput;
```

Extend `NormalizedAddTextOptions` and `NormalizedTextInput` with a normalized margin object. Call `normalizeTextBoxMargins(options.margin, 'Text margin')` exactly once before any package mutation and pass the result to `addTextShape()`/`textShapeXml()` for both plain and rich paths.

- [ ] **Step 3: Serialize direct margins on new text shapes**

Change `textShapeXml()` to insert `renderTextBoxMarginAttributes(margins)` between `wrap="square"` and the existing `rtlCol="0" anchor="ctr"`. Do not change the current body-properties defaults when margin is omitted.

- [ ] **Step 4: Add invalid-input mutation-isolation cases**

Add scalar `NaN`, infinities, raw-overflow values, `null`, string, boolean, short/long tuples, a tuple with one invalid side, and objects with invalid values or an unknown key to the SDK invalid matrix. For both `addText()` and `shape.textMargins = value`, assert throws leave slide bytes, mutation journal, live text/margin snapshots, and object identity unchanged.

Use exact raw boundary cases:

```ts
const min = -2_147_483_648 / 12_700;
const max = 2_147_483_647 / 12_700;
```

Assert those quantize successfully while one-EMU-outside inputs fail.

- [ ] **Step 5: Run typecheck and focused model/SDK tests**

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
- Consumes: Task 1/2 public types and live API plus the existing adapter fixture and packed-package smoke.
- Produces: PptxGenJS 4.0.1 real-output evidence, corrected compatibility matrix, public usage documentation, and Node/browser/declaration smoke coverage.

- [ ] **Step 1: Add PptxGenJS-generated margin shapes**

Generate named shapes for omitted, zero, scalar 10, symmetric tuple `[4, 8, 4, 8]`, fractional `0.125`, and negative `-0.5` margins. Import the actual public `write({ outputType: 'uint8array' })` result and assert direct snapshots before and after this library writes/reopens the deck. Use the symmetric tuple for shared functional parity. Add a separate `[1, 2, 3, 4]` probe and assert its imported snapshot is `{ left: 1, top: 4, right: 2, bottom: 3 }`, documenting PptxGenJS 4.0.1's actual top/left swap without using that order in native creation.

- [ ] **Step 2: Correct and advance the compatibility surface**

Replace the inaccurate combined “普通段落 margin、first-line indent” row with two rows:

```md
| 文本框 `margin` scalar/TRBL 与 direct 四边编辑 | `AddTextOptions.margin` / `ShapeModel.textMargins` | 已支持 |
| paragraph 左右 margin、first-line/hanging indent（非 bullet） | 尚无完整公开 API | 尚未支持，后续逐项补齐 |
```

Document scalar/tuple/object semantics in `docs/api/README.md`, mention text-box margins in the npm README, and add one Unreleased changelog bullet. Keep fit/wrap/vertical/direction/RTL fields in the remaining partial-support row.

- [ ] **Step 3: Extend actual tarball smoke**

In Node smoke, create a scalar margin, assert the live snapshot, set a partial object, and verify the new snapshot. In browser smoke, create an explicit-zero tuple and verify all four sides. In the generated TypeScript smoke, import and use:

```ts
type TextBoxMarginInput,
type TextBoxMargins,

const margin: TextBoxMarginInput = [4, 8, 4, 8];
const marginSnapshot: TextBoxMargins | undefined = createdText.textMargins;
```

Ensure the packed declarations accept both the `AddTextOptions.margin` input and `ShapeModel.textMargins` setter.

- [ ] **Step 4: Run adapter, typecheck, and full repository tests**

Run the adapter test, full typecheck, full Vitest suite, and isolated performance suite. Expected: all functional tests pass, the default performance test remains skipped in the full run, and the explicit performance run passes.

### Task 4: Package, CLI, LibreOffice, review, commit, and push

**Files:**
- Review every file listed in Tasks 1–3; do not stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation/tests/docs.
- Produces: reviewed `feat: support text box margins` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual npm tarball**

Build from `packages/pptx` using the pinned local tools:

```sh
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
npm pack --ignore-scripts
```

Run `node scripts/smoke-npm-package.mjs packages/pptx/jiayunxie-pptx-0.1.0.tgz` from the repository root. Expected: API, browser, types, and CLI checks are all true.

- [ ] **Step 2: Validate and render native/PptxGenJS comparison decks**

Generate a native deck containing zero, uniform, and asymmetric named-object margins, plus a PptxGenJS deck containing zero, uniform, and symmetric tuple margins. Validate both with:

```sh
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-text-margin-validation/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-text-margin-validation/pptxgenjs.pptx --profile powerpoint-2010
```

Expected: zero errors and zero warnings. Export both through `soffice --headless --convert-to pdf`, confirm no repaired output or source overwrite, render PDFs with `pdftoppm -png`, and visually verify zero/uniform inset direction plus symmetric PptxGenJS/native parity.

- [ ] **Step 3: Final review**

Run `git diff --check`; inspect the complete diff; verify every changed line traces to margins; confirm no generic body-properties abstraction or paragraph-margin behavior leaked into this item; verify `git status --short` lists only the intended implementation files plus untracked `.pnpm-store/`. Re-run focused tests after any review fix.

- [ ] **Step 4: Commit and push**

Stage only the intended implementation files, commit with:

```sh
git commit -m "feat: support text box margins"
```

Push `main` with the repository SSH-over-443 command and verify `git status --short --branch` reports `main...origin/main` plus only untracked `.pnpm-store/`.
