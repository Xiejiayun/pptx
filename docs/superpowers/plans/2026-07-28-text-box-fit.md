# Text Box Fit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native create/read/edit/clear support for PptxGenJS-compatible text-box none, shrink, and resize fit modes.

**Architecture:** Add a focused fit codec that owns only the direct `bodyPr` text-autofit choice children, preserves PowerPoint-calculated shrink metadata on same-mode writes, and inserts canonical children in schema order. Keep creation wire-compatible with PptxGenJS by representing omitted/none as no child, then expose strict direct reads and transactional live edits through `ShapeModel.textFit`.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation API is `AddTextOptions.fit?: TextBoxFit`; live editing API is `ShapeModel.textFit`.
- Valid values are exactly `none | shrink | resize`.
- Creation omitted/none writes no autofit child; shrink writes direct `a:normAutofit`; resize writes direct `a:spAutoFit`.
- Read an existing unique direct `a:noAutofit` as none, `a:normAutofit` as shrink, and `a:spAutoFit` as resize; absent or multiple supported children return `undefined`.
- Live setter none/undefined removes all direct supported fit children; same-mode shrink/resize preserves the existing child byte-for-byte.
- Preserve unsupported children, bodyPr attributes, PowerPoint-calculated shrink metadata, and all non-fit content during getters and unrelated edits.
- Do not calculate fontScale, line-spacing reduction, text overflow, or shape dimensions.
- Table-cell fit, inheritance, overflow, word breaking, columns, and deprecated native aliases are outside this item.
- Adapter conformance must still import PptxGenJS deprecated `shrinkText` / `autoFit` public output.
- PptxGenJS is used only by the adapter conformance test and never becomes a non-adapter dependency.
- Do not stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.
- Execute inline without questions or subagents under the user's delegated authority.

---

### Task 1: Public fit type, direct-child codec, and live accessors

**Files:**
- Create: `packages/model/src/text-box-fit.internal.ts`
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `requireTextBodyProperties()`, `LosslessXmlDocument`, `XmlElement`, `SlideModel` delegation, stable shape identity, and bodyPr child source spans.
- Produces: `TextBoxFit`, fit normalize/render/read/replace functions, `SlideModel.getShapeTextFit()` / `setShapeTextFit()`, and `ShapeModel.textFit`.

- [ ] **Step 1: Write failing lossless model tests**

Inject an expanded direct bodyPr with fit metadata and post-fit children:

```ts
const bodyProperties = [
  `<a:bodyPr wrap="none" lIns="127000" anchor="b" vert="vert270" custom="KEEP">`,
  '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>',
  `<a:normAutofit fontScale='85000' lnSpcReduction="20000"/>`,
  '<a:scene3d><a:camera prst="orthographicFront"/></a:scene3d>',
  '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
  '</a:bodyPr>',
].join('');
```

Assert repeated reads return `shrink` without mutations. Reassign shrink and assert exact bytes/journal remain unchanged, including mixed quotes and calculated attributes. Switch to resize and verify only the old fit child becomes `<a:spAutoFit/>` at the same position; run plain/rich text, margin, vertical-alignment, wrapping, direction, and transform edits and verify fit plus all unowned XML survive. Roll back a mode change, clear with undefined, then add shrink into a fit-less expanded bodyPr and assert insertion before `scene3d`.

- [ ] **Step 2: Run the model test and confirm the red state**

Run:

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new tests fail because the type, codec, and `ShapeModel.textFit` do not exist.

- [ ] **Step 3: Add the public type and strict codec**

In `packages/model/src/text.ts` add:

```ts
export type TextBoxFit = 'none' | 'shrink' | 'resize';
```

Create `text-box-fit.internal.ts` with:

```ts
export function normalizeTextBoxFit(value: unknown, context: string): TextBoxFit;
export function renderTextBoxFitChild(value: TextBoxFit): string;
export function readTextBoxFit(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): TextBoxFit | undefined;
export function replaceTextBoxFit(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: TextBoxFit | undefined,
  partUri: string,
): void;
```

Normalization accepts only the exact union. Render returns `''` for none, `<a:normAutofit/>` for shrink, and `<a:spAutoFit/>` for resize.

Use a local direct-child helper and map:

```ts
const FROM_OOXML = new Map([
  ['noAutofit', 'none'],
  ['normAutofit', 'shrink'],
  ['spAutoFit', 'resize'],
] as const);
```

Read only when the bodyPr has exactly one supported direct child. Replace by parsing `xml.original(bodyProperties)` into a focused document, then:

- remove every supported direct child for none/undefined;
- return the original unchanged for a sole matching norm/sp child;
- otherwise replace the first old choice with a canonical child and remove extras;
- when no choice exists, expand a self-closing bodyPr with `appendChildXml()`, or insert before the first direct `scene3d`, `sp3d`, or `extLst`, otherwise append before the closing tag;
- use the bodyPr element prefix for the canonical child instead of assuming `a:`.

Replace the outer bodyPr element with the focused document serialization. Identical serialization is safe because `OpcPackage.setPart()` does not journal equal bytes.

- [ ] **Step 4: Add live SlideModel/ShapeModel delegation**

Add:

```ts
getShapeTextFit(id: number): TextBoxFit | undefined;
setShapeTextFit(id: number, value: TextBoxFit | undefined): void;
```

The setter validates non-undefined input before resolving/patching, runs inside `opcPackage.transaction()`, and saves only the serialized slide XML. Add to `ShapeModel`:

```ts
get textFit(): TextBoxFit | undefined;
set textFit(value: TextBoxFit | undefined);
```

- [ ] **Step 5: Complete direct-choice and malformed coverage**

Use fresh fixtures for absent, sole no/norm/sp, case variants, nested descendants, unknown children, duplicate same children, and mixed supported children. Assert the three sole choices surface, every ambiguous/unsupported structure returns undefined, and reads do not mutate. For an explicit `noAutofit`, run all unrelated text-body/transform edits and verify it survives unchanged. Extend malformed direct `txBody/bodyPr` tests to fit getter/setter with exact byte/journal preservation.

- [ ] **Step 6: Run focused model regressions**

Run the Task 1 command. Expected: all model tests pass, including margins, vertical alignment, wrapping, and direction.

### Task 2: Optional plain/rich creation and invalid-input isolation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public type/codec, `AddTextOptions`, shared text-shape creation, duplicate/write/reopen transactions.
- Produces: optional `AddTextOptions.fit`, PptxGenJS-compatible creation XML, and lifecycle evidence.

- [ ] **Step 1: Write failing creation/lifecycle tests**

Create omitted, none, shrink, and resize shapes across plain/rich paths:

```ts
const omitted = slide.addText('Omitted');
const none = slide.addText('None', { fit: 'none' });
const shrink = slide.addRichText(
  [{ align: 'center', runs: [{ text: 'Shrink', style: { bold: true } }] }],
  { fit: 'shrink', wrap: false, margin: 8, valign: 'bottom', vert: 'vert' },
);
const resize = slide.addText('Resize', { fit: 'resize' });
```

Assert omitted/none snapshots are undefined and have self-closing bodyPr without a fit child; shrink/resize expose exact values and have expanded bodyPr with one child. Duplicate before edits, preserve fit through all other text-body edits, switch modes, clear one, add back to a self-closing bodyPr, test transaction rollback, then write/reopen and compare original/duplicate snapshots.

- [ ] **Step 2: Normalize optional creation fit and render the child**

Extend `AddTextOptions` with `readonly fit?: TextBoxFit`. Extend `NormalizedAddTextOptions` / `NormalizedTextInput` with optional `textFit`. Normalize only when provided:

```ts
const textFit = options.fit === undefined
  ? undefined
  : normalizeTextBoxFit(options.fit, 'Text fit');
```

Pass it through `addTextShape()` to `textShapeXml()`. Build the bodyPr XML as:

```ts
const fitChild = textFit === undefined ? '' : renderTextBoxFitChild(textFit);
const bodyProperties = fitChild
  ? `<a:bodyPr${attributes}>${fitChild}</a:bodyPr>`
  : `<a:bodyPr${attributes}/>`;
```

Omitted/none must leave the existing serialized bodyPr bytes identical.

- [ ] **Step 3: Add invalid-input mutation-isolation cases**

Test null, booleans, numbers, empty string, case/whitespace variants, unknown strings, arrays, objects, and symbols. For `addText()`, `addRichText()`, and live setter, assert `TypeError` and preservation of bytes, journal, text, margins, vertical alignment, wrapping, direction, fit snapshot, and stable identity.

- [ ] **Step 4: Run typecheck and focused model/SDK tests**

Run:

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck exits 0 and both suites pass.

### Task 3: PptxGenJS conformance and public surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/api/README.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 1/2 API and existing PptxGenJS public-write fixture.
- Produces: real 4.0.1 fit evidence, release docs, and packed Node/browser/type smoke.

- [ ] **Step 1: Add PptxGenJS fit conformance shapes**

Generate named shapes for omitted, none, shrink, resize, invalid outer fit, ignored run-level fit/legacy flags, legacy outer `shrinkText: true`, and legacy outer `autoFit: true`. Import public bytes and assert snapshots:

```ts
[
  undefined,
  undefined,
  'shrink',
  'resize',
  undefined,
  undefined,
  'shrink',
  'resize',
]
```

Assert raw XML has no child for omitted/none/invalid/run, exactly norm/sp for modern/legacy valid shapes, and no invalid value serialization. Duplicate/write/reopen and verify identical snapshots. Keep the existing 10-second timeout unless real full-suite contention requires a local increase.

- [ ] **Step 2: Update compatibility and API documentation**

Add:

```md
| 文本框 `fit: none/shrink/resize` 与 direct 编辑 | `AddTextOptions.fit` / `ShapeModel.textFit` | 已支持 |
```

Remove ordinary text-box fit from the remaining partial row and explicitly leave table-cell fit unsupported. Document omitted/none equivalence, existing noAutofit reads, direct clear semantics, PowerPoint dynamic scaling, strict invalid input, and deprecated adapter coverage. Update the npm README and add one Unreleased changelog bullet.

- [ ] **Step 3: Extend packed-package smoke**

In Node smoke, create shrink, set resize, set none, then clear while checking snapshots. In browser smoke, create resize and assert it. In generated TypeScript smoke, import/use:

```ts
type TextBoxFit,

const fit: TextBoxFit = 'shrink';
const fitted = createdDocument.addSlide().addText('Fitted', { fit });
const fitSnapshot: TextBoxFit | undefined = fitted.textFit;
fitted.textFit = 'resize';
fitted.textFit = 'none';
fitted.textFit = undefined;
```

- [ ] **Step 4: Run adapter, typecheck, full tests, and performance**

Run adapter, full typecheck, complete Vitest, and isolated `RUN_PERF=1` performance. Expected: every functional test passes, default performance remains skipped, and isolated performance passes.

### Task 4: Package, CLI, LibreOffice, review, commit, and push

**Files:**
- Review all Task 1–3 files; never stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation/tests/docs.
- Produces: reviewed `feat: support text box fit modes` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the real tarball**

Run pinned tsup 8.5.1 Node/browser builds plus declaration builder from `packages/pptx`, then:

```sh
package_dir=$(mktemp -d /tmp/pptx-text-fit-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$package_dir"
node ../../scripts/smoke-npm-package.mjs "$package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser conditional export, declarations, and installed CLI all pass.

- [ ] **Step 2: Validate and render native/PptxGenJS comparison decks**

Generate wide-layout decks in a fresh `/tmp/pptx-text-fit-validation.*` directory with equal-size boxes for none, shrink, and resize. Use the same deliberately overflowing text, margin, wrap, vertical alignment, and dimensions. Validate both with repository CLI `powerpoint-2010`: 0 errors, 0 warnings.

Record source hashes, export both through LibreOffice, verify hashes unchanged and no repair output, render page PNGs at 150 DPI, and visually compare clipping/shrinking/shape resizing. Record that PowerPoint may calculate final normAutofit scale only after a user edit/resize, as documented by PptxGenJS.

- [ ] **Step 3: Final review**

Run `git diff --check`; inspect the complete diff and the new untracked codec; verify the codec owns only the direct fit choice; confirm omitted/none creation is byte-identical; confirm calculated normAutofit attributes survive same-mode and unrelated edits; confirm other text-body codecs/tests did not regress; ensure every changed line belongs to this item and status shows only intended files plus `.pnpm-store/`. Re-run focused tests after review fixes.

- [ ] **Step 4: Commit and push**

Stage intended files only and commit:

```sh
git commit -m "feat: support text box fit modes"
```

Push through SSH-over-443, verify `origin/main...HEAD` is `0 0`, and immediately continue with the next unsupported PptxGenJS surface.
