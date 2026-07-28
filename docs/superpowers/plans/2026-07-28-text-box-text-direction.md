# Text Box Text Direction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native create/read/edit/clear support for all seven valid text-box directions with PptxGenJS 4.0.1 output conformance.

**Architecture:** Add a focused text-direction codec that validates the seven public `vert` tokens and patches only direct `bodyPr@vert` through the shared body-properties helper. Keep creation optional so omitted direction produces no attribute, then expose live reads/edits as `ShapeModel.textDirection` without interpreting inheritance or rotating shape transforms.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation API is `AddTextOptions.vert?: TextBoxTextDirection`; live editing API is `ShapeModel.textDirection`.
- Valid values are exactly `eaVert | horz | mongolianVert | vert | vert270 | wordArtVert | wordArtVertRtl` and serialize unchanged to direct `bodyPr@vert`.
- Omitted creation writes no `vert`; live setter `undefined` removes the direct override.
- Read only exact direct valid tokens; absent, empty, case variants, whitespace, and unknown tokens return `undefined`.
- Preserve unsupported vert tokens and all unowned body properties during getters and unrelated edits.
- Table-cell `textDirection`, paragraph/run RTL, language, `rtlCol`, fit/overflow, columns, and shape rotation are outside this item.
- Do not expose PptxGenJS 4.0.1's ignored ordinary-text `textDirection` alias or run-level direction as working APIs.
- PptxGenJS is used only by the adapter conformance test and never becomes a non-adapter dependency.
- Do not stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.
- Execute inline without questions or subagents under the user's delegated authority.

---

### Task 1: Public direction type, strict codec, and live accessors

**Files:**
- Create: `packages/model/src/text-box-text-direction.internal.ts`
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `requireTextBodyProperties()`, `updateTextBodyAttribute()`, `LosslessXmlDocument`, `XmlElement`, `SlideModel` delegation, and stable shape identity.
- Produces: `TextBoxTextDirection`, direction normalize/render/read/replace functions, `SlideModel.getShapeTextDirection()` / `setShapeTextDirection()`, and `ShapeModel.textDirection`.

- [ ] **Step 1: Write a failing lossless-edit model test**

Inject expanded direct body properties with single-quoted direction and unrelated attributes/children:

```ts
const bodyProperties = [
  `<a:bodyPr wrap="none" lIns="127000" anchor="b" vert='vert270' rtlCol="1" custom="KEEP">`,
  '<a:normAutofit fontScale="90000"/>',
  '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
  '</a:bodyPr>',
].join('');
```

Assert `shape.textDirection === 'vert270'`, repeated reads add no mutation, and text margins / vertical alignment / wrapping snapshots remain correct. Set `wordArtVert` and verify the existing quote style remains single; then set `eaVert`. Run plain/rich text, margin, vertical-alignment, wrapping, and transform edits and verify only their owned fields change. Roll back a direction change, then clear with `undefined`, preserving all unowned metadata and stable identity.

- [ ] **Step 2: Run the model test and confirm the red state**

Run:

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new tests fail because the type, codec, and `ShapeModel.textDirection` do not exist.

- [ ] **Step 3: Add the public type and strict direction codec**

In `packages/model/src/text.ts` add:

```ts
export type TextBoxTextDirection =
  | 'eaVert'
  | 'horz'
  | 'mongolianVert'
  | 'vert'
  | 'vert270'
  | 'wordArtVert'
  | 'wordArtVertRtl';
```

Create `text-box-text-direction.internal.ts` with an exact readonly list/set and:

```ts
export function normalizeTextBoxTextDirection(
  value: unknown,
  context: string,
): TextBoxTextDirection;
export function renderTextBoxTextDirectionAttribute(value: TextBoxTextDirection): string;
export function readTextBoxTextDirection(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): TextBoxTextDirection | undefined;
export function replaceTextBoxTextDirection(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: TextBoxTextDirection | undefined,
  partUri: string,
): void;
```

Normalization accepts only exact list members. Render returns ` vert="${value}"`. Read only an exact direct token. Replace only `vert` through `updateTextBodyAttribute()`.

- [ ] **Step 4: Add live SlideModel/ShapeModel delegation**

Add:

```ts
getShapeTextDirection(id: number): TextBoxTextDirection | undefined;
setShapeTextDirection(id: number, value: TextBoxTextDirection | undefined): void;
```

The setter validates non-undefined input before patching, executes within `opcPackage.transaction()`, resolves the live shape, and saves with `setXml()`.

Add to `ShapeModel`:

```ts
get textDirection(): TextBoxTextDirection | undefined;
set textDirection(value: TextBoxTextDirection | undefined);
```

- [ ] **Step 5: Complete seven-token, unknown-token, and malformed coverage**

Use fresh fixtures for all seven valid tokens plus absent, empty, `Vert`, ` vert `, `vertical`, and unknown. Assert only valid exact values surface and reads do not mutate. For `vertical`, run `.text`, `.richText`, `textMargins`, `verticalAlignment`, `textWrap`, and transform edits and verify the unsupported token survives. Extend malformed direct `txBody/bodyPr` tests to direction getter/setter with exact byte/journal preservation.

- [ ] **Step 6: Run focused model regressions**

Run the Task 1 command. Expected: all model tests pass, including margin, vertical-alignment, and wrapping suites.

### Task 2: Optional plain/rich creation and invalid-input isolation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 public type/codec, `AddTextOptions`, shared text-shape creation, duplicate/write/reopen transactions.
- Produces: optional `AddTextOptions.vert`, all seven creation values, and lifecycle evidence.

- [ ] **Step 1: Write failing creation/lifecycle tests**

Create one omitted shape plus all seven valid directions, using both plain and rich text and combining selected cases with margins, wrap, vertical alignment, paragraph alignment, and run style:

```ts
const omitted = slide.addText('Omitted');
const vertical = slide.addRichText(
  [{ align: 'center', runs: [{ text: 'Vertical', style: { bold: true } }] }],
  { vert: 'vert', wrap: false, margin: 8, valign: 'bottom' },
);
const rotated = slide.addText('Rotated', { vert: 'vert270' });
```

Assert omitted is `undefined` and every explicit shape returns its exact token. Verify raw XML, duplicate before edits, replace directions, clear one, add back to a vert-less self-closing bodyPr, and verify edited/duplicated snapshots after write/reopen. Ensure text/margin/vertical-alignment/wrapping/transform edits preserve direction.

- [ ] **Step 2: Normalize optional creation direction and render it**

Extend `AddTextOptions` with `readonly vert?: TextBoxTextDirection`. Extend `NormalizedAddTextOptions` / `NormalizedTextInput` with optional `textDirection`. Normalize only when provided:

```ts
const textDirection = options.vert === undefined
  ? undefined
  : normalizeTextBoxTextDirection(options.vert, 'Text direction');
```

Pass the optional value through `addTextShape()` to `textShapeXml()`. Append `renderTextBoxTextDirectionAttribute(textDirection)` after the existing vertical-anchor attribute only when present; omission must leave current serialized bytes unchanged.

- [ ] **Step 3: Add invalid-input mutation-isolation cases**

Test null, booleans, numbers, empty string, case/whitespace variants, invalid strings, arrays, objects, and symbols. For `addText()`, `addRichText()`, and live setter, assert `TypeError` and preservation of bytes, journal, text, margin, vertical alignment, wrapping, direction snapshot, and stable identity.

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
- Produces: real 4.0.1 direction evidence, release docs, and packed Node/browser/type smoke.

- [ ] **Step 1: Add PptxGenJS direction conformance shapes**

Generate named shapes for omitted, all seven valid `vert` values, an invalid passthrough string, an outer `textDirection: 'vert'` alias with no `vert`, and one `TextProps[]` run carrying `vert` / `textDirection`. Import public bytes and assert `undefined`, the exact seven values, then `undefined` for invalid/ignored cases. Assert raw XML preserves the invalid token but omits direction for ignored alias/run cases. Duplicate/write/reopen and verify identical snapshots.

Increase only this existing large integration test's explicit timeout if full-suite contention exceeds its current 10 seconds; do not relax global timeouts.

- [ ] **Step 2: Update compatibility and API documentation**

Add:

```md
| 文本框 `vert` 七种文本方向与 direct 编辑 | `AddTextOptions.vert` / `ShapeModel.textDirection` | 已支持 |
```

Remove ordinary text-box vert from the remaining partial row and explicitly leave table-cell `textDirection` unsupported. Document omitted-vs-explicit-horz and clear semantics, update the npm README, and add one Unreleased changelog bullet.

- [ ] **Step 3: Extend packed-package smoke**

In Node smoke, create `vert270`, set `wordArtVert`, then clear while checking all snapshots. In browser smoke, create `vert` and assert it. In generated TypeScript smoke, import/use:

```ts
type TextBoxTextDirection,

const direction: TextBoxTextDirection = 'vert270';
const directed = createdDocument.addSlide().addText('Directed', { vert: direction });
const directionSnapshot: TextBoxTextDirection | undefined = directed.textDirection;
directed.textDirection = 'wordArtVert';
directed.textDirection = undefined;
```

- [ ] **Step 4: Run adapter, typecheck, full tests, and performance**

Run adapter, full typecheck, complete Vitest, and isolated `RUN_PERF=1` performance. Expected: every functional test passes, default performance remains skipped, and isolated performance passes.

### Task 4: Package, CLI, LibreOffice, review, commit, and push

**Files:**
- Review all Task 1–3 files; never stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation/tests/docs.
- Produces: reviewed `feat: support text box text direction` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the real tarball**

Run pinned tsup 8.5.1 Node/browser builds plus declaration builder from `packages/pptx`, then:

```sh
package_dir=$(mktemp -d /tmp/pptx-text-direction-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$package_dir"
node ../../scripts/smoke-npm-package.mjs "$package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser conditional export, declarations, and installed CLI all pass.

- [ ] **Step 2: Validate and render native/PptxGenJS comparison decks**

Generate wide-layout decks in a fresh `/tmp/pptx-text-direction-validation.*` directory with equal-size filled boxes for `horz`, `vert`, `vert270`, and `wordArtVert`, using the same text, margin, wrap, and vertical alignment. Validate both with repository CLI `powerpoint-2010`: 0 errors, 0 warnings.

Record source hashes, export both through LibreOffice, verify hashes unchanged and no repair output, render page PNGs at 150 DPI, and visually compare direction. Record any identical LibreOffice degradation for `eaVert`, `mongolianVert`, or `wordArtVertRtl` if extra diagnostic boxes are included.

- [ ] **Step 3: Final review**

Run `git diff --check`; inspect complete diff; verify codec owns only `vert`; confirm omitted creation is byte-identical; confirm other text-body codecs/tests did not regress; ensure every changed line belongs to this item and status shows only intended files plus `.pnpm-store/`. Re-run focused tests after review fixes.

- [ ] **Step 4: Commit and push**

Stage intended files only and commit:

```sh
git commit -m "feat: support text box text direction"
```

Push through SSH-over-443, verify `origin/main...HEAD` is `0 0`, and immediately continue with the next unsupported PptxGenJS surface.
