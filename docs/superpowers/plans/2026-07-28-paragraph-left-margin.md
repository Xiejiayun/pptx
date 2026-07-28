# Paragraph Left Margin Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native non-list paragraph left-margin create/read/edit/clear support without changing bullet/numbering indentation.

**Architecture:** Extend the existing paragraph-property pipeline in `rich-text.internal.ts` with a strict point-to-EMU `marL` codec. Resolve outer defaults and paragraph overrides before rendering, reject a numeric left margin when the same paragraph has an active bullet, patch `marL` only after bullet rendering, and suppress the public field when direct list markup owns the attribute. Keep the existing canonical `marL="0"` creation default while allowing rich replacement to clear direct state.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 conformance fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation default is `AddTextOptions.paragraphMarginLeft?: number`; paragraph snapshot/override is `RichTextParagraph.marginLeft?: number | false`.
- Values use points, must be finite and within `0..4032`, and serialize as `Math.round(points * 12700)` in `a:pPr@marL`.
- New non-list paragraphs with no explicit margin retain canonical direct `marL="0"`; `marginLeft: false` writes no direct `marL`.
- Existing-shape `richText` replacement treats omitted `marginLeft` or false as clear for a non-list paragraph; plain `.text` continues to preserve and clone the first paragraph template.
- Direct `buChar`, `buAutoNum`, or `buBlip` owns `marL`; the getter suppresses `marginLeft` and rendering never clears the bullet-computed value.
- A resolved numeric left margin plus an active bullet/numbering on the same paragraph fails before package mutation; false may suppress an outer margin so a bullet paragraph remains valid.
- Do not add `marR`, first-line/hanging `indent`, inheritance resolution, RTL side swapping, or unrelated refactors.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, and LibreOffice rendering pass.

---

### Task 1: Strict paragraph model and lossless `marL` codec

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `RichTextParagraph`, `NormalizedRichTextParagraph`, `renderRichTextParagraphs()`, `renderParagraphProperties()`, `updateParagraphAttribute()`, `readRichText()`, bullet direct-child parsing, and paragraph templates.
- Produces: `RichTextParagraph.marginLeft?: number | false`, exported strict normalization, resolved margin rendering, direct getter, and exact non-list `pPr@marL` replacement/removal.

- [ ] **Step 1: Add failing strict-read and bullet-isolation tests**

Create existing paragraphs whose `marL` values are `0`, `152400`, `1`, missing, empty, `1.5`, `-1`, `51206401`, and `yes`; add four `marL="228600"` paragraphs containing direct `buNone`, `buChar`, `buAutoNum`, and `buBlip`. Assert:

```ts
expect(shape.richText.map(({ marginLeft }) => marginLeft)).toEqual([
  0,
  12,
  1 / 12700,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  18,
  undefined,
  undefined,
  undefined,
]);
expect(pkg.mutations).toEqual(journal);
```

Use plain `.text = 'First\nSecond'` on the valid 12-point paragraph and assert both output paragraphs preserve exact `marL="152400"` plus unrelated attributes/children.

- [ ] **Step 2: Run the model suite and confirm the new assertions fail**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: failure because `marginLeft` is not part of `RichTextParagraph` or the getter.

- [ ] **Step 3: Add the public field and strict normalization**

Add to `packages/model/src/text.ts`:

```ts
export interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly align?: TextAlignment;
  readonly rtl?: boolean;
  readonly marginLeft?: number | false;
  readonly bullet?: ParagraphBullet;
  readonly level?: number;
  readonly spacing?: ParagraphSpacing | false;
  readonly tabStops?: readonly ParagraphTabStop[] | false;
}
```

In `rich-text.internal.ts`, allow `marginLeft` in the supported key list and normalize it before rendering:

```ts
export function normalizeParagraphMarginLeft(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > 4032) {
    throw new RangeError(`${context} must be between 0 and 4032 points`);
  }
  return value;
}
```

Normalize the paragraph field with the exact three-state branch:

```ts
const marginLeft = candidate.marginLeft === undefined
  ? undefined
  : candidate.marginLeft === false
    ? false
    : normalizeParagraphMarginLeft(
        candidate.marginLeft,
        `Rich text paragraph ${paragraphIndex} marginLeft`,
      );
```

Preserve explicit false in the normalized paragraph object; do not round the public point value before EMU conversion.

- [ ] **Step 4: Resolve paragraph defaults and reject list conflicts**

Add `defaultMarginLeft?: number` to `RenderRichTextOptions`. For each paragraph, resolve the active bullet and local/default margin before calling `renderParagraphProperties()`:

```ts
const resolvedBullet = bullet === false
  ? undefined
  : bullet ?? (options.defaultBullet === false ? undefined : options.defaultBullet);
const resolvedMarginLeft = marginLeft === false
  ? false
  : marginLeft ?? options.defaultMarginLeft;
if (resolvedBullet && typeof resolvedMarginLeft === 'number') {
  throw new TypeError('Paragraph left margin cannot be combined with an active bullet');
}
```

Pass the resolved bullet and `number | false | undefined` to `renderParagraphProperties()`.

- [ ] **Step 5: Patch `marL` after bullet rendering**

After `renderParagraphBullet()`, update only a non-list paragraph:

```ts
function renderParagraphLeftMargin(
  template: string,
  marginLeft: number | false | undefined,
  hasActiveBullet: boolean,
  isNewParagraph: boolean,
): string {
  if (hasActiveBullet) return template;
  if (marginLeft === undefined && isNewParagraph) return template;
  const value = typeof marginLeft === 'number'
    ? String(Math.round(marginLeft * EMU_PER_POINT))
    : undefined;
  return updateParagraphAttribute(template, 'marL', value);
}
```

Call this before tab-stop rendering. `isNewParagraph` is `template === undefined`; it preserves the existing canonical direct zero only for newly rendered paragraphs. Existing paragraph templates clear `marL` when the public field is omitted or false.

- [ ] **Step 6: Read only valid direct non-list values**

Add `readParagraphMarginLeft()` that finds direct `pPr`, returns undefined when any direct active bullet child is present, parses a full decimal integer, accepts only `0..51206400`, and returns `value / EMU_PER_POINT`. Include `marginLeft` in `readRichText()` only when defined. Re-run the model suite and expect it to pass after updating exact snapshots that now legitimately include direct zero/nonzero margins.

### Task 2: Plain/rich creation defaults and lifecycle validation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 normalization/rendering, `AddTextOptions`, `validateAddTextOptions()`, `validateTextInput()`, `textParagraphXml()`, transactions, duplicate, and write/reopen.
- Produces: `AddTextOptions.paragraphMarginLeft?: number`, plain/rich defaults and overrides, mutation-isolated failures, and native lifecycle evidence.

- [ ] **Step 1: Add failing plain/rich creation tests**

Create plain shapes for omitted, `0`, `12.5`, and `4032` across CR/LF paragraphs and empty lines. Create rich text with an outer 24-point default, a 12-point override, and false suppression. Assert public snapshots and exact XML values:

```ts
expect(shape.richText.map(({ marginLeft }) => marginLeft)).toEqual([24, 12, undefined]);
expect(slideXml).toContain('marL="304800"');
expect(slideXml).toContain('marL="152400"');
```

Combine margin paragraphs with alignment, RTL, spacing, and tab stops. In another paragraph of the same rich shape, use `marginLeft: false` plus a bullet and assert its existing bullet `marL`/`indent` pair remains correct.

- [ ] **Step 2: Add and normalize the outer creation option**

Add to `AddTextOptions`:

```ts
readonly paragraphMarginLeft?: number;
```

Add `marginLeft?: number` to `NormalizedAddTextOptions` and `NormalizedTextInput`; normalize with `normalizeParagraphMarginLeft(options.paragraphMarginLeft, 'Paragraph left margin')`. Pass it to every plain `textParagraphXml()` call and as `defaultMarginLeft` for rich rendering.

- [ ] **Step 3: Reject only resolved plain conflicts**

`validateAddTextOptions()` must allow outer bullet and margin defaults because individual rich paragraphs can suppress either one. In `validateTextInput()`, where plain text has no local override, reject the resolved pair:

```ts
if (defaults.bullet && defaults.marginLeft !== undefined) {
  throw new TypeError('Paragraph left margin cannot be combined with an active bullet');
}
```

Pass the normalized margin to `textParagraphXml()` and then `renderParagraphProperties()`; keep `bullet: false` plus a margin valid.

- [ ] **Step 4: Cover edit, clear, duplicate, reopen, and rollback**

Replace existing non-list paragraphs with positive, zero, false, and omitted margins; assert false/omitted remove only direct `marL`. Duplicate before edits and prove the duplicate retains the original values. Verify plain `.text` clones the first template margin. Roll back an outer transaction after a margin update and assert exact bytes, mutation journal, slide identity, and shape identity are restored. Write/reopen and validate all supported values.

- [ ] **Step 5: Add invalid-input mutation isolation**

Test outer values `null`, true, false, strings, objects, arrays, symbols, NaN, Infinity, `-0.01`, and `4032.01`. Test the same invalid paragraph values except false, which is the valid suppression form already covered above. Also test resolved bullet/margin conflicts for plain, rich outer defaults, and local overrides. Assert TypeError/RangeError as designed and unchanged package bytes/journal before any shape is allocated.

- [ ] **Step 6: Run typecheck and focused suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck exits 0 and both suites pass.

### Task 3: Conformance regression, documentation, and packed API

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public API and existing PptxGenJS import/write/reopen fixtures.
- Produces: default-zero and bullet-indentation regression evidence, partial-support documentation, and Node/browser/declaration packed coverage.

- [ ] **Step 1: Extend PptxGenJS real-output conformance**

Generate PptxGenJS 4.0.1 ordinary plain/rich text plus bullet and numbered paragraphs. Assert imported ordinary paragraphs expose direct `marginLeft: 0`, direct bullet paragraphs suppress it while preserving `bullet.indent`, and raw XML retains the expected `marL="0" indent="0"` and bullet `marL`/negative `indent` pairs. Reopen the imported document and repeat the public assertions. Do not access PptxGenJS private fields.

- [ ] **Step 2: Update compatibility and API documentation**

Change the paragraph margin/indent matrix row from unsupported to partial: direct non-list left margin is supported, right margin and first-line/hanging indent remain unsupported. Document point units, zero versus false/absence, text-box margin separation, bullet conflict behavior, strict direct reads, and PptxGenJS's lack of a non-list public option. Add one precise changelog bullet and examples to API/package README.

- [ ] **Step 3: Extend packed Node/browser/declaration smoke**

In Node smoke, create 24/12/false rich paragraphs, edit to 6/0/omitted, and assert snapshots plus bullet independence. In browser smoke, create and clear one left margin. In the declaration fixture, compile:

```ts
const paragraphMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { runs: [{ text: 'Override' }], marginLeft: 12 },
  { runs: [{ text: 'Inherited' }], marginLeft: false },
];
createdDocument.addSlide().addRichText(paragraphMargins, { paragraphMarginLeft: 24 });
```

- [ ] **Step 4: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/performance/src/performance.test.ts
```

Expected: adapter and all functional tests pass, the default performance case is the only full-run skip, and isolated performance passes.

### Task 4: Package, visual validation, final review, and delivery

**Files:**
- Review all Task 1–3 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and all earlier evidence.
- Produces: reviewed `feat: support paragraph left margin` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
paragraph_margin_package_dir=$(mktemp -d /tmp/pptx-paragraph-left-margin-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$paragraph_margin_package_dir"
node ../../scripts/smoke-npm-package.mjs "$paragraph_margin_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke checks all report true.

- [ ] **Step 2: Generate and validate native comparison files**

Generate `/tmp/pptx-paragraph-left-margin-native/native.pptx` with visible guide lines and 0/12/24/48-point paragraph left margins. Generate a baseline with the same valid `pPr@marL` values using a hand-patched copy because PptxGenJS has no non-list API. Validate both with:

```sh
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-paragraph-left-margin-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-paragraph-left-margin-baseline/baseline.pptx --profile powerpoint-2010
```

Expect zero errors and zero warnings, and inspect `/ppt/slides/slide1.xml` to confirm only intended non-list `marL` values differ.

- [ ] **Step 3: Export and inspect every rendered page**

Use `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf` for both files, render PDFs with `pdftoppm -png`, and visually inspect every page. Confirm text starts progress monotonically at 0/12/24/48 points, bullet anchors remain unchanged, and there is no repair output, clipping, overlap, or overflow.

- [ ] **Step 4: Review, commit, push, and verify remote state**

Run `git diff --check`, inspect the entire diff, verify omission/false/direct-zero semantics, bullet isolation, and no `marR`/`indent` feature expansion. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, run `git diff --cached --check`, commit:

```sh
git commit -m "feat: support paragraph left margin"
```

Push through SSH-over-443 and verify:

```sh
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.
