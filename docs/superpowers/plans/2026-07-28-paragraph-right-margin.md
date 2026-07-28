# Paragraph Right Margin Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native paragraph right-margin create/read/edit/clear support without changing left-margin or bullet/numbering indentation behavior.

**Architecture:** Extend the existing paragraph-property pipeline in `rich-text.internal.ts` with a strict point-to-EMU `marR` codec parallel to `marL`. Resolve outer defaults and paragraph overrides before rendering, patch `marR` independently after existing bullet and left-margin rendering, and read any valid direct value even when the paragraph has a bullet because list indentation does not own `marR`. Keep `marR` absent for omitted new paragraphs and clear direct state on omitted existing-shape replacement.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 conformance fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation default is `AddTextOptions.paragraphMarginRight?: number`; paragraph snapshot/override is `RichTextParagraph.marginRight?: number | false`.
- Values use points, must be finite and within `0..4032`, and serialize as `Math.round(points * 12700)` in `a:pPr@marR`.
- New paragraphs with omitted right margin do not gain direct `marR`; explicit zero writes `marR="0"`; `marginRight: false` writes no direct value.
- Existing-shape `richText` replacement treats omitted `marginRight` or false as clear; plain `.text` continues to preserve and clone the first paragraph template.
- Right margin can coexist with direct `buChar`, `buAutoNum`, or `buBlip`; it never changes bullet-owned `marL`/`indent` and creates no bullet conflict error.
- Direct getter accepts only integer `marR` values in `0..51206400` EMU and does not resolve inherited paragraph state.
- Do not add first-line/hanging `indent`, inheritance resolution, RTL side swapping, picture-bullet writing, or unrelated refactors.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, and LibreOffice rendering pass.

---

### Task 1: Strict paragraph model and lossless `marR` codec

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `RichTextParagraph`, `NormalizedRichTextParagraph`, `renderRichTextParagraphs()`, `renderParagraphProperties()`, `updateParagraphAttribute()`, `readRichText()`, and paragraph templates.
- Produces: `RichTextParagraph.marginRight?: number | false`, strict normalization, resolved right-margin rendering, direct getter, and exact `pPr@marR` replacement/removal.

- [ ] **Step 1: Add failing strict-read tests, including list paragraphs**

Create existing paragraphs whose `marR` values are `0`, `152400`, `1`, missing, empty, `1.5`, `-1`, `51206401`, and `yes`; add four `marR="228600"` paragraphs containing direct `buNone`, `buChar`, `buAutoNum`, and `buBlip`. Assert:

```ts
expect(shape.richText.map(({ marginRight }) => marginRight)).toEqual([
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
  18,
  18,
  18,
]);
expect(pkg.mutations).toEqual(journal);
```

Use plain `.text = 'First\nSecond'` on the valid 12-point paragraph and assert both output paragraphs preserve exact `marR="152400"` plus unrelated attributes/children.

- [ ] **Step 2: Run the model suite and confirm the new assertions fail**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: failure because `marginRight` is not part of `RichTextParagraph` or the getter.

- [ ] **Step 3: Add the public field and strict normalization**

Add to `packages/model/src/text.ts`:

```ts
export interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly align?: TextAlignment;
  readonly rtl?: boolean;
  readonly marginLeft?: number | false;
  readonly marginRight?: number | false;
  readonly bullet?: ParagraphBullet;
  readonly level?: number;
  readonly spacing?: ParagraphSpacing | false;
  readonly tabStops?: readonly ParagraphTabStop[] | false;
}
```

In `rich-text.internal.ts`, allow `marginRight` in the supported key list and normalize it with a new exported function:

```ts
export function normalizeParagraphMarginRight(value: unknown, context: string): number {
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
const marginRight = candidate.marginRight === undefined
  ? undefined
  : candidate.marginRight === false
    ? false
    : normalizeParagraphMarginRight(
        candidate.marginRight,
        `Rich text paragraph ${paragraphIndex} marginRight`,
      );
```

Preserve explicit false in the normalized paragraph object; do not round the public point value before EMU conversion.

- [ ] **Step 4: Resolve paragraph defaults without introducing bullet conflicts**

Add `defaultMarginRight?: number` to `RenderRichTextOptions`. For each paragraph, resolve the local/default value independently from `resolvedBullet`:

```ts
const resolvedMarginRight = marginRight === false
  ? false
  : marginRight ?? options.defaultMarginRight;
```

Keep the existing `resolvedBullet` versus numeric `resolvedMarginLeft` conflict check unchanged. Pass `resolvedMarginRight` to `renderParagraphProperties()` for ordinary and list paragraphs alike.

- [ ] **Step 5: Patch `marR` independently after left-margin rendering**

Append `marginRight?: number | false` to `renderParagraphProperties()` and update it after `renderParagraphLeftMargin()` but before tab-stop rendering:

```ts
function renderParagraphRightMargin(
  template: string,
  marginRight: number | false | undefined,
  isNewParagraph: boolean,
): string {
  if (marginRight === undefined && isNewParagraph) return template;
  return updateParagraphAttribute(
    template,
    'marR',
    typeof marginRight === 'number' ? String(Math.round(marginRight * EMU_PER_POINT)) : undefined,
  );
}
```

`isNewParagraph` is `template === undefined`. Omitted new rendering preserves absence; existing paragraph templates clear `marR` when the public field is omitted or false. Do not inspect or branch on bullet state in this function.

- [ ] **Step 6: Read only valid direct values on every paragraph kind**

Add `readParagraphMarginRight()` that finds direct `pPr`, parses a full decimal integer, accepts only `0..51206400`, and returns `value / EMU_PER_POINT`. It must not suppress values for `buChar`, `buAutoNum`, or `buBlip`. Include `marginRight` in `readRichText()` only when defined. Re-run the model suite and update exact snapshots only where source XML already contains a valid direct `marR`.

### Task 2: Plain/rich creation defaults and lifecycle validation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 normalization/rendering, `AddTextOptions`, `validateAddTextOptions()`, `validateTextInput()`, `textParagraphXml()`, transactions, duplicate, and write/reopen.
- Produces: `AddTextOptions.paragraphMarginRight?: number`, plain/rich defaults and overrides, list coexistence, mutation-isolated failures, and native lifecycle evidence.

- [ ] **Step 1: Add failing plain/rich creation tests**

Create plain shapes for omitted, `0`, `12.5`, and `4032` across CR/LF paragraphs and empty lines. Create rich text with an outer 24-point default, a 12-point override, false suppression, a bullet that inherits 24 points, and a numbered paragraph with an 18-point override. Assert public snapshots and exact XML values:

```ts
expect(shape.richText.map(({ marginRight }) => marginRight)).toEqual([
  24,
  12,
  undefined,
  24,
  18,
]);
expect(slideXml).toContain('marR="304800"');
expect(slideXml).toContain('marR="152400"');
expect(slideXml).toContain('marR="228600"');
```

Combine left/right margins with alignment, RTL, spacing, tab stops, bullet, and numbering. Assert bullet `marL`/negative `indent` pairs remain byte-for-byte correct while `marR` is independently present.

- [ ] **Step 2: Add and normalize the outer creation option**

Add to `AddTextOptions`:

```ts
readonly paragraphMarginRight?: number;
```

Add `marginRight?: number` to `NormalizedAddTextOptions` and `NormalizedTextInput`; normalize with `normalizeParagraphMarginRight(options.paragraphMarginRight, 'Paragraph right margin')`. Pass it to every plain `textParagraphXml()` call and as `defaultMarginRight` for rich rendering.

- [ ] **Step 3: Keep plain and rich list combinations valid**

Do not add a right-margin/bullet conflict in `validateTextInput()` or rich rendering. Verify these both create successfully:

```ts
slide.addText('Bullet', { bullet: true, paragraphMarginRight: 12 });
slide.addRichText(
  [{ bullet: { kind: 'number' }, marginRight: 18, runs: [{ text: 'Numbered' }] }],
  { paragraphMarginRight: 24 },
);
```

Pass the normalized margin to `textParagraphXml()` and then `renderParagraphProperties()` without changing existing left-margin conflict validation.

- [ ] **Step 4: Cover edit, clear, duplicate, reopen, and rollback**

Replace existing paragraphs with positive, zero, false, and omitted right margins; assert false/omitted remove only direct `marR`. Duplicate before edits and prove the duplicate retains original values. Verify plain `.text` clones the first template right margin. Roll back an outer transaction after a right-margin update and assert exact bytes, mutation journal, slide identity, and shape identity are restored. Write/reopen and validate all supported values and list combinations.

- [ ] **Step 5: Add invalid-input mutation isolation**

Test outer values `null`, true, false, strings, objects, arrays, symbols, NaN, Infinity, `-0.01`, and `4032.01`. Test the same invalid paragraph values except false, which is the valid suppression form already covered above. Assert TypeError for non-number/non-finite values, RangeError for finite out-of-range values, and unchanged package bytes/journal before any shape is allocated. Include bullet and numbering options in valid boundary cases to prove no conflict was introduced.

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
- Produces: no-default-`marR` and bullet-indentation regression evidence, partial-support documentation, and Node/browser/declaration packed coverage.

- [ ] **Step 1: Extend PptxGenJS real-output conformance**

Use the existing real PptxGenJS 4.0.1 ordinary plain/rich, bullet, and numbered paragraphs. Assert every imported paragraph has `marginRight === undefined`, raw XML contains no `marR=`, and ordinary direct `marL="0"` plus list `marL`/negative `indent` pairs remain unchanged. Reopen the imported document and repeat the public assertions. Do not access PptxGenJS private fields.

- [ ] **Step 2: Update compatibility and API documentation**

Change the paragraph margin/indent matrix row to direct left and right margins supported, first-line/hanging indent still pending. Document point units, zero versus false/absence, text-box margin separation, bullet coexistence for right margin, strict direct reads, physical-side semantics under RTL, and PptxGenJS's lack of a paragraph-right public option. Add one precise changelog bullet and examples to API/package README.

- [ ] **Step 3: Extend packed Node/browser/declaration smoke**

In Node smoke, create 24/12/false/bullet rich paragraphs, edit to 6/0/omitted, and assert snapshots plus unchanged bullet left-indent behavior. In browser smoke, create and clear one right margin. In the declaration fixture, compile:

```ts
const paragraphRightMargins: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { marginRight: 12, runs: [{ text: 'Override' }] },
  { marginRight: false, runs: [{ text: 'Suppressed' }] },
  { bullet: true, marginRight: 18, runs: [{ text: 'Bullet' }] },
];
createdDocument.addSlide().addRichText(paragraphRightMargins, { paragraphMarginRight: 24 });
```

- [ ] **Step 4: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: adapter and all functional tests pass, the default performance case is the only full-run skip, and isolated performance passes.

### Task 4: Package, visual validation, final review, and delivery

**Files:**
- Review all Task 1–3 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and all earlier evidence.
- Produces: reviewed `feat: support paragraph right margin` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
paragraph_right_margin_package_dir=$(mktemp -d /tmp/pptx-paragraph-right-margin-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$paragraph_right_margin_package_dir"
node ../../scripts/smoke-npm-package.mjs "$paragraph_right_margin_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke checks all report true.

- [ ] **Step 2: Generate and validate native comparison files**

Generate `/tmp/pptx-paragraph-right-margin-native/native.pptx` with equal-width text boxes, visible right guide lines, long wrapping text, and 0/12/24/48-point paragraph right margins. Generate a baseline with the same valid `pPr@marR` values using a hand-patched copy because PptxGenJS has no public paragraph-right API. Validate both with:

```sh
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-paragraph-right-margin-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-paragraph-right-margin-baseline/baseline.pptx --profile powerpoint-2010
```

Expect zero errors and zero warnings. Decompress both files and compare every package part; after normalizing ZIP metadata, native and hand-patched baseline content must match exactly.

- [ ] **Step 3: Export and inspect every rendered page**

Use `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf` for both files, render PDFs with `pdftoppm -png`, and visually inspect every page. Confirm long text wraps earlier as right margins increase through 0/12/24/48 points, left anchors and bullet `marL`/`indent` remain unchanged, native and baseline renders match, and there is no repair output, clipping, overlap, or overflow.

- [ ] **Step 4: Review, commit, push, and verify remote state**

Run `git diff --check`, inspect the entire diff, verify omission/false/direct-zero semantics, bullet coexistence, direct strict reads, physical-side RTL semantics, and no first-line/hanging `indent` expansion. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, run `git diff --cached --check`, and commit:

```sh
git commit -m "feat: support paragraph right margin"
```

Push through SSH-over-443 and verify:

```sh
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.
