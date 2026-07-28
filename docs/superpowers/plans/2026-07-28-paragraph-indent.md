# Paragraph Indent Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native signed ordinary-paragraph first-line/hanging indent create/read/edit/clear support without changing bullet/numbering indentation.

**Architecture:** Extend the existing paragraph-property pipeline in `rich-text.internal.ts` with one signed point-to-EMU codec for direct `pPr@indent`. Resolve outer defaults and local overrides before rendering, reject numeric ordinary indent when the resolved paragraph has an active bullet, and patch ordinary indent only after the existing bullet and margin stages. Strict reads suppress active-list values so bullet-owned negative indent remains represented only by `ParagraphBullet.indent`.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, PptxGenJS 4.0.1 conformance fixtures, tsup, npm tarball smoke, `pptx-inspect`, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation default is `AddTextOptions.paragraphIndent?: number`; paragraph snapshot/override is `RichTextParagraph.indent?: number | false`.
- Values use points, must be finite and within `-4032..4032`, and serialize as `Math.round(points * 12700)` in `a:pPr@indent`.
- Positive values are first-line indent; negative values are hanging indent; direct zero remains distinguishable from false/absence.
- New non-list paragraphs with omitted indent retain canonical direct `indent="0"`; explicit `indent: false` removes it.
- Existing-shape `richText` replacement treats omitted `indent` or false as clear; plain `.text` preserves and clones the first paragraph template.
- Direct `buChar`, `buAutoNum`, or `buBlip` owns `indent`; the ordinary getter suppresses it and rendering never clears an active bullet's value.
- Numeric ordinary indent and a resolved active bullet conflict. `indent: false` may suppress an outer ordinary default, and `bullet: false` may suppress an outer bullet default.
- Direct getter accepts only signed integer `indent` values in `-51206400..51206400` EMU and does not resolve inherited paragraph state.
- Do not add inheritance resolution, RTL sign changes, picture-bullet writing, alternative first-line/hanging aliases, or unrelated refactors.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Strict paragraph model and lossless signed `indent` codec

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `RichTextParagraph`, `NormalizedRichTextParagraph`, `renderRichTextParagraphs()`, `renderParagraphProperties()`, `renderParagraphBullet()`, `updateParagraphAttribute()`, `readRichText()`, direct bullet parsing, and paragraph templates.
- Produces: `RichTextParagraph.indent?: number | false`, exported signed normalization, resolved indent rendering, direct getter, exact `pPr@indent` replacement/removal, and bullet ownership isolation.

- [ ] **Step 1: Add failing strict-read and preservation tests**

Create existing paragraphs whose direct `indent` values are `0`, `152400`, `-228600`, `1`, missing, empty, `1.5`, `51206401`, `-51206401`, and `yes`. Add four `indent="228600"` paragraphs containing direct `buNone`, `buChar`, `buAutoNum`, and `buBlip`. Assert:

```ts
expect(shape.richText.map(({ indent }) => indent)).toEqual([
  0,
  12,
  -18,
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

Use plain `.text = 'First\nSecond'` on the valid `-18` paragraph and assert both output paragraphs preserve exact `indent="-228600"` plus unrelated `marL`, `marR`, attributes, and children.

- [ ] **Step 2: Run the model suite and confirm the new assertions fail**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: failure because `indent` is not part of `RichTextParagraph` or the ordinary getter.

- [ ] **Step 3: Add the public field and strict signed normalization**

Add to `packages/model/src/text.ts`:

```ts
export interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly align?: TextAlignment;
  readonly rtl?: boolean;
  readonly marginLeft?: number | false;
  readonly marginRight?: number | false;
  readonly indent?: number | false;
  readonly bullet?: ParagraphBullet;
  readonly level?: number;
  readonly spacing?: ParagraphSpacing | false;
  readonly tabStops?: readonly ParagraphTabStop[] | false;
}
```

In `rich-text.internal.ts`, allow `indent` in the supported key list and normalize it with:

```ts
export function normalizeParagraphIndent(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < -4032 || value > 4032) {
    throw new RangeError(`${context} must be between -4032 and 4032 points`);
  }
  return value;
}
```

Normalize the rich paragraph field with the exact three-state branch:

```ts
const indent = candidate.indent === undefined
  ? undefined
  : candidate.indent === false
    ? false
    : normalizeParagraphIndent(candidate.indent, `Rich text paragraph ${paragraphIndex} indent`);
```

Preserve explicit false in `NormalizedRichTextParagraph`; do not round the public point value before EMU serialization.

- [ ] **Step 4: Resolve defaults and reject the resolved bullet conflict**

Add `defaultIndent?: number` to `RenderRichTextOptions`. Resolve each local/default value after resolving bullet:

```ts
const resolvedIndent = indent === false ? false : indent ?? options.defaultIndent;
if (resolvedBullet && typeof resolvedIndent === 'number') {
  throw new TypeError('Paragraph indent cannot be combined with an active bullet');
}
```

This check intentionally permits an outer bullet plus `bullet: false`/numeric indent and an outer indent plus `indent: false`/local bullet. Keep the independent left-margin/bullet conflict check and pass `resolvedIndent` to `renderParagraphProperties()`.

- [ ] **Step 5: Patch ordinary indent after bullet and margin rendering**

Append `indent?: number | false` to `renderParagraphProperties()`. After bullet, left-margin, and right-margin processing but before tab stops, call:

```ts
function renderParagraphIndent(
  template: string,
  indent: number | false | undefined,
  hasActiveBullet: boolean,
  isNewParagraph: boolean,
): string {
  if (hasActiveBullet) return template;
  if (indent === undefined && isNewParagraph) return template;
  return updateParagraphAttribute(
    template,
    'indent',
    typeof indent === 'number' ? String(Math.round(indent * EMU_PER_POINT)) : undefined,
  );
}
```

`isNewParagraph` is `template === undefined`. Omitted new ordinary rendering preserves canonical `indent="0"`; false removes it. Existing templates clear direct indent when the field is omitted/false. Active bullets retain the exact value written or preserved by `renderParagraphBullet()`.

- [ ] **Step 6: Read only valid direct ordinary values**

Add `readParagraphIndent()` that finds direct `pPr`, returns undefined if any direct active `buChar`, `buAutoNum`, or `buBlip` exists, parses a full signed decimal integer, accepts only `-51206400..51206400`, and returns `value / EMU_PER_POINT`. `buNone` remains readable. Include `indent` in `readRichText()` only when defined. Re-run the model suite and update exact snapshots only where source XML already contains a valid direct ordinary indent.

### Task 2: Plain/rich creation, conflict handling, and lifecycle validation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 normalization/rendering, `AddTextOptions`, `validateAddTextOptions()`, `validateTextInput()`, `textParagraphXml()`, transactions, duplicate, and write/reopen.
- Produces: `AddTextOptions.paragraphIndent?: number`, plain/rich defaults and overrides, pre-mutation conflict failures, and native lifecycle evidence.

- [ ] **Step 1: Add failing plain/rich creation tests**

Create plain shapes for omitted, `0`, `12.5`, `-12.5`, `4032`, and `-4032` across CR/LF paragraphs and empty lines. Create rich text with outer 24-point default, positive/negative/zero overrides, false suppression, an outer-bullet paragraph with `bullet: false`/`indent: -18`, and a local bullet with `indent: false`. Assert snapshots and XML:

```ts
expect(shape.richText.map(({ indent }) => indent)).toEqual([
  24,
  12,
  -18,
  0,
  undefined,
  -18,
  undefined,
]);
expect(slideXml).toContain('indent="304800"');
expect(slideXml).toContain('indent="152400"');
expect(slideXml).toContain('indent="-228600"');
```

Combine ordinary indent with left/right margins, alignment, RTL, spacing, and tab stops. Assert local bullet/numbering paragraphs retain their existing `marL`/negative `indent` pairs while the ordinary getter suppresses those list values.

- [ ] **Step 2: Add and normalize the outer creation option**

Add to `AddTextOptions`:

```ts
readonly paragraphIndent?: number;
```

Import `normalizeParagraphIndent`; add `indent?: number` to `NormalizedAddTextOptions` and `NormalizedTextInput`; normalize with `normalizeParagraphIndent(options.paragraphIndent, 'Paragraph indent')`. Pass it to every plain `textParagraphXml()` call and as `defaultIndent` for rich rendering.

- [ ] **Step 3: Enforce plain and rich resolved conflicts before mutation**

In `validateTextInput()`, reject the plain-text resolved pair:

```ts
if (defaults.bullet && defaults.indent !== undefined) {
  throw new TypeError('Paragraph indent cannot be combined with an active bullet');
}
```

Do not reject outer bullet plus outer indent in `validateAddTextOptions()`, because individual rich paragraphs may suppress either value. Rely on Task 1's per-paragraph resolved check. Verify these succeed:

```ts
slide.addRichText(
  [{ bullet: false, indent: -18, runs: [{ text: 'Ordinary hanging' }] }],
  { bullet: true },
);
slide.addRichText(
  [{ bullet: true, indent: false, runs: [{ text: 'List-owned indent' }] }],
  { paragraphIndent: 24 },
);
```

Verify plain `bullet: true` plus `paragraphIndent: 0` and each unresolved rich numeric-indent/active-bullet combination throw before shape allocation.

- [ ] **Step 4: Cover edit, clear, duplicate, reopen, and rollback**

Replace existing ordinary paragraphs with positive, negative, zero, false, and omitted values; assert false/omitted remove only direct `indent`. Change a bullet paragraph to ordinary with omitted indent and assert recognized bullet `marL` is zeroed while `indent` is absent. Duplicate before edits and prove the duplicate retains original values. Verify plain `.text` clones the first template's signed indent. Roll back an outer transaction after an indent update and assert exact bytes, mutation journal, slide identity, and shape identity are restored. Write/reopen and validate all ordinary/list cases.

- [ ] **Step 5: Add invalid-input mutation isolation**

Test outer values `null`, true, false, strings, objects, arrays, symbols, NaN, and Infinity. Test the same invalid paragraph values except false, which is the valid suppression form. Test signed ranges `-4032.01` and `4032.01`. Assert TypeError for non-number/non-finite values, RangeError for finite out-of-range values, TypeError for resolved list conflicts, and unchanged package bytes/journal before any shape is allocated.

- [ ] **Step 6: Run typecheck and focused suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck exits 0 and both suites pass.

### Task 3: PptxGenJS conformance, documentation, and packed API

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public API and existing real PptxGenJS import/write/reopen fixtures.
- Produces: canonical ordinary-zero/list-negative regression evidence, updated direct-indent compatibility status, and Node/browser/declaration packed coverage.

- [ ] **Step 1: Extend real PptxGenJS 4.0.1 conformance**

Use the existing ordinary plain/rich, character-bullet, and numbered-paragraph fixture. Assert imported ordinary paragraphs expose `indent: 0`, while active bullet/numbering paragraphs expose `indent === undefined` and continue exposing `bullet.indent`. Assert raw XML retains ordinary `indent="0"` and list negative `indent`/positive `marL` pairs. Reopen the imported document and repeat public assertions. Do not access PptxGenJS private fields.

- [ ] **Step 2: Update compatibility and API documentation**

Change the paragraph margin/indent matrix row to direct left/right margins plus signed first-line/hanging indent supported. Document `paragraphIndent`/`indent`, point units, `-4032..4032`, sign meaning, zero versus false/absence, canonical new zero, text-box margin separation, coexistence with ordinary margins, resolved bullet conflict, strict direct reads, RTL sign stability, and PptxGenJS's lack of an ordinary indent public option. Add one precise changelog bullet and examples to API/package README.

- [ ] **Step 3: Extend packed Node/browser/declaration smoke**

In Node smoke, create outer 24, local `-18`, false, and bullet-with-false paragraphs; edit to `6`, `-6`, `0`, false, and omitted; assert snapshots and unchanged bullet indentation. In browser smoke, create and clear one signed indent. In the declaration fixture, compile:

```ts
const paragraphIndents: readonly RichTextParagraph[] = [
  { runs: [{ text: 'Default' }] },
  { indent: 18, runs: [{ text: 'First-line' }] },
  { indent: -18, runs: [{ text: 'Hanging' }] },
  { indent: false, runs: [{ text: 'Suppressed' }] },
  { bullet: true, indent: false, runs: [{ text: 'Bullet' }] },
];
createdDocument.addSlide().addRichText(paragraphIndents, { paragraphIndent: 24 });
```

- [ ] **Step 4: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: adapter and all functional tests pass, the default performance case is the only full-run skip, and isolated performance passes.

### Task 4: Package, native visual validation, final review, and delivery

**Files:**
- Review all Task 1–3 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and all earlier evidence.
- Produces: reviewed `feat: support paragraph indent` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
paragraph_indent_package_dir=$(mktemp -d /tmp/pptx-paragraph-indent-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$paragraph_indent_package_dir"
node ../../scripts/smoke-npm-package.mjs "$paragraph_indent_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke checks all report true.

- [ ] **Step 2: Generate native and hand-patched comparison files**

Generate `/tmp/pptx-paragraph-indent-native/native.pptx` with equal-width text boxes, visible left guide lines, long wrapping text, a common 48-point left margin, and ordinary `-24/-12/0/12/24` point indents. Generate `/tmp/pptx-paragraph-indent-baseline/baseline.pptx` from the same omitted-indent base and patch only direct `pPr@indent` to the expected signed EMU values. Include separate bullet and numbering rows to prove their `marL`/negative `indent` pairs are unchanged and suppressed by the ordinary getter.

Verify the offline inspector, validate both packages, and compare them:

```sh
command -v pptx-inspect
pptx-inspect --json doctor
pptx-inspect --json package validate /tmp/pptx-paragraph-indent-native/native.pptx --profile powerpoint-2010
pptx-inspect --json package validate /tmp/pptx-paragraph-indent-baseline/baseline.pptx --profile powerpoint-2010
pptx-inspect --json package diff /tmp/pptx-paragraph-indent-native/native.pptx /tmp/pptx-paragraph-indent-baseline/baseline.pptx
```

Expected: doctor succeeds, both packages report zero errors/warnings, and package diff is empty. Also validate through `node packages/pptx/dist/cli.js --json package validate ...` to cover the just-built local CLI.

- [ ] **Step 3: Render, inspect, and run overflow checks**

Use `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf` for both files and `pdftoppm -png` for page images. Inspect every rendered page at full size. Confirm negative values move the first line left of later lines, positive values move it right, zero aligns them, `marL` stays fixed, list rows remain unchanged, native/baseline renders match, and there is no repair output, clipping, or unintended overlap.

Run the presentation overflow helper on both files:

```sh
python /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-paragraph-indent-native/native.pptx
python /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-paragraph-indent-baseline/baseline.pptx
```

Expected: both checks pass with no overflow.

- [ ] **Step 4: Review, commit, push, and verify remote state**

Run `git diff --check`, inspect the entire diff, and verify signed range/rounding, canonical new zero, setter omission/false clearing, active-bullet ownership, resolved conflict handling, strict direct reads, RTL sign stability, and no inheritance/picture-bullet expansion. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, run `git diff --cached --check`, and commit:

```sh
git commit -m "feat: support paragraph indent"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected: push succeeds and remote count is `0 0`.
