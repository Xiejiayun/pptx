# Rich Text Language Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native plain/rich create, direct-run read, and rich replacement support for PptxGenJS 4.0.1 `lang` semantics.

**Architecture:** Extend the existing text creation and `RichTextRunStyle` vertical slices. Normalize outer/run language strings before mutation, pass one optional creation default through plain/rich renderers, resolve run language with `style.lang ?? defaultLanguage ?? 'en-US'`, and conservatively expose only direct non-empty `rPr@lang`. Preserve the current whole-run replacement and paragraph-template boundaries without adding a general locale or inheritance engine.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation API is `AddTextOptions.lang?: string`; live run API is `RichTextRunStyle.lang?: string` through `ShapeModel.richText`.
- Outer language applies to plain runs, rich runs without a local override, and creation-time `endParaRPr`; run language wins only for that run.
- Omitted language writes `lang="en-US"` without `altLang`; explicit outer/run language writes `altLang="en-US"` on runs only.
- Accept non-empty strings without trimming or case normalization; reject XML 1.0 invalid controls and escape attribute metacharacters.
- Read only non-empty direct `rPr@lang`; never infer language from `altLang`, `endParaRPr`, paragraphs, layouts, masters, or presentation defaults.
- Getter and non-rich-text mutations preserve source `lang`/`altLang`; `shape.richText` remains whole-run replacement and preserves the existing first `endParaRPr` template.
- PptxGenJS is used only by the adapter conformance test and never becomes a non-adapter dependency.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice open/export checks pass.

---

### Task 1: Run API, normalization, serialization, and direct read

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: existing `RichTextRunStyle`, `normalizeStyle()`, `renderRichTextParagraphs()`, `renderRun()`, `readStyle()`, `escapeXmlAttribute()`, and rich replacement template behavior.
- Produces: `RichTextRunStyle.lang?: string`, exported `normalizeTextLanguage(value, context)`, optional `RenderRichTextOptions.defaultLanguage`, PptxGenJS-compatible run attributes, and direct language snapshots.

- [ ] **Step 1: Add a failing direct-read and preservation test**

Add a model fixture test with runs containing `lang="fr-CA" altLang="it-IT"`, `lang="x-private"`, `lang=""`, and no `lang`. Assert only non-empty direct values are returned, `altLang` does not affect the result, another style on the empty-language run remains readable, and getter access does not mutate the package.

```ts
expect(shape.richText[0]!.runs.map(({ style }) => style?.lang)).toEqual([
  'fr-CA',
  'x-private',
  undefined,
  undefined,
]);
expect(pkg.mutations).toEqual(journal);
```

Then set `shape.text = 'First\nSecond'` and assert the first run template's exact `lang="fr-CA" altLang="it-IT"` appears twice and the original first `endParaRPr` is preserved.

- [ ] **Step 2: Run the model test and confirm it fails before implementation**

Run:

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: failure because `RichTextRunStyle.lang` is absent and the getter does not expose `rPr@lang`.

- [ ] **Step 3: Add the public property and shared normalizer**

In `packages/model/src/text.ts`, add:

```ts
readonly lang?: string;
```

In `packages/model/src/rich-text.internal.ts`, allow `lang` in the strict style key list and normalize it before returning the immutable style snapshot:

```ts
export function normalizeTextLanguage(value: unknown, context: string): string {
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`${context} must be a non-empty string`);
  }
  if (containsInvalidXmlCharacter(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value;
}
```

Normalize and copy with these exact expressions:

```ts
const language = candidate.lang === undefined
  ? undefined
  : normalizeTextLanguage(candidate.lang, `Rich text run ${paragraphIndex},${runIndex} lang`);
...(language !== undefined ? { lang: language } : {}),
```

- [ ] **Step 4: Thread the optional default through rich rendering**

Extend `RenderRichTextOptions` with `defaultLanguage?: string`, pass it into every `renderRun()`, compute this exact first attribute fragment, and replace the current hard-coded `'lang="en-US"'` list entry with `languageAttributes`:

```ts
const language = style.lang ?? defaultLanguage ?? 'en-US';
const explicitLanguage = style.lang !== undefined || defaultLanguage !== undefined;
const languageAttributes = `lang="${escapeXmlAttribute(language)}"${
  explicitLanguage ? ' altLang="en-US"' : ''
}`;
```

Build the default `endParaRPr` from `options.defaultLanguage ?? 'en-US'`, escaped as an attribute, without `altLang`. `replaceRichText()` continues passing no default and continues supplying the exact existing first `endParaRPr` when present.

- [ ] **Step 5: Read the direct attribute and update intentional exact snapshots**

In `readStyle()`, add:

```ts
const language = xml.attribute(properties, 'lang')?.value;
...(language ? { lang: language } : {}),
```

The spread expression is inserted into the existing `RichTextRunStyle` object literal; it is not a new wrapper object.

Update existing exact run-style assertions whose fixture has direct `lang="en-US"` so the new public field is expected. Re-run the model test and expect it to pass.

### Task 2: Outer creation defaults and native lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `normalizeTextLanguage()`, `RenderRichTextOptions.defaultLanguage`, `AddTextOptions`, plain `textParagraphXml()`, rich `addRichText()`, transactions, duplication, and write/reopen.
- Produces: `AddTextOptions.lang?: string`, plain/rich default inheritance, creation-time `endParaRPr` language, escaping, and lifecycle evidence.

- [ ] **Step 1: Add failing plain/rich creation coverage**

Create a default plain shape, a multi-paragraph plain shape with `lang: 'fr-CA'`, and a rich shape with outer `fr-CA` plus local `de-DE`, explicit `en-US`, and a metacharacter value such as `'x-private&"quoted'`. Assert:

```ts
expect(rich.richText[0]!.runs.map(({ style }) => style?.lang)).toEqual([
  'fr-CA',
  'de-DE',
  'en-US',
  'x-private&"quoted',
]);
```

Inspect XML for default `lang="en-US"` without `altLang`, inherited/local attributes with `altLang="en-US"`, `x-private&amp;&quot;quoted`, and `endParaRPr lang="fr-CA"` on every outer-language paragraph including an empty one.

- [ ] **Step 2: Thread normalized outer language through plain and rich creation**

Add `readonly lang?: string` to `AddTextOptions`, `language?: string` to both normalized option interfaces, and normalize with:

```ts
const language = options.lang === undefined
  ? undefined
  : normalizeTextLanguage(options.lang, 'Text language');
```

Return the normalized field, pass it from `validateTextInput()` into every `textParagraphXml()` call, and pass it from `addRichText()` as `defaultLanguage`.

- [ ] **Step 3: Serialize plain run and paragraph language**

Extend `textParagraphXml()` and `defaultTextRunXml()` with `language?: string`. Use escaped `language ?? 'en-US'` for run and `endParaRPr`; add `altLang="en-US"` only when `language !== undefined` on a run. Empty paragraphs contain no run but must still use the outer language on `endParaRPr`.

- [ ] **Step 4: Cover edit/default reset, duplicate, reopen, and rollback**

Duplicate the rich-language slide, replace the original shape's rich text with one local `ja-JP` run and one omitted-language run, and assert the omitted value reads back as direct `en-US`. Verify the duplicate retains the creation languages, write/reopen retains both results, a transaction rollback restores the previous bytes and object identity, and `validatePackage()` has no errors.

- [ ] **Step 5: Add invalid outer/run mutation-isolation cases**

Test outer `lang` values `null`, booleans, numbers, objects, arrays, empty string, and `bad\u0000lang` against both `addText()` and `addRichText()`. Add the same invalid values under `style.lang` to the existing invalid rich-text matrix and exercise the setter. After all failures, assert slide bytes and mutation journal are unchanged.

- [ ] **Step 6: Run typecheck and focused model/SDK tests**

Run:

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck exits 0 and both suites pass.

### Task 3: PptxGenJS conformance, documentation, and release surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public fields and the existing PptxGenJS import/write/reopen fixture.
- Produces: actual 4.0.1 outer/run/default evidence, migration documentation, and packed Node/browser/type smoke coverage.

- [ ] **Step 1: Add a PptxGenJS-generated language shape**

Append one rich shape with outer `lang: 'fr-CA'` and runs with missing language, local `de-DE`, explicit `en-US`, and empty string. Assert imported and reopened languages are `['fr-CA', 'de-DE', 'en-US', 'fr-CA']`, and inspect the raw slide XML to confirm `endParaRPr lang="fr-CA"` plus the expected `altLang` attributes.

```ts
generatedSlide.addText([
  { text: 'Inherited', options: {} },
  { text: ' German', options: { lang: 'de-DE' } },
  { text: ' Explicit default', options: { lang: 'en-US' } },
  { text: ' Empty inherits', options: { lang: '' } },
], { x: 9, y: 10, w: 3, h: 1, lang: 'fr-CA', objectName: 'Language outer' });
```

Update the existing exact PptxGenJS run snapshots to include their newly exposed direct `lang: 'en-US'` fields.

- [ ] **Step 2: Update public compatibility documentation**

Add a completed baseline row mapping PptxGenJS outer/run `lang` to `AddTextOptions.lang` / `RichTextRunStyle.lang`; remove run language from the partial-support row. Document outer inheritance, local override, direct getter behavior, `altLang` output, default `en-US`, strict empty/type rejection, and XML escaping. Add one Unreleased changelog bullet and update both API/package README examples or capability summaries.

- [ ] **Step 3: Extend packed-package smoke**

In the Node smoke, create a rich shape with outer `fr-CA` and local `de-DE`, then assert both resolved values. In the browser smoke, add a local `ja-JP` run and assert it. In the TypeScript smoke, prove both declaration entry points compile:

```ts
const languageStyle: RichTextRunStyle = { lang: 'de-DE' };
createdDocument.addSlide().addRichText(
  [{ runs: [{ text: 'Typed language', style: languageStyle }] }],
  { lang: 'fr-CA' },
);
```

- [ ] **Step 4: Run adapter and full repository tests**

Run the adapter suite, then `pnpm test` and `pnpm test:performance`. Expected: all functional suites pass, the default performance case is skipped only in the full run, and the explicit performance run passes.

### Task 4: Package, compatibility, final review, commit, and push

**Files:**
- Review all files listed in Tasks 1–3; never stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation, tests, and docs.
- Produces: reviewed `feat: support rich text language` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual npm tarball**

From `packages/pptx`, run the pinned Node/browser tsup builds and declaration builder, then pack into a validated temporary directory:

```sh
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
package_dir=$(mktemp -d /tmp/pptx-rich-language-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$package_dir"
node ../../scripts/smoke-npm-package.mjs "$package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser conditional export, declarations, and packed CLI doctor all pass.

- [ ] **Step 2: Validate native and PptxGenJS comparison decks**

Generate `/tmp/pptx-rich-language-native/native.pptx` and `/tmp/pptx-rich-language-pptxgenjs/pptxgenjs.pptx` with English, French-Canadian, German, Japanese, and Chinese runs. Validate both with:

```sh
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-rich-language-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-rich-language-pptxgenjs/pptxgenjs.pptx --profile powerpoint-2010
```

Expected: zero errors and zero warnings. Read `ppt/slides/slide1.xml` from both ZIPs and confirm the valid default/outer/override/end-paragraph mappings from the design.

- [ ] **Step 3: LibreOffice open/export and visual sanity check**

Open/export both comparison decks headlessly, render every PDF page, and inspect the rendered PNGs:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir /tmp/pptx-rich-language-native /tmp/pptx-rich-language-native/native.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf --outdir /tmp/pptx-rich-language-pptxgenjs /tmp/pptx-rich-language-pptxgenjs/pptxgenjs.pptx
pdftoppm -png /tmp/pptx-rich-language-native/native.pdf /tmp/pptx-rich-language-native/render
pdftoppm -png /tmp/pptx-rich-language-pptxgenjs/pptxgenjs.pdf /tmp/pptx-rich-language-pptxgenjs/render
```

Expected: no repair/error output, all scripts remain visible and uncorrupted, and language metadata causes no clipping or layout regression.

- [ ] **Step 4: Final review**

Run `git diff --check`, inspect the complete diff, check for duplicated normalization or accidental inheritance parsing, verify all public strings are escaped exactly once, and ensure `git status --short` lists only the intended implementation files plus untracked `.pnpm-store/`. Re-run focused tests after any review fix.

- [ ] **Step 5: Commit and push**

Stage only the intended implementation files, commit with:

```sh
git commit -m "feat: support rich text language"
```

Push `main` using the repository SSH-over-443 command and verify `origin/main...HEAD` reports `0 0`.
