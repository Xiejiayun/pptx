# Rich Text Transparency Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, reversible rich-text main-fill transparency create/read/edit support aligned with PptxGenJS's `0..100` percentage API.

**Architecture:** Add one optional scalar to `RichTextRunStyle`, normalize it through the OOXML alpha integer, and render alpha only inside the main text `solidFill` color choice. Read transparency with an independent strict helper that requires one unambiguous direct fill/color/alpha path, while leaving the existing base-color getter and all outline/highlight/underline/glow paths unchanged.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, DrawingML color transforms, PptxGenJS 4.0.1 real-output fixtures, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public API is `RichTextRunStyle.transparency?: number`; do not add outer `AddTextOptions.transparency` in this item.
- Values are finite percentages in `0..100`; 0 is opaque and 100 is fully transparent.
- Serialize `alpha = Math.round((100 - transparency) * 1000)` and normalize the public value back to `100 - alpha / 1000`.
- Omission writes no alpha. Explicit 0 writes direct `alpha val="100000"`; explicit 100 writes `alpha val="0"`.
- Transparency affects only the main text `rPr/solidFill` color; outline, glow, highlight, underline fill, shape fill, and table-cell fill are independent.
- With no explicit run color, transparency applies to the existing default `schemeClr tx1` fill.
- Getter requires one unique direct solid fill, one valid direct sRGB/scheme color, and one strict direct alpha; it does not resolve inheritance.
- Plain `.text` preserves run templates; `shape.richText` replacement rebuilds run styles and omission clears direct alpha.
- Do not add gradient/pattern text fill, other color transforms, table-cell APIs, inheritance, or unrelated refactors.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, actual tarball smoke, CLI validation, package diff, overflow checks, and LibreOffice rendering pass.

---

### Task 1: Strict public style, alpha normalization, rendering, and getter

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `RichTextRunStyle`, `normalizeStyle()`, `renderRun()`, `renderColorChoice()`, `readStyle()`, `readIntegerAttribute()`, `SCHEME_COLORS`, and direct-child helpers.
- Produces: `RichTextRunStyle.transparency?: number`, `normalizeTextTransparency()`, main-fill alpha rendering, strict direct transparency reads, and plain-text preservation evidence.

- [ ] **Step 1: Add failing strict-read and plain-preservation tests**

Build existing runs with valid main fills containing alpha values `0`, `1`, `49445`, `75000`, and `100000`; then separately cover no alpha transform, alpha without `val`, empty/decimal/scientific/negative/`100001` values, repeated alpha, lone `alphaMod`, lone `alphaOff`, alpha plus tint, alpha with an extra attribute, alpha with an element child, color with an extra attribute, duplicate solid fills, duplicate color choices, invalid sRGB, and unknown scheme. Assert the first five values exactly and assert every malformed case is `undefined`:

```ts
const transparencies = shape.richText[0]!.runs.map(({ style }) => style?.transparency);
expect(transparencies.slice(0, 5)).toEqual([100, 99.999, 50.555, 25, 0]);
expect(transparencies.slice(5)).toEqual(Array(18).fill(undefined));
expect(pkg.mutations).toEqual(journal);
```

In a separate fixture, give the first run `<a:srgbClr val="FF0000"><a:alpha val="75000"/></a:srgbClr>` plus unrelated `rPr` XML, set plain `.text = 'First\nSecond'`, and assert both paragraphs preserve the exact alpha and unrelated style markup.

- [ ] **Step 2: Run the model suite and verify red**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new transparency assertions fail because the field/getter do not exist.

- [ ] **Step 3: Add the public field and canonical normalizer**

Add to `RichTextRunStyle`:

```ts
readonly transparency?: number;
```

Allow `transparency` in the style key list and normalize it before returning the normalized style. Keep the helper private to `rich-text.internal.ts` because no other style surface consumes it:

```ts
function normalizeTextTransparency(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > 100) {
    throw new RangeError(`${context} must be between 0 and 100 percent`);
  }
  const alpha = Math.round((100 - value) * 1000);
  return 100 - alpha / 1000;
}
```

Preserve explicit zero with an `!== undefined` object branch. Do not add false/string aliases.

- [ ] **Step 4: Render alpha only in the main text fill**

Keep `renderColorChoice()` unchanged for outline, highlight, underline, and other existing callers. Add a dedicated helper used only by `renderRun()`:

```ts
function renderMainTextColorChoice(
  color: RichTextColor,
  prefix: string,
  transparency: number | undefined,
): string {
  const tag = color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  if (transparency === undefined) return `<${prefix}${tag} val="${color.value}"/>`;
  const alpha = Math.round((100 - transparency) * 1000);
  return `<${prefix}${tag} val="${color.value}"><${prefix}alpha val="${alpha}"/></${prefix}${tag}>`;
}
```

Replace only the main `colorXml` call in `renderRun()` with this helper. `style.color ?? { kind: 'scheme', value: 'tx1' }` remains the resolved color, so transparency without explicit color is valid. Do not pass transparency to outline/glow/highlight/underline rendering.

- [ ] **Step 5: Read one strict direct alpha path**

Add `readMainTextTransparency(xml, properties)` with these exact checks:

1. `properties` has exactly one direct `solidFill`.
2. That fill has exactly one direct child, either `srgbClr` with six hex digits or a supported `schemeClr`.
3. The color has only its `val` attribute apart from namespace declarations.
4. The color has exactly one direct element child named `alpha`; `alphaMod`, `alphaOff`, tint, and mixed transforms are rejected.
5. Alpha has only `val`, has no direct element child, and the strict integer is in `0..100000`.

Return `100 - alpha / 1000`; otherwise return undefined. In `readStyle()`, add transparency independently from the existing base-color result:

```ts
const transparency = readMainTextTransparency(xml, properties);
// ...
...(transparency !== undefined ? { transparency } : {}),
```

Do not change how invalid/unmodeled base color transforms are preserved in source XML during read-only or plain-text operations.

- [ ] **Step 6: Re-run the model suite and fix only expected snapshots**

Existing fixtures with a valid main-fill alpha may now legitimately include `transparency`. Update only those exact expected snapshots; do not expose alpha from glow, outline, highlight, or underline fills.

### Task 2: Native create/edit/clear lifecycle and mutation isolation

**Files:**
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 style codec, `addRichText()`, `ShapeModel.richText`, duplicate, transactions, and write/reopen.
- Produces: creation/lifecycle/boundary evidence without any `SlideModel` API changes.

- [ ] **Step 1: Add rich creation and quantization coverage**

Create runs for omitted, 0, 25, `50.5555`, 100, scheme 40, default-color 60, and empty text 75. Combine one transparent run with font, baseline, character spacing, glow, highlight, outline, underline, strike, paragraph indent, RTL, spacing, and tab stops. Assert snapshots and exact XML:

```ts
expect(shape.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
  undefined,
  0,
  25,
  50.555,
  100,
  40,
  60,
  75,
]);
expect(slideXml).toContain('<a:alpha val="100000"/>');
expect(slideXml).toContain('<a:alpha val="75000"/>');
expect(slideXml).toContain('<a:alpha val="49445"/>');
expect(slideXml).toContain('<a:alpha val="0"/>');
```

Assert the glow alpha remains its independent opacity value and outline/highlight/underline colors gain no alpha.

- [ ] **Step 2: Cover replacement, clear, duplicate, reopen, and rollback**

Duplicate before editing. Replace runs with transparency 10, 90, 0, 100, and omitted; assert omission has no direct alpha. Mutate a returned snapshot locally and prove model isolation. Roll back an outer transaction after setting transparency and assert exact bytes, journal, slide/shape identity, and live snapshot restoration. Write/reopen and compare edited and duplicated values.

- [ ] **Step 3: Add invalid-input isolation**

Test `null`, true, false, strings, objects, arrays, symbols, NaN, Infinity, `-0.001`, and `100.001` for both `addRichText()` and `shape.richText` replacement. Assert TypeError for non-number/non-finite values, RangeError for finite out-of-range values, and unchanged bytes/journal/shape identity.

- [ ] **Step 4: Run typecheck and focused suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck and both suites pass.

### Task 3: PptxGenJS conformance, compatibility docs, and packed API

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public style and real PptxGenJS output import.
- Produces: 4.0.1 conformance evidence, split compatibility status, and Node/browser/declaration tarball coverage.

- [ ] **Step 1: Add real PptxGenJS 4.0.1 conformance**

Generate one rich text shape through public PptxGenJS APIs with explicit sRGB runs for omitted, 0, 25, `50.5555`, and 100; add scheme accent1 at 40 and a run with transparency 60 but no explicit color. Import through `importPptxGenJS()` and assert:

```ts
expect(runs.map(({ style }) => style?.transparency)).toEqual([
  undefined,
  undefined,
  25,
  50.555,
  100,
  40,
  60,
]);
```

Assert the no-color run imports PptxGenJS's direct black color, scheme remains scheme, raw alpha values match, and write/reopen preserves values. Do not inspect PptxGenJS private fields.

- [ ] **Step 2: Update compatibility and API documentation**

Split the matrix row into:

```md
| rich run `transparency` 0–100% | `RichTextRunStyle.transparency` | 已支持 |
| table-cell fit / `textDirection` | 尚无完整公开 API | 部分支持，后续逐项补齐 |
```

Document percentage direction, 0.001% quantization, explicit zero versus absence, default tx1 behavior, main-fill-only scope, strict direct read, and PptxGenJS zero/default-color differences. Add a precise changelog bullet plus API/package README examples.

- [ ] **Step 3: Extend packed Node/browser/declaration smoke**

In Node smoke, create 25/fractional/100/default-color runs, edit to 0/75/omitted, and assert canonical snapshots. In browser smoke, create 50 then clear by omission. Compile this declaration fixture:

```ts
const transparentParagraphs: readonly RichTextParagraph[] = [{
  runs: [
    { text: 'Opaque', style: { transparency: 0 } },
    { text: 'Quarter', style: { color: { kind: 'srgb', value: 'FF0000' }, transparency: 25 } },
    { text: 'Theme', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 100 } },
  ],
}];
createdDocument.addSlide().addRichText(transparentParagraphs);
```

- [ ] **Step 4: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: all functional tests pass, only default performance is skipped in the full run, and isolated performance passes.

### Task 4: Package, visual validation, final review, and delivery

**Files:**
- Review all Task 1–3 files; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and all earlier evidence.
- Produces: reviewed `feat: support rich text transparency` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
rich_transparency_package_dir=$(mktemp -d /tmp/pptx-rich-transparency-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$rich_transparency_package_dir"
node ../../scripts/smoke-npm-package.mjs "$rich_transparency_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected: Node API, browser API, declarations, and CLI smoke all report true.

- [ ] **Step 2: Generate native and hand-patched baseline PPTX files**

Generate `/tmp/pptx-rich-transparency-native/native.pptx` with large identical red text rows at 0/25/50/75/100 transparency plus a theme-color row. Generate `/tmp/pptx-rich-transparency-baseline/baseline.pptx` from the same omitted-alpha document and patch only the main color choices with alpha `100000/75000/50000/25000/0` and the theme alpha. Validate both and compare:

```sh
command -v pptx-inspect
pptx-inspect --json doctor
pptx-inspect --json package validate /tmp/pptx-rich-transparency-native/native.pptx --profile powerpoint-2010
pptx-inspect --json package validate /tmp/pptx-rich-transparency-baseline/baseline.pptx --profile powerpoint-2010
pptx-inspect --json package diff /tmp/pptx-rich-transparency-native/native.pptx /tmp/pptx-rich-transparency-baseline/baseline.pptx
node packages/pptx/dist/cli.js --json doctor
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-rich-transparency-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-rich-transparency-baseline/baseline.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff /tmp/pptx-rich-transparency-native/native.pptx /tmp/pptx-rich-transparency-baseline/baseline.pptx
```

Expected: zero errors/warnings and empty added/removed/changed diff arrays.

- [ ] **Step 3: Render, inspect, and run overflow checks**

Export both through `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf`, rasterize every PDF page with `pdftoppm -png`, and inspect at full size. Confirm red text fades progressively from 0 through 100, the theme row renders, fully transparent text leaves no unexpected glyph artifact, native/baseline pages match, and there is no repair output or clipping.

Run both with the presentation overflow helper:

```sh
python /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-rich-transparency-native/native.pptx
python /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py /tmp/pptx-rich-transparency-baseline/baseline.pptx
```

Expected: both checks pass with no overflow.

- [ ] **Step 4: Review, commit, push, and verify remote state**

Run `git diff --check`, inspect the full diff, and verify main-fill-only scope, percentage/alpha direction, explicit zero, fractional canonicalization, strict direct reads, invalid-input isolation, and no outer/table-cell/color-type expansion. Confirm only intended files plus `.pnpm-store/` are listed. Stage explicit files, run cached diff checks, commit, and push through the verified SSH-over-443 channel:

```sh
git commit -m "feat: support rich text transparency"
git rev-list --left-right --count origin/main...HEAD
```

Expected: remote count is `0 0`.
