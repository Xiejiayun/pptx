# Paragraph RTL Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native plain/rich paragraph RTL create/read/edit/clear support with PptxGenJS 4.0.1 valid-output conformance.

**Architecture:** Extend the existing paragraph-property pipeline in `rich-text.internal.ts`. Normalize strict booleans before mutation, resolve `paragraph.rtl ?? defaultRtl`, patch only direct `pPr@rtl`, and read only recognized boolean lexical forms. Reuse the existing paragraph template, transaction, duplicate, adapter, package, and release-smoke paths without touching `bodyPr@rtlCol` or presentation RTL.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation default is `AddTextOptions.rtlMode?: boolean`; paragraph snapshot/override is `RichTextParagraph.rtl?: boolean`.
- Resolve paragraph value before rendering; true writes `rtl="1"`, false writes `rtl="0"`, undefined removes or omits the direct attribute.
- Read only direct `pPr@rtl`: `1/true/on` are true, `0/false/off` are false, all other values are undefined and preserved during unrelated edits.
- Never read or modify `bodyPr@rtlCol` or `p:presentation@rtl`; never infer RTL from text, language, or alignment.
- `shape.richText` remains whole-public-paragraph replacement; plain `.text` preserves paragraph templates.
- PptxGenJS run-level duplicate-`pPr` output is conformance evidence only and must not become a supported run API.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.

---

### Task 1: Paragraph API, strict read, and lossless rendering

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `RichTextParagraph`, `NormalizedRichTextParagraph`, `renderParagraphProperties()`, `updateParagraphAttribute()`, `readRichText()`, and paragraph templates.
- Produces: `RichTextParagraph.rtl?: boolean`, exported `normalizeParagraphRtl(value, context)`, `RenderRichTextOptions.defaultRtl`, strict direct getter, and exact `pPr@rtl` updates.

- [ ] **Step 1: Add a failing strict-read and preservation test**

Create paragraphs with `rtl="1"`, `true`, `on`, `0`, `false`, `off`, missing, empty, and `rtl="yes"`. Assert snapshots are:

```ts
expect(shape.richText.map(({ rtl }) => rtl)).toEqual([
  true, true, true, false, false, false, undefined, undefined, undefined,
]);
expect(pkg.mutations).toEqual(journal);
```

Set plain `.text` to multiple paragraphs and assert the first template's exact `rtl="1"` plus unrelated `pPr` content is copied without mutation loss.

- [ ] **Step 2: Run the model suite and confirm the new test fails**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: failure because `RichTextParagraph.rtl` is absent.

- [ ] **Step 3: Add public and normalized fields with strict validation**

Add `readonly rtl?: boolean` to `RichTextParagraph`, add `rtl?: boolean` to `NormalizedRichTextParagraph`, allow `rtl` in the paragraph key list, and normalize with:

```ts
export function normalizeParagraphRtl(value: unknown, context: string): boolean {
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}
```

Copy with `rtl !== undefined` so false survives.

- [ ] **Step 4: Render resolved direction through the existing property pipeline**

Add `defaultRtl?: boolean` to `RenderRichTextOptions`; destructure paragraph `rtl`; call `renderParagraphProperties()` with `rtl ?? options.defaultRtl`. Extend that function with `rtl?: boolean` and patch after alignment:

```ts
const directed = updateParagraphAttribute(
  aligned,
  'rtl',
  rtl === undefined ? undefined : rtl ? '1' : '0',
);
```

Pass `directed` into the existing level/spacing/bullet/tab-stop sequence. No new XML parser or codec file is needed.

- [ ] **Step 5: Read recognized direct boolean tokens**

Add `readParagraphRtl()` that locates direct `pPr`, reads `rtl`, returns true for `1/true/on`, false for `0/false/off`, and undefined otherwise. Include the field in `readRichText()` only when defined. Re-run the model suite and expect it to pass.

### Task 2: Outer creation defaults and native lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 normalization/rendering, `AddTextOptions`, plain paragraph creation, rich defaults, transactions, duplicate, and write/reopen.
- Produces: `AddTextOptions.rtlMode?: boolean`, mixed paragraph overrides, explicit false, clearing, and mutation-isolation evidence.

- [ ] **Step 1: Add failing plain/rich creation coverage**

Create omitted, true, and false plain shapes including empty paragraphs. Create rich paragraphs under outer true with omitted, true, and false local values. Assert public snapshots and XML contain `rtl="1"`, `rtl="0"`, and no direct attribute for omission without a default. Combine RTL with align, bullet, spacing, level, tab stops, and `lang` to prove property composition.

- [ ] **Step 2: Normalize and thread the outer default**

Add `readonly rtlMode?: boolean` to `AddTextOptions`, `rtl?: boolean` to normalized option/input interfaces, and normalize with `normalizeParagraphRtl(options.rtlMode, 'Text RTL mode')` only when provided. Pass it to every plain `textParagraphXml()` call and as `defaultRtl` for rich rendering.

- [ ] **Step 3: Extend plain paragraph rendering**

Add an `rtl?: boolean` parameter to `textParagraphXml()` and pass it to `renderParagraphProperties()`. Do not change `bodyPr rtlCol="0"`, language, end-paragraph properties, or run rendering.

- [ ] **Step 4: Cover edit/clear/duplicate/reopen/rollback**

Duplicate the slide, replace original rich paragraphs with true, false, and omitted RTL, assert omitted clears direct state, and verify the duplicate retains original defaults. Check rollback restores bytes/journal/identity, write/reopen retains all direct states, and `validatePackage()` returns no errors.

- [ ] **Step 5: Add invalid outer/paragraph mutation isolation**

Test null, numbers, strings, objects, arrays, and symbols for outer `rtlMode` through both plain/rich creation and for paragraph `rtl` through creation/setter. Assert TypeError and unchanged part bytes/journal.

- [ ] **Step 6: Run typecheck and focused suites**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck exits 0 and both suites pass.

### Task 3: PptxGenJS conformance, docs, and packed surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public fields and existing PptxGenJS import/write/reopen fixture.
- Produces: real valid-output conformance, documented divergence from run-level bugs, and Node/browser/declaration smoke.

- [ ] **Step 1: Add PptxGenJS outer RTL shapes**

Generate plain text with outer true across two paragraphs, plus false and omitted controls. Assert imported/reopened paragraph snapshots are `[true, true]`, `[undefined]`, and `[undefined]`; inspect XML for valid first-position `pPr rtl="1"`. Add a TextProps run-level probe only to assert it does not create a supported `RichTextRunStyle` field or corrupt adapter reads.

- [ ] **Step 2: Update compatibility documentation**

Add a completed baseline row mapping `rtlMode` to `AddTextOptions.rtlMode` / `RichTextParagraph.rtl`; remove text-box RTL from the partial row. Document paragraph scope, explicit false, direct getter tokens, run-level PptxGenJS bug, and separation from presentation/body RTL. Add one changelog bullet and update API/package README examples.

- [ ] **Step 3: Extend packed smoke**

In Node smoke, create outer true with a paragraph false override and assert `[true, false]`. In browser smoke, assert a true paragraph. In TypeScript smoke, compile:

```ts
const rtlParagraphs: readonly RichTextParagraph[] = [
  { rtl: true, runs: [{ text: 'RTL' }] },
  { rtl: false, runs: [{ text: 'LTR' }] },
];
createdDocument.addSlide().addRichText(rtlParagraphs, { rtlMode: true });
```

- [ ] **Step 4: Run adapter, full, and performance suites**

Run the adapter test, pinned full Vitest entry, and `RUN_PERF=1` performance test. Expected: all functional tests pass, only the default performance case is skipped in the full run, and the isolated performance test passes.

### Task 4: Package, compatibility, review, commit, and push

**Files:**
- Review all Task 1–3 files; never stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and evidence.
- Produces: reviewed `feat: support paragraph rtl` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

From `packages/pptx`, run:

```sh
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
rtl_package_dir=$(mktemp -d /tmp/pptx-paragraph-rtl-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$rtl_package_dir"
node ../../scripts/smoke-npm-package.mjs "$rtl_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expect API/browser/types/CLI checks true.

- [ ] **Step 2: CLI and LibreOffice comparison**

Generate `/tmp/pptx-paragraph-rtl-native/native.pptx` and `/tmp/pptx-paragraph-rtl-pptxgenjs/pptxgenjs.pptx` with Arabic/Hebrew RTL paragraphs and explicit English LTR paragraphs. Validate with:

```sh
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-paragraph-rtl-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-paragraph-rtl-pptxgenjs/pptxgenjs.pptx --profile powerpoint-2010
```

Export both with `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf`, render every PDF page with `pdftoppm -png`, and inspect all text for direction, clipping, and repair output.

- [ ] **Step 3: Final review and commit**

Run `git diff --check`, inspect the complete diff, verify direct false/clear behavior and no `rtlCol`/presentation changes, and confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, commit `feat: support paragraph rtl`, push through SSH-over-443, and verify `origin/main...HEAD` is `0 0`.
