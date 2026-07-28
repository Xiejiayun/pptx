# Rich Text Character Spacing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native create/read/edit/clear support for point-based rich-text character spacing with PptxGenJS 4.0.1 output conformance.

**Architecture:** Extend the existing `RichTextRunStyle` vertical slice in `packages/model/src/rich-text.internal.ts`; normalize point input to OOXML hundredths of a point before mutation, render `spc` plus `kern="0"`, and conservatively read a strict Int32 `spc`. Reuse the current rich-text snapshot, transaction, duplicate, adapter, package, and release-smoke paths without introducing a generic typography codec.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, LibreOffice headless.

## Global Constraints

- Public API is `RichTextRunStyle.characterSpacing?: number` in points.
- Write `spc = round(points * 100)` and `kern="0"`; explicit zero must be written.
- Accept only finite input whose quantized raw value fits signed Int32.
- Read only strict integer `rPr@spc` within signed Int32; return `raw / 100`.
- A lone or malformed `kern` never creates public character spacing.
- Getter and non-rich-text mutations preserve source XML; `shape.richText` remains whole-run replacement.
- PptxGenJS is used only by the adapter conformance test and never becomes a non-adapter dependency.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.

---

### Task 1: Model API, validation, serialization, and strict read

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: existing `RichTextRunStyle`, `readIntegerAttribute()`, `MIN_COORDINATE_32`, `MAX_COORDINATE_32`, `renderRun()`, and `readStyle()`.
- Produces: `RichTextRunStyle.characterSpacing?: number`, normalized hundredth-point values, serialized `spc`/`kern`, and strict public snapshots.

- [ ] **Step 1: Add a failing strict-read regression test**

Add a model fixture test containing valid `spc="250"`, `spc="-125"`, `spc="0" kern="1200"`, signed Int32 boundaries, and invalid decimal/scientific/empty/out-of-range values. Assert valid snapshots are `[2.5, -1.25, 0, -21474836.48, 21474836.47]`, invalid values are `undefined`, other styles remain readable, and getter access does not add mutations.

```ts
expect(shape.richText[0]!.runs.map(({ style }) => style?.characterSpacing)).toEqual([
  2.5, -1.25, 0, -21_474_836.48, 21_474_836.47,
  undefined, undefined, undefined, undefined, undefined,
]);
expect(pkg.mutations).toEqual(journal);
```

- [ ] **Step 2: Run the model test and confirm it fails before implementation**

Run:

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: failure because `style.characterSpacing` is absent.

- [ ] **Step 3: Add the public property and normalizer**

In `packages/model/src/text.ts` add:

```ts
readonly characterSpacing?: number;
```

In `normalizeStyle()`, allow `characterSpacing` and normalize it with:

```ts
function normalizeCharacterSpacing(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const raw = Math.round(value * 100);
  if (raw < MIN_COORDINATE_32 || raw > MAX_COORDINATE_32) {
    throw new RangeError(`${context} must fit the OOXML Int32 point range`);
  }
  return raw / 100;
}
```

Use `!== undefined` when copying the normalized property so zero survives.

- [ ] **Step 4: Serialize and read the OOXML attributes**

In `renderRun()`, add both attributes when the property is present:

```ts
style.characterSpacing === undefined ? '' : `spc="${Math.round(style.characterSpacing * 100)}"`,
style.characterSpacing === undefined ? '' : 'kern="0"',
```

In `readStyle()`, call:

```ts
function readCharacterSpacing(xml: LosslessXmlDocument, properties: XmlElement): number | undefined {
  if (!xml.attribute(properties, 'spc')) return undefined;
  const raw = readIntegerAttribute(xml, properties, 'spc');
  return raw !== undefined && raw >= MIN_COORDINATE_32 && raw <= MAX_COORDINATE_32
    ? raw / 100
    : undefined;
}
```

- [ ] **Step 5: Verify model read and plain-text preservation**

Extend the fixture test to set `shape.text` to two paragraphs and assert the first run's original `spc`/`kern` template appears twice. Re-run the model test and expect all cases to pass.

### Task 2: Native SDK lifecycle and validation isolation

**Files:**
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: `SlideModel.addRichText()`, `ShapeModel.richText`, `PptxDocument.transaction()`, duplication, write/reopen, and the Task 1 property.
- Produces: evidence for create/edit/clear/duplicate/rollback/quantization/boundary behavior.

- [ ] **Step 1: Add lifecycle coverage**

Create runs with `2.5`, `-1.25`, `0`, `0.004`, both signed Int32 point boundaries, and a combined baseline/glow/outline/underline/strike run. Assert `0.004` normalizes to `0`, XML contains `spc="0" kern="0"`, editing replaces spacing, omission clears it, duplication preserves it, rollback restores bytes and identity, and write/reopen retains the values.

```ts
const shape = slide.addRichText([{ runs: [
  { text: 'Expanded', style: { characterSpacing: 2.5 } },
  { text: 'Condensed', style: { characterSpacing: -1.25 } },
  { text: 'Normal', style: { characterSpacing: 0 } },
  { text: 'Quantized', style: { characterSpacing: 0.004 } },
] }]);
expect(shape.richText[0]!.runs.map(({ style }) => style?.characterSpacing))
  .toEqual([2.5, -1.25, 0, 0]);
```

- [ ] **Step 2: Add invalid-input mutation-isolation cases**

Add `null`, boolean, string, object, array, `NaN`, positive infinity, `-21474836.49`, and `21474836.48` to the existing invalid rich-text matrix. Assert both `addRichText()` and the setter throw and that slide bytes plus the mutation journal remain unchanged.

```ts
[{ runs: [{ text: 'x', style: { characterSpacing: Number.NaN } }] }],
[{ runs: [{ text: 'x', style: { characterSpacing: 21_474_836.48 } }] }],
```

- [ ] **Step 3: Run typecheck and model/SDK tests**

Run:

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: typecheck exits 0 and both suites pass.

### Task 3: PptxGenJS import conformance and public release surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 1 public property and existing PptxGenJS import/write/reopen fixture.
- Produces: actual 4.0.1 compatibility evidence and Node/browser/types package smoke coverage.

- [ ] **Step 1: Add a PptxGenJS-generated spacing shape**

Append runs with `charSpacing: 2.5`, `-1.25`, `0.004`, exact `0`, spacing plus `baseline: 600`, and no spacing. Assert imported/reopened values are `[2.5, -1.25, 0, undefined, 3, undefined]` and the combined run still reads `baseline: 'superscript'`.

```ts
generatedSlide.addText([
  { text: 'Positive', options: { charSpacing: 2.5 } },
  { text: ' Negative', options: { charSpacing: -1.25 } },
  { text: ' Combined', options: { charSpacing: 3, baseline: 600 } },
], { x: 9, y: 9, w: 3, h: 1 });
```

- [ ] **Step 2: Update the compatibility surface**

Add a completed matrix row mapping PptxGenJS `charSpacing` to `RichTextRunStyle.characterSpacing`, remove character spacing from the partial-support row, mention the feature in README, and add one Unreleased changelog bullet.

- [ ] **Step 3: Extend release smoke**

Exercise positive spacing in Node, explicit zero in browser, and import `RichTextRunStyle` through the packed declaration path with a typed `characterSpacing` value. Add runtime assertions for normalized values.

```ts
const style: RichTextRunStyle = { characterSpacing: 2.5 };
const run = created.slides[0].addRichText([{ runs: [{ text: 'Spaced', style }] }])
  .richText[0]!.runs[0]!;
if (run.style?.characterSpacing !== 2.5) throw new Error('Character spacing smoke failed');
```

- [ ] **Step 4: Run adapter and full repository tests**

Run the adapter test, then the full Vitest suite and the isolated performance suite. Expected: all functional tests pass, the default performance test remains skipped in the full run, and the explicit performance run passes.

### Task 4: Package, compatibility, review, commit, and push

**Files:**
- Review all files listed in Tasks 1–3; do not stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation and tests.
- Produces: reviewed `feat: support rich text character spacing` commit on `origin/main`.

- [ ] **Step 1: Build and smoke the actual npm tarball**

Build Node/browser bundles and declarations with the pinned tsup/TypeScript runtimes, pack with `npm pack --ignore-scripts`, and run `node scripts/smoke-npm-package.mjs <tarball>`. Expected: API, browser, types, and CLI checks are all true.

- [ ] **Step 2: Validate and render native/PptxGenJS comparison decks**

Generate expanded and condensed runs from both libraries, validate both with `packages/pptx/dist/cli.js --json package validate --profile powerpoint-2010`, export both through LibreOffice, render the PDFs with `pdftoppm`, and visually confirm matching spacing direction and magnitude without repair prompts.

- [ ] **Step 3: Final review**

Run `git diff --check`, inspect the complete diff, verify `git status --short` lists only the intended nine implementation files plus untracked `.pnpm-store/`, and re-run the focused tests after any review fix.

- [ ] **Step 4: Commit and push**

Stage only the intended implementation files, commit with:

```sh
git commit -m "feat: support rich text character spacing"
```

Push `main` using the repository's SSH-over-443 command and verify `main...origin/main` is synchronized.
