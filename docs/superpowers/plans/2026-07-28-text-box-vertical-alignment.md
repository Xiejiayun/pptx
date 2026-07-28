# Text Box Vertical Alignment Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native create/read/edit/clear support for top, middle, and bottom text-box vertical alignment with PptxGenJS 4.0.1 output conformance.

**Architecture:** Extract only the already-duplicated direct `txBody/bodyPr` locator and source-span attribute updater into a shared internal helper, leaving margin and vertical-alignment normalization/codecs separate. Normalize `valign` before mutation, serialize `anchor=t|ctr|b`, strictly read only those direct tokens, and expose live editing through `ShapeModel.verticalAlignment` without interpreting inheritance or rebuilding text content.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation API is `AddTextOptions.valign?: TextBoxVerticalAlignment`; live editing API is `ShapeModel.verticalAlignment`.
- Public values are exactly `top | middle | bottom`; serialize as direct `bodyPr@anchor` tokens `t | ctr | b`.
- Omitted creation remains explicit middle (`anchor="ctr"`); live setter `undefined` removes the direct anchor.
- Read only exact direct `t/ctr/b`; absent, `just`, `dist`, empty, case variants, whitespace, and unknown tokens return `undefined`.
- Preserve `just/dist` and all unowned body properties during getters and unrelated edits.
- Shared `text-body-properties.internal.ts` may only locate direct body properties and update one attribute template; it must not model fit children, inheritance, or the full bodyPr schema.
- Paragraph alignment, table cells, placeholders/master inheritance, `anchorCtr`, fit, wrap, direction, RTL, and columns are outside this item.
- PptxGenJS is used only by the adapter conformance test and never becomes a non-adapter dependency.
- Do not stage, delete, or otherwise modify `.pnpm-store/`.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.

---

### Task 1: Shared body-properties helper and strict vertical-alignment model codec

**Files:**
- Create: `packages/model/src/text-body-properties.internal.ts`
- Modify: `packages/model/src/text-box-margins.internal.ts`
- Create: `packages/model/src/text-box-vertical-alignment.internal.ts`
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: existing lossless bodyPr margin implementation, `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `SlideModel` delegation, and stable `ShapeModel` identity.
- Produces: `TextBoxVerticalAlignment`, shared `requireTextBodyProperties()`/`updateTextBodyAttribute()`, vertical normalize/render/read/replace functions, and live accessors.

- [ ] **Step 1: Write a failing strict-read and lossless-edit model test**

Inject expanded direct body properties containing margins, `anchor="t"`, `anchorCtr="1"`, an autofit child, and an unknown child:

```ts
const bodyProperties = [
  '<a:bodyPr wrap="square" lIns="127000" anchor="t" anchorCtr="1" custom="KEEP">',
  '<a:normAutofit fontScale="90000"/>',
  '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
  '</a:bodyPr>',
].join('');
```

Assert `shape.verticalAlignment === 'top'`, repeated getter access adds no mutation, and reads remain live after edits. Set middle then bottom and verify only the anchor value changes. Set `undefined`, assert direct anchor disappears while `anchorCtr`, margin, autofit/unknown children, paragraphs/runs, text, and stable identity survive.

- [ ] **Step 2: Run the model test and confirm it fails before implementation**

Run:

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: failure because `ShapeModel.verticalAlignment` does not exist.

- [ ] **Step 3: Extract the narrow shared helper without changing margin behavior**

Move only these responsibilities from `text-box-margins.internal.ts` into `text-body-properties.internal.ts`:

```ts
export function requireTextBodyProperties(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): XmlElement;

export function updateTextBodyAttribute(
  template: string,
  name: string,
  value: string | undefined,
): string;
```

The locator must require direct `txBody/bodyPr`; the updater must preserve existing quote style on replacement, remove only the target attribute plus its preceding horizontal whitespace, and safely append to expanded or self-closing templates. Rewire the margin codec to import these two functions, with no public or serialized behavior change.

- [ ] **Step 4: Add the public vertical-alignment type and internal codec**

In `packages/model/src/text.ts` add:

```ts
export type TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom';
```

Create `text-box-vertical-alignment.internal.ts` with exact maps:

```ts
const TO_OOXML = { top: 't', middle: 'ctr', bottom: 'b' } as const;
const FROM_OOXML = new Map([
  ['t', 'top'],
  ['ctr', 'middle'],
  ['b', 'bottom'],
]);
```

Implement:

```ts
normalizeTextBoxVerticalAlignment(value: unknown, context: string): TextBoxVerticalAlignment;
renderTextBoxVerticalAlignmentAttribute(value: TextBoxVerticalAlignment): string;
readTextBoxVerticalAlignment(xml, shape, partUri): TextBoxVerticalAlignment | undefined;
replaceTextBoxVerticalAlignment(xml, shape, value, partUri): void;
```

Normalization must accept only the three exact strings. Read only the direct exact token. Replace the single `anchor` attribute through the shared helper.

- [ ] **Step 5: Delegate live reads and edits through SlideModel/ShapeModel**

Add to `SlideModel`:

```ts
getShapeTextVerticalAlignment(id: number): TextBoxVerticalAlignment | undefined;
setShapeTextVerticalAlignment(id: number, value: TextBoxVerticalAlignment | undefined): void;
```

The setter must validate non-undefined input before applying any patch, execute in `opcPackage.transaction()`, and save through `setXml()`.

Add asymmetric accessors to `ShapeModel`:

```ts
get verticalAlignment(): TextBoxVerticalAlignment | undefined;
set verticalAlignment(value: TextBoxVerticalAlignment | undefined);
```

- [ ] **Step 6: Complete strict-token, preservation, malformed-shape, and rollback coverage**

Use fresh fixtures for `t`, `ctr`, `b`, absent, `just`, `dist`, empty, `T`, ` middle `, and unknown tokens. Assert only `t/ctr/b` map to public values and getter access never mutates. Assert plain/rich replacement, `textMargins`, and transform edits preserve unsupported anchors byte-for-byte. Extend malformed `txBody/bodyPr` coverage to getter/setter. Verify inner failure and outer rollback restore exact bytes, journal, margins, text, vertical snapshot, and identity.

- [ ] **Step 7: Run focused model and margin regressions**

Run the Task 1 model command. Expected: all model tests, including the existing text-box-margin suite after helper extraction, pass.

### Task 2: Plain/rich creation lifecycle and invalid-input isolation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 type/codec, `AddTextOptions`, shared plain/rich `addTextShape()` path, duplicate/write/reopen transactions.
- Produces: `AddTextOptions.valign`, explicit middle default, top/middle/bottom creation, and native lifecycle evidence.

- [ ] **Step 1: Write failing creation and lifecycle tests**

Create plain/rich shapes for omitted, top, middle, and bottom, combining top with margins and rich bottom with paragraph alignment/bullet/run styles:

```ts
const omitted = slide.addText('Omitted');
const top = slide.addText('Top', { valign: 'top', margin: 8 });
const middle = slide.addRichText([{ runs: [{ text: 'Middle' }] }], { valign: 'middle' });
const bottom = slide.addRichText([{ align: 'center', runs: [{ text: 'Bottom', style: { bold: true } }] }], {
  valign: 'bottom',
});
```

Assert snapshots are middle/top/middle/bottom and XML contains `anchor="ctr"`, `anchor="t"`, and `anchor="b"` on the expected body properties. Duplicate the slide, edit top to bottom, clear middle, add top to a self-closing bodyPr that lacks anchor, and verify edited versus duplicated values after write/reopen.

- [ ] **Step 2: Normalize creation options and serialize the target token**

Extend `AddTextOptions` with `readonly valign?: TextBoxVerticalAlignment`. Extend normalized creation options with a required `verticalAlignment`; use middle when omitted, otherwise call the strict normalizer. Pass the normalized value through `addTextShape()` to `textShapeXml()` and replace the hard-coded `anchor="ctr"` with `renderTextBoxVerticalAlignmentAttribute(verticalAlignment)`.

- [ ] **Step 3: Add invalid-input mutation-isolation cases**

Test `'center'`, `'t'`, `'ctr'`, `'b'`, empty, case variants, whitespace, boolean, number, null, object, array, and symbol. For `addText()`, `addRichText()`, and the live setter, assert throws leave slide bytes, mutation journal, text, margins, vertical snapshot, and identity unchanged.

- [ ] **Step 4: Run typecheck and focused model/SDK tests**

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
- Consumes: Task 1/2 public API and existing PptxGenJS public-write fixture.
- Produces: real 4.0.1 vertical-alignment evidence, corrected compatibility matrix, public docs, and Node/browser/declaration smoke.

- [ ] **Step 1: Add PptxGenJS-generated vertical-alignment shapes**

Generate named shapes for omitted, top, middle, bottom, and one `TextProps[]` run with `options.valign: 'bottom'` but no text-box valign. Import public bytes and assert snapshots are middle/top/middle/bottom/middle. Assert raw direct tokens, then write/reopen and verify the same values on the duplicated slide.

- [ ] **Step 2: Advance compatibility and API documentation**

Add a completed matrix row:

```md
| 文本框 `valign: top/middle/bottom` 与 direct 编辑 | `AddTextOptions.valign` / `ShapeModel.verticalAlignment` | 已支持 |
```

Remove valign from the remaining text-box partial row. Document creation default middle versus live clear semantics in `docs/api/README.md`, mention vertical alignment in the npm README, and add one Unreleased changelog bullet.

- [ ] **Step 3: Extend packed-package smoke**

In Node smoke, create top and set bottom through the live property. In browser smoke, create bottom and assert the snapshot. In generated TypeScript smoke, import/use:

```ts
type TextBoxVerticalAlignment,

const valign: TextBoxVerticalAlignment = 'top';
const aligned = createdDocument.addSlide().addText('Aligned', { valign });
aligned.verticalAlignment = 'bottom';
```

Ensure packed declarations accept the creation option, getter, setter, and clear assignment.

- [ ] **Step 4: Run adapter, typecheck, full tests, and performance**

Run the adapter test, full typecheck, full Vitest suite, and isolated `RUN_PERF=1` performance suite. Expected: all functional tests pass, default performance remains skipped in the full run, and the isolated performance test passes.

### Task 4: Package, CLI, LibreOffice, review, commit, and push

**Files:**
- Review all files listed in Tasks 1–3; do not stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed implementation/tests/docs.
- Produces: reviewed `feat: support text box vertical alignment` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the real npm tarball**

From `packages/pptx`, run the pinned Node/browser tsup builds, declaration builder, and `npm pack --ignore-scripts --pack-destination /tmp/pptx-text-valign-package`. Run `node scripts/smoke-npm-package.mjs /tmp/pptx-text-valign-package/jiayunxie-pptx-0.1.0.tgz`. Expected: API, browser, types, and CLI checks are true.

- [ ] **Step 2: Validate and render native/PptxGenJS comparison decks**

Generate wide-layout decks at:

```text
/tmp/pptx-text-valign-validation/native.pptx
/tmp/pptx-text-valign-validation/pptxgenjs.pptx
```

Each contains equal-height filled boxes with top, middle, and bottom alignment plus explicit margins. Validate both with the repository CLI under `powerpoint-2010`; expect zero errors/warnings. Export through `soffice --headless --convert-to pdf`, confirm source hashes do not change and no repair output appears, render with `pdftoppm -png -r 150`, and visually confirm three distinct vertical positions plus native/PptxGenJS parity.

- [ ] **Step 3: Final review**

Run `git diff --check`; inspect the complete diff; verify the shared helper contains only direct bodyPr location/attribute update; verify margin serialization/tests are unchanged; confirm every other changed line traces to vertical alignment; verify status lists only intended files plus untracked `.pnpm-store/`. Re-run focused tests after any review fix.

- [ ] **Step 4: Commit and push**

Stage only intended files, commit with:

```sh
git commit -m "feat: support text box vertical alignment"
```

Push `main` through SSH-over-443 and verify `main...origin/main` plus only untracked `.pnpm-store/`.
