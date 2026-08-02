# Text Shape Preset Geometry Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict `AddTextOptions.shape` preset geometry to every text creation surface, reuse the existing live preset geometry editor, and prove PptxGenJS 4.0.1 compatible behavior in source, packed Node/browser/type/CLI, and real PPTX validation.

**Architecture:** Extract `normalizePresetShapeType()` and `renderPresetShapeGeometry()` as internal primitives from the existing preset-shape codec. Normalize one canonical token inside `validateAddTextOptions()`, render it during the original text-shape transaction, and continue using `ShapeModel.presetType` for strict reading and live replacement. Layout/master/placeholder wrappers share `SlideModel`; declarative definitions only extend their closed option key set.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, real Chrome, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- `AddTextOptions.shape` is exactly `PresetShapeType | undefined`; there is no text-only token type or alias namespace.
- Omitted or own data property `shape: undefined` renders canonical `rect`, preserving current text geometry bytes.
- The legal set is the existing 178-token `PRESET_SHAPE_TYPES`: include `foldedCorner`; reject `folderCorner`, `custGeom`, unknown strings, empty strings, coercions, and non-strings.
- Read the new field as an own data property. Reject an own accessor without invocation; do not consume inherited `shape`.
- Normalize before any relationship, part bytes, shape order, live cache, or mutation journal change.
- Plain/rich text, empty/multiline text, placeholder creation/population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects share one implementation.
- Render one direct `<a:prstGeom prst="TOKEN"><a:avLst/></a:prstGeom>` during initial creation. Do not create rect first and invoke the live setter afterward.
- Existing `ShapeModel.presetType` remains the unified strict read/edit surface. Same-value assignment is an exact no-op; replacement clears old adjustments and preserves all unrelated state.
- Geometry does not own relationships or parts. Fill, line, arrows, shadow, hyperlink, transform, placeholder identity, text body, and run hyperlinks remain isolated.
- Preserve current `p:cNvSpPr txBox="1"` behavior. `isTextBox`, `rectRadius`, text-specific adjustment shortcuts, custom points, and `breakLine` remain later slices.
- Compare all 177 common valid PptxGenJS tokens. Record `folderCorner`, `custGeom`, falsy fallback, unknown/coercion, and line-without-line-option failures as upstream behavior, not native behavior to copy.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, browser artifacts, or build output.
- Use repository-local binaries. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0`.

---

### Task 1: Shared preset geometry primitive and native text creation

**Files:**
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `PRESET_SHAPE_TYPES`, `PresetShapeType`, `NormalizedShapeAdjustments`, `renderShapeAdjustmentList()`, existing `validateAddTextOptions()`, `textShapeXml()`, and `ShapeModel.presetType`.
- Produces: `normalizePresetShapeType(value, context)`, `renderPresetShapeGeometry(type, prefix?, adjustments?)`, `AddTextOptions.shape?: PresetShapeType`, normalized default `rect`, and one-pass shaped-text rendering.

- [ ] **Step 1: Add failing codec tests**

Extend `preset-shape.internal.test.ts` to require the shared normalizer to accept every token and reject all non-canonical values:

```ts
for (const type of PRESET_SHAPE_TYPES) {
  expect(normalizePresetShapeType(type, 'Text shape geometry')).toBe(type);
  expect(renderPresetShapeGeometry(type)).toBe(
    `<a:prstGeom prst="${type}"><a:avLst/></a:prstGeom>`,
  );
}
for (const value of ['', 'folderCorner', 'custGeom', 'RECT', 1, false, null, {}]) {
  expect(() => normalizePresetShapeType(value, 'Text shape geometry')).toThrow(TypeError);
}
```

Require `renderPresetShapeXml()` to remain byte-identical after switching to the shared helper, including non-empty ordered adjustments and alternate prefixes.

- [ ] **Step 2: Add failing text creation and all-token tests**

In `model.test.ts`, create default, explicit rect, ellipse, line, lineInv, flowchart, callout, action-button, folded-corner, plain, rich, empty, and multiline text. Require immediate `presetType`, direct `prstGeom`, one empty `avLst`, unchanged text, and the existing `txBox="1"` state.

Create all 178 native tokens on one slide:

```ts
const shapes = PRESET_SHAPE_TYPES.map((shape, index) => slide.addText(
  `Geometry ${index}`,
  { name: `text_geometry_${index}`, shape },
));
expect(shapes.map(({ presetType }) => presetType)).toEqual(PRESET_SHAPE_TYPES);
```

Combine a representative shaped text with transform, fill, line, arrows, shadow, outer hyperlink, rich-run hyperlink, margins, vertical alignment, direction, fit, and wrap. Require all values to read back unchanged.

- [ ] **Step 3: Add failing strict-input and zero-mutation tests**

Snapshot slide bytes, relationships, part URI list, mutations, shape order, and an existing live model before each invalid call. Reject:

```ts
{ shape: '' }
{ shape: 'folderCorner' }
{ shape: 'custGeom' }
{ shape: 'unknown' }
{ shape: 1 }
{ shape: false }
{ shape: null }
```

Define an own getter that throws and require rejection without getter invocation. Require an inherited `shape: 'ellipse'` to behave as omitted `rect`. Require runtime own `shape: undefined` to produce exactly the omitted geometry bytes.

- [ ] **Step 4: Add failing live read/edit tests for text owners**

For shaped text, require:

- same token assignment preserves exact bytes, mutation journal, adjustments, text, and model identity;
- another token replaces only geometry and clears adjustments;
- `customGeometry` conversion and subsequent preset replacement preserve text and styles;
- fill/line/arrows/shadow/hyperlink/richText edits preserve geometry;
- malformed unknown, repeated, qualified-lookalike, missing, or preset-plus-custom geometry reads `undefined` and rejects replacement with exact zero mutation;
- duplicate, outer rollback, write/reopen, and source/duplicate independence preserve tokens.

- [ ] **Step 5: Run focused tests and confirm intended failure**

```sh
node_modules/.bin/vitest run packages/model/src/preset-shape.internal.test.ts packages/model/src/model.test.ts -t "preset geometry primitive|text shape preset geometry" --reporter=dot
```

Expected: compile/runtime failures because the shared helpers and `AddTextOptions.shape` do not exist.

- [ ] **Step 6: Extract the shared normalizer and renderer**

In `preset-shape.internal.ts`, keep the token set private and add:

```ts
const EMPTY_SHAPE_ADJUSTMENTS: NormalizedShapeAdjustments = Object.freeze([]);

export function normalizePresetShapeType(
  value: unknown,
  context: string,
): PresetShapeType {
  if (typeof value !== 'string' || !PRESET_SHAPE_TYPE_SET.has(value)) {
    throw new TypeError(`${context} must be a canonical preset shape string`);
  }
  return value as PresetShapeType;
}

export function renderPresetShapeGeometry(
  type: PresetShapeType,
  prefix = 'a:',
  adjustments: NormalizedShapeAdjustments = EMPTY_SHAPE_ADJUSTMENTS,
): string {
  return `<${prefix}prstGeom prst="${escapeXmlAttribute(type)}">` +
    `${renderShapeAdjustmentList(adjustments, prefix)}</${prefix}prstGeom>`;
}
```

Use the normalizer from `normalizePresetShape()` and `setShapePresetType()`. Use the renderer from `renderPresetShapeXml()` without changing its output.

- [ ] **Step 7: Add strict shape option normalization**

In `slide.ts`, add `shape?: PresetShapeType` to `AddTextOptions`, `shape: PresetShapeType` to both normalized text interfaces, and a small field-specific reader:

```ts
function normalizeTextShapeType(options: AddTextOptions): PresetShapeType {
  const descriptor = Object.getOwnPropertyDescriptor(options, 'shape');
  if (!descriptor) return 'rect';
  if (!Object.hasOwn(descriptor, 'value')) {
    throw new TypeError('Text shape geometry must be a data property');
  }
  if (descriptor.value === undefined) return 'rect';
  return normalizePresetShapeType(descriptor.value, 'Text shape geometry');
}
```

Do not access `options.shape` directly. Normalize alongside fill/line/arrows/shadow/hyperlink before any relationship creation.

- [ ] **Step 8: Thread the normalized token through the renderer**

Pass `shape` through `addText()`, both `addPlaceholder()` branches, `addRichText()`, `addTextShape()`, and `textShapeXml()`. Replace only the hard-coded geometry:

```ts
${renderPresetShapeGeometry(shape)}
```

Do not change transform, non-visual, fill, line, effect, hyperlink, or text-body rendering. Do not invoke `ShapeModel.presetType` after creation.

- [ ] **Step 9: Run model regression gates**

```sh
node_modules/.bin/vitest run packages/model/src/preset-shape.internal.test.ts packages/model/src/model.test.ts packages/model/src/custom-geometry.internal.test.ts packages/model/src/shape-adjustments.internal.test.ts packages/model/src/shape-hyperlink.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 10: Review, commit, push, and verify**

Review token-set reuse, accessor non-invocation, default rect byte stability, validation-before-relationship ordering, all creation call sites, direct child order, live identity, and malformed-state atomicity. Stage only the four files, commit `feat: create shaped text geometry`, push, fetch, and require divergence `0 0`.

---

### Task 2: Public owners, declarative definitions, root types, and lifecycle

**Files:**
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddTextOptions.shape`, `AddPlaceholderOptions`, `SlideLayoutModel`, `SlideMasterModel`, `SlideMasterObject`, `PRESENTATION_FORMAT_PROFILES`, and existing placeholder population.
- Produces: declarative `shape` acceptance, compile-time root exports, and public lifecycle evidence across every text owner.

- [ ] **Step 1: Add failing owner-surface tests**

Create distinct geometries through slide plain/rich text, layout plain text, master rich text, layout/master placeholders, and populated slide placeholders:

```ts
layout.addPlaceholder('Prompt', {
  name: 'geometry_title',
  type: 'title',
  shape: 'roundRect',
});
const slide = document.addSlide({ masterName: layout.name });
const populated = slide.addText('Populated', {
  placeholder: { name: 'geometry_title' },
  shape: 'ellipse',
});
expect(populated.presetType).toBe('ellipse');
expect(layout.shapes.find(({ name }) => name === 'geometry_title')?.presetType)
  .toBe('roundRect');
```

Require owner part isolation, inherited identity/name/transform, stable wrapper/model identity, and no changes to sibling objects.

- [ ] **Step 2: Add failing declarative and asynchronous-detachment tests**

Use `defineSlideMaster()` with shaped text and shaped placeholders plus delayed image resolution. Mutate the source options after the call and require the detached original token. Pass invalid/accessor values and require rejection before observable package mutation.

Expected initial failure: `TEXT_OPTION_KEYS` and `PLACEHOLDER_OPTION_KEYS` reject `shape`.

- [ ] **Step 3: Add six-format and transaction lifecycle tests**

For each `PRESENTATION_FORMAT_PROFILES` entry, create representative slide/layout/master shaped text, write/reopen, and compare name, kind, text, `presetType`, transform, placeholder identity, and styles. Duplicate and move shaped slides, then throw from an outer transaction after an edit; require exact source restoration and source/duplicate independence.

- [ ] **Step 4: Add root-package type and runtime tests**

In `packages/pptx/src/index.test.ts`, compile and run:

```ts
const typedShape: PresetShapeType = 'ellipse';
const options: AddTextOptions = { shape: typedShape };
const text = PptxDocument.create().addSlide().addText('Root geometry', options);
expect(text.presetType).toBe('ellipse');
// @ts-expect-error malformed upstream spelling is intentionally excluded
const invalidFolder: AddTextOptions = { shape: 'folderCorner' };
// @ts-expect-error custom geometry is not a preset text shape token
const invalidCustom: AddTextOptions = { shape: 'custGeom' };
```

Import `PresetShapeType` from the root package and verify write/reopen plus live replacement.

- [ ] **Step 5: Add declarative key support**

Add exactly one entry to `TEXT_OPTION_KEYS`:

```ts
'shape',
```

`PLACEHOLDER_OPTION_KEYS` inherits it. Do not add a second clone or validator; existing declarative data cloning and Task 1 normalization own those responsibilities.

- [ ] **Step 6: Run public/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape preset geometry|shaped text" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 7: Review, commit, push, and verify**

Review every public owner, placeholder source isolation, declarative cloning before async work, exact closed-key expansion, six-format persistence, compile-time exclusions, rollback, and root export reachability. Commit `test: cover public shaped text geometry`, push, fetch, and require divergence `0 0`.

---

### Task 3: PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `ShapeType`, `addText()`, `write()`, native `PRESET_SHAPE_TYPES`, package import, and `ShapeModel.presetType`.
- Produces: all-token supported parity and raw evidence for strict divergences.

- [ ] **Step 1: Add a public-only representative fixture**

Generate omitted, runtime undefined, ellipse, roundRect, line, lineInv, flowchart, callout, action-button, star, plain, multiline, and rich text via public PptxGenJS only. Supply `{ line: {} }` for `shape: 'line'` so the supported fixture writes successfully. Inspect raw geometry and import it through native reader.

Require omitted/undefined rect, exact direct token, empty `avLst`, intact text, and no geometry dependence on `isTextBox`.

- [ ] **Step 2: Compare all common legal tokens**

Use the existing catalog and exclude native-only `foldedCorner`:

```ts
const commonTypes = PRESET_SHAPE_TYPES.filter((type) => type !== 'foldedCorner');
for (const type of commonTypes) {
  const publicType = generated.ShapeType[type];
  expect(publicType, type).toBe(type);
  slide.addText(type, {
    shape: publicType!,
    ...(type === 'line' ? { line: {} } : {}),
  });
}
```

Require 177 imported text shapes in order and exact `presetType` equality. Create native equivalents and compare direct geometry token plus adjustment-list emptiness.

- [ ] **Step 3: Lock upstream defect evidence and native strictness**

Using runtime-only casts where required, prove:

- `folderCorner` writes invalid `prst="folderCorner"`, while native rejects it and writes valid `foldedCorner`;
- `custGeom` enters a points-based custom renderer and is excluded from `AddTextOptions.shape`;
- `false`/empty string fallback to rect in PptxGenJS but native rejects them;
- unknown string and number are serialized by PptxGenJS but native rejects them with exact zero mutation;
- `shape: 'line'` without a line option throws in PptxGenJS, while native emits valid line geometry with canonical no-fill line style;
- `isTextBox` changes only `cNvSpPr@txBox`; `rectRadius` changes only adjustments and remains outside this slice.

- [ ] **Step 4: Run conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text shape preset geometry" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape preset geometry|shaped text" --reporter=dot
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only generation, 177-token order, line fixture handling, raw malformed evidence, geometry-only semantic comparison, and the absence of adapter production changes. Commit `test: compare shaped text with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 4: Packed Node, declarations, browser, and CLI proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual packed `@jiayunxie/pptx`, Node/browser conditional exports, emitted declarations, installed CLI, and `pptx-inspect` commands.
- Produces: `textShapePresetGeometry: true` across Node/browser smoke plus a validated shaped-text deck.

- [ ] **Step 1: Extend packed Node runtime coverage**

Using only the installed tarball, create plain/rich/placeholder/layout/master/declarative text with default rect, ellipse, line, foldedCorner, and combined styles. Assert immediate `presetType`, same-value no-op, replacement, placeholder source isolation, duplicate independence, write/reopen, and six-format behavior. Save a shaped-text smoke deck inside the script's temporary consumer and add `textShapePresetGeometry` to final JSON.

- [ ] **Step 2: Extend installed declaration checks**

In the generated TypeScript consumer, import `AddTextOptions` and `PresetShapeType`, compile valid rect/ellipse/foldedCorner values, and add `@ts-expect-error` cases for folderCorner, custGeom, unknown strings, number, and boolean. Require emitted root types to expose the existing preset union without an internal helper export.

- [ ] **Step 3: Extend installed CLI inspection and validation**

Run the installed CLI on the generated deck:

```sh
pptx-inspect package validate text-shape-preset-geometry-smoke.pptx --profile powerpoint-2010
pptx-inspect package slides text-shape-preset-geometry-smoke.pptx --json
pptx-inspect package part-read text-shape-preset-geometry-smoke.pptx /ppt/slides/slide1.xml --json
```

Require zero errors/warnings, expected slide count, and exact shaped-text `prstGeom` tokens in the installed CLI output.

- [ ] **Step 4: Extend real Chrome coverage**

Through the browser bundle, create default/ellipse/line/foldedCorner shaped text across plain/rich/placeholder owners, edit one live token, writeBlob, reopen, and validate. Add detailed browser state plus `textShapePresetGeometry: true`; require zero page, console, network, and validation errors.

- [ ] **Step 5: Run packed smoke and build gates**

```sh
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
node_modules/.bin/tsc -b --pretty false
pnpm --filter @jiayunxie/pptx run build
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review installed-only imports, temporary artifact containment, type negatives, Node/browser parity, CLI part evidence, Chrome error capture, and no checked-in generated artifacts. Commit `test: verify packed shaped text geometry`, push, fetch, and require divergence `0 0`.

---

### Task 5: Compatibility QA, documentation, and completion audit

**Files:**
- Modify: `README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: Tasks 1–4 implementation and evidence.
- Produces: user-facing contract, compatibility boundary, progress accounting, release note, and final reproducible verification record.

- [ ] **Step 1: Run focused compatibility validation with the pptx-inspect workflow**

Read the repository `pptx-inspect` skill before this step. Generate a representative deck containing rectangle, ellipse, roundRect, line, flowchart, callout, action-button, star, foldedCorner, placeholder, rich text, and combined styles. Validate with PowerPoint 2010 profile, inspect exact slide parts, and verify only expected parts differ after live geometry edits.

- [ ] **Step 2: Run LibreOffice round-trip and visual render**

Open/save the representative deck through LibreOffice, reopen it with native model, and require every text token and text value to survive. Render to PDF/PNG and inspect that representative shapes contain visible text without clipping or missing geometry. Treat `txBox` appearance as existing behavior; do not broaden into the later `isTextBox` slice.

- [ ] **Step 3: Update user-facing documentation**

Document:

- `AddTextOptions.shape?: PresetShapeType`, default rect, 178 canonical tokens, and `ShapeModel.presetType` editing;
- all supported owners and lifecycle behavior;
- correct `foldedCorner` versus PptxGenJS `folderCorner` defect;
- strict rejection of falsy/unknown/coercible/custGeom values;
- geometry/style ownership and line geometry versus line style;
- remaining `rectRadius`, `isTextBox`, `breakLine`, advanced text/table, output/runtime, and peer-range work.

Move `shape` from unsupported to supported in compatibility/progress files without marking the broader advanced-text or PptxGenJS parity goal complete.

- [ ] **Step 4: Run complete repository and release gates**

```sh
node_modules/.bin/vitest run --reporter=dot
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
node_modules/.bin/tsup --config packages/pptx/tsup.config.ts
node_modules/.bin/tsup --config packages/pptx/tsup.browser.config.ts
pnpm --filter @jiayunxie/pptx run build
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
git diff --check
```

Record full test counts, performance duration, validation diagnostics, tarball file count/hash, Node/types/browser/CLI booleans, and Chrome error counts.

- [ ] **Step 5: Review documentation and completion claims**

Cross-check every supported claim against a test or smoke assertion. Search for stale statements that still list `shape` as missing. Ensure `rectRadius`, `isTextBox`, and `breakLine` remain explicitly outstanding and that no document claims full PptxGenJS parity.

- [ ] **Step 6: Commit, push, and verify**

Stage only the five documentation files, commit `docs: document shaped text geometry`, push, fetch, and require divergence `0 0`.

- [ ] **Step 7: Report completion and next slice**

Report this slice as complete only after all gates pass. Summarize commits, exact tests, validation, installed/runtime evidence, remaining advanced-text items, and updated overall progress. Set the next slice to `AddTextOptions.rectRadius` without marking the long-running full-parity goal complete.
