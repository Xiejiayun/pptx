# Text Shape Rectangle Radius Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict `AddTextOptions.rectRadius?: Emu` creation support for rounded-rectangle text across every owner, reuse `ShapeModel.adjustments` for live editing, and prove supported PptxGenJS 4.0.1 final semantics in source, packed Node/browser/type/CLI, and real PPTX compatibility.

**Architecture:** Read `rectRadius` as a descriptor-safe own data property during the existing text-option validation, quantize it to EMU, require `shape: 'roundRect'`, and carry the normalized radius through the one-pass text renderer. At render time, combine the radius with the final rounded owner extent to materialize one direct `adj` guide through the existing preset-geometry/adjustment codec; afterward the guide remains ordinary direct OOXML state exposed only through `ShapeModel.adjustments`.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, actual npm tarball smoke, real Chrome, `pptx-inspect` PowerPoint 2010 validation, LibreOffice, and PDF/PNG visual QA.

## Global Constraints

- `AddTextOptions.rectRadius` is exactly `Emu | undefined`; callers use the existing root-exported `inches()` for physical units.
- Omitted, inherited, or own data property `rectRadius: undefined` renders an empty direct `a:avLst`; an own accessor is rejected without invoking its getter.
- A defined radius must be a finite number in the inclusive raw range `0..914400`, is rounded once to EMU, normalizes `-0` to `0`, and requires normalized `shape === 'roundRect'`.
- Explicit zero materializes `{ name: 'adj', value: 0 }`; do not copy PptxGenJS 4.0.1 truthiness loss.
- Compute `Math.round(rectRadius * 100000 / Math.min(finalWidth, finalHeight))` from the final rounded transform. Placeholder population must use the owner extent, not caller-supplied width/height.
- Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects share the same normalizer and renderer.
- Render the guide during initial creation. Do not create an empty rounded rectangle and invoke the live adjustment setter afterward.
- Do not add `ShapeModel.rectRadius`, `AddTextOptions.adjustments`, `angleRange`, `arcThicknessRatio`, custom geometry, or automatic resize recomputation.
- Creation results are read and edited through existing `ShapeModel.adjustments`. Resize preserves the materialized guide until the caller explicitly replaces it.
- Radius owns no relationships or parts. Geometry token, fill, line, arrows, shadow, hyperlinks, text body, transform, placeholder identity, shape order, and live cache remain independently owned.
- PptxGenJS conformance compares legal positive final guide values. Lock zero loss, string coercion, wrong-shape passthrough, negative/out-of-range/non-finite behavior as intentional upstream divergences that native does not copy.
- `isTextBox` and combined `breakLine` behavior remain later slices; completing this plan does not claim full PptxGenJS parity.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, browser artifacts, LibreOffice profiles, render output, or build output.
- Use repository-local binaries. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Core model normalization, rendering, and live lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `Emu`, `EMU_PER_INCH`, `PresetShapeType`, `normalizeShapeAdjustments()`, `renderPresetShapeGeometry()`, `validateAddTextOptions()`, `textShapeXml()`, and `ShapeModel.adjustments`.
- Produces: `AddTextOptions.rectRadius?: Emu`, descriptor-safe `normalizeTextRectRadius()`, normalized radius fields on plain/rich inputs, final-extent guide calculation, and one-pass rounded-text creation.

- [ ] **Step 1: Add failing normalization and atomicity tests**

In `model.test.ts`, snapshot part bytes, relationships, part URIs, mutations, shape order, and an existing live model before each invalid call. Require rejection and exact zero mutation for:

```ts
{ shape: 'rect', rectRadius: inches(0.5) }
{ rectRadius: inches(0.5) }
{ shape: 'ellipse', rectRadius: inches(0.5) }
{ shape: 'roundRect', rectRadius: -1 as Emu }
{ shape: 'roundRect', rectRadius: (EMU_PER_INCH + 1) as Emu }
{ shape: 'roundRect', rectRadius: NaN as Emu }
{ shape: 'roundRect', rectRadius: Infinity as Emu }
{ shape: 'roundRect', rectRadius: '0.5' as never }
{ shape: 'roundRect', rectRadius: null as never }
```

Define an own accessor that throws, require no getter call, and require an inherited radius to behave as omitted. Verify own `undefined` equals omitted bytes, `-0` reads back as adjustment value `0`, and fractional EMU values quantize exactly once.

- [ ] **Step 2: Add failing guide-calculation and one-pass creation tests**

Create plain, rich, empty, multiline, and placeholder rounded text. Require:

```ts
expect(slide.addText('2x1', {
  shape: 'roundRect', width: inches(2), height: inches(1),
  rectRadius: inches(0.5),
}).adjustments).toEqual([{ name: 'adj', value: 50_000 }]);

expect(slide.addText('4x2', {
  shape: 'roundRect', width: inches(4), height: inches(2),
  rectRadius: inches(0.5),
}).adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
```

Cover explicit zero, one-inch boundary, radius and extent fractional-EMU rounding, portrait/non-square extents, and the default extent. Inspect raw XML for exactly one `a:prstGeom prst="roundRect"`, one `a:avLst`, and one `a:gd name="adj" fmla="val N"` only when the radius is defined.

- [ ] **Step 3: Add failing live editing and ownership tests**

Require immediate deep-frozen adjustment snapshots, exact same-list bytes/journal no-op, whole replacement, `[]` clear, same `presetType` preservation, and different-preset clearing. After creation call `setTransform()` and prove the guide does not recalculate. Combine radius with fill, line, arrows, shadow, outer/run hyperlinks, margins, valign, direction, fit, wrap, and rich paragraphs; edit each unrelated property and require the guide and text to stay unchanged. Cover duplicate independence, move, outer rollback, write/reopen, and malformed adjustment-state behavior through the existing live reader.

- [ ] **Step 4: Run focused tests and confirm intended failure**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts -t "text shape rectangle radius|text rect radius" --reporter=dot
```

Expected: compile/runtime failures because `AddTextOptions.rectRadius` and radius guide creation do not exist.

- [ ] **Step 5: Add the public field and strict radius normalizer**

In `slide.ts`, import `EMU_PER_INCH` plus `type Emu`, add the public field, and add the normalized field to both text-option snapshots:

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly rectRadius?: Emu;
}

function normalizeTextRectRadius(
  options: AddTextOptions,
  shape: PresetShapeType,
): Emu | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(options, 'rectRadius');
  if (!descriptor) return undefined;
  if (!Object.hasOwn(descriptor, 'value')) {
    throw new TypeError('Text rectangle radius must be a data property');
  }
  const value = descriptor.value;
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Text rectangle radius must be finite');
  }
  if (value < 0 || value > EMU_PER_INCH) {
    throw new RangeError('Text rectangle radius must be between 0 and 914400 EMU');
  }
  if (shape !== 'roundRect') {
    throw new TypeError('Text rectangle radius requires roundRect geometry');
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError('Text rectangle radius must round to a safe EMU integer');
  }
  return (Object.is(rounded, -0) ? 0 : rounded) as Emu;
}
```

Normalize `shape` first, then radius, inside `validateAddTextOptions()` before any relationship or part mutation. Preserve `undefined` instead of collapsing it to numeric zero.

- [ ] **Step 6: Thread radius through every model call site and renderer**

Add `rectRadius` immediately after `shape` in `NormalizedTextInput`, `NormalizedAddTextOptions`, all four `addTextShape()` call sites, `addTextShape()`, and `textShapeXml()`. After final `width`/`height` rounding, materialize the existing adjustment value type:

```ts
const adjustments = rectRadius === undefined
  ? undefined
  : normalizeShapeAdjustments([{
      name: 'adj',
      value: Math.round(rectRadius * 100_000 / Math.min(width, height)),
    }], 'Text rectangle radius adjustments');
```

Render only through:

```ts
renderPresetShapeGeometry(shape, 'a:', adjustments)
```

Do not add a second XML serializer or post-creation setter call. Because populated placeholders already pass `placeholderTextOptions(owner)`, the calculation must use those final owner dimensions.

- [ ] **Step 7: Run core regression gates**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/shape-adjustments.internal.test.ts packages/model/src/preset-shape.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 8: Review, commit, push, and verify**

Review descriptor safety, raw range before rounding, explicit-zero preservation, final owner extent, validation-before-relationship ordering, all renderer call sites, canonical child order, existing adjustment codec reuse, resize semantics, rollback, and unrelated-state isolation. Stage only the two files, commit `feat: create rounded text radius`, push, fetch, and require divergence `0 0`.

---

### Task 2: Public owners, declarative definitions, root types, and lifecycle

**Files:**
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddTextOptions.rectRadius`, `AddPlaceholderOptions`, `ShapeModel.adjustments`, `SlideLayoutModel`, `SlideMasterModel`, `SlideMasterObject`, `PRESENTATION_FORMAT_PROFILES`, and existing placeholder population.
- Produces: declarative radius acceptance, root-package type evidence, owner-extent population behavior, and public lifecycle coverage across all text owners.

- [ ] **Step 1: Add failing owner and placeholder-population tests**

Create distinct radii through slide plain/rich text, layout plain text, master rich text, layout/master placeholders, and populated slide placeholders. For a 4×2-inch source owner populated with caller width/height 1×1 inch, require radius 0.5 inch to produce `adj=25000`, proving the owner extent wins. Require the layout/master source guide, text, identity, transform, and sibling objects to remain unchanged.

- [ ] **Step 2: Add failing declarative and asynchronous-detachment tests**

Use `defineSlideMaster()` with rounded text and rounded placeholders plus delayed image resolution. Mutate the source radius and nested option object after the call and require the detached original value. Pass accessor, wrong-shape, and invalid-range values and require rejection before observable package mutation.

Expected initial failure: `TEXT_OPTION_KEYS` and therefore `PLACEHOLDER_OPTION_KEYS` reject `rectRadius`.

- [ ] **Step 3: Add six-format, duplicate, move, and rollback tests**

For every `PRESENTATION_FORMAT_PROFILES` entry, create representative slide/layout/master rounded text, write/reopen, and compare name, kind, text, preset token, adjustments, transform, placeholder identity, and styles. Duplicate and move a slide; then throw from an outer transaction after adjustment/transform edits and require exact source restoration, live identity restoration, and source/duplicate independence.

- [ ] **Step 4: Add root-package type and runtime tests**

In `packages/pptx/src/index.test.ts`, compile and run:

```ts
const options: AddTextOptions = {
  shape: 'roundRect',
  rectRadius: inches(0.5),
  width: inches(4),
  height: inches(2),
};
const text = PptxDocument.create().addSlide().addText('Rounded', options);
expect(text.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);

// @ts-expect-error radius uses branded EMU, not an implicit inch number
const invalidNumber: AddTextOptions = { shape: 'roundRect', rectRadius: 0.5 };
// @ts-expect-error radius is numeric EMU, not a coercible string
const invalidString: AddTextOptions = { shape: 'roundRect', rectRadius: '0.5' };
```

Import `Emu`, `AddTextOptions`, and `inches` from the root package and verify live replacement plus write/reopen.

- [ ] **Step 5: Add declarative key support**

Add exactly one key to `TEXT_OPTION_KEYS`:

```ts
'rectRadius',
```

`PLACEHOLDER_OPTION_KEYS` inherits it. Do not add a second clone or validator; existing declarative cloning and Task 1 normalization own those responsibilities.

- [ ] **Step 6: Run public/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape rectangle radius|rounded text" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 7: Review, commit, push, and verify**

Review every public owner, owner-extent calculation, placeholder source isolation, declarative clone timing, exact closed-key expansion, six-format persistence, root type branding, duplicate/move/rollback, and absence of unrelated option expansion. Commit `test: cover public rounded text radius`, push, fetch, and require divergence `0 0`.

---

### Task 3: PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()`/`write()`, native `inches()`, `AddTextOptions.rectRadius`, package import, raw slide XML helpers, and `ShapeModel.adjustments`.
- Produces: exact supported positive-radius parity and raw evidence for every intentional strict divergence.

- [ ] **Step 1: Add public-only positive-radius fixtures**

Generate text shapes through public PptxGenJS only for 2×1 inches at radius 0.5, 4×2 at radius 0.5, radius 1.0, and a representative fractional radius. Import each deck and require one `roundRect` preset plus the exact final adjustment integer. Create native equivalents with `inches(radius)` and compare `{ presetType, adjustments, text }` rather than irrelevant lexical XML differences.

- [ ] **Step 2: Prove geometry-only isolation**

For identical PptxGenJS controls with omitted versus positive radius, normalize only the `a:avLst` fragment and require all remaining shape XML to match. Repeat natively with combined fill, line, arrows, shadow, shape/run hyperlinks, and rich text; require radius to change only the direct adjustment list.

- [ ] **Step 3: Lock upstream divergence and native atomicity evidence**

Using runtime-only casts where needed, record PptxGenJS behavior for zero, numeric `NaN`, infinities, negative, over-one-inch, string-coercible radius, and positive radius on `rect`/`ellipse`. Require raw formulas or empty lists exactly as emitted. Against the same cases, require native explicit zero to produce `adj=0`, legal positive `roundRect` values to match final integers, and every invalid/wrong-shape input to throw with package state unchanged.

- [ ] **Step 4: Run conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text shape rectangle radius" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape rectangle radius|rounded text" --reporter=dot
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review that all upstream decks use public PptxGenJS APIs, final guide calculations use written extents, assertions separate supported semantics from malformed/permissive behavior, native zero intent remains explicit, invalid native calls remain atomic, and adapter production code is untouched. Commit `test: compare rounded text radius with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 4: Packed Node, declarations, browser, and CLI proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual packed `@jiayunxie/pptx`, Node/browser conditional exports, emitted declarations, installed CLI, and existing text-shape smoke sections.
- Produces: `textShapeRectRadius: true` across installed Node/browser smoke plus a validated rounded-text deck and exact part evidence.

- [ ] **Step 1: Extend actual-tarball Node runtime coverage**

Using only the installed tarball, create plain/rich/placeholder/layout/master/declarative rounded text for omitted, explicit zero, 2×1/0.5, and 4×2/0.5 radii. Assert immediate adjustments, same-list no-op, replacement/clear, resize non-recalculation, placeholder owner extent, source isolation, duplicate independence, write/reopen, and all six formats. Save the representative deck only inside the temporary consumer and add `textShapeRectRadius` to final JSON.

- [ ] **Step 2: Extend installed declaration checks**

In the generated TypeScript consumer, import `AddTextOptions`, `Emu`, and `inches`; compile legal branded values and add `@ts-expect-error` cases for unbranded `0.5`, string, boolean, and object values. Require the emitted root declaration to contain `readonly rectRadius?: Emu` and no `ShapeModel.rectRadius` alias.

- [ ] **Step 3: Extend installed CLI inspection and validation**

Run the installed CLI against the generated deck:

```sh
pptx-inspect --json package validate text-shape-rect-radius-smoke.pptx --profile powerpoint-2010
pptx-inspect --json slides list text-shape-rect-radius-smoke.pptx
pptx-inspect --json part read text-shape-rect-radius-smoke.pptx /ppt/slides/slide1.xml
```

Require zero errors/warnings, expected slide count, and exact `roundRect` guide formulas for zero, 25000, and 50000 in installed CLI output.

- [ ] **Step 4: Extend real Chrome coverage**

Through the browser bundle, create omitted/zero/positive rounded text across plain, rich, and placeholder owners; edit and clear one live adjustment list; writeBlob, reopen, and validate. Add detailed browser state plus `textShapeRectRadius: true`; require zero page, console, network, and validation errors.

- [ ] **Step 5: Run packed and build gates**

```sh
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
node_modules/.bin/tsup --config packages/pptx/tsup.config.ts
node_modules/.bin/tsup --config packages/pptx/tsup.browser.config.ts
pnpm --filter @jiayunxie/pptx run build
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review installed-only imports, type branding negatives, Node/browser parity, CLI raw-part evidence, real Chrome error capture, temporary artifact containment, and preservation of every existing smoke field. Commit `test: verify packed rounded text radius`, push, fetch, and require divergence `0 0`.

---

### Task 5: Compatibility, LibreOffice/visual QA, documentation, and completion audit

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: Tasks 1–4 implementation/evidence plus the repository `pptx-inspect` and presentation verification workflows.
- Produces: user-facing API contract, compatibility boundary, exact progress accounting, release note, and reproducible final QA record.

- [ ] **Step 1: Run focused PowerPoint compatibility validation**

Read and follow the local `pptx-inspect` skill before this step. Generate a representative deck with omitted, zero, 0.25-, 0.5-, and 1-inch radii across landscape/portrait extents, plain/rich text, layout/master placeholders, populated placeholders, and combined styles. Validate under the PowerPoint 2010 profile, inspect exact slide/layout/master parts, and prove a live adjustment edit changes only its owning part and direct guide bytes.

- [ ] **Step 2: Run LibreOffice round-trip and visual QA**

Open/save the representative deck through LibreOffice with an isolated temporary profile. Reopen it natively and require every `(owner, text, presetType, adjustments)` tuple to survive. Render source and round-tripped decks to PDF/PNG; inspect every page for visible text, distinguishable corner radii, no clipping, no overflow, and no unexpected geometry/style changes. Keep profiles, decks, PDFs, and PNGs outside the repository.

- [ ] **Step 3: Update public documentation and progress accounting**

Document:

- `AddTextOptions.rectRadius?: Emu`, `inches()` usage, allowed `0..914400` range, and `roundRect` requirement;
- omitted/undefined versus explicit-zero direct state and exact final-extent formula;
- all supported owners, placeholder owner-extent behavior, and one-pass creation;
- `ShapeModel.adjustments` as the only live read/edit surface and resize non-recalculation;
- supported PptxGenJS positive final semantics plus strict zero/type/range/wrong-shape divergences;
- remaining `isTextBox`, `breakLine`, advanced line/effect/text/table, `tableToSlides`, output/runtime helpers, and peer-range audit work.

Move `rectRadius` from pending to supported without marking advanced text or full PptxGenJS parity complete. Set `isTextBox` as the next smallest advanced-text slice.

- [ ] **Step 4: Run the complete release gates**

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

Record full test counts, performance duration, PowerPoint/LibreOffice diagnostics, tarball file count/hash, Node/types/browser/CLI booleans, Chrome validation/console/page/network counts, tuple preservation, overflow results, and page-by-page visual outcome.

- [ ] **Step 5: Review documentation and completion claims**

Cross-check every supported claim against a permanent test, smoke assertion, raw OOXML record, or compatibility artifact. Search for stale text that still lists `rectRadius` as missing. Ensure `isTextBox`, `breakLine`, and full parity remain explicitly outstanding and that no generated artifact is staged.

- [ ] **Step 6: Commit, push, and verify**

Stage only the six documentation files, commit `docs: document rounded text radius`, push, fetch, and require divergence `0 0`.

- [ ] **Step 7: Report completion and next slice**

Report this slice complete only after every gate passes. Summarize all commits, exact test/validation/package/browser evidence, completed capability, remaining roadmap, and updated overall progress. Continue inline with `AddTextOptions.isTextBox`; do not stop the full-parity program or claim the long-running goal complete.
