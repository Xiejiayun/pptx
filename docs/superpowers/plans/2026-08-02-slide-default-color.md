# Slide Default Color Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict `SlideModel.color` state whose current value is materialized into subsequently created plain/rich text runs, with PptxGenJS 4.0.1 public-output parity and complete runtime/package verification.

**Architecture:** `PresentationModel` owns a transient per-slide color map because OOXML has no legal slide-level default-text-color field. `SlideModel.color` normalizes `RichTextColor` into that map, while the existing plain/rich text renderers receive an optional captured default and materialize it as standard run-level `a:solidFill`; explicit run color remains higher priority.

**Tech Stack:** TypeScript 5.8, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, JSZip, PptxGenJS 4.0.1, tsup, real-Chrome browser smoke, `pptx-inspect`, LibreOffice/Poppler.

## Global Constraints

- Execute inline in the current task; do not delegate implementation or review.
- Public property is exactly `SlideModel.color`; its value is existing `RichTextColor | undefined`.
- State is transient and must not create a custom extension, `p:clrMapOvr`, `p:txStyles`, layout/master write, or any other slide-level OOXML color field.
- New plain/rich run precedence is local run color, then current slide default, then existing canonical `schemeClr tx1`.
- Existing shapes, `ShapeModel.richText` edits, tables, charts, notes, slide numbers, placeholders, layouts, masters, and sibling slides must not read or change the slide default.
- Setter input is ordinary/null-prototype, own-data, closed, copied, normalized, frozen, and validated before state change.
- sRGB accepts optional `#`, normalizes to six uppercase hex digits; scheme tokens reuse the existing supported `RichTextColor` set.
- Equal assignment and absent clear preserve snapshot identity and produce zero OPC part/relationship/graph/journal mutation.
- Duplicate copies the source transient default only after package transaction success; move retains it by part URI; delete removes it only after transaction success.
- Reopened documents have `slide.color === undefined`; colors already materialized into runs must remain readable and visually unchanged.
- Keep native zero-input/cleared behavior as theme-aware `tx1`; do not copy PptxGenJS hard-coded black or invalid-string fallback.
- Apply identical output behavior to `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm` without changing format profiles.
- Production packages do not import or call PptxGenJS; the locked 4.0.1 package remains adapter/test evidence only.
- Use fresh `/tmp/pptx-slide-default-color-*` directories for tarballs, fixtures, galleries, renders, validation output, and client round trips.
- Every task ends with focused tests, typecheck where applicable, diff review, one isolated commit, SSH push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage or commit `.pnpm-store/`, generated decks, renders, packed tarballs, or temporary client artifacts.

---

## File Map

- Modify `packages/model/src/rich-text.internal.ts`: expose strict shared color normalization and accept an internal render default.
- Modify `packages/model/src/presentation.ts`: own transient default-color map and duplicate/move/delete lifecycle.
- Modify `packages/model/src/slide.ts`: expose `SlideModel.color` and pass captured defaults into new plain/rich text rendering.
- Modify `packages/model/src/model.test.ts`: state, lifecycle, renderer, isolation, rollback, all-format, and reopen coverage.
- Modify `packages/sdk/src/index.test.ts`: public create/write/open behavior across SDK boundary.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: public PptxGenJS output conformance and intentional differences.
- Modify `packages/validator/src/index.test.ts`: canonical native/imported packages validate without new diagnostics.
- Modify `packages/pptx/src/index.test.ts`: root-package public type/runtime coverage.
- Modify `scripts/build-npm-package-types.mjs`: compile installed `Slide.color`/`RichTextColor` consumer.
- Modify `scripts/smoke-npm-package.mjs`: actual tarball Node/CLI/default-color evidence and `slideDefaultColor: true`.
- Modify `scripts/playwright-browser-smoke.js`: real-Chrome state/materialization/duplicate/reopen checks.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: final public contract, evidence, remaining scope, and next roadmap item.

---

### Task 1: Strict public state and lifecycle

**Files:**

- Modify: `packages/model/src/rich-text.internal.ts`
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: existing `RichTextColor`, `PresentationModel` stable slide identity, and package transaction lifecycle.
- Produces:

```ts
export function normalizeRichTextColor(
  value: unknown,
  context: string,
): Readonly<RichTextColor>;

export class SlideModel {
  get color(): Readonly<RichTextColor> | undefined;
  set color(value: RichTextColor | undefined);
}
```

`PresentationModel` package-internal state methods:

```ts
getSlideDefaultColor(partUri: string): Readonly<RichTextColor> | undefined;
setSlideDefaultColor(
  partUri: string,
  value: Readonly<RichTextColor> | undefined,
): void;
```

- [ ] **Step 1: Write failing normalization/state tests**

Add a `describe('slide default text color state', ...)` block after slide-number lifecycle coverage in
`packages/model/src/model.test.ts`. The first test must assert this exact public contract:

```ts
const { pkg, model } = emptyPresentationModel();
const slide = model.addSlide();
const input = { kind: 'srgb' as const, value: '#ff3399' };
const before = packageSnapshot(pkg);

slide.color = input;
expect(slide.color).toEqual({ kind: 'srgb', value: 'FF3399' });
expect(slide.color).not.toBe(input);
expect(Object.isFrozen(slide.color)).toBe(true);
expect(packageSnapshot(pkg)).toEqual(before);

const snapshot = slide.color;
slide.color = { kind: 'srgb', value: 'FF3399' };
expect(slide.color).toBe(snapshot);
expect(packageSnapshot(pkg)).toEqual(before);

slide.color = undefined;
expect(slide.color).toBeUndefined();
slide.color = undefined;
expect(packageSnapshot(pkg)).toEqual(before);
```

Run:

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts -t "slide default text color state"
```

Expected: FAIL because `SlideModel.color` does not exist.

- [ ] **Step 2: Write failing strict-input tests**

Cover valid scheme, null-prototype sRGB, and all invalid categories without package mutation:

```ts
const invalid = [
  null,
  'FF3399',
  [],
  { kind: 'srgb' },
  { kind: 'srgb', value: 'FFF' },
  { kind: 'scheme', value: 'unknown' },
  { kind: 'srgb', value: 'FF3399', extra: true },
  Object.create({ kind: 'srgb', value: 'FF3399' }),
  Object.defineProperty({ kind: 'srgb' }, 'value', { get: () => 'FF3399' }),
  { kind: 'srgb', value: 'FF3399', [Symbol('extra')]: true },
];
```

For every candidate, assert a throw, unchanged previous `slide.color`, and unchanged `packageSnapshot(pkg)`.
Also use a revoked Proxy and assert the underlying error leaves state/package unchanged.

- [ ] **Step 3: Implement descriptor-safe shared color normalization**

Replace private `normalizeColor()` with exported internal `normalizeRichTextColor()` and make existing rich-text style,
highlight, underline, outline, and glow callers use it. Read exactly two data descriptors using this shape:

```ts
function readColorDataObject(value: unknown, context: string): {
  readonly kind: unknown;
  readonly value: unknown;
} {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const keys = Reflect.ownKeys(value);
  if (keys.some((key) => typeof key !== 'string' || !['kind', 'value'].includes(key))) {
    throw new TypeError(`${context} contains unsupported properties`);
  }
  const kind = Object.getOwnPropertyDescriptor(value, 'kind');
  const colorValue = Object.getOwnPropertyDescriptor(value, 'value');
  if (!kind || !Object.hasOwn(kind, 'value') || !colorValue || !Object.hasOwn(colorValue, 'value')) {
    throw new TypeError(`${context} must contain kind and value data properties`);
  }
  return { kind: kind.value, value: colorValue.value };
}
```

Return `Object.freeze({ kind, value })`, uppercase sRGB without `#`, and frozen scheme values. Do not export this helper
from `packages/model/src/index.ts`; only sibling implementation files consume it.

- [ ] **Step 4: Implement `SlideModel.color` and transient map methods**

Add to `PresentationModel`:

```ts
readonly #slideDefaultColors = new Map<string, Readonly<RichTextColor>>();

/** @internal */
getSlideDefaultColor(partUri: string): Readonly<RichTextColor> | undefined {
  return this.#slideDefaultColors.get(partUri);
}

/** @internal */
setSlideDefaultColor(
  partUri: string,
  value: Readonly<RichTextColor> | undefined,
): void {
  const current = this.#slideDefaultColors.get(partUri);
  if (colorsEqual(current, value)) return;
  if (value === undefined) this.#slideDefaultColors.delete(partUri);
  else this.#slideDefaultColors.set(partUri, value);
}
```

Keep `colorsEqual()` local to `presentation.ts` and compare only `kind`/`value`. Add to `SlideModel`:

```ts
get color(): Readonly<RichTextColor> | undefined {
  return this.presentation.getSlideDefaultColor(this.partUri);
}

set color(value: RichTextColor | undefined) {
  const normalized = value === undefined
    ? undefined
    : normalizeRichTextColor(value, 'Slide default text color');
  this.presentation.setSlideDefaultColor(this.partUri, normalized);
}
```

Do not parse or require the slide part in either accessor.

- [ ] **Step 5: Write failing lifecycle/rollback tests**

Test sibling isolation, duplicate copy, move retention, successful delete cleanup with reused `/ppt/slides/slideN.xml`,
failed duplicate no-entry leak, and failed delete retention. The stable lifecycle expectation is:

```ts
source.color = { kind: 'scheme', value: 'accent2' };
const duplicate = model.duplicateSlide(model.slides.indexOf(source));
expect(duplicate.color).toEqual(source.color);
expect(duplicate.color).toBe(source.color);

model.moveSlide(model.slides.indexOf(duplicate), 0);
expect(model.slides[0]).toBe(duplicate);
expect(duplicate.color).toEqual({ kind: 'scheme', value: 'accent2' });
```

Inject `pkg.setPart` failures using the existing Vitest spy pattern and assert failed operations do not alter the map-visible
values.

- [ ] **Step 6: Commit lifecycle changes only after focused review**

Run:

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts -t "slide default text color state"
pnpm typecheck
git diff --check
git diff -- packages/model/src/rich-text.internal.ts packages/model/src/presentation.ts packages/model/src/slide.ts packages/model/src/model.test.ts
```

Expected: all pass; diff contains state/normalizer/lifecycle only, no renderer behavior.

Commit and push:

```bash
git add packages/model/src/rich-text.internal.ts packages/model/src/presentation.ts packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: add slide default color state"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 2: Materialize defaults into new plain and rich text

**Files:**

- Modify: `packages/model/src/rich-text.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: Task 1 `SlideModel.color` and `normalizeRichTextColor()`.
- Produces an internal renderer option:

```ts
interface RenderRichTextOptions {
  readonly defaultColor?: Readonly<RichTextColor>;
}
```

- [ ] **Step 1: Add failing plain-text materialization tests**

Create a slide with sRGB default and `addText('First\nSecond')`. Assert every non-empty run reads back
`{ kind: 'srgb', value: 'FF3399' }`, direct slide XML has exactly two matching `a:srgbClr`, and the getter stays set.
Add theme default and zero-input controls; zero-input must still read `scheme tx1`.

- [ ] **Step 2: Thread default color through the plain-text renderer**

Extend only these private signatures:

```ts
function textParagraphXml(
  value: string,
  prefix: string,
  /* existing arguments */
  defaultColor?: Readonly<RichTextColor>,
): string;

function defaultTextRunXml(
  value: string,
  prefix: string,
  language?: string,
  defaultColor?: Readonly<RichTextColor>,
): string;
```

Use the same canonical color-choice renderer as rich text; do not duplicate sRGB/scheme XML selection. Capture
`this.color` once before building all paragraphs so a single `addText()` call cannot observe changing state mid-render.

- [ ] **Step 3: Add failing rich inheritance/override/transparency tests**

Use one `addRichText()` call containing:

```ts
[
  { text: 'Inherited' },
  { text: 'Override', style: { color: { kind: 'srgb', value: '00AA00' } } },
  { text: 'Transparent inherited', style: { transparency: 25 } },
]
```

With slide default `accent1`, assert colors `[accent1, 00AA00, accent1]`, transparencies
`[undefined, undefined, 25]`, and canonical alpha `75000` only on the third run.

- [ ] **Step 4: Add `defaultColor` to rich-text rendering**

Change `renderRichTextParagraphs()` to call:

```ts
renderRun(run, prefix, options.defaultLanguage, options.defaultColor)
```

and resolve:

```ts
const color = style.color ?? defaultColor ?? { kind: 'scheme' as const, value: 'tx1' };
```

Pass the captured `this.color` only from `SlideModel.addRichText()`. Do not pass it from `replaceRichText()`, table
rendering, slide-number rendering, or any edit path.

- [ ] **Step 5: Lock temporal and non-target isolation**

Add tests that create text under color A, change to B, create more text, clear, create a final text, then assert A/B/tx1
by object. Verify changing/clearing does not change existing XML bytes. Verify `shape.richText = ...` after setting slide
color still uses explicit/tx1 behavior, and `addTable()` remains tx1 under native rendering rather than slide color.

- [ ] **Step 6: Lock duplicate/write/reopen and all six formats**

For every `PRESENTATION_FORMAT_PROFILES` entry:

1. create/open the format fixture;
2. set source slide default and create plain/rich text;
3. duplicate and create new text on duplicate without resetting color;
4. write/reopen;
5. assert materialized run colors survive and reopened `slide.color` is `undefined`.

Add SDK coverage using `PptxDocument.create()` for the ordinary `pptx` path and ensure strict write has zero errors.

- [ ] **Step 7: Review, test, commit, and push materialization**

Run:

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts -t "slide default"
pnpm typecheck
pnpm build
git diff --check
```

Review changed XML snapshots for `solidFill`, scheme/sRGB choice, alpha placement, and no layout/master/custom extension.

Commit and push:

```bash
git add packages/model/src/rich-text.internal.ts packages/model/src/slide.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git commit -m "feat: materialize slide default colors"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 3: Lock PptxGenJS public-output conformance

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/validator/src/index.test.ts`

**Interfaces:**

- Consumes: public PptxGenJS 4.0.1 `slide.color`, `addText()`, `addTable()`, and write bytes; native
  `PptxDocument.open()`/`ShapeModel.richText`.
- Produces: permanent final-state conformance evidence without private-field access.

- [ ] **Step 1: Extend only the local public-test facade**

Add to `PptxGenJSSlide`:

```ts
color?: string;
```

This is test typing only. Do not read it after assignment and do not access runtime private fields.

- [ ] **Step 2: Generate valid public cases**

Build independent presentations for sRGB lowercase, scheme `accent1`, rich inherited/local override, color change over
time, cleared default, and table isolation. Use public `write()` through existing `importPptxGenJS()`.

- [ ] **Step 3: Compare native semantic snapshots**

Assert imported run snapshots match expected normalized colors. Build native equivalents with `slide.color`, write/reopen,
and compare final `richText` color/transparency arrays for all valid custom-default cases.

For clear/default controls, lock the intentional difference explicitly:

```ts
expect(pptxGenClearedColor).toEqual({ kind: 'srgb', value: '000000' });
expect(nativeClearedColor).toEqual({ kind: 'scheme', value: 'tx1' });
```

- [ ] **Step 4: Lock strict invalid-input correction**

Generate PptxGenJS public output with `slide.color = 'BAD'` while spying on its warning, and assert its final run becomes
black. Separately assert native setter throws immediately, does not add a shape, does not change its old default, and emits
no warning/fallback package.

- [ ] **Step 5: Add validator evidence**

Validate native sRGB/theme/override/alpha packages and valid PptxGenJS packages under PowerPoint 2010 and current profiles.
Expected: zero errors; no new warning code is introduced for transient state.

- [ ] **Step 6: Review, test, commit, and push conformance**

Run:

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/validator/src/index.test.ts -t "slide default"
pnpm typecheck
git diff --check
```

Commit and push:

```bash
git add packages/pptxgenjs-adapter/src/index.test.ts packages/validator/src/index.test.ts
git commit -m "test: compare slide default colors with pptxgenjs"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 4: Verify root package, actual tarball, browser, galleries, and clients

**Files:**

- Modify: `packages/pptx/src/index.test.ts`
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Consumes: Tasks 1-3 public API and materialized run semantics.
- Produces: root/runtime flag `slideDefaultColor: true`, declaration evidence, deterministic package manifest, browser state,
  validation/render/client evidence.

- [ ] **Step 1: Add root export/runtime test**

Import `type RichTextColor`, assign it through `slide.color`, create inherited/override runs, duplicate, write/reopen, and
assert materialized colors plus reopened transient `undefined`. Root test must use only `./index.js` exports.

- [ ] **Step 2: Extend the installed TypeScript consumer**

In `scripts/build-npm-package-types.mjs`, add:

```ts
const defaultColor: RichTextColor = { kind: 'scheme', value: 'accent1' };
slide.color = defaultColor;
const currentDefault: Readonly<RichTextColor> | undefined = slide.color;
```

Compile with the existing strict installed-consumer command. Include a `@ts-expect-error` proving string assignment is not
accepted.

- [ ] **Step 3: Extend actual npm tarball Node/CLI smoke**

Create sRGB/theme/override/duplicate cases from the installed package, write/reopen, and compute:

```js
const slideDefaultColor = /* exact state, materialized colors, duplicate, reopened undefined */;
```

Add it to `checks`, the final JSON summary, and the all-true assertion. Inspect/validate the generated deck using installed
CLI commands and include `slideDefaultColorInspect`, `slideDefaultColorValidate`, and slide-part-read booleans.

- [ ] **Step 4: Extend real-Chrome browser smoke**

Inside the browser module evaluation:

1. set theme default;
2. create inherited/local override/transparency runs;
3. duplicate and confirm duplicate transient state before write;
4. writeBlob/open;
5. return materialized color arrays and reopened transient values;
6. assert zero console/page/network errors.

Add the exact object to `expected`; avoid timing-dependent values.

- [ ] **Step 5: Run focused, full, performance, type, and build gates**

Run:

```bash
pnpm --config.verify-deps-before-run=false exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts packages/validator/src/index.test.ts packages/pptx/src/index.test.ts -t "slide default"
pnpm test
pnpm test:performance
pnpm typecheck
pnpm build
```

Record exact pass/skip counts in a temporary evidence note, not in source until the final docs task.

- [ ] **Step 6: Pack twice and verify deterministic manifests**

Use a fresh `/tmp/pptx-slide-default-color-pack-*` directory. Run the repository package build/pack smoke twice, compare
sorted `dist` path/size/SHA-256 manifests, and record tarball file count, dist file count, and manifest hash. Run the real-
Chrome smoke against the built browser export.

- [ ] **Step 7: Build and inspect native/PptxGenJS galleries**

Before PPTX inspection, read and follow `.codex/skills/pptx-inspect/SKILL.md`. Generate fresh native and independent
PptxGenJS galleries in `/tmp`, covering zero-input, sRGB, scheme, multi-paragraph, inherited/override, alpha, temporal
change, clear, duplicate/move, and table isolation.

Use `pptx-inspect` dry-run/read-only commands to record slide count, part/relationship counts, run colors, and PowerPoint
2010 validation. Render every slide at 180 DPI, run edge-contact checks, and visually inspect every page.

- [ ] **Step 8: Run factual LibreOffice and PowerPoint checks**

Save both galleries through LibreOffice 26.8, reopen with native library, verify slide order/text/materialized run colors,
and validate the saved packages. Attempt PowerPoint 16.112 open/save/PDF only with a fresh target; report success only if
an active presentation and output files are actually produced.

- [ ] **Step 9: Commit only source verification changes**

Review that no `/tmp` artifacts or generated PPTX/PDF/PNG files are staged. Run:

```bash
git diff --check
git status --short
```

Commit and push:

```bash
git add packages/pptx/src/index.test.ts scripts/build-npm-package-types.mjs scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git commit -m "test: verify packed slide default colors"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 5: Publish docs and close the specialty

**Files:**

- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**

- Consumes: final verified API, exact test/package/browser/gallery/client evidence from Tasks 1-4.
- Produces: public examples, lifecycle/persistence warning, compatibility-matrix status, and next roadmap item.

- [ ] **Step 1: Add concise public examples**

Document:

```ts
slide.color = { kind: 'scheme', value: 'accent1' };
slide.addText('Uses accent1');
slide.addRichText([{
  runs: [
    { text: 'Inherited' },
    { text: ' override', style: { color: { kind: 'srgb', value: '00AA00' } } },
  ],
}]);
slide.color = undefined;
```

State prominently that default applies only to subsequently created `addText()` / `addRichText()` runs, is not serialized
as slide state, does not recolor existing shapes/tables, and reopened getters are `undefined` while materialized colors remain.

- [ ] **Step 2: Update compatibility and progress matrices**

Move `slide.color` from unsupported to supported for direct transient state, strict sRGB/theme values, inheritance,
duplicate/move/delete, Node/browser/types/CLI, six formats, PptxGenJS valid output, and client verification. Keep
layout/master/placeholder theme inheritance in the next specialty.

Set next sequence to:

```text
master/layout/placeholder → advanced text → advanced table/tableToSlides → output/runtime helpers → peer-range full-suite audit
```

- [ ] **Step 3: Publish exact evidence only**

Write actual focused/full/performance counts, tarball/dist counts and SHA-256, deterministic build results, browser returned
state, gallery part/relationship counts, validation diagnostics, render/edge results, LibreOffice normalization, and factual
PowerPoint outcome. Do not reuse Slide Number numbers or claim a client pass without output evidence.

- [ ] **Step 4: Run final stale-language and diff review**

Run:

```bash
git diff --check
rg -n -i "next .*default color|default color.*next|下一.*default color|slide\.color.*未支持" README.md packages/pptx/README.md docs CHANGELOG.md
git diff -- README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

Every roadmap reference must say default color is complete and master/layout/placeholder is next.

- [ ] **Step 5: Commit, push, and verify 100% specialty closure**

```bash
git add README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --cached --check
git commit -m "docs: complete slide default color support"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only pre-existing `.pnpm-store/` remains untracked. Report Default Color 7/7 (100%),
completed/remaining items, exact commits, and automatically begin master/layout/placeholder.

---

## Plan Self-Review

- Spec coverage: public state, strict normalization, precedence, temporal semantics, non-target isolation, lifecycle,
  persistence boundary, six formats, PptxGenJS, actual package, browser, CLI, gallery, validation, clients, and docs each map to
  an explicit task.
- Placeholder scan: every implementation and test step contains concrete files, interfaces, assertions, commands, and expected
  results.
- Type consistency: all tasks use existing `RichTextColor`, public `SlideModel.color`, internal
  `normalizeRichTextColor(value, context)`, and internal renderer `defaultColor` consistently.
- Scope check: table/master/layout/placeholder and advanced text remain excluded; no unrelated refactor or schema extension is
  introduced.
