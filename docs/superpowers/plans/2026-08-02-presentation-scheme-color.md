# Presentation `SchemeColor` Runtime Helper Implementation Plan

> **For agentic workers:** Execute this plan task-by-task inline in the current session. The current user direction prohibits subagents; review, commit, push, and verify each task before starting its successor.

**Goal:** Add an immutable native runtime equivalent of PptxGenJS 4.0.1's ten-entry `SchemeColor` helper, prove it through source and actual-package Node/browser/Chrome evidence, and move the helper from pending to supported.

**Architecture:** `@pptx/model` owns one frozen key-to-value mapping and a union derived from its values. Existing `export *` layers expose the same object through `@pptx/sdk` and `@jiayunxie/pptx`; no document getter, extra facade, serializer change, or `RichTextColor` narrowing is introduced. Tests compare the exact upstream helper, exercise existing theme-color serialization, and verify catalog/package isolation.

**Tech Stack:** TypeScript 5.8, Vitest, PptxGenJS 4.0.1, tsup, packed declaration smoke, Playwright with installed Google Chrome, `pptx-inspect` CLI.

## Global Constraints

- Exact public order is `text1`, `text2`, `background1`, `background2`, `accent1`, `accent2`, `accent3`, `accent4`, `accent5`, `accent6`.
- Exact values are `tx1`, `tx2`, `bg1`, `bg2`, `accent1`, `accent2`, `accent3`, `accent4`, `accent5`, `accent6`.
- The public name is `SCHEME_COLORS`; its derived value union is `SchemeColor`.
- `SCHEME_COLORS` is frozen, model-owned, environment-independent, and shared by identity through model, SDK, and root exports.
- Do not add `PptxDocument.SchemeColor`, a class static, a TypeScript enum, a mutable alias, or a second catalog copy.
- Do not narrow `RichTextColor`, consolidate the existing seventeen-token validators, or alter existing OOXML output.
- Reading or attempting to mutate the catalog must not change a document, diagnostics, or the package mutation journal.
- Every task ends with review, a scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage or commit `.pnpm-store/` or generated workspace tarballs.

---

## File map

- Create `packages/model/src/scheme-color.ts`: sole runtime catalog and derived value union.
- Create `packages/model/src/scheme-color.test.ts`: exact value, immutability, and compile-time contract.
- Modify `packages/model/src/index.ts`: public model export.
- Modify `packages/pptx/src/index.test.ts`: SDK/root identity, package isolation, real color use, and root type closure.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: real PptxGenJS 4.0.1 helper and legal-output conformance.
- Modify `scripts/build-npm-package-types.mjs`: require the packed public declaration.
- Modify `scripts/smoke-npm-package.mjs`: packed declaration, Node, browser-condition, TypeScript, and CLI evidence.
- Modify `scripts/playwright-browser-smoke.js`: real-Chrome immutable catalog and write/reopen evidence.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, `docs/implementation-progress.md`, and `CHANGELOG.md`: public usage, deliberate divergence, release evidence, and remaining work.

---

### Task 1: Core catalog, types, exports, and conformance

**Files:**
- Create: `packages/model/src/scheme-color.ts`
- Create: `packages/model/src/scheme-color.test.ts`
- Modify: `packages/model/src/index.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: existing model -> SDK -> aggregate-root `export *` chain, current `RichTextColor` APIs, `PptxDocument`, `ShapeModel`, `importPptxGenJS()`, and the public PptxGenJS 4.0.1 `SchemeColor` getter.
- Produces: `SCHEME_COLORS: Readonly<{ text1: 'tx1'; text2: 'tx2'; background1: 'bg1'; background2: 'bg2'; accent1: 'accent1'; accent2: 'accent2'; accent3: 'accent3'; accent4: 'accent4'; accent5: 'accent5'; accent6: 'accent6' }>` and `SchemeColor`, available through model, SDK, and root with one object identity.

- [ ] **Step 1: Add the failing model contract test**

Create `packages/model/src/scheme-color.test.ts` with exact runtime and type expectations:

```ts
import { describe, expect, it } from 'vitest';
import { SCHEME_COLORS, type SchemeColor } from './scheme-color.js';

const EXPECTED = {
  text1: 'tx1',
  text2: 'tx2',
  background1: 'bg1',
  background2: 'bg2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
} as const;

describe('SCHEME_COLORS', () => {
  it('publishes the complete frozen PptxGenJS helper mapping', () => {
    expect(Object.entries(SCHEME_COLORS)).toEqual(Object.entries(EXPECTED));
    expect(new Set(Object.values(SCHEME_COLORS))).toHaveLength(10);
    expect(Object.isFrozen(SCHEME_COLORS)).toBe(true);
    expect(() => {
      (SCHEME_COLORS as { accent1: string }).accent1 = 'changed';
    }).toThrow(TypeError);
    expect(() => Object.defineProperty(SCHEME_COLORS, 'accent7', {
      value: 'accent7',
    })).toThrow(TypeError);
    expect(() => {
      delete (SCHEME_COLORS as { accent6?: string }).accent6;
    }).toThrow(TypeError);
    expect(SCHEME_COLORS).toEqual(EXPECTED);
  });

  it('derives SchemeColor from the helper values', () => {
    const values: readonly SchemeColor[] = Object.values(SCHEME_COLORS);
    expect(values).toEqual(Object.values(EXPECTED));
    if (false) {
      // @ts-expect-error helper key is not a helper value
      const key: SchemeColor = 'text1';
      // @ts-expect-error extended DrawingML token is outside this PptxGenJS helper
      const extended: SchemeColor = 'hlink';
      // @ts-expect-error raw sRGB is not a scheme helper value
      const srgb: SchemeColor = 'FF0000';
      // @ts-expect-error helper keys are closed
      SCHEME_COLORS.text3;
      void [key, extended, srgb];
    }
  });
});
```

- [ ] **Step 2: Add failing root identity and lifecycle coverage**

In `packages/pptx/src/index.test.ts`, import `SCHEME_COLORS as MODEL_SCHEME_COLORS` from `@pptx/model`, import `SCHEME_COLORS as SDK_SCHEME_COLORS` from `@pptx/sdk`, and import root `SCHEME_COLORS` plus `SchemeColor`. Add a stable-export test that:

```ts
it('exports the frozen SCHEME_COLORS helper from the root package', async () => {
  expect(SCHEME_COLORS).toBe(MODEL_SCHEME_COLORS);
  expect(SCHEME_COLORS).toBe(SDK_SCHEME_COLORS);
  expect(Object.entries(SCHEME_COLORS)).toEqual([
    ['text1', 'tx1'],
    ['text2', 'tx2'],
    ['background1', 'bg1'],
    ['background2', 'bg2'],
    ['accent1', 'accent1'],
    ['accent2', 'accent2'],
    ['accent3', 'accent3'],
    ['accent4', 'accent4'],
    ['accent5', 'accent5'],
    ['accent6', 'accent6'],
  ]);
  expect(Object.isFrozen(SCHEME_COLORS)).toBe(true);

  const isolated = PptxDocument.create();
  const journal = JSON.stringify(isolated.opcPackage.mutations);
  Object.values(SCHEME_COLORS);
  expect(JSON.stringify(isolated.opcPackage.mutations)).toBe(journal);

  const document = PptxDocument.create();
  document.addSlide().addText('Scheme helper', {
    color: { kind: 'scheme', value: SCHEME_COLORS.text1 },
    fill: { kind: 'solid', color: { kind: 'scheme', value: SCHEME_COLORS.accent1 } },
  });
  const reopened = await PptxDocument.open(await document.write());
  const shape = reopened.slides[0]?.shapes[0];
  expect(shape).toBeInstanceOf(ShapeModel);
  expect((shape as ShapeModel).richText[0]?.runs[0]?.style.color)
    .toEqual({ kind: 'scheme', value: 'tx1' });
  expect((shape as ShapeModel).fill)
    .toEqual({ kind: 'solid', color: { kind: 'scheme', value: 'accent1' } });

  if (false) {
    const text: SchemeColor = SCHEME_COLORS.text1;
    const accent: SchemeColor = SCHEME_COLORS.accent6;
    // @ts-expect-error SchemeColor excludes key labels
    const invalid: SchemeColor = 'background1';
    // @ts-expect-error the runtime catalog is readonly
    SCHEME_COLORS.accent1 = 'accent2';
    void [text, accent, invalid];
  }
});
```

- [ ] **Step 3: Add failing real-PptxGenJS conformance**

In `packages/pptxgenjs-adapter/src/index.test.ts`:

1. Import `SCHEME_COLORS` from `@pptx/sdk`.
2. Expand `PptxGenJSInstance['SchemeColor']` to all ten exact properties.
3. Add a test beside the existing `AlignV`/layout/runtime-helper tests:

```ts
it('matches the public PptxGenJS SchemeColor helper and legal output', async () => {
  const generated = new PptxGenJS();
  const second = new PptxGenJS();
  expect(Object.entries(generated.SchemeColor)).toEqual(Object.entries(SCHEME_COLORS));
  expect(second.SchemeColor).toBe(generated.SchemeColor);
  expect(Object.isFrozen(generated.SchemeColor)).toBe(false);
  expect(Object.getOwnPropertyDescriptor(
    Object.getPrototypeOf(generated),
    'SchemeColor',
  )).toMatchObject({ set: undefined, enumerable: false });
  expect(Object.isFrozen(SCHEME_COLORS)).toBe(true);

  const slide = generated.addSlide();
  slide.addText([
    { text: 'Text1', options: { color: generated.SchemeColor.text1 } },
    { text: 'Accent1', options: { color: generated.SchemeColor.accent1 } },
  ], { x: 1, y: 1, w: 5, h: 1 });
  const imported = await importPptxGenJS(generated);
  const shape = imported.slides[0]?.shapes[0];
  expect(shape).toBeInstanceOf(ShapeModel);
  expect((shape as ShapeModel).richText[0]?.runs.map(({ style }) => style.color)).toEqual([
    { kind: 'scheme', value: SCHEME_COLORS.text1 },
    { kind: 'scheme', value: SCHEME_COLORS.accent1 },
  ]);
});
```

- [ ] **Step 4: Run RED and require missing-module/export failures**

Run:

```sh
node_modules/.bin/vitest run packages/model/src/scheme-color.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
```

Expected: FAIL because `packages/model/src/scheme-color.ts`, `SCHEME_COLORS`, and `SchemeColor` do not exist. Do not weaken any expectation.

- [ ] **Step 5: Implement the model-owned catalog**

Create `packages/model/src/scheme-color.ts`:

```ts
export const SCHEME_COLORS = Object.freeze({
  text1: 'tx1',
  text2: 'tx2',
  background1: 'bg1',
  background2: 'bg2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
} as const);

export type SchemeColor = (typeof SCHEME_COLORS)[keyof typeof SCHEME_COLORS];
```

Add exactly one line to `packages/model/src/index.ts`:

```ts
export * from './scheme-color.js';
```

Do not modify `RichTextColor`, internal scheme-color sets, SDK forwarding code, or `packages/pptx/src/index.ts`; the current export chain must provide the public surface.

- [ ] **Step 6: Run focused GREEN, type, and bundle gates**

Run:

```sh
node_modules/.bin/vitest run packages/model/src/scheme-color.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
```

Require all commands to pass. Inspect `packages/pptx/dist/types/model/scheme-color.d.ts` and require the ten readonly literal properties plus the derived `SchemeColor` type. Confirm `packages/pptx/dist/browser.js` has no static `node:` import.

- [ ] **Step 7: Run complete regression and performance gates**

Run:

```sh
node_modules/.bin/vitest run --reporter=dot
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Record exact test totals and elapsed performance. The performance case must remain below its existing 5-second budget.

- [ ] **Step 8: Review, commit, push, and verify**

Review exact key/value order, ten unique values, freeze semantics, upstream descriptor evidence, model/SDK/root identity, no package mutation on reads, real `tx1`/`accent1` write/reopen, no `RichTextColor` narrowing, and no generated artifact scope.

Stage only the five core files, commit, push, fetch, and verify:

```sh
git add packages/model/src/scheme-color.ts packages/model/src/scheme-color.test.ts \
  packages/model/src/index.ts packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "feat: add presentation scheme color helper"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 2: Actual-package Node, declarations, browser, CLI, and Chrome proof

**Files:**
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: Task 1 `SCHEME_COLORS`, `SchemeColor`, the packed model declaration, existing stable smoke JSON, and the actual browser conditional export.
- Produces: stable `schemeColors: true` and `schemeColorState` evidence in installed Node and Chrome, plus declaration/type/browser-condition/CLI proof.

- [ ] **Step 1: Lock the packed declaration**

In `scripts/build-npm-package-types.mjs`, add `model/scheme-color.d.ts` to `requiredPublicDeclarations` with required names `SCHEME_COLORS` and `SchemeColor`.

In `scripts/smoke-npm-package.mjs`, read and whitespace-normalize the same declaration. Require all ten property/value fragments and:

```ts
export type SchemeColor = (typeof SCHEME_COLORS)[keyof typeof SCHEME_COLORS];
```

Also require the packed model index to export `./scheme-color.js` and the SDK declaration to preserve its model re-export.

- [ ] **Step 2: Extend installed Node runtime smoke**

Add `SCHEME_COLORS` and `SchemeColor` to the generated consumer imports. Build this stable Node state:

```ts
const schemeColorEntries = Object.entries(SCHEME_COLORS);
const schemeColorIsolationDocument = PptxDocument.create();
const schemeColorJournal = JSON.stringify(schemeColorIsolationDocument.opcPackage.mutations);
Object.values(SCHEME_COLORS);
const schemeColorDocument = PptxDocument.create();
schemeColorDocument.addSlide().addText('Packed scheme colors', {
  color: { kind: 'scheme', value: SCHEME_COLORS.text1 },
  fill: { kind: 'solid', color: { kind: 'scheme', value: SCHEME_COLORS.accent1 } },
});
const schemeColorReopened = await PptxDocument.open(await schemeColorDocument.write());
const schemeColorShape = schemeColorReopened.slides[0].shapes[0];
const schemeColorState = {
  entries: schemeColorEntries,
  frozen: Object.isFrozen(SCHEME_COLORS),
  mutationIsolation: JSON.stringify(
    schemeColorIsolationDocument.opcPackage.mutations,
  ) === schemeColorJournal,
  textColor: schemeColorShape instanceof ShapeModel
    ? schemeColorShape.richText[0]?.runs[0]?.style.color
    : undefined,
  fill: schemeColorShape instanceof ShapeModel ? schemeColorShape.fill : undefined,
};
const schemeColors = JSON.stringify(schemeColorState) === JSON.stringify({
  entries: [
    ['text1', 'tx1'], ['text2', 'tx2'], ['background1', 'bg1'],
    ['background2', 'bg2'], ['accent1', 'accent1'], ['accent2', 'accent2'],
    ['accent3', 'accent3'], ['accent4', 'accent4'], ['accent5', 'accent5'],
    ['accent6', 'accent6'],
  ],
  frozen: true,
  mutationIsolation: true,
  textColor: { kind: 'scheme', value: 'tx1' },
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
});
```

Add both fields to the existing stable checks object and fail with detailed state when `schemeColors` is false.

- [ ] **Step 3: Extend installed TypeScript and browser-condition consumers**

In the generated TypeScript fixture, require:

```ts
const typedTextScheme: SchemeColor = SCHEME_COLORS.text1;
const typedAccentScheme: SchemeColor = SCHEME_COLORS.accent6;
// @ts-expect-error key labels are not SchemeColor values
const invalidSchemeKey: SchemeColor = 'text1';
// @ts-expect-error extended DrawingML values are not in the PptxGenJS helper
const invalidHelperValue: SchemeColor = 'hlink';
// @ts-expect-error the mapping is readonly
SCHEME_COLORS.accent1 = 'accent2';
```

In the browser-condition consumer, require exact entries, frozen state, create/writeBlob/reopen of `tx1` and `accent1`, and no mutation of a separate document. Keep the existing conditional-export check and all prior runtime helper checks intact.

- [ ] **Step 4: Extend the real-Chrome callback**

Near the other runtime-helper checks in `scripts/playwright-browser-smoke.js`, construct `schemeColorState` with exact entries, frozen state, catalog mutation isolation, `text1` and `accent1` use, `writeBlob()`/reopen state, and zero validation errors. Add `schemeColors` and `schemeColorState` to both actual and expected JSON. Do not mutate the frozen catalog merely to prove it is frozen; `Object.isFrozen()` plus existing source tests cover rejection without creating a browser console error.

- [ ] **Step 5: Run syntax and actual-package gates**

Run:

```sh
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
git diff --check
node scripts/build-npm-package-types.mjs
(cd packages/pptx && npm pack --ignore-scripts)
node scripts/smoke-npm-package.mjs packages/pptx/jiayunxie-pptx-0.1.0.tgz
```

Require installed Node/types/browser-condition/CLI to report `schemeColors: true` without changing any prior smoke field.

- [ ] **Step 6: Run installed Google Chrome against the extracted tarball**

Extract the actual tarball under a fresh `/tmp/pptx-scheme-color-artifacts.XXXXXX` site that maps the packed `dist/browser.js` to `/packages/pptx/dist/browser.js`. Serve it over loopback, run `scripts/playwright-browser-smoke.js` with the bundled Playwright runtime and installed Google Chrome, and retain compact result JSON.

Require:

```json
{
  "schemeColors": true,
  "schemeColorState": {
    "entries": [
      ["text1", "tx1"],
      ["text2", "tx2"],
      ["background1", "bg1"],
      ["background2", "bg2"],
      ["accent1", "accent1"],
      ["accent2", "accent2"],
      ["accent3", "accent3"],
      ["accent4", "accent4"],
      ["accent5", "accent5"],
      ["accent6", "accent6"]
    ],
    "frozen": true,
    "mutationIsolation": true,
    "textColor": { "kind": "scheme", "value": "tx1" },
    "fill": { "kind": "solid", "color": { "kind": "scheme", "value": "accent1" } },
    "validationErrors": 0
  },
  "errorCounts": { "console": 0, "page": 0, "network": 0 }
}
```

Record the tarball file count and SHA-256. Move the tarball out of the workspace into the evidence directory.

- [ ] **Step 7: Review, commit, push, and verify**

Review actual-package-only imports, exact packed declarations, stable JSON, prior-field preservation, root availability in Node and browser, real write/reopen, catalog/package isolation, HTTP success, and zero Chrome errors.

Stage only the three scripts, commit, push, fetch, and require divergence `0 0`:

```sh
git add scripts/build-npm-package-types.mjs scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js
git commit -m "test: verify packed scheme color helper"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 3: Release documentation and progress closeout

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: final source, upstream conformance, full tests, performance, actual-tarball, Node/browser/CLI, and Chrome evidence from Tasks 1–2.
- Produces: public helper examples, precise immutable-root divergence, supported status, final evidence, and the next runtime-helper item.

- [ ] **Step 1: Document public use and scope**

Add examples using `SCHEME_COLORS.text1`, `SCHEME_COLORS.background1`, and `SCHEME_COLORS.accent1`, plus a `SchemeColor` annotation. State exact key/value order, frozen model/SDK/root identity, and environment-independent discovery.

State explicitly that the helper contains the ten PptxGenJS public values while existing native color APIs continue to accept their wider validated DrawingML subset. Do not claim that `SchemeColor` is the exhaustive native OOXML scheme-color type.

- [ ] **Step 2: Document compatibility and deliberate divergence**

Move `SchemeColor` runtime helper from pending to supported in the compatibility matrix. Record exact PptxGenJS 4.0.1 keys/values/order and legal-output import. Explain that native uses a frozen root mapping instead of the upstream prototype getter and shared mutable enum object.

- [ ] **Step 3: Close progress with actual evidence**

In `docs/implementation-progress.md`, add a completed section with exact focused/full test totals, performance time, tarball file count/SHA-256, packed/Chrome state, zero Chrome errors, commit boundaries, and retained `/tmp` evidence path. Name the next smallest runtime-helper item based on the remaining compatibility audit; do not raise the approximately 97% figure without quantified justification.

Add matching concise changelog bullets and update the two READMEs plus API reference with the same facts.

- [ ] **Step 4: Review documentation consistency**

Run:

```sh
rg -n 'SCHEME_COLORS|SchemeColor|schemeColors|runtime helper|PptxGenJS|remaining|剩余' \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --check
```

Require no current section to list `SchemeColor` as pending. Historical checkpoint language may remain only when it explicitly points forward to the completed section.

- [ ] **Step 5: Commit, push, verify, and report**

Stage only the six documentation files, commit, push, fetch, and require divergence `0 0`:

```sh
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: document presentation scheme color helper"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Report the completed helper, exact evidence, remaining helper/features, and unchanged or justified overall progress. Immediately begin the next selected runtime-helper design under the same one-item review/commit/push discipline.
