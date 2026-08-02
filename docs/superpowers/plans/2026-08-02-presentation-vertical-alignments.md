# Presentation Vertical Alignment Runtime Catalog Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with independent review, commit, push, fetch, and remote-divergence verification.

**Goal:** Expose the three native vertical text alignment tokens as one frozen runtime catalog and derive the existing `TextBoxVerticalAlignment` type from it across source, root, actual tarball Node/browser/types, CLI inspection, and real Chrome.

**Architecture:** Add `TEXT_VERTICAL_ALIGNMENTS` beside the existing type in `packages/model/src/text.ts`; keep all current normalizers and text-body/table-cell OOXML mappings unchanged. Existing model → SDK → aggregate root export chains publish the catalog automatically. Permanent tests compare PptxGenJS 4.0.1 public `AlignV` values while preserving native tuple naming and immutability.

**Tech Stack:** TypeScript strict mode, Vitest, model/SDK/root packages, PptxGenJS 4.0.1 public runtime, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, installed CLI, and real Google Chrome.

## Global Constraints

- Exact order is `top`, `middle`, `bottom`.
- `TEXT_VERTICAL_ALIGNMENTS` is a frozen readonly tuple created once at module initialization.
- `TextBoxVerticalAlignment` must be exactly `(typeof TEXT_VERTICAL_ALIGNMENTS)[number]` with no widened string or duplicated union.
- Do not add `PptxDocument.AlignV`, a generic `AlignV` export, enum-shaped object, setter, registry, or mutable alias.
- Do not change `normalizeTextBoxVerticalAlignment`, text-body/table-cell OOXML token mappings, document bytes, error messages, defaults, or fallback behavior.
- Existing text box, slide-number, table, and table-cell APIs retain the same three accepted values.
- SDK/root expose the model constant only through existing `export *` chains; do not add duplicate constants.
- PptxGenJS conformance uses only public constructor, `AlignV`, `addText()`, `addTable()`, and `write()`.
- Never stage `.pnpm-store/`, dist output, tarballs, browser artifacts, temporary consumers, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Add the catalog, type derivation, public exports, and PptxGenJS conformance

**Files:**
- Modify: `packages/model/src/text.test.ts`
- Modify: `packages/model/src/text.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Produces `TEXT_VERTICAL_ALIGNMENTS: readonly ['top', 'middle', 'bottom']`.
- Preserves `TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom'` through derived typing.
- Consumes existing model → SDK → root export chains without facade changes.

- [ ] **Step 1: Write the failing model catalog tests**

Extend `packages/model/src/text.test.ts` with:

```ts
import {
  TEXT_ALIGNMENTS,
  TEXT_VERTICAL_ALIGNMENTS,
  type TextAlignment,
  type TextBoxVerticalAlignment,
} from './text.js';

describe('TEXT_VERTICAL_ALIGNMENTS', () => {
  it('publishes the complete frozen vertical alignment catalog', () => {
    expect(TEXT_VERTICAL_ALIGNMENTS).toEqual(['top', 'middle', 'bottom']);
    expect(new Set(TEXT_VERTICAL_ALIGNMENTS)).toHaveLength(3);
    expect(Object.isFrozen(TEXT_VERTICAL_ALIGNMENTS)).toBe(true);
    expect(() => {
      (TEXT_VERTICAL_ALIGNMENTS as unknown as string[]).push('top');
    }).toThrow(TypeError);
    expect(TEXT_VERTICAL_ALIGNMENTS).toEqual(['top', 'middle', 'bottom']);
  });

  it('keeps the runtime catalog and TextBoxVerticalAlignment synchronized', () => {
    const values: readonly TextBoxVerticalAlignment[] =
      TEXT_VERTICAL_ALIGNMENTS;
    expect(values).toBe(TEXT_VERTICAL_ALIGNMENTS);
    if (false) {
      // @ts-expect-error unknown vertical alignment is not supported
      const invalid: TextBoxVerticalAlignment = 'distributed';
      void invalid;
    }
  });
});
```

- [ ] **Step 2: Add failing SDK lifecycle coverage**

Import `TEXT_VERTICAL_ALIGNMENTS` and `TextBoxVerticalAlignment` in `packages/sdk/src/index.test.ts`. Add one test that creates three text shapes and one three-cell table from the catalog, checks immediate snapshots, writes/reopens, and requires both owner types to preserve the exact ordered values:

```ts
const document = PptxDocument.create();
const slide = document.addSlide();
const shapes = TEXT_VERTICAL_ALIGNMENTS.map((alignment) => {
  const typed: TextBoxVerticalAlignment = alignment;
  return slide.addText(typed, { valign: typed });
});
const table = slide.addTable([
  TEXT_VERTICAL_ALIGNMENTS.map((alignment) => ({
    text: alignment,
    options: { valign: alignment },
  })),
]);

expect(shapes.map(({ verticalAlignment }) => verticalAlignment))
  .toEqual(TEXT_VERTICAL_ALIGNMENTS);
expect(table.rows[0]?.cells.map(({ verticalAlignment }) => verticalAlignment))
  .toEqual(TEXT_VERTICAL_ALIGNMENTS);
```

After reopen, identify the three text shapes and table through their stable public collection order and require the same two arrays. Capture `[...document.opcPackage.mutations]` before catalog reads and prove that reading/iterating the tuple alone changes no package state.

- [ ] **Step 3: Add failing aggregate-root export coverage**

In `packages/pptx/src/index.test.ts`, import the SDK value as `SDK_TEXT_VERTICAL_ALIGNMENTS` and the aggregate-root value/type normally. Add a test requiring:

```ts
const values: readonly TextBoxVerticalAlignment[] = TEXT_VERTICAL_ALIGNMENTS;
expect(TEXT_VERTICAL_ALIGNMENTS).toBe(SDK_TEXT_VERTICAL_ALIGNMENTS);
expect(values).toBe(TEXT_VERTICAL_ALIGNMENTS);
expect(values).toEqual(['top', 'middle', 'bottom']);
expect(Object.isFrozen(TEXT_VERTICAL_ALIGNMENTS)).toBe(true);
```

Include an unreachable `@ts-expect-error` assignment for `'distributed'` so the root declaration surface is checked by TypeScript.

- [ ] **Step 4: Add failing PptxGenJS public comparison**

Extend the test-only `PptxGenJSInstance` in `packages/pptxgenjs-adapter/src/index.test.ts` with:

```ts
readonly AlignV: Readonly<Record<string, string>>;
```

Add a test requiring `Object.keys(generated.AlignV)` and `Object.values(generated.AlignV)` to equal `TEXT_VERTICAL_ALIGNMENTS` in order. Use all three public PptxGenJS values as `valign` for three text boxes and three table cells, call public `write()`, import with `importPptxGenJS()`, and require native shape/table snapshots to preserve `top`, `middle`, `bottom`. Do not mutate or inspect PptxGenJS private fields.

- [ ] **Step 5: Run RED**

```sh
node_modules/.bin/vitest run packages/model/src/text.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "TEXT_VERTICAL_ALIGNMENTS|vertical alignment runtime catalog" --reporter=dot
```

Expected: FAIL because `TEXT_VERTICAL_ALIGNMENTS` does not exist.

- [ ] **Step 6: Implement the single source of truth**

Replace the handwritten union in `packages/model/src/text.ts` with:

```ts
export const TEXT_VERTICAL_ALIGNMENTS = Object.freeze([
  'top',
  'middle',
  'bottom',
] as const);

export type TextBoxVerticalAlignment =
  typeof TEXT_VERTICAL_ALIGNMENTS[number];
```

Do not modify model index, SDK index, root index, normalizers, codecs, or OOXML mappings; the existing export chains are sufficient.

- [ ] **Step 7: Run focused, package, and type gates**

```sh
node_modules/.bin/vitest run packages/model/src/text.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "TEXT_VERTICAL_ALIGNMENTS|vertical alignment runtime catalog" --reporter=dot
node_modules/.bin/vitest run packages/model/src/text.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 8: Review, commit, push, and verify**

Review exact ordering, freeze/mutation behavior, derived union, zero facade duplication, unchanged text/table OOXML lifecycle, root closure, public-only PptxGenJS comparison, and no unrelated mapping changes. Stage only the five listed files, commit `feat: expose vertical alignment catalog`, push, fetch, and require divergence `0 0`.

---

### Task 2: Actual-tarball Node, browser, declarations, TypeScript, CLI, and Chrome proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes installed `TEXT_VERTICAL_ALIGNMENTS` and `TextBoxVerticalAlignment`.
- Produces stable `verticalAlignments: true` plus `verticalAlignmentState` in packed Node and real Chrome output.

- [ ] **Step 1: Extend packed declaration checks**

Read `dist/types/model/text.d.ts` and require both declarations with exact token order and derived typing:

```text
export declare const TEXT_VERTICAL_ALIGNMENTS: readonly ['top', 'middle', 'bottom'];
export type TextBoxVerticalAlignment = typeof TEXT_VERTICAL_ALIGNMENTS[number];
```

Accept the declaration generator's quote style but require readonly tuple semantics. Keep the root declaration re-export check unchanged.

- [ ] **Step 2: Extend installed Node and browser-condition smoke**

Add `TEXT_VERTICAL_ALIGNMENTS` to generated consumer imports. Build `verticalAlignmentState` from a fresh same-package dynamic import and real create/reopen snapshots:

```ts
const verticalAlignmentState = {
  values: [...TEXT_VERTICAL_ALIGNMENTS],
  textReopened: reopenedTextShapes.map(
    ({ verticalAlignment }) => verticalAlignment,
  ),
  tableReopened: reopenedTable.rows[0]?.cells.map(
    ({ verticalAlignment }) => verticalAlignment,
  ),
  frozen: Object.isFrozen(TEXT_VERTICAL_ALIGNMENTS),
  shared: TEXT_VERTICAL_ALIGNMENTS === reimportedVerticalAlignments,
};
```

Set `verticalAlignments` true only when all value arrays equal `['top', 'middle', 'bottom']` and `frozen/shared` are true. Add both fields to stable packed JSON. The generated deck must still pass installed CLI package inspection.

- [ ] **Step 3: Extend installed TypeScript consumer**

Import runtime and type, then add:

```ts
const verticalAlignments: readonly TextBoxVerticalAlignment[] =
  TEXT_VERTICAL_ALIGNMENTS;
for (const alignment of verticalAlignments) {
  alignment satisfies TextBoxVerticalAlignment;
}
// @ts-expect-error runtime catalog is readonly
TEXT_VERTICAL_ALIGNMENTS.push('top');
// @ts-expect-error catalog index is readonly
TEXT_VERTICAL_ALIGNMENTS[0] = 'bottom';
// @ts-expect-error unknown alignment is rejected
const invalidVerticalAlignment: TextBoxVerticalAlignment = 'distributed';
```

Run the installed consumer under its existing strict TypeScript configuration.

- [ ] **Step 4: Extend real Chrome coverage**

In `scripts/playwright-browser-smoke.js`, read `api.TEXT_VERTICAL_ALIGNMENTS`, require exact values and `Object.isFrozen()`, create/writeBlob/reopen three vertically aligned text shapes and a three-cell table, and return:

```ts
verticalAlignments: true,
verticalAlignmentState: {
  values: ['top', 'middle', 'bottom'],
  textReopened: ['top', 'middle', 'bottom'],
  tableReopened: ['top', 'middle', 'bottom'],
  frozen: true,
},
```

Chrome validation/console/page/network errors must remain zero.

- [ ] **Step 5: Build, pack, and run installed gates**

```sh
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
(cd packages/pptx && npm pack --ignore-scripts)
node scripts/smoke-npm-package.mjs packages/pptx/jiayunxie-pptx-0.1.0.tgz
```

Run the repository browser smoke callback in installed Google Chrome against the built browser module. Record actual tarball file count and SHA-256, then move tarball/browser artifacts out of the workspace and confirm no generated tracked changes.

- [ ] **Step 6: Review, commit, push, and verify**

Review installed-only imports, declaration exactness, tuple immutability, Node/browser text and table round-trip, TypeScript negatives, CLI inspection, real Chrome zero-error state, stable JSON, and clean artifact scope. Stage only the two listed scripts, commit `test: verify packed vertical alignments`, push, fetch, and require divergence `0 0`.

---

### Task 3: Full release gates, compatibility status, and documentation

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes verified Task 1–2 behavior and actual tarball/Chrome evidence.
- Produces the final supported/remaining status without claiming exact PptxGenJS namespace shape or unfinished output/runtime helpers.

- [ ] **Step 1: Run complete release gates**

```sh
node_modules/.bin/vitest run --reporter=dot
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
```

The chart catalog test already has its explicit 15-second budget; any new failure must be reproduced and resolved, not hidden by a global timeout change.

- [ ] **Step 2: Update exactly the six release documents**

Document all of the following:

- `TEXT_VERTICAL_ALIGNMENTS` exact values/order/frozen state and derived `TextBoxVerticalAlignment`;
- existing text-box/slide-number/table use and unchanged `t`/`ctr`/`b` OOXML mappings;
- native tuple naming versus PptxGenJS instance `AlignV` object;
- deliberate omission of instance getter and mutable enum alias;
- source/root, packed Node/browser/types, installed CLI, and real Chrome evidence;
- final tests, performance, tarball file count/SHA, and Chrome error counts;
- remaining output/runtime helpers, advanced text/table, `tableToSlides`, and final audit.

- [ ] **Step 3: Search stale status and self-review**

Search non-plan docs for `AlignV` still marked unsupported. Require exactly the six intended files modified, balanced Markdown fences, no `TODO`/`TBD`/`FIXME`, `git diff --check`, no `.pnpm-store/`, and no generated artifacts staged.

- [ ] **Step 4: Commit, push, and verify**

Stage only the six listed documents, commit `docs: document vertical alignment catalog`, push, fetch, and require divergence `0 0`.

- [ ] **Step 5: Synchronize progress and continue**

Report design/plan plus all three implementation tasks complete, update overall parity progress from measured coverage, list remaining gaps, and select the next smallest independent output/runtime helper item without waiting for another instruction.
