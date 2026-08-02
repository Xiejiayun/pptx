# Presentation Horizontal Alignment Runtime Catalog Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with independent review, commit, push, fetch, and remote-divergence verification.

**Goal:** Expose the four native horizontal text alignment tokens as one frozen runtime catalog and derive the existing `TextAlignment` type from it across source, root, actual tarball Node/browser/types, CLI inspection, and real Chrome.

**Architecture:** Add `TEXT_ALIGNMENTS` beside the existing type in `packages/model/src/text.ts`; keep all current normalizers and OOXML mappings unchanged. Existing model → SDK → aggregate root export chains publish the catalog automatically. Permanent tests compare PptxGenJS 4.0.1 public `AlignH` values while preserving native tuple naming and immutability.

**Tech Stack:** TypeScript strict mode, Vitest, model/SDK/root packages, PptxGenJS 4.0.1 public runtime, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, installed CLI, and real Google Chrome.

## Global Constraints

- Exact order is `left`, `center`, `right`, `justify`.
- `TEXT_ALIGNMENTS` is a frozen readonly tuple created once at module initialization.
- `TextAlignment` must be exactly `(typeof TEXT_ALIGNMENTS)[number]` with no widened string or duplicated union.
- Do not add `PptxDocument.AlignH`, a generic `AlignH` export, enum-shaped object, setter, registry, or mutable alias.
- Do not change `normalizeTextAlignment`, OOXML token mappings, document bytes, error messages, or fallback behavior.
- Existing plain/rich text and table/table-cell APIs retain the same four accepted values.
- SDK/root expose the model constant only through existing `export *` chains; do not add duplicate constants.
- PptxGenJS conformance uses only public constructor and `AlignH` property.
- Never stage `.pnpm-store/`, dist output, tarballs, browser artifacts, temporary consumers, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Add the catalog, type derivation, public exports, and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/text.test.ts`
- Modify: `packages/model/src/text.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Produces `TEXT_ALIGNMENTS: readonly ['left', 'center', 'right', 'justify']`.
- Preserves `TextAlignment = 'left' | 'center' | 'right' | 'justify'` through derived typing.
- Consumes existing model → SDK → root export chains without facade changes.

- [ ] **Step 1: Write the failing model catalog test**

Create `packages/model/src/text.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { TEXT_ALIGNMENTS, type TextAlignment } from './text.js';

describe('TEXT_ALIGNMENTS', () => {
  it('publishes the complete frozen horizontal alignment catalog', () => {
    expect(TEXT_ALIGNMENTS).toEqual(['left', 'center', 'right', 'justify']);
    expect(new Set(TEXT_ALIGNMENTS)).toHaveLength(4);
    expect(Object.isFrozen(TEXT_ALIGNMENTS)).toBe(true);
    expect(() => {
      (TEXT_ALIGNMENTS as unknown as string[]).push('left');
    }).toThrow(TypeError);
    expect(TEXT_ALIGNMENTS).toEqual(['left', 'center', 'right', 'justify']);
  });

  it('keeps the runtime catalog and TextAlignment union synchronized', () => {
    const values: readonly TextAlignment[] = TEXT_ALIGNMENTS;
    expect(values).toBe(TEXT_ALIGNMENTS);
    if (false) {
      // @ts-expect-error unknown horizontal alignment is not supported
      const invalid: TextAlignment = 'distributed';
      void invalid;
    }
  });
});
```

- [ ] **Step 2: Add failing SDK/root lifecycle tests**

Import `TEXT_ALIGNMENTS` and `TextAlignment` in SDK/root tests. In SDK, create one text shape for each token, write/reopen, and require paragraph snapshots to retain the four values in order:

```ts
const document = PptxDocument.create();
const slide = document.addSlide();
for (const align of TEXT_ALIGNMENTS) {
  const typed: TextAlignment = align;
  slide.addText(typed, { align: typed });
}
expect(slide.shapes.map(({ richText }) => richText[0]?.align))
  .toEqual(TEXT_ALIGNMENTS);
const reopened = await PptxDocument.open(await document.write());
expect(reopened.slides[0]?.shapes.map(({ richText }) => richText[0]?.align))
  .toEqual(TEXT_ALIGNMENTS);
```

In root tests, require exact values, frozen state, and tuple iteration through `TextAlignment`. This proves both runtime and declarations close through `@jiayunxie/pptx`.

- [ ] **Step 3: Add failing PptxGenJS public comparison**

Extend test-only `PptxGenJSInstance` with:

```ts
readonly AlignH: Readonly<Record<string, string>>;
```

Add a test requiring:

```ts
const generated = new PptxGenJS();
expect(Object.keys(generated.AlignH)).toEqual(TEXT_ALIGNMENTS);
expect(Object.values(generated.AlignH)).toEqual(TEXT_ALIGNMENTS);
expect(TEXT_ALIGNMENTS).toEqual(['left', 'center', 'right', 'justify']);
expect(Object.isFrozen(TEXT_ALIGNMENTS)).toBe(true);
```

Use the four public PptxGenJS values to create text, write bytes, import natively, and require the four paragraph alignments. Do not mutate or inspect PptxGenJS private fields.

- [ ] **Step 4: Run RED**

```sh
node_modules/.bin/vitest run packages/model/src/text.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "TEXT_ALIGNMENTS|horizontal alignment runtime catalog" --reporter=dot
```

Expected: FAIL because `TEXT_ALIGNMENTS` does not exist.

- [ ] **Step 5: Implement the single source of truth**

Replace the handwritten union in `packages/model/src/text.ts` with:

```ts
export const TEXT_ALIGNMENTS = Object.freeze([
  'left',
  'center',
  'right',
  'justify',
] as const);

export type TextAlignment = typeof TEXT_ALIGNMENTS[number];
```

Do not modify model index, SDK index, root index, normalizers, or OOXML mappings; the existing export chains are sufficient.

- [ ] **Step 6: Run focused, package, and type gates**

```sh
node_modules/.bin/vitest run packages/model/src/text.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "TEXT_ALIGNMENTS|horizontal alignment runtime catalog" --reporter=dot
node_modules/.bin/vitest run packages/model/src/text.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 7: Review, commit, push, and verify**

Review exact ordering, freeze/mutation behavior, derived union, zero facade duplication, existing text OOXML lifecycle, root closure, public-only PptxGenJS comparison, and no unrelated mapping changes. Commit `feat: expose horizontal alignment catalog`, push, fetch, and require divergence `0 0`.

---

### Task 2: Actual-tarball Node, browser, declarations, TypeScript, CLI, and Chrome proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes installed `TEXT_ALIGNMENTS` and `TextAlignment`.
- Produces stable `horizontalAlignments: true` plus `horizontalAlignmentState` in packed Node and real Chrome output.

- [ ] **Step 1: Extend packed declaration checks**

Read `dist/types/model/text.d.ts` and require:

```text
export declare const TEXT_ALIGNMENTS: readonly ["left", "center", "right", "justify"];
export type TextAlignment = typeof TEXT_ALIGNMENTS[number];
```

Accept the declaration generator's quote style but require exact token order, readonly tuple semantics, and derived type. Require root declarations to re-export the model tree as before.

- [ ] **Step 2: Extend installed Node/browser-condition smoke**

Import `TEXT_ALIGNMENTS` in generated Node and browser-condition consumers. Build:

```ts
const horizontalAlignmentState = {
  values: [...TEXT_ALIGNMENTS],
  frozen: Object.isFrozen(TEXT_ALIGNMENTS),
  shared: TEXT_ALIGNMENTS === TEXT_ALIGNMENTS,
};
const horizontalAlignments = JSON.stringify(horizontalAlignmentState.values)
  === JSON.stringify(['left', 'center', 'right', 'justify'])
  && horizontalAlignmentState.frozen
  && horizontalAlignmentState.shared;
```

Create four aligned text shapes, write/reopen, and include their paragraph values in the state. Add both fields to the stable packed JSON and require installed CLI package inspection of the generated deck to succeed.

- [ ] **Step 3: Extend installed TypeScript consumer**

Import runtime and type, then add:

```ts
const horizontalAlignments: readonly TextAlignment[] = TEXT_ALIGNMENTS;
for (const alignment of horizontalAlignments) alignment satisfies TextAlignment;
// @ts-expect-error runtime catalog is readonly
TEXT_ALIGNMENTS.push('left');
// @ts-expect-error catalog index is readonly
TEXT_ALIGNMENTS[0] = 'right';
// @ts-expect-error unknown alignment is rejected
const invalidHorizontalAlignment: TextAlignment = 'distributed';
```

- [ ] **Step 4: Extend real Chrome coverage**

Read `api.TEXT_ALIGNMENTS`, require exact values and `Object.isFrozen()`, create/reopen four aligned shapes, and return:

```ts
horizontalAlignments: true,
horizontalAlignmentState: {
  values: ['left', 'center', 'right', 'justify'],
  reopened: ['left', 'center', 'right', 'justify'],
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

Review installed-only imports, declaration exactness, tuple immutability, Node/browser text round-trip, TypeScript negatives, CLI inspection, real Chrome zero-error state, stable JSON, and clean artifact scope. Commit `test: verify packed horizontal alignments`, push, fetch, and require divergence `0 0`.

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
- Produces the final supported/remaining status without claiming `AlignV` or exact PptxGenJS namespace shape.

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

The 178-preset catalog test already has its explicit 10-second budget; any new failure must be reproduced and resolved, not hidden by a global timeout change.

- [ ] **Step 2: Update exactly the six release documents**

Document:

- `TEXT_ALIGNMENTS` exact values/order/frozen state and derived `TextAlignment`;
- existing text/rich-text/table use and unchanged OOXML mappings;
- native tuple naming versus PptxGenJS instance `AlignH` object;
- deliberate omission of instance getter and mutable enum alias;
- source/root, packed Node/browser/types, installed CLI, and real Chrome evidence;
- final tests, performance, tarball file count/SHA, and Chrome error counts;
- remaining `AlignV`, output/runtime helpers, advanced text/table, `tableToSlides`, and final audit.

- [ ] **Step 3: Search stale status and self-review**

Search non-plan docs for `AlignH` still marked unsupported. Require exactly the six intended files modified, balanced Markdown fences, no placeholders, `git diff --check`, no `.pnpm-store/`, and no generated artifacts staged.

- [ ] **Step 4: Commit, push, and verify**

Commit `docs: document horizontal alignment catalog`, push, fetch, and require divergence `0 0`.

- [ ] **Step 5: Synchronize progress and continue**

Report this专项 3/3 implementation tasks plus design/plan completion, update overall parity progress, list remaining gaps, and select vertical alignment runtime catalog (`AlignV`) as the next smallest independent item.
