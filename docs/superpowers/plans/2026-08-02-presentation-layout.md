# Presentation `presLayout` Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add a getter-only `PptxDocument.presLayout` projection that exposes the current canvas name and EMU dimensions across source, root, packed Node/browser, declarations, CLI inspection, and PptxGenJS conformance evidence.

**Architecture:** Keep `slideSize` and `p:sldSz` as the only source of truth. A small browser-safe SDK module maps exact standard dimensions to canonical PptxGenJS names and returns a new `{ name, width, height }` snapshot; `PptxDocument` delegates its getter to that pure projection. Packed smoke and real Chrome validate the same API without runtime metadata or a layout registry.

**Tech Stack:** TypeScript strict mode, Vitest, SDK/root packages, PptxGenJS 4.0.1 public runtime, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, installed CLI, and real Google Chrome.

## Global Constraints

- Width and height use native `Emu`, exactly matching `SlideSize` and PptxGenJS `presLayout` runtime output.
- `p:sldSz` remains the only persistent source of truth; do not add registry, cache, instance field, package extension, or custom XML.
- Exact 9,144,000 × 6,858,000, 9,144,000 × 5,143,500, and 9,144,000 × 5,715,000 map to `screen4x3`, `screen16x9`, and `screen16x10` respectively.
- Every other valid size, including 12,192,000 × 6,858,000 wide, maps to `custom`.
- `presLayout` and its fields are readonly in declarations; every read returns a detached, unfrozen plain-object snapshot.
- Do not expose PptxGenJS runtime-private `_sizeW` or `_sizeH`.
- Do not add PptxGenJS `layout` setter or `defineLayout()` registry to native; `slideSize` remains the create/edit API.
- Malformed `p:sldSz` reuses the existing `slideSize` error and never falls back or repairs bytes.
- Reading `presLayout` must not change parts, relationships, graph, mutation journal, diagnostics, or output bytes.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, browser artifacts, build output, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Add the SDK projection and document getter

**Files:**
- Create: `packages/sdk/src/presentation-layout.ts`
- Create: `packages/sdk/src/presentation-layout.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes `Emu`, `SlideSize`, and `PptxDocument.slideSize` from `@pptx/model`.
- Produces `PresentationLayoutName`, `PresentationLayout`, internal `presentationLayoutFromSlideSize()`, and getter-only `PptxDocument.presLayout`.

- [ ] **Step 1: Write the failing pure projection tests**

Create `packages/sdk/src/presentation-layout.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { inches } from '@pptx/model';
import { presentationLayoutFromSlideSize } from './presentation-layout.js';

describe('presentationLayoutFromSlideSize', () => {
  it.each([
    [inches(10), inches(7.5), 'screen4x3'],
    [inches(10), inches(5.625), 'screen16x9'],
    [inches(10), inches(6.25), 'screen16x10'],
    [inches(13 + 1 / 3), inches(7.5), 'custom'],
    [inches(11.7), inches(8.3), 'custom'],
  ] as const)('maps %s × %s to %s', (width, height, name) => {
    expect(presentationLayoutFromSlideSize({ width, height })).toEqual({
      name,
      width,
      height,
    });
  });

  it('uses exact dimensions and returns detached snapshots', () => {
    const input = { width: inches(10), height: inches(5.625) };
    const first = presentationLayoutFromSlideSize(input);
    const second = presentationLayoutFromSlideSize(input);
    expect(first).toEqual(second);
    expect(first).not.toBe(second);
    expect(presentationLayoutFromSlideSize({
      width: input.width,
      height: input.height + 1,
    })).toEqual({ name: 'custom', width: input.width, height: input.height + 1 });
  });
});
```

- [ ] **Step 2: Write failing document lifecycle tests**

In `packages/sdk/src/index.test.ts`, import the two public types and add one focused test beside runtime version. Require:

```ts
const sizes = [
  ['4:3', 'screen4x3', 9_144_000, 6_858_000],
  ['16:9', 'screen16x9', 9_144_000, 5_143_500],
  ['16:10', 'screen16x10', 9_144_000, 5_715_000],
  ['wide', 'custom', 12_192_000, 6_858_000],
] as const;

for (const [slideSize, name, width, height] of sizes) {
  const document = PptxDocument.create({ slideSize });
  const before = await sdkPackageSnapshot(document);
  const layout: PresentationLayout = document.presLayout;
  const layoutName: PresentationLayoutName = layout.name;
  expect({ ...layout, name: layoutName }).toEqual({ name, width, height });
  expect(document.presLayout).not.toBe(layout);
  expect(await sdkPackageSnapshot(document)).toEqual(before);
}
```

Then cover detached mutation, valid edit, invalid-edit rollback, write/reopen, descriptor, and malformed input:

```ts
const custom = PptxDocument.create({
  slideSize: { width: inches(11.7), height: inches(8.3) },
});
const detached = custom.presLayout as { name: string; width: number; height: number };
detached.name = 'changed';
detached.width = 1;
expect(custom.presLayout).toEqual({
  name: 'custom', width: inches(11.7), height: inches(8.3),
});
custom.slideSize = { width: inches(10), height: inches(6.25) };
expect(custom.presLayout.name).toBe('screen16x10');
expect(() => {
  custom.slideSize = { width: 0 as never, height: inches(7.5) };
}).toThrow(RangeError);
expect(custom.presLayout.name).toBe('screen16x10');
expect((await PptxDocument.open(await custom.write())).presLayout)
  .toEqual(custom.presLayout);
expect(Object.getOwnPropertyDescriptor(PptxDocument.prototype, 'presLayout'))
  .toMatchObject({ set: undefined, enumerable: false });
```

Add unreachable compile-time negatives for assignment to the property and snapshot fields. Corrupt a fresh deck's `p:sldSz@cx` to `0` and require `document.presLayout` to throw the existing slide-width error.

- [ ] **Step 3: Run RED**

```sh
node_modules/.bin/vitest run packages/sdk/src/presentation-layout.test.ts packages/sdk/src/index.test.ts -t "presentationLayoutFromSlideSize|presentation layout projection" --reporter=dot
```

Expected: FAIL because the module, exports, and getter do not exist.

- [ ] **Step 4: Implement the minimal browser-safe projection**

Create `packages/sdk/src/presentation-layout.ts`:

```ts
import { inches, type Emu, type SlideSize } from '@pptx/model';

export type PresentationLayoutName =
  | 'screen4x3'
  | 'screen16x9'
  | 'screen16x10'
  | 'custom';

export interface PresentationLayout {
  readonly name: PresentationLayoutName;
  readonly width: Emu;
  readonly height: Emu;
}

const STANDARD_LAYOUTS = [
  { name: 'screen4x3', width: inches(10), height: inches(7.5) },
  { name: 'screen16x9', width: inches(10), height: inches(5.625) },
  { name: 'screen16x10', width: inches(10), height: inches(6.25) },
] as const satisfies readonly PresentationLayout[];

export function presentationLayoutFromSlideSize(
  slideSize: Readonly<SlideSize>,
): PresentationLayout {
  const standard = STANDARD_LAYOUTS.find(
    ({ width, height }) => width === slideSize.width && height === slideSize.height,
  );
  return {
    name: standard?.name ?? 'custom',
    width: slideSize.width,
    height: slideSize.height,
  };
}
```

In `packages/sdk/src/index.ts`, import the function/type, re-export only the two public types, and add:

```ts
get presLayout(): PresentationLayout {
  return presentationLayoutFromSlideSize(this.slideSize);
}
```

- [ ] **Step 5: Run focused, package, and type gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/presentation-layout.test.ts packages/sdk/src/index.test.ts -t "presentationLayoutFromSlideSize|presentation layout projection" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/presentation-layout.test.ts packages/sdk/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review exact mapping, EMU types, getter-only descriptor, detached mutation, edit/rollback, malformed input, write/reopen, mutation isolation, lack of private fields, and browser-safe imports. Commit `feat: expose presentation layout projection`, push, fetch, and require divergence `0 0`.

---

### Task 2: Root types and PptxGenJS public conformance

**Files:**
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes Task 1 exports through the existing root `export * from '@pptx/sdk'` chain.
- Produces root declaration closure and public-only PptxGenJS comparison for default, four built-ins, custom, descriptor, and write stability.

- [ ] **Step 1: Add root public API tests**

Import `PresentationLayout` and `PresentationLayoutName` in `packages/pptx/src/index.test.ts`. Add a root stable-export test that creates wide, reads `{ name: 'custom', width: 12_192_000, height: 6_858_000 }`, edits to 16:9, writes/reopens, and verifies the canonical `screen16x9` snapshot. Include `@ts-expect-error` checks for assigning `document.presLayout` and `layout.width`.

- [ ] **Step 2: Extend the local PptxGenJS public interface**

Add to `PptxGenJSInstance`:

```ts
readonly presLayout: {
  readonly name: string;
  readonly width: number;
  readonly height: number;
};
defineLayout(layout: {
  readonly name: string;
  readonly width: number;
  readonly height: number;
}): void;
```

This test-only interface must not add `_sizeW` or `_sizeH`.

- [ ] **Step 3: Add public runtime conformance**

In the adapter suite, instantiate a fresh PptxGenJS and native document for each case:

```ts
const cases = [
  ['LAYOUT_4x3', '4:3', 'screen4x3', 9_144_000, 6_858_000],
  ['LAYOUT_16x9', '16:9', 'screen16x9', 9_144_000, 5_143_500],
  ['LAYOUT_16x10', '16:10', 'screen16x10', 9_144_000, 5_715_000],
  ['LAYOUT_WIDE', 'wide', 'custom', 12_192_000, 6_858_000],
] as const;
```

For default and each case, compare exactly `{ name, width, height }` after selecting the PptxGenJS `layout`. Require native keys to be exactly `['name', 'width', 'height']`, repeated native reads to be detached, and both prototype descriptors to have no setter.

For custom:

```ts
generated.defineLayout({ name: 'CUSTOM', width: 11.7, height: 8.3 });
generated.layout = 'CUSTOM';
const native = PptxDocument.create({
  slideSize: { width: inches(11.7), height: inches(8.3) },
});
expect(generated.presLayout).toMatchObject({
  name: 'CUSTOM', width: inches(11.7), height: inches(8.3),
});
expect(native.presLayout).toEqual({
  name: 'custom', width: inches(11.7), height: inches(8.3),
});
```

Write both decks, require the public values to remain stable, open the PptxGenJS bytes natively, and require its `presLayout` width/height to match with canonical name `custom`.

- [ ] **Step 4: Run focused and package gates**

```sh
node_modules/.bin/vitest run packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "presentation layout|presLayout" --reporter=dot
node_modules/.bin/vitest run packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review root type exports, readonly negatives, public-only PptxGenJS access, all layout mappings, custom-name boundary, descriptor semantics, write stability, and absence of private-field assertions. Commit `test: verify presentation layout parity`, push, fetch, and require divergence `0 0`.

---

### Task 3: Actual-tarball Node, browser, declarations, TypeScript, CLI, and Chrome proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes installed `PresentationLayout`, `PresentationLayoutName`, `PptxDocument.presLayout`, browser conditional export, generated declarations, and CLI package inspection.
- Produces stable `presentationLayouts: true` and `presentationLayoutState` evidence in packed Node and real Chrome output.

- [ ] **Step 1: Extend packed declaration checks**

Read `dist/types/sdk/presentation-layout.d.ts` and require declarations for the four-name union plus readonly `name`, `width`, and `height`. Require `dist/types/sdk/index.d.ts` to contain `get presLayout(): PresentationLayout;` and to export both public types. Do not look for `_sizeW` or `_sizeH`.

- [ ] **Step 2: Extend the installed Node consumer**

In generated `smoke.mjs`, record default, wide, custom, edited, and reopened states:

```ts
const packedDefaultLayout = PptxDocument.create().presLayout;
const packedWideLayout = PptxDocument.create({ slideSize: 'wide' }).presLayout;
const packedLayoutDeck = PptxDocument.create({
  slideSize: { width: inches(11.7), height: inches(8.3) },
});
const packedCustomLayout = packedLayoutDeck.presLayout;
packedLayoutDeck.slideSize = { width: inches(10), height: inches(6.25) };
const packedEditedLayout = packedLayoutDeck.presLayout;
const packedReopenedLayout = (
  await PptxDocument.open(await packedLayoutDeck.write())
).presLayout;
```

Build `presentationLayoutState` from these values, verify exact names/dimensions and detached reads, add `presentationLayouts` to the checks object, and include both fields in the final smoke JSON. Write the deck to a temporary PPTX and run installed `pptx-inspect --json package inspect`; require successful parsing and the presentation part to be present.

- [ ] **Step 3: Extend browser-condition and TypeScript consumers**

In `browser-smoke.mjs`, create a custom deck, edit it to 16:10, `writeBlob()`/reopen, and require custom → screen16x10 state without Node APIs.

In `smoke.ts`, import `PresentationLayout` and `PresentationLayoutName` and add:

```ts
const layoutDocument = PptxDocument.create({ slideSize: 'wide' });
const presentationLayout: PresentationLayout = layoutDocument.presLayout;
const presentationLayoutName: PresentationLayoutName = presentationLayout.name;
presentationLayoutName satisfies 'screen4x3' | 'screen16x9' | 'screen16x10' | 'custom';
// @ts-expect-error presLayout is getter-only
layoutDocument.presLayout = presentationLayout;
// @ts-expect-error presentation layout width is readonly
presentationLayout.width = inches(1);
```

- [ ] **Step 4: Extend real Chrome state**

In `scripts/playwright-browser-smoke.js`, create one 4:3 deck and one custom deck, edit the custom deck to wide, writeBlob/reopen, and return:

```ts
presentationLayouts: true,
presentationLayoutState: {
  standard: { name: 'screen4x3', width: 9144000, height: 6858000 },
  custom: { name: 'custom', width: 10698480, height: 7589520 },
  edited: { name: 'custom', width: 12192000, height: 6858000 },
  reopened: { name: 'custom', width: 12192000, height: 6858000 },
},
```

Require validation/console/page/network error counts to stay zero.

- [ ] **Step 5: Build, pack, and run installed gates**

```sh
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
(cd packages/pptx && npm pack --ignore-scripts)
node scripts/smoke-npm-package.mjs packages/pptx/jiayunxie-pptx-0.1.0.tgz
node scripts/playwright-browser-smoke.js packages/pptx/jiayunxie-pptx-0.1.0.tgz
```

Record actual tarball file count and SHA-256, remove the generated tarball after evidence is captured, and confirm `git status` contains no generated tracked changes.

- [ ] **Step 6: Review, commit, push, and verify**

Review installed-only imports, exact declaration surface, readonly type negatives, Node/browser custom/edit/reopen state, CLI package inspection, real Chrome zero-error evidence, stable smoke JSON, and absence of build artifacts. Commit `test: verify packed presentation layouts`, push, fetch, and require divergence `0 0`.

---

### Task 4: Full release gates, compatibility status, and documentation

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes verified Task 1–3 behavior and actual tarball/Chrome evidence.
- Produces the final supported/remaining status without claiming named layout registry support or full parity.

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

Any isolated timeout must be reproduced alone and the full suite rerun without competing build load; do not conceal an assertion failure.

- [ ] **Step 2: Update exactly the six release documents**

Document:

- `PresentationLayoutName`, `PresentationLayout`, and getter-only `PptxDocument.presLayout`;
- exact `screen4x3` / `screen16x9` / `screen16x10` / `custom` mapping and EMU units;
- `slideSize` as the only create/edit source of truth and detached snapshot semantics;
- default/four built-ins/custom/edit/write/reopen/malformed/mutation-isolation coverage;
- PptxGenJS public runtime parity and deliberate omission of `_sizeW` / `_sizeH`;
- custom registry name loss across PPTX and why canonical `custom` is correct;
- root/types, packed Node/browser, installed CLI inspection, actual tarball, and real Chrome evidence;
- final tests, performance, tarball file count/SHA-256, and Chrome error counts;
- remaining runtime constants, output types/stream/compression, advanced text/table, `tableToSlides`, and final peer/client audit.

- [ ] **Step 3: Search stale status and self-review**

Search non-plan docs for `presLayout` still marked unsupported or pending. Require only the six intended documents modified, balanced Markdown fences, no placeholder text, `git diff --check`, no staged `.pnpm-store/`, and no generated tarball/build output.

- [ ] **Step 4: Commit, push, and verify**

Commit `docs: document presentation layout projection`, push, fetch, and require divergence `0 0`.

- [ ] **Step 5: Synchronize progress and continue**

Report this专项 4/4 implementation tasks plus design/plan completion, update overall parity progress from about 94%, list remaining runtime/output/content gaps, and select the next smallest independent PptxGenJS parity item without stopping at the milestone.
