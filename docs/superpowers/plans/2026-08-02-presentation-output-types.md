# Presentation Output Type Runtime Catalog Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with independent review, commit, push, fetch, and remote-divergence verification.

**Goal:** Expose the six PptxGenJS-compatible presentation output-type tokens as one frozen SDK/root runtime catalog and derive a single public `OutputType` union without changing native write behavior.

**Architecture:** Add a focused `packages/sdk/src/output-type.ts` module because output types belong to the public I/O layer rather than the document model. Export its value/type through SDK `index.ts`; the aggregate package reuses that SDK export. Permanent source, packed Node/browser/types, installed CLI, and real-Chrome checks lock exact values, order, identity, immutability, and package-mutation isolation.

**Tech Stack:** TypeScript strict mode, Vitest, SDK/root packages, PptxGenJS 4.0.1 public runtime, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, installed CLI, and real Google Chrome.

## Global Constraints

- Exact order is `arraybuffer`, `base64`, `binarystring`, `blob`, `nodebuffer`, `uint8array`.
- `OUTPUT_TYPES` is a frozen readonly tuple created once at SDK module initialization.
- `OutputType` is exactly `(typeof OUTPUT_TYPES)[number]`, with no widened string or duplicate handwritten union.
- Do not include `STREAM`; it belongs to PptxGenJS `WRITE_OUTPUT_TYPE`/`stream()` but is not an instance `OutputType` enum member.
- Do not add `PptxDocument.OutputType`, a generic `OutputType` runtime value, enum-shaped object, setter, registry, or mutable alias.
- Do not add `WriteOptions.outputType`, write overloads, output conversion, stream support, or compression behavior in this item.
- Preserve current `write(): Promise<Uint8Array>`, `writeBlob()`, `writeFile()`, `download()`, diagnostics, bytes, MIME, and errors exactly.
- SDK owns the catalog; aggregate root must reuse the exact SDK object without a facade copy.
- PptxGenJS conformance uses only public constructor and `OutputType` getter.
- Never stage `.pnpm-store/`, dist output, tarballs, browser artifacts, temporary consumers, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Add the SDK catalog, derived type, public exports, and PptxGenJS conformance

**Files:**
- Create: `packages/sdk/src/output-type.ts`
- Create: `packages/sdk/src/output-type.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Produces `OUTPUT_TYPES: readonly ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']`.
- Produces `OutputType = typeof OUTPUT_TYPES[number]`.
- Preserves every existing `PptxDocument` output method signature and runtime behavior.
- Supplies the single type source consumed by the later six-output implementation.

- [ ] **Step 1: Write the focused SDK module tests**

Create `packages/sdk/src/output-type.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { OUTPUT_TYPES, type OutputType } from './output-type.js';

describe('OUTPUT_TYPES', () => {
  it('publishes the complete frozen output type catalog', () => {
    expect(OUTPUT_TYPES).toEqual([
      'arraybuffer',
      'base64',
      'binarystring',
      'blob',
      'nodebuffer',
      'uint8array',
    ]);
    expect(new Set(OUTPUT_TYPES)).toHaveLength(6);
    expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);
    expect(() => {
      (OUTPUT_TYPES as unknown as string[]).push('uint8array');
    }).toThrow(TypeError);
  });

  it('keeps the runtime catalog and OutputType synchronized', () => {
    const values: readonly OutputType[] = OUTPUT_TYPES;
    expect(values).toBe(OUTPUT_TYPES);
    if (false) {
      // @ts-expect-error STREAM is handled by the separate stream API
      const stream: OutputType = 'STREAM';
      // @ts-expect-error unknown output type is not supported
      const unknown: OutputType = 'buffer';
      void [stream, unknown];
    }
  });
});
```

- [ ] **Step 2: Add failing SDK public-export and mutation-isolation coverage**

Import `OUTPUT_TYPES` and `OutputType` from `./index.js` in `packages/sdk/src/index.test.ts`. Add:

```ts
it('publishes the frozen output type catalog without package mutation', () => {
  const document = PptxDocument.create();
  const journal = [...document.opcPackage.mutations];
  const values: readonly OutputType[] = OUTPUT_TYPES;

  expect(values).toBe(OUTPUT_TYPES);
  expect([...values]).toEqual([
    'arraybuffer',
    'base64',
    'binarystring',
    'blob',
    'nodebuffer',
    'uint8array',
  ]);
  expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);
  expect(document.opcPackage.mutations).toEqual(journal);
});
```

Do not call a new output selector; this test proves catalog discovery is package-independent.

- [ ] **Step 3: Add failing aggregate-root identity and type coverage**

In `packages/pptx/src/index.test.ts`, import SDK value as `SDK_OUTPUT_TYPES`, and import root `OUTPUT_TYPES` plus `OutputType`. Add:

```ts
it('exports the frozen OUTPUT_TYPES catalog from the root package', () => {
  const values: readonly OutputType[] = OUTPUT_TYPES;

  expect(OUTPUT_TYPES).toBe(SDK_OUTPUT_TYPES);
  expect(values).toBe(OUTPUT_TYPES);
  expect([...values]).toEqual([
    'arraybuffer',
    'base64',
    'binarystring',
    'blob',
    'nodebuffer',
    'uint8array',
  ]);
  expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);
});
```

Include compile-only negatives for `'STREAM'`, `'buffer'`, tuple `push`, and index assignment.

- [ ] **Step 4: Add failing public-only PptxGenJS comparison**

In `packages/pptxgenjs-adapter/src/index.test.ts`:

1. import `OUTPUT_TYPES` from `@pptx/sdk`;
2. extend the local public-only `PptxGenJSInstance` shape with:

```ts
readonly OutputType: Readonly<Record<string, string>>;
```

3. add a test requiring both keys and values to match in order:

```ts
it('matches the PptxGenJS output type runtime catalog', () => {
  const generated = new PptxGenJS();

  expect(Object.keys(generated.OutputType)).toEqual(OUTPUT_TYPES);
  expect(Object.values(generated.OutputType)).toEqual(OUTPUT_TYPES);
  expect(OUTPUT_TYPES).toEqual([
    'arraybuffer',
    'base64',
    'binarystring',
    'blob',
    'nodebuffer',
    'uint8array',
  ]);
  expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);
});
```

Do not inspect `_outputType`, source files, JSZip internals, or other private state.

- [ ] **Step 5: Run RED**

```sh
node_modules/.bin/vitest run packages/sdk/src/output-type.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "OUTPUT_TYPES|output type runtime catalog" --reporter=dot
```

Expected: FAIL because `OUTPUT_TYPES` and `OutputType` do not exist.

- [ ] **Step 6: Implement the focused SDK source and export**

Create `packages/sdk/src/output-type.ts`:

```ts
export const OUTPUT_TYPES = Object.freeze([
  'arraybuffer',
  'base64',
  'binarystring',
  'blob',
  'nodebuffer',
  'uint8array',
] as const);

export type OutputType = typeof OUTPUT_TYPES[number];
```

Add only these export lines to `packages/sdk/src/index.ts`:

```ts
export { OUTPUT_TYPES } from './output-type.js';
export type { OutputType } from './output-type.js';
```

Do not import the catalog into the `PptxDocument` implementation or modify `WriteOptions`.

- [ ] **Step 7: Run focused, package, and type gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/output-type.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "OUTPUT_TYPES|output type runtime catalog" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/output-type.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 8: Review, commit, push, and verify**

Review exact order, six unique values, freeze/mutation behavior, derived union, `STREAM` exclusion, SDK ownership, root identity, unchanged write API, public-only PptxGenJS comparison, and unrelated-diff absence. Stage only the six listed files, commit `feat: expose output type catalog`, push, fetch, and require divergence `0 0`.

---

### Task 2: Actual-tarball Node, browser, declarations, TypeScript, CLI, and Chrome proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes installed `OUTPUT_TYPES` and `OutputType`.
- Produces stable `outputTypes: true` plus `outputTypeState` in packed Node and real Chrome output.
- Proves catalog discovery does not mutate a presentation or require environment-specific output globals.

- [ ] **Step 1: Extend packed declaration checks**

Read `dist/types/sdk/output-type.d.ts`, normalize quote style and whitespace, then require:

```text
export declare const OUTPUT_TYPES: readonly ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array'];
export type OutputType = typeof OUTPUT_TYPES[number];
```

Also require `dist/types/sdk/index.d.ts` to re-export both surfaces. Do not accept a widened `readonly string[]` or duplicated literal union.

- [ ] **Step 2: Extend installed Node and browser-condition smoke**

Add `OUTPUT_TYPES` to generated consumer imports. Use a fresh same-package dynamic import and a new document journal snapshot:

```ts
const outputTypeDocument = PptxDocument.create();
const outputTypeJournal = [...outputTypeDocument.opcPackage.mutations];
const reimportedOutputTypes = (
  await import('@jiayunxie/pptx')
).OUTPUT_TYPES;
const outputTypeState = {
  values: [...OUTPUT_TYPES],
  frozen: Object.isFrozen(OUTPUT_TYPES),
  shared: OUTPUT_TYPES === reimportedOutputTypes,
  mutationIsolation: JSON.stringify(outputTypeDocument.opcPackage.mutations)
    === JSON.stringify(outputTypeJournal),
};
```

Set `outputTypes` true only when exact values, frozen/shared, and mutation isolation all pass. Add both fields to stable packed JSON. The browser-condition consumer must resolve `dist/browser.js`, expose the same exact frozen values, and keep existing writeBlob/reopen checks green.

- [ ] **Step 3: Extend installed TypeScript consumer**

Import runtime and type, then add:

```ts
const outputTypes: readonly OutputType[] = OUTPUT_TYPES;
for (const outputType of outputTypes) {
  outputType satisfies OutputType;
}
// @ts-expect-error runtime output type catalog is readonly
OUTPUT_TYPES.push('uint8array');
// @ts-expect-error output type catalog indexes are readonly
OUTPUT_TYPES[0] = 'uint8array';
// @ts-expect-error STREAM is handled by the separate stream API
const streamOutputType: OutputType = 'STREAM';
// @ts-expect-error unknown output type is rejected
const invalidOutputType: OutputType = 'buffer';
```

Run the installed consumer under its existing strict TypeScript configuration.

- [ ] **Step 4: Extend real Chrome coverage**

In `scripts/playwright-browser-smoke.js`, read `api.OUTPUT_TYPES`, capture package mutations before and after iteration, and return:

```ts
outputTypes: true,
outputTypeState: {
  values: [
    'arraybuffer',
    'base64',
    'binarystring',
    'blob',
    'nodebuffer',
    'uint8array',
  ],
  frozen: true,
  mutationIsolation: true,
},
```

Do not call `write({ outputType })` because that selector belongs to the next implementation item. Existing create/writeBlob/reopen and Chrome validation/console/page/network checks must remain green.

- [ ] **Step 5: Run complete release and installed gates**

```sh
node_modules/.bin/vitest run --reporter=dot
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
(cd packages/pptx && npm pack --ignore-scripts)
node scripts/smoke-npm-package.mjs packages/pptx/jiayunxie-pptx-0.1.0.tgz
```

Run the repository browser smoke callback in installed Google Chrome against the built browser module. Record actual test/file totals, performance time, tarball file count and SHA-256, and Chrome state. Move tarball/browser artifacts out of the workspace and confirm no generated tracked files remain.

- [ ] **Step 6: Review, commit, push, and verify**

Review installed-only imports, declaration exactness, six-value immutability, dynamic-import identity, TypeScript negatives, package-mutation isolation, CLI inspection, real Chrome zero-error state, stable JSON, and clean artifact scope. Stage only the two listed scripts, commit `test: verify packed output types`, push, fetch, and require divergence `0 0`.

---

### Task 3: Compatibility status and release documentation

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes verified Task 1–2 behavior and actual tarball/Chrome evidence.
- Produces final supported/remaining status without claiming six-output write semantics, stream, compression, or exact PptxGenJS namespace shape.

- [ ] **Step 1: Update exactly the six release documents**

Document all of the following:

- `OUTPUT_TYPES` exact six values/order/frozen state and derived `OutputType`;
- SDK ownership and aggregate-root identity;
- PptxGenJS instance `OutputType` value/order conformance;
- deliberate omission of instance getter, mutable enum alias, and `STREAM` from the catalog;
- unchanged native `write(): Uint8Array`, `writeBlob()`, `writeFile()`, and `download()` behavior;
- source/root, packed Node/browser/types, installed CLI, and real Chrome evidence;
- final tests, performance, tarball file count/SHA, and Chrome error counts;
- remaining six-output selector/return semantics, Node readable stream, compression policy, scheme-color/other helpers, advanced text/table, `tableToSlides`, and final audit.

- [ ] **Step 2: Search stale status and self-review**

Search non-plan docs for `OutputType` still listed wholly unsupported. Keep distinctions that say only the catalog is complete while output conversion remains pending. Require exactly the six intended files modified, balanced Markdown fences, no placeholders, `git diff --check`, no `.pnpm-store/`, and no generated artifacts staged.

- [ ] **Step 3: Commit, push, and verify**

Stage only the six listed documents, commit `docs: document output type catalog`, push, fetch, and require divergence `0 0`.

- [ ] **Step 4: Synchronize progress and continue**

Report design, plan, source, packed proof, and docs as 5/5 complete; update overall parity progress only if measured coverage justifies a change. List remaining gaps and immediately begin the next coherent output item: six-value `write({ outputType })` return semantics, with stream/compression kept separate unless its design proves they must move together.
