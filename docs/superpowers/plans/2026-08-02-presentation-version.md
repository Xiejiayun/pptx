# Presentation Runtime `version` Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Expose the current native package version through a read-only `PptxDocument.version` instance property and `PPTX_VERSION` root constant, then prove manifest, Node, browser, declarations, CLI, and PptxGenJS public-surface behavior stay synchronized.

**Architecture:** Keep one browser-safe compile-time literal in `packages/sdk/src/version.ts`; return it from a getter on `PptxDocument`, re-export it through SDK/root, and replace CLI version literals with the same constant. Repository tests compare all three manifests to the constant, while actual-tarball smoke compares the installed manifest to Node/browser/CLI runtime values.

**Tech Stack:** TypeScript strict mode, Vitest, SDK/root packages, Commander CLI, PptxGenJS 4.0.1 public runtime, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, and real Google Chrome.

## Global Constraints

- Current version is exactly `0.1.0`, matching root, `packages/sdk`, and `packages/pptx` manifests.
- Runtime source must not read `package.json`, use JSON module assertions, inspect filesystem paths, or branch by Node/browser.
- `PptxDocument.version` is a getter with no setter and always returns `PPTX_VERSION`.
- The getter reports the current library runtime, never OOXML `AppVersion`, producer metadata, a source filename, or PptxGenJS's version.
- Create/open/import/write/reopen and all six presentation formats share the same value without package mutation.
- `PPTX_VERSION` is exported from `@pptx/sdk` and `@jiayunxie/pptx`; no generic lower-case `version` export is added.
- PptxGenJS conformance compares availability, readonly typing, stability, and each library's own manifest value; it never expects both version strings to be equal.
- CLI `.version()` and JSON doctor output must consume `PPTX_VERSION`, removing duplicate hard-coded strings.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, browser artifacts, build output, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Add the SDK version constant and read-only document getter

**Files:**
- Create: `packages/sdk/src/version.ts`
- Create: `packages/sdk/src/version.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `PPTX_VERSION = '0.1.0' as const`.
- Produces `PptxVersion = typeof PPTX_VERSION`.
- Produces `PptxDocument.version: PptxVersion` with getter-only runtime semantics.

- [ ] **Step 1: Write the failing constant/manifest sync test**

Create `packages/sdk/src/version.test.ts` using `readFile`, `fileURLToPath`, and repository-relative manifest paths. Require all values to equal the literal and a strict stable-semver expression:

```ts
import { readFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { PPTX_VERSION } from './version.js';

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('PPTX_VERSION', () => {
  it('matches every public package manifest', async () => {
    const versions = await Promise.all([
      'package.json',
      'packages/sdk/package.json',
      'packages/pptx/package.json',
    ].map(async (path) => JSON.parse(
      await readFile(resolve(repositoryRoot, path), 'utf8'),
    ).version));
    expect(PPTX_VERSION).toMatch(/^\d+\.\d+\.\d+$/u);
    expect(versions).toEqual([PPTX_VERSION, PPTX_VERSION, PPTX_VERSION]);
  });
});
```

- [ ] **Step 2: Write the failing document lifecycle test**

In `packages/sdk/src/index.test.ts`, import the constant/type and require:

```ts
const created = PptxDocument.create();
expect(created.version).toBe(PPTX_VERSION);
expect(created.version).toBe('0.1.0');

const before = packageSnapshot(created.opcPackage);
expect(created.version).toBe(created.version);
expect(packageSnapshot(created.opcPackage)).toEqual(before);

const reopened = await PptxDocument.open(await created.write());
expect(reopened.version).toBe(PPTX_VERSION);
expect(Object.getOwnPropertyDescriptor(PptxDocument.prototype, 'version')).toMatchObject({
  set: undefined,
  enumerable: false,
});
```

Add an unreachable compile-time negative:

```ts
if (false) {
  // @ts-expect-error version is read-only
  created.version = '9.9.9';
}
```

- [ ] **Step 3: Run RED**

```sh
node_modules/.bin/vitest run packages/sdk/src/version.test.ts packages/sdk/src/index.test.ts -t "PPTX_VERSION|runtime version" --reporter=dot
```

Expected: FAIL because `version.ts`, the export, and the getter do not exist.

- [ ] **Step 4: Implement the minimal constant and getter**

Create `packages/sdk/src/version.ts`:

```ts
export const PPTX_VERSION = '0.1.0' as const;
export type PptxVersion = typeof PPTX_VERSION;
```

In `packages/sdk/src/index.ts`, export the symbols, import the value/type for local use, and add to `PptxDocument`:

```ts
get version(): PptxVersion {
  return PPTX_VERSION;
}
```

Do not store the version on each instance or add setter/mutation code.

- [ ] **Step 5: Run focused and type gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/version.test.ts packages/sdk/src/index.test.ts -t "PPTX_VERSION|runtime version" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/version.test.ts packages/sdk/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review literal/manifests, browser-safe source, getter descriptor, compile-time readonly behavior, create/open/write stability, and zero package mutation. Commit `feat: expose presentation runtime version`, push through the working remote transport, fetch, and require divergence `0 0`.

---

### Task 2: Root declarations, CLI reuse, and PptxGenJS conformance

**Files:**
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/cli/src/index.ts`
- Modify: `packages/cli/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes Task 1 root re-export through `export * from '@pptx/sdk'`.
- Produces CLI `--version` and doctor JSON from `PPTX_VERSION`.
- Produces public PptxGenJS/native runtime conformance evidence without equating their strings.

- [ ] **Step 1: Add root type/runtime tests**

In `packages/pptx/src/index.test.ts`, import `PPTX_VERSION` and `PptxVersion`, then require the root value, created/opened instances, literal assignment, and readonly negative:

```ts
const current: PptxVersion = PPTX_VERSION;
const document = PptxDocument.create();
expect(current).toBe('0.1.0');
expect(document.version).toBe(current);
expect((await PptxDocument.open(await document.write())).version).toBe(current);
```

- [ ] **Step 2: Replace CLI literals and extend tests**

Import `PPTX_VERSION` from `@pptx/sdk` in `packages/cli/src/index.ts`. Replace both `.version('0.1.0')` and doctor `{ version: '0.1.0' }` with the constant.

Extend `packages/cli/src/index.test.ts` so JSON doctor requires `data.version === PPTX_VERSION`, and run `['--version']` with captured stdout requiring `${PPTX_VERSION}\n` and exit code 0.

- [ ] **Step 3: Add public PptxGenJS behavior comparison**

In the adapter test, instantiate PptxGenJS through its public constructor and Native through `PptxDocument.create()`. Require:

```ts
expect(pptx.version).toBe('4.0.1');
expect(native.version).toBe(PPTX_VERSION);
expect(pptx.version).not.toBe(native.version);

await pptx.write({ outputType: 'uint8array' });
await native.write();
expect(pptx.version).toBe('4.0.1');
expect(native.version).toBe(PPTX_VERSION);
```

Use only PptxGenJS public constructor, property, slide creation if needed, and write method. Do not inspect private fields.

- [ ] **Step 4: Run focused and package gates**

```sh
node_modules/.bin/vitest run packages/pptx/src/index.test.ts packages/cli/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "runtime version|doctor" --reporter=dot
node_modules/.bin/vitest run packages/pptx/src/index.test.ts packages/cli/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review root declaration closure, literal type, CLI single-source use, `--version` exit semantics, doctor JSON stability, public-only PptxGenJS probe, and absence of any expected cross-library string equality. Commit `test: verify presentation runtime version`, push, fetch, and require divergence `0 0`.

---

### Task 3: Actual-tarball Node, browser, declarations, and installed CLI proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes installed `PPTX_VERSION`, `PptxDocument.version`, package manifest, browser conditional export, declarations, and CLI.
- Produces `presentationVersion: true` plus detailed stable state in Node/browser smoke output.

- [ ] **Step 1: Extend installed manifest/declaration checks**

After loading the installed manifest, stop hard-coding `0.1.0` in the identity assertion; require name and stable version shape, then compare imported `PPTX_VERSION` and document version with `manifest.version` inside the installed consumer.

Require generated declarations to contain:

```text
export declare const PPTX_VERSION: "0.1.0";
export type PptxVersion = typeof PPTX_VERSION;
get version(): PptxVersion;
```

The exact declaration file paths may follow the current generated type tree; do not assume symbols remain in one bundled `.d.ts` if the generator preserves source structure.

- [ ] **Step 2: Extend installed Node and CLI checks**

Import `PPTX_VERSION` in generated `smoke.mjs`. Create, write, and reopen one document, then require all three version values and the installed manifest version to match. Add stable output:

```ts
presentationVersion: true,
presentationVersionState: {
  constant: PPTX_VERSION,
  created: created.version,
  reopened: reopened.version,
  manifest: manifestVersion,
  cli: doctor.data.version,
},
```

Pass `manifest.version` into the generated consumer without adding a runtime path dependency, for example by serializing the already-read string into the generated script. Require installed `pptx-inspect --version` and JSON doctor to report the same value.

- [ ] **Step 3: Extend real browser coverage**

In `scripts/playwright-browser-smoke.js`, read `api.PPTX_VERSION`, create/writeBlob/reopen a document, and require identical version values with zero package diagnostics. Add `presentationVersion` and `presentationVersionState` to actual and expected objects. The browser must not fetch package metadata or use Node APIs.

- [ ] **Step 4: Extend installed TypeScript consumer**

Import `PPTX_VERSION`, `PptxVersion`, and `PptxDocument`; accept literal assignment and add:

```ts
const version: PptxVersion = PPTX_VERSION;
const versionDocument = PptxDocument.create();
versionDocument.version satisfies PptxVersion;
// @ts-expect-error document version is read-only
versionDocument.version = '9.9.9';
```

- [ ] **Step 5: Build, pack, and run packed gates**

Run Node/browser tsup directly from `packages/pptx`, then declaration generation, a fresh `npm pack --ignore-scripts`, installed smoke, and real Chrome. Record tarball file count and SHA-256. Do not use the workspace filter build entry that downloads unrelated platform packages.

- [ ] **Step 6: Review, commit, push, and verify**

Review actual installed-only imports, manifest comparison, declaration literal/readonly surface, CLI output, browser no-fetch behavior, stable JSON additions, and absence of generated artifacts from git. Commit `test: verify packed presentation runtime version`, push, fetch, and require divergence `0 0`.

---

### Task 4: Release gates, compatibility status, and documentation

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes verified Task 1–3 behavior and actual tarball evidence.
- Produces the final supported/remaining status without claiming other runtime helpers or full parity.

- [ ] **Step 1: Run complete release gates**

```sh
node_modules/.bin/vitest run --reporter=dot
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
```

Then rerun both tsup builds and `node scripts/build-npm-package-types.mjs`. Any isolated timeout must be reproduced alone and the full suite rerun without competing build load; do not conceal a real assertion failure.

- [ ] **Step 2: Update exactly the six release documents**

Document:

- `PPTX_VERSION`, `PptxVersion`, and read-only `PptxDocument.version`;
- compile-time constant and manifest drift guard;
- create/open/write/reopen, root/types, Node/browser/CLI/actual-tarball coverage;
- PptxGenJS `4.0.1` versus native `0.1.0` and why unequal values are correct;
- absence of OOXML/file-producer coupling and zero package mutation;
- remaining `presLayout`, runtime helper constants, output types/stream/compression, advanced text/table/`tableToSlides`, and final audit work;
- actual final tests, performance, tarball file count/SHA, Chrome/CLI results.

- [ ] **Step 3: Search stale status and self-review**

Search all non-plan docs for `version` marked unsupported or partial. Require only the six intended files modified, balanced Markdown fences, no placeholder text, `git diff --check`, and no staged `.pnpm-store/`.

- [ ] **Step 4: Commit, push, and verify**

Commit `docs: document presentation runtime version`, push, fetch, and require divergence `0 0`.

- [ ] **Step 5: Synchronize progress and continue**

Report this专项 4/4 and overall parity progress, list the remaining runtime/output and content gaps, then select the next smallest independent PptxGenJS parity item without stopping at the milestone.
