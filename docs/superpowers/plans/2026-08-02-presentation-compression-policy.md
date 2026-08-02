# Presentation Compression Policy Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with independent review, commit, push, fetch, and remote-divergence verification.

**Goal:** Add one deterministic boolean ZIP compression policy to every `PptxDocument` output surface while preserving unchanged-package original bytes and browser-safe output types.

**Architecture:** Extend `WriteBaseOptions` with `compression?: boolean`, validate it once in the SDK before diagnostics and ZIP generation, and pass it into a minimal `PackageWriteOptions` owned by OPC. OPC alone chooses unchanged-original, STORE, or DEFLATE level 6; all SDK output representations and convenience methods continue to consume the same canonical `Uint8Array`.

**Tech Stack:** TypeScript strict mode, JSZip 3.10.1, Vitest, PptxGenJS 4.0.1 public probes, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, installed CLI, and real Google Chrome.

## Global Constraints

- Public selector is exactly `readonly compression?: boolean` on `WriteBaseOptions`.
- Omitted/undefined and false generate STORE for created or changed packages; true generates DEFLATE level 6.
- An opened, unchanged package with omitted/undefined selector returns a defensive copy of original bytes.
- Explicit false or true always requests regeneration, including for an opened, unchanged package.
- `write()`, all six output types, `stream()`, `writeFile()`, `writeBlob()`, and `download()` consume the same policy.
- Invalid non-boolean values fail with `PptxDocument output compression must be a boolean` before diagnostics or OPC write.
- Direct OPC invalid values fail with `Package compression must be a boolean`.
- Browser `stream()` platform rejection remains earlier than compression validation.
- Do not add levels, algorithm strings, per-entry rules, file-name object overloads, output types, ZIP streaming, or Node-only compression code.
- Preserve entry order/date, part bytes, relationships, content types, diagnostics, package journal, MIME contracts, and return types.
- Do not copy PptxGenJS 4.0.1's bug where explicit `write()` output types ignore compression, or its truthy-value coercion.
- Never stage `.pnpm-store/`, dist output, tarballs, browser profiles, temporary consumers, probe pages, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: OPC and SDK core compression policy

**Files:**
- Modify: `packages/opc/src/index.ts`
- Modify: `packages/opc/src/index.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Create: `packages/sdk/src/compression-policy.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes existing `OpcPackage.changed`, original bytes, JSZip generation, `WriteBaseOptions`, `WriteOptions<T>`, and private `PptxDocument.#writeBytes()`.
- Produces `PackageWriteOptions`, `OpcPackage.write(options?)`, public `compression?: boolean`, and one canonical policy shared by every output method.
- Leaves `convertWriteOutput()`, `OUTPUT_TYPES`, MIME rules, Node Readable chunking, and input APIs unchanged.

- [ ] **Step 1: Add a stable ZIP central-directory test helper**

In OPC and SDK compression tests, parse the EOCD instead of relying on JSZip private fields. Use this exact shape locally in each test file so no production helper becomes public:

```ts
function zipCompressionMethods(bytes: Uint8Array): number[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.byteLength - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x0605_4b50) eocd -= 1;
  if (eocd < 0) throw new Error('ZIP EOCD not found');
  const entries = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const methods: number[] = [];
  for (let index = 0; index < entries; index += 1) {
    if (view.getUint32(offset, true) !== 0x0201_4b50) {
      throw new Error('ZIP central directory entry not found');
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (!name.endsWith('/')) methods.push(view.getUint16(offset + 10, true));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return methods;
}
```

- [ ] **Step 2: Write failing OPC policy tests**

Extend `packages/opc/src/index.test.ts`. Create packages with fixed `entryDate: new Date('1980-01-01T00:00:00.000Z')` and a repetitive `/data.xml` payload. Require:

```ts
const defaultBytes = await pkg.write();
const storedBytes = await pkg.write({ compression: false });
const deflatedBytes = await pkg.write({ compression: true });
expect(defaultBytes).toEqual(storedBytes);
expect(new Set(zipCompressionMethods(storedBytes))).toEqual(new Set([0]));
expect(new Set(zipCompressionMethods(deflatedBytes))).toEqual(new Set([8]));
expect(deflatedBytes.byteLength).toBeLessThan(storedBytes.byteLength);
expect(await pkg.write({ compression: true })).toEqual(deflatedBytes);
await expect(OpcPackage.open(storedBytes)).resolves.toBeInstanceOf(OpcPackage);
await expect(OpcPackage.open(deflatedBytes)).resolves.toBeInstanceOf(OpcPackage);
```

Open a DEFLATE fixture without mutation. Require omitted and explicit undefined return original bytes, explicit false regenerates STORE, and explicit true contains DEFLATE. Open a STORE fixture and require omitted preserves it while explicit true changes the method to DEFLATE. Pass `'yes'`, `1`, `null`, `{}`, and `new Boolean(true)` through a runtime cast; require the exact OPC `TypeError` and unchanged mutations.

- [ ] **Step 3: Write failing SDK cross-surface tests**

Create `packages/sdk/src/compression-policy.test.ts`. Build a one-slide document with text `Compression policy 你好` plus a 131,072-byte repeated custom part. Decode `write()` output types with the existing public representation rules. Require default equals explicit false, false uses only method 0, true uses only method 8, true is smaller, both reopen with the same text, and all six true outputs decode to the exact same true bytes.

Collect `await document.stream({ compression: true })`, read `await document.writeBlob({ compression: true })`, and write to a temporary path with `await document.writeFile(path, { compression: true })`; require all three equal `write({ compression: true })`. Repeat STORE method assertions for false. Capture diagnostics and package journal around every path.

For an unchanged reopened document, require omitted output equals the input while explicit false switches a DEFLATE input to STORE and explicit true switches a STORE input to DEFLATE.

- [ ] **Step 4: Write failing validation-order and type tests**

For each invalid value `['true', 1, 0, null, {}, [], new Boolean(true)]`, create a fresh document, spy on `document.opcPackage.write`, capture diagnostics/journal, and require:

```ts
await expect(document.write({ compression: value } as never)).rejects.toThrow(
  new TypeError('PptxDocument output compression must be a boolean'),
);
expect(write).not.toHaveBeenCalled();
expect(document.diagnostics).toEqual(diagnostics);
expect(document.opcPackage.mutations).toEqual(journal);
```

In SDK and aggregate-root compile-only branches, require `compression` on `WriteBaseOptions`, generic `WriteOptions`, `write`, `stream`, `writeFile`, `writeBlob`, and `download`; add `@ts-expect-error` assertions for string/number/null selectors. Retain `stream({ outputType })` rejection.

- [ ] **Step 5: Add failing public PptxGenJS evidence**

In `packages/pptxgenjs-adapter/src/index.test.ts`, use only public `addSlide`, `addText`, `stream`, and `write`. Require PptxGenJS stream false/default uses STORE, stream true contains DEFLATE and is smaller, but `write({ outputType: 'uint8array', compression: false/true })` outputs are equal and STORE. Require native `write({ outputType: 'uint8array', compression: true })` contains DEFLATE and false contains STORE. This locks legal API parity and documents the upstream explicit-output defect without comparing producer bytes.

- [ ] **Step 6: Run RED**

```sh
node_modules/.bin/vitest run packages/opc/src/index.test.ts packages/sdk/src/compression-policy.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "compression" --reporter=dot
node_modules/.bin/tsc -b --pretty false
```

Expected: typecheck/tests fail because `compression` and `PackageWriteOptions` do not exist and OPC still always DEFLATEs changed packages.

- [ ] **Step 7: Implement OPC policy minimally**

Add before `OpcPackage`:

```ts
export interface PackageWriteOptions {
  readonly compression?: boolean;
}
```

Replace `OpcPackage.write()` with:

```ts
async write(options: PackageWriteOptions = {}): Promise<Uint8Array> {
  const compression = resolvePackageCompression(options.compression);
  if (!this.changed && compression === undefined) return new Uint8Array(this.#original);
  return this.#zip.generateAsync({
    type: 'uint8array',
    compression: compression === true ? 'DEFLATE' : 'STORE',
    ...(compression === true ? { compressionOptions: { level: 6 } } : {}),
    platform: 'DOS',
  });
}
```

Add a private module function that accepts only undefined or primitive boolean and otherwise throws `new TypeError('Package compression must be a boolean')`. Do not change ZIP file insertion or original-byte storage.

- [ ] **Step 8: Implement SDK validation and propagation minimally**

Add `readonly compression?: boolean` to `WriteBaseOptions`. At the start of `#writeBytes`, before compatibility selection and diagnostics, normalize with:

```ts
function resolveOutputCompression(value: unknown): boolean | undefined {
  if (value !== undefined && typeof value !== 'boolean') {
    throw new TypeError('PptxDocument output compression must be a boolean');
  }
  return value;
}
```

Pass the normalized value to `this.opcPackage.write({ compression })`. Do not validate before the existing Node-only `stream()` preflight and do not change `write()` output-type validation order.

- [ ] **Step 9: Run focused, type, and bundle gates**

```sh
node_modules/.bin/vitest run packages/opc/src/index.test.ts packages/sdk/src/compression-policy.test.ts packages/sdk/src/write-output.test.ts packages/sdk/src/node-readable-stream.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
git diff --check
```

Require focused tests green, both typechecks green, both bundles green, and no new static Node dependency in `dist/browser.js`. Remove generated untracked artifacts before staging.

- [ ] **Step 10: Review, commit, push, and verify**

Review the full decision matrix, strict primitive validation, output-type/compression orthogonality, original fast path, explicit regeneration, deterministic bytes, all output surfaces, browser stream ordering, public upstream evidence, and unrelated-diff absence. Stage only the six listed source/test files, commit `feat: add presentation compression policy`, push, fetch, and require divergence `0 0`.

---

### Task 2: Actual-tarball Node, browser, declarations, CLI, and Chrome proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes the packed public compression selector and central-directory method helper.
- Produces stable `compressionPolicy: true` plus methods, sizes, cross-surface equality, reopen, fast-path, invalid-isolation, and download evidence.
- Preserves all existing packed and browser smoke fields.

- [ ] **Step 1: Extend actual declaration verification**

Require packed declarations contain `readonly compression?: boolean` in `WriteBaseOptions`, `PackageWriteOptions`, and `write(options?: PackageWriteOptions)`. Retain exact generic `WriteOptions`/`WriteOutput` and structural stream checks. Reject algorithm/level fields and new Node references.

- [ ] **Step 2: Extend installed Node runtime smoke**

In the generated installed consumer, add a central-directory parser equivalent to Task 1. Create a repetitive document and produce default, false, and true bytes. Decode all six `write({ outputType, compression: true })` values, collect a true stream, read a true presentation Blob, and read a true file. Record:

```ts
const compressionPolicyState = {
  defaultEqualsFalse: true,
  storeMethods: [0],
  deflateMethods: [8],
  deflateSmaller: true,
  sixOutputEquality: true,
  streamEquality: true,
  blobEquality: true,
  fileEquality: true,
  reopenTitles: ['Compression policy 你好', 'Compression policy 你好'],
  unchangedOriginal: true,
  explicitStore: true,
  explicitDeflate: true,
  mutationIsolation: true,
};
```

Use one DEFLATE input and one STORE input to prove omitted-original and explicit opposite-policy regeneration. Add `compressionPolicy` and state to output JSON and fail the smoke immediately if any field differs.

- [ ] **Step 3: Extend installed TypeScript and browser-condition consumers**

In `smoke.ts`, compile every public method with true/false and reject string/number/null selectors. In `browser-smoke.mjs`, create a repetitive deck, require default/false STORE, true DEFLATE, true smaller, `write()`/`writeBlob()` equality, successful reopen, exact invalid error, unchanged diagnostics/journal, and later successful write. Retain browser conditional resolution and no-static-Node scans.

- [ ] **Step 4: Extend real Chrome and download evidence**

Add the same browser compression state to the initial `page.evaluate()` and expected result. Trigger a separate `document.download('compression-policy.pptx', { compression: true })`, collect its Playwright download stream outside the page, require suggested filename, parse method 8, then pass its base64 bytes back into the page and require reopen title `Chrome compression download`. Add the result to `compressionPolicyState`; require zero validation/console/page/network errors.

- [ ] **Step 5: Run complete release gates**

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

Serve the actual packed browser module over loopback, execute `scripts/playwright-browser-smoke.js` in installed Google Chrome, and retain result JSON. Require all resource responses 200 and console/page/network errors 0. Record full test totals, performance time, tarball file count, SHA-256, installed Node/type/browser/CLI state, and Chrome state. Move evidence to a fresh `/tmp/pptx-compression-policy-artifacts.*` directory and remove the workspace tarball.

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-package-only imports, declaration fidelity, method 0/8 evidence, size reduction, six outputs, stream/file/Blob/download equality, reopen, unchanged-original and explicit regeneration, invalid failure isolation, stable JSON, zero Chrome errors, and artifact scope. Stage only the two scripts, commit `test: verify packed compression policy`, push, fetch, and require divergence `0 0`.

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
- Consumes final source, packed, and Chrome evidence from Tasks 1–2.
- Produces public compression examples, exact default/fast-path rules, deliberate upstream divergence, release evidence, and next-item ordering.

- [ ] **Step 1: Document the public selector and matrix**

Add concise examples using `write({ outputType: 'uint8array', compression: true })`, `stream({ compression: true })`, `writeFile(path, { compression: true })`, and browser `writeBlob({ compression: false })`. State omitted/false STORE, true DEFLATE level 6, and consistent behavior across every output surface.

- [ ] **Step 2: Document lossless editing and upstream divergence**

Explain the sole omitted-policy exception for opened unchanged packages, and that explicit false/true forces regeneration. Record PptxGenJS 4.0.1's explicit-output bug and truthy coercion as deliberate non-copies while confirming legal boolean parity.

- [ ] **Step 3: Update compatibility, progress, and changelog**

Move compression from partial to supported. Add exact final test totals, performance, tarball count/SHA-256, Node/browser/Chrome evidence, STORE/DEFLATE sizes/methods, and the next item `SchemeColor` plus remaining runtime helpers. Update overall parity only if the completed coverage justifies it.

- [ ] **Step 4: Review documentation integrity**

```sh
git diff --check
git diff --name-only
rg -n 'compression|STORE|DEFLATE|fast path|PptxGenJS|SchemeColor' README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

Require exactly six modified documentation files, balanced Markdown fences, no placeholder terms, no stale current statement saying compression is pending, historical checkpoints preserved, and current evidence matching Task 2 exactly.

- [ ] **Step 5: Commit, push, verify, and continue**

Stage only the six documentation files, commit `docs: document presentation compression policy`, push, fetch, and require divergence `0 0`. Report completed item, remaining items, and updated overall progress; immediately continue to the `SchemeColor` runtime-helper design.
