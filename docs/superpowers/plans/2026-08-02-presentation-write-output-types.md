# Presentation `write({ outputType })` Return Semantics Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with independent review, commit, push, fetch, and remote-divergence verification.

**Goal:** Make all six public `OUTPUT_TYPES` tokens select real, byte-identical `PptxDocument.write()` return formats with source-compatible defaults and exact Node/browser behavior.

**Architecture:** Keep `OpcPackage.write()` as the single canonical `Uint8Array` producer. Add one SDK-internal conversion module, a conditional public return type, and a private validated byte-writing path so `write()` can select a representation while `writeBlob()`, `writeFile()`, and `download()` preserve their contracts. Runtime-only dynamic Node Buffer loading keeps the browser bundle free of static Node imports.

**Tech Stack:** TypeScript strict mode, Vitest, JSZip-backed OPC package bytes, PptxGenJS 4.0.1 public runtime, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, installed CLI, and real Google Chrome.

## Global Constraints

- Exact tokens/order remain `arraybuffer`, `base64`, `binarystring`, `blob`, `nodebuffer`, `uint8array`.
- Omitted options, `{}`, validation-only options, and runtime explicit undefined continue to return plain `Uint8Array`.
- Explicit blob output uses `application/zip`; existing `writeBlob()` keeps the presentation format MIME.
- Node `nodebuffer` passes `Buffer.isBuffer()`; browser `nodebuffer` rejects with `nodebuffer is not supported by this platform` before validation or ZIP generation.
- Public `WriteOutput<'nodebuffer'>` is browser-safe `Uint8Array`; runtime proof distinguishes Buffer from plain Uint8Array.
- Base64 is raw standard base64 without data-URI prefix or whitespace. Binary string has one unsigned byte per code unit.
- Every successful representation decodes to the exact canonical bytes and reopens as a valid presentation.
- Do not change `OpcPackage.write()`, compression policy, ZIP ordering/date/level, diagnostics content, package mutation, stream behavior, file/download returns, or `PptxInput`.
- Do not include `STREAM`, add a Buffer global dependency, add static `node:` imports to browser output, or copy the PptxGenJS `write({})` defect.
- Never stage `.pnpm-store/`, dist output, tarballs, browser profiles, temporary consumers, probe pages, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Core conditional API, conversion, defaults, and public conformance

**Files:**
- Create: `packages/sdk/src/write-output.ts`
- Create: `packages/sdk/src/write-output.test.ts`
- Modify: `packages/sdk/src/output-type.ts`
- Modify: `packages/sdk/src/output-type.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes `OUTPUT_TYPES`, `OutputType`, existing validation/diagnostics, and canonical `OpcPackage.write(): Promise<Uint8Array>`.
- Produces `WriteOutput<TOutputType>`, `WriteBaseOptions`, generic `WriteOptions<TOutputType>`, and generic `PptxDocument.write<TOutputType>()`.
- Produces internal `resolveWriteOutputType()` and `convertWriteOutput()` with no root export.
- Preserves no-argument write and all convenience-method runtime results.

- [ ] **Step 1: Write failing pure conversion and conditional-type tests**

Create `packages/sdk/src/write-output.test.ts`. Use bytes containing low/high values and enough data to cross the binary chunk boundary:

```ts
import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { convertWriteOutput, resolveWriteOutputType } from './write-output.js';

const bytes = Uint8Array.from(
  { length: 65_541 },
  (_, index) => [0x00, 0x7f, 0x80, 0xff, index & 0xff][index % 5]!,
);

it('converts canonical bytes into all six Node output representations', async () => {
  const arraybuffer = await convertWriteOutput(bytes, 'arraybuffer');
  const base64 = await convertWriteOutput(bytes, 'base64');
  const binarystring = await convertWriteOutput(bytes, 'binarystring');
  const blob = await convertWriteOutput(bytes, 'blob');
  const nodebuffer = await convertWriteOutput(bytes, 'nodebuffer');
  const uint8array = await convertWriteOutput(bytes, 'uint8array');

  arraybuffer satisfies ArrayBuffer;
  base64 satisfies string;
  binarystring satisfies string;
  blob satisfies Blob;
  nodebuffer satisfies Uint8Array;
  uint8array satisfies Uint8Array;

  expect(new Uint8Array(arraybuffer)).toEqual(bytes);
  expect(Uint8Array.from(Buffer.from(base64, 'base64'))).toEqual(bytes);
  expect(Uint8Array.from(binarystring, (value) => value.charCodeAt(0))).toEqual(bytes);
  expect(new Uint8Array(await blob.arrayBuffer())).toEqual(bytes);
  expect(blob.type).toBe('application/zip');
  expect(Buffer.isBuffer(nodebuffer)).toBe(true);
  expect(nodebuffer).toEqual(bytes);
  expect(uint8array).toEqual(bytes);
  expect(Buffer.isBuffer(uint8array)).toBe(false);
});
```

Add resolver tests for omitted/undefined → `uint8array`, every catalog member, and strict rejection of `STREAM`, `buffer`, `BLOB`, null, number, object, and symbol. Require the generic result assignments above to compile without casts.

- [ ] **Step 2: Extend failing catalog-derived return typing**

In `packages/sdk/src/output-type.ts`, the test should expect this public type, but do not implement it yet:

```ts
type ArrayBufferOutput = WriteOutput<'arraybuffer'>;
type StringOutput = WriteOutput<'base64' | 'binarystring'>;
type BlobOutput = WriteOutput<'blob'>;
type ByteOutput = WriteOutput<'nodebuffer' | 'uint8array'>;
```

Use assignments inside the existing compile-only block to require `ArrayBuffer`, `string`, `Blob`, and `Uint8Array` respectively, and negative assignments for wrong targets.

- [ ] **Step 3: Add failing document lifecycle and source-compatibility coverage**

In `packages/sdk/src/index.test.ts`, add one focused test that creates a presentation, adds text, captures its package mutation journal, obtains canonical default bytes, and requests all six tokens. Require:

```ts
const defaultOutput = await document.write();
const emptyOutput = await document.write({});
const permissiveOutput = await document.write({ mode: 'permissive' });
const arraybuffer = await document.write({ outputType: 'arraybuffer' });
const base64 = await document.write({ outputType: 'base64' });
const binarystring = await document.write({ outputType: 'binarystring' });
const blob = await document.write({ outputType: 'blob' });
const nodebuffer = await document.write({ outputType: 'nodebuffer' });
const uint8array = await document.write({ outputType: 'uint8array' });
```

Require default/empty/permissive/uint8array to be non-Buffer `Uint8Array`; decode all representations and compare them with `defaultOutput`; require the explicit Blob MIME to be `application/zip`; reopen all decoded outputs and require the same slide text; require the mutation journal to remain unchanged by conversions. Separately require `writeBlob()` to retain the PPTX format MIME and `writeFile()` bytes to remain canonical.

Add runtime casts that pass invalid output values and assert `TypeError` before `opcPackage.write` or diagnostics changes. Use a temporary method spy only inside the test and restore it in `finally`.

- [ ] **Step 4: Add failing exact generic/root type coverage**

Import `WriteBaseOptions`, `WriteOptions`, and `WriteOutput` in both SDK/root test surfaces. Add compile assertions:

```ts
document.write() satisfies Promise<Uint8Array>;
document.write({ mode: 'permissive' }) satisfies Promise<Uint8Array>;
document.write({ outputType: 'arraybuffer' }) satisfies Promise<ArrayBuffer>;
document.write({ outputType: 'base64' }) satisfies Promise<string>;
document.write({ outputType: 'binarystring' }) satisfies Promise<string>;
document.write({ outputType: 'blob' }) satisfies Promise<Blob>;
document.write({ outputType: 'nodebuffer' }) satisfies Promise<Uint8Array>;
document.write({ outputType: 'uint8array' }) satisfies Promise<Uint8Array>;
```

Require `WriteOptions<'blob'>` to accept only blob, a `WriteOptions<OutputType>` variable to produce `Promise<WriteOutput<OutputType>>`, and `WriteBaseOptions` to remain accepted by `writeBlob()`/`writeFile()`. Keep `STREAM`, unknown tokens, and direct `outputType` on convenience object literals as negative cases.

- [ ] **Step 5: Add failing public-only PptxGenJS return-kind conformance**

In `packages/pptxgenjs-adapter/src/index.test.ts`, widen the local public interface to accept `OutputType` and return `unknown`; do not import PptxGenJS private types. Convert each result into a stable public kind:

```ts
function publicOutputKind(value: unknown): string {
  if (Buffer.isBuffer(value)) return 'nodebuffer';
  if (value instanceof ArrayBuffer) return 'arraybuffer';
  if (value instanceof Blob) return `blob:${value.type}`;
  if (value instanceof Uint8Array) return 'uint8array';
  return typeof value;
}
```

For all six tokens, create equivalent one-slide PptxGenJS/native decks, call only public `write()`, and require matching kinds. Require both strings to decode to valid packages, both blobs to use `application/zip`, nodebuffer values to pass `Buffer.isBuffer()`, uint8array values not to be Buffer, and every result to reopen through public native input after decoding.

- [ ] **Step 6: Run RED**

```sh
node_modules/.bin/vitest run packages/sdk/src/write-output.test.ts packages/sdk/src/output-type.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "write output|output return|PptxGenJS output" --reporter=dot
node_modules/.bin/tsc -b --pretty false
```

Expected: tests/typecheck fail because `WriteOutput`, generic options, internal conversion, and selected return semantics do not exist.

- [ ] **Step 7: Implement the conditional type and internal converter**

Add to `packages/sdk/src/output-type.ts`:

```ts
export type WriteOutput<TOutputType extends OutputType = OutputType> =
  TOutputType extends 'arraybuffer' ? ArrayBuffer
    : TOutputType extends 'base64' | 'binarystring' ? string
      : TOutputType extends 'blob' ? Blob
        : Uint8Array;
```

Create `packages/sdk/src/write-output.ts` with these exact responsibilities:

```ts
export function resolveWriteOutputType(value: unknown): OutputType {
  const outputType = value === undefined ? 'uint8array' : value;
  if (typeof outputType !== 'string' ||
      !OUTPUT_TYPES.includes(outputType as OutputType)) {
    throw new TypeError('PptxDocument.write() received an unsupported outputType');
  }
  if (outputType === 'nodebuffer' && !isNodeRuntime()) {
    throw new Error('nodebuffer is not supported by this platform');
  }
  return outputType as OutputType;
}
```

Implement `convertWriteOutput<T>()` as one switch. Use `Uint8Array.from(bytes).buffer` for standalone ArrayBuffer; a 32,768-byte `String.fromCharCode` chunk loop for binary string; a chunked standard-alphabet triplet encoder for base64; `new Blob([standaloneArrayBuffer], { type: 'application/zip' })`; a runtime-only `import(['node:', 'buffer'].join(''))` whose structural constructor returns `Uint8Array`; and unchanged bytes for uint8array. Localize unavoidable conditional-return casts inside this module.

- [ ] **Step 8: Refactor document writing around one private byte writer**

In `packages/sdk/src/index.ts`:

1. export `WriteOutput` from `output-type.ts`;
2. add `WriteBaseOptions` with the two existing validation fields;
3. make `WriteOptions<TOutputType extends OutputType = 'uint8array'>` extend it and add optional `outputType`;
4. move the current entire validation/diagnostics/OPC body into `#writeBytes(options: WriteBaseOptions): Promise<Uint8Array>`;
5. implement generic public `write()` as resolve → `#writeBytes()` → convert;
6. call `#writeBytes()` directly from `writeFile()` and `writeBlob()`;
7. leave `download()` delegating to unchanged `writeBlob()`.

Do not modify validation ordering inside `#writeBytes`, exception classes, `OpcPackage`, or format MIME lookup.

- [ ] **Step 9: Run focused and package gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/write-output.test.ts packages/sdk/src/output-type.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "write output|output return|PptxGenJS output" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/write-output.test.ts packages/sdk/src/output-type.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 10: Review, commit, push, and verify**

Review default source compatibility, one canonical byte writer, strict token validation before diagnostics, six exact returns, high-byte strings, standalone buffers, Blob MIME separation, Buffer/plain-Uint8Array distinction, no static Node import, PptxGenJS public-only comparison, package mutation isolation, and unrelated-diff absence. Stage only the eight listed files, commit `feat: support write output types`, push, fetch, and require divergence `0 0`.

---

### Task 2: Actual-tarball Node, declarations, browser, CLI, and Chrome proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes installed generic `write()` and all six token conversions.
- Produces stable `writeOutputTypes: true` plus constructor/MIME/byte-equality/reopen state in packed Node and real Chrome output.
- Proves actual browser bundle rejects Node Buffer output without static Node imports.

- [ ] **Step 1: Extend packed declaration verification**

Normalize `dist/types/sdk/output-type.d.ts` and `dist/types/sdk/index.d.ts`, then require the conditional `WriteOutput` branches, exported `WriteBaseOptions`, generic `WriteOptions<TOutputType extends OutputType = 'uint8array'>`, optional selector, and generic `write<TOutputType extends OutputType = 'uint8array'>` return. Reject widened `Promise<string | ArrayBuffer | Blob | Uint8Array>` and any static `node:buffer` declaration dependency.

- [ ] **Step 2: Extend installed Node runtime smoke**

Create a packed document with non-ASCII text. Capture default bytes and request six outputs. Decode each with local smoke helpers and require:

```ts
const writeOutputTypeState = {
  defaultKind: 'uint8array',
  arraybufferKind: 'arraybuffer',
  base64Kind: 'string',
  binarystringKind: 'string',
  blobKind: 'blob',
  blobType: 'application/zip',
  nodebufferKind: 'nodebuffer',
  uint8arrayKind: 'uint8array',
  byteEquality: true,
  reopenTitles: Array(6).fill('Packed output types 你好'),
  writeBlobType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};
```

Use `Buffer.isBuffer()` only in the installed Node consumer, require explicit uint8array not to be Buffer, keep current catalog state, and include `writeOutputTypes`/state in stable JSON.

- [ ] **Step 3: Extend installed TypeScript consumer**

Import `WriteBaseOptions`, `WriteOptions`, and `WriteOutput`. Add the same eight `satisfies Promise<...>` assertions as source tests, one dynamic union assertion, positive generic option assignments, and negative `STREAM`, unknown, wrong literal generic, and selector-on-convenience object-literal cases. The consumer must compile without importing Node types for the nodebuffer branch.

- [ ] **Step 4: Extend browser-condition and real-Chrome checks**

In the browser condition consumer, retain resolved `dist/browser.js`, no static `node:` scan, existing `writeBlob()` reopen, and catalog checks.

In `scripts/playwright-browser-smoke.js`, create one deck and obtain canonical uint8array. Request arraybuffer/base64/binarystring/blob/uint8array, decode inside Chrome, compare all bytes, reopen all five, and require explicit Blob `application/zip`. Request nodebuffer and capture exact error name/message; verify the failed request did not alter diagnostics or the mutation journal. Return `writeOutputTypes: true` and a compact state matching the Node field names plus `nodebufferError`.

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

Serve the actual extracted tarball browser module and fixture over loopback, execute `scripts/playwright-browser-smoke.js` in installed Google Chrome, and record the returned JSON. Require all resource responses 200 and console/page/network errors 0. Record full test totals, performance time, tarball file count, SHA-256, installed Node/type/browser/CLI state, and Chrome state. Move retained evidence to a fresh `/tmp/pptx-write-output-types-artifacts.*` directory and leave no generated tracked files.

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-package-only imports, exact declarations, six Node results, five browser results plus exact nodebuffer rejection, high-byte-safe decoders, byte equality, reopen, MIME separation, no static Node import, stable JSON, zero Chrome errors, and clean artifact scope. Stage only the two scripts, commit `test: verify packed write output types`, push, fetch, and require divergence `0 0`.

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
- Produces accurate public examples, compatibility status, release evidence, remaining work, and next-item ordering.

- [ ] **Step 1: Document the selected-return API in both package READMEs and API docs**

Add one concise example showing default Uint8Array, explicit ArrayBuffer/base64/binarystring/Blob/nodebuffer/uint8array, and type inference. State explicit Blob uses `application/zip`, `writeBlob()` retains presentation MIME, nodebuffer is Node-only, and default remains Uint8Array. Do not claim stream or compression support.

- [ ] **Step 2: Update compatibility, progress, and changelog**

Move six actual write return semantics from partial to supported. Keep `stream()` and compression explicitly pending. Add a completed `write({ outputType })` stage with exact final test totals, performance, tarball count/SHA-256, packed checks, Chrome results, intentional default difference from PptxGenJS, and next item Node readable stream.

- [ ] **Step 3: Review documentation integrity**

```sh
git diff --check
git diff --name-only
rg -n "output type|outputType|stream|compression" README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

Require exactly six modified documentation files, balanced Markdown fences, no placeholder terms, no stale statement saying all output types remain unsupported, historical test numbers preserved in their own checkpoints, and current evidence matching Task 2.

- [ ] **Step 4: Commit, push, and verify**

Stage only the six documentation files, commit `docs: document write output types`, push, fetch, and require divergence `0 0`. Report completed item, remaining items, and overall progress; immediately continue to the Node readable stream design.
