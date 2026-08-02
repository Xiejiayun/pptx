# Presentation Node Readable Stream Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Every repository-changing task ends with independent review, commit, push, fetch, and remote-divergence verification.

**Goal:** Add a real Node.js binary `Readable` output to `PptxDocument.stream()` while preserving canonical bytes, validation behavior, browser purity, and every existing output contract.

**Architecture:** Keep `#writeBytes()` and `OpcPackage.write()` as the single in-memory canonical ZIP producer. After a Node-only preflight, `stream()` dynamically loads `node:stream` and exposes those bytes through `Readable.from()` in fixed 64 KiB non-object-mode chunks. A browser-safe structural `PptxNodeReadableStream` type exposes the core binary-readable surface without static Node imports or Node namespaces.

**Tech Stack:** TypeScript strict mode, Node.js 20 `Readable`/`Writable`, Vitest, PptxGenJS 4.0.1 public runtime probes, tsup Node/browser bundles, generated declarations, actual npm tarball smoke, installed CLI, and real Google Chrome.

## Global Constraints

- Public API is `stream(options?: WriteBaseOptions): Promise<PptxNodeReadableStream>`.
- Runtime output in Node is a real `node:stream` `Readable` with `readableObjectMode === false`.
- Output is delivered in ordered chunks no larger than 65,536 bytes; this size is private and not configurable.
- Concatenated stream bytes must be byte-identical to `write()` for the same document state and reopen successfully.
- Browser/other runtimes reject exactly with `PptxDocument.stream() is only supported in Node.js` before validation, diagnostics replacement, ZIP generation, or dynamic import.
- `STREAM` remains excluded from `OUTPUT_TYPES`, `OutputType`, and `write({ outputType })`.
- Do not change `write()`, `writeBlob()`, `writeFile()`, `download()`, compression, ZIP ordering/date/level, format MIME, or input-stream behavior.
- Do not claim constant-memory ZIP generation or time-to-first-byte streaming; the complete canonical `Uint8Array` still exists before the Readable is returned.
- Browser declarations and bundle must not reference static `node:stream`, `node:buffer`, `NodeJS`, or a Buffer global.
- PptxGenJS 4.0.1 `stream()` returning Buffer is a documented upstream divergence; native `write({ outputType: 'nodebuffer' })` already covers that byte-result behavior.
- Never stage `.pnpm-store/`, dist output, tarballs, browser profiles, temporary consumers, probe pages, or generated decks.
- Every task is reviewed, committed, pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Core Node Readable API, lifecycle, types, and public conformance

**Files:**
- Modify: `packages/sdk/src/index.ts`
- Create: `packages/sdk/src/node-readable-stream.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes `WriteBaseOptions` and the existing private `#writeBytes(options): Promise<Uint8Array>` canonical path.
- Produces public `PptxNodeReadableStream` and `PptxDocument.stream()`.
- Uses private `NODE_STREAM_CHUNK_SIZE`, `isNodeRuntime()`, `chunkPptxBytes()`, and a structural dynamic `NodeStreamModule` only inside SDK.
- Leaves aggregate-root runtime export unchanged because `PptxDocument` flows through existing `export * from '@pptx/sdk'`.

- [ ] **Step 1: Write failing Node stream lifecycle tests**

Create `packages/sdk/src/node-readable-stream.test.ts` with Node `Readable`/`Writable`, `ValidationError`, `PptxDocument`, and `vi`. Use this deterministic collector:

```ts
async function collect(source: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of source) chunks.push(new Uint8Array(chunk));
  const length = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const bytes = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}
```

Create a document with one non-ASCII text shape and a 196,613-byte deterministic xorshift payload at `/custom/stream.bin` with content type `application/octet-stream`. Capture `expected = await document.write()` and package journal. Require:

```ts
const readable = await document.stream();
expect(readable).toBeInstanceOf(Readable);
expect(readable.readableObjectMode).toBe(false);
expect(await collect(readable)).toEqual(expected);
expect(document.opcPackage.mutations).toEqual(journal);
```

Use a second stream with `data`/`end` listeners to record every emitted chunk, require at least two chunks and every `byteLength <= 65_536`, concatenate and compare. Use a third stream piped to a `Writable` whose `write()` copies chunks; require `stream.pipe(destination) === destination`, byte equality, and successful reopen with the original text. Destroy a fourth stream before consumption, await `close`, and require package journal plus a later `write()` to remain unchanged.

- [ ] **Step 2: Write failing captured-state, validation, and platform tests**

Require bytes are captured when `await document.stream()` resolves: create the stream, then add another slide, consume the stream, reopen it, and require only the pre-mutation slide. Require a later `stream()` contains both slides.

For validation, replace `/_rels/.rels` with one office-document relationship targeting `ppt/missing.xml`. Require strict `stream()` rejects with `ValidationError`; require `{ mode: 'permissive' }` resolves to a readable with non-empty bytes and the same diagnostic errors as `write({ mode: 'permissive' })`.

For browser preflight:

```ts
vi.stubGlobal('process', undefined);
const diagnostics = [...document.diagnostics];
const journal = [...document.opcPackage.mutations];
const write = vi.spyOn(document.opcPackage, 'write');
await expect(document.stream()).rejects.toThrow(
  new Error('PptxDocument.stream() is only supported in Node.js'),
);
expect(write).not.toHaveBeenCalled();
expect(document.diagnostics).toEqual(diagnostics);
expect(document.opcPackage.mutations).toEqual(journal);
vi.unstubAllGlobals();
```

Restore the spy/globals in `finally` so failures cannot contaminate other tests.

- [ ] **Step 3: Add failing SDK/root compile-time coverage**

Require `PptxNodeReadableStream` to be exported from SDK and aggregate root. Add compile-only assertions:

```ts
document.stream() satisfies Promise<PptxNodeReadableStream>;
document.stream({ mode: 'permissive' }) satisfies Promise<PptxNodeReadableStream>;
document.stream({ compatibility: 'powerpoint-current' })
  satisfies Promise<PptxNodeReadableStream>;
void document.stream().then((readable) => {
  readable satisfies AsyncIterable<Uint8Array>;
  const destination = { tag: 'destination' } as const;
  readable.pipe(destination) satisfies typeof destination;
  readable.pause().resume().destroy();
});
// @ts-expect-error stream does not consume write output selectors
document.stream({ outputType: 'uint8array' });
```

Require `PptxNodeReadableStream` itself contains no dependency on `NodeJS.ReadableStream` or imported Node types.

- [ ] **Step 4: Add failing PptxGenJS public-runtime evidence**

In `packages/pptxgenjs-adapter/src/index.test.ts`, extend the local public PptxGenJS interface with:

```ts
stream(options?: { readonly compression?: boolean }): Promise<unknown>;
```

Create equivalent one-slide PptxGenJS/native decks. Call only public methods and require:

```ts
const generatedStream = await generated.stream();
expect(Buffer.isBuffer(generatedStream)).toBe(true);

const nativeNodeBuffer = await native.write({ outputType: 'nodebuffer' });
expect(Buffer.isBuffer(nativeNodeBuffer)).toBe(true);

const nativeStream = await native.stream();
expect(nativeStream).toBeInstanceOf(Readable);
expect(Buffer.isBuffer(nativeStream)).toBe(false);
```

Collect/reopen both upstream Buffer and native Readable. Do not inspect PptxGenJS private fields, JSZip, or bundled source. This test records actual behavior; it does not require upstream/native bytes to be identical because package producers differ.

- [ ] **Step 5: Run RED**

```sh
node_modules/.bin/vitest run packages/sdk/src/node-readable-stream.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "Node readable|stream output|public stream" --reporter=dot
node_modules/.bin/tsc -b --pretty false
```

Expected: tests/typecheck fail because `PptxNodeReadableStream` and `PptxDocument.stream()` do not exist.

- [ ] **Step 6: Implement the public structural type**

Add before `WriteBaseOptions` in `packages/sdk/src/index.ts`:

```ts
export interface PptxNodeReadableStream extends AsyncIterable<Uint8Array> {
  readonly destroyed: boolean;
  readonly readable: boolean;
  readonly readableEnded: boolean;
  readonly readableObjectMode: false;
  pipe<TDestination>(
    destination: TDestination,
    options?: { readonly end?: boolean },
  ): TDestination;
  pause(): this;
  resume(): this;
  isPaused(): boolean;
  read(size?: number): Uint8Array | null;
  destroy(error?: Error): this;
  on(event: 'data', listener: (chunk: Uint8Array) => void): this;
  on(event: 'end' | 'close', listener: () => void): this;
  on(event: 'error', listener: (error: Error) => void): this;
  once(event: 'data', listener: (chunk: Uint8Array) => void): this;
  once(event: 'end' | 'close', listener: () => void): this;
  once(event: 'error', listener: (error: Error) => void): this;
}
```

Do not import from `node:stream` or reference the `NodeJS` namespace.

- [ ] **Step 7: Implement Node preflight and chunked readable creation**

Add next to other SDK-private adapters:

```ts
const NODE_STREAM_CHUNK_SIZE = 64 * 1024;

interface NodeStreamModule {
  readonly Readable: {
    from(
      iterable: Iterable<Uint8Array>,
      options: { readonly objectMode: false },
    ): unknown;
  };
}

function isNodeRuntime(): boolean {
  return typeof process !== 'undefined' && Boolean(process.versions?.node);
}

function* chunkPptxBytes(bytes: Uint8Array): Iterable<Uint8Array> {
  for (let offset = 0; offset < bytes.byteLength; offset += NODE_STREAM_CHUNK_SIZE) {
    yield bytes.subarray(offset, Math.min(offset + NODE_STREAM_CHUNK_SIZE, bytes.byteLength));
  }
}
```

Add to `PptxDocument`, adjacent to `write()`/`writeFile()`:

```ts
async stream(options: WriteBaseOptions = {}): Promise<PptxNodeReadableStream> {
  if (!isNodeRuntime()) {
    throw new Error('PptxDocument.stream() is only supported in Node.js');
  }
  const bytes = await this.#writeBytes(options);
  const { Readable } = await loadNodeModule<NodeStreamModule>(['node:stream'].join('/'));
  return Readable.from(chunkPptxBytes(bytes), { objectMode: false })
    as PptxNodeReadableStream;
}
```

Keep platform preflight before `#writeBytes()`. Do not expose chunk size or alter `loadNodeModule()` error behavior for existing path APIs.

- [ ] **Step 8: Run focused and package gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/node-readable-stream.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
rg -n "(?:from|import).*node:stream|NodeJS" packages/pptx/dist/browser.js packages/pptx/dist/types/sdk/index.d.ts
git diff --check
```

The final `rg` must have no matches. Remove generated dist output before staging if it is not ignored.

- [ ] **Step 9: Review, commit, push, and verify**

Review actual `Readable` identity, non-object mode, chunk ordering/size, three consumption modes, captured state, validation parity, exact early browser failure, dynamic-only Node loading, structural declarations, upstream divergence evidence, package/diagnostics isolation, and unrelated-diff absence. Stage only the four listed source/test files, commit `feat: add node readable stream`, push, fetch, and require divergence `0 0`.

---

### Task 2: Actual-tarball Node, declarations, browser, CLI, and Chrome proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes installed `PptxDocument.stream()` and `PptxNodeReadableStream`.
- Produces stable `nodeReadableStream: true` plus constructor/object-mode/chunk/pipe/byte/reopen/isolation evidence in packed Node output.
- Produces stable browser exact-rejection evidence without static Node imports.

- [ ] **Step 1: Extend packed declaration verification**

Normalize `dist/types/sdk/index.d.ts` and require all `PptxNodeReadableStream` fields/methods, plus:

```ts
stream(options?: WriteBaseOptions): Promise<PptxNodeReadableStream>;
```

Reject declaration references to `node:stream`, `node:buffer`, `NodeJS.`, or `Buffer`. Retain all existing write-output declaration checks.

- [ ] **Step 2: Extend installed Node runtime smoke**

In the generated installed consumer, import `Readable` and `Writable` from `node:stream`. Create a one-slide deck with text `Packed node stream 你好`, capture its journal and canonical bytes, then consume separate streams by async iteration and pipe. Record `data` chunk lengths on another stream. Require:

```ts
const nodeReadableStreamState = {
  readable: stream instanceof Readable,
  buffer: Buffer.isBuffer(stream),
  objectMode: stream.readableObjectMode,
  chunks: chunkLengths.length,
  chunkLimit: chunkLengths.every((length) => length <= 65_536),
  byteEquality: true,
  pipeEquality: true,
  reopenTitle: 'Packed node stream 你好',
  mutationIsolation: true,
};
```

Require `readable: true`, `buffer: false`, `objectMode: false`, at least one chunk, and all booleans true. Add `nodeReadableStream: true` plus state to installed JSON. Existing Node/type/browser/CLI checks remain intact.

- [ ] **Step 3: Extend installed TypeScript consumer**

Import `PptxNodeReadableStream`. Add the exact compile assertions from Task 1, including `AsyncIterable<Uint8Array>`, generic pipe destination return inference, pause/resume/destroy chaining, validation options, and `@ts-expect-error` for `{ outputType: 'uint8array' }`. The consumer must compile without Node types being required by the package declaration itself.

- [ ] **Step 4: Extend browser-condition and real-Chrome checks**

Retain the browser entry/no-static-Node scan. In Chrome, create a document, capture diagnostics and mutation journal, spy by replacing only the public package write method if the smoke harness already supports safe temporary replacement, call `await document.stream()`, and capture:

```ts
{
  name: 'Error',
  message: 'PptxDocument.stream() is only supported in Node.js',
}
```

Require diagnostics/journal unchanged and a later `write()` still succeeds/reopens. Return `nodeReadableStream: true` only when exact rejection and isolation pass. Do not attempt to import a Node stream polyfill.

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

Serve the extracted actual tarball browser module and fixture over loopback, execute `scripts/playwright-browser-smoke.js` in installed Google Chrome, and retain result JSON. Require all resource responses 200 and console/page/network errors 0. Record full test totals, performance time, tarball file count, SHA-256, installed Node/type/browser/CLI state, and Chrome state. Move evidence to a fresh `/tmp/pptx-node-readable-stream-artifacts.*` directory and leave no generated workspace tarball.

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-package-only imports, declaration purity, true Readable identity, Buffer distinction, non-object mode, chunk limit, async/pipe byte equality, reopen, mutation isolation, exact Chrome rejection, no polyfill/static Node import, stable JSON, zero Chrome errors, and clean artifact scope. Stage only the two scripts, commit `test: verify packed node readable stream`, push, fetch, and require divergence `0 0`.

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
- Produces public stream examples, honest memory/backpressure boundaries, compatibility status, release evidence, and next-item ordering.

- [ ] **Step 1: Document the Node stream API**

Add a concise `for await` and `pipe(createWriteStream(...))` example to both package READMEs and API docs. State that `stream()` is Node-only, returns a true non-object-mode Readable, accepts `WriteBaseOptions`, chunks at an internal 64 KiB maximum, and is byte-identical to `write()`. State clearly that ZIP generation still buffers the complete package before the Readable is returned.

- [ ] **Step 2: Document upstream divergence and browser behavior**

Explain that PptxGenJS 4.0.1 `stream()` actually returns Buffer; native covers that result through `write({ outputType: 'nodebuffer' })` and gives `stream()` real readable semantics. Record browser exact rejection and direct users to `write()`, `writeBlob()`, or `download()` there. Do not claim Web Readable output.

- [ ] **Step 3: Update compatibility, progress, and changelog**

Move Node readable output from partial to supported. Keep compression explicitly pending. Add a completed stream stage with exact final test totals, performance, tarball count/SHA-256, packed checks, Chrome results, non-constant-memory boundary, and next item compression policy.

- [ ] **Step 4: Review documentation integrity**

```sh
git diff --check
git diff --name-only
rg -n 'Node readable|stream\(\)|compression|constant.memory|Buffer' README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

Require exactly six modified documentation files, balanced Markdown fences, no placeholder terms, no stale current statement saying Node readable output is pending, historical checkpoint numbers preserved, and current evidence matching Task 2 exactly.

- [ ] **Step 5: Commit, push, verify, and continue**

Stage only the six documentation files, commit `docs: document node readable stream`, push, fetch, and require divergence `0 0`. Report completed item, remaining items, and updated overall progress; immediately continue to the compression-policy design.
