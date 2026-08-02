# Presentation Node Readable Stream 设计

日期：2026-08-02  
状态：已批准实施

## 1. 目标

为 `PptxDocument` 增加真正的 Node.js readable 输出，使生成的演示文稿可以直接 pipe 到文件、HTTP response、压缩/加密 transform 或其他 Node writable，同时保留 async iteration。该 API 必须复用当前 validation 与 canonical ZIP bytes，不改变 `write()`、`writeBlob()`、`writeFile()` 或 `download()` 的既有契约。

本项只完成 Node readable contract。它不实现 ZIP streaming、constant-memory generation、compression selector、Web `ReadableStream` 输出、writable-target convenience API 或 `STREAM` output token。

## 2. 当前状态与上游证据

Native 当前具有：

- `write({ outputType })` 的六种非流式表示；
- Node `writeFile()`；
- browser `writeBlob()` / `download()`；
- path、Web `ReadableStream` 和 async iterable 输入，其中 Node Readable 通过 async iterable contract 打开。

PptxGenJS 4.0.1 的 public declaration 暴露 `stream(props?: WriteBaseProps): Promise<string | ArrayBuffer | Blob | Uint8Array>`。真实 Node probe 和公开实现表明它把内部 `STREAM` 分支交给 JSZip `generateAsync({ type: 'nodebuffer' })`，因此 runtime 返回 Buffer，而不是 Node Readable。Native 已通过 `write({ outputType: 'nodebuffer' })` 精确覆盖这个 byte-result 能力；本项实现方法名称所表达的真正流式消费能力，并记录该 deliberate divergence。

## 3. 方案比较

### A. 复制 PptxGenJS，`stream()` 返回 Buffer

调用名称最接近上游，但与 `write({ outputType: 'nodebuffer' })` 完全重复，也无法 pipe、逐块消费或遵守 readable backpressure。拒绝。

### B. 返回 Web `ReadableStream<Uint8Array>`

浏览器与 Node 18+ 都可使用，类型也不依赖 `@types/node`，但 Node 生态的 `pipe()`、events、`createWriteStream()` 与传统 HTTP response 需要额外转换，不满足本阶段明确的 Node readable 目标。留作未来独立 portable-stream 需求，不在本项实施。

### C. 动态加载 `node:stream` 并返回真实 `Readable`（采用）

`stream()` 在 Node 中完成 validation 与 canonical ZIP generation，再以固定 64 KiB binary chunks 建立 `Readable.from(..., { objectMode: false })`。Node runtime 获得真正的 `Readable`，browser bundle 不含 static Node import。公开 declaration 使用小型 structural `PptxNodeReadableStream`，避免统一 browser type entry 引入 `node:stream` 或 `@types/node` contract。

## 4. 公共 API

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

export class PptxDocument {
  stream(options?: WriteBaseOptions): Promise<PptxNodeReadableStream>;
}
```

`stream()` 采用与 `writeBlob()` / `writeFile()` 相同的 `WriteBaseOptions`，因此当前接受 `compatibility` 与 `mode`。后续 compression 小项可以在一个共享位置扩展 base options，不需要改变 stream 方法形状。

不增加 `STREAM` 到 `OutputType` 或 `OUTPUT_TYPES`。`write({ outputType: 'STREAM' })` 继续拒绝，避免返回 type union 混入环境专属 stream。

## 5. Runtime 与数据语义

调用顺序固定为：

1. 检查 Node runtime；browser/非 Node 环境在 validation 和 ZIP write 前拒绝；
2. 调用现有 `#writeBytes(options)` 一次，复用 compatibility diagnostics、strict/permissive gate 与 canonical `OpcPackage.write()`；
3. 动态加载 `node:stream`；
4. 以 64 KiB 上限的非 object-mode chunks 暴露 canonical bytes；
5. 返回真实 Node `Readable`。

所有 stream chunks 按顺序拼接后必须与同一状态下 `write()` byte-identical，并能由 `PptxDocument.open()` 重开。建立 stream 后对 document 的修改不改变已捕获输出。消费或 destroy stream 不修改 package、diagnostics 或 mutation journal。

当前 OPC/ZIP writer 仍产生完整 `Uint8Array`，所以 `stream()` 降低下游一次性接口耦合并提供 backpressure-aware delivery，但不降低 ZIP generation 的峰值内存，也不宣称 time-to-first-byte streaming。chunks 使用原 canonical byte buffer 的 non-overlapping views；stream 拥有这次 write 的私有 bytes，调用方无法修改 document/package state。

## 6. 平台与错误

- browser 调用精确拒绝：`PptxDocument.stream() is only supported in Node.js`；
- 环境拒绝发生在 validation、diagnostics replacement、OPC write 和动态 import 前；
- strict validation error 继续由 `stream()` Promise 拒绝，不创建半有效 readable；
- permissive mode 与 `write()` 一样返回 stream 并保留 diagnostics；
- Node module load failure 由 Promise 原样拒绝；
- consumer-side destination/error/destroy 使用 Node Readable 标准传播，不包装或吞掉；
- output 建立后没有 background package mutation，因此不存在晚到的 package error。

公开 structural type 覆盖 binary-readable 的核心消费面：pipe、async iteration、data/end/close/error、pause/resume/read/destroy。runtime 对象仍是完整 Node `Readable`；本项不重新声明整个 Node stream 类型体系。

## 7. 模块边界

- `packages/sdk/src/index.ts` 拥有公开 interface、`stream()` 与 Node-only dynamic adapter；
- OPC、model、codecs、validator 和 lossless XML 不修改；
- aggregate root 继续通过 `export * from '@pptx/sdk'` 自动复用类型和方法；
- browser bundle 不增加 static `node:` import、Buffer global 或 eager runtime detection；
- fixed chunking helper 保持 SDK private，不成为可配置 public policy。

## 8. 验证策略

### Source 与类型

- `stream()` 返回真实 Node `Readable`，`readableObjectMode === false`；
- async iteration、`data/end` events 与 `pipe()` 三种消费方式均得到 byte-identical canonical bytes；
- 大于 64 KiB 的 opaque part 锁定 multi-chunk 顺序、chunk 上限和可重开；
- captured state、package journal、diagnostics、destroy isolation 与 repeated stream/write 覆盖；
- strict validation rejection 与 permissive success 对齐 `write()`；
- simulated browser 在 package write 前得到 exact error；
- TypeScript 锁定 `Promise<PptxNodeReadableStream>`、`AsyncIterable<Uint8Array>`、pipe return inference 和 `WriteBaseOptions`，拒绝 outputType。

### PptxGenJS conformance 与 deliberate divergence

- public-only probe 锁定 PptxGenJS 4.0.1 `stream()` runtime 为 Buffer、bytes 可重开；
- native `write({ outputType: 'nodebuffer' })` 对等其实际 byte result；
- native `stream()` 明确验证为 `Readable` 而非 Buffer，提供上游命名暗示但未实现的能力。

### Packed Node/browser

- actual tarball Node consumer 将 stream 分别 pipe 到 file/writable、async iterate、reopen，并检查 byte equality；
- declarations 包含 structural stream interface 且不引用 `node:stream`、`node:buffer` 或 `NodeJS` namespace；
- browser conditional export 可导入且调用在 ZIP write 前 exact reject；
- real Chrome 检查 exact error、diagnostics/package mutation isolation 和 zero console/page/network errors；
- full Vitest、performance、两套 TypeScript、Node/browser bundles、declarations、tarball count/SHA-256 与 installed CLI 全部记录。

## 9. 实施与提交边界

实施分为三个可独立 review 的 repository change：

1. core API、structural type、chunked Node Readable、source/type/PptxGenJS conformance；
2. actual-tarball Node/browser/declaration/CLI 与 real-Chrome proof；
3. README/API/compatibility/progress/changelog 发布文档。

每个 change 都必须 review、commit、push、fetch，并确认远端同步后再继续。

本项完成标准：Node `stream()` 是真实 binary Readable；三种消费方式 byte-identical 且可重开；validation/platform/destroy 边界明确；browser package 保持纯净；release gates 全绿。完成后下一项进入 compression policy。
