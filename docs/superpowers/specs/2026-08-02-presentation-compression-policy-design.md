# Presentation Compression Policy 设计

日期：2026-08-02
状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

为所有 `PptxDocument` 输出路径增加统一的 ZIP compression policy，使调用方可以明确选择 STORE 或 DEFLATE，并使合法的 PptxGenJS 4.0.1 `compression?: boolean` 调用在 Node 与 browser 中获得一致结果。

公开语义固定为：

| `compression` | ZIP policy | 适用输出 |
| --- | --- | --- |
| omitted / `undefined` | STORE；仅未修改的已打开 package 保留原始 bytes | `write`、`stream`、`writeFile`、`writeBlob`、`download` |
| `false` | 强制重新打包为 STORE | 同上 |
| `true` | 强制重新打包为 DEFLATE level 6 | 同上 |

本项只支持 PptxGenJS 的 boolean contract，不增加压缩级别、算法名称、per-entry policy、constant-memory ZIP generation、Web stream、文件名对象 overload 或新的 output type。

## 2. 当前状态与上游证据

Native 当前状态：

- `WriteBaseOptions` 只有 `compatibility` 与 strict/permissive `mode`；
- `write({ outputType })` 的六种表示都来自一次 canonical `Uint8Array`；
- `stream()`、`writeFile()`、`writeBlob()` 与 `download()` 复用同一个 private byte writer；
- `OpcPackage.write()` 对未修改的已打开 package 返回原始 bytes，对创建或修改后的 package 固定使用 DEFLATE level 6；
- 因此调用方目前不能选择 STORE，且 native 默认与 PptxGenJS 声明的 default false 不一致。

锁定依赖 PptxGenJS 4.0.1 的 public declarations、public methods 与真实 probe 表明：

- `WriteBaseProps.compression?: boolean`，文档默认 `false`；
- `stream({ compression: false })` 产生 STORE，`stream({ compression: true })` 产生 DEFLATE；
- `writeFile({ compression })` 和 browser default-Blob branch 使用相同 boolean selector；
- `write({ outputType, compression })` 的实现错误地只把 `outputType` 传给 JSZip，导致合法的 true/false 结果完全相同且都为 STORE；
- runtime 还把任意 truthy 值当作 true。

Native 对等公开的合法 boolean 能力，但不复制显式 output type 忽略 selector 的缺陷，也不接受 string、number、null、object 或 boxed boolean。

## 3. 方案比较

### A. 逐字复制 PptxGenJS 4.0.1 分支行为

`stream()` 与 `writeFile()` 使用 selector，而 `write({ outputType })` 忽略它。这样虽然复制了当前实现，却让同一组选项因输出表示不同而改变 ZIP method，无法形成可靠 API。拒绝。

### B. 保留 native tri-state 默认

省略继续代表当前 DEFLATE/original fast path，只有显式 false 才 STORE。它最少改变既有 bytes，但违背 PptxGenJS 声明的 default false，并让新建文稿的 omitted 与 false 不等价。拒绝。

### C. 统一 boolean policy，并保留无损编辑 fast path（采用）

省略与 false 对创建/修改后的 package 都使用 STORE，true 使用 DEFLATE level 6。唯一例外是打开后完全未修改且省略 selector 时继续返回原始 bytes；这是 native 编辑能力特有的 lossless guarantee，PptxGenJS 没有对应场景。显式 false/true 表示调用方要求 ZIP policy，因此即使 package 未修改也重新打包。

## 4. 公共 API

`WriteBaseOptions` 增加只读 boolean：

```ts
export interface WriteBaseOptions {
  readonly compatibility?: CompatibilityProfile;
  readonly compression?: boolean;
  readonly mode?: 'strict' | 'permissive';
}
```

`WriteOptions<TOutputType>` 继续继承 `WriteBaseOptions`。以下既有方法不改变参数位置或返回类型，只获得同一 compression field：

```ts
write<TOutputType extends OutputType = 'uint8array'>(
  options?: WriteOptions<TOutputType>,
): Promise<WriteOutput<TOutputType>>;

stream(options?: WriteBaseOptions): Promise<PptxNodeReadableStream>;
writeFile(path: string, options?: WriteBaseOptions): Promise<void>;
writeBlob(options?: WriteBaseOptions): Promise<Blob>;
download(fileName?: string, options?: WriteBaseOptions): Promise<void>;
```

Native 保留 positional `writeFile(path, options)` 与 `download(fileName, options)`；不复制 PptxGenJS 的 `{ fileName, compression }` object shape，因为现有 typed API 已覆盖文件名和 compression 能力，改变形状会破坏兼容性。

## 5. OPC policy 与数据流

OPC 层增加最小 public write options：

```ts
export interface PackageWriteOptions {
  readonly compression?: boolean;
}

write(options?: PackageWriteOptions): Promise<Uint8Array>;
```

调用顺序固定为：

1. SDK 验证 `compression` 为 `undefined` 或 primitive boolean；
2. 执行现有 compatibility diagnostics 与 strict/permissive gate；
3. private byte writer 把 selector 传给 `OpcPackage.write({ compression })`；
4. OPC 根据 package state 与 selector 选择 original、STORE 或 DEFLATE；
5. `write()` 再把同一 canonical bytes 转换成目标 output type，其他 convenience methods 直接消费这些 bytes。

OPC 决策矩阵：

| Package state | selector | 结果 |
| --- | --- | --- |
| opened + unchanged | omitted / undefined | defensive copy of original bytes |
| opened + unchanged | false | JSZip STORE regeneration |
| opened + unchanged | true | JSZip DEFLATE level 6 regeneration |
| created or changed | omitted / undefined | JSZip STORE generation |
| created or changed | false | JSZip STORE generation |
| created or changed | true | JSZip DEFLATE level 6 generation |

STORE generation 不传无意义的 `compressionOptions`。DEFLATE 固定 level 6，与 JSZip/PptxGenJS 当前默认 level 和 native 既有 writer 一致。entry order、entry date、part bytes、content types、relationships、diagnostics 与 mutation journal 不因 selector 改变。

## 6. 一致性、错误与隔离

- 同一 document state 与同一 selector 下，六种 `write()` output、stream、file、presentation Blob 和 download Blob 的解码 bytes 必须逐 byte 相同；
- `compression: true` 和 false 的 package parts、relationships 与 presentation semantics 必须相同，差别只允许存在于 ZIP container；
- 具有足够重复内容的 fixture 中，DEFLATE 输出必须小于 STORE，并且 ZIP local entries 证明 compression method 分别为 8 与 0；
- true/false 输出都必须能由 `PptxDocument.open()` 重开；
- selector 不修改 package、mutation journal 或诊断内容；重复相同 write 必须 deterministic；
- invalid selector 在 diagnostics replacement 与 `OpcPackage.write()` 前抛出 `TypeError('PptxDocument output compression must be a boolean')`；
- `stream()` 的 browser-only rejection 仍先于 compression validation，保持其既有平台边界；
- direct `OpcPackage.write()` 对非法 selector 也抛出 `TypeError('Package compression must be a boolean')`，防止 SDK 以外调用绕过 contract；
- explicit `compression: undefined` 等价 omitted，不视为强制重新打包。

## 7. 模块边界

- `packages/opc/src/index.ts` 只负责 original/STORE/DEFLATE 决策和 JSZip generation；
- `packages/sdk/src/index.ts` 只负责公开 field、输入验证、diagnostics 顺序和向 OPC 传递 selector；
- `packages/sdk/src/write-output.ts` 不修改，确保 compression 与返回表示正交；
- aggregate `packages/pptx` 继续通过 SDK re-export 自动获得类型；
- model、codecs、validator、lossless XML 与 adapter 不承担 native ZIP policy；
- browser bundle 不增加 Node import、Buffer dependency 或环境专属 compression 实现。

## 8. 验证策略

### Source 与类型

- OPC focused tests 锁定 changed/default/false/true、opened-unchanged fast path、显式重新打包、method 0/8、size、determinism 与 reopen；
- SDK focused tests 锁定所有输出 surface 的 same-selector byte equality、六 output representations、strict/permissive parity、invalid early failure、diagnostic/journal isolation；
- TypeScript checks 接受 `compression?: boolean` 于 `WriteBaseOptions`、generic `WriteOptions` 和全部 convenience methods，并拒绝非 boolean；
- public-only PptxGenJS probe 记录 stream true/false 的实际 method/size，以及显式 output type 忽略 selector 的 deliberate divergence。

### Packed Node 与 browser

- actual tarball Node consumer 检查 STORE/DEFLATE methods、size reduction、六 output representations、stream、file、Blob、reopen、original fast path、declarations 与 CLI；
- browser conditional export 检查 false/true 的 method、size、`write`/`writeBlob` 一致性、reopen 与无 Node dependency；
- real Chrome 检查 `write`、`writeBlob`、download 捕获 bytes、invalid failure isolation、later write/reopen，以及 zero validation/console/page/network errors；
- full Vitest、performance、两套 TypeScript、Node/browser bundles、declaration build、actual tarball file count/SHA-256 全部记录。

## 9. 实施与提交边界

实施分为三个可独立 review 的 repository change：

1. OPC/SDK core policy、source/type/PptxGenJS conformance；
2. actual-tarball Node/browser/declaration/CLI 与 real-Chrome proof；
3. README/API/compatibility/progress/changelog 发布文档。

每个 change 都必须 review、commit、push、fetch，并确认远端同步后再继续。

本项完成标准：所有输出 surface 使用同一 boolean policy；STORE/DEFLATE container 证据明确；合法 PptxGenJS compression 能力对等；上游显式-output 缺陷不复制；无损编辑 fast path 保留；Node/browser release gates 全绿。完成后下一项进入 scheme-color runtime helper。
