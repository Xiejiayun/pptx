# Presentation `write({ outputType })` Return Semantics Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项让现有 `OUTPUT_TYPES` 六值 runtime catalog 成为 `PptxDocument.write()` 的真实选择器，并为每个 literal token 提供可推导的返回类型：

| `outputType` | Node runtime | Browser runtime | 公共 TypeScript 返回类型 |
| --- | --- | --- | --- |
| `arraybuffer` | standalone `ArrayBuffer` | standalone `ArrayBuffer` | `ArrayBuffer` |
| `base64` | raw base64 string | raw base64 string | `string` |
| `binarystring` | byte-per-code-unit binary string | byte-per-code-unit binary string | `string` |
| `blob` | `Blob`，MIME `application/zip` | `Blob`，MIME `application/zip` | `Blob` |
| `nodebuffer` | Node `Buffer` | reject | `Uint8Array` |
| `uint8array` | plain `Uint8Array` | plain `Uint8Array` | `Uint8Array` |

`nodebuffer` 的 runtime 必须通过 `Buffer.isBuffer()`；公共 declaration 使用它的标准超类 `Uint8Array`，避免 browser consumer 为一个不可用分支安装 Node type declarations。真实 Node 行为、packed Node 行为和 browser rejection 分别由 runtime tests 锁定。

本项不实现 `STREAM`、`stream()`、compression selector、write-file compression、下载策略变化或新的文件格式。它只完成六种非流式返回语义。

## 2. 当前状态与权威基线

Native 当前边界：

- `OpcPackage.write()` 始终返回 canonical `Uint8Array`；未改动 package 会返回原始 package 的 defensive copy，已改动 package 通过 JSZip 生成 DEFLATE level 6 bytes；
- `PptxDocument.write(options?)` 执行 validation/diagnostics 后返回底层 bytes；
- `writeBlob()` 使用 presentation format 的真实 MIME；
- `writeFile()` 把 bytes 写到 Node path；
- `download()` 通过 `writeBlob()` 触发 browser download；
- `WriteOptions` 目前只有 compatibility 和 strict/permissive mode。

锁定依赖 PptxGenJS 4.0.1 的 public constructor、`addSlide()`、`addText()` 和 `write()` 实测结果：

- 六个显式 token 的 runtime 种类依次为 `ArrayBuffer`、string、string、`Blob`、`Buffer`、`Uint8Array`；
- base64 不带 data-URI prefix；binary string 的每个 code unit 为一个 byte；
- explicit blob MIME 为 `application/zip`；
- real Chrome 对 `nodebuffer` 抛出 `Error: nodebuffer is not supported by this platform`；
- PptxGenJS `write()` 无参数默认返回 Blob，但 4.0.1 的显式空对象会因缺少 `outputType` 抛出 `TypeError`。

Native 已公开并广泛使用无参数 `write(): Promise<Uint8Array>`，且仓库内 create/edit/reopen、adapter、CLI 和用户代码都依赖它。本项对等六个显式 token，但保留 native default：omitted options、`{}`、`outputType: undefined` 以及只有 validation fields 的 options 均返回 `Uint8Array`。不复制 PptxGenJS 空对象缺陷，也不进行破坏性的默认值迁移。

## 3. 方案比较

### A. 让 `OpcPackage.write(outputType)` 直接调用 JSZip 六种 generator type

这能减少 SDK 转换代码，但会把 Blob、Buffer、base64 和 presentation-specific public contract 泄漏到通用 OPC package 层；未修改 package 的 original-byte fast path 也必须为六种类型重复实现。Browser bundle 还会承担 Node type/runtime 分支。拒绝。

### B. SDK 从单一 canonical bytes 做环境安全转换（采用）

`PptxDocument` 先验证 token，再执行现有 validation 和一次 `OpcPackage.write()`，最后转换 bytes。OPC 层、ZIP policy、diagnostics、package mutation 和 deterministic bytes 都保持原样。所有类型都能从同一 bytes 做交叉相等验证，`nodebuffer` 单独使用 runtime-only dynamic Node import。

### C. 继续扩展 `writeBlob()` 并新增 `writeArrayBuffer()` / `writeBase64()` 等方法

独立方法容易理解，但不消费 `OUTPUT_TYPES`，无法提供 PptxGenJS `write({ outputType })` 对等调用形状，并扩大六套 convenience API。现有 convenience methods 保留，但统一 selector 必须由 `write()` 提供。拒绝。

## 4. 公共 API 与类型推导

`packages/sdk/src/output-type.ts` 增加公开 conditional type：

```ts
export type WriteOutput<TOutputType extends OutputType = OutputType> =
  TOutputType extends 'arraybuffer' ? ArrayBuffer
    : TOutputType extends 'base64' | 'binarystring' ? string
      : TOutputType extends 'blob' ? Blob
        : Uint8Array;
```

`packages/sdk/src/index.ts` 把原有 validation fields 提取为 base options，并让 `WriteOptions` 以 `uint8array` 为 source-compatible default：

```ts
export interface WriteBaseOptions {
  readonly compatibility?: CompatibilityProfile;
  readonly mode?: 'strict' | 'permissive';
}

export interface WriteOptions<
  TOutputType extends OutputType = 'uint8array',
> extends WriteBaseOptions {
  readonly outputType?: TOutputType;
}

async write<TOutputType extends OutputType = 'uint8array'>(
  options?: WriteOptions<TOutputType>,
): Promise<WriteOutput<TOutputType>>;
```

精确推导：

```ts
await document.write();                              // Uint8Array
await document.write({ mode: 'permissive' });        // Uint8Array
await document.write({ outputType: 'arraybuffer' }); // ArrayBuffer
await document.write({ outputType: 'base64' });      // string
await document.write({ outputType: 'blob' });        // Blob
await document.write({ outputType: 'nodebuffer' });  // Uint8Array type; Buffer in Node
```

`writeBlob()`、`writeFile()` 和 `download()` 接受 `WriteBaseOptions`，从 private byte writer 读取 canonical bytes，不受 selector 影响。已有 object literals 和 `WriteOptions` variables 如果只含原 validation fields，继续结构兼容。

`OutputType`、`WriteOutput`、`WriteBaseOptions` 和 generic `WriteOptions` 由 SDK 与 aggregate root 导出。无 instance enum/getter、mutable alias 或 `STREAM` union。

## 5. Runtime 转换与错误语义

调用顺序：

1. 从 options 读取 `outputType`，omitted/undefined 规范化为 `uint8array`；
2. 严格检查它是 `OUTPUT_TYPES` 的精确成员；未知值、大小写变体、`STREAM`、null、number 和 object 在 validation/ZIP generation 前抛出 `TypeError`；
3. browser 遇到 `nodebuffer` 时，在 validation/ZIP generation 前抛出与 PptxGenJS 相同的 `Error('nodebuffer is not supported by this platform')`；
4. 执行一次现有 validation、diagnostic replacement 和 `OpcPackage.write()`；
5. 进行目标转换。

转换 contract：

- `uint8array` 直接返回 canonical plain bytes；不能是 Buffer；
- `arraybuffer` 返回只覆盖 presentation bytes 的 standalone copy，byte offset 为 0，不暴露更大 backing buffer；
- `binarystring` 对每个 byte 产生一个 0–255 code unit，分块构造以避免参数栈溢出；
- `base64` 使用 environment-neutral byte encoder，不依赖 Buffer、`btoa` 或 DOM；输出无换行、无 prefix、标准 `+/=` alphabet；
- `blob` 从 standalone bytes 创建，MIME 固定为 PptxGenJS/JSZip 对等的 `application/zip`；
- `nodebuffer` 只在 Node runtime 通过非静态 `node:buffer` import 创建 defensive Buffer copy；browser bundle 不出现 static Node import。

所有成功类型解码后的 bytes 必须逐 byte 相同，并能由 `PptxDocument.open()` 重开。转换不修改 OPC package、mutation journal、format、diagnostics 内容或后续 write 结果。

显式 blob selector 与 `writeBlob()` 的 MIME 有意不同：前者对等 JSZip output type 的 `application/zip`，后者继续提供 presentation-specific MIME，例如 PPTX 的 `application/vnd.openxmlformats-officedocument.presentationml.presentation`。

## 6. 隔离、兼容性与非目标

- 不修改 `OpcPackage.write()`、ZIP compression、entry order/date、original-byte fast path 或 format profile；
- 不修改 `PptxInput`；所有六种成功输出解码后继续能作为 open input；
- 不让 output conversion 进入 model/codecs/validator；
- 不把 Node Buffer constructor 放到 browser module initialization；
- 不接受 `STREAM`，独立 `stream()` 小项仍负责 backpressure、error propagation 和 Node readable contract；
- 不增加 `compression`，下一小项会决定 boolean/default、unchanged-package fast path 和 PptxGenJS byte conformance；
- 不改变 `writeFile()` 的 void return、`download()` 行为或 `writeBlob()` format MIME；
- 不复制 PptxGenJS `write({})` TypeError、宽泛 union return declaration 或 browser-only default。

## 7. 验证策略

### Source 与 types

- focused conversion tests 覆盖六值 constructor、MIME、Buffer identity、byte equality、open/reopen 和 package mutation isolation；
- base64/binary string 覆盖 bytes `0x00`、`0x7f`、`0x80`、`0xff`，不是只验证 ASCII；
- default、empty object、mode-only、compatibility-only 和 explicit undefined 均保持 plain `Uint8Array`；
- generic type tests 锁定六个 literal 返回类型、dynamic union 返回 union、readonly options 和 invalid tokens；
- invalid runtime values 与 browser `nodebuffer` 在 package write 前失败；
- public-only PptxGenJS conformance 锁定六值 runtime kind、blob MIME 和可重开输出，不读取私有字段或 JSZip internals。

### Packed Node/browser

- actual tarball declarations 包含 `WriteOutput` conditional、generic `WriteOptions` 和 generic `write()`；
- installed Node 检查 default、六值、`Buffer.isBuffer()`、逐 byte equality、reopen、CLI 和现有 convenience methods；
- installed TypeScript consumer 用 `satisfies` 锁定每个 literal Promise 类型，并拒绝 `STREAM`/unknown；
- browser conditional export 检查无 static `node:` import；
- real Google Chrome 检查五个可用 token、`nodebuffer` exact rejection、base64/binary high bytes、Blob MIME、reopen、zero validation/console/page/network errors；
- actual tarball file count、SHA-256、full Vitest、performance、两套 TypeScript、Node/browser bundles 与 declarations 全部记录。

## 8. 实施边界与完成门禁

实施分成三个可独立 review 的 repository change：

1. core API、转换、source/type/PptxGenJS conformance；
2. actual-tarball Node/browser/types/CLI 与 real-Chrome proof；
3. README/API/compatibility/progress/changelog 发布文档。

每个 change 都必须 review、commit、push、fetch，并确认远端同步后再进入下一项。

本小项完成标准：六个 catalog token 都有真实返回语义；默认和四个 convenience contracts 无回归；Node/browser 边界明确；所有成功输出解码后 byte-identical 且可重开；release gates 全绿。完成后总体进度仍按约 97% 报告，下一项进入 Node readable stream，再进入 compression policy。
