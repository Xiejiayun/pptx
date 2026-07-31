# Raster Image Source Loading and Inspection Design

## 1. 目标

在已经完成的 strict、同步、平台无关 `SlideModel.addImage(bytes, options)` 之上，增加
PptxGenJS 风格的高层图片来源入口。调用方可以从本地路径、HTTP/HTTPS URL、浏览器相对 URL、
strict data URI、`Uint8Array`、`ArrayBuffer`、`Blob`/`File`、Web `ReadableStream` 或 async iterable
添加 PNG、JPEG、GIF 图片。

本项同时建立唯一的 raster signature 与 intrinsic pixel-dimension inspector。图片格式始终由 bytes
决定；显式 content type 只作为必须与检测结果一致的 assertion。所有 I/O、解码和检查必须在 package
mutation 之前完成，因此失败不会留下 part、relationship 或 picture XML。

本项不改变现有 1-inch 默认 transform。检测出的 pixel width/height 通过公共 inspector 暴露，并供
下一小项 contain/cover/crop sizing 使用。

## 2. 当前状态与问题

当前 native core 已经能够：

- 从 detached `Uint8Array` 创建完整 embedded PNG/JPEG/GIF picture；
- 创建 media part、content type、image relationship 与 canonical `p:pic`；
- 原子回滚，并参与 transform、replace、duplicate、move、delete 和 clone-on-write 生命周期；
- 从 blank presentation 和 existing deck 工作。

但调用方仍必须自己读取 bytes 并传入 `contentType`。这与 PptxGenJS 的常用 `path`/`data` 输入存在明显
差距，也没有一个可复用、严格的 intrinsic-size parser。若 sizing 或 SVG fallback 各自实现格式推断，
会产生多个不一致的 truth source。

## 3. 方案比较

### A. 扩展 `SlideModel.addImage()`，让它直接接收所有异步来源

该方案表面 API 最少，但 model mutation 会变成 async，并把 Node 文件系统、Fetch、Blob、stream 和
data URI 解析引入 OOXML model。现有同步 transaction 语义也会被破坏。拒绝。

### B. 复用通用 media loader，并继续依据扩展名或 MIME 决定格式

现有 audio/video loader 的字符串 URL 表示 external media relationship，且 MIME/扩展名推断允许默认值；
图片必须真正下载并 embed，格式错误也不能伪装成合法 part。两者的 source policy 与验证强度不同，直接
复用会保留错误语义。拒绝。

### C. SDK source loader + 独立 byte inspector + existing model core（采用）

新增 browser-safe raster source/inspection 模块。SDK 先解析来源、收集 detached bytes、以 signature
检测格式和 pixel dimensions，再调用 existing `SlideModel.addImage()`。Model 层保持同步且只负责 package
mutation；loader 和 inspector 可独立测试；下一阶段 sizing 与 SVG fallback 复用同一个 inspector。

## 4. 公共 API

在 SDK 根入口导出：

```ts
export type RasterImageByteChunk =
  | number
  | Uint8Array
  | ArrayBuffer
  | ArrayBufferView;

export type RasterImageByteStream =
  | ReadableStream<RasterImageByteChunk>
  | AsyncIterable<RasterImageByteChunk>;

export type RasterImageSource =
  | string
  | Uint8Array
  | ArrayBuffer
  | Blob
  | RasterImageByteStream;

export interface RasterImageInfo {
  readonly contentType: RasterImageContentType;
  readonly width: number;
  readonly height: number;
}

export interface AddImageSourceOptions
  extends Omit<AddImageOptions, 'contentType'> {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
}

export function inspectRasterImage(bytes: Uint8Array): RasterImageInfo;

export class PptxDocument {
  addImage(
    slideIndex: number,
    source: RasterImageSource,
    options?: AddImageSourceOptions,
  ): Promise<ImageModel>;
}
```

`slideIndex` 保持 SDK 现有 zero-based convention，并在执行任何 source I/O 前验证。`options` 默认 `{}`。
除 `contentType` 和 `signal` 外，其余字段与 existing `AddImageOptions` 完全一致。SDK 只把 inspector 检测出的
content type 与 image creation fields 传给 model core，绝不把 `signal` 传入 strict model normalizer。

`inspectRasterImage()` 只接受直接 `Uint8Array`，不读取路径或 URL。它不保留 input identity，也不修改
bytes。返回的 width/height 是正安全整数 pixel count。

## 5. Source resolution

### 5.1 字符串分派

字符串按以下固定顺序解释：

1. `data:`：只接受本设计定义的 strict base64 image data URI；
2. `http://` 或 `https://`：Node 和 browser 都通过 Fetch 下载并 embed；
3. 其他字符串：Node 解释为 local path，browser 交给 Fetch，允许相对 URL；
4. 其他显式 URI scheme 一律拒绝，不回退为文件路径。

空字符串在读取前拒绝。Node local path 使用动态 `node:fs/promises` import，browser bundle 不产生静态
Node import。Fetch 必须检查 `response.ok`；redirect 使用平台默认行为；网络、HTTP 和 abort 错误保持为
明确失败，不生成 external image relationship。

### 5.2 Strict data URI

只接受以下 MIME 的 `;base64,` form：

- `image/png`；
- `image/jpeg`；
- `image/gif`。

Media type 与 `base64` token 大小写不敏感，但不接受额外 parameter、percent-encoded raw payload、
URL-safe alphabet、空 payload、非 canonical padding、空白或 trailing data。解码使用 browser-safe API/逻辑，
不依赖 `Buffer`。Data URI 声明的 MIME 必须与 signature inspector 一致。

### 5.3 Byte-like inputs

- `Uint8Array`：立即复制；
- `ArrayBuffer`：复制全部 bytes；
- `Blob`/`File`：await `arrayBuffer()` 后复制；
- Web stream / async iterable：逐 chunk 检查并收集，结束后合并成单一 detached `Uint8Array`。

Stream chunk 接受 byte number、`Uint8Array`、`ArrayBuffer` 或任意 `ArrayBufferView`。Byte number 必须是
0..255 整数。非法 chunk、non-stream object、reader error 或 abort 均失败。Web reader 总是在 finally
release lock；async iterable 在 abort 时通过 iterator closing 结束。

本项不引入 arbitrary size limit，也不承诺 streaming ZIP/image decode；source resolution 的输出是一个
完整、detached byte array，与现有 image core contract 一致。

### 5.4 MIME assertion policy

Signature 是 content type 的唯一事实来源。以下值是 assertion，若不一致则拒绝：

- `options.contentType`；
- data URI media type。

`Blob.type` 和 HTTP `Content-Type` 可能由平台、服务器或代理错误设置，因此不用于接受未知格式，也不作为
hard assertion。文件扩展名、URL suffix、query string 和 `File.name` 不参与格式判断。

## 6. Signature 与 dimensions

### 6.1 PNG

- 必须包含完整 8-byte PNG signature；
- 第一个 chunk 必须是 length `13` 的 `IHDR`；
- 至少包含可读取 width/height 的 24 bytes；
- width/height 从 IHDR 以 unsigned big-endian 32-bit 读取；
- 任一 dimension 为 0 或超过 JavaScript safe integer contract 时拒绝。

本项只检查识别和 dimensions 所需 header，不做 CRC 或完整 IDAT/IEND 验证。

### 6.2 GIF

- signature 必须是 `GIF87a` 或 `GIF89a`；
- 至少包含 10-byte logical screen descriptor prefix；
- width/height 从 offset 6/8 以 unsigned little-endian 16-bit 读取；
- 任一 dimension 为 0 时拒绝。

Dimensions 表示 logical screen，不解析 animation frames。

### 6.3 JPEG

- 必须以 SOI `FF D8` 开始；
- 按 marker segment 长度安全遍历，允许合法 fill bytes 和 standalone marker；
- 接受所有携带 sample dimensions 的 SOF marker：C0-C3、C5-C7、C9-CB、CD-CF；
- SOF segment 必须足够长，height/width 以 unsigned big-endian 16-bit 读取且非 0；
- truncated segment、非法 length、在 SOF 前遇到 SOS/EOI 或始终无 SOF 时拒绝。

APP metadata、EXIF orientation 与 ICC profile 不改变 raw pixel dimensions。本项不自动旋转 transform。

未知 signature、truncated header 和 zero dimension 都抛出 deterministic `TypeError`。错误消息包含问题类别，
不包含完整 source data 或 data URI payload。

## 7. Mutation 与 detachment

`PptxDocument.addImage()` 顺序固定为：

1. 验证 `slideIndex` 并取得 target slide；
2. descriptor-safe normalize `AddImageSourceOptions`，拒绝 unknown、inherited、accessor 与 symbol property；
3. resolve source，响应 `AbortSignal`；
4. inspect signature/dimensions；
5. 检查显式 MIME assertions；
6. 调用 `slide.addImage(detachedBytes, normalizedModelOptions)`。

前五步不修改 package。第六步继续使用 existing model transaction，因此 allocation 或 XML 失败会整体回滚。
SDK source options 的 transform/name/altText validation 与 model core 保持同一 contract；SDK normalization 只负责
安全拆分 loader-only fields，最终 semantic validation 仍由 core 完成。

调用完成后再修改 caller `Uint8Array`、`ArrayBuffer`、Blob backing data 或已 yield 的 stream chunk，不得改变
package payload。Response/Blob/stream 也不保留在 document model 中。

## 8. PptxGenJS 4.0.1 对等边界

Conformance coverage 使用 public `slide.addImage()` 验证其 `path` 与 `data` workflow，并比较 native output：

- local PNG/JPEG/GIF path；
- HTTP/HTTPS URL 下载后 embed；
- PNG/JPEG/GIF data URI；
- default 与 explicit x/y/w/h、rotate、flip、name、alt text；
- embedded payload、content type、relationship 和 picture order。

PptxGenJS browser XHR 与 native Fetch 的 transport 实现不要求逐行相同，验收点是相同来源最终形成同语义
embedded picture。Native 额外支持 Blob、ArrayBuffer 和 streams，并对 signature/MIME mismatch、truncation
与非法 data URI 提供 strict zero-mutation rejection。

PptxGenJS 当前不会依据 natural dimensions 自动设置 w/h；native 本项同样保留 1-inch default。Intrinsic
dimensions 是下一小项 sizing API 的基础，不在本项隐式改变布局。

## 9. 测试与发布门禁

### Inspector

- PNG/JPEG/GIF 全部 signature 与 dimensions；
- GIF87a/GIF89a、JPEG 所有合法 SOF family、marker fill/standalone/APP traversal；
- unknown、empty、truncated、zero dimension、bad JPEG length、SOS/EOI-before-SOF；
- caller bytes unchanged，重复调用结果稳定。

### Source resolver

- copied Uint8Array/ArrayBuffer；
- Blob/File including empty/incorrect MIME metadata；
- Web stream 与 async iterable 的所有合法 chunk types、ordering、detachment 和 reader release；
- Node local path；Node/browser HTTP URL；browser relative URL；redirect；HTTP failure；network failure；abort；
- strict data URI valid cases 与 MIME mismatch、bad alphabet/padding/whitespace/parameter/raw encoding rejection；
- unsupported object/scheme 和 empty source。

### SDK/model integration

- blank/existing deck addImage returns live `ImageModel`；
- default/explicit transform、name、alt text 与 detected content type；
- invalid slide fails before source consumption；
- every read/parse/assertion/options failure leaves package bytes、parts、relationships、slide XML and journal unchanged；
- caller mutation after resolve cannot affect written payload；
- six presentation formats write/reopen；Node and browser public types compile。

### Conformance/release

- PptxGenJS 4.0.1 public path/data output comparison；
- actual npm tarball Node/browser/types smoke；
- generated multi-source gallery strict validation；
- LibreOffice headless open/save/reopen with payload and direct-state preservation evidence；
- focused tests、`pnpm typecheck`、`pnpm test`、`pnpm test:performance`、`pnpm build`、dist reproducibility、
  `git diff --check`。

## 10. 实施小项

1. Pure raster signature/dimension inspector；
2. In-memory、Blob 与 stream source resolution；
3. Strict data URI resolution；
4. Node path 与 Fetch URL resolution，包括 abort；
5. `PptxDocument.addImage()` atomic integration；
6. PptxGenJS conformance、tarball/browser smoke、gallery/client evidence 与文档。

每项完成 focused tests、self-review、commit 和 push 后再进入下一项。

## 11. Self-review

- 三种架构已有明确取舍，选定方案不改变 model 的同步 transaction contract。
- String dispatch 在 Node/browser 中无歧义；HTTP image 总是下载并 embed。
- Signature、asserted MIME 与 metadata hint 的优先级只有一种解释。
- PNG/GIF/JPEG dimensions 的 byte offsets、endianness、truncation 和 zero-size policy 完整。
- I/O/inspection 全部发生在 mutation 前，core transaction 继续负责 package atomicity。
- 本项不虚构 natural-size layout，对下一阶段 sizing 提供稳定公共 inspector。
- Public API 覆盖 PptxGenJS path/data，并额外覆盖现代 browser/stream inputs。
