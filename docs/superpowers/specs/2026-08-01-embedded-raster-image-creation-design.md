# Embedded Raster Image Creation Design

## 1. 目标

在现有 `ImageModel` 读取、transform 编辑和 clone-on-write `replaceData()` 基础上，增加
一个严格、同步、平台无关的 native embedded raster image 创建入口。首个 image/SVG
小项只接收调用方已持有的 `Uint8Array`，为 PNG、JPEG 和 GIF 创建完整 `p:pic`、image
relationship 与 `/ppt/media/*` part，并立即返回 live `ImageModel`。

本项必须从 zero-input presentation 工作，也必须能向 existing deck 的任意合法 slide
追加图片。创建结果需要支持现有 transform、duplicate、move、delete、rollback、write/reopen
和 `replaceData()` 生命周期，不引入 Node-only I/O，也不在运行时依赖 PptxGenJS。

用户已要求实现方连续决定后续内容，因此本设计完成 self-review 后直接进入实施计划，
不设置交互式确认点。

## 2. 当前状态与问题

`ImageModel` 已支持：

- 读取 embedded `sourcePartUri` 或 external `externalUrl`；
- 读取和编辑 common shape transform；
- 对 embedded payload 执行 atomic clone-on-write `replaceData()`；
- 在 slide duplicate/delete dependency lifecycle 中保留 shared image target。

缺少的是创建入口。调用方目前无法从 blank presentation 添加图片，只能打开已有 `p:pic`
再编辑。PptxGenJS 4.0.1 的 `slide.addImage()` 同时混合 path/data I/O、source-type 推断、
SVG fallback、contain/cover/crop、rounding、transparency、alt text、hyperlink、shadow 和
placeholder。一次实现全部语义会把可独立验证的 package mutation 与异步 I/O、图片解析、
layout 和 client compatibility 绑在一起。

## 3. 方案比较

### A. 一次实现完整 `addImage({ path | data, ... })`

表面最接近 PptxGenJS，但 local path、URL、Blob、data URI、SVG raster fallback、pixel-size
解析和 OOXML mutation 会落入一个大入口。浏览器/Node 分支、失败回滚和测试矩阵无法形成
清晰的小项 review gate。拒绝。

### B. 先只做 `PptxDocument.addImage(slideIndex, source, options)`

可以复用现有 media source loader，但把核心图片创建隐藏在 SDK facade，无法先建立路线图要求
的 `SlideModel.addImage()`，也会迫使 model 依赖异步 I/O。拒绝作为第一层。

### C. `SlideModel.addImage(bytes, options)` 作为同步核心（采用）

Model 层只拥有 strict bytes、direct transform、part、relationship 和 `p:pic`。后续异步
document-level path/URL/Blob/data-URI loader 只负责把 source 解析成 bytes/content type，再调用
这个核心。SVG 项可复用同一 raster fallback 创建路径，并额外管理 SVG part/extension
relationship。该方案保持平台边界和原子 transaction 最清晰。

## 4. 公开 API

```ts
export type RasterImageContentType =
  | 'image/png'
  | 'image/jpeg'
  | 'image/gif';

export interface AddImageOptions extends Partial<Transform> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
}

export class SlideModel {
  addImage(bytes: Uint8Array, options: AddImageOptions): ImageModel;
}
```

`contentType` 在本项为 required，避免在没有图片 signature/dimension parser 时根据文件名或
bytes 猜测。后续 source loader 和 content-type detection 可以在调用这个核心前填入同一个
字段，不需要改变 model 方法的职责。

`AddImageOptions` 使用现有 native direct units：x/y/width/height 是 EMU，rotation 是 OOXML
`1/60000°`，flip fields 是 boolean。可继续用 `inches()` / `degrees()` 转换。

默认值与 PptxGenJS 4.0.1 image output 对齐：

- x = 0；
- y = 0；
- width = 1 inch；
- height = 1 inch；
- rotation = 0；
- flipHorizontal = false；
- flipVertical = false；
- name = `Image ${zeroBasedImageCountBeforeInsert}`；
- omitted alt text 写为 `preencoded.png`，匹配其 data-input fallback；显式 `''` 保留为空。

`ImageModel` 的 existing `name`、`transform`、`sourcePartUri` 和 `replaceData()` 立即适用于返回
对象。本项不增加 payload getter 或 alt-text editor；后者作为独立 direct-state 小项实施。

## 5. Strict normalization

`bytes` 必须是直接 `Uint8Array`，长度至少为 1。创建前复制为 detached
`new Uint8Array(bytes)`；调用方后续修改原 buffer 不得改变 package payload。

本项只验证 explicit content type，不解析或验证 PNG/JPEG/GIF signature。Signature、pixel
dimensions、animated-image/client diagnostics 属于后续 detection 小项。Content type 与 part
extension 使用固定表：

| Content type | Extension |
| --- | --- |
| `image/png` | `.png` |
| `image/jpeg` | `.jpeg` |
| `image/gif` | `.gif` |

Options 必须是 ordinary 或 null-prototype object，只允许 own data
`contentType/name/altText/x/y/width/height/rotation/flipHorizontal/flipVertical`。Accessor、inherited、
symbol、unknown property 和 runtime-`undefined` required content type 在 mutation 前拒绝。

`name` / `altText` 必须是 string 且不含非法 XML 1.0 character；empty string 合法。Transform
沿用 preset/custom shape 的 strict contract：x/y 为 finite safe integer，width/height 为 positive
safe integer，rotation 为 safe integer 且落在 `-21600000..21600000`，flip 必须是 boolean。

Normalized definition 与 nested/copied bytes 在进入 transaction 前脱离 caller。结构只保存创建
所需字段，不保留 options object identity。

## 6. OOXML 与 package ownership

每次创建分配一个新的 `/ppt/media/imageN.<ext>` part，并按 explicit content type 写入 detached
bytes。本项不做 hash deduplication；这与 PptxGenJS data input 的实际输出一致，也避免在首项引入
异步 hashing。后续 source-loader dedup 可以在调用核心前选择已有 target，或扩展内部 helper，
不改变 public bytes API。

Slide 获得一条 internal image relationship：

```xml
<Relationship
  Id="rIdN"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  Target="../media/imageN.png"/>
```

Picture 使用 canonical direct tree：

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="ID" name="Image 0" descr="preencoded.png"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rIdN"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="914400" cy="914400"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

Picture 必须插入 direct `p:cSld/p:spTree`，若存在 direct `p:extLst` 则插在它之前。ID 使用
slide-wide safe allocator，与 shape/text/table/custom geometry 共用 collision contract。创建后通过
`slide.shapes` 解码回同一 live `ImageModel`；若解码不成 image，则整个 transaction 回滚。

Part、content-type entry、relationship、slide XML、shape cache identity 和 mutation journal 必须处于
同一个 outer transaction。任何错误都不能留下 orphan media part 或 relationship。

## 7. Lifecycle 与无损边界

- 创建只追加一个 `p:pic`，不重写现有 shape XML、unknown children、opaque parts 或 relationships。
- `ImageModel.setTransform()` 必须继续只编辑返回对象的 transform。
- `replaceData()` 对刚创建的 exclusive target 原位写入；duplicate 后的 shared target 继续使用已有
  clone-on-write 行为。
- Slide duplicate 保留 shared image target；对任一副本调用 `replaceData()` 时自动隔离。
- Slide move 不改变 image target；delete 只在现有 dependency policy 允许时回收 unreferenced target。
- Outer transaction rollback 恢复新 part、relationship、slide XML、shape ID allocation 和 journal。
- `.pptx/.pptm/.ppsx/.ppsm/.potx/.potm` 六种 package profile 的 create/write/reopen 语义一致。

## 8. PptxGenJS 4.0.1 对等边界

Conformance fixture 使用其 public `slide.addImage({ data, x, y, w, h, rotate, flipH, flipV,
objectName, altText })` 与 public `write()`，不读取内部对象。对 PNG/JPEG/GIF 各建立至少一个
合法 data URI，并通过 adapter 比较：

- picture kind/order；
- embedded bytes 与 content type；
- internal image relationship；
- x/y/width/height/rotation/flips；
- name 与 alt text；
- rect geometry、no-change-aspect lock 与 stretch fill。

Native API 使用 direct EMU/OOXML angle，而 PptxGenJS public input 使用 inch/degree；测试先将
输入配对到相同最终 values。PptxGenJS 的 falsy numeric fallback、invalid base64 passthrough、
path-derived extension 和 console-only rejection 不进入 native contract；native 在 mutation 前严格
拒绝 invalid input。

本项不宣称 path/URL/data-URI public input、automatic content-type/dimension detection、contain/cover/
crop、rounding、transparency、hyperlink、shadow、placeholder 或 SVG 对等。这些是后续独立小项，
但都必须复用本项建立的 picture/part/relationship core。

## 9. 测试与发布门禁

### Internal normalization/renderer

- PNG/JPEG/GIF extension mapping；
- default/explicit transform、name、alt text；
- null-prototype options；
- descriptor-safe/getter-free normalization；
- caller bytes/options detachment；
- empty/non-`Uint8Array` bytes、missing/unknown content type、unknown/accessor/inherited/symbol option、invalid XML、
  unsafe transform 的 zero-mutation rejection；
- canonical XML escaping 与 required picture child order。

### Model lifecycle

- blank slide create returns live `ImageModel` immediately；
- part bytes/content type、relationship target、sourcePartUri、name/transform 和 shape order；
- insertion before extLst and preservation of opaque state；
- multiple images receive unique IDs/parts/default names；
- setTransform、replaceData、duplicate/clone-on-write、move/delete；
- nested rollback and failure after resource allocation；
- six-format write/reopen and stable model identity。

### SDK/package/conformance

- public `@jiayunxie/pptx` export/types compile；
- PptxGenJS 4.0.1 public-output semantic comparison for PNG/JPEG/GIF；
- actual npm tarball Node/browser/types/CLI smoke verifies `SlideModel.addImage()`；
- PowerPoint 2010 strict validation for the generated raster gallery；
- LibreOffice headless open/save and reopen evidence, with any direct-state rewrite recorded explicitly。

最终 release gate：focused tests、`pnpm typecheck`、`pnpm test`、`pnpm test:performance`、`pnpm build`、
dist clean check、actual tarball smoke、`git diff --check`。

## 10. 后续小项顺序

1. Embedded raster bytes creation（本项）；
2. raster source loader：path/URL/data URI/Blob/ArrayBuffer/stream；
3. content-type 与 intrinsic pixel-dimension detection；
4. contain/cover/crop sizing 与 direct srcRect editing；
5. SVG part + PNG fallback + `asvg:svgBlip` lifecycle；
6. rounding 与 transparency；
7. image alt-text/name direct editing；
8. image hyperlink、shadow 与 placeholder；
9. image delete/relationship GC 专用 public entry 与 full peer-range audit。

每个小项单独经过 tests、review、commit 和 push。

## 11. Self-review

- 无占位标记或未定义 public type。
- Public model method 保持同步、browser-safe；异步 I/O 有明确后续 facade boundary。
- Required content type 与“不做 signature detection”不冲突。
- Defaults、direct units、part ownership、relationship、rollback 和 PptxGenJS semantic comparison 均有
  单一解释。
- 首项可独立发布，不虚构 path/SVG/sizing/style 对等；后续扩展不要求破坏当前 method signature。
