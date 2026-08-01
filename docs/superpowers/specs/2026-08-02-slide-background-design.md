# Slide Background Design

## 1. 目标

完成一个可独立验收的 slide background 小项，使库能够从零创建、读取、编辑和清除 PowerPoint 页面背景，
并对齐 PptxGenJS 4.0.1 的公开背景能力：

- 无填充背景；
- sRGB 或 theme scheme 纯色背景与 `0..100` transparency；
- 现有 linear/path gradient；
- PNG、JPEG、GIF 内嵌图片背景；
- SDK 从 path、HTTP/HTTPS、browser-relative URL、strict data URI、`Uint8Array`、`ArrayBuffer`、
  `Blob`/`File`、Web stream 和 async iterable 设置图片背景；
- duplicated slide 的图片背景首次编辑隔离，替换/清除/删除后的 relationship 与 media part 正确回收；
- `pptx/pptm/potx/potm/ppsx/ppsm` 六种格式、Node、browser、实际 npm tarball、declarations、CLI、
  PowerPoint/LibreOffice 和 PptxGenJS imported-output 全链路验证。

本项只处理 slide 自身的 direct background。layout/master background 编辑、pattern fill、group fill、theme
`p:bgRef` 的语义编辑、图片 crop/tile/effects 和 default color 属于其他小项。读取遇到这些 opaque 状态时不猜测、
不修复；显式赋新值或清除仍按调用方意图替换整个 direct background。

用户已要求实现方连续决定后续技术取舍，因此本设计完成 self-review 后直接进入实施计划与逐项实现，不设置
交互式停顿点。

## 2. 当前状态与权威行为审计

当前 `SlideModel.background` 只暴露 `GradientFill | undefined`，并为每次读写创建一个 `GradientCodec`。
`GradientCodec.getSlideBackground()` 使用宽松 descendant 搜索，只能读 gradient；setter 能写 `p:bg/p:bgPr`，
但不能清除、不能处理 solid/noFill/image，也不管理图片 relationship 或 media part 生命周期。

现有基础设施已经覆盖本项的大部分底层能力：

- `simple-fill.internal.ts` 提供严格 `none | solid` normalization、read/render/equality；
- `GradientCodec` 已提供 gradient type、validation/encoding、OOXML color 与 compatibility diagnostics；
- `RasterImageContentType` 和 model image creation 已固定 PNG/JPEG/GIF MIME 与扩展名；
- SDK raster source loader 已提供跨 Node/browser source loading、signature inspection、abort、detached bytes 和
  asserted MIME 校验；
- OPC transaction、graph incoming、relationship allocation/update/remove 和 part allocation 可用于原子写入；
- slide duplication 将 image target 视为 shared，适合在首次 mutation 时 clone-on-write；
- image replacement 已验证 relationship id 被同页多个 XML node 引用时必须新建 relationship 的规则；
- slide deletion 已有 media target 收集和 graph-based GC，但当前只扫描 `p:pic`，不会收集 background image。

对锁定版本 PptxGenJS 4.0.1 的实际输出审计结论为：

- `{ color: 'FF3399' }` 写 `<a:solidFill><a:srgbClr val="FF3399"/></a:solidFill>`；
- `{ color: 'FF3399', transparency: 50 }` 在 color 下写 `<a:alpha val="50000"/>`；
- `{ data: PNG_DATA_URI }` 写 internal image relationship、`/ppt/media/*.png` 和
  `a:blipFill/a:blip + a:stretch/a:fillRect`；
- deprecated `{ fill: '00FF00' }` 等价于 solid color；
- `{ type: 'none' }` 不生成 direct background；
- `{ color: 'FF3399', type: 'none' }` 会留下空 `p:bgPr`。这是非法/不完整 choice，不复制；native
  `kind: 'none'` 写合法 `a:noFill`，`undefined` 才表示删除 direct background 并恢复继承。

PptxGenJS adapter 当前通过 public `write()` bytes 导入整包，不需要为背景增加 production translation；本项会在
adapter test 中用实际 PptxGenJS 生成 solid/transparency/image fixtures，验证 native reader 和写回结果。

## 3. 方案比较

### A. 扩展 `GradientCodec` 为通用背景 codec

改动文件少，但会让名为 gradient 的 codec 同时拥有 simple fill、image part、relationship、GC 和 SDK source
semantics。codec ownership/capabilities 与真实职责不再一致，后续 shape gradient 改动也会被背景生命周期牵连。

### B. Model-owned strict slide-background 模块，SDK 复用 source loader（采用）

新增聚焦的内部模块，负责 direct background 的严格 projection、normalization、OOXML patch、relationship isolation
和 media GC；gradient encode/decode/diagnostics 继续复用 `GradientCodec`，simple fill 继续复用现有 helper。
`SlideModel` 暴露同步 bytes-level property，SDK 只增加异步图片 source convenience API。职责边界与现有 raster
image 架构一致，也能单独测试生命周期。

### C. 长期存活的 `SlideBackgroundModel`

可把 background 变成有 identity 的对象并提供 `replaceData()`，但背景是 slide 的单值属性，不是 shape collection
成员；stable handle 不解决额外问题，反而增加 stale/clear/recreate identity 规则。当前 union property 更直接。

## 4. 公共 API

`@pptx/model` 根入口新增并导出：

```ts
export interface SlideBackgroundImage {
  readonly kind: 'image';
  readonly contentType: RasterImageContentType;
  readonly bytes: Uint8Array;
}

export type SlideBackground =
  | SimpleFill
  | GradientFill
  | SlideBackgroundImage;

export class SlideModel {
  get background(): SlideBackground | undefined;
  set background(value: SlideBackground | undefined);
}
```

`SimpleFill` 从内部类型提升为 model public type，但 normalization/render helper 仍保持 internal。纯色继续使用统一
`RichTextColor`：`{ kind: 'srgb', value: 'FF3399' }` 或 `{ kind: 'scheme', value: 'accent1' }`。
`transparency` 是百分比 `0..100`，写入时按 DrawingML alpha `100000..0` 映射。

Image getter 每次读取当前 part bytes，返回新 `Uint8Array`；返回对象及 color/gradient stop/rectangle 等嵌套结构均
detached，调用方修改它们不会改变 package。Setter 在 validation 开始时复制 image bytes，不持有调用方 buffer。

`undefined` 具有明确的 edit 语义：删除 direct `p:cSld/p:bg`，使页面重新使用 layout/master inheritance。
`{ kind: 'none' }` 则保留 direct background 并写 `<a:noFill/>`。重复赋完全相同的 supported value 是 exact no-op，
不改变 bytes、relationship id、target URI、mutation journal 或对象状态。

`@pptx/sdk` 根入口新增并导出：

```ts
export interface SetSlideBackgroundImageOptions {
  readonly contentType?: RasterImageContentType;
  readonly signal?: AbortSignal;
}

export class PptxDocument {
  setSlideBackgroundImage(
    slideIndex: number,
    source: RasterImageSource,
    options?: SetSlideBackgroundImageOptions,
  ): Promise<void>;
}
```

该入口只负责 source resolution 与 signature/MIME 校验，最终调用同一个 `slide.background = image` 同步核心。
普通 none/solid/gradient/clear 直接使用 `document.slides[index].background`，不增加重复 facade。

## 5. Validation 与 normalization

`normalizeSlideBackground(value)` 接受 `undefined` 或 descriptor-safe ordinary/null-prototype own-data object；拒绝
array、class instance、accessor、symbol key 和未知字段，且在 package transaction 前完成全部 validation。

- `none` 只能有 `kind`；
- `solid` 必须有合法 `RichTextColor`，可选 transparency 必须 finite 且在 `0..100`；
- linear/path gradient 复用并补齐现有 gradient validation：至少两个 stops，offset/alpha/angle/rectangle 均 finite，
  enum 合法，输入数组与嵌套对象全部 detached；
- image 必须有非空 `Uint8Array` bytes 和 PNG/JPEG/GIF content type；model 不重复实现 signature parser，SDK
  source API 会在进入 model 前执行 signature inspection 与 asserted MIME 一致性校验；
- 任何错误必须是 zero-mutation，outer transaction rollback 后 getter 重新读取恢复状态。

SDK options 同样是 closed ordinary own-data object，只接受 `contentType` 和 `signal`。`slideIndex` 不存在时先抛
`RangeError`，不读取 source。source load/abort/signature/MIME 失败发生在 model mutation 前。

## 6. Strict direct background reader

新增 `slide-background.internal.ts`。Reader 只沿 namespace-aware direct chain 识别：

```text
p:sld / p:cSld / p:bg / p:bgPr / one DrawingML fill choice
```

它不使用全局 `elements('bg')` 或 descendant fill 搜索，避免把 layout extension、effect child 或其他 opaque node
误当成 slide background。`p:bgRef`、缺失/重复 choice、错误 namespace、dangling relationship、external image、
非 image relationship、缺失 target、unsupported MIME、pattern/group fill 或 ambiguous structure 均投影为
`undefined`，读取时不修复源文件。

Supported choice 规则：

- direct `a:noFill` 和 `a:solidFill` 使用 `readSimpleFillChoice()` 的严格结构；
- direct `a:gradFill` 使用 `GradientCodec.decode()`，但候选必须是唯一 direct fill choice；
- direct `a:blipFill` 必须有一个可解析的 direct `a:blip@r:embed`，对应当前 slide 的 internal standard image
  relationship，target part 存在且 MIME 为 PNG/JPEG/GIF；其他 crop/stretch/tile/effect siblings 被 losslessly 保留，
  但 public state 只投影 payload bytes/content type。

Getter 返回 `undefined` 既可能表示没有 direct background，也可能表示 direct background 不在本项可安全投影的
子集。Setter 的含义不依赖 getter：显式新值替换 direct background choice，显式 `undefined` 清除 direct `p:bg`。

## 7. Canonical OOXML 写入

所有写入在一个 OPC transaction 中执行。新增或替换 supported background 时确保 direct 结构为：

```xml
<p:bg>
  <p:bgPr>
    <!-- exactly one a:noFill | a:solidFill | a:gradFill | a:blipFill -->
    <a:effectLst/>
  </p:bgPr>
</p:bg>
```

新建 `p:bg` 时插入 direct `p:cSld` 的第一个 child 位置，保持 schema order 在 `p:spTree` 之前。已有合法
`p:bgPr` 时只替换 direct fill choice，保留 `p:bgPr` attributes、unknown non-fill children、effect list 和周边 XML
字节；已有 `p:bgRef` 或 ambiguous background 时按显式赋值替换整个 direct `p:bg` 为 canonical `p:bgPr`。

none 使用 `<a:noFill/>`。solid 使用 `<a:solidFill>` 和 direct sRGB/scheme color；只有 transparency 显式提供时写
`a:alpha`，归一化后相同值不重写。gradient 继续使用现有 canonical encoder和 diagnostics。

Image canonical choice 为：

```xml
<a:blipFill>
  <a:blip r:embed="rIdN"/>
  <a:stretch><a:fillRect/></a:stretch>
</a:blipFill>
```

target 位于 `/ppt/media/backgroundN.png|jpeg|gif`，content type 与 extension 精确匹配。关系类型固定为 standard
Office image relationship，target mode 为 Internal。PptxGenJS 使用的 media stem 不构成 contract；选择
`background` stem 便于 package inspection，读取仍不依赖文件名。

清除时删除 direct `p:bg`；不删除 `p:cSld`、shape tree 或 layout relationship。清除前识别当前 image reference，
XML 更新后再按 relationship/reference 与 graph 规则回收，避免 dangling 或误删共享 payload。

## 8. 图片关系、替换与 clone-on-write

图片背景替换先比较当前 safely projected image 的 exact content type 和 bytes；完全相同时返回，不分配任何资源。
不同值按以下规则执行：

1. 解析当前 background 的 image relationship id、target 与当前 slide XML 中该 id 的全部 relationship references；
2. 若 relationship id 还被本页其他 XML node 引用，创建新 image relationship，只把 background 的
   `r:embed` 改为新 id；不得 retarget 旧 id；
3. 若 id 只由 background 使用，可以保留该 relationship id 并 retarget；
4. 只有 target graph incoming 恰好只含当前 relationship、relationship id 只被 background 引用、content type
   与扩展名不变时，才允许原位改写 target part；
5. duplicated slide、多个关系共享 target、content type/extension 改变或 shared relationship id 时，分配新 part；
6. 替换 XML/relationship 完成后，旧 relationship 仅在本页无 XML reference 时删除；旧 `/ppt/media/*` target
   仅在 package graph incoming 为零时删除。

Non-image → image 创建新 part/relationship。Image → none/solid/gradient/clear 删除 background 自己不再使用的
relationship，并按 graph 回收旧 target。Image → image 在安全时复用 relationship id；duplicate source 与 clone
最初共享 target，首次编辑后只分离被编辑页。

Slide deletion 的 shared media target 收集扩展为同时扫描：

- 现有 `p:pic` media references；
- strict direct background image reference。

删除 slide part 和 relationships 后，对预先收集的 `/ppt/media/*` targets 执行现有 incoming-zero GC。这样最后一页
删除时能回收图片背景，被其他 slide/shape/media 引用的 target 则保留。

## 9. Lossless、atomic 与 compatibility 规则

- Reader 永不修改 imported deck；
- invalid input 和 unsafe source state 在 mutation 前失败；
- supported edit 只改 direct background choice、必要 relationship 和必要 media part；
- unrelated shapes、notes、hyperlinks、layout、unknown extensions、relationship ordering 与 opaque parts 保持不变；
- public getter 始终 live-read，不缓存 bytes、relationship id 或 target URI；
- transaction rollback 精确恢复 slide XML、relationships、parts、content types、graph 和 mutation journal；
- duplication/move 不主动复制背景 payload；move 保持原 slide 与背景关系不变；
- 六种 presentation format 使用同一 slide-background contract，宏、template/slideshow 主 part 不改变；
- `powerpoint-2010` 和 `powerpoint-current` 应为 0 error/0 warning；gradient 的既有 profile diagnostics 继续生效；
- LibreOffice open/render/save/reopen 后 supported background 仍可读，并允许 LibreOffice 自身对 shared image 去重；
- `kind: 'none'` 的合法 noFill 输出优先于复制 PptxGenJS 空 `bgPr` 缺陷。

## 10. 测试与验收

### 内部 model 单元测试

- strict direct path：缺失、inheritance、noFill、solid sRGB/scheme、transparency、linear/path gradient、image；
- 拒绝 descendant trap、wrong namespace、multiple choice、`p:bgRef`、pattern/group、external/dangling/wrong-type image；
- normalization：ordinary/null-prototype、accessor/symbol/unknown key、color、transparency、gradient edge、bytes copy；
- exact no-op、none 与 clear 区别、opaque → supported replacement、schema order 和 unknown sibling preservation；
- image create、same MIME in-place、MIME change、same-slide shared relationship id、cross-slide shared target COW；
- image → fill/clear cleanup、duplicate isolation、move、delete final/shared GC、rollback、write/reopen。

### SDK 与 public-package 测试

- zero-input `PptxDocument.create()` 添加 slide 后设置四类 background；
- `setSlideBackgroundImage()` 覆盖全部 raster source forms、abort、signature/asserted MIME 和 invalid index；
- getter detachment、immediate read、clear、六种 formats 与 reopen；
- Node ESM、browser bundle/真实 Chrome、generated declarations、actual packed `@jiayunxie/pptx` tarball；
- CLI `package validate` 与 `package inspect` 可以处理 native 和 PptxGenJS fixtures。

### PptxGenJS 与客户端证据

- 锁定 `pptxgenjs@4.0.1` 生成 default、solid、50% transparency、none、PNG data URI 和 deprecated fill；
- native 打开并读取 supported variants，编辑后重新打开；
- 比较 public-output 的 fill kind、color/alpha、relationship、content type、payload hash 和 stretch/fillRect；
- native gallery 同时展示 inherited、none、solid、transparent solid、linear/path gradient、PNG/JPEG/GIF；
- PowerPoint/LibreOffice 打开与渲染检查背景覆盖、透明度、渐变和图片，无修复提示；
- 全仓 `typecheck`、tests、performance、build、pack smoke 和 CLI validation 通过。

## 11. 文档与完成定义

更新根 README、`packages/pptx/README.md`、API README、compatibility baseline、implementation progress 和
CHANGELOG，明确：

- public union 与 SDK image-source 入口；
- `undefined` inheritance/clear 与 `kind: 'none'` 的区别；
- direct background supported subset 与 layout/master scope；
- PptxGenJS 4.0.1 对等项及 intentionally corrected empty-`bgPr` behavior；
- 下一剩余项为 slide number，之后为 default color。

小项完成的必要条件是：全部新增与现有测试通过，review 未发现 correctness/lifecycle/public API 问题，真实输出
完成 package/client 验证，文档与进度表同步，工作树只包含既有未跟踪缓存目录，提交已推送且远端分支一致。
