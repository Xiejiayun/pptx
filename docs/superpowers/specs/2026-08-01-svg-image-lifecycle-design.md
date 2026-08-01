# SVG Image Lifecycle Design

## 1. 目标

在现有 PNG/JPEG/GIF 创建、source loading、sizing、`a:srcRect` 与 clone-on-write 生命周期之上，补齐
SVG 图片的创建、读取、编辑和跨运行时加载，使 native SDK 达到 PptxGenJS 4.0.1 的 SVG 公共功能对等：

- 从 path、HTTP/HTTPS、browser-relative URL、strict base64 data URI、`Uint8Array`、`ArrayBuffer`、
  `Blob`/`File`、Web stream 与 async iterable 加载 SVG；
- 从零创建 Office SVG picture，写入 SVG part、真实 PNG fallback、两个 internal image relationships 与
  canonical `asvg:svgBlip` extension；
- 读取 PptxGenJS、PowerPoint 与 LibreOffice 已有 SVG picture，不依赖固定 namespace prefix 或单一 SVG MIME；
- 原子替换 SVG + fallback payload，并在 duplicated/shared slide 中对两个 target 分别 clone-on-write；
- 复用 transform、rotation、flip、name、alt text、contain/cover/crop 与 direct source rectangle；
- Node、browser、actual npm tarball、declarations、CLI、六种 presentation format、PowerPoint 2010 profile、
  LibreOffice open/render/save/reopen 与 gallery 全部通过。

SVG 脚本执行、外链资源抓取、SVG DOM 编辑、任意图片格式互转、旧版客户端中的像素级 fallback 等价不属于
本小项。调用方可通过显式 PNG fallback 获得可控的旧客户端视觉结果。

## 2. 权威行为审计

PptxGenJS 4.0.1 对每张 SVG 写入两个 internal image relationships：fallback PNG 在前，SVG 在后。
Picture 的 direct blip 结构为：

```xml
<a:blip r:embed="fallbackPngRid">
  <a:extLst>
    <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
      <asvg:svgBlip
        xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"
        r:embed="svgRid"/>
    </a:ext>
  </a:extLst>
</a:blip>
```

`[Content_Types].xml` 必须为 `.svg` part 提供有效 SVG content type。PptxGenJS 使用
`image/svg+xml`；LibreOffice 26.2 回存时会改为 `image/svg`，但继续保留相同的 `svgBlip` extension 和双
relationship 结构。

实际审计文件覆盖 data URI、local path、cover sizing、rotation 与 flip：

- 原包为 1 slide、3 pictures、3 SVG parts、3 fallback parts、6 image relationships；
- `pptx-inspect package validate --profile powerpoint-2010` 为 0 errors / 0 warnings；
- LibreOffice 原包渲染与回存件渲染均显示全部三张 SVG；
- LibreOffice 将三份相同 SVG 和 fallback 分别去重为一个共享 target，保留每个 picture 的两条关系；
- PptxGenJS Node data URI 使用内置 broken-image PNG；local path 分支却把 SVG bytes 写进 `.png` part。

最后一项是 PptxGenJS fallback 实现缺陷，不是需要复制的 public contract。Native 输出必须始终保证
`.png` part 具有真实 PNG signature。

## 3. 方案比较

### A. 调用方始终提供 PNG fallback

Model 边界最简单，输出也完全可控，但 `PptxDocument.addImage(pathOrData)` 无法像 PptxGenJS 一样只给 SVG
就完成创建，不满足 high-level 功能对等。拒绝作为唯一方案。

### B. Node 与 browser 都引入 SVG rasterizer

可自动产生高保真 fallback，但 Node rasterizer 通常包含 native binary 或 WASM，增加平台矩阵、安装、bundle
尺寸和供应链风险。当前 `@jiayunxie/pptx` 会把依赖全部 bundle，仓库又只允许 esbuild build dependency；在
SVG 主能力并不依赖 fallback raster fidelity 的情况下不引入该依赖。

### C. Strict paired model + explicit/browser/placeholder fallback（采用）

同步 model 只接受已经准备好的 SVG bytes 与 PNG fallback bytes，负责双 part、双关系、XML 和原子生命周期。
异步 SDK 自动识别 SVG：调用方显式提供 fallback 时严格使用它；browser DOM 可用时通过 `Image + Canvas`
生成 PNG；Node 或无可用 Canvas 时使用库内固定、有效、透明 PNG placeholder。

该方案满足 PptxGenJS 的“只给 SVG 即可写入”行为，避免 native dependency，并允许对旧客户端有要求的调用方
提供真实 fallback。与 PptxGenJS 相比，Node 默认 fallback 的像素不要求相同，但 part 格式更严格有效。

## 4. Model 公共 API

在 `@pptx/model` 根入口新增：

```ts
export type SvgImageContentType = 'image/svg+xml';

export interface AddSvgImageOptions extends Partial<Transform> {
  readonly name?: string;
  readonly altText?: string;
  readonly sourceRectangle?: ImageSourceRectangle;
}

export class SlideModel {
  addSvgImage(
    svgBytes: Uint8Array,
    fallbackPngBytes: Uint8Array,
    options?: AddSvgImageOptions,
  ): ImageModel;
}

export class ImageModel {
  get svgPartUri(): string | undefined;
  get fallbackPartUri(): string | undefined;
  get isSvg(): boolean;
  replaceSvgData(svgBytes: Uint8Array, fallbackPngBytes: Uint8Array): void;
}
```

现有 `SlideModel.addImage(bytes, AddImageOptions)` 和 `RasterImageContentType` 不改变，继续是 strict raster
low-level API。`ImageModel.sourcePartUri` 继续表示 direct `a:blip` 的 target；对 SVG picture 它等于
`fallbackPartUri`，避免改变已有属性语义。`svgPartUri` 表示 extension 中的 vector target。

`AddSvgImageOptions` 与 raster options 共享 name、alt text、transform 和 source rectangle 规则，但不暴露
content type：两个 payload content type 固定为 `image/svg+xml` 与 `image/png`。两个输入必须是非空
`Uint8Array`，在 mutation 前复制；options 必须是 descriptor-safe ordinary/null-prototype own data object。

默认值与 raster creation 保持一致：x/y 为 0，width/height 为 1 inch，rotation 为 0，两个 flip 为 false，
默认 name 为当前 image count，默认 alt text 为 `preencoded.svg`。

## 5. Canonical 创建结构

`SlideModel.addSvgImage()` 在一个 package transaction 中按以下顺序执行：

1. 完成全部 payload/options validation 与 detached copies；
2. 分配 `/ppt/media/imageN.png` 并写 `image/png` fallback；
3. 分配 `/ppt/media/imageN.svg` 并写 `image/svg+xml` vector part；
4. 在 slide 上分别创建两个 internal image relationships；
5. 分配唯一 shape id，生成 canonical `p:pic` 并插入 shape-tree `p:extLst` 之前；
6. 更新 slide XML，并返回 `slide.shapes` 中同一个 live `ImageModel`。

`a:blip@r:embed` 指向 fallback relationship。其 direct `a:extLst/a:ext/asvg:svgBlip@r:embed` 指向 SVG
relationship，extension URI 和 SVG namespace 使用审计得到的固定值。`a:srcRect` 若存在，继续位于 blip 与
stretch 之间。所有 name/alt text/XML attribute 按现有 escape contract 写入。

任何 validation、part allocation、relationship、shape-tree ownership、XML serialization 或 outer transaction
失败，parts、relationships、content types、slide bytes、shape cache、graph 与 mutation journal 必须精确回滚。

## 6. Namespace-aware 读取

SVG 识别不能使用固定 `asvg`/`r` prefix，也不能只按 local name 扫描。新增内部 namespace resolver：

- element namespace 由自身 qualified name prefix 和最近祖先上的 `xmlns`/`xmlns:prefix` 解析；
- attribute namespace 由 attribute prefix 解析，default namespace 不作用于 unprefixed attribute；
- 只接受 DrawingML namespace 中 direct `a:blip` 下的 direct extension chain；
- `a:ext@uri` 必须精确等于固定 SVG extension URI；
- `svgBlip` 必须属于 Office SVG namespace；
- 两个 relationship attributes 必须属于 Office document relationship namespace；
- 两个 relationships 必须是 internal image relationships，且 target part 存在。

有效结构不要求 part MIME 必须为 `image/svg+xml`，以兼容 LibreOffice 的 `image/svg`；读取也不依赖 media
part 文件名。缺失、external、duplicate、namespace-confused、错误 extension URI、错误 relationship type、
dangling target 或多个候选 extension 都视为 unsafe SVG state：`isSvg` 为 false，SVG getter 返回
`undefined`，写 API 拒绝，不猜测、不修改原 XML。

## 7. SVG + fallback 原子替换

`replaceSvgData()` 只接受当前可安全识别的 SVG picture。它先验证并复制两个非空 payload，再在一个 transaction
中解析两个 relationship reference。两个 target 独立执行 clone-on-write：

- target 仅由当前 relationship/reference 使用且扩展名匹配时原位更新；
- target 被其他 slide relationship 或当前 XML 中其他 reference 共享时分配新 target；
- relationship id 被其他 shape reference 共享时创建新 relationship，并只改当前 shape attribute；
- SVG replacement 始终写 canonical `image/svg+xml` + `.svg`；fallback 始终写 `image/png` + `.png`；
- 原 SVG/fallback target、其他 pictures、extension siblings、source rectangle、transform 与 metadata 不变。

当 SVG 与 fallback 任一侧不安全时整个操作失败且零变化。现有 `replaceData()` 保持 raster payload API；只要
picture 中存在 Office SVG extension candidate，无论 pair 可否安全解析，都明确报错并指向
`replaceSvgData()`，避免 malformed SVG 只替换 fallback 造成 pair 进一步不一致。

Duplicated slides 最初共享两个 media targets；第一次在任一 duplicate 上调用 `replaceSvgData()` 后，该 picture
获得一对新 targets，原 slide 的两个 payload 保持不变。重新打开后两边仍应分别可识别和继续编辑。

## 8. SDK source 与 inspection

SDK 根入口新增并保留全部 raster exports：

```ts
export type ImageContentType = RasterImageContentType | SvgImageContentType;
export type ImageByteChunk = RasterImageByteChunk;
export type ImageByteStream = RasterImageByteStream;
export type ImageSource = RasterImageSource;

export interface SvgImageInfo {
  readonly contentType: 'image/svg+xml';
  readonly width: number;
  readonly height: number;
}

export type ImageInfo = RasterImageInfo | SvgImageInfo;

export function inspectSvgImage(bytes: Uint8Array): SvgImageInfo;
export function inspectImage(bytes: Uint8Array): ImageInfo;
```

`inspectRasterImage()` 与 `resolveRasterImageSource()` 继续只接受 raster。Generic inspector 先检查 raster
signatures，再以 fatal UTF-8 解码并用 lossless XML parser 检查 SVG。SVG 必须有且只有一个 SVG namespace
root element；DTD/ENTITY、非 UTF-8、HTML/XML 冒充、错误 namespace、空/truncated XML 均拒绝。

Intrinsic ratio 按以下优先级确定：

1. positive `viewBox` width/height；
2. positive root width/height，支持 unitless、px、in、cm、mm、pt、pc、Q 并换算到 96 CSS px；
3. 仅有一边且有 viewBox 时按 viewBox ratio 推导另一边；
4. 无可用 intrinsic size 时使用 SVG replaced-element 默认 300 × 150。

百分比、`auto`、负数、零、非有限或不支持单位不作为 intrinsic dimension。公开 width/height 是 detached
positive finite numbers；不执行 SVG 脚本，也不加载外部资源。

Generic source loader 复用已有 path/fetch/Blob/stream/abort/copy 语义。Data URI 只接受 canonical
`data:image/svg+xml;base64,...`，继续拒绝 raw percent encoding、额外参数、whitespace、URL-safe alphabet 与
非 canonical padding。File extension、response MIME 与 Blob name 不覆盖 bytes inspection result。

## 9. High-level `PptxDocument.addImage()`

统一后的 public surface 为：

```ts
type AddImageSourceBaseOptions = Omit<
  AddSvgImageOptions,
  'sourceRectangle' | 'width' | 'height'
> & {
  readonly contentType?: ImageContentType;
  readonly fallback?: ImageSource;
  readonly signal?: AbortSignal;
};

export type AddImageSourceOptions = AddImageSourceBaseOptions & (
  | {
      readonly sizing?: undefined;
      readonly width?: number;
      readonly height?: number;
    }
  | {
      readonly sizing: ImageSizing;
      readonly width?: never;
      readonly height?: never;
    }
);

export class PptxDocument {
  addImage(
    slideIndex: number,
    source: ImageSource,
    options?: AddImageSourceOptions,
  ): Promise<ImageModel>;
}
```

`ImageSizing`、`ImageCropRegion`、`ImageSizingResult` 与 `calculateImageSizing()` 是现有 Raster sizing contract
的 format-neutral 名称；现有 `RasterImageSizing`、`RasterImageCropRegion`、`RasterImageSizingResult` 与
`calculateRasterImageSizing()` 保留为兼容 API。Generic sizing 对 raster 继续要求 safe-integer pixels，对 SVG
接受 inspector 返回的 positive finite CSS pixel dimensions，最终仍量化为相同 direct `a:srcRect`。

High-level 流程：

1. 在 I/O 前验证 slide index、options container、known keys、signal、contentType、sizing 互斥规则；
2. 异步解析并复制 primary source，检查 asserted/declared content type；
3. raster 分支拒绝 `fallback`，沿用现有 sizing 与 `slide.addImage()`；
4. SVG 分支计算 sizing，然后解析 explicit fallback 或自动生成 fallback；
5. explicit fallback 必须严格检测为 PNG，不能是 JPEG/GIF/SVG，也不能只靠声明通过；
6. 调用同步 `slide.addSvgImage()` 完成唯一 package mutation。

Primary/fallback resolution、inspection、Canvas、abort 或 semantic validation 失败都不能修改 package。Primary
被识别为 raster 时，在不读取 fallback getter/Blob/stream 的情况下拒绝 fallback，以避免无意义 I/O。

## 10. 自动 fallback 策略

Fallback 优先级固定：

1. `options.fallback`；
2. browser DOM 中可用的 `Image`、`document.createElement('canvas')` 与 2D context；
3. library 内置固定透明 PNG placeholder。

Browser 分支使用 detached SVG bytes 创建同源 object URL，等待 image decode，按 intrinsic dimensions 创建
canvas，并导出 `image/png`。canvas dimension 必须是 positive safe integer，并限制单边与总像素预算，防止恶意
SVG 触发过量分配。decode、tainted canvas、缺失 2D context、预算超限或 `toBlob`/`toDataURL` 失败时回退到
内置 placeholder，而不是写入错误格式 part。

内置 placeholder bytes 为不可变源码常量，每次返回 detached copy，并通过同一个 PNG inspector 自测。它只用于
不理解 Office SVG extension 的客户端；现代 PowerPoint 和 LibreOffice 使用 SVG target。对 fallback 视觉有要求的
调用方必须显式提供 PNG。

## 11. 测试与验收

### 11.1 Model

- normalization/XML unit tests 覆盖 canonical extension、固定 URI/namespace、escaping、transform、flip、
  source rectangle、defaults、payload copies 与全部 unsafe inputs；
- zero-input package 创建后立即得到 live identity、两个 parts、两个 relationships 与 effective content types；
- PptxGenJS package、alternate prefixes、LibreOffice `image/svg`、shared target、relationship reuse、unknown
  extension siblings 与 malformed/namespace-confused fixtures；
- `replaceSvgData()` exclusive update、两侧 clone-on-write、same-slide shared rId、duplicate slide、rollback、
  write/reopen/edit-again、move/delete dependency lifecycle；
- pptx/pptm/potx/potm/ppsx/ppsm 六格式 round-trip。

### 11.2 SDK

- SVG inspector 覆盖 viewBox、physical/CSS units、one-side derivation、300×150 default、UTF-8 与拒绝矩阵；
- 每种 source kind、strict data URI、path/URL/browser-relative URL、abort、copy isolation 与 MIME assertion；
- explicit fallback、Node placeholder、browser Canvas success/failure fallback；
- raster existing tests保持不变，raster + fallback 在 fallback 未读取前拒绝；
- contain/cover/equal ratio/crop final transform 与 `srcRect` 对照 PptxGenJS public final state。

### 11.3 制品与客户端

- model/sdk focused tests、全量 `pnpm test`、`pnpm typecheck`、`git diff --check`；
- `@jiayunxie/pptx` Node/browser build 无 static Node import，clean build SHA-256 deterministic；
- actual `pnpm pack` tarball 的 Node、browser、types 与 CLI smoke 覆盖 SVG create/open/edit/reopen；
- gallery 覆盖 data/path/bytes/Blob/stream/URL、explicit/default fallback、contain/cover/crop、rotation/flips、
  duplicate + replacement、特殊字符 name 与非空 alt text；
- 原件与 LibreOffice round-trip 均 strict reopen，PowerPoint 2010 validation 为 0 errors / 0 warnings；
- 解包检查每个 original SVG picture 都有真实 PNG fallback、SVG target 与两个 internal relationships；
- LibreOffice open/render/save/reopen 后记录 shared-target 去重、`image/svg` MIME normalization、transform 和
  source-rectangle quantization；逐页 full-size 视觉检查，无 clipping、broken image 或意外空白；
- 更新 root/package README、CHANGELOG、implementation progress 与 PptxGenJS compatibility baseline。

## 12. 完成定义

只有在 public types、runtime、model editing、source loading、sizing、fallback、六格式、actual package、browser、
CLI、PptxGenJS conformance、LibreOffice round-trip、文档与全量 release gates 同时通过后，SVG 小项才标记完成。
不得以“SVG part 能写入”替代完整双向生命周期，也不得把 PptxGenJS 的伪 PNG fallback 缺陷带入 native 输出。
