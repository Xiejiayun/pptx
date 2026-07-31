# Raster Image Sizing and Source Rectangle Design

## 1. 目标

在已经完成的 PNG/JPEG/GIF source loader 与 intrinsic pixel inspector 之上，增加可用于从零创建和编辑
PPT 的图片 sizing 能力：

- high-level `contain` 把完整图片等比放入目标框；
- high-level `cover` 用居中裁剪填满目标框；
- high-level `crop` 把明确的 source pixel region 映射到目标框；
- low-level direct `a:srcRect` create/read/edit/clear；
- PptxGenJS 4.0.1 `sizing` public-output final-state conformance。

现有未传 sizing 的行为完全不变：图片默认 x/y 为 0、width/height 为 1 inch，不依据 natural pixels
隐式改变布局。Sizing 只在调用方显式提供时生效。

## 2. 当前状态与问题

当前 native 已经能够：

- 从 path、URL、data URI、bytes、Blob/File 和 streams 加载 PNG/JPEG/GIF；
- 以 signature 检测 canonical content type 与 raw pixel width/height；
- 原子创建 media part、image relationship 和 canonical `p:pic`；
- 读取并编辑 transform，clone-on-write 替换 embedded payload；
- 在 source resolution、inspection、assertion 或 model mutation 失败时保持 package 零变化。

但 picture fill 仍固定写成 `<a:stretch><a:fillRect/></a:stretch>`，`ImageModel` 没有 direct source
rectangle API。PptxGenJS 4.0.1 的 `ImageProps.sizing` 会输出 `a:srcRect`，同时把 outer `w/h` 当成
source coordinate size、把 nested `sizing.w/h` 当成 target frame；`crop` 又把 nested `x/y/w/h`
复用为裁剪窗口。这种行为可以产生有效最终 OOXML，但不能作为 native API 的无歧义单位 contract。

## 3. 方案比较

### A. 只在 SDK 计算 transform，不公开 direct `srcRect`

Contain 可以通过缩小 transform 得到视觉近似，但目标 frame 不再稳定，cover/crop 也无法表达。已有 PPT 中的
`a:srcRect` 仍不可读取或编辑，不满足 bidirectional editing。拒绝。

### B. 把 PptxGenJS `sizing` 原样加入 `SlideModel.addImage()`

Model 必须知道 intrinsic pixels，因而会依赖 SDK inspector 或自行重复图片解析；outer 与 nested width/height
也继续存在两个含混 truth sources。同步 OOXML model 不应承担 source inspection。拒绝。

### C. Direct source rectangle model + intrinsic-aware SDK sizing（采用）

Model 公开严格、平台无关的 `ImageSourceRectangle` create/read/edit/clear，数值直接对应 DrawingML
`a:srcRect` 百分比。SDK 提供纯 sizing calculator，并在 source 已解析、intrinsic dimensions 已知后把
contain/cover/crop 归一为 target width/height + direct source rectangle，再调用 existing atomic model core。

该边界同时满足：

- existing deck 可编辑；
- source loader 不进入 model；
- sizing 公式可纯测试；
- PptxGenJS final XML 语义可比较；
- 后续 SVG fallback 可以复用相同 frame/source-rectangle contract。

## 4. 公共 API

### 4.1 Model direct state

在 model 根入口导出：

```ts
export interface ImageSourceRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface AddImageOptions extends Partial<Transform> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
  readonly sourceRectangle?: ImageSourceRectangle;
}

export class ImageModel {
  get sourceRectangle(): Readonly<ImageSourceRectangle> | undefined;
  set sourceRectangle(value: ImageSourceRectangle | undefined);
}
```

四个值使用百分数：`1` 表示 1%，序列化时乘 1000 成为 DrawingML integer percentage。正值从对应边缘
裁去图片，负值把 source rectangle 扩展到图片之外，因而能够表达 contain letterboxing。Public value
统一量化到 0.001%。

`sourceRectangle: undefined` 表示没有 direct `a:srcRect`；显式四个 0 表示存在一个 zero rectangle。
Getter 返回 detached frozen snapshot。Creation 与 setter 都只接受 descriptor-safe ordinary/null-prototype
own data properties，四个字段必须齐全，不接受 unknown、inherited、symbol 或 accessor property。

### 4.2 SDK sizing

在 SDK 根入口导出：

```ts
export interface RasterImageCropRegion {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type RasterImageSizing =
  | {
      readonly type: 'contain' | 'cover';
      readonly width: number;
      readonly height: number;
    }
  | {
      readonly type: 'crop';
      readonly width: number;
      readonly height: number;
      readonly source: RasterImageCropRegion;
    };

export interface RasterImageSizingResult {
  readonly width: number;
  readonly height: number;
  readonly sourceRectangle: Readonly<ImageSourceRectangle>;
}

export function calculateRasterImageSizing(
  info: RasterImageInfo,
  sizing: RasterImageSizing,
): Readonly<RasterImageSizingResult>;
```

`RasterImageSizing.width/height` 是最终 picture frame 的 native EMU，与 `Transform.width/height` 一致。
Crop `source` 使用 raw source pixels；允许 finite fractional pixel boundaries，但必须完整落在 detected image
范围内。这样 source coordinates 与 target layout units 永不混用。

High-level `AddImageSourceOptions` 从 `Omit<AddImageOptions, 'contentType' | 'sourceRectangle'>` 派生并增加
`contentType?`、`signal?` 与 `sizing?`。Type surface 与 runtime 同时保证：

- 无 `sizing` 时继续允许 top-level `width`、`height`，并沿用已有 direct transform behavior；
- 有 `sizing` 时 top-level `width`、`height` 不可同时提供；
- top-level `x`、`y` 仍是 picture frame placement；rotation 与 flips 保持独立；
- `sizing.width/height` 成为最终 transform extents。

Direct `sourceRectangle` 只属于 low-level model creation 和 live `ImageModel` editing；high-level source loader
通过 `sizing` 生成它，避免一个调用里出现两个 source-rectangle truth sources。

## 5. Direct `a:srcRect` contract

### 5.1 归一化

每个 edge 必须是 finite number。归一化为：

```ts
const raw = Math.round(percent * 1_000);
const normalized = raw / 1_000;
```

Raw value 必须落在 OOXML signed Int32 范围。每个 normalized edge 必须 `< 100`，且：

```text
left + right < 100
top + bottom < 100
```

这保证 source width/height 为正，同时允许 contain 所需的任意合法负扩展。`NaN`、infinity、numeric
string、missing field、extra field、正向 100% 或 pair sum 达到/超过 100% 都在 mutation 前拒绝。

### 5.2 读取

只读取 picture direct `p:blipFill/a:srcRect`。没有 direct element 返回 `undefined`。一个安全 direct element
满足以下条件时公开 snapshot：

- 至多一个 direct `a:srcRect`；
- 没有 child content；
- `l/t/r/b` 缺省按 0；
- present attribute 是 canonical signed Int32 lexical value；
- normalized result 满足同一 source-area contract。

Malformed、duplicate、namespace-confused 或 unsafe state 返回 `undefined`，读取不修改 bytes。Noncanonical
但合法百分号 lexical form 本阶段无损保留但不猜测为 public direct state。

### 5.3 创建与编辑

`SlideModel.addImage()` 收到 source rectangle 时，在 direct `a:blip` 后写：

```xml
<a:srcRect l="..." t="..." r="..." b="..."/>
<a:stretch><a:fillRect/></a:stretch>
```

Getter/setter 只拥有 direct `a:srcRect`：

- assignment whole-replace 这一个 element；
- `undefined` 只移除 direct `a:srcRect`；
- absent-to-value insertion 位于 direct `a:blip` 之后、`a:tile`/`a:stretch` 之前，并复用 in-scope
  DrawingML prefix，缺失时使用 canonical `a` binding；
- 不修改 `a:blip`、`a:tile`/`a:stretch`、fillRect、effects、extensions、relationships 或 transform；
- normalized same-value assignment 是 exact bytes/journal no-op；
- 一个 malformed direct element 可由 valid replacement/clear 修复；
- duplicate direct elements 因 ownership ambiguous 而拒绝写入；
- 所有 mutation 使用 existing package transaction 并可由 outer transaction 回滚。

## 6. Pure sizing formulas

设 image dimensions 为 `Iw × Ih` pixels，target frame 为 `Fw × Fh` EMU。所有计算只使用 ratio，最后
通过 direct source-rectangle normalizer 量化到 0.001%。

### 6.1 Cover

图片等比放大到覆盖 frame，居中裁去超出部分：

```ts
const imageRatio = Ih / Iw;
const frameRatio = Fh / Fw;
if (frameRatio > imageRatio) {
  const renderedWidth = Fh / imageRatio;
  const horizontal = 50 * (1 - Fw / renderedWidth);
  rect = { left: horizontal, right: horizontal, top: 0, bottom: 0 };
} else {
  const renderedHeight = Fw * imageRatio;
  const vertical = 50 * (1 - Fh / renderedHeight);
  rect = { left: 0, right: 0, top: vertical, bottom: vertical };
}
```

### 6.2 Contain

图片等比缩小到完整落入 frame，居中用 negative source rectangle 表达空白区域：

```ts
if (frameRatio > imageRatio) {
  const renderedHeight = Fw * imageRatio;
  const vertical = 50 * (1 - Fh / renderedHeight); // negative
  rect = { left: 0, right: 0, top: vertical, bottom: vertical };
} else {
  const renderedWidth = Fh / imageRatio;
  const horizontal = 50 * (1 - Fw / renderedWidth); // negative
  rect = { left: horizontal, right: horizontal, top: 0, bottom: 0 };
}
```

Equal aspect ratio 产生显式 zero rectangle。极端 ratio 若量化后超出 Int32 percentage range 则确定性拒绝，
不 clamp 成另一种视觉结果。

### 6.3 Crop

设 pixel region 为 `x/y/width/height`：

```ts
left = 100 * x / Iw;
top = 100 * y / Ih;
right = 100 * (Iw - x - width) / Iw;
bottom = 100 * (Ih - y - height) / Ih;
```

`x/y >= 0`、`width/height > 0`、`x + width <= Iw`、`y + height <= Ih`。目标 frame 与 source
region aspect ratio 可以不同；crop 明确允许把选区 stretch 到调用方指定的 final frame。

## 7. `PptxDocument.addImage()` data flow

执行顺序固定为：

1. 验证 slide index；
2. descriptor-safe normalize source options 和 sizing shape；
3. resolve source bytes；
4. inspect signature、content type 与 pixel dimensions；
5. 检查 MIME assertions；
6. 若有 sizing，调用 pure calculator 得到 width/height/sourceRectangle；
7. 调用 `SlideModel.addImage()` 执行 existing atomic package mutation。

前六步不修改 package。Invalid sizing 不消费额外 source，也不产生 part、relationship、shape id 或 journal
entry。Caller 在 Promise pending 期间修改 sizing/source objects 不得影响最终结果；normalization 必须在第一
个 await 之前完成并 detach nested crop region。

## 8. PptxGenJS 4.0.1 对等边界

Conformance 使用 public `slide.addImage({ path|data, x, y, w, h, sizing })` 生成实际文件，并比较：

- contain 横向和纵向 letterbox；
- cover 横向和纵向 crop；
- equal-ratio zero source rectangle；
- crop source window；
- final transform extents、rotation、flips、name、alt text；
- direct `a:srcRect` percentages、embedded payload、content type 与 internal relationship。

PptxGenJS case 会把 outer `w/h` 设置为与 fixture intrinsic ratio 等价的 source coordinate size，再把 nested
`sizing.w/h` 设置为 target frame。Native 依据真实 intrinsic pixels 计算，因此两者在 valid、明确输入上得到
相同 final `srcRect` 与 transform。

明确记录而不复制的 divergence：

- PptxGenJS 接受 outer/nested 两套含混 width/height；native sizing 只有一个 target frame；
- PptxGenJS crop coordinates 与 inches/percentage smart parsing 共用；native crop source 明确使用 pixels；
- PptxGenJS 使用 truthy fallback，可能把 0 或 invalid runtime values 静默替换；native 在 mutation 前拒绝；
- Native source rectangle snapshot/editing 是 PptxGenJS 没有的 bidirectional extension。

## 9. 测试矩阵

### Model normalization/creation

- positive crop、negative contain、zero rectangle 与 0.001% quantization；
- detached/frozen inputs，ordinary/null-prototype options；
- missing/extra/inherited/symbol/accessor、NaN/infinity/Int32 overflow、edge/pair 100% rejection；
- canonical child ordering，default creation bytes unchanged when omitted；
- failure allocation/outer transaction rollback。

### Existing image lifecycle

- absent/valid/explicit-zero/negative direct reads；
- missing attributes default zero；malformed/duplicate/namespace-confused reads；
- whole replace、repair、clear、exact no-op；
- preserve blip/tile/stretch/effects/extensions/relationships/transform；
- duplicate isolation、move/delete/reopen 与六种 formats。

### Pure calculator

- landscape/portrait/square image × landscape/portrait/square frame；
- cover/contain horizontal and vertical branches；
- exact expected 0.001% rounding；
- crop full image, centered subset, fractional pixels and edge-aligned region；
- invalid info/sizing descriptors, unsafe frame extents, out-of-bounds/empty crop and extreme ratio overflow；
- detached frozen result and caller-object mutation isolation。

### SDK/conformance/release

- source options type union forbids width/height conflicts with sizing，并始终拒绝 high-level direct
  `sourceRectangle`；
- source failure and sizing failure are zero-mutation；
- all portable source kinds and six presentation formats；
- PptxGenJS actual tarball public-output comparison；
- packed Node/browser/declaration/CLI smoke and browser no-static-Node-import check；
- multi-ratio gallery, strict validation, LibreOffice open/save/reopen, render/overflow and direct-state diff；
- focused tests、typecheck、full test、performance、build、dist reproducibility 与 `git diff --check`。

## 10. 实施小项

1. Model `ImageSourceRectangle` normalization 与 picture creation；
2. Existing `ImageModel.sourceRectangle` read/edit/clear lifecycle；
3. Public pure contain/cover/crop calculator；
4. `PptxDocument.addImage()` sizing integration 与 strict source-option union；
5. PptxGenJS conformance 和 packed artifact verification；
6. Gallery/client evidence 与 public docs/status update。

每项完成 focused tests、self-review、commit 和 push 后再进入下一项。

## 11. Self-review

- 没有 placeholder 或未决单位；source pixels、percentage 与 target EMU 各自只有一种含义。
- Direct source rectangle 与 high-level sizing ownership 分离，model 不依赖 source I/O 或 image parser。
- Contain 允许合法 negative percentages，cover/crop 保持 positive visible area。
- Runtime、TypeScript 与 transaction conflict rules 一致。
- Omitted sizing 不改变现有 default transform 或 canonical picture bytes。
- Existing `a:srcRect` 可读取、编辑、修复、清除并保留无关 fill/relationship state。
- PptxGenJS 对等声明只覆盖可明确映射的合法 final state，已列出不复制的 runtime divergences。
- Scope 不包含 SVG、rounding、transparency、alt-text editing、image hyperlink/shadow/placeholder 或 public deletion/GC。
