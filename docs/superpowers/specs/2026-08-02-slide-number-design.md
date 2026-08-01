# Slide Number Design

## 1. 目标

完成一个可独立验收的 slide number 小项，使库能够从零创建、严格读取、编辑、清除和重排 PowerPoint 页码字段，
并覆盖 PptxGenJS 4.0.1 实际可见的公开能力：

- slide direct page-number field；
- layout 与 slide master direct page-number placeholder；
- x/y/width/height、horizontal/vertical alignment、RTL 和 text-box margin；
- font family、font size、sRGB/theme color、transparency、bold、italic 和 language；
- presentation `firstSlideNum` 的创建、读取、编辑、清除与缓存文本同步；
- duplicate/move/delete 后 direct field cached text 与最终页序一致；
- `pptx/pptm/potx/potm/ppsx/ppsm` 六种格式、Node、browser、实际 npm tarball、declarations、CLI、
  PowerPoint/LibreOffice 与 PptxGenJS imported-output 全链路验证。

本项不实现 declarative `defineSlideMaster()`、按名称选择 layout、date/footer/header、页码前后缀、章节内重新编号、
Roman/字母编号或自定义 field formula。普通 text shape 中不是 `p:ph@type="sldNum"` 的 `a:fld@type="slidenum"`
继续作为普通/opaque text content 保留。default text color、完整 master/layout/placeholder definition 和 advanced text
属于后续小项。

用户已明确要求实现方持续决定最佳方案且不设置询问停顿点。本设计完成 self-review 后直接写实施计划，并按每个
可独立 review 的小项执行 tests、commit 与 push。

## 2. 当前状态与权威行为审计

当前库没有 `slideNumber` 或 `firstSlideNumber` public API。`SlideModel.shapes` 会把 imported slide-number placeholder
当作普通 text `ShapeModel`，因此内容可以被看到，但不能以 field-aware 方式创建或安全编辑。`MasterModel` 与
`LayoutModel` 只公开 name/theme/layout/placeholder inspection 和 raw part lifecycle，没有页码属性。

现有基础设施可复用：

- `LosslessXmlDocument` 可做 namespace-aware direct-child projection 与 span-level patch；
- OPC transaction 与 mutation journal 可保证跨 presentation/slide part 的 rollback；
- presentation 已有 add/duplicate/move/delete、section ordering 和 stable slide model lifecycle；
- text-box margin、vertical alignment、paragraph alignment、RichText color/style validation 已锁定单位和范围；
- master/layout codec 已拥有对应 part model、relationship traversal、copy/delete/relink 和 live model cache；
- validator、actual-package smoke、real-Chrome、gallery、LibreOffice 与 PowerPoint control 流程可沿用。

### PptxGenJS 4.0.1 实际输出

审计只通过 public constructor、`addSlide()`、`slide.slideNumber`、`defineSlideMaster()` 和 `write()` bytes 进行，
production conformance test 不读取 `_slides`、`_slideNumberProps` 等私有字段。

有效 `slide.slideNumber = { ... }` 会在当前 slide 末尾生成一个 direct `p:sp`：

- `p:ph type="sldNum" sz="quarter" idx="4294967295"`；
- fixed `p:cNvPr id="25"` 与 fixed field GUID；
- `a:xfrm/a:off+a:ext`；
- exactly one `a:fld type="slidenum"`，cached text 为当前一基页序；
- align 写在 direct `a:pPr`，bold/lang 写在 field `a:rPr`；
- font size/color/family 写在 `a:lstStyle/a:lvl1pPr/a:defRPr`；
- margin 与 vertical alignment 写在 `a:bodyPr`。

空对象的 effective defaults 是 x/y `0`、width `800000` EMU、height `300000` EMU、left alignment、
`lang="en-US"` 和 `bold=false`。合法 sRGB/theme color、font face/size、bold、left/center/right、top/middle/bottom、
uniform/four-side margin 都有可见输出。

Setter 还隐式把同一个 options 对象写入唯一 master 与名为 `DEFAULT` 的 layout。该副作用有以下已观察问题：

- 只影响当前 slide 的 direct field；之后新增普通 slide 不获得 direct field；
- 后一次 setter 覆盖 shared master/default-layout state，清除一个 slide 也会清除祖先 state；
- master cached text 是字符串 `null`，layout cached text 是内部编号 `1000+`；
- fixed shape id `25` 可能与已有 shape 冲突；
- master 的 `p:hf@sldNum` 仍为 `0`；
- zero width/height 被 falsy fallback 改为 defaults，而不是拒绝；
- `justify` fallback 为 left；
- declared `italic`、`transparency`、RTL 等字段被忽略；
- position 没有 viewport validation，完全位于页面外也会静默输出。

这些缺陷不复制。Native 提供显式 slide/layout/master ownership、合法唯一 shape id、合法 cached text、strict input、
explicit zero preservation 和完整 supported style。Adapter 必须能读取 PptxGenJS 的合法 direct slide/layout shape，
并锁定 master `hf=0`、`null/1000` cache 与 ignored-style 差异。

### 客户端继承实验

在同一 zero-input native deck 中分别只向 slide、layout 或 master 写入 namespace-correct `sldNum/slidenum` field，
并对 master `p:hf@sldNum` 组合 `0/1`，LibreOffice 26.8 headless PDF 的五组结果为：

- direct slide field 可见且 field text 被渲染；
- layout-only 在 `hf=0` 与 `hf=1` 都不可见；
- master-only 在 `hf=0` 与 `hf=1` 都不可见。

因此 portable creation 不能只依赖继承。`SlideModel.slideNumber` 必须写 direct slide field；layout/master 属性仍用于
精确编辑 PowerPoint hierarchy，但不会替代 direct slide API，也不会偷偷改写其他 owner。

## 3. 方案比较

### A. 只实现 direct slide field

实现最小且 LibreOffice 可见，但 imported/master-created deck 仍无法语义编辑 layout/master page-number placeholder，
与既定 compatibility matrix 的 master-level 要求不符。

### B. 只实现 master/layout inheritance

OOXML ownership 接近 PowerPoint UI，但跨客户端实测不可移植，也无法与 PptxGenJS 每个 configured slide 的 direct
output 对等。

### C. Shared strict codec + 三个显式 owner（采用）

在 `@pptx/codecs` 增加一个 package-independent slide-number value/reader/writer，`SlideModel`、`LayoutModel` 与
`MasterModel` 各自调用同一实现。Slide setter 只改 direct slide；layout/master setter 只改自己的 part。Master setter
另同步自己 direct `p:hf@sldNum`，但 slide setter 不产生 hidden ancestor mutation。Presentation lifecycle 只同步
direct slide field 的 cache，不重写 layout/master。

该方案同时满足 portable creation、hierarchy editing、lossless explicit ownership 和 PptxGenJS final-state import，
职责也不会把 page-number field 伪装成普通 text shape feature。

## 4. 公共 API

`@pptx/codecs` 根入口新增并导出：

```ts
export type SlideNumberColor =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string };

export interface SlideNumberMargins {
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly left?: number;
}

export type SlideNumberMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | SlideNumberMargins;

export interface SlideNumberTextStyleOptions {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

export interface SlideNumberTextStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

export interface SlideNumberOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly align?: 'left' | 'center' | 'right' | 'justify';
  readonly rtl?: boolean;
  readonly valign?: 'top' | 'middle' | 'bottom';
  readonly margin?: SlideNumberMarginInput;
  readonly style?: SlideNumberTextStyleOptions;
}

export interface SlideNumber {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly align: 'left' | 'center' | 'right' | 'justify';
  readonly rtl: boolean;
  readonly valign?: 'top' | 'middle' | 'bottom';
  readonly margin?: SlideNumberMargins;
  readonly style: SlideNumberTextStyle;
}
```

`x/y/width/height` 使用 native EMU，与 shape/image API 一致；调用方用 `inches()` 显式转换。Margin 使用 points，
与现有 text-box margin API 一致。Defaults 固定为：

```ts
{
  x: 0,
  y: 0,
  width: 800_000,
  height: 300_000,
  align: 'left',
  rtl: false,
  style: { lang: 'en-US', bold: false, italic: false },
}
```

`@pptx/model` 的 `SlideModel` 新增：

```ts
export class SlideModel {
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}
```

`@pptx/codecs` 的 live models 新增同名 direct owner API：

```ts
export class LayoutModel {
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}

export class MasterModel {
  get slideNumber(): Readonly<SlideNumber> | undefined;
  set slideNumber(value: SlideNumberOptions | undefined);
}
```

`@pptx/model` 的 `PresentationModel` 与 `@pptx/sdk` 的 create options 新增：

```ts
export class PresentationModel {
  get firstSlideNumber(): number | undefined;
  set firstSlideNumber(value: number | undefined);
}

export interface CreatePresentationOptions {
  readonly firstSlideNumber?: number;
}
```

Getter 只返回 direct `p:presentation@firstSlideNum`；`undefined` 表示使用 OOXML default `1`。Setter 接受 signed
Int32 safe integer；`undefined` 只删除 direct attribute 并重新使用 `1`。改变该值后所有 safely recognized direct
slide-number fields 的 cached text 以 `(firstSlideNumber ?? 1) + slideIndex` 同步。

所有 value getter live-read、detached、deep-frozen。相同 semantic value、清除 absent direct state 或设置相同 direct
`firstSlideNum` 都是 parts/XML/relationships/graph/journal exact no-op。

## 5. Validation 与 normalization

`normalizeSlideNumberOptions(value)` 只接受 descriptor-safe ordinary/null-prototype own-data object；拒绝 array、class
instance、accessor、symbol key、unknown field 和 nested exotic object，并在 package access 前完成全部 validation。

- x/y 必须 finite，量化到整数 EMU，并落在 DrawingML signed coordinate range；explicit zero 保留；
- width/height 必须 finite、量化到整数 EMU、严格大于 0 且不超过 DrawingML positive coordinate maximum；
- align、valign 是 closed enums，RTL/bold/italic 是 strict boolean；
- margin 复用 signed Int32 EMU point conversion，tuple 必须 dense ordinary length 4，空 object canonicalize 为 absent；
- font family/lang 必须 non-empty、XML-safe string；
- font size 必须 `1..4000pt` 并量化到 `0.01pt`；
- sRGB 必须六位 hex 并转 uppercase，scheme 必须是支持的 theme token；
- transparency 必须 finite `0..100` 并量化到 `0.001%`；只有 transparency 时显式使用 scheme `tx1`；
- caller object、margin tuple/object、style 与 color 都在 mutation 前复制并冻结。

`normalizeFirstSlideNumber()` 接受 signed Int32 safe integer。Cached text 使用 JavaScript safe integer 计算；
`firstSlideNum` 本身超范围或非法 lexical direct attribute 读取为 `undefined`，显式合法 setter 可以 repair 该单一
attribute。

## 6. Strict direct field reader

Reader 只识别 owner part 根下 namespace-correct direct chain：

```text
p:sld|p:sldLayout|p:sldMaster
  / p:cSld
  / p:spTree
  / exactly one direct p:sp whose p:nvSpPr/p:nvPr/p:ph@type is sldNum
```

该 shape 必须有 unique valid `p:cNvPr@id`、unique direct `p:spPr`、unique `a:xfrm/a:off+a:ext`、unique direct
`p:txBody/a:bodyPr`、一个 direct paragraph，以及该 paragraph 中 exactly one direct `a:fld@type="slidenum"`。
Field cached `a:t` 必须存在且是唯一 direct text child，但不进入 public value equality。

Reader 支持两种 style placement：

1. Native canonical direct field `a:rPr`；
2. PptxGenJS `a:lstStyle/a:lvl1pPr/a:defRPr` defaults + field `a:rPr` override。

它按 DrawingML precedence 合并支持字段。Unknown harmless outer siblings 保留；wrong namespace、duplicate placeholder、
duplicate field、缺失 transform/text body、无效 coordinate/style、多个 conflicting default style、unsupported field type 或
ambiguous direct structure 返回 `undefined`，且读取不修复 package。普通 `a:r`、break、其他 field 或额外 paragraph 不被
误识别为 slide number。

Master getter 还要求 direct `p:hf@sldNum` 不为 false；PptxGenJS 的 shape + `sldNum="0"` 因此记录为 disabled/unsupported
master state，但它的 direct slide 与 layout shape 仍可分别读取。显式 master assignment canonicalize 该状态并启用；
显式 clear 可删除 unique direct slide-number placeholder 并把 master flag 设为 false。

## 7. Canonical OOXML 写入

新建 direct field 时在 owner `p:spTree` 的 `p:extLst` 前追加一个 `p:sp`；无 extLst 时追加到 shape tree 末尾。
Shape id 使用当前 part 全部 namespace-correct `p:cNvPr@id` 的最小可用正整数，不使用 PptxGenJS fixed `25`。Placeholder
index 使用 `4294967295`，若该 index 已被其他 direct placeholder 占用，则使用当前 part 最小可用 unsigned index。

Canonical shape 为：

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="N" name="Slide Number N"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="sldNum" sz="quarter" idx="I"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="..." y="..."/><a:ext cx="..." cy="..."/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr .../>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="..." rtl="..."/>
      <a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum">
        <a:rPr ...><!-- supported style --></a:rPr>
        <a:t>...</a:t>
      </a:fld>
      <a:endParaRPr lang="..."/>
    </a:p>
  </p:txBody>
</p:sp>
```

Color/transparency 使用 direct `a:solidFill/a:srgbClr|a:schemeClr/a:alpha`。font family 同时写 latin/ea/cs，防止
客户端 script fallback 改变外观。No color 表示继承；transparency-only canonicalize 为 `tx1 + alpha`。Native
italic/transparency/justify/RTL 明确写合法 OOXML，不复制 PptxGenJS ignored/fallback behavior。

对 safely recognized existing shape：same-value 不写；不同 supported assignment 只 patch transform、bodyPr、paragraph、
field cache 和 supported style spans，保留 shape id/name、placeholder index、field GUID、unknown attributes/children、
extension、周边 shape order 与所有 unrelated parts。PptxGenJS `defRPr` 中本 API own 的 style spans 会被清除或更新，
避免旧 default 与新的 field `rPr` 叠加；unknown default style 继续保留。对 unique `p:ph type="sldNum"` 但内部结构
不支持的 shape，explicit supported assignment 可按调用方意图 canonicalize 整个 shape并保留 valid unique id/index；
若 id/index 也 ambiguous 则在 mutation 前拒绝。`undefined` 只删除 unique direct slide-number placeholder；没有 direct
placeholder 时 exact no-op，multiple placeholders 时拒绝而不是猜测。

Master supported assignment 同时把 unique/created direct `p:hf@sldNum` 设为 `1`，保留 hdr/ftr/dt 与 unknown attributes；
clear 设为 `0`。缺失 `p:hf` 时按 schema order 插在 direct `p:txStyles` 或 `p:extLst` 前；两者都不存在时插在
`p:sldLayoutIdLst` 后或 root 尾部。

## 8. Cached text 与 presentation lifecycle

Field 是 dynamic `slidenum`；cached `a:t` 只用于未计算 field 的 renderer/fallback。Native 永远按 presentation order 写
可读十进制 text：

```text
(direct firstSlideNum, default 1) + zero-based slide index
```

Layout/master canonical cache 使用 `‹#›`，不写 PptxGenJS 的 internal `1000+` 或 `null`。Public `SlideNumber` 不暴露
cache，因此客户端刷新 cache 不会使 getter semantic value变化。

Lifecycle hook：

- direct slide setter 使用当前 live index；
- add slide 追加时既有 ordinal 不变，新页无 direct field，不产生多余写入；
- duplicate 复制 direct field 后只更新 duplicate cache；source bytes 不变；
- move 更新 ordinal 发生变化且 safely recognized 的 direct fields；
- delete 在 part/relationship 删除后更新剩余 affected fields；
- firstSlideNumber setter 更新全部 recognized direct fields；
- section move/rename/assignment 不改变 presentation order，不更新 cache；
- open/get/write 永不为刷新 cache 而 mutation；unsupported imported fields 原样保留，PowerPoint 可自行计算 dynamic field。

所有跨 presentation 与多个 slide part 的同步在一个 outer OPC transaction 中完成；任何 injection failure 恢复
presentation XML、全部 slide parts、relationships、parts、model identity 和 mutation journal。

## 9. Lossless、atomic 与 compatibility 规则

- Reader 永不修改 imported deck；
- invalid input 与 ambiguous owner 在 mutation 前失败；
- supported write 只改 owner slide-number shape、master 自己的 hf flag 或 presentation 自己的 firstSlideNum；
- notes 的独立 slide-number placeholder、header/footer、普通 field、shapes、relationships、extensions 与 unknown XML 保持；
- `SlideModel.shapes` 继续反映 physical placeholder shape，不在本项引入 collection filtering 或新 model identity；
- duplicate/move/delete 保持现有 sections、hyperlinks、media/charts 和 dependency GC 行为；
- 六种 presentation format 共用同一 codec，宏、template/slideshow 主 part/content type 不改变；
- `powerpoint-2010` 与 `powerpoint-current` 对 native canonical output 必须 0 error/0 warning；
- PptxGenJS fixed-id collision、master hf mismatch 和 invalid cache 作为兼容诊断/intentional difference，不复制；
- LibreOffice direct slide field 必须 open/render/save/reopen 可见；layout/master-only degradation 明确记录；
- PowerPoint control 只有实际 open/render/save/reopen 成功时才报告通过。

## 10. 测试与验收

### Codec 单元测试

- defaults、all supported fields、quantization、input detachment/deep freeze；
- descriptor-safe ordinary/null-prototype 与 accessor/symbol/unknown/sparse/exotic rejection；
- strict direct path、alternate prefixes、PptxGenJS defRPr+rPr precedence、sRGB/scheme/alpha；
- wrong namespace、descendant trap、multiple placeholder/field/paragraph、ordinary field、malformed transform/style；
- create、same-value exact no-op、whole replacement、clear、opaque canonicalization、schema order、id/index allocation；
- master hf enable/disable、missing hf insertion、other hf attributes preservation；
- complete rollback and unrelated byte isolation。

### Model/SDK lifecycle

- zero-input create immediately adds/reads/edits/clears direct slide number；
- firstSlideNumber absent/custom/zero/negative/max/range/repair/clear；
- add/duplicate/move/delete + section combinations keep expected cache and stable models；
- unsupported imported field remains unchanged on unrelated edits；
- all six formats write/open twice with preserved format profile and validator errors `[]`；
- layout/master live properties create/read/edit/clear/copy/relink without cross-owner mutation；
- outer transaction failure after multi-slide cache changes restores all parts/journal。

### PptxGenJS public-output conformance

- omitted、empty、styled sRGB、theme color、left/center/right、uniform/tuple margin、top/middle/bottom；
- `defineSlideMaster({ slideNumber })` + named-layout slide；
- public output imported without private-field access；
- compare geometry、alignment、margin、font/size/color/bold、field type、cached slide text and owner count；
- lock intentional corrections：fixed id collision、master hf=0、`null/1000` cache、zero width/height fallback、justify fallback、
  ignored italic/transparency/RTL and hidden ancestor mutation；
- native edit/reopen retains neighbors and validates 0/0。

### Packed/browser/client evidence

- actual npm tarball Node ESM/CJS surface、generated declarations、installed CLI 和 smoke flag `slideNumbers: true`；
- real Chrome create/writeBlob/open 覆盖 defaults、style、firstSlideNumber、move/cache，0 diagnostics/console/page/network errors；
- clean-build SHA-256 manifest deterministic；
- native gallery 覆盖 default/custom start、four alignments、three vertical alignments、margins、sRGB/theme/transparency、
  bold/italic/font/lang、duplicate/move/delete final state；
- independent PptxGenJS control；
- CLI inspect/validate、PowerPoint 2010 profile、逐页 render、overflow 与 visual inspection；
- LibreOffice save/reopen 检查 slide count/order、field type/cache/style/direct visibility；
- PowerPoint automation/control 结果按事实记录，不把自动化失败写成 round-trip pass。

## 11. 文档与完成定义

更新根 README、package README、API README、compatibility baseline、implementation progress 和 CHANGELOG，明确：

- slide/layout/master 的 explicit direct ownership；
- public units、defaults、style、firstSlideNumber 和 cache lifecycle；
- PptxGenJS 4.0.1 对等范围与 intentional corrections；
- direct slide 的 LibreOffice portability 与 inherited-only degradation；
- 下一剩余项为 default color，之后进入 master/layout/placeholder。

小项完成必须同时满足：全部新增/现有 tests、performance、typecheck、build、packed/browser/CLI gates 通过；native 与
PptxGenJS gallery 完成结构、validation、render 和客户端检查；review 未发现 correctness/lifecycle/public API 问题；
文档与路线同步；工作树只保留既有未跟踪缓存目录；每个独立小项均已 commit、push，最终远端 divergence 为 `0 0`。
