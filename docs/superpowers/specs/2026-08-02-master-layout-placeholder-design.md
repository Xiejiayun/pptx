# Master、Layout 与 Placeholder 设计

日期：2026-08-02

## 1. 目标

完成一个可独立验收的 master/layout/placeholder 专项，使库既能从零定义可复用的命名页面布局，也能在已有
PPTX 中读取、选择和编辑 master/layout 内容，并覆盖 PptxGenJS 4.0.1 的 `defineSlideMaster()`、
`addSlide({ masterName })` 与 placeholder 公开语义。

完成后必须满足：

- `defineSlideMaster()` 可创建命名 layout，支持 background、margin、slide number，以及 chart/image/line/rect/
  text/placeholder 六类公开对象；
- `addSlide({ masterName })` 按 exact layout name 选择继承链，并与 `sectionTitle` 正交组合；
- master/layout 的 background 和内容可语义读取、创建与编辑，placeholder 可稳定识别、物化和填充；
- layout 普通对象保持在继承层，不复制到每个 slide；placeholder owner 在 slide 创建时立即物化；
- duplicate/move/delete/relink、layout replace/delete、relationship 和 owned-part GC 保持事务安全；
- `pptx/pptm/potx/potm/ppsx/ppsm`、Node、browser、types、CLI、PptxGenJS public output 与真实客户端证据齐全；
- table/master/layout/placeholder 继承完成后，下一专项进入 advanced text。

## 2. 当前状态

### 2.1 已有能力

- `MasterLayoutThemeCodec` 已列出并缓存稳定的 raw `MasterModel`、`LayoutModel`、`ThemeModel`；
- raw theme/master/layout 可 create/copy/delete/relink，slide 可 relink 到 layout；
- `materializeInheritedStyle()` 可按 placeholder `type + idx` 返回 slide/layout/master 三层 XML；
- master/layout placeholder 可读取 `shapeId/type/index/text`；
- direct layout/master slide number 已有 strict read/edit/clear；
- slide、shape、text、image、chart、table、media 的 native creator、relationship 管理、transaction 与 GC 可复用；
- zero-input presentation 已有 canonical master/layout/theme chain，默认 layout 名为 `DEFAULT`。

### 2.2 缺口

- `AddSlideOptions` 只有 `sectionTitle`，没有 named layout 选择；
- raw `createLayout()` 只接受调用者提供的 XML，没有 declarative definition；
- master/layout 没有统一的 semantic content model，background codec 目前只认 `p:sld` root；
- placeholder 只读 snapshot 不表达 name、stable shape identity、slide materialization 或 object population；
- layout image/chart 创建需要 async preparation，现有 creator 把 async 与 mutation 封装在单个 slide 方法内，不能直接组成
  一次全定义 atomic commit；
- PptxGenJS `margin` 是只存在于运行时 layout definition 的 auto-page page margin，OOXML 没有对应字段。

## 3. PptxGenJS 4.0.1 public behavior audit

审计只使用 public constructor、`defineSlideMaster()`、`addSlide({ masterName })`、各公开 object creator 与
`write()` bytes，不读取 `_slideLayouts`、`_slideObjects` 等私有字段。真实行为如下：

- `defineSlideMaster({ title })` 实际新增 `slideLayoutN.xml`，所有命名 layout 共享 `slideMaster1.xml` 和 theme；
- layout `p:cSld@name` 等于 `title`，slide relationship 指向按 exact name 找到的 layout；
- unknown `masterName` 静默回退到默认 layout，重复 title 选择第一个；
- background 和普通 rect/line/text/image/chart 对象只写在 layout，不复制到 slide；
- placeholder 在 layout 中是 text shape，`name` 只用于运行时匹配，`idx` 为 `100 + objects` 中的 ordinal；
- write 时为每张使用该 layout 的 slide 补一个空 placeholder owner；已由 object creator 填充的同名 placeholder 不重复；
- text/image/chart/shape/table/media 的 `placeholder` 选项继承 layout placeholder geometry，并写匹配的 `p:ph`；
- top-level `margin` 不序列化，只被 table auto-page/`tableToSlides` 当作 page margin 使用；
- layout slide number 还会被复制到 master 和每张 slide，固定 shape id 25、随机/noncanonical cache 和 disabled master
  等缺陷已在 Slide Number 专项锁定。

Native 只对合法 public final semantics 对等，不复制 silent fallback、first-duplicate selection、JSON clone/coercion、
write-time delayed mutation、fixed IDs、disabled master 或 noncanonical cache。

## 4. 方案比较

### A. 每个 `defineSlideMaster()` 创建真实 slide master

名称与 API 表面一致，但与 PptxGenJS 实际 package graph 不一致；会复制 theme、master text styles 和 layout chain，
还会让 named selection、delete/relink 与多 layout master 的已有 PPTX 互操作变复杂。拒绝。

### B. 在真实 parent master 下创建命名 layout（采用）

把 PptxGenJS 的 “master title” 明确定义为 layout name。普通对象留在 layout；slide 只保存 layout relationship 和
需要编辑的 placeholder owner。已有 raw codec 继续负责 package graph，SDK 增加 semantic wrapper 和 async definition
orchestrator。该方案与真实 public output 一致，也能服务 existing-deck editing。

### C. 把定义对象复制到每张新 slide

实现简单，但失去继承、layout 后续编辑、文件体积、relationship dedup 和标准 PowerPoint master behavior；无法视为
master/layout 支持。拒绝。

## 5. 公共 API

### 5.1 Named slide creation

扩展现有类型：

```ts
export interface AddSlideOptions {
  readonly masterName?: string;
  readonly sectionTitle?: string;
}
```

`masterName` 仍沿用 PptxGenJS public spelling，但文档明确它选择的是 `LayoutModel.name`。未提供时继续使用首张 slide
的 layout，否则使用 presentation 的第一个安全 default layout。提供时必须：

- 是 non-empty、XML-safe string；
- 在 presentation 可达 layouts 中 exact match 唯一；
- relationship 为 internal、target 存在且 content type 正确。

零匹配抛 `RangeError`，多匹配或不安全 graph 抛 `ModelParseError`，都发生在分配 slide part 前。
Named resolver 只接受 unique direct `p:cSld@name`；raw `LayoutModel.name` 为显示而提供的 root-name/
basename fallback 不会被当成可选 master name。

### 5.2 Placeholder types and selectors

```ts
export const PLACEHOLDER_TYPES = [
  'title',
  'body',
  'pic',
  'chart',
  'tbl',
  'media',
] as const;

export type PlaceholderType = typeof PLACEHOLDER_TYPES[number];

export interface PlaceholderIdentity {
  readonly type: PlaceholderType;
  readonly index: number;
}

export type PlaceholderSelector =
  | string
  | PlaceholderIdentity;

export interface AddPlaceholderOptions extends Omit<AddTextOptions, 'name'> {
  readonly name: string;
  readonly type: PlaceholderType;
  readonly index?: number;
}
```

String selector 按 layout placeholder 的 direct `cNvPr@name` exact match；identity selector 按 canonical `type + index`
匹配。两种 selector 都要求唯一结果。`index` 省略时使用 `100 + object ordinal`，显式值必须是
`0..4294967294` safe integer；`4294967295` 保留给既有 slide-number compatibility state，不用于新 placeholder。

`ShapeModel.placeholder` 增加只读 `Readonly<PlaceholderIdentity> | undefined`，使 slide、layout、master 中的
placeholder owner 都可通过 stable shape identity 检查。Name 继续使用现有 live `ShapeModel.name`。

### 5.3 Declarative definition

```ts
export interface SlideMasterMargin {
  readonly top: Emu;
  readonly right: Emu;
  readonly bottom: Emu;
  readonly left: Emu;
}

export type SlideMasterMarginInput =
  | Emu
  | readonly [Emu, Emu, Emu, Emu];

export type SlideMasterBackground =
  | SlideBackground
  | {
      readonly kind: 'image-source';
      readonly source: RasterImageSource;
      readonly contentType?: RasterImageContentType;
      readonly signal?: AbortSignal;
    };

export type SlideMasterObject =
  | { readonly kind: 'rect'; readonly options?: AddShapeOptions }
  | { readonly kind: 'line'; readonly options?: AddShapeOptions }
  | {
      readonly kind: 'text';
      readonly text: string | readonly RichTextParagraph[];
      readonly options?: AddTextOptions;
    }
  | {
      readonly kind: 'placeholder';
      readonly text?: string | readonly RichTextParagraph[];
      readonly options: AddPlaceholderOptions;
    }
  | {
      readonly kind: 'image';
      readonly source: ImageSource;
      readonly options?: AddImageSourceOptions;
    }
  | {
      readonly kind: 'chart';
      readonly groups: readonly ChartGroupInput[];
      readonly options?: AddChartOptions;
    };

export interface DefineSlideMasterOptions {
  readonly title: string;
  readonly master?: SlideMasterModel;
  readonly background?: SlideMasterBackground;
  readonly margin?: SlideMasterMarginInput;
  readonly slideNumber?: SlideNumberOptions;
  readonly objects?: readonly SlideMasterObject[];
}
```

Root method：

```ts
await document.defineSlideMaster(options): Promise<SlideLayoutModel>;
```

省略 `master` 时使用第一个 presentation-attached、internal、content-type-correct master；显式 model 必须属于当前
document 且仍 attached。`title` 在全部 presentation layouts 中必须唯一。

Native 使用 discriminated union，不复制 PptxGenJS 的 single-key object、truthy fallback 或 unknown-key ignore。
`rect` 固定映射 preset `rect`，`line` 固定映射 preset `line`；其他 native preset/custom shape 继续通过 layout semantic
content API 创建，不塞入这个 PptxGenJS-focused definition union。

### 5.4 Semantic master/layout wrappers

保留 `document.masterLayoutTheme` 的 low-level raw codec API。`PptxDocument.masters` 与 `layouts` 改为 SDK stable wrappers：

```ts
export class SlideLayoutModel {
  readonly partUri: string;
  readonly name: string;
  readonly masterPartUri?: string;
  readonly shapes: readonly SemanticShape[];
  readonly placeholders: readonly ShapeModel[];
  background: SlideBackground | undefined;
  margin: Readonly<SlideMasterMargin> | undefined;
  slideNumber: Readonly<SlideNumber> | undefined;

  addPlaceholder(
    value: string | readonly RichTextParagraph[],
    options: AddPlaceholderOptions,
  ): ShapeModel;
}

export class SlideMasterModel {
  readonly partUri: string;
  readonly layouts: readonly SlideLayoutModel[];
  readonly shapes: readonly SemanticShape[];
  readonly placeholders: readonly ShapeModel[];
  readonly theme?: ThemeModel;
  background: SlideBackground | undefined;
  slideNumber: Readonly<SlideNumber> | undefined;

  addPlaceholder(
    value: string | readonly RichTextParagraph[],
    options: AddPlaceholderOptions,
  ): ShapeModel;
}
```

两个 wrapper 还转发与 `SlideModel` 完全相同签名和返回类型的 `addText()`、`addRichText()`、`addShape()`、
`addImage()` 与 `addChart()`；implementation 不另建一套宽松 options。

Wrapper 复用 package part URI 作为 identity key；collection 重读、无关编辑和 move/relink 保持 `===`。Low-level raw
`LayoutModel` / `MasterModel` 继续从 `@pptx/codecs` 与 `document.masterLayoutTheme` 可用，避免把 semantic model 反向依赖
塞进 codec package。

Wrapper 的 common-slide content 只暴露适用于 owner 的内容操作，不暴露 slide-only `hidden`、notes、sections、
transient color 或 direct slide number cache。

### 5.5 Definition editing and deletion

```ts
await document.replaceSlideMaster(
  layout: SlideLayoutModel,
  options: DefineSlideMasterOptions,
): Promise<void>;

document.deleteSlideMaster(
  layout: SlideLayoutModel,
  replacement?: SlideLayoutModel,
): void;
```

`replaceSlideMaster()` 保持 layout part URI、wrapper identity 和所有 incoming slide relationships，只 whole-replace
owned background/content/relationships/margin/slideNumber。新的 `master` 可触发 parent relink；新的 `title` 仍必须唯一。
等值 canonical definition 是 exact parts/relationships/graph/journal no-op。

Delete 在存在 incoming slides 时要求 replacement；retarget 后再删除 layout relationship、owned parts 和 transient
margin。Replacement 必须属于同一 document 且不是 target。删除 default layout 后，无参数 `addSlide()` 使用新的第一个安全
layout；若 presentation 已无安全 layout，明确失败而不创建 dangling slide。

## 6. Architecture

### 6.1 Low-level graph remains in codecs

`MasterLayoutThemeCodec` 继续负责 raw part create/copy/delete/relink、master/layout/theme relationships、ID list 和
stable raw model。新增 strict named-layout resolver、safe attached checks 和 relationship validation，但不依赖 `@pptx/model`。

### 6.2 SDK wrapper composes model content

SDK 为每个 master/layout part URI 缓存 wrapper，并为其创建一个 internal common-slide content facade。Facade 复用现有
shape parser/renderer/relationship logic，但只暴露合法 owner 操作。不会把 `SlideModel` 的 slide ID、section、notes、hidden
或 default color 伪装成 layout/master state。

为避免一次性重写 `SlideModel`，第一版使用小型 `CommonSlideContentModel` adapter，把已有通用 shape/text/image/chart
实现提取为 owner-neutral helpers；slide 与 wrapper 调用同一 helper。只提取本专项需要的函数，不重构无关 table/media
代码。

### 6.3 Generic background owner

现有 background internal codec 增加 owner kind：

```ts
type BackgroundOwnerKind = 'slide' | 'layout' | 'master';
```

它分别要求唯一 namespace-correct root `p:sld`、`p:sldLayout`、`p:sldMaster`，并只拥有 direct
`p:cSld/p:bg`。Fill、image relationship、duplicate/shared target COW 和 reference-safe GC 规则完全复用；不把 `p:bgRef`
误读为 supported direct background。

### 6.4 Prepare then commit

`defineSlideMaster()` / `replaceSlideMaster()` 分两阶段：

1. descriptor-safe normalization、source loading、signature/MIME 检查、SVG fallback、chart workbook build 和所有 async
   工作在 package mutation 前完成；
2. 一个 synchronous outer OPC transaction 创建或更新 layout，插入已准备对象，更新 relationships/ID lists/margin map。

现有 chart/image creator 拆出 internal prepare plan 与 synchronous commit helper；public slide creator 继续调用同一条路径，
保证输出和行为不分叉。任一 prepare 或 commit 失败都保持 parts、relationships、ZIP、graph、model caches、transient margin
与 mutation journal 不变。

## 7. Placeholder semantics

### 7.1 Layout creation

Layout placeholder 是 canonical `p:sp`，包含唯一 positive shape ID、direct `cNvPr@name`、`p:ph@type/idx`、合法
transform、plain/rich prompt 和支持的 text/body style。Name 在同一 layout 内唯一；`type + index` 也唯一。

Prompt 只留在 layout。它可由 PowerPoint 显示为编辑提示，不当作新 slide 的实际内容。

### 7.2 Slide materialization

`addSlide({ masterName })` 在同一 create transaction 内扫描 selected layout 的 strict supported placeholders，并为每个
创建 canonical empty owner：

- 保持 name、type、index、geometry 和合法 body properties；
- 不复制 prompt run；
- 不复制 layout-only普通对象；
- 复制过程中 shape IDs 在新 slide 内唯一；
- unsupported/ambiguous layout placeholder 使显式 named creation 在 mutation 前失败。

这与 PptxGenJS write-time final output 语义一致，但 native 在 `addSlide()` 返回前完成，避免 write 改变模型。

### 7.3 Population

以下 options 增加 `placeholder?: PlaceholderSelector`：

- `AddTextOptions` / rich text；
- `AddShapeOptions`；
- `AddImageOptions` / high-level image source options；
- `AddChartOptions`；
- `AddTableOptions`；
- `AddMediaOptions`。

指定 selector 时，creator 必须解析当前 slide 的 unique internal layout relationship，再匹配 layout 与 slide owner。
Geometry 取 layout placeholder；显式 object transform 不覆盖 placeholder geometry，避免同一个 placeholder 出现两套坐标。
内容/style 仍使用 object-specific options。创建操作在原 tree position 用同 shape ID 替换 empty owner，并写相同 `p:ph`：

- text/rich text 保持 `p:sp`；
- image/media 使用 `p:pic`；
- chart/table 使用 `p:graphicFrame`；
- rect/line 使用 `p:sp`。

旧 empty `ShapeModel` handle 因 owner kind 改变时按现有 stale-handle 规则失效，新 creator 返回 live replacement model。
重复填充已经非空的 owner、unknown selector、type 与 object domain 明显冲突、missing/ambiguous relationship 都在 mutation
前失败。`title/body` 可填 text/shape，`pic` 填 image，`chart` 填 chart，`tbl` 填 table，`media` 填 audio/video；native
不复制 PptxGenJS 对不匹配 domain 的 loose passthrough。

### 7.4 Inherited style inspection

`materializeInheritedStyle()` 保留 raw XML escape hatch。Semantic wrapper 另提供 placeholder owner 的 direct state；本专项
不计算完整 theme cascade，也不把 master/layout prompt 当作 slide text。高级多级 text-style cascade 留在 advanced text，
但 layout/master direct text/body properties必须无损保留并在 placeholder population 时生效。

## 8. Slide number and margin integration

- Definition `slideNumber` 写 direct layout field；selected slide 创建时同步 materialize direct slide field 和当前 canonical
  cache，使 PowerPoint/LibreOffice 可见性不依赖客户端继承实现；
- parent master 的 direct `p:hf@sldNum` 被安全启用，但不强行复制某个 layout 的 style 到共享 master；
- 不使用固定 shape id 25，不写 random cache，不生成 disabled master；
- `margin` 归一化为 frozen `{ top,right,bottom,left }` EMU，保存在 document 的 transient layout map；
- scalar margin 同时用于四边，tuple 顺序固定为 top/right/bottom/left；每边必须是 non-negative safe
  integer，水平与垂直 pair sum 必须分别小于 slide width 与 height；
- margin 不序列化，reopen 后 `layout.margin === undefined`，与 OOXML/PptxGenJS 的非持久化事实一致；
- replace 更新、delete 清理；后续 `tableToSlides` 直接消费该 map。改变 margin 不重排已有 slide 或 table。

## 9. Lifecycle and isolation

- slide duplicate 保留原 layout relationship，并复制 slide placeholder owners/filled objects；
- slide move 不改 layout 或 placeholder identity；
- slide delete 只 GC slide-owned dependencies，layout-owned image/chart targets仍有 incoming 时保留；
- layout ordinary objects和其 media/chart/workbook targets由 layout relationship拥有，多张 slide 不复制这些 targets；
- replace 采用 relationship-aware reuse/COW，删除 superseded relationship 前检查 XML reference count，删除 target 前检查
  package graph incoming；
- raw master/layout copy/delete/relink 后 SDK wrapper collection 与 stable identity同步；
- model cache 不让已删除 part URI 的旧 wrapper污染以后重新分配的 URI；旧 handle 对缺失/新 generation part 抛 stale error。

## 10. Validation and intentional differences

Strict native behavior：

- options 只接受 ordinary/null-prototype own-data object，拒绝 accessor、symbol、unknown field、class instance 和 malformed
  nested input；明确允许的 `SlideMasterModel` wrapper instance 单独验证归属；
- duplicate/empty/unsafe title、placeholder name、identity 和 unsupported object kind 在 mutation 前拒绝；
- unknown/duplicate `masterName` 不 silent fallback；
- percentage coordinates 不在本专项引入，所有 native geometry 继续使用 EMU；
- PptxGenJS invalid placeholder type、fixed IDs、truthy zero loss、random cache、disabled master、JSON coercion 和 write-time
  mutation 只作为 conformance difference 锁定，不复制；
- external/dangling/wrong-content-type layout/master/image/chart relationship 保持 bytes，危险 semantic edit 拒绝。

Compatibility diagnostics 新增：

- `LAYOUT_NAME_DUPLICATE`；
- `LAYOUT_RELATIONSHIP_INVALID`；
- `PLACEHOLDER_IDENTITY_AMBIGUOUS`；
- `PLACEHOLDER_OWNER_MISSING`；
- `PLACEHOLDER_DOMAIN_MISMATCH`；
- `LAYOUT_MARGIN_TRANSIENT` 只作为 info，写出不报 warning。

## 11. Testing and evidence

### 11.1 Unit and integration

- strict normalization、frozen/detached values、equal no-op、invalid zero mutation；
- named selection、unknown/duplicate、section combination、default fallback；
- master/layout wrapper stable identity、background、content create/edit、raw codec coexistence；
- placeholder six types、explicit/allocated index、prompt、empty materialization、all six object-domain population；
- duplicate/move/delete/relink、replace/delete with replacement、target GC、rollback、URI reuse；
- margin transient lifecycle与 reopen boundary；
- all six presentation formats、五个 compatibility profiles、alternate prefixes/wrong namespaces/opaque preservation。

### 11.2 PptxGenJS conformance

Public-output cases覆盖：

- default and multiple named layouts；
- solid/transparency/image background；
- scalar/TRBL margin；
- slide number；
- rect、line、plain/rich text、PNG/SVG image、standard chart；
- title/body/pic/chart/tbl/media placeholders；
- populated and empty placeholders；
- unknown/duplicate/invalid/zero/falsy intentional differences。

比较最终 layout/slide/master relationship、shape order、type/index、transform、text/style、payload、chart/workbook、background
和 visible semantics，不比较随机 GUID、relationship ID 或无语义 XML formatting。

### 11.3 Package/runtime/client

- focused tests、全量 Vitest、performance、typecheck、全仓 build；
- actual npm tarball Node、real Chrome、browser conditional export、TypeScript consumer、installed CLI；
- 两次 clean package build 的 dist manifest 与 tarball SHA-256；
- native 与 PptxGenJS galleries 的 slide/part/relationship/owned-target/orphan counts；
- `pptx-inspect` inspect/validate/slides/part-read/diff，PowerPoint 2010 profile；
- 180-DPI 全页 render、overflow、edge/minimum-margin 和逐页视觉检查；
- LibreOffice 26.8 与本机 PowerPoint 16.112 只记录实际产生的 output，不把自动化失败写成客户端通过。

## 12. 小项拆分与提交门禁

1. Named layout selection：strict `masterName` resolution、section/default behavior、lifecycle与六格式。
2. Semantic wrappers：stable `SlideMasterModel`/`SlideLayoutModel`、owner-neutral content与 raw codec coexistence。
3. Layout/master backgrounds：generic owner read/edit/image lifecycle与 PptxGenJS layout background import。
4. Placeholder identity/materialization：types、name/index、empty slide owners、read/reopen。
5. Placeholder population：text/shape/image/chart/table/media replacement、domain validation与 rollback。
6. Declarative synchronous definition：title、margin、slide number、rect/line/text/placeholder。
7. Async definition objects：image/chart/background source prepare-then-commit、owned targets与 atomicity。
8. Definition replace/delete：stable identity、retarget、GC、transient cleanup与 no-op。
9. PptxGenJS audit/conformance：valid public output与 intentional differences。
10. Packed/runtime/client verification：tarball、browser、types、CLI、gallery、validator、LibreOffice/PowerPoint。
11. Public docs/final closure：README/API/baseline/progress/changelog，下一专项切到 advanced text。

每项先写 failing test，最小实现后运行 focused tests、typecheck/相关 build 和 diff review；review 无问题立即独立 commit、
push，并同步完成项、剩余项与专项进度。临时 deck、render、tarball、client round-trip 和 `.pnpm-store/` 不提交。

## 13. Success criteria

- 从零可定义至少一个含 background、margin、slide number、六类 layout object 和六类 placeholder 的命名 layout；
- `addSlide({ masterName })` 立即建立正确 inheritance chain 和 empty placeholder owners；
- supported object 可填充对应 placeholder，geometry/type/index/name/reopen保持；
- existing master/layout 可读取并通过 wrapper 做受支持的 content/background/placeholder edit；
- replace/delete/relink、slide lifecycle、owned dependencies、rollback、no-op 和 URI reuse 无泄漏；
- valid PptxGenJS 4.0.1 public output 在适用 final semantics 上对等，差异有永久测试；
- 六格式、全量 tests、performance、typecheck、build、actual package、browser、CLI、gallery、validator 和客户端证据完成；
- 文档把 master/layout/placeholder 从下一项移入已支持，并把 advanced text 设为下一专项。

## 14. Self-review

- Placeholder scan：未发现占位标记、未定义接口或留给实施者自行决定的行为。
- Internal consistency：PptxGenJS “master” 始终映射 named layout；ordinary objects留在 layout，只有 placeholder owner与
  direct slide-number visibility state物化到 slide。
- Scope：不实现 full theme text-style cascade、percentage coordinate、advanced text/table/media/chart style 或
  `tableToSlides`；只提供它们需要的 placeholder placement和 transient margin基础。
- Ambiguity：name/identity resolution、unknown/duplicate behavior、owner domain、margin persistence、async atomicity、
  replace/delete和客户端门禁均有唯一规则。
