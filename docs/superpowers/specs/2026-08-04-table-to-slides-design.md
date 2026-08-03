# HTML Table to Slides Design

日期：2026-08-04
状态：已批准实施

## 目标与范围

为 native `PptxDocument` 增加 `tableToSlides()`，从当前浏览器文档中的 `<table>` 读取 `thead`、全部 `tbody`、`tfoot`、cell text、CSS cell style、colspan/rowspan、可见列宽和 `data-pptx-min-width` / `data-pptx-width`，生成一张或多张普通、可编辑、可重开的 PowerPoint table slides。

本项复用已经完成的 strict table creation、automatic content measurement、rich row fragmentation、repeated headers、same-layout continuation insertion、section/slide-number synchronization、relationship ownership和 transaction rollback。它同时覆盖 PptxGenJS 4.0.1 `TableToSlidesProps` 的 master layout、slide margins、auto-page weights、first/continuation placement，以及每张生成页附加 image、shape、table、text 的能力。

本项不实现通用 HTML/CSS renderer，不保留 DOM 节点或 CSS rule引用，不导入 word-level HTML markup，不嵌套 PowerPoint table，也不把网页截图冒充 editable table。Cell `innerText`是唯一 text source；nested table只贡献其可见文字。完成本项后，PptxGenJS公开能力缺口归零，但在最终 peer/client full-surface audit 完成前仍不宣称完整 parity 认证。

用户已明确授权实现方自主选择最佳方案、持续推进且不设置询问点，因此本设计完成方案比较、自审和真实 runtime 基线后直接批准。

## PptxGenJS 4.0.1 权威基线

基线来自已安装的 PptxGenJS 4.0.1 declarations、`dist/pptxgen.cjs.js`、官方 HTML-to-PowerPoint 文档和真实 Google Chrome 探针。

公开声明是：

```ts
tableToSlides(eleId: string, props?: PptxGenJS.TableToSlidesProps): void;

interface TableToSlidesProps extends TableProps {
  addImage?: { image: DataOrPathProps; options: PositionProps };
  addShape?: { shapeName: SHAPE_NAME; options: ShapeProps };
  addTable?: { rows: TableRow[]; options: TableProps };
  addText?: { text: TextProps[]; options: TextPropsOptions };
  autoPage?: boolean;
  autoPageCharWeight?: number;
  autoPageLineWeight?: number;
  autoPageRepeatHeader?: boolean;
  autoPageSlideStartY?: number;
  colW?: number | number[];
  masterSlideName?: string;
  slideMargin?: Margin;
}
```

4.0.1 runtime按 HTML section顺序读取 rows，使用 first-row `offsetWidth`比例计算列宽，从 computed style读取 color/background/font/padding/border/alignment，调用现有 table paginator，并在每张生成页上按 image → shape → table → text顺序添加附加对象。官方 `data-pptx-min-width`和`data-pptx-width`单位是英寸。

真实 Chrome 基线确认：

- 返回值是 `undefined`；
- 传入 options、nested addText options和 parsed cells会被 runtime写入；
- `autoPage: false`仍生成与 true相同的 10 张分页 slide，属于实现缺陷；
- `x: 0.75`在第一张后被写成 top margin `0.4`；
- 900px、25%/75%的两列在 `w: 12`和左右 margin 0.5/0.7下变成 `[2.7, 8.1]`英寸；
- 18px Arial、600 weight、RGB text/fill、2px四边 border、7/11/7/11px padding、right/bottom alignment均被复制到 cell；
- `innerText`会在分页器中被拆成 mutable word/line arrays。

Native要求合法常用输入得到同等 editable output能力，但不复制 caller mutation、`autoPage: false`失效、truthy/coercible controls、selector escaping缺陷、colspan width filter错误、CSS decimal stripping、silent NaN或 mutable line queues。

## 公共 API

API位于 `@pptx/sdk`的 `PptxDocument`，并通过 root `@jiayunxie/pptx`导出：

```ts
export interface TableToSlidesAddImage {
  readonly source: ImageSource;
  readonly options?: AddImageSourceOptions;
}

export interface TableToSlidesAddShape {
  readonly type: PresetShapeType;
  readonly options?: AddShapeOptions;
}

export interface TableToSlidesAddTable {
  readonly rows: readonly (readonly AddTableCellInput[])[];
  readonly options?: AddTableOptions;
}

export interface TableToSlidesAddText {
  readonly text: string | readonly RichTextParagraph[];
  readonly options?: AddTextOptions;
}

export interface TableToSlidesOptions {
  readonly name?: string;
  readonly masterSlideName?: string;
  readonly autoPage?: boolean;
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
  readonly autoPageRepeatHeader?: boolean;
  readonly autoPageHeaderRows?: number;
  readonly autoPageSlideStartY?: number;
  readonly slideMargin?: TableAutoPageMarginInput;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly addImage?: TableToSlidesAddImage;
  readonly addShape?: TableToSlidesAddShape;
  readonly addTable?: TableToSlidesAddTable;
  readonly addText?: TableToSlidesAddText;
}

export class PptxDocument {
  tableToSlides(
    elementId: string,
    options?: TableToSlidesOptions,
  ): Promise<readonly SlideModel[]>;
}
```

所有 native geometry字段继续使用 EMU；CSS pixel values和 HTML `data-pptx-*` inches只在 DOM import boundary转换。`autoPage`默认 true；显式 false创建一张 ordinary automatic-height table slide，不进行分页。返回值是 frozen readonly snapshot，包含 HTML table直接生成的全部 pages；这比 PptxGenJS void result更可验证。由于附加 image支持 path、URL、data URI、bytes、Blob和streams，方法始终 async。

`height`表示每页 table layout region的最大高度，不是 fixed table row-height总和。其 effective bottom edge为 `min(y + height, slideHeight - bottomMargin)`；分页时通过等价的 stricter bottom margin传给现有 planner。省略时使用 layout/explicit/canonical bottom margin。`width`是 table实际宽度；省略时为 `slideWidth - x - rightMargin`。

Options和所有 nested descriptors必须是 ordinary或 null-prototype data objects/arrays。Accessor、symbol、inherited、class、sparse、unknown property、numeric coercion和 invalid finite/range input在 mutation前拒绝。DOM/CSS platform objects自身通过标准读取API访问，不被当作 caller option records。

## 方案选择

1. **纯 DOM/CSS snapshot parser + 现有 native table auto-page + 单事务提交；采用。** Parser只产生 detached rows、style和width constraints；已经验证的 table engine继续负责 measurement、fragmentation、header、layout、section、relationship和 OOXML。DOM读取、外部 image resolution和 package mutation边界清晰。
2. **逐行移植 PptxGenJS `genTableToSlides()` / `getSlidesForTableRows()`。** 可快速复刻简单输出，但会复制 options mutation、word queue、false-autoPage缺陷、colspan filter错误、rowspan line-height零化、CSS decimal bug和第二套 pagination engine，无法维持现有 strict transaction/relationship contract，故不采用。
3. **把 HTML table rasterize为图片。** CSS视觉覆盖更广，但丢失 editable cells、text、merge、accessibility和 table semantic state，不满足“创建完整PPT并可编辑”的最终目标，故不采用。

## 模块边界

新增 `packages/sdk/src/table-to-slides.ts`，只负责：

- descriptor-safe `TableToSlidesOptions` normalization；
- 从 real或fixture DOM-like table读取 detached `HtmlTableSnapshot`；
- computed CSS到 native `AddTableCellOptions`映射；
- visible pixel widths、HTML width constraints和 target EMU width求解；
- 生成交给 `SlideModel.addTable()`的 immutable rows/options；
- 预解析附加 image source，供每个生成页复用。

`packages/sdk/src/index.ts`负责 `PptxDocument.tableToSlides()` orchestration：解析当前 global document中的 table、解析 master/layout margin、在 mutation前 resolve image、开启一个 outer OPC transaction、append first slide、调用 source `addTable()`、获取 `newAutoPagedSlides`、按固定顺序添加附加对象、返回 frozen slides。

`packages/model`不增加 DOM或 CSS依赖。Existing `SlideModel.addTable()`、`PresentationModel.addSlide()`、layout margin map、same-layout insertion和 addImage/addShape/addTable/addText/addRichText APIs保持唯一 OOXML mutation owners。

## DOM Source 与 Row Ordering

`elementId`必须是 non-empty string。Browser document、`getElementById()`、owner document default view和 `getComputedStyle()`必须存在；target必须是 `<table>`。Native不拼 CSS selector，因此包含引号、冒号、空格或 CSS metacharacter的合法 ID仍可工作。

Row顺序固定为：

1. `table.tHead?.rows`；
2. 每个 `table.tBodies[i].rows`，保持 DOM顺序；
3. `table.tFoot?.rows`。

至少需要一行且每行至少一个 cell。Cell text读取一次 `innerText`并立即 normalize CRLF/CR为 LF；之后 DOM变化不影响 operation。`thead` row count单独记录，供 repeated-header default使用。`tfoot`只位于最后，不重复。多个 `tbody`完整保留。Nested table不递归创建 table，只保留 outer cell `innerText`。

HTML `colSpan` / `rowSpan`读取 normalized DOM property；值1省略，大于1映射到 cell options。最终 grid完整性、overlap、physical cell expansion和 merge topology仍由 existing strict table normalizer验证。

## CSS Cell Mapping

每个 cell只调用一次 `getComputedStyle()`，缓存所需属性：

- `color` → `{ kind: 'srgb', value: RRGGBB }`；
- transparent background → canonical white；其他 `background-color` → solid sRGB fill；
- first computed font family → `fontFamily`，去除外围引号；generic/inherit/initial/empty值不写；
- positive computed `font-size` px数值按 PptxGenJS兼容边界作为 point值；
- `font-weight: bold|bolder`或 numeric `>= 500` → `bold: true`，否则 false；
- `text-align: left|center|right|justify|start|end` → native alignment；start/end根据 computed direction解析；
- `vertical-align: top|middle|bottom` → native vertical alignment；
- 四边 `padding-*` px按 1px = 1pt映射，保留到 OOXML point量化；
- 四边 border的 style/width/color分别映射；none/hidden为 no-border，solid为 solid，dash/dotted/double等可见非实线降级到 native `dash`，width以 point量化；
- rgba alpha除完全透明背景外与 PptxGenJS一样不改变最终 sRGB alpha，因为当前 table direct fill/text API不从 CSS公开 opacity。

CSSOM通常返回 `rgb()` / `rgba()`；parser也接受空格和小数通道并 round到 0–255。Unsupported CSS color spaces、currentColor ambiguity、non-finite size或 invalid platform return必须明确报错，不生成 `NaN` OOXML。

Word-level HTML style、pseudo elements、background image/gradient、border radius、text decoration、writing-mode、box shadow、CSS variables原始tokens和 layout outside cells不导入。它们不是 PptxGenJS 4.0.1 tableToSlides声明的可编辑 cell state。

## Column Width Resolution

若 options显式提供 `columnWidths`，它拥有最高优先级；existing strict table normalization验证 scalar/vector、column count、positive safe EMU和 `width` sum。

否则选择第一条 non-empty row作为 width source，按每个 cell `offsetWidth / colSpan`展开为 physical-column pixel weights。总 pixel width必须 positive；hidden/zero-layout table明确拒绝并提示使用 `columnWidths`或显示 table，不继续传播 NaN。

Target table width依次为：显式 `width`，否则 `slideWidth - effectiveX - rightMargin`。先按 pixel比例用 integer largest-remainder分配，使所有列 positive且 sum exact。

第一条 header row的 `<th>`可以提供官方 attributes：

- `data-pptx-width="N"`：该 header cell覆盖范围的 fixed total width，N为 positive finite inches；
- `data-pptx-min-width="N"`：该范围的 minimum total width，N为 non-negative finite inches；
- 同时存在时 fixed优先；跨 colspan的 total constraint按 quotient/remainder分配到 covered physical columns。

Fixed columns保持 exact，minimum columns通过 iterative water-filling得到不小于 minimum的 proportional width。若 fixed+minimum已超过 target，actual table width扩大到约束总和；否则 flexible columns填满 target。Final columnWidths全为 positive safe EMU，table width始终等于 exact sum。Native修正4.0.1把 fixed attr误作 minimum、只查 `nth-child`和约束后 sum漂移的缺陷。

## Layout、Pagination 与 Header

Effective margins优先级与 existing auto-page一致：explicit `slideMargin` → selected runtime layout margin → canonical 0.5 inch。`masterSlideName`交给 `addSlide({ masterName })`并在 mutation前确认唯一 live layout。未提供 master时沿用现有 default/inherited layout规则。

First Y默认为 top margin，continuation Y默认为 existing planner的 layout top margin；显式 `y`和 `autoPageSlideStartY`分别覆盖。X默认为 left margin。

`autoPage: true`时，HTML rows以 omitted rowHeights进入 deterministic measurement。Table/cell font size、margin、colspan和 text参与现有 estimator。`autoPageCharWeight` / `autoPageLineWeight`继续 strict `[-1, 1]`。Repeated header规则：

- false时不重复，且显式 `autoPageHeaderRows`非法；
- true且显式 header count时使用该 count；
- true且省略时，有 `thead`则重复全部 thead rows，否则重复第一行；
- header count不能超过 total rows，merge不能跨 header/body boundary。

Generated slides是 ordinary same-layout pages，保持 section membership、slide-number cache和 editable table state。每页 table transform height是 measured rows exact sum。内部 link与 page-local relationship behavior完全继承 existing table engine。

`autoPage: false`时创建一张 slide和一张 automatic-row table；不设置 auto-page controls、不创建 continuation、不声称 overflow protection。该语义遵循 public option文档并修正4.0.1 runtime忽略 false的缺陷。

## Additional Objects

HTML table先创建，随后每张 HTML-generated page按下列顺序添加最多一个对象：

1. `addImage`；
2. `addShape`；
3. `addTable`；
4. `addText` / `addRichText`。

Image source在任何 package mutation前只 resolve一次；PNG/JPEG/GIF/SVG signature、fallback、content type和 sizing使用现有 SDK规则。每页创建自己的 picture relationship，同时 content-aware media dedup共享相同 payload。Placeholder-aware sizing在每个 page local owner上重新计算。

Shape、table和text使用现有 native types与 validators。Rich text调用 `addRichText()`。附加 table可启用自己的 auto-page；其 continuation pages只属于该附加 table，不加入 HTML-generated result snapshot，也不自动获得其他附加对象，与 PptxGenJS逐页 addTable调用边界一致。

附加对象的 internal slide hyperlink在所有 HTML pages创建后解析，因此 slide identity稳定。所有 templates保持 detached且不被写入 default、relationship ID、line queue或 rendered metadata。

## Transaction 与 Failure Isolation

执行顺序固定：

1. normalize element ID/options和 nested descriptors；
2. resolve DOM table、snapshot rows/styles/width constraints；
3. resolve master/layout margin、placement、column widths和 header count；
4. resolve/inspect optional image及 SVG fallback；
5. preflight immutable table rows/options；
6. 在一个 outer `OpcPackage.transaction()`中添加 first slide、HTML table pages和附加对象；
7. 成功后冻结并返回 HTML page models。

步骤1–5不得改变 package、diagnostics、mutation journal、layout maps、slide/shape caches或 DOM/options。步骤6任意 source/generated/additional-object/relationship/section/slide-number failure恢复 exact package bytes/relationships/order/sections/journal。失败期间创建后又 detached的 `SlideModel`从 presentation cache清除；retry必须得到与 clean run byte-identical的结果。

Async image abort/load/signature/fallback错误发生在 mutation前。CSS getter或 DOM platform错误被包装为含 element/cell位置的 deterministic error。Native不 console.warn后继续，也不留下部分 slides。

## PptxGenJS 对等边界

共同能力要求：

- string ID选择 HTML table；
- thead/tbody/tfoot、plain visible text、colspan/rowspan；
- visible proportional columns、fixed/min width attributes；
- cell text/fill/font/padding/border/align/valign；
- automatic multi-slide pagination、weights、repeated headers、first/continuation placement和 slide margins；
- named master/layout；
- 每页附加 image、shape、table和text；
- editable OOXML tables和 ordinary generated slides。

Deliberate native differences：

- EMU public geometry与 async Promise result；
- strict data records/ranges和 pre-mutation rejection；
- `autoPage: false`真正禁用分页；
- caller options、DOM和 additional templates不被修改；
- exact column sum、correct colspan range、fixed vs minimum语义；
- decimal CSS解析、start/end direction和 justify支持；
- no selector interpolation、no truthy/coercion、no silent warning/NaN；
- deterministic existing estimator而非 mutable PptxGenJS word queue。

这些差异提高正确性、transaction safety和跨 Node/browser可验证性，不减少合法 public use case的最终能力。

## 测试与验收

1. **Option normalization：** omitted/default/false、bounds、master/margin/placement/height、column widths、四类 add descriptors、accessor/symbol/inherited/class/sparse/unknown/caller detachment。
2. **DOM selection：** missing browser/document/view/style function、empty/unsafe ID、missing/non-table target、special-character ID、empty table、thead/multiple tbody/tfoot order和 snapshot detachment。
3. **CSS mapping：** RGB/RGBA/transparent、font family/size/weight、left/center/right/justify/start/end、top/middle/bottom、decimal padding、four border sides/styles/colors、CR/LF、nested visible text和 invalid CSSOM returns。
4. **Widths：** offset proportions、colspan expansion、exact integer sum、explicit scalar/vector override、fixed/min attributes、mixed fixed/min/flexible、constraint overflow、hidden table rejection和 safe integer overflow。
5. **Pagination：** default true、explicit false、weights、all thead repeat、explicit header count、body fallback header、header/body merge rejection、rich fragmentation、first/continuation Y、height/bottom edge和 zero overflow。
6. **Master/lifecycle：** named layout/margin/background/placeholders、same-layout pages、sections、slide numbers、append order、stable returned identities、frozen snapshot、move/delete/duplicate/write/reopen。
7. **Additional objects：** raster/SVG image resolve-once + per-page relationships/dedup, shape, plain/rich text, nested table, placeholder placement, internal links, add order and nested auto-page boundary。
8. **Failure isolation：** DOM/CSS/image preflight zero mutation、first/generated/additional object injected failures、outer rollback/cache cleanup、retry bytes和 caller/DOM isolation。
9. **PptxGenJS conformance：** real Chrome fixtures compare legal CSS/cell/width/header/master/add-object final states；记录 mutation、false-autoPage、fixed-width和 estimator differences而不误报。
10. **Public/release proof：** SDK/root/type declarations、six presentation formats、typecheck/build/full/performance、two clean manifests/tarballs、installed Node/NodeNext/browser/CLI/Inspector、real Chrome、PowerPoint 2010 validation、LibreOffice/render/overflow、docs和 remote `0 0`。

## 完成标准与后续

本项完成时，真实 browser DOM table可以通过 root package生成 1–N张 editable slides；CSS cell state、column constraints、merge、pagination、headers、layout、section、追加对象、transaction和 write/reopen均有 focused/full/packed/client证据。兼容性文档把 `tableToSlides`从 unsupported移到 supported，公开能力覆盖率达到100%。

随后执行最终 peer/client full-surface audit：按完整 PptxGenJS 4.0.1 public declaration矩阵重新核验 supported/extension/deliberate-difference/unsupported四类，运行最终 consumer与客户端 corpus。只有审计证明无公开能力缺口时，才更新 full-parity claim并完成总目标。
