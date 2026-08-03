# Table Auto-Page and Repeated Headers Design

日期：2026-08-04
状态：已批准实施

## 目标与范围

为 native `SlideModel.addTable()` 增加严格的跨幻灯片分页与重复表头，使一组具有明确 direct row heights 的 logical table rows 能从当前 slide 开始、在可用垂直区域耗尽时创建后续 slides，并在每个 continuation table 前重复指定的 header rows。新增 slides 必须位于 source slide 之后、继承同一 layout、保持 section 顺序和 stable `SlideModel` identity；所有 table text、rich style、merge、hyperlink、column width 与 relationship ownership 必须在每个 page 上保持合法。

本小项完成 **explicit-row-height structural pagination**。它不猜测 automatic row 的 rendered height，不拆分一个 row 的文字内容，也不实现 `autoPageCharWeight` / `autoPageLineWeight`。这些依赖字体、段落、软换行、margin 与字符宽度的能力进入下一项 content measurement/layout recomputation；完成后会移除本阶段“必须提供 positive `rowHeights`”的限制。`tableToSlides` 的 HTML/CSS 导入和附加对象仍是再下一项。

用户已明确授权实现方自主选择最佳方案并连续推进，因此本设计在完成方案比较与自审后直接批准，不设置额外确认点。

## 公共 API

```ts
export type TableAutoPageMarginInput =
  | number
  | readonly [number, number, number, number];

export interface AddTableOptions {
  readonly autoPage?: boolean;
  readonly autoPageRepeatHeader?: boolean;
  readonly autoPageHeaderRows?: number;
  readonly autoPageSlideStartY?: number;
  readonly slideMargin?: TableAutoPageMarginInput;
  // existing fields unchanged
}

export class SlideModel {
  get newAutoPagedSlides(): readonly SlideModel[];
}
```

所有 native geometry 值继续使用 EMU。`slideMargin` scalar 广播四边，tuple 顺序固定为 `[top, right, bottom, left]`；每项是 non-negative safe integer，水平和垂直 margin 总和必须分别小于 slide width/height。未提供时优先使用当前 runtime named-layout margin，否则使用 canonical 0.5-inch 四边 margin。

`autoPageSlideStartY` 是 continuation slides 上 table 的 absolute Y，必须是 non-negative safe EMU，并在 bottom margin 之前。首张 slide 始终使用 table `y`。`autoPageRepeatHeader` 默认为 false；为 true 时 `autoPageHeaderRows` 默认为 1，显式值必须是 `1..rowCount` 的 positive safe integer。分页控制字段只在 `autoPage: true` 时合法，防止无效果配置被静默接受。

`SlideModel.newAutoPagedSlides` 返回最近一次成功 `addTable()` 调用创建的 continuation slides，不包含 source slide。下一次成功的 non-auto-page 或无 overflow table creation 将其重置为 frozen empty array；失败调用保留上一次成功结果。该状态仅存在于当前 runtime，write/reopen 后为空。Move 保持集合与 identity，删除 generated slide 后 getter 不再返回已脱离 presentation 的对象，duplicate source slide 从 empty runtime state 开始。

## 方案选择

1. **按 explicit direct row heights 分页，后续接入独立测量器；采用。** 当前库已经有 strict row-height vector、merge topology、rich-cell renderer 和 relationship lifecycle。先建立 deterministic page partition、slide insertion 与 repeated-header ownership，能以 exact EMU 验证行为，同时为下一项测量器提供稳定 consumer。
2. **一次实现 PptxGenJS 式字符估算、row splitting 与 pagination。** 表面进度更快，但会把 page lifecycle、字体测量、段落断行、row fragment rendering 和 layout recomputation 混为一个不可独立 review 的改动；PptxGenJS 4.0.1 本身还存在 colspan width、truthy option、input mutation 与 existing-slide reuse 缺陷，因此不采用。
3. **新增独立 `paginateTable()` / `addAutoPagedTable()` API。** 返回值可以更整洁，但无法覆盖 PptxGenJS 声明的 `addTable({ autoPage })` 与 `slide.newAutoPagedSlides` public surface，且会制造第二套 creation path，故不采用。

## 输入归一化与前置条件

现有 descriptor-safe table option reader增加五个 canonical fields。Unknown、symbol、accessor、inherited、class-instance、sparse tuple、coercible numeric 或非 boolean input 在 package读取前拒绝。所有 caller arrays/objects 在 normalization 时脱离，后续 mutation caller 不改变 page plan。

`autoPage: true` 在本阶段要求：

- `rowHeights` 已显式提供，且每个 physical row height 都大于 0；automatic zero row 在任何 mutation 前以清晰错误拒绝；
- 不使用 `placeholder`。Continuation slides 的 placeholder replacement 与 runtime page geometry存在独立 ownership问题，留到测量/layout阶段统一定义；
- normalized table width 等于 exact column-width sum，full table height 等于 full row-height sum；
- source `y`、continuation Y、bottom margin 形成 positive usable height；
- repeated header rows本身完整，不与 body 之间存在跨 boundary rowspan；
- 每个不可拆 row block 在 continuation usable area中可完整容纳。

Auto-page metadata只参与 runtime partition，不写入 DrawingML、extension、custom property或 retained table model state。每个 page 的最终 PPTX 仍只是普通 `p:graphicFrame/a:tbl`。

## 分页几何

首张 page 的 usable vertical interval 是 `[table.y, slide.height - bottomMargin)`；continuation page 使用 `[autoPageSlideStartY ?? topMargin, slide.height - bottomMargin)`。Table X、column widths、total width、name和 direct cell formatting在所有 pages保持；每个 page transform height精确等于其 included row heights之和，不保留 full-table input height。

分页以 safe integer EMU加法完成。Header rows先占用 page height，body block按原顺序追加；下一个 block超出当前 capacity时完成当前 page并在新 page从 repeated headers重新开始。没有 overflow时只创建 source table，`newAutoPagedSlides` 为空。

Repeated headers关闭时，source page必须至少容纳第一个 row block，否则拒绝，因为 native `addTable()` 必须返回 source slide上的 `TableModel`。Repeated headers开启时，source page至少包含 headers；如果首个 body block只适合 continuation capacity，可以生成 header-only source table，随后 body从第一张 generated slide开始。Continuation page必须包含至少一个 body block，避免只含 header的尾页。

## Merge 与 Row Block

分页输入使用已展开的 physical rows，但不能在任意 physical index切断 merge。Internal partitioner从 normalized anchor `rowspan` 构造最小 contiguous row blocks：一个 block包含从起始 row到所有 active rowspan结束的完整区间；相邻或嵌套 span会扩展同一 block。Colspan不影响纵向分组。

Header boundary不能穿过任意 rowspan。合法 header-only merge会在每个 continuation page完整复制；body merge始终随所属 block移动，anchor/continuation `rowSpan/gridSpan/vMerge/hMerge` tokens保持相同局部语义。Page slice不重建 cell，不改变 rich content/style、hidden continuation、unknown XML语义或 hyperlink definition。

单个 block加 repeated-header height超过 continuation capacity时严格拒绝；本阶段不缩放、截断、拆 row或生成超出 slide的 table。下一项 content measurement可以把没有 rowspan的超高 text row拆成 canonical fragments，但不能改变本设计的 merge atomicity。

## Slide 创建、顺序与 Layout

新增一个 model-internal “blank slide after source” primitive，不通过 append-then-public-move暴露中间状态。它：

- 沿 source slide 唯一 direct internal slide-layout relationship创建相同 layout owner；missing、ambiguous、external或错误 content type在 mutation前拒绝；
- 创建 canonical blank slide，materialize与普通 `addSlide()` 相同的合法 layout placeholders和 slide-number state；
- 把新的 `p:sldId` 直接插入 source/previous generated slide之后；
- 精确复制 source section membership，并重新按 presentation order排序 section members；
- 同步 direct slide-number caches，保留既有 slides、IDs、relationships、sections和 unknown presentation XML；
- 返回 presentation cache中的 stable `SlideModel`。

Auto-page生成多张 slide时始终连续插入，不复用或修改原先位于 source之后的既有 slides。这是对 PptxGenJS 4.0.1 runtime会把 page table添加到既有 next slide缺陷的明确修正。Source slide不是 presentation最后一张时，原后续 slides整体后移但内容、identity和section保持。

## Table 与 Relationship 提交

完整 normalized definition先被 partition为 ordinary page definitions。所有 URL/internal-slide hyperlink targets在插入任何 slide之前解析成 stable relationship inputs，避免新 slide插入导致 one-based internal slide index漂移。随后每个 target slide独立分配 relationship IDs并渲染其 table：

- repeated header hyperlink在每张 continuation slide拥有该 slide自己的 relationship；
- page内相同 target可继续遵循当前 relationship分配/共享 contract；
- URL、internal slide、tooltip、run-local/outer link和underline保持；
- relationship只能由本 page的 XML引用，不把 source `rId`复制到另一 slide；
- rollback删除全部新 relationships/parts/slides，不留下 orphan或 dangling link。

现有 `addTable()` 的 normalized commit逻辑提取为可供 source和generated slide复用的单一路径。Page definition不重新经过 public input normalization，避免 resolved table/cell defaults、rich text或merge state被第二次解释。

## Transaction、缓存与失败隔离

Normalization、page partition、layout/section topology检查、hyperlink target resolution和全部 page plans在第一处 package mutation前完成。Source table、generated slides、relationships、presentation order、section membership与slide-number cache在一个 outer OPC transaction内提交。

任何 `setPart`、relationship、placeholder materialization、slide insertion、table render或outer transaction failure都恢复 parts、relationships、ZIP dates和mutation journal。JS-side slide model cache对本次新建 part URI执行显式 cleanup；`newAutoPagedSlides`只在 transaction成功后替换。失败前后的 package bytes、slide collection identity、source table collection和先前 runtime page result保持一致。

## PptxGenJS 4.0.1 边界

对等 public surface包括 `TableProps.autoPage`、`autoPageRepeatHeader`、`autoPageHeaderRows`、`autoPageSlideStartY`、`slideMargin`和`Slide.newAutoPagedSlides`。本阶段以显式 row heights提供比 PptxGenJS估算器更确定的 structural pagination；实际 PptxGenJS plain/rich/linked/merged合法 output仍用于 final table state、header copy、relationship与slide ordering conformance。

Native不复制以下 runtime行为：

- `autoPageLineWeight`超范围时静默 clamp，或用 truthy fallback折叠 explicit zero；
- 修改 caller options/cell objects并注入 `_lines`、`_lineHeight`、`_arrObjTabHeadRows`；
- colspan width过滤错误、rowspan line-height suppression和过高 row的近似拆分；
- 在 source后已有 slide时复用并修改该 slide；
- invalid header count、margin或numeric token的 coercion/warning-only fallback。

`autoPageCharWeight`、`autoPageLineWeight`、automatic row height与 text row fragmentation会在下一项以 strict native measurement API完成后加入，不在当前文档中虚报支持。

## 测试与验收

1. Option normalization：boolean/cross-field/header count/startY/margin/descriptor/dense tuple/caller detachment和 automatic-row rejection。
2. Pure partition：first/continuation capacity、exact boundary、multi-page、header-only first page、no overflow、safe-sum overflow、oversized block和zero usable area。
3. Merge atomicity：vertical/rectangular/nested-adjacent blocks、header-local merge、header/body crossing rejection、token preservation和reopen。
4. Slide lifecycle：source middle/end、same layout、placeholder/slide-number materialization、same section、contiguous order、stable identity、move/delete/duplicate和runtime reset。
5. Relationships：plain/rich/default/local URL与internal links、repeated header IDs per slide、pre-insertion target stability、shared/unique ownership和zero orphan links。
6. Public contracts：model/SDK/root declarations、six formats、`newAutoPagedSlides` readonly runtime behavior和PptxGenJS 4.0.1 public-output comparison。
7. Failure isolation：invalid inputs/topology/preflight zero mutation、injected package failures、outer rollback、model-cache cleanup、prior runtime-state preservation和deterministic write。
8. Release proof：focused/typecheck/build/full/performance、two clean dist manifests、two byte-identical actual tarballs、installed Node/NodeNext/browser/CLI/Inspector、real Chrome、PowerPoint 2010、docs/review/commit/push和remote divergence `0/0`。

## 完成标准与后续

本项完成时，explicit positive row heights可以通过 `addTable({ autoPage: true })` deterministic分页，repeated headers、continuation Y/margins、same-layout slide insertion、section/identity、rich/merge/hyperlink state与`newAutoPagedSlides`全部通过完整生命周期和实际包/浏览器验证。

下一项接入 content measurement/layout recomputation，支持 omitted/automatic row heights、strict `autoPageCharWeight` / `autoPageLineWeight`和无 rowspan text-row fragmentation；随后完成 `tableToSlides`，最后执行 peer/client full-surface audit。
