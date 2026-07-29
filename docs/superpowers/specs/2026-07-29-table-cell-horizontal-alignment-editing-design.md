# Table Cell Horizontal Alignment Editing Design

日期：2026-07-29
状态：已批准实施

## 目标与范围

为既有 table 的严格单段落 physical cell 增加 direct 水平对齐 snapshot、编辑和清除能力。公开值复用 `TextAlignment` 的 `left`、`center`、`right`、`justify`，并只拥有选中 cell 唯一 direct `a:txBody/a:p/a:pPr@algn`。

本小项承接已经完成的 cell/table 创建期 `align`。Native 创建后返回的 live `TableModel`、PptxGenJS 4.0.1 输出和已有 deck 都能通过同一 API 读取、替换与清除 alignment。它不把 direct 缺失推断为 effective left，也不读取 table style、master/layout/theme 或保留创建期 default source。

当前 `TableCell` 只公开 collapsed plain text，不公开逐段 rich-text 模型。为避免把异构多段落错误压缩成一个值，本小项严格要求一个 direct text body 和一个 direct paragraph；rich/multi-paragraph table-cell snapshot/editing 将在后续独立小项与完整 rich-text value model 一起实现。

## 公共 API

复用已有公共类型：

```ts
export type TextAlignment = 'left' | 'center' | 'right' | 'justify';

export interface TableCell {
  readonly text: string;
  readonly horizontalAlignment?: TextAlignment;
  // existing borders/fill/margins/textDirection/textFit/verticalAlignment remain unchanged
}

export class TableModel {
  setCellHorizontalAlignment(
    rowIndex: number,
    columnIndex: number,
    value: TextAlignment | undefined,
  ): void;
}
```

用法：

```ts
const table = slide.shapes.find((shape) => shape.kind === 'table') as TableModel;

table.rows[0]?.cells[0]?.horizontalAlignment; // direct value or undefined
table.setCellHorizontalAlignment(0, 0, 'left');
table.setCellHorizontalAlignment(0, 0, 'center');
table.setCellHorizontalAlignment(0, 0, 'right');
table.setCellHorizontalAlignment(0, 0, 'justify');
table.setCellHorizontalAlignment(0, 0, undefined); // clear direct algn
```

`rows` 继续返回 detached immutable snapshots；不新增 live `TableCellModel`。索引沿用其他 table-cell setter 的 zero-based physical row/cell 语义，不折算 merge/span 逻辑坐标。

## 方案选择

考虑三个方案：

1. 对严格单段落 cell 增加 scalar `horizontalAlignment` 和 scalar setter；采用此方案。它准确对应当前 public cell text model、已经支持的 native table creation，以及 PptxGenJS cell-level scalar `align`，不会隐藏多段落信息损失。
2. 对任意多段落 cell 返回共同 alignment，并让 setter 批量覆盖全部 paragraphs；不采用。heterogeneous、部分缺失与 malformed direct state 都会被压成 `undefined`，调用方无法区分，编辑还会无提示覆盖每段独立意图。
3. 立即增加 `horizontalAlignments[]` 和 paragraph-index setter；不采用。该 API 只暴露 alignment 而不暴露对应 runs、paragraph text 和其他 paragraph properties，会成为完整 table-cell rich-text model 的一次性中间表面。

Rich/multi-paragraph support 不是被删除的目标，而是下一阶段与 `TableCell.richText`/paragraph identity 一起设计，避免当前小项锁定一个不完整公共模型。

用户已授权实现方持续自主决定最佳方案并推进全部缺口，因此本设计按推荐方案直接批准实施。

## PptxGenJS 4.0.1 基线

PptxGenJS 4.0.1 的 `TableProps.align` 与 `TableCellProps.align` 都复用 `TextBaseProps.align`。公开 runtime 输出已证明：

- `left`、`center`、`right`、`justify` 分别物化为 cell paragraph direct `pPr@algn="l"`、`"ctr"`、`"r"`、`"just"`；
- table value 复制到未覆盖 cell，cell value 优先，最终文件不保留 table-level alignment metadata；
- omitted/`undefined` 不写 direct token；
- unknown runtime table/cell value 被静默忽略，而 native setter 继续严格拒绝。

PptxGenJS 没有 existing-deck object model 或 editor，因此 snapshot/editing 是本库的 lossless extension。对等证据聚焦 PptxGenJS 已 materialize 的 supported direct cell state，不声称存在 PptxGenJS 读取/编辑 API。

## OOXML ownership 与 strict read

目标结构：

```xml
<a:tc>
  <a:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r><a:t>Center</a:t></a:r>
    </a:p>
  </a:txBody>
  <a:tcPr/>
</a:tc>
```

getter 只接受：

1. cell 恰好一个 direct `txBody`；
2. text body 恰好一个 direct `p`；
3. paragraph 零个或一个 direct `pPr`；
4. properties 上零个或一个 unqualified `algn`；
5. token 精确为 `l`、`ctr`、`r` 或 `just`。

合法 token 映射为 public value。缺少 `pPr`/`algn` 返回 `undefined`；missing/repeated `txBody`、zero/repeated paragraph、repeated `pPr`、repeated unqualified attribute、namespaced-only attribute、empty、case/whitespace variant、long-form value、`dist`、`thaiDist`、`justLow` 和 unknown token 也返回 `undefined`。getter 不抛出 malformed cell error，避免一个坏 cell 阻断整张 table snapshot；只读不产生 package 或 mutation-journal 变化。

`bodyPr@algn`、`tcPr@algn`、descendant/foreign namespace elements、table style 和 effective defaults 都不是本字段的 state。

## 无损编辑

新增窄内部 codec `table-cell-horizontal-alignment.internal.ts`，负责 direct paragraph 定位、wire mapping 和 source-span patch；`TableModel` 只负责 public normalization、physical index、transaction 与保存。

setter 在 mutation 前完成 strict value normalization。`undefined` 清除 direct unqualified `algn`；四个 public value canonicalize 为 `l/ctr/r/just`。

结构要求：

- 恰好一个 direct `txBody` 和一个 direct `p`；否则抛 `ModelParseError`。
- paragraph 最多一个 direct `pPr`；重复时抛 `ModelParseError`。
- direct `pPr` 最多一个 unqualified `algn`；重复时抛 `ModelParseError`。
- 缺少 `pPr` 且设置合法值时，在 paragraph 第一个 child 前插入与 paragraph 相同 namespace prefix 的 `<pPr algn="..."/>`。
- 缺少 `pPr` 且清除时为 exact no-op。
- 既有 self-closing 或 expanded `pPr` 只增加、替换或删除目标 attribute；清除后保留空 `pPr`，不擅自删除可能有继承语义的 element。
- same canonical assignment 和 absent clear 是 exact no-op，不改变 bytes 或 mutation journal。
- 单个 unknown unqualified token 可被合法值 canonicalize，或由 `undefined` 清除。
- namespaced `x:algn` 与所有非目标 attribute/child 原样保留；setter 只拥有 unqualified `algn`。

属性顺序、quote style、paragraph runs、fields、breaks、end properties、bullet、spacing、margin/indent、RTL、tab stops、extensions、body properties、cell margins/vertical alignment/direction/fit/border/fill、merge state、neighbor cells、geometry、relationships 与 object identity 都必须保持。

## 错误与事务语义

`value` 除 exact `left`、`center`、`right`、`justify` 或 `undefined` 外全部严格拒绝，包括 null、boolean、number、empty string、case/whitespace variant、wire token、`dist`、array、object 和 symbol。复用 `normalizeTextAlignment(value, 'Table cell horizontal alignment')`，不复制 public value validation。

不存在的 physical row/cell 抛 `RangeError`。不安全的 paragraph 结构抛带 slide part URI 的 `ModelParseError`。value、索引和结构失败必须是 zero mutation；outer `PptxDocument.transaction()` rollback 恢复 exact bytes、journal 与 fresh snapshots。

合法操作放在 OPC transaction 中。只在 codec 报告实际变化时调用 `slide.setXml()`，保证 same-value/absent-clear no-op。

## 生命周期与互操作

创建期 cell/table `align` 物化出的 direct token 必须立即通过 `horizontalAlignment` 可见。后续 `setCellText()`、cell border/fill/margin/vertical alignment/direction/fit、column/row sizing、transform、slide duplicate、write/reopen 都保留它；horizontal editor 也不得改变这些字段。

设置/清除目标 cell 只能改变 owning slide part。duplicate 后编辑 clone 不改变 source；outer rollback 恢复 source/clone；write/reopen 后 snapshot 与 canonical direct XML 一致。

## 测试与验收

1. Model fixture 覆盖 `l/ctr/r/just`、absent、missing/repeated `txBody`、zero/repeated paragraphs、missing/repeated `pPr`、empty/case/whitespace/long-form/unsupported/unknown token、namespaced/repeated attribute 与 descendant impostors。
2. Snapshot 覆盖 detached immutable rows、direct-only read、getter zero mutation，以及 native-created table/cell alignment 的即时可见性。
3. Editor 覆盖 missing/self-closing/expanded `pPr` 的 add、replace、四值、clear、same-value no-op、unknown canonicalization、namespace-prefix preservation、invalid index/structure/value 与 outer rollback。
4. Isolation 覆盖 paragraph properties/runs/fields/breaks/extensions、cell properties、neighbor cells、other package parts、stable identity 和 mutation journal。
5. Lifecycle 覆盖所有现有 cell/property/geometry mutation、slide duplicate、write/reopen 与 transaction rollback 保留 direct alignment。
6. SDK 覆盖 public Node API、invalid runtime values/coordinates、package isolation、created/existing tables 和 stable identity。
7. PptxGenJS 4.0.1 fixture 覆盖 table/cell valid materialized values、cell precedence、omission 和 invalid fallback divergence；adapter 导入 snapshot 与 wire XML 一致。
8. Packed Node/browser/declaration smoke 覆盖 snapshot、四值 editor、clear、creation→edit→write→reopen，CLI 保持 `0.1.0`。
9. 文档更新 changelog、API README、compatibility baseline 与 package README，明确 strict single-paragraph direct ownership 和 rich/multi-paragraph 后续边界。
10. TypeScript project references、focused/full Vitest、performance、actual tarball、PowerPoint 2010 profile、package mutation diff、LibreOffice/Poppler render、original-resolution inspection 与 overflow checks 全部通过。

## 非目标与后续

本小项不新增 table-level alignment getter/editor、effective inheritance、multi/rich paragraph table-cell model、per-paragraph editor、`dist`/`thaiDist`/`justLow`、direction/fit creation、hyperlink、merge/span mutation、row insertion/deletion、table styles、auto-page、repeated headers、内容测量或 layout recomputation。

下一 table-cell text slice 将以完整 paragraph/run value model 支持 rich/multi-paragraph snapshot、creation 和 replacement，并在那里提供逐段 alignment，而不是扩展本 scalar API 去猜测 heterogeneous state。
