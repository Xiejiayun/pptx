# Table Cell Fit Design

日期：2026-07-28
状态：已批准实施

## 目标与范围

为既有 table cell 增加 direct text-autofit 状态的读取、编辑和清除能力，覆盖 OOXML `a:bodyPr` 的 `noAutofit`、`normAutofit`、`spAutoFit` choice。该能力复用本库已公开的 `TextBoxFit = 'none' | 'shrink' | 'resize'` 语义，但 ownership 是 cell 的 direct `a:tc/a:txBody/a:bodyPr`，不是普通文本 shape。

PptxGenJS 4.0.1 的 `TableCellProps` / `TableProps` 不声明 `fit`，runtime 对 table/cell 上透传的 `fit`、`shrinkText`、`autoFit` 也全部忽略并固定生成 `<a:bodyPr/>`。因此本小项是既有 PPTX 的原生编辑能力，不虚构 PptxGenJS 不存在的 table creation option；adapter 仍忠实导入其公开 `addTable()` 输出。

本小项不增加 table creation、table-level fit default、cell text-direction/vertical-alignment/margin/fill/border/overflow、merge mutation、rich text、row/column mutation、继承解析、字体缩放计算或 shape 尺寸重算。`textDirection` 已由前一独立小项完成。

## 公共 API

沿用已有三值类型，不增加同义 union：

```ts
export type TextBoxFit = 'none' | 'shrink' | 'resize';

export interface TableCell {
  readonly text: string;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
}

export class TableModel {
  readonly rows: readonly TableRow[];

  setCellTextFit(
    rowIndex: number,
    columnIndex: number,
    value: TextBoxFit | undefined,
  ): void;
}
```

用法：

```ts
const table = slide.shapes.find((shape) => shape.kind === 'table') as TableModel;

table.rows[0]?.cells[0]?.textFit; // direct value or undefined
table.setCellTextFit(0, 0, 'shrink');
table.setCellTextFit(0, 0, 'resize');
table.setCellTextFit(0, 0, 'none');      // clear direct supported choice
table.setCellTextFit(0, 0, undefined);   // clear direct supported choice
```

`rows` 继续返回 immutable value snapshot，不引入 live `TableCellModel`。`setCellTextFit()` 延续现有 physical zero-based row/cell 索引式 mutation；修改旧 snapshot 不影响 source，重新读取 `rows` 才看到当前 direct state。

## 方案选择

考虑过三个方案：

1. 新增独立 `TableCellFit` union。ownership 更显眼，但三值与 wire choice 完全相同，会制造重复公共类型和不必要转换。
2. 把每个 cell 升级为 live `TableCellModel`。长期可承载完整 cell API，但本小项会被迫引入 identity cache、merge 坐标和 row/column 生命周期。
3. 复用 `TextBoxFit`，在 `TableCell` snapshot 暴露 `textFit`，由 `TableModel.setCellTextFit()` 负责 indexed edit；采用此方案。它复用稳定的 fit 语义与 codec，同时保持 cell snapshot + indexed mutation 架构。

实现层新增窄 `table-cell-text-fit.internal.ts` wrapper：它只验证 cell 的 unique direct `txBody/bodyPr` 边界，然后复用现有 `text-box-fit.internal.ts` 对 `bodyPr` choice 的 strict read/lossless patch。这样不会复制 fit child 排序、same-mode no-op 和 PowerPoint-calculated metadata 保留逻辑，也不重构公开 text-box API。

`none` 有两种可选 wire 方案：写 `<a:noAutofit/>`，或像已有 text-box/PptxGenJS 语义一样不写 choice。采用后者：setter 的 `none` 与 `undefined` 都移除 direct supported fit children；getter 仍能把既有唯一 `<a:noAutofit/>` 读为 `none`。这保持 PowerPoint 2013 安全语义，也避免同一三值类型在 shape 与 cell 上出现相反写法。

用户已授权实现方持续选择最佳方案并逐项推进，因此本设计按 OOXML ownership、现有 API 一致性、最小改动和可验证性定稿。

## PptxGenJS 4.0.1 基线

类型与公开 writer 实测结果：

- `TableCellProps` 和 `TableProps` 都只继承 `TextBaseProps`，不包含普通 `TextPropsOptions.fit`，也不包含 deprecated `autoFit` / `shrinkText`。
- table/cell runtime 透传 `fit: 'none' | 'shrink' | 'resize'` 均被忽略。
- table/cell runtime 透传 `autoFit: true`、`shrinkText: true` 或三者冲突组合也被忽略。
- table-level runtime 透传不会实体化到 cells。
- 上述七种 cell 的 `a:bodyPr` 均为 `<a:bodyPr/>`，`noAutofit` / `normAutofit` / `spAutoFit` 数量全为 0。
- table-cell writer 的 `genXmlBodyProperties()` 分支固定返回 `<a:bodyPr/>`；普通 text shape 的 fit serializer 不参与 table 输出。

adapter conformance 只使用公开 `addTable()` / `write()`，证明这些 ignored runtime values 导入后均为 `undefined`，cell text 和其他属性不变，不访问 `_slides` 或其他私有字段。本库原生 setter 生成的 fit child 是编辑能力扩展，不宣称是 PptxGenJS table option 对等。

## OOXML 与 direct 语义

三个可读取状态：

```xml
<!-- none: existing explicit source only -->
<a:bodyPr><a:noAutofit/></a:bodyPr>

<!-- shrink -->
<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="20000"/></a:bodyPr>

<!-- resize -->
<a:bodyPr><a:spAutoFit/></a:bodyPr>
```

getter 先要求 cell 恰好一个 direct `txBody`，且其中恰好一个 direct `bodyPr`。随后只检查与 `bodyPr` 相同 prefix 的 direct `noAutofit` / `normAutofit` / `spAutoFit` children：恰好一个返回对应值；absent、多个 choice、duplicate、case variant、不同 namespace/prefix、descendant 或相似 unknown child 返回 `undefined`。它不解析 table style、master/layout/theme、effective default、text length 或 cell size，只读不产生 mutation。

missing/repeated `txBody` 或 `bodyPr` 在 snapshot read 中返回 `undefined`，避免一个 malformed cell 阻断整张 table 的 rows snapshot；setter 对同一结构抛 `ModelParseError`，因为不能安全选择 mutation target。

## 无损编辑与错误边界

setter 在 OPC transaction 内定位目标 table 的物理 row/cell。row 或 cell 不存在时抛 `RangeError`；cell 缺少 unique direct `txBody/bodyPr` 时抛 `ModelParseError`。merged placeholder 只要拥有合法 direct text body 就可编辑；本项不做 logical span 折算。

实际 patch 复用现有 fit codec：

- `undefined` / `none` 移除所有 direct supported fit choice，不新增 `noAutofit`。
- `shrink` / `resize` 产生唯一 `normAutofit` / `spAutoFit`。
- 唯一 same-mode shrink/resize 是 semantic no-op，原样保留 `fontScale`、`lnSpcReduction`、quote style、namespace 与 unknown attributes。
- 切换或规范化冲突 choice 时，在第一个旧 choice 位置写 canonical child，并移除其余 supported choice。
- fit-less self-closing `bodyPr` 被最小展开；expanded `bodyPr` 在 `scene3d` / `sp3d` / `extLst` 前插入，否则追加到末尾，保持 schema order。

`bodyPr` 的 wrap/margin/anchor/vert/RTL/columns/warp/3D/extension/unknown XML、cell `tcPr`、text、paragraph/run、merge state、neighboring cells、table relationships 和 `TableModel` identity 保持不变。`setCellText()`、`setCellTextDirection()`、transform、duplicate、write/reopen 等非 fit mutation 必须保留 fit。

value 对 null、boolean、number、空字符串、case/whitespace variant、未知字符串、array、object 和 symbol 抛 `TypeError`。validation、索引与结构检查均在 part mutation 前完成；失败与 outer transaction rollback 保持 exact bytes、mutation journal、slide/table identity 和 fresh rows snapshots。

## 测试与发布门禁

1. model fixture 覆盖 absent、唯一 no/norm/sp、case/namespace/descendant/unknown、duplicate/mixed choice、missing/repeated `txBody/bodyPr`；getter 不产生 mutation。
2. `setCellTextFit()` 覆盖 self-closing/expanded bodyPr 的 shrink/resize/none/clear、same-mode calculated metadata、conflict normalization、schema-order insertion、merged placeholder、invalid index/structure 和 outer rollback。
3. `setCellText()`、`setCellTextDirection()`、transform、duplicate、write/reopen 保留 fit；只编辑一个 cell 不改变相邻 cell、`tcPr` 或 relationships。
4. SDK 覆盖 immutable snapshots、public transaction、invalid runtime values/coordinates、package isolation 与 stable identity。
5. PptxGenJS 4.0.1 conformance 覆盖 table/cell-level `fit`、`autoFit`、`shrinkText` runtime passthrough 全部 ignored，raw bodyPr 不含 fit child。
6. API、npm README、changelog 与兼容矩阵明确：既有 table-cell fit 原生编辑已支持；PptxGenJS 4.0.1 本身无对应 table API；table creation 仍未支持。
7. packed Node/browser/declaration smoke 覆盖 snapshot、shrink/resize/none/clear；typecheck、全仓 tests、独立 performance、actual tarball 和 CLI 全部通过。
8. 同源 native/hand-patched table 文件通过 PowerPoint 2010 profile 的 0 error / 0 warning 验证且 package diff 为空；LibreOffice 无修复导出并逐图检查，两份 overflow checks 通过。动态字体 scale/shape resize 不由本库计算。
