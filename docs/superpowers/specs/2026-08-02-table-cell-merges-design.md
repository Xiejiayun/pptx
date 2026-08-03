# Table-Cell Merge, Colspan, and Rowspan Design

日期：2026-08-02
状态：已批准实施

## 目标与范围

为 native `slide.addTable()` 增加 PptxGenJS 4.0.1 对等的 cell `colspan` / `rowspan` 创建能力，并为已有 PPTX table 增加严格的合并拓扑读取、detached snapshot、`mergeCells()` 与 `unmergeCell()`。合法二维矩形合并必须生成 DrawingML 的 `gridSpan` / `rowSpan` anchor 与 `hMerge` / `vMerge` continuation 组合；已有表编辑只改变合并属性，保留每个 physical cell 的内容、样式和 relationship，确保 unmerge 可无损恢复。

本小项不增加 row/column CRUD、cell insert/delete、auto-page、repeated headers、content measurement、layout recomputation 或 `tableToSlides`。用户已授权实现方自主确定后续内容并连续推进，因此本设计按推荐方案直接批准实施，不设置额外等待确认点。

## 公共 API

```ts
export interface AddTableCellOptions {
  // existing fields remain unchanged
  readonly colspan?: number;
  readonly rowspan?: number;
}

export interface TableMergeRegion {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly rowspan: number;
  readonly colspan: number;
}

export interface TableCellMerge extends TableMergeRegion {
  readonly isAnchor: boolean;
}

export interface TableCell {
  // existing fields remain unchanged
  readonly merge?: Readonly<TableCellMerge>;
}

export class TableModel {
  get mergeRegions(): readonly Readonly<TableMergeRegion>[] | undefined;
  mergeCells(
    rowIndex: number,
    columnIndex: number,
    rowspan: number,
    colspan: number,
  ): void;
  unmergeCell(rowIndex: number, columnIndex: number): void;
}
```

`colspan` / `rowspan` intentionally retain PptxGenJS and HTML table spelling instead of adding a second alias family. They accept positive safe integers; omitted and `1` mean an unmerged dimension and emit no attribute. `mergeCells()` requires the requested rectangle to cover at least two cells. All public coordinates are zero-based physical row/column coordinates, matching every existing indexed table-cell editor.

`mergeRegions` returns `[]` for a recognized unmerged table, a row-major detached deeply frozen list for a recognized merged table, and `undefined` when direct table structure or merge topology is not safely recognizable. `TableCell.merge` is present for every physical cell in a recognized merged region, points to the same anchor coordinates and spans, and distinguishes the anchor with `isAnchor`; it is omitted for unmerged cells and for all cells when the table topology is unsupported.

## 方案选择

1. **严格逻辑输入布局、physical snapshot、lossless attribute-only editing；采用。** 创建期允许 logical rows 因 `rowspan` 而变短，但先展开为完整 physical matrix；读取和编辑继续使用现有 physical coordinates。已有 cell bytes 和 relationships 不被 merge/unmerge 隐式删除，最符合本库的 lossless editor contract。
2. **始终要求调用方显式提供 continuation cells。** 实现较短，但与 PptxGenJS `colspan` / `rowspan` surface 不对等，调用方必须手写 `hMerge` / `vMerge` 拓扑，且容易生成非法 OOXML，因此不采用。
3. **改成 logical-only rows 并让已有 indexed editors 自动重定向到 anchor。** 表面更像 HTML，但会破坏当前 physical row/cell 数量和坐标语义，隐藏 continuation 中可恢复的数据，也使既有编辑器发生不兼容变化，因此不采用。

## 创建期逻辑布局

Outer rows 仍必须是 non-empty dense data array，首行必须包含至少一个 logical cell。首行每个 cell 的 normalized `colspan` 之和定义 physical column count。后续 logical row 可为空；布局器按输入顺序把每个 cell 放到当前行最左侧尚未被上方 active `rowspan` 占用的 physical column，并为其完整 `rowspan × colspan` 矩形预留位置。

每行布局结束后，来自既有 row spans 和本行 anchors 必须恰好覆盖全部 physical columns。任何空洞、越过首行列数、越过总行数、重叠 rectangle、不完整最后 continuation，或展开后超过 1,000,000 个 physical cells 都在创建前拒绝；总量上限保证单个巨大 span 也不会在校验阶段触发无界分配。这样可表达 PptxGenJS 的常见 lopsided input：

```ts
slide.addTable([
  [{ text: 'A', options: { colspan: 2, rowspan: 2 } }, 'B'],
  ['C'],
]);
```

它展开成 `2 × 3` physical matrix：第二行前两格由 A 占用，`C` 自动放到 physical column 2。一个完全被上方 row spans 覆盖的后续 logical row 可写为 `[]`。Input row/cell objects、rich text、options 和 arrays 全部 descriptor-safe、detached，并且不会像 PptxGenJS writer 那样被 splice 或注入内部字段。

Column widths、table width 和 physical grid count 使用展开后的 physical column count；row heights 和 table height 仍使用 physical row count。Placeholder scaling、transform、table defaults 和 anchor cell 的 content/style/hyperlink precedence 保持现有行为。

## OOXML 渲染

每个 logical anchor 使用原 cell content、`txBody`、`tcPr` 和 hyperlinks。由布局器生成的 continuation cells 不接受 caller content/style/relationships，并 canonical render 为只有 `<a:tcPr/>` 的 cell。四类位置按以下规则写属性，顺序固定为 `rowSpan`、`gridSpan`、`vMerge`、`hMerge`：

| 位置 | 属性 |
| --- | --- |
| top-left anchor | `rowSpan="R"` when R > 1；`gridSpan="C"` when C > 1 |
| top-row horizontal continuation | anchor `rowSpan` when R > 1；`hMerge="1"` |
| left-column vertical continuation | anchor `gridSpan` when C > 1；`vMerge="1"` |
| interior continuation | `vMerge="1" hMerge="1"` |

一维 horizontal/vertical merge 是同一规则在 R=1 或 C=1 时的自然退化。Hyperlink relationship matrices 改为 physical dimensions；只为 anchor cells 分配 default/run relationships，continuations 永远不创建 relationship。未使用 span 的所有 legacy table output 必须 byte-identical。

## 读取与拓扑识别

新增单一内部 merge boundary，解析 direct `a:tbl`、direct `a:tblGrid/a:gridCol`、direct `a:tr` 和 direct `a:tc`。安全 topology 要求：

- 每行 physical cell 数相同，并等于唯一 direct grid 的 column count；
- `gridSpan` / `rowSpan` 是唯一 unqualified attribute，值为 positive safe integer；leading zeros 可识别；omitted 或 `1` 视为单格；
- `hMerge` / `vMerge` 是唯一 unqualified XML Schema boolean，接受 `1` / `true` / `0` / `false`；false 视为未合并；
- 每个 anchor rectangle 在 table bounds 内且不重叠；
- 每个 continuation 的四个 semantic values 与其 anchor 和相对位置完全一致；
- 不存在 orphan continuation、continuation-owned new span、span anchor 缺失或同一 cell 的 repeated merge attribute。

识别采用 whole-table all-or-nothing contract。未知 attributes、namespace declarations、cell children 和 opaque extensions 不影响 merge recognition；merge snapshot 不推断 effective style 或 content ownership。Malformed/ambiguous state 仍可由既有 `rows` 读取 physical cell 内容，但 `mergeRegions` 为 `undefined`、所有 `TableCell.merge` 均省略，semantic merge/unmerge 拒绝修改。

## 已有表合并与拆分

`mergeCells(rowIndex, columnIndex, rowspan, colspan)` 以指定 top-left physical cell 为 anchor。它先完成完整 topology、integer、bounds 和 overlap validation：

- 如果请求与一个现有 region 完全相同，package bytes、relationships、ZIP dates、model identity 和 mutation journal exact no-op；
- 如果请求 rectangle 与任何现有 region 部分或完整相交但不完全相同，抛出明确错误，调用方必须先 unmerge；
- 不相交的现有 merges 保持原 bytes，可继续新增另一个 region；
- unsupported topology、1×1 请求、越界或非法数值在 mutation 前拒绝。

`unmergeCell(rowIndex, columnIndex)` 接受 region 中任意 physical member；它定位所属 anchor 并移除整个 rectangle 的四类 merge attributes。Unmerged cell 是 exact no-op；unsupported topology 或越界拒绝。

两个 editor 都只重写受影响 `a:tc` start tag 中的 `gridSpan`、`rowSpan`、`hMerge`、`vMerge`，并 canonicalize 这些已编辑 tokens。所有其他 attributes 及其 values、start tag 外的 whitespace、`txBody`、`tcPr`、unknown children、cell content、styles 和 slide relationships 原样保留。因而 continuation 的隐藏内容和 hyperlink 不会丢失，unmerge 后重新可见，也不需要 relationship GC。既有 `setCellText()`、`setCellRichText()` 和 style editors 继续使用 physical coordinates，并允许显式编辑 continuation 的保留状态；它们不自动重定向到 anchor。

## PptxGenJS 4.0.1 基线

锁定的 `types/index.d.ts` 公开 `TableCellProps.colspan?: number` 与 `rowspan?: number`。真实 writer 与 public `write()` 输出确认：

- first-row span sum 用于 grid width；logical rows 会被 writer 原地 splice 成 physical rows；
- 2×2 anchor 写 `rowSpan="2" gridSpan="2"`，top continuation 写 `rowSpan="2" hMerge="1"`，left continuation 写 `gridSpan="2" vMerge="1"`，interior 写 `vMerge="1" hMerge="1"`；
- continuation cells 只有 `<a:tcPr/>`，不会继承 anchor/table cell text style；
- 合法 offset、horizontal、vertical 和 rectangular spans 均可由 native 生成同一 semantic final state。

PptxGenJS 对 lopsided non-span rows、negative/fractional spans 和 out-of-bounds rowspans 缺少严格验证，可输出 grid/row cell counts 不一致、zero grid columns、fractional `rowSpan` 或缺失 continuation。Native 不复制这些缺陷：所有非法布局在 shape ID、relationship、part bytes或 mutation journal变化前拒绝。

## 错误、事务与兼容边界

- `colspan` / `rowspan` unknown、symbol、accessor、inherited 或 class-instance fields 继续由 descriptor-safe readers 拒绝；non-number、non-finite、fractional、zero、negative 和 unsafe integer 分别使用明确 type/range errors。
- 逻辑布局、rich text、style、hyperlink target 和 all spans 全部在 add-table transaction 的 observable mutation 前完成。
- Merge/unmerge 运行在既有 OPC transaction 中；任一 injected write failure 完整恢复 part bytes、relationships、dates、cache identity 和 journal。
- Duplicate/move/write/reopen 自然保留 direct merge attributes；duplicate table wrapper identity 与 sibling isolation 沿用现有 slide lifecycle。
- Existing PptxGenJS/PowerPoint valid merge topology 可读取、编辑和重开；unsupported state 保留原 bytes，不尝试猜测或自动修复。

## 测试与验收

1. Internal topology TDD：horizontal、vertical、2D、offset、多 region、full-row coverage、leading-zero/boolean lexical forms、detached/frozen snapshots、foreign-namespace impostor isolation，以及 orphan/overlap/out-of-range/repeated/mismatched grid/row rejection。
2. Creation TDD：logical-to-physical expansion、empty covered row、anchor styles/rich text/hyperlinks、physical relationship matrices、sizes、strict invalid spans/layouts、input detachment 和 no-span legacy byte parity。
3. Model lifecycle：`mergeRegions` / `TableCell.merge`、merge/unmerge exact no-op、hidden cell state preservation、continuation physical editing、non-overlap merge、overlap rejection、opaque XML preservation、rollback、duplicate、six formats 与 write/reopen。
4. SDK/root/declaration：new fields/types/methods 的 positive/negative TypeScript contract，Node/browser exports 和 generated declarations。
5. PptxGenJS adapter：4.0.1 legal horizontal/vertical/rectangular/offset final-state conformance，以及 input mutation和非法 output 的 documented strict divergence。
6. Package proof：actual tarball Node/NodeNext/browser/CLI/Inspector、真实 Chrome create/read/merge/unmerge/reopen、PowerPoint 2010 validation、part inspection、determinism 和 zero console/page/network errors。
7. Final gates：focused Vitest、project typecheck/build、full Vitest with two workers、independent performance、docs review、commit、push 与 local/remote `0/0`。

## 后续边界

本小项完成 table merge/colspan/rowspan 的创建、读取与编辑。后续小项依次处理 row/column CRUD、auto-page/repeated headers、content measurement/layout recomputation、`tableToSlides` 与最终 peer/client audit。
