# Table Cell Vertical Alignment Design

日期：2026-07-28
状态：已批准实施

## 目标与范围

为既有 table cell 增加文字垂直对齐的读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 table/cell `valign` 的 `top`、`middle`、`bottom` 三个公开值。该能力映射选中物理 cell 的 direct `a:tc/a:tcPr@anchor`，不复用普通文本框的 `a:txBody/a:bodyPr@anchor` 存储位置。

本小项只扩展 `TableModel` 对既有表格的 immutable snapshot 和 indexed mutation。不增加 table creation、table-level default、effective inheritance、distributed/justified anchor、margin、fill、border、merge mutation、rich text、row/column mutation 或动态布局。已完成的 table-cell `textDirection`、`textFit` 与本字段保持独立。

## 公共 API

复用现有三值类型，不增加重复 union：

```ts
export type TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom';

export interface TableCell {
  readonly text: string;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}

export class TableModel {
  readonly rows: readonly TableRow[];

  setCellVerticalAlignment(
    rowIndex: number,
    columnIndex: number,
    value: TextBoxVerticalAlignment | undefined,
  ): void;
}
```

用法：

```ts
const table = slide.shapes.find((shape) => shape.kind === 'table') as TableModel;

table.rows[0]?.cells[0]?.verticalAlignment; // direct value or undefined
table.setCellVerticalAlignment(0, 0, 'top');
table.setCellVerticalAlignment(0, 0, 'middle');
table.setCellVerticalAlignment(0, 0, 'bottom');
table.setCellVerticalAlignment(0, 0, undefined); // clear direct anchor
```

`rows` 继续返回 detached immutable value snapshots，不引入 live `TableCellModel`。setter 延续 `setCellText()`、`setCellTextDirection()` 和 `setCellTextFit()` 的 physical zero-based row/cell 索引表面。

## 方案选择

考虑过三个方案：

1. 直接对 cell 复用 text-box vertical-alignment codec。公共类型正确，但 codec 会修改 `bodyPr@anchor`，与 table cell 的真实 ownership 不符，因此不采用。
2. 新增 `TableCellVerticalAlignment` 类型。可以强调 cell 边界，但三值与 `TextBoxVerticalAlignment` 完全相同，会制造重复公共 union 和转换逻辑。
3. 复用 `TextBoxVerticalAlignment` 公共类型，新增窄 `table-cell-vertical-alignment.internal.ts` codec，公开 `TableCell.verticalAlignment` 和 `setCellVerticalAlignment()`；采用此方案。它保留统一的 top/middle/bottom 语义，同时在实现层严格区分 `tcPr@anchor` 与 `bodyPr@anchor`。

不抽象通用 arbitrary-attribute patcher；现有 table-cell direction codec 已证明专用小 codec 更容易保持 strict read、错误消息和无损边界。`TableModel` 只负责定位、事务和公开 snapshot/mutation。

用户已授权实现方持续选择最佳方案并逐项推进，因此本设计按 PptxGenJS 对等、OOXML ownership 和最小可验证改动定稿。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()` 和安装包类型声明实测：

- `TableCellProps` 与 `TableProps` 都继承 `TextBaseProps.valign?: 'top' | 'middle' | 'bottom'`。
- valid `top`、`middle`、`bottom` 分别写为 direct `tcPr@anchor="t"`、`"ctr"`、`"b"`；cell `bodyPr` 不写 anchor。
- table-level 值在生成时复制到每个未覆盖 cell，cell-level 值覆盖 table-level 值，最终文件没有单独可编辑的 table-level valign 状态。
- table 和 cell 都省略 `valign` 时，实际 writer 仍在每个 cell 实体化 `anchor="ctr"`，因此导入值是 direct `middle`。这与共享类型注释中的 table-cell default top 不一致，本库以真实公开输出为 conformance 基线。
- runtime 透传类型外 `mid`、`distributed` 会原样生成 `anchor="mid"` / `"distributed"`。本库 getter 对这些 token 返回 `undefined` 并在无关 mutation 中保留，setter 严格拒绝。

adapter conformance 只使用公开输出，不读取 `_slides` 或其他私有字段。PptxGenJS omitted 与 explicit middle 都导入为 `middle`，因为它们具有相同 wire form；本库 `undefined` 表示清除 direct anchor，而不是请求 PptxGenJS 的生成默认值。

## OOXML 与 direct 语义

```xml
<a:tc>
  <a:txBody>
    <a:bodyPr/>
    ...
  </a:txBody>
  <a:tcPr marL="91440" marR="91440" anchor="ctr">
    ...
  </a:tcPr>
</a:tc>
```

getter 只检查 cell 唯一 direct `tcPr` 上唯一 unqualified `anchor` attribute。wire 值完全等于 `t`、`ctr`、`b` 时分别返回 `top`、`middle`、`bottom`；absent、empty、case/whitespace variant、long-form `top/middle/bottom`、`just`、`dist`、namespaced/duplicate attribute 或未知 token 返回 `undefined`。不读取 descendant `tcPr`、`bodyPr@anchor`、table style、master/layout/theme 或 effective default，只读不产生 mutation。

missing/repeated direct `tcPr` 在 snapshot read 中返回 `undefined`，避免一个 malformed cell 阻断整张 table；setter 对同一结构抛 `ModelParseError`，因为不能安全选择 mutation target。

## 无损编辑与错误边界

setter 在 OPC transaction 内定位目标 table 的物理 row/cell。row 或 cell 索引不存在时抛 `RangeError`；cell 缺少唯一 direct `tcPr` 或包含重复 unqualified `anchor` attribute 时抛 `ModelParseError`。merged placeholder 只要有合法 direct `tcPr` 就可编辑；本项不把 merge span 折算为逻辑坐标。

更新只增加、替换或删除 selected direct `tcPr@anchor`：

- `top` / `middle` / `bottom` canonicalize 为 `t` / `ctr` / `b`。
- `undefined` 删除 direct unqualified `anchor`。
- same canonical value 是 semantic no-op，保持 exact bytes 与 mutation journal。
- 单个 unknown direct token 可由合法值 canonicalize，或由 `undefined` 清除。

既有 attribute 顺序、quote style、margins、`vert`、horizontal overflow、fill、border、3D、extension、namespace、unknown child、cell `bodyPr`、fit、text、paragraph/run、merge state、neighboring cells、relationships 和 `TableModel` identity 保持不变。`setCellText()`、`setCellTextDirection()`、`setCellTextFit()`、transform、duplicate、write/reopen 等非 alignment mutation 必须保留 anchor。

输入 value 对 null、boolean、number、空字符串、case/whitespace variant、`mid`、`center`、`just`、`dist`、array、object 和 symbol 抛 `TypeError`。validation、索引和结构检查均在 part mutation 前完成；失败与 outer transaction rollback 保持 exact part bytes、mutation journal、slide/table identity 和 fresh rows snapshots。

## 测试与发布门禁

1. model fixture 覆盖 `t/ctr/b`、absent、empty、case/whitespace、long-form/just/dist/unknown、namespaced/duplicate attribute、bodyPr descendant 和 repeated/missing `tcPr`；getter 不产生 mutation。
2. `setCellVerticalAlignment()` 覆盖 self-closing/expanded `tcPr` 的 add/replace/top/middle/bottom/clear、same-mode no-op、unknown canonicalization、merged placeholder、invalid index/structure 和 outer rollback。
3. `setCellText()`、`setCellTextDirection()`、`setCellTextFit()`、transform、duplicate、write/reopen 保留 vertical alignment；只编辑一个 cell 不改变相邻 cell、bodyPr 或 relationships。
4. SDK 覆盖 immutable snapshots、public transaction、invalid runtime values/coordinates、package isolation 与 stable identity。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted default middle、table-level top/middle/bottom、per-cell override、valid values 和 invalid passthrough，证明 anchor 只在 `tcPr`。
6. API、npm README、changelog 与兼容矩阵明确 direct 编辑、physical index、PptxGenJS default materialization，以及 table creation/table-level mutation 仍未支持。
7. packed Node/browser/declaration smoke 覆盖 snapshot、三值编辑和 clear；typecheck、全仓 tests、独立 performance、actual tarball 与 CLI 全部通过。
8. 同源 native/hand-patched table 文件通过 PowerPoint 2010 profile 的 0 error / 0 warning 验证且 package diff 为空；LibreOffice 无修复导出并逐图检查 top/middle/bottom 位置，两份 overflow checks 通过。
