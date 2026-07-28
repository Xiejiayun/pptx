# Table Cell Text Direction Design

日期：2026-07-28
状态：已批准实施

## 目标与范围

为既有 table cell 增加文字方向的读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TableCellProps.textDirection` 的四个公开值，并正确导入 table-level default 已实体化到各 cell 的输出。该能力映射 direct `a:tc/a:tcPr@vert`，不复用普通文本框的 `a:txBody/a:bodyPr@vert` 存储位置。

本小项只扩展当前 `TableModel` 对既有表格的语义读取和局部编辑。不增加 table 创建 API、table-level default、cell fit/autofit、vertical alignment、margin、fill、border、merge 操作、rich text、row/column mutation、继承解析或普通文本框方向；这些继续按独立小项实现。

## 公共 API

```ts
export type TableCellTextDirection =
  | 'horz'
  | 'vert'
  | 'vert270'
  | 'wordArtVert';

export interface TableCell {
  readonly text: string;
  readonly textDirection?: TableCellTextDirection;
}

export class TableModel {
  readonly rows: readonly TableRow[];

  setCellTextDirection(
    rowIndex: number,
    columnIndex: number,
    value: TableCellTextDirection | undefined,
  ): void;
}
```

用法：

```ts
const table = slide.shapes.find((shape) => shape.kind === 'table') as TableModel;

table.rows[0]?.cells[0]?.textDirection; // direct value or undefined
table.setCellTextDirection(0, 0, 'vert270');
table.setCellTextDirection(0, 0, 'horz');
table.setCellTextDirection(0, 0, undefined); // clear direct override
```

`rows` 仍是 immutable value snapshot，不引入 live `TableCellModel` identity。setter 延续现有 `setCellText(row, column, value)` 的索引式表面；读取 snapshot 与后续重新读取 `rows` 之间不共享可变对象。

## 方案选择

考虑过三个方案：

1. 增加专用 `TableCellTextDirection` 和 `setCellTextDirection()`；采用此方案。它精确匹配 PptxGenJS table-cell 的四个 token，并保持当前 table snapshot + indexed mutation 架构。
2. 直接复用七值 `TextBoxTextDirection`。代码更少，但会错误承诺 `eaVert`、`mongolianVert`、`wordArtVertRtl` 与 PptxGenJS table-cell 全功能对等，也掩盖 `tcPr@vert` 和 `bodyPr@vert` 的不同 ownership。
3. 把每个 cell 升级为 live `TableCellModel`。长期可能适合完整表格 API，但为单一方向字段引入 identity cache、merge 坐标和 row/column 生命周期，超出本小项。

实现层新增窄 `table-cell-text-direction.internal.ts` codec，负责 token validation、strict direct read 和单 attribute patch。`TableModel` 只负责定位 cell、事务与公开 snapshot/mutation；不重构既有 `setCellText()`。

用户已授权实现方持续选择最佳方案并逐项推进，因此本设计按 PptxGenJS 对等、存储边界准确和最小变更定稿。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()` 实测：

- table 和 cell 都使用 `textDirection` 名称，公开 union 是 `horz | vert | vert270 | wordArtVert`。
- table-level 非水平值会在生成阶段复制到每个未覆盖 cell，最终文件没有单独的 table-level direction 状态。
- cell 显式 `horz` 可抑制 table-level 非水平 default，但与无 default 时的 omitted 一样，不写 `tcPr@vert`。
- `vert`、`vert270`、`wordArtVert` 原样写为 direct `a:tcPr@vert`。
- runtime 会把类型外 truthy 字符串（实测 `eaVert`）原样写入；本库不复制该宽松行为，getter 对它返回 `undefined`，原 XML保持不变，setter 严格拒绝。
- direction 不写在 cell `a:txBody/a:bodyPr`，也不是普通 shape `TextPropsOptions.vert`。

adapter 仍通过 PptxGenJS 公开输出导入，不读取 `_slides` 或其他私有字段。导入显式 `horz` 与 omitted 时都只能得到 `undefined`，因为两者 wire form 相同；effective appearance 都是水平文字。

## OOXML 与 direct 语义

```xml
<a:tc>
  <a:txBody>...</a:txBody>
  <a:tcPr marL="91440" marR="91440" vert="vert270">
    ...
  </a:tcPr>
</a:tc>
```

getter 只检查 cell 唯一 direct `tcPr` 上唯一 unqualified `vert` attribute。值完全等于四个公共 token 之一时返回；absent、empty、case/whitespace variant、namespaced attribute、duplicate attribute 或未知 token 返回 `undefined`。不读取 descendant `tcPr`、`bodyPr@vert`、table style、master/layout/theme 或 effective inheritance，只读不产生 mutation。

显式 `horz` 与 absent 在公共编辑 API 中不同：`setCellTextDirection(..., 'horz')` 写 direct `vert="horz"`，重新读取返回 `horz`；`undefined` 删除 direct attribute。该选择与本库普通文本框的可逆 direct editing 一致，同时保持与 PptxGenJS omitted/horz 输出的 effective appearance 对等。

## 无损编辑与错误边界

setter 在 outer OPC transaction 内定位目标 table 的物理 row/cell。row 或 cell 索引不存在时抛 `RangeError`；cell 缺少唯一 direct `tcPr` 时抛 `ModelParseError`，不创建推测结构。merged placeholder cell 只要有合法 direct `tcPr` 就可局部设置方向；本项不把 merge span 折算为逻辑坐标。

更新只增加、替换或删除 direct `tcPr@vert`。既有 attribute 顺序、quote style、margins、anchor、horizontal overflow、fill、border、3D、extension、namespace、unknown child、`txBody`、paragraph/run、merge attribute 和相邻 cell 保持原字节与顺序。existing supported value 重设为相同值是 semantic no-op；unknown direct token 可被显式合法值替换，或由 `undefined` 清除。

`setCellText()`、table/shape transform、slide reorder/duplicate 和所有非 direction 读取或 mutation 保留原 vert。direction setter 不改变 cell text、table relationships、shape id 或 live `TableModel` identity。

输入 value 对 null、boolean、number、空字符串、case/whitespace variant、七值文本框独有 token、array、object 和 symbol 抛 `TypeError`。validation、索引解析和结构检查都必须在 part mutation 前完成；失败及 outer transaction rollback 保持 exact part bytes、mutation journal、slide/table identity 和 live snapshot 状态。

## 测试与发布门禁

1. model fixture 覆盖四个 direct token、absent、empty、case/whitespace、namespaced/duplicate/unknown attribute、bodyPr descendant 和 repeated/missing `tcPr`；getter 不产生 mutation。
2. `setCellTextDirection()` 覆盖 self-closing/expanded `tcPr` 的 add/replace/explicit horz/clear、unknown replacement、merged placeholder、invalid index/structure、snapshot isolation、exact unrelated XML preservation 和 outer rollback。
3. `setCellText()`、transform、duplicate、write/reopen 均保留 directions；只编辑一个 cell 不改变相邻 cell。
4. PptxGenJS 4.0.1 真实输出覆盖 table-level inheritance、per-cell override、omitted、horz、vert、vert270、wordArtVert 和 invalid passthrough；不访问私有字段。
5. API、npm README、changelog 与兼容矩阵标记 table-cell textDirection 已支持，把 table-cell fit 保留为独立 partial 项。
6. packed Node/browser/declaration smoke 覆盖 snapshot、四值编辑与 clear；全仓 typecheck/test、独立 performance 和 tarball CLI 全部通过。
7. 同源 native/hand-patched table 文件通过 PowerPoint 2010 profile 的 0 error / 0 warning 验证且 package diff 为空；LibreOffice 无修复导出，逐 cell 检查 horz/vert/vert270/wordArtVert 的方向与 PptxGenJS 基线一致，两份 overflow 检查通过。
