# Table Cell Margins Design

日期：2026-07-28
状态：已批准实施

## 目标与范围

为既有 table cell 增加四边文字内边距的读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 table/cell `margin` 最终写入的 direct `a:tc/a:tcPr@marL`、`marR`、`marT`、`marB`。该能力只表达文件中的 direct cell margin，不读取或修改普通文本框 `a:bodyPr@lIns/tIns/rIns/bIns`、paragraph margin/indent、table style 或 PowerPoint effective default。

本小项只扩展 `TableModel` 对既有表格的 detached snapshot 与 physical-cell indexed mutation。不增加 table creation、table-level margin default、effective inheritance、row/column sizing、fill、border、merge mutation、rich text、auto paging 或布局计算。已完成的 table-cell text direction、fit 和 vertical alignment 与本字段保持独立。

## 公共 API

复用现有 point-based margin value types，不增加含义重复的 table-only value object：

```ts
export interface TextBoxMargins {
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly left?: number;
}

export type TextBoxMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | TextBoxMargins;

export interface TableCell {
  readonly text: string;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
  readonly margins?: TextBoxMargins;
}

export class TableModel {
  readonly rows: readonly TableRow[];

  setCellMargins(
    rowIndex: number,
    columnIndex: number,
    value: TextBoxMarginInput | undefined,
  ): void;
}
```

用法：

```ts
const table = slide.shapes.find((shape) => shape.kind === 'table') as TableModel;

table.rows[0]?.cells[0]?.margins;        // direct point snapshot or undefined
table.setCellMargins(0, 0, 7.2);        // uniform 0.1 inch equivalent
table.setCellMargins(0, 1, [3.6, 7.2, 10.8, 14.4]);
table.setCellMargins(0, 2, { top: 4, left: 8 });
table.setCellMargins(0, 2, {});          // clear all four direct sides
table.setCellMargins(0, 2, undefined);   // clear all four direct sides
```

所有公共数值使用 point，与现有文本框 margin API、字体和段落度量保持一致。标量应用到四边；tuple 使用 `[top, right, bottom, left]`；具名对象只保留提供的 direct sides。setter 是四个受支持属性的整体替换，`{}` 或 `undefined` 清除四边。`rows` 继续返回 detached value snapshots，外部修改旧对象不影响 model。

## 方案选择

考虑过三个方案：

1. 原样复制 PptxGenJS table-cell `margin` 的 number/tuple 输入和运行时单位分支。迁移表面最像，但 4.0.1 会根据 tuple 第一项是否 `>= 1` 在 inches 与 points 间切换，同一个数值缺少稳定单位，无法作为新的双向 API 契约。
2. 公开 raw EMU 或新的 table-only coordinate type。它最接近 OOXML，却把底层存储泄漏到常用 API，并与现有 point-based `TextBoxMargins` 形成不必要的重复。
3. 复用 `TextBoxMargins` / `TextBoxMarginInput` 的 point value semantics，新增窄 `table-cell-margins.internal.ts` codec，公开 `TableCell.margins` 与 `setCellMargins()`；采用此方案。它保持高层单位稳定，同时让 adapter 忠实读取 PptxGenJS 已经生成的 direct EMU。

不复用 text-box margin codec 的 storage path。两者虽然共享公共 value shape 和 normalization，但 text box 拥有 `bodyPr@*Ins`，cell 拥有 `tcPr@mar*`；读写 ownership、malformed structure 和未知 XML 保留边界不同。

用户已授权实现方持续选择最佳方案并逐项推进，因此本设计按清晰单位、PptxGenJS 输出兼容、OOXML ownership 和最小可验证改动定稿。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()`、安装包类型声明和 PowerPoint 2010 profile validation 实测：

- `TableCellProps.margin` 与 `TableProps.margin` 声明为 cell margin；cell 值覆盖 table-level 值，table-level 值在生成时复制到未覆盖 cells，最终文件没有独立 table-level margin 状态。
- table/cell 都省略 margin 时，每个普通 cell 仍写 direct narrow defaults：top/bottom `0.05in = 45,720 EMU = 3.6pt`，left/right `0.1in = 91,440 EMU = 7.2pt`。
- 显式 scalar `0` 写四个 direct zero；scalar `0.1` 写四个 `91,440` EMU。
- tuple 按 `[top, right, bottom, left]` 映射，例如 `[0.05, 0.1, 0.15, 0.2]` 写 top `45,720`、right `91,440`、bottom `137,160`、left `182,880`。
- 4.0.1 保留旧版本兼容分支：只检查 margin 第一项；第一项 `< 1` 时整组按 inches，第一项 `>= 1` 时整组按 points。因此 scalar `1` 写 `12,700` EMU，而 scalar `0.99` 写 `905,256` EMU；`[1, 2, 3, 4]` 写 1/2/3/4pt。
- runtime 接受负数和类型范围外值，可能生成负 margin 或无效 token；本库 adapter 保留实际 OOXML，strict native setter 不复制 non-finite、结构错误或 Int32 越界行为。
- merged placeholder cells 由 PptxGenJS 生成独立 `<a:tcPr/>`，不继承或复制可见起始 cell 的 margin attributes。本库继续按 physical cells 编辑，不做 logical-span 折算。

adapter conformance 只使用公开输出，不读取 `_slides` 或其他私有字段。导入后 getter 统一把 direct EMU 除以 12,700 返回 point，因此 `0.1in` 输出读取为 `7.2`，而遗留分支的输入 `1` 读取为 `1`。这表达文件事实，不尝试逆推出调用方最初使用的单位。

## OOXML 与 direct 语义

```xml
<a:tc>
  <a:txBody>
    <a:bodyPr/>
    ...
  </a:txBody>
  <a:tcPr
    marL="91440"
    marR="91440"
    marT="45720"
    marB="45720"
    anchor="ctr"
    vert="horz">
    ...
  </a:tcPr>
</a:tc>
```

字段映射：

| 公共字段 | OOXML attribute |
| --- | --- |
| `left` | `tcPr@marL` |
| `right` | `tcPr@marR` |
| `top` | `tcPr@marT` |
| `bottom` | `tcPr@marB` |

point 使用 `Math.round(value * 12700)` 转为 EMU。每个 direct raw attribute 必须是 strict decimal integer 且位于 signed Int32 `-2,147,483,648..2,147,483,647`；显式 zero 保留，absence 才表示没有该 side 的 direct override。

getter 首先要求 cell 恰好一个 direct `tcPr`。对每个 side 独立读取唯一 unqualified attribute：strict integer 且位于 signed Int32 时返回 `raw / 12700` point；absent、empty、decimal、scientific notation、whitespace/case variant、namespaced、duplicate 或越界 state 不产生该 side。其他合法 side 仍可形成 partial snapshot；四边都没有合法 direct value 时返回 `undefined`。它不读取 descendant `tcPr`、`bodyPr` inset、table style、master/layout/theme 或 effective default，只读不产生 mutation。

missing/repeated direct `tcPr` 在 snapshot read 中返回 `undefined`，避免一个 malformed cell 阻断整张 table；setter 对同一结构抛 `ModelParseError`，因为不能安全选择 mutation target。

## 无损编辑与错误边界

setter 在任何 package mutation 前完整 normalize scalar/tuple/object。随后在 OPC transaction 内定位目标 table 的 physical row/cell；row 或 cell 不存在时抛 `RangeError`。selected cell 必须有唯一 direct `tcPr`，且四个受管 unqualified attributes 均不得重复，否则抛 `ModelParseError`。

更新只增加、替换或删除 selected direct `tcPr@marL/marR/marT/marB`：

- scalar/tuple 写四个 canonical integer tokens；具名对象写提供的 sides 并清除其余受管 direct sides。
- `{}` / `undefined` 清除所有四个受管 unqualified attributes。
- exact normalized values 是 semantic no-op，保持 exact bytes 与 mutation journal；非 canonical 但数值相同的合法 integer spelling 也可保留。
- 单个 malformed/unknown direct token 可由合法值替换，或因该 side 缺省而清除。
- 新增多个 attributes 时使用稳定的 `marL`、`marR`、`marT`、`marB` 顺序；既有 attributes 的相对位置与 quote style 由 lossless patch 保留。

既有 anchor、vert、horizontal overflow、fill、border、3D、extension、namespace、unknown attribute/child、cell `bodyPr`、fit、text、paragraph/run、merge state、neighboring cells、relationships 和 `TableModel` identity 保持不变。`setCellText()`、`setCellTextDirection()`、`setCellTextFit()`、`setCellVerticalAlignment()`、transform、duplicate、write/reopen 等非 margin mutation 必须保留 margins。

输入中的每个提供值必须是 finite number，量化后位于 signed Int32；tuple 必须恰好四项；对象只接受 top/right/bottom/left keys。null、boolean、string、错误长度数组、未知 key、non-finite、symbol 和越界值在 mutation 前失败。失败与 outer transaction rollback 必须保持 exact part bytes、mutation journal、slide/table identity 和 fresh rows snapshots。

## 测试与发布门禁

1. model fixture 覆盖完整/partial/zero/negative/fractional/signed-Int32-boundary direct margins，以及 absent、empty、decimal、scientific、whitespace、case、namespaced、duplicate、越界 attributes 和 repeated/missing `tcPr`；getter 不产生 mutation。
2. `setCellMargins()` 覆盖 scalar、TRBL、object、partial replacement、clear、same-value no-op、malformed canonicalization、self-closing/expanded `tcPr`、merged placeholder、invalid index/structure/input 和 outer rollback。
3. `setCellText()`、direction、fit、vertical alignment、transform、duplicate、write/reopen 保留 margins；只编辑一个 physical cell 不改变 neighboring cell、bodyPr 或 relationships。
4. SDK 覆盖 detached snapshots、public transaction、invalid runtime values/coordinates、package isolation 与 stable identity。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted narrow defaults、table-level scalar/tuple、cell scalar/tuple/zero override、第一项 `<1` inches 分支、第一项 `>=1` points 分支和 negative runtime output，证明 margin 只写在 `tcPr`。
6. API、npm README、changelog 与兼容矩阵明确 point-native API、PptxGenJS dual-unit legacy branch、physical index、direct clear semantics，以及 table creation/table-level mutation 仍未支持。
7. packed Node/browser/declaration smoke 覆盖 snapshot、scalar/TRBL/object editing、partial replacement、clear 和 neighbor isolation；typecheck、全仓 tests、独立 performance、actual tarball 与 CLI 全部通过。
8. 同源 native/PptxGenJS table margin 文件通过 PowerPoint 2010 profile 的 0 error / 0 warning validation 和空 package diff；LibreOffice 无修复导出并逐图检查 zero/uniform/asymmetric/clear/neighbor layout，两份 overflow checks 通过。
