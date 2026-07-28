# Table Cell Borders Design

日期：2026-07-28
状态：已批准实施

## 目标与范围

为既有 table cell 增加四个 direct 边框的读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TableCellProps.border` / `TableProps.border` 最终写入的 `a:tc/a:tcPr/a:lnL|a:lnR|a:lnT|a:lnB`。本小项支持各边显式无边框、solid/dash 线型、point 宽度和 sRGB/theme color，并区分 direct no-line 与没有 direct side override。

该能力只表达选中物理 cell 的 left/right/top/bottom direct line。它不解析 table style、conditional banding、相邻 cell 的 shared-edge precedence、theme 解析后的 effective color 或渲染后的最终边框，也不把 diagonal border、cell fill 或 text outline 当作四边边框。

本小项不增加 table creation、table-level border default、diagonal border 编辑、其他 OOXML dash preset、compound/cap/join/alignment/arrow/effect 编辑、border transparency、merge mutation、row/column sizing、rich cell text、auto paging 或布局计算。已完成的 cell text、direction、fit、vertical alignment、margins 和 fill 与 borders 保持独立。

## 公共 API

复用现有 `RichTextColor` 的 strict sRGB/theme value；width 统一使用 point，不公开 EMU 或 raw OOXML token：

```ts
export type TableCellBorderStyle = 'solid' | 'dash';

export type TableCellBorder =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly width: number;
      readonly style?: TableCellBorderStyle;
    };

export interface TableCellBorders {
  readonly top?: TableCellBorder;
  readonly right?: TableCellBorder;
  readonly bottom?: TableCellBorder;
  readonly left?: TableCellBorder;
}

export type TableCellBorderInput =
  | TableCellBorder
  | readonly [
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
    ]
  | TableCellBorders;

export interface TableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}

export class TableModel {
  setCellBorders(
    rowIndex: number,
    columnIndex: number,
    value: TableCellBorderInput | undefined,
  ): void;
}
```

用法：

```ts
table.rows[0]?.cells[0]?.borders; // detached direct-state snapshot or undefined

table.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: 'FF0000' },
  width: 2,
  style: 'solid',
}); // all four sides

table.setCellBorders(0, 1, [
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1, style: 'dash' },
  { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 1.5 },
  { kind: 'none' },
  undefined,
]); // top, right, bottom, left; undefined clears one direct side

table.setCellBorders(0, 2, { top: { kind: 'none' } }); // whole replacement
table.setCellBorders(0, 2, undefined);                 // clear all four direct sides
```

标量 border 应用于四边；tuple 必须恰好四项并使用 `[top, right, bottom, left]`；named object 是四边的 whole replacement，省略的 side 被清除，`{}` 与 `undefined` 都清除四边。snapshot 只包含能够严格读取的 direct sides；没有 supported side 时返回 `undefined`。

`kind: 'none'` 写 direct line element + `noFill`；side omission/`undefined` 删除该 direct line element，两者与 table-style fallback 的语义不同。`style` 缺省表示不写 direct `prstDash`；`solid` 写 `prstDash val="solid"`；`dash` 写 PptxGenJS 对等的 `prstDash val="sysDash"`。width 是 finite `0..1584` points，量化到最近 EMU；零宽 line 仍是 line state，不自动改成 `none`。

`rows` 继续返回 detached nested snapshots。调用方修改旧 borders、某个 side 或嵌套 color 不会修改 model。setter 继续使用零基 physical row/cell index，对 merged placeholder 不做 logical-span 折算。

## 方案选择

考虑过三个方案：

1. 直接复制 PptxGenJS `BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]`。表面迁移最直接，但 `type` 同时承担 none/line style、缺省值在 table-level 与 cell-level 的 runtime 分支不同，也无法表达 direct side absence，因此不适合作为双向编辑契约。
2. 使用 strict `none | line` side value，加 scalar/TRBL/named whole-replacement input；采用此方案。它覆盖 PptxGenJS typed table border 能力，复用本库 color/point 约定，保留 explicit noFill 与 absence，并把 unsupported line details 留在 XML 中。
3. 立即抽象通用 DrawingML line engine，覆盖所有 dash、gradient/pattern line fill、compound/cap/join/arrows/effects 和 diagonals。OOXML 覆盖更广，但会把一个 table-cell 小项扩大为 shape/text/chart 的跨域重构，不符合当前逐项可验证的推进方式。

用户已授权实现方按小项持续选择最佳方案，因此本设计按 strict native API、PptxGenJS 输出兼容、direct-state 可逆和最小可验证改动定稿。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()`、安装包类型声明和实际 slide XML 校准：

- `TableCellProps.border` 与 `TableProps.border` 接受一个 `BorderProps` 或 `[top, right, bottom, left]` tuple；cell-level value 覆盖 table-level value，最终文件中没有独立 table-level border 状态。
- table/cell 都省略 border 时，PptxGenJS 仍给每个 physical cell 的四边写 direct `lnL/lnR/lnT/lnB`，每边是 `w="0"`、neutral line attributes 和 `noFill`。因此 adapter snapshot 是四个 `{ kind: 'none' }`，不是 `undefined`。
- scalar border 被复制到四边；tuple 的 missing/null side 被实体化为 noFill。边框 XML 必须按 L/R/T/B schema 顺序写入，即使 public tuple 是 TRBL。
- `type` 缺省为 `solid`，color 缺省为 `666666`，point width 缺省为 `1`。solid 写 sRGB `solidFill` + `prstDash val="solid"`；dash 写 `prstDash val="sysDash"`。
- visible line 固定写 `cap="flat" cmpd="sng" algn="ctr"`、`round` 和 none/medium head/tail metadata。none 写相同 neutral attributes、`w="0"` 和 `noFill`，不写 dash/join/end metadata。
- cell-level numeric `pt: 0` 被保留并写为 zero-width solid line；table-level tuple normalization 以 truthy fallback 处理 `pt: 0`，会退回默认 1pt。adapter 忠实读取最终 XML，不逆推输入层级。
- `BorderProps.color` 的公开类型是 hex string；native API 额外支持既有 `RichTextColor` theme tokens，以便读取和编辑合法 direct theme border。adapter conformance 不依赖未声明的 runtime theme input。
- runtime 对 negative/non-finite/over-range point value 没有稳定 strict contract。本库 adapter 保留实际 XML，strict getter 不伪造合法 snapshot，native setter 在 mutation 前拒绝。
- merged placeholders 仍是独立 physical cells。本库只按实际 row/cell index 读取与编辑，不合并 shared edge，也不把可见起始 cell 的 border 复制到 placeholder。

adapter 只从公开输出的 OOXML 读取，不访问 `_slides` 等私有字段。它暴露最终 direct sides，不推测 border 最初来自 table-level scalar、TRBL 还是 cell override。

## OOXML 与 strict direct 语义

```xml
<a:tcPr>
  <a:lnL w="12700" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
    <a:prstDash val="solid"/>
    <a:round/>
    <a:headEnd type="none" w="med" len="med"/>
    <a:tailEnd type="none" w="med" len="med"/>
  </a:lnL>
  <a:lnR w="19050" cap="flat" cmpd="sng" algn="ctr">
    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
    <a:prstDash val="sysDash"/>
  </a:lnR>
  <a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT>
  <a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>
</a:tcPr>
```

getter 首先要求 cell 恰好一个 direct `tcPr`，然后独立检查同 prefix 的 direct `lnL`、`lnR`、`lnT`、`lnB`。某一 side missing、repeated、wrong-prefix、unsupported 或 malformed 时只省略该 side；其他合法 side 仍可形成 partial snapshot。四边都没有 supported snapshot 时返回 `undefined`。`lnTlToBr` / `lnBlToTr` 始终不属于此 API。

每个 line element 只接受以下 strict 子集：

- 除 namespace declaration 外，必须恰好一个 unqualified `w`，可选 `cap="flat"`、`cmpd="sng"`、`algn="ctr"` 各一次；unknown、repeated、namespaced substitute 或其他值使该 side unsupported。`w` 必须是 `0..20116800` strict decimal integer，snapshot 以 point 返回。
- line fill choice 必须唯一且使用 `tcPr` 的同一 lexical prefix。`noFill` 必须无非 namespace attribute 和 element child，并且是 line 的唯一 element child，才返回 `{ kind: 'none' }`。
- visible line 必须有唯一 strict `solidFill`，其中恰好一个 sRGB/theme color，不能有 alpha 或其他 transform。sRGB 必须六位 hex并在 snapshot 中大写；theme token 使用现有 `RichTextColor` 集合。
- `prstDash` 可缺省、或唯一 direct value 为 `solid` / `sysDash`，分别映射 omitted / `solid` / `dash`。其他 preset 不被折叠为近似 dash。
- 可选 `round` 必须为空；可选 `headEnd` / `tailEnd` 必须严格为 `type="none" w="med" len="med"` 且为空。额外 join、arrow、fill、dash、effect 或 unknown child 使该 side 返回 `undefined`，避免忽略影响最终线条的状态。

snapshot 读取永不产生 mutation。missing/repeated direct `tcPr` 使整个 borders snapshot 返回 `undefined`；setter 对同一结构抛 `ModelParseError`，因为无法安全选择 mutation target。

## 无损编辑与错误边界

setter 在 package mutation 前完整 normalize public input：

- scalar strict border 复制为四个 detached values；tuple 必须恰好四项；named object 只允许 top/right/bottom/left keys，并以 whole replacement 归一化。
- `kind: 'none'` 不允许额外字段。
- `kind: 'line'` 必须提供 strict `RichTextColor` 和 finite width；sRGB 允许可选 `#` 并规范化为大写，width 量化为 EMU 后必须在 `0..20116800`，style 只允许 omitted/solid/dash。
- null、boolean、string、错误长度/稀疏 tuple、未知 key/kind/style、缺少 color/width、非法 color、non-finite、negative 或 over-range width 在 mutation 前失败。

随后在 OPC transaction 内定位选中 physical cell。row/cell 不存在时抛 `RangeError`。selected direct `tcPr` 必须唯一；每个受管同-prefix side 最多一个，否则抛 `ModelParseError`。四边以一个 whole replacement 操作完成，任何失败和 outer rollback 都保持 atomic。

更新只替换、插入或删除 selected direct `lnL/lnR/lnT/lnB`：

- none 写 PptxGenJS-compatible `w="0"` + neutral attributes + `noFill`；line 写 quantized width、neutral attributes、strict solid color、optional dash 和 neutral join/end metadata；undefined 删除该 direct side。
- supported normalized state 相同时 exact no-op，保留 lowercase sRGB、leading-zero width、quote style、whitespace 和 PptxGenJS neutral metadata。
- malformed/unsupported 但唯一的 existing side 可由显式 supported value 整体替换，或因 desired side undefined 而清除；不局部修补其内部节点。
- 既有 side 尽量原位替换/删除；new sides 按 L/R/T/B schema 顺序插入，并位于 diagonals、cell fill、cell3D 和 `extLst` 之前。当前 `tcPr` lexical prefix 被复用，自闭合 `tcPr` 安全展开。
- wrong-prefix line、diagonals 和其他非受管 children 原样保留。

既有 fill、margins、anchor、vert、overflow、diagonal borders、cell3D、extensions、unknown attributes/children、text body、fit、text、paragraph/run、merge state、neighboring cells、relationships 和 `TableModel` identity 必须保持不变。`setCellText()`、direction、fit、vertical alignment、margins、fill、transform、duplicate、write/reopen 等非 border mutation 必须保留 borders。

失败与 outer transaction rollback 必须保持 exact part bytes、mutation journal、slide/table identity 和 fresh rows snapshots。

## 测试与发布门禁

1. model fixture 覆盖四边 absent/none/solid/dash、sRGB/theme、zero/max/fractional point width、optional/direct solid dash，以及 repeated/wrong-prefix side、unsupported dash/fill/join/arrow、非法 width/color、diagonals 和 repeated/missing `tcPr`；getter 不产生 mutation。
2. `setCellBorders()` 覆盖 scalar、TRBL、partial named object、empty/undefined clear、none、style omitted/solid/dash、width 0/fractional/max、same-value no-op、unsupported/malformed replacement、schema-order insertion、自闭合/expanded `tcPr`、merged placeholder、invalid index/structure/input 和 outer rollback。
3. 交叉能力测试证明 text、direction、fit、vertical alignment、margins、fill、transform、duplicate、write/reopen 保留 borders；只编辑一个 physical cell 不修改 diagonals、cell/text fill、neighbor 或 relationships。
4. SDK 覆盖 detached multi-level snapshots、public transaction、invalid runtime values/coordinates、package isolation 和 stable identity。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted border materialized none、table/cell scalar、TRBL overrides、missing tuple side、default type/color/width、solid/dash、cell zero width、table zero-width fallback 和 malformed runtime output；只读取公开生成结果。
6. API、npm README、changelog 与兼容矩阵明确 TRBL input/LRTB XML order、point range、none/absence、style omission/direct solid、physical index、table-level materialization，并且不声称 table creation、shared-edge/effective style 或 advanced line 已支持。
7. packed Node/browser/declaration smoke 覆盖 snapshot、scalar/TRBL/named input、sRGB/theme、solid/dash/none/clear、neighbor/cross-property isolation；typecheck、全仓 tests、独立 performance、actual tarball 与 CLI 全部通过。
8. 同源 native/baseline 表格文件覆盖 all-side solid、per-side mixed solid/dash/none/clear 和 untouched neighbor；PowerPoint 2010 profile 为 0 error / 0 warning，package diff 为空，LibreOffice 无修复导出，逐图检查线宽/颜色/虚线/四边/文字/填充/邻居，两份 overflow 通过且 raster hash 一致。
