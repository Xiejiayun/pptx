# Table Border Creation Design

## 目标与范围

本小项为 native `slide.addTable()` 增加 table-level `border` 创建默认值，使原生 API 能表达 PptxGenJS 4.0.1 `TableProps.border` 的最终 direct-cell 状态。table border 在创建时物化到没有有效 cell-level border 输入的每个物理 cell；cell border 采用 whole-value 优先级。

本小项复用现有 strict `TableCellBorderInput`、cell border renderer、snapshot 和 editor，不增加第二套 PptxGenJS-shaped native 类型。它不增加 table-level getter/editor、diagonal 或 advanced line、border transparency、effective shared-edge/style 解析、horizontal alignment、direction/fit creation、merge、hyperlink、rich cell text、auto-page、repeated headers或 layout recomputation。

用户已授权实现方自主选择最佳小项、设计和实现并持续推进，因此本规格完成内部设计评审后直接进入计划与实施，不设置人工确认停顿。

## 公共 API

`AddTableOptions` 增加：

```ts
export interface AddTableOptions {
  readonly border?: TableCellBorderInput;
}
```

value model 完全复用现有类型：

```ts
export type TableCellBorder =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly width: number;
      readonly style?: 'solid' | 'dash';
    };

export type TableCellBorderInput =
  | TableCellBorder
  | readonly [
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
    ]
  | TableCellBorders;
```

Scalar 应用于四边；tuple 使用 `[top, right, bottom, left]`；named object 只提供指定 sides。line 支持 strict sRGB/theme color、finite `0..1584` point width 和 optional `solid`/`dash` style。`undefined` 表示没有该输入层，不是 explicit none。

不增加 `TableModel.border`、`TableModel.borders` 或 table-level setter。创建后 source layer 不再存在，只有每个 cell 的 direct L/R/T/B state。

## 方案选择

考虑三个方案：

1. **创建期 whole-value overlay，复用 cell renderer（采用）**。严格归一化 table border；只在 `cell.borders === undefined` 时复制到 cell，再由现有 renderer 在 canonical noFill 基线上输出。验证发生在 package mutation 前，PptxGenJS source-layer precedence 与现有 native cell semantics 都保持一致。
2. **按 side 合并 table 与 cell border**。这会让 cell named partial 或含 `undefined` 的 tuple 继承 table 的其他边，改变现有“cell border 在 canonical noFill 上完整渲染”的语义，也偏离 PptxGenJS 任何 truthy cell border 都阻断 table border 的行为，不采用。
3. **保留 table metadata 或先创建后循环调用 `setCellBorders()`**。前者会产生 OOXML 中不存在的 re-inheritance source layer；后者重复 parse/serialize、放大 mutation journal，并绑定到 editor 的 direct-absence 语义，不采用。

## 严格归一化与覆盖顺序

`table-create.internal.ts` 的 table option allowlist 增加 `border`。`readOptions()` 继续只读取 ordinary/null-prototype object 的 own data properties，不调用 getter。随后调用：

```ts
const tableBorders = normalizeTableCellBorders(
  normalizedOptions.border,
  'Table border',
);
```

覆盖规则只有两层，且是 whole-value precedence：

1. table border；
2. cell border。

具体语义：

- table omitted、runtime `undefined`、`{}` 或全 `undefined` tuple：不做 border overlay，保持现有 canonical bytes。
- table scalar、非空 tuple 或非空 named：物化到 `cell.borders === undefined` 的 cells。
- cell scalar、非空 tuple 或非空 named：完整阻断 table border；其未提供 sides 由现有 renderer 写 canonical direct noFill，不从 table 继承。
- cell omitted、runtime `undefined`、empty options、`border: {}` 或全 `undefined` tuple：没有有效 cell border 层，继承 supplied table border。
- 显式 cell scalar none：四边 direct noFill，并完整覆盖 table border。

table partial named 或含 `undefined` 的 tuple 也只定义自己的 sides；在继承 cells 上，renderer 为其余 sides 写 canonical noFill。这样 table 与 cell 的 partial value 仍使用相同现有 value model，但不会引入逐边跨层合并。

归一化结果立即与 caller 脱离。scalar 的四个 sides、tuple/named side、nested color、width 和 style 都由 `normalizeTableCellBorders()` 复制与量化；后续修改 table/cell source border 或 color 不影响定义和 OOXML。

table border overlay 在 cell normalization 后、table fill/margin/valign overlay 前执行。四个能力彼此独立：border 只设置 `cell.borders`，fill 只设置 `cell.fill`，margin 只设置 `cell.margins`，valign 只设置 `cell.verticalAlignment`。

## OOXML 所有权与序列化

不写任何 table-level border OOXML 或 custom metadata。每个 resolved physical cell 继续由 `renderTableCellBorders()` 输出四条 direct sides：

1. `lnL`；
2. `lnR`；
3. `lnT`；
4. `lnB`；
5. optional direct cell fill。

None 写现有 zero-width `noFill` line。Line 写量化后的 width、strict sRGB/theme `solidFill`、optional `prstDash`（`solid` 或 `sysDash`）以及现有 neutral round/head/tail metadata。显式 width `0` 仍是 line，不折叠为 none；omitted style 不写 direct dash token。

table border omitted/runtime undefined/empty 时，所有原有 table creation bytes 必须保持不变。border 与 table fill、margin、valign 同时存在时，start-tag margins、optional anchor、L/R/T/B borders、fill 的 schema 顺序不变。

## 生命周期与编辑语义

`slide.addTable()` 返回的 live `TableModel` 立即通过每个 `TableCell.borders` 暴露物化后的 detached snapshot。duplicate、outer transaction rollback、write/reopen 都只处理 direct cell state。

创建后调用：

```ts
table.setCellBorders(0, 0, undefined);
```

只清除该 cell 的四条 direct border，不重新应用原始 table border。复制 slide 后编辑 source cell 不影响 duplicate；write/reopen 保留各自最终 border matrix。不能从相同 cell borders 反推它们最初来自 table 还是 cell。

## PptxGenJS 4.0.1 对等与差异

PptxGenJS 4.0.1 会把 table-level scalar 或 TRBL border 复制到没有 cell border 的 cells；任何 truthy cell border 都完整覆盖 table border。对 explicit scalar/full TRBL none、sRGB line、scalar zero/positive/fractional width、solid/dash 和 cell override，native 与 PptxGenJS 最终 direct L/R/T/B state、geometry、text、fill、margin 和 valign snapshot 对等。

保留 strict native 差异：

- Native empty table/cell border 与全 `undefined` tuple 表示没有有效 input layer；PptxGenJS `{}` 会物化默认 gray `666666`、1pt、solid，cell `{}` 还会阻断 table default。
- Native line style omitted 不写 direct dash；PptxGenJS omitted/default type 写 direct solid。需要 direct-state 对等时 native 使用 `style: 'solid'`。
- Native scalar 与 tuple width `0` 都保留 zero-width line；PptxGenJS scalar `pt: 0` 保留 zero，但 TRBL item `pt: 0` 因 fallback 变成 1pt。
- Native tuple 必须恰好四项且 dense；PptxGenJS runtime 可接受 short tuple并把缺项实体化为 none。
- Native 额外支持 named sides 和 theme color；PptxGenJS public border 只声明 scalar/TRBL 与 hex color。
- Native 拒绝 coercion、null、non-finite/negative/over-range width、unsupported line type、extra/symbol key 和 malformed object；adapter 继续 losslessly 读取 PptxGenJS 最终 OOXML。

## 错误处理与无 mutation 保证

table border 与 cell border 使用同一 `normalizeTableCellBorders()` contract：

- only ordinary/null-prototype objects and ordinary exact-length tuples；
- own data properties only；
- strict kind、color、width、style；
- no extra/symbol keys；
- no getter invocation；
- deep detached normalized value。

非法 table border 必须在 slide part mutation、shape ID allocation或 transaction write 前失败。package bytes、mutation journal、slides/shapes identity 和已有 table snapshot 保持不变。table option accessor、border side accessor 和 nested color accessor 都必须零调用。

## 测试与验收

实现必须覆盖：

1. Internal normalization：table scalar/TRBL/named、none、sRGB/theme、style omitted/solid/dash、zero/fractional/max width；string/object/empty/undefined cells 继承；cell scalar/TRBL/named/none whole-value override；caller deep detachment。
2. Exact OOXML：L/R/T/B 顺序、noFill/line metadata、width/color/dash、border-before-fill；omitted/undefined/empty table border byte-identical。
3. Invalid input：table option accessor、border/tuple item/side/color accessor、inherited/class/array、extra/symbol key、sparse/wrong-length tuple、missing kind/color/width、invalid style/color/width；getter count zero且无 mutation。
4. Public model lifecycle：typed `AddTableOptions.border`、immediate snapshots、cell clear/replace、duplicate isolation、rollback、write/reopen、stable identity 和 width/height vectors。
5. SDK lifecycle：public package exports、transaction rollback、invalid table/cell border、non-target part isolation 和 no mutation。
6. PptxGenJS 4.0.1：table scalar/TRBL/none/zero/solid/dash 与 cell override final-state parity；empty/default/style omission/TRBL-zero differences有测试和文档。
7. Packed Node/browser/types：真实 tarball 创建、继承、覆盖、清除、reopen；新增 `tableBorderCreation: true`。
8. 文档：changelog、API README、compatibility baseline、package README 标记 table-level border creation supported，table getter/editor 与 advanced borders 仍 unsupported。
9. Full QA：TypeScript、全仓 Vitest、performance、tarball、PowerPoint 2010 validate、package diff、LibreOffice/Poppler、overflow 和逐图检查。

## 提交边界

每个可独立 review 的小项单独 commit + push + fetch，并要求 `origin/main...HEAD = 0 0`：

1. 设计规格；
2. 实现计划；
3. internal normalization/materialization；
4. public API/model lifecycle；
5. SDK；
6. PptxGenJS parity；
7. packed smoke；
8. docs；
9. defect fixes（仅在 QA 发现缺陷时）。

QA-only 成功不创建空 commit。`.pnpm-store/` 永不修改、删除或暂存。

## 非目标

本小项不改变现有 `TableCell.borders` snapshot、cell creation border 在 canonical noFill 基线上的渲染或 `setCellBorders()` whole-replacement 语义。它不支持 table-level border read/edit、diagonal、gradient/pattern line fill、border transparency、advanced dash、compound/cap/join/alignment/arrow/effect、shared-edge precedence、effective table style、horizontal alignment、text direction/fit creation、merge/span、hyperlink、rich text、row insertion/deletion、auto-page/repeated headers、内容测量或 layout recomputation。
