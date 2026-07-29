# Table Fill Creation Design

## 目标与范围

本小项为 native `slide.addTable()` 增加 table-level `fill` 创建默认值，使原生 API 能表达 PptxGenJS 4.0.1 `TableProps.fill` 的最终 direct-cell 状态。table fill 在创建时物化到没有 cell-level fill 的每个物理 cell；cell fill 优先。

本小项复用现有 strict `TableCellFill`、cell fill renderer、snapshot 和 editor，不增加第二套 PptxGenJS-shaped native 类型。它不增加 table-level getter/editor、effective table-style 解析、gradient/pattern/picture/group fill、table border、horizontal alignment、direction/fit creation、merge、hyperlink、rich cell text、auto-page、repeated headers或 layout recomputation。

用户已明确授权实现方自主选择最佳小项和设计并持续推进，因此本规格完成内部设计评审后直接进入计划与实施，不设置人工确认停顿。

## 公共 API

`AddTableOptions` 增加：

```ts
export interface AddTableOptions {
  readonly fill?: TableCellFill;
}
```

value model 完全复用现有类型：

```ts
export type TableCellFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };
```

因此 native table/cell creation、`TableCell.fill` 和 `TableModel.setCellFill()` 使用同一 strict direct-state 语义：

- `{ kind: 'none' }` 写 direct `a:noFill`。
- solid 支持 sRGB/theme color 和 optional finite `0..100` transparency。
- transparency 量化到 `0.001%`；omitted 不写 alpha，explicit zero 写 `alpha=100000`。
- `undefined` 表示没有该输入层，不是 explicit no-fill。
- `{}`、null、array、class、accessor、额外或 symbol key 均非法。

不增加 `TableModel.fill` 或 `setFill()`。创建后 source layer 不再存在，只有每个 cell 的 direct state。

## 方案选择

考虑三个方案：

1. **创建期归一化 overlay，复用 cell renderer（采用）**。在 table option normalizer 中严格读取一次 table fill，把它复制到没有 normalized cell fill 的 cells，继续由现有 `renderTableCellFill()` 输出。验证发生在 package mutation 前，输出模型与 cell editor 一致，改动最小。
2. **保留 table-level metadata 并在 getter/editor 时计算 effective fill**。这会引入当前 OOXML 不存在的 source layer，使清除 cell 后重新继承，违背 direct-state snapshot 与 PptxGenJS 最终物化行为，不采用。
3. **先创建基础表格，再循环调用 `setCellFill()`**。会重复 parse/serialize、放大 mutation journal，并让创建行为依赖 editor insertion path，不采用。

## 严格归一化与覆盖顺序

`table-create.internal.ts` 的 table option allowlist 增加 `fill`。`readOptions()` 继续 descriptor-safe 地复制 own data properties，不调用 getter。随后调用：

```ts
const tableFill = normalizeTableCellFill(
  normalizedOptions.fill,
  'Table fill',
);
```

覆盖规则只有两层：

1. table fill；
2. cell fill。

具体语义：

- table omitted 或 runtime `undefined`：不做 fill overlay，保持原有 bytes。
- table solid：所有没有 cell fill 的物理 cells 得到该 solid direct state。
- table none：所有没有 cell fill 的物理 cells 得到 direct `a:noFill`。
- cell solid 或 none：完整覆盖 table fill。
- cell omitted、runtime `undefined` 或 empty cell options：继承 table fill。
- empty table fill object 非法；fill 没有 margin 风格的 partial/empty 语义。

归一化结果立即与 caller 脱离。solid nested color 和 transparency 已由 `normalizeTableCellFill()` 复制/量化；后续修改 table/cell source fill 或 color 不影响定义和 OOXML。

table fill overlay 在 cell normalization 后、table margin 和 valign overlay 前执行。三个能力互相独立：fill 只设置 `cell.fill`，margin 只设置 `cell.margins`，valign 只设置 `cell.verticalAlignment`。renderer 顺序保持：

1. `marL/marR/marT/marB` attributes；
2. optional `anchor` attribute；
3. L/R/T/B border children；
4. optional direct cell fill。

## 生命周期与编辑语义

`slide.addTable()` 返回的 live `TableModel` 立即通过每个 `TableCell.fill` 暴露物化后的 detached snapshot。duplicate、outer transaction rollback、write/reopen 都只处理 direct cell state。

创建后调用：

```ts
table.setCellFill(0, 0, undefined);
```

只清除该 cell 的 direct fill，不重新应用原始 table fill。复制 slide 后编辑 source cell 不影响 duplicate；write/reopen 保留各自最终 matrix。没有 table-level fill getter/editor，也不尝试从相同 cells 反推原始 default。

## PptxGenJS 4.0.1 对等与差异

PptxGenJS 4.0.1 会把 table-level solid fill 复制到未覆盖 cells，cell solid fill 优先，最终文件不保留独立 table fill metadata。对 sRGB/theme solid、non-zero/fractional/full transparency，native 与 PptxGenJS 最终 `tcPr` fill choice、color、alpha、geometry、text、border、margin 和 valign snapshot 对等。

保留 strict native 差异：

- PptxGenJS `{ type: 'none' }` 折叠为没有 direct fill；native `{ kind: 'none' }` 写 direct `a:noFill`。
- PptxGenJS explicit transparency 0 省略 alpha；native explicit zero 写 `alpha=100000`。
- PptxGenJS deprecated `alpha`、string fill、runtime coercion 和越界 transparency 不进入 native API。

adapter 继续忠实读取最终 OOXML，不逆推输入来自 table 还是 cell。

## 错误处理与无 mutation 保证

table fill 与 cell fill 使用同一 `normalizeTableCellFill()` contract：

- only ordinary/null-prototype objects；
- own data properties only；
- strict kind、color、transparency；
- no extra/symbol keys；
- no getter invocation；
- deep detached normalized value。

非法 table fill 必须在 slide part mutation、shape ID allocation或 transaction write 前失败。package bytes、mutation journal、slides/shapes identity 和已有 table snapshot 保持不变。

## 测试与验收

实现必须覆盖：

1. Internal normalization：table sRGB/theme/none/explicit-zero/fractional/full transparency；string/object/empty/undefined cells 继承；cell solid/none 覆盖；caller deep detachment。
2. Exact OOXML：fill 保持在 margins/anchor/borders 之后；solid/noFill 精确输出；omitted/undefined table fill byte-identical。
3. Invalid input：table option accessor、fill/color accessor、inherited/class/array、extra/symbol key、missing kind/color、非法 color/transparency；getter count zero且无 cell getter side effect。
4. Public model lifecycle：typed `AddTableOptions.fill`、immediate snapshots、cell clear/replace、duplicate isolation、rollback、write/reopen 和 width/height vectors。
5. SDK lifecycle：public package exports、transaction rollback、invalid cell/table fill、stable identities、no mutation。
6. PptxGenJS 4.0.1：table solid sRGB/theme/transparency 与 cell override final-state parity；native direct none和 explicit-zero 差异明确。
7. Packed Node/browser/types：真实 tarball 创建、继承、覆盖、清除、reopen；新增 `tableFillCreation: true`。
8. 文档：changelog、API README、compatibility baseline、package README 标记 table-level fill creation supported，table getter/editor 与 advanced fill 仍 unsupported。
9. Full QA：TypeScript、全仓 Vitest、performance、tarball、PowerPoint 2010 validate、package diff、LibreOffice/Poppler、overflow 和逐图检查。

## 提交边界

按既有用户约束，每个可独立 review 的小项单独 commit + push + fetch，并要求 `origin/main...HEAD = 0 0`：

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

本小项不改变现有 cell fill snapshot/editor 的合法 value 或 OOXML 语义，不支持 table-level fill read/edit、effective table-style fill、gradient/pattern/picture/group fill、table-level border、horizontal alignment、text direction/fit creation、merge/span、hyperlink、rich text、row insertion/deletion、auto-page/repeated headers、内容测量或 layout recomputation。
