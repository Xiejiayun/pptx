# Table Cell Horizontal Alignment Creation Design

## 目标与范围

本小项为 native `slide.addTable()` 的 cell object 增加 `options.align`，让新建的单段落 cell 可以表达 PptxGenJS 4.0.1 的 left、center、right、justify 四种水平对齐。调用方可以继续混用 string、`{ text }` 与已有 border/fill/margin/valign cell object，也可以在同一 cell 上组合这些选项。

水平对齐属于 cell text body 内唯一 paragraph 的 direct `a:p/a:pPr@algn`，不写入 `a:tcPr`，也不计算 master、layout、table style 或其他 effective default。本小项只增加 cell-level 创建输入；table-level `align` 创建与传播、已有 cell 的读取和编辑分别作为后续独立小项，避免把 default precedence 与多段落 existing-deck 语义混入同一 review 单元。

## 公共 API

复用已有公共四值类型，不增加重复 union：

```ts
export type TextAlignment = 'left' | 'center' | 'right' | 'justify';

export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly valign?: TextBoxVerticalAlignment;
}
```

创建字段沿用 PptxGenJS 的 `align` 名称：

```ts
slide.addTable([[
  { text: 'Left', options: { align: 'left' } },
  { text: 'Center', options: { align: 'center' } },
  { text: 'Right', options: { align: 'right' } },
  { text: 'Justify this longer sentence', options: { align: 'justify' } },
]]);
```

`AddTableCell`、`AddTableCellInput`、`SlideModel.addTable()` 与 `TableModel` 的其他签名不变。本小项不伪造尚未存在的 cell horizontal-alignment snapshot 或 editor。

## 方案选择

考虑三个方案：

1. 在 `AddTableCellOptions` 增加 `align`，复用现有 `TextAlignment` normalizer 与 rich-text paragraph renderer，把值作为新建单段落的 default alignment。采用此方案；它没有重复 codec 或 OOXML 映射，且与普通 `addText()` 的创建语义一致。
2. 同时增加 table-level `align` 并把它传播到所有未覆盖 cells。最终价值更高，但会把 cell 严格输入、table default overlay 与 precedence 测试合成一个较大的 review 单元，不采用。
3. 一次增加创建、existing-deck snapshot 与 editor。已有 cell 可能含多个 paragraph 且每段 direct alignment 不同，必须先设计不会丢失信息的读取/编辑值模型；不应让这个问题阻塞无歧义的单段落创建能力，因此不采用。

用户已授权实现方持续自主选择最佳方案并推进全部缺口，因此本设计按推荐方案直接批准实施，完成后继续 table-level horizontal alignment 和 existing-deck editing。

## 严格输入与内部模型

Cell `options` own keys 从 `border` / `fill` / `margin` / `valign` 扩展为 `align` / `border` / `fill` / `margin` / `valign`。现有 descriptor-safe 规则不变：options 必须是 ordinary 或 null-prototype object，只允许 supported own data properties；accessor、inherited、class、array、extra key 与 symbol key 均在 getter-free 状态下拒绝。

`align` 可省略或为 `undefined`。显式值只接受 exact `left`、`center`、`right`、`justify`；null、boolean、number、空字符串、case/whitespace variant、OOXML token `l` / `ctr` / `r` / `just`、`dist`、`thaiDist`、`justLow`、array、object 与 symbol 均抛 `TypeError`。所有归一化在 geometry、rendering 和 package mutation 前完成。

创建内部 cell 扩展为：

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly alignment?: TextAlignment;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}
```

归一化复用 `normalizeTextAlignment(value, context)`。归一化后的 definition 只保留 detached primitive，不保留 caller options object，也不增加 table metadata。

## OOXML 与渲染

公开值映射沿用普通 text paragraph 的 canonical 映射：

| API | OOXML `a:pPr@algn` |
| --- | --- |
| `left` | `l` |
| `center` | `ctr` |
| `right` | `r` |
| `justify` | `just` |

`renderTableCell()` 继续用 `normalizeRichText()` 构造一个 paragraph，只把 `cell.alignment` 传给现有 `renderRichTextParagraphs(..., { defaultAlign })`。不得复制 token map、手工拼接 `pPr`，也不得修改 rich-text renderer 的普通 shape 行为。

Omitted `align`、`align: undefined`、empty cell options、string 与等价 `{ text }` 必须保持现有 byte-identical output。显式 alignment 只向该 paragraph 的 existing canonical `a:pPr` 增加 `algn`；不改变 body properties、list style、run properties、end paragraph properties、cell margins、vertical alignment、border/fill、geometry、row/column sizes 或 relationships。

`TableModel.setCellText()` 只替换目标 `a:t`，因此创建的 `algn` 必须在文字编辑、其他 cell property 编辑、duplicate、write 和 reopen 后保持。后续 horizontal-alignment editor 会拥有 `pPr@algn` 的 direct add/replace/clear 语义，本小项不预先增加该 API。

## PptxGenJS 4.0.1 对等与差异

PptxGenJS 4.0.1 的 `TableCellProps` 与 `TableProps` 都继承 `TextBaseProps.align`。真实 4.0.1 输出确认：

- cell `left`、`center`、`right`、`justify` 分别生成 direct `pPr@algn="l"`、`"ctr"`、`"r"`、`"just"`；
- cell alignment 省略时不写 direct `algn`；
- table-level valid value 在生成时复制到未覆盖 cell，cell-level valid value 覆盖 table value，最终只保留 paragraph direct state；
- runtime 非法值如 `dist` 被静默忽略并生成没有 `algn` 的 paragraph。

本小项对 cell-level 四种合法 final state 与 omission 完全对等。Native 严格拒绝 PptxGenJS 的 invalid runtime fallback，不把拼写错误折叠为无 alignment。Table-level native `align` 尚不在本小项中；adapter 继续忠实读取 PptxGenJS 已 materialize 的 paragraph XML。

Native 新建 table 的 paragraph properties 与 PptxGenJS 仍可包含其他 canonical direct-state 差异；对等断言聚焦 alignment token、cell text、geometry、margins、vertical alignment、border/fill 和最终可见结果，不要求整份 slide XML 字节相同。

## 错误与事务语义

非法 cell `align` 必须在任何 slide bytes、package relationships、ZIP state、live model identity 或 mutation journal 变化前抛 `TypeError`。Accessor-backed options 在不调用 getter 的情况下拒绝。现有 `SlideModel.addTable()` package transaction、矩阵验证顺序和其他 cell option 错误保持不变。

本小项不新增 fallback、coercion、alias 或错误恢复分支。有效 input 只通过现有创建路径产生一次 slide mutation。

## 测试与验收

实现必须覆盖：

1. Internal creation：mixed string/object、omitted/undefined/left/center/right/justify、null-prototype options，以及与 margin/valign/border/fill 组合。
2. Exact output：omitted/undefined/empty options byte-identical；`l/ctr/r/just` token 与 canonical `pPr` child 顺序精确；`tcPr` 不出现 paragraph alignment。
3. Invalid input：options accessor、class/inherited/extra/symbol state 及全部非法 runtime alignment；getter count zero，normalize failure 不改变 package。
4. Public lifecycle：创建、`setCellText()`、cell border/fill/margin/valign/direction/fit 编辑、duplicate isolation、outer rollback、stable identity、geometry 与 write/reopen 保留 `algn`。
5. SDK：Node public input、invalid values/coordinates、non-target part isolation 及已有 table property 互操作。
6. PptxGenJS 4.0.1：omitted/left/center/right/justify final direct alignment、text、geometry、margins、valign、border/fill 对等；invalid fallback 差异明确。
7. Packed Node/browser/declaration/CLI smoke：`AddTableCellOptions.align` 可编译，并在创建、文字编辑、写出和重开后保持正确。
8. 文档：changelog、API README、compatibility baseline 与 package README 更新 supported/unsupported 边界。
9. TypeScript project references、focused/full tests、performance gate、actual tarball 与 staged diff review 全部通过。
10. 真实 native source/edited/reopened 与 PptxGenJS baseline 均通过 PowerPoint 2010 profile zero error/zero warning。
11. Omitted/undefined package diff 为零；single-cell text/property edit 只改变目标 slide；write/reopen 再次写出为零差异。
12. LibreOffice/Poppler 原图确认 left/center/right/justify、border/fill/margin/valign、非等宽列/非等高行，无 repair、裁切、异常换行、重叠或越界；overflow checks 通过。

## 文档与发布边界

文档明确 native cell creation 新增 `align`，合法值及 direct paragraph ownership；省略不实体化 effective left default。兼容矩阵区分 cell-level 创建、尚未完成的 table-level default 与 existing-deck horizontal-alignment editing。

## 非目标

本小项不增加 table-level `align`、`TableCell.horizontalAlignment`、`TableModel.setCellHorizontalAlignment()`、effective inheritance、`dist` / `thaiDist` / `justLow`、多段落/rich cell 创建、text direction/fit creation、hyperlink、merge/span、auto-page、repeated headers、内容测量或 layout recomputation。它不改变已有 cell editor、table geometry、margin/valign/border/fill contract、PptxGenJS adapter 读取或其他 shape paragraph API。
