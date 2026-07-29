# Table Cell Vertical Alignment Creation Design

## 目标与范围

本小项为 native `slide.addTable()` 的 cell object 增加 `options.valign`，让新建 cell 可以直接表达 top、middle、bottom 三种垂直对齐。调用方可以继续混用 string、`{ text }` 与已有 border/fill/margin cell object，也可以在同一 cell 上组合 `valign`、margin、border 和 fill。

垂直对齐写入选中物理 cell 的 direct `a:tcPr@anchor`，不写 `a:txBody/a:bodyPr@anchor`，不计算 effective default。本小项只增加 cell-level 创建输入；table-level `valign` 创建与传播作为紧随其后的独立小项，避免把 per-cell 严格归一化和 table default overlay 混入同一 review 单元。

## 公共 API

复用已有公共三值类型，不增加重复 union：

```ts
export type TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom';

export interface AddTableCellOptions {
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly valign?: TextBoxVerticalAlignment;
}
```

创建字段沿用 PptxGenJS 的 `valign` 名称，live snapshot/editor 继续使用清晰的 `verticalAlignment`：

```ts
const table = slide.addTable([[
  { text: 'Top', options: { valign: 'top' } },
  { text: 'Middle', options: { valign: 'middle' } },
  { text: 'Bottom', options: { valign: 'bottom' } },
]]);

table.rows[0]?.cells[0]?.verticalAlignment; // 'top'
table.setCellVerticalAlignment(0, 0, 'bottom');
table.setCellVerticalAlignment(0, 0, undefined); // clear direct anchor
```

`AddTableCell`、`AddTableCellInput`、`SlideModel.addTable()`、`TableCell.verticalAlignment` 与 `TableModel.setCellVerticalAlignment()` 的其余签名不变。

## 方案选择

考虑三个方案：

1. 在 `AddTableCellOptions` 增加 `valign`，复用现有严格三值 normalizer，并由 table-cell vertical-alignment codec 渲染 direct anchor。采用此方案；它与 PptxGenJS 创建命名、native text-box creation/live 命名模式及已有 table-cell editor 完全一致。
2. 使用 `verticalAlignment` 作为 cell 创建字段。名称更完整，但会破坏本库已采用的 `valign` creation / `verticalAlignment` live convention，并提高 PptxGenJS 迁移成本，不采用。
3. 先创建基础 table，再为每个有值的 cell 调用 `setCellVerticalAlignment()`。这会重复解析和序列化 slide、产生多次 mutation，并让创建依赖 editor 生命周期，不采用。

用户已授权实现方持续自主选择最佳方案并推进全部缺口，因此本设计按上述推荐方案直接批准实施。

## 严格输入与内部模型

Cell `options` own keys 从 `border` / `fill` / `margin` 扩展为 `border` / `fill` / `margin` / `valign`。现有 descriptor-safe 规则不变：options 必须是 ordinary 或 null-prototype object，只允许 supported own data properties；accessor、inherited、class、array、extra key 与 symbol key 均在 getter-free 状态下拒绝。

`valign` 可省略或为 `undefined`。显式值只接受 exact `top`、`middle`、`bottom`；null、boolean、number、空字符串、case/whitespace variant、`mid`、`center`、`just`、`dist`、`distributed`、array、object 与 symbol 均抛 `TypeError`。所有归一化在 geometry、rendering 和 package mutation 前完成。

创建内部 cell 扩展为：

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}
```

字符串是 immutable primitive，不需要额外 detached copy。归一化后的 definition 不保留 caller options object。

## OOXML 与渲染顺序

在 `table-cell-vertical-alignment.internal.ts` 增加 creation renderer：

```ts
export function renderTableCellVerticalAlignmentAttribute(
  value: TextBoxVerticalAlignment | undefined,
): string;
```

映射固定为：

- `undefined` → 空字符串，不写 direct anchor；
- `top` → ` anchor="t"`；
- `middle` → ` anchor="ctr"`；
- `bottom` → ` anchor="b"`。

Table renderer 先写 canonical `marL/marR/marT/marB`，再写 optional `anchor`，随后按 L/R/T/B 输出 border children，最后输出 optional fill：

```xml
<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="ctr">
  <a:lnL>...</a:lnL>
  <a:lnR>...</a:lnR>
  <a:lnT>...</a:lnT>
  <a:lnB>...</a:lnB>
  <a:solidFill>...</a:solidFill>
</a:tcPr>
```

Omitted `valign`、`valign: undefined`、empty cell options、string 与等价 `{ text }` 必须保持现有 byte-identical output。显式对齐只增加一个 `tcPr` attribute，不改变 text body、margins、border/fill、geometry、row/column sizes 或 relationships。

新建 table 的 `TableModel.rows` 立即从 direct anchor 读回 `verticalAlignment`，并可继续通过 `setCellVerticalAlignment()` add/replace/clear。Editor 仍保持 whole direct-attribute ownership、same-value no-op、malformed structure protection、transaction rollback 和 physical zero-based indices。

## PptxGenJS 4.0.1 对等与差异

PptxGenJS 4.0.1 的 `TableCellProps` 与 `TableProps` 都继承 `TextBaseProps.valign`。公开输出已经验证：

- cell `top`、`middle`、`bottom` 分别生成 direct `tcPr@anchor="t"`、`"ctr"`、`"b"`；
- table/cell 都省略时不写 direct anchor；
- table-level valid value 在生成时复制到未覆盖 cell，cell-level value 覆盖 table-level value；最终文件只保留每个 cell 的 direct state；
- runtime 透传类型外 `mid`、`distributed` 会原样生成 invalid/unsupported anchor token。

本小项对 cell-level final state 完全对等：省略保持 direct absence，三种合法值生成相同 anchor，并通过 adapter/reopen 暴露相同 snapshot。Native 严格拒绝 PptxGenJS 的 invalid runtime passthrough，不复制静默写出缺陷。

Table-level native `valign` 尚不在本小项中；它会在下一独立小项把 valid default overlay 到未显式提供 cell value 的 normalized cells。Adapter 继续忠实读取已经 materialize 的 PptxGenJS table-level 结果。

## 测试与验收

实现必须覆盖：

1. Internal creation：mixed string/object、omitted/undefined/top/middle/bottom、null-prototype options，以及与 margin/border/fill 组合。
2. Exact output：omitted/undefined/empty options byte-identical；`t/ctr/b` token 与 margins → anchor → borders/fill 顺序精确。
3. Invalid input：options accessor、class/inherited/extra/symbol state 及全部非法 valign runtime values；getter count zero，normalize failure 不改变 package。
4. Public model：创建后 snapshot、immediate add/replace/clear、same-value no-op、duplicate isolation、outer rollback、stable identity、geometry 与 write/reopen。
5. SDK：Node public lifecycle、invalid values/coordinates、non-target part isolation 及已有 margin/border/fill/direction/fit/alignment editor 互操作。
6. PptxGenJS 4.0.1：omitted/top/middle/bottom final direct anchor、snapshot、text、geometry、margins、border/fill 对等；invalid passthrough 差异明确。
7. Packed Node/browser/declaration/CLI smoke：`AddTableCellOptions.valign` 可编译，并在创建、编辑、写出、重开后保持正确。
8. 文档：changelog、API README、compatibility baseline 与 package README 更新 supported/unsupported 边界。
9. TypeScript project references、focused/full tests、performance gate、actual tarball 与 staged diff review 全部通过。
10. 真实 native source/edited/reopened 与 PptxGenJS baseline 均通过 PowerPoint 2010 profile zero error/zero warning。
11. Omitted/undefined package diff 为零；single-cell valign edit 只改变目标 slide；write/reopen 再次写出为零差异。
12. LibreOffice/Poppler 原图确认 top/middle/bottom 位置、border/fill/margin、非等宽列/非等高行，无 repair、裁切、异常换行、重叠或越界；overflow checks 通过。

## 文档与发布边界

文档明确 native cell creation 新增 point-independent `valign`，合法值及 direct `tcPr@anchor` ownership；省略不实体化 effective top default；创建字段与 live property 分别为 `valign` / `verticalAlignment`。兼容矩阵区分 cell-level 创建、已有 direct 编辑与尚未完成的 native table-level default。

## 非目标

本小项不增加 table-level `valign`、effective inheritance、`just`/`dist`、`anchorCtr`、horizontal paragraph alignment、text direction/fit creation、hyperlink、merge/span、rich text、auto-page、repeated headers 或内容测量。它不改变已有 editor、table geometry、margin/border/fill contract、PptxGenJS adapter 读取、existing-deck lossless editing 或其他 cell options。
