# Table Cell Text Direction Creation Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 native `slide.addTable()` 的 cell object 增加 `options.textDirection`，支持 PptxGenJS 4.0.1 table-cell 的 `horz`、`vert`、`vert270`、`wordArtVert` 四个公开值。调用方可以在同一严格单段落 cell 上组合现有 `align`、`border`、`fill`、`margin`、`valign` 与新增方向。

方向属于物理 cell 唯一 direct `a:tcPr@vert`，不写普通文本框使用的 `a:txBody/a:bodyPr@vert`。本小项只增加 cell-level 创建输入；table-level `textDirection` 的 default propagation 紧随其后独立实现，以保持 review、precedence 和 conformance 边界清晰。

## 公共 API

复用已有专用四值类型：

```ts
export type TableCellTextDirection =
  | 'horz'
  | 'vert'
  | 'vert270'
  | 'wordArtVert';

export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}
```

创建字段与 PptxGenJS、现有 immutable snapshot 都使用 `textDirection`：

```ts
const table = slide.addTable([[
  { text: 'Horizontal', options: { textDirection: 'horz' } },
  { text: 'Vertical', options: { textDirection: 'vert' } },
  { text: 'Rotate 270', options: { textDirection: 'vert270' } },
  { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
]]);

table.rows[0]?.cells[1]?.textDirection; // 'vert'
table.setCellTextDirection(0, 1, 'horz'); // existing direct editor
```

`AddTableCell`、`AddTableCellInput`、`SlideModel.addTable()`、`TableCell.textDirection` 与 `TableModel.setCellTextDirection()` 的其他签名不变。

## 方案选择

考虑三个方案：

1. 在 `AddTableCellOptions` 增加 `textDirection`，复用现有专用 normalizer，并增加只负责创建期 attribute 的窄 renderer。采用此方案；公开命名、严格值域和 OOXML ownership 与 PptxGenJS 及既有 editor 一致。
2. 复用普通文本框的 `vert` 字段与七值 `TextBoxTextDirection`。不采用；它会错误承诺 `eaVert`、`mongolianVert`、`wordArtVertRtl`，并混淆 `bodyPr@vert` 与 `tcPr@vert`。
3. 在创建出基础 table 后逐 cell 调用 `setCellTextDirection()`。不采用；它重复解析和序列化 slide、产生多次 mutation，并无法在后续 table-level default 中保留显式 `horz` 的 precedence 意图。

用户已授权实现方自行选择最佳方案、持续推进且不再等待常规确认，因此本设计按方案 1 直接批准实施。

## 严格输入与归一化

Cell `options` own keys 扩展为 `align`、`border`、`fill`、`margin`、`textDirection`、`valign`。现有 descriptor-safe 规则保持：options 必须是 ordinary 或 null-prototype object，只接受 supported own data properties；accessor、inherited、class、array、extra key 与 symbol key 在 getter-free 状态下拒绝。

`textDirection` 可省略或为 `undefined`。显式值只接受 exact `horz`、`vert`、`vert270`、`wordArtVert`；null、boolean、number、空字符串、case/whitespace variant、普通文本框独有 token、array、object 和 symbol 均抛 `TypeError`。归一化复用 `normalizeTableCellTextDirection()`，并在 geometry、rendering 与 package mutation 前完成。

创建内部 cell 增加：

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly alignment?: TextAlignment;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly textDirection?: TableCellTextDirection;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}
```

显式 `horz` 必须作为归一化字段保留，不能在 normalization 阶段折叠为 `undefined`。后续 table-level default 将据此识别 cell override；只有最终 renderer 按 PptxGenJS wire semantics 省略 horizontal attribute。

## OOXML 与创建语义

在 `table-cell-text-direction.internal.ts` 增加：

```ts
export function renderTableCellTextDirectionAttribute(
  value: TableCellTextDirection | undefined,
): string;
```

映射固定为：

| API value | 创建期 direct `tcPr@vert` |
| --- | --- |
| omitted / `undefined` | omitted |
| `horz` | omitted |
| `vert` | `vert="vert"` |
| `vert270` | `vert="vert270"` |
| `wordArtVert` | `vert="wordArtVert"` |

创建 renderer 先写 canonical `marL/marR/marT/marB`，再写 optional `anchor`，然后写 optional `vert`，随后输出 L/R/T/B borders 与 fill：

```xml
<a:tcPr
  marL="91440" marR="91440" marT="45720" marB="45720"
  anchor="ctr" vert="vert270">
  <a:lnL>...</a:lnL>
  <a:lnR>...</a:lnR>
  <a:lnT>...</a:lnT>
  <a:lnB>...</a:lnB>
  <a:solidFill>...</a:solidFill>
</a:tcPr>
```

Omitted、runtime-`undefined`、string、`{ text }`、empty options 与显式 `horz` 在没有 table-level default 时生成相同 direct direction bytes。方向不改变 text body、paragraph alignment、margins、vertical alignment、border/fill、geometry、row/column sizes、relationships 或 shape identity。

创建后的 live snapshot 对三个非水平值立即返回 exact token。显式 `horz` 因 final OOXML 与 omission 相同而读取为 `undefined`；这与导入 PptxGenJS 输出一致。已有 `setCellTextDirection(..., 'horz')` 继续写显式 `vert="horz"`，用于可逆 direct-state 编辑；creation conformance 与 lossless direct editing 的差异必须在 API 文档中说明。

## PptxGenJS 4.0.1 对等

公开 `addTable()` / `write()` 基线已经证明：

- cell `vert`、`vert270`、`wordArtVert` 原样写入 direct `tcPr@vert`；
- omitted 与显式 `horz` 在无 table default 时都不写 direct `vert`；
- table-level 非水平值会物化到未覆盖 cells；cell 显式 `horz` 抑制该 default，但仍不写 direct `vert`；
- runtime 类型外 truthy token 可被原样写出，native API 不复制该宽松行为。

本小项要求 cell-level supported final state、omission、explicit-horz collapse、text、geometry 和其他已支持 cell options 与 PptxGenJS 4.0.1 对等。Adapter 仍只消费公开输出，不访问私有字段。Table-level propagation 在下一小项覆盖 inherited/default/override 矩阵。

## 错误、事务与生命周期

非法方向必须在任何 slide bytes、relationships、ZIP state、live identity 或 mutation journal 改变前抛错。Accessor-backed property 不得执行 getter。现有 `SlideModel.addTable()` transaction、matrix validation 和 geometry validation 顺序不改变。

创建的 direct direction 必须在 `setCellText()`、horizontal/vertical alignment、margin、border、fill、fit、column/row sizing、transform、slide duplicate、outer rollback、write 与 reopen 中保持。只修改 clone 不得改变 source；失败与 rollback 恢复 exact bytes 和 fresh snapshots。

## 测试与验收

1. Internal normalization 覆盖 string/object、omitted/undefined、四值、null-prototype、组合 options、detachment 与 explicit-horz marker。
2. Exact rendering 覆盖 omission/undefined/horz byte identity、三个 direct token、margins → anchor → vert → borders/fill 顺序，以及 `bodyPr@vert` absence。
3. Invalid input 覆盖 accessor、inherited/class/extra/symbol options 及所有非法 runtime token；getter count 为零且 package zero mutation。
4. Model/SDK lifecycle 覆盖即时 snapshot、existing editor interoperability、same-value/no-op、cell property mutations、duplicate isolation、rollback、stable identity、write/reopen 与 non-target part isolation。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted、horz、vert、vert270、wordArtVert 和 invalid passthrough 差异；只读取公开生成 bytes。
6. Packed Node/browser/declaration smoke 覆盖 typed creation、snapshot、editor、write/reopen，CLI 保持 `0.1.0`。
7. 文档更新 changelog、API README、compatibility baseline 与 package README，明确 cell-level creation supported、explicit horz collapse、direct editor distinction和 table-level pending。
8. TypeScript project references、focused/full Vitest、performance、actual tarball、PowerPoint 2010 profile、package diff、LibreOffice/Poppler 原图 inspection 与 overflow checks 全部通过。

## 非目标与下一项

本小项不增加 table-level `textDirection`、普通文本框七值方向、fit creation、rich/multi-paragraph table cells、hyperlink、merge/span、row insertion/deletion、table styles、auto-page、repeated headers、content measurement 或 layout recomputation。

下一独立小项为 `AddTableOptions.textDirection`：把四值 default 物化到没有 cell override 的 physical cells，并确保显式 cell `horz` 抑制非水平 table default 而最终仍省略 direct attribute。
