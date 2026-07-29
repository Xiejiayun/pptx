# Table Text Direction Creation Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 native `slide.addTable()` 增加 table-level `textDirection` 创建输入。调用方可以用一个严格的 `horz`、`vert`、`vert270` 或 `wordArtVert` 默认值覆盖没有 cell-level override 的所有 physical cells；合法 cell `options.textDirection` 始终优先。

该 default 只参与创建期归一化。最终状态实体化为各 physical cell 唯一 direct `a:tcPr@vert`，不写 `a:bodyPr@vert`，也不保留 table metadata。这样既匹配 PptxGenJS 4.0.1 的 supported final state，也保持现有 direct snapshot/editor 的无损边界。

本小项不增加 table-level getter/editor。创建完成后只存在每个 cell 的 direct state；后续清除一个 cell 的方向不会重新继承原始 table default。

## 公共 API

`AddTableOptions` 增加：

```ts
export interface AddTableOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
  // existing geometry and name fields remain unchanged
}
```

复用已有专用四值类型：

```ts
export type TableCellTextDirection =
  | 'horz'
  | 'vert'
  | 'vert270'
  | 'wordArtVert';
```

示例：

```ts
const table = slide.addTable([[
  'Inherits vertical',
  { text: 'Blocks with horizontal', options: { textDirection: 'horz' } },
  { text: 'Overrides with stacked', options: { textDirection: 'wordArtVert' } },
]], {
  textDirection: 'vert270',
});

table.rows[0]?.cells.map(({ textDirection }) => textDirection);
// ['vert270', undefined, 'wordArtVert']
```

第二个 snapshot 为 `undefined`，因为创建期 explicit `horz` 按 PptxGenJS wire semantics 阻断 table default，但最终仍省略 direct attribute。已有 `setCellTextDirection(..., 'horz')` 继续显式写 `vert="horz"`；创建期 collapse 与 direct editor 的可逆状态语义保持区分。

## 方案选择

考虑三个方案：

1. 在 `normalizeTableDefinition()` 中验证 table value，再把 resolved value 实体化到没有 cell marker 的 normalized cells。采用此方案；它复用现有 default propagation 模式，只序列化一次，并能让 explicit cell `horz` 正确阻断非水平 default。
2. 创建基础 table 后逐 cell 调用 `setCellTextDirection()`。不采用；它会重复解析和序列化 slide，而且 editor 会把 resolved `horz` 显式写到 XML，无法匹配创建期 final state。
3. 把 default 存入 table extension，并让 getter 计算 effective value。不采用；PptxGenJS 不保留这种 metadata，它会改变后续 clear 语义并扩大 OOXML ownership。

用户已授权实现方自行选择最佳方案、持续推进且不再等待常规确认，因此本设计按方案 1 直接批准实施。

## 严格输入与归一化

Table options own keys 增加 `textDirection`。现有 descriptor-safe 规则保持：options 必须是 ordinary 或 null-prototype object，只接受 supported own data properties；accessor、inherited、class、array、extra key 与 symbol key 在 getter-free 状态下拒绝。

Table `textDirection` 可省略或为 `undefined`。显式值只接受 exact `horz`、`vert`、`vert270`、`wordArtVert`；null、boolean、number、空字符串、case/whitespace variant、普通文本框独有 token、array、object 和 symbol 均抛 `TypeError`。归一化复用 `normalizeTableCellTextDirection()`，并在 name、geometry、rendering 与 package mutation 前完成。

Cell 已在前一小项中保留 explicit `horz` marker。因此 precedence 直接基于 normalized field presence：

| table input | cell input | resolved normalized cell | direct `tcPr@vert` |
| --- | --- | --- | --- |
| omitted / `undefined` | omitted / `undefined` | absent | absent |
| `horz` | omitted / `undefined` | `horz` | absent |
| `vert` / `vert270` / `wordArtVert` | omitted / `undefined` | table value | exact table token |
| any valid value | `horz` | `horz` | absent |
| any valid value | non-horizontal valid value | cell value | exact cell token |

A string cell、`{ text }`、empty options 与 runtime-`undefined` cell value 都视为没有 override。归一化通过 immutable row/cell copies 实体化 final values，不修改 caller rows、cells、options 或 table options。

## OOXML ownership 与序列化

不增加新的 table-level OOXML property。现有 cell renderer 接收 resolved `textDirection`，继续按以下顺序输出：

1. `marL`、`marR`、`marT`、`marB` attributes；
2. optional `anchor` attribute；
3. optional non-horizontal `vert` attribute；
4. `lnL`、`lnR`、`lnT`、`lnB` children；
5. optional cell fill child。

例如 table `vert270` 与 cell middle valign 的 resolved cell 为：

```xml
<a:tcPr
  marL="91440" marR="91440" marT="45720" marB="45720"
  anchor="ctr" vert="vert270">
  <a:lnL>...</a:lnL>
  <a:lnR>...</a:lnR>
  <a:lnT>...</a:lnT>
  <a:lnB>...</a:lnB>
</a:tcPr>
```

Table omitted、runtime-`undefined` 与 table explicit `horz` 在相同 rows/options 下生成 byte-identical native OOXML。方向不改变 text、paragraph properties、margins、vertical alignment、borders、fill、geometry、row/column sizes、relationships 或 shape identity。

## PptxGenJS 4.0.1 对等

已有 public-output baseline 已证明：

- table-level 非水平值实体化到 cell-level omitted/undefined 的 cells；
- cell explicit `horz` 阻断非水平 table default，但不写 direct `vert`；
- cell `vert`、`vert270`、`wordArtVert` 覆盖 table value；
- table `horz` 的最终 direct state 与 omission 相同；
- runtime 类型外 token 可以被 PptxGenJS 原样写出。

Native API 对 supported values、precedence、omission、text、geometry 与其他已支持 cell options 的最终 direct state 保持对等。Native 不复制非法 token passthrough：未知 table value 在任何 mutation 前严格拒绝。Adapter 仍只消费 PptxGenJS 公开 `write()` 输出，不读取私有字段，并对未知 imported token 暴露 `undefined` snapshot、原样保留 XML。

## 错误、事务与生命周期

非法 table direction 必须在任何 slide bytes、relationships、ZIP state、live identity 或 mutation journal 改变前抛错。Accessor-backed property 不得执行 getter。现有 `SlideModel.addTable()` transaction、cell-first validation 与 geometry validation 边界不改变。

创建后，每个 cell 只保留 materialized direct state：

- `TableCell.textDirection` 立即读取最终 direct token；
- `setCellTextDirection(..., undefined)` 只清除该 cell，不重新继承；
- `setCellTextDirection(..., 'horz')` 继续写 explicit direct horizontal token；
- text/alignment/margin/border/fill/fit/size/transform mutations 保持方向；
- duplicate、outer rollback、write 与 reopen 保持 final state；
- clone edits 不影响 source，失败恢复 exact bytes 与 fresh snapshots。

## 测试与验收

1. Internal normalization 覆盖 table 四值、omitted/undefined、null-prototype、detachment，以及 string/object/empty/undefined cell inheritance。
2. Precedence matrix 覆盖每个 table value × cell omitted/undefined/四值；explicit cell `horz` 必须阻断 default 且 final attribute absent。
3. Exact rendering 覆盖 table omitted/undefined/horz byte identity、三个 non-horizontal tokens、margins → anchor → vert → borders/fill 顺序，以及 `bodyPr@vert` absence。
4. Invalid table options 覆盖 accessor、inherited/class/extra/symbol options 与全部非法 runtime tokens；getter count 为零且 package zero mutation。
5. Model/SDK lifecycle 覆盖即时 snapshot、cell editor/clear、不重新继承、unrelated cell mutations、duplicate isolation、rollback、stable identity、write/reopen 与 non-target part isolation。
6. PptxGenJS 4.0.1 conformance 覆盖 supported final-state equality、table/cell precedence 与 invalid passthrough 差异；只读取公开生成 bytes。
7. Packed Node/browser/declaration smoke 覆盖 typed table default、snapshot、cell editor、write/reopen，CLI 保持 `0.1.0`。
8. 文档更新 changelog、API README、compatibility baseline 与 package README，明确 table-level creation supported、explicit-horz precedence、materialized direct state 和 table-level editor absence。
9. TypeScript project references、focused/full Vitest、performance、actual tarball、PowerPoint 2010 profile、package diff、LibreOffice/Poppler inspection 与 overflow checks 全部通过。

## 非目标与下一项

本小项不增加 table-level direction getter/editor、table-cell fit creation、rich/multi-paragraph cells、hyperlinks、merge/span、row insertion/deletion、table styles、auto-page、repeated headers、content measurement 或 layout recomputation。它不改变普通文本框七值方向或既有 direct cell editor。

下一小项将在完成全量对等清单的优先级复核后，从剩余 table creation/editor 缺口中选择依赖最少且可独立验证的一项。
