# Table Cell Fill Creation Design

## 目标与范围

本小项在已支持的 table cell object 基础上增加第一个创建时 `options` 能力：cell-level `fill`。调用方可以继续传 string、`{ text }`，也可以传 `{ text, options: { fill } }`；三者可在同一矩形 matrix 中混用。

本小项只支持现有原生 `TableCellFill` 的 direct no-fill 与 strict solid fill，不同时加入 border、margin、vertical alignment、text direction、text fit、hyperlink、merge/span、rich text、table-level fill、auto-page 或 repeated headers。后续每个 cell option 仍按独立小项补齐。

## 公共 API

新增并导出：

```ts
export interface AddTableCellOptions {
  readonly fill?: TableCellFill;
}

export interface AddTableCell {
  readonly text: string;
  readonly options?: AddTableCellOptions;
}
```

`AddTableCellInput = string | AddTableCell` 与 `SlideModel.addTable()` 签名保持不变。`fill` 复用已有读取/编辑类型：

```ts
export type TableCellFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };
```

因此创建、snapshot 和 `setCellFill()` 使用同一个 value model，不引入第二套 PptxGenJS-shaped native type。

## 方案选择

考虑三个方案：

1. 扩展 cell object 的 `options.fill`，复用 `TableCellFill`，归一化后由 table renderer 一次写出。采用此方案；公共模型一致、验证发生在 package mutation 前，且可稳定扩展后续 options。
2. 创建基础表格后循环调用 `setCellFill()`。该方案会反复 parse/serialize slide、产生嵌套 transaction，并让 creation output 依赖 editor insertion path，不采用。
3. 直接接受 PptxGenJS `fill: { color, transparency, type }`。这会让 native creation 与已有 native editor 使用两套类型，并保留 PptxGenJS 的 truthy/coercion 行为，不采用；adapter 仍负责读取其 public output。

## 严格输入归一化

Cell object 允许的 own keys 从只有 `text` 扩展为 `text` 加 optional `options`：

- object prototype 只能是 `Object.prototype` 或 null。
- `text` 必须存在且为 own data property；原有 single-paragraph/XML 规则不变。
- `options` 可省略或显式为 `undefined`；若存在 object，prototype 只能是 `Object.prototype` 或 null，own keys 只能是 `fill`。
- `{ text, options: {} }` 与 `{ text, options: { fill: undefined } }` 合法，均表示没有 direct cell fill。
- `fill`、solid `color` 及其所有字段必须来自 ordinary/null-prototype object 的 own data properties；accessor、继承 property、class instance、array、extra/symbol key 均拒绝且不得调用 getter。
- `kind: 'none'` 只允许 `kind`；`kind: 'solid'` 必须有合法 sRGB/theme `color`，optional transparency 必须为 finite `0..100`，沿用现有 `0.001%` quantization。
- 所有 normalized cell、options、fill 和 color 都与 caller object 脱离；后续 caller mutation 不影响定义或渲染。

共享 `normalizeTableCellFill()` 将改成 descriptor-safe，因此已有 `setCellFill()` 同时获得相同的 getter-free strict input contract。普通 object、null-prototype object 和冻结 object 继续合法；class/exotic/accessor state 不再被隐式读取。这是安全修复，不改变任何合法 fill 的 normalized value 或 OOXML。

## 内部模型与渲染

`NormalizedTableDefinition.rows` 从 detached string matrix 调整为：

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly fill?: TableCellFill;
}

readonly (readonly NormalizedTableCell[])[]
```

这只影响 table-creation internal surface；公开 `TableModel.rows` snapshot 不变。每个 string、`{ text }` 或 empty-options object 都归一化为 `{ text }`，filled cell 才携带 `fill`。

共享 fill encoder 从 `table-cell-fill.internal.ts` 导出给 renderer 与 editor 使用，防止两条序列化路径分叉。Cell OOXML 保持既有顺序：text body、`tcPr` margins、四条 direct no-fill border，最后是 optional cell fill。Solid fill 写 `a:solidFill` 和 strict `a:srgbClr` / `a:schemeClr`；explicit transparency 写 direct `a:alpha`。`kind: 'none'` 写 direct `a:noFill`。

没有 fill 时不增加任何 whitespace 或 element，因此 string、`{ text }`、`{ text, options: {} }` 和 `{ text, options: { fill: undefined } }` 对相同 text 必须生成 byte-identical native table XML。

## PptxGenJS 对等与差异

PptxGenJS 4.0.1 cell `options.fill` 使用 `{ color, transparency, type }`。对 supported solid sRGB/theme color 与 non-zero transparency，native output 在 direct fill choice、color token、alpha value 和最终 snapshot 上对等。

保留两项 intentional difference：

- PptxGenJS explicit transparency `0` 因 truthy 判断省略 alpha；native `transparency: 0` 沿用 editor contract，写 direct `alpha=100000`，从而保留 explicit direct state。
- PptxGenJS `{ type: 'none' }` 折叠为没有 direct fill；native `{ kind: 'none' }` 写 direct `a:noFill`，与 omission/table-style fallback 保持可区分。

PptxGenJS deprecated `alpha`、越界 transparency 和 runtime coercion 不进入 native API。adapter 仍 losslessly 保留其 XML，strict snapshot 只暴露合法状态。

## 测试与验收

实现必须覆盖：

1. Internal normalization：mixed string/object、empty options、solid sRGB/theme、none、explicit zero/fractional/full transparency、null-prototype 与 detached mutation。
2. Exact output：string、`{ text }`、empty options 与 undefined fill 的 native XML byte-identical；filled cell 的 XML 顺序、color 和 alpha 精确。
3. Invalid matrix：cell/options/fill/color accessor、inherited/class/array、extra/symbol key、missing kind/color、非法 color/transparency；所有 getter invocation count 为 zero。
4. Existing editor regression：descriptor-safe `setCellFill()`、numeric normalization、no-op、rollback 和 unknown OOXML preservation 不变。
5. Public model/SDK：创建后 live snapshot、immediate edit、duplicate isolation、outer rollback、write/reopen 与 non-target parts 稳定。
6. PptxGenJS 4.0.1：solid sRGB/theme/non-zero transparency 的 direct fill、snapshot 与 geometry 对等；explicit zero/none 差异有测试和文档。
7. Packed Node/browser/declaration/CLI smoke：`AddTableCellOptions` 可导入，filled cells 可创建、编辑、写出和重开。
8. TypeScript project references、focused/full tests与 performance gate。
9. PowerPoint 2010 validation 对 native source/edited/reopened/PptxGenJS baseline 均 zero error/zero warning。
10. Empty-options 等价 package diff 为 zero changed parts；单一 fill/edit 只改变目标 slide；write/reopen 全 part 稳定。
11. LibreOffice/Poppler 原图确认 sRGB/theme/transparent/no-fill cells、row/column geometry 和既有 text direction/alignment 无裁切或异常，overflow checker 通过。

## 文档与发布表面

更新 changelog、API README、PptxGenJS compatibility baseline、package README 和 packed smoke。文档明确 native creation 现在接受 empty cell options 与 `options.fill`，但 border、margin、alignment、direction、fit、hyperlink、merge、rich text、table-level options、auto-page 和 repeated headers 仍未支持。

## 非目标

本小项不改变 `TableCell.fill` snapshot 或 `setCellFill()` 的合法 value/OOXML 语义，不支持 gradient/pattern/picture/group fill、PptxGenJS deprecated `alpha`、table-level fill 或 effective table-style computation。它不改变 table geometry、column/row size、stable identity、existing-deck lossless editing、text renderer 或其他 cell editors。
