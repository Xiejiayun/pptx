# Table Cell Border Creation Design

## 目标与范围

本小项在已支持的 table cell `options.fill` 基础上增加创建时 `options.border`。调用方可以继续传 string、`{ text }`、只有 fill 的 object，也可以在同一矩形 matrix 中使用 `{ text, options: { border } }` 或同时提供 border 与 fill。

本小项只支持现有原生 `TableCellBorderInput` 能表达的四个 physical direct sides：none 或 strict solid-color line、point width，以及 omitted/solid/dash style。它不同时加入 margin、vertical alignment、text direction、text fit、hyperlink、merge/span、rich text、table-level border、auto-page、repeated headers、diagonal border 或 effective shared-edge layout。后续 cell option 继续独立补齐。

## 公共 API

扩展现有创建 option：

```ts
export interface AddTableCellOptions {
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
}
```

属性名使用 PptxGenJS 的 singular `border`；value 复用本库已有读取/编辑模型：

```ts
export type TableCellBorder =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly width: number;
      readonly style?: 'solid' | 'dash';
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
```

Scalar 应用于四边；tuple 使用 `[top, right, bottom, left]`；named object 覆盖指定 sides。创建器始终以当前 canonical 四边 direct noFill 为基线，因此 tuple/named 中的 omitted 或 `undefined` side 保持 noFill，而不会从新 cell 删除该 side。需要 direct absence 时，创建后可用现有 `setCellBorders()` whole-replacement editor 清除。这一创建语义与 PptxGenJS 的 materialized no-border sides、以及本库现有未指定 border 的 native bytes 一致。

`AddTableCell`、`AddTableCellInput` 和 `SlideModel.addTable()` 签名不变。创建、snapshot 与 editor 使用同一 side value，不引入第二套 PptxGenJS-shaped native border type。

## 方案选择

考虑三个方案：

1. 在 `AddTableCellOptions` 增加 `border?: TableCellBorderInput`，共享严格 normalizer/renderer，并把结果叠加到 canonical no-border 四边。采用此方案；它与现有 native API 一致，覆盖 PptxGenJS scalar/TRBL 能力，也允许 native named input。
2. 创建基础表后循环调用 `setCellBorders()`。这会为每个 styled cell 重复 parse/serialize slide、产生多次 mutation，并让 creation output 依赖 whole-replacement editor 的 direct-absence 语义，不采用。
3. 原样接受 PptxGenJS `{ type, color, pt }`。迁移表面最直接，但会与已有 `kind/color/width/style` editor 形成两套 native 类型，并继承 PptxGenJS 的 defaults、truthy 和 runtime coercion，不采用；adapter 继续负责读取其公开输出。

## 严格输入归一化

Cell `options` own keys 从只有 `fill` 扩展为 `border` 与 `fill`，其余规则不变：

- cell/options object prototype 只能是 `Object.prototype` 或 null，所有受支持字段必须是 own data property。
- `border` 可省略或为 `undefined`；`{ text, options: {} }`、`border: undefined`、`border: {}` 和全 undefined tuple 均归一化为 canonical 四个 noFill sides，并与同 text 的既有无 border 创建 bytes 相同。
- Scalar border 复制为四个 detached side；tuple 必须恰好四项、dense、无 extra/symbol key；named object只允许 top/right/bottom/left。
- `kind: 'none'` 只允许 `kind`；`kind: 'line'` 必须提供合法 sRGB/theme color 与 finite `0..1584` point width，量化到最近 EMU；style 只允许 omitted、solid 或 dash。
- scalar/named/tuple、每个 side、nested color 只接受 ordinary/null-prototype data objects；class、exotic prototype、accessor、inherited state、array-as-side 和 extra/symbol key 均拒绝，且不得调用 getter。
- fill 继续使用上一小项的相同 descriptor-safe contract；同时提供 border/fill 时两者分别深度脱离，caller 后续 mutation 不影响 normalized definition 或 live table。

共享 `normalizeTableCellBorders()` 将升级为 descriptor-safe。现有 `setCellBorders()` 因此获得同一 getter-free contract；所有原本合法的 ordinary/null-prototype scalar、TRBL、named 和 nested color value 保持 normalized value 与 OOXML 不变。这是输入安全加固，不扩大 editor 的 mutation target。

## 内部模型与渲染

创建内部 cell 扩展为：

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
}
```

`borders` 只保存显式 normalized overrides；renderer 在四个 canonical `{ kind: 'none' }` sides 上应用这些 overrides。公开 `TableModel.rows` 不变，创建后返回四边 detached snapshot。

`table-cell-borders.internal.ts` 导出一个 canonical creation encoder，editor 与 creation 共用单 side encoder，避免 color、width、dash 和 metadata 分叉。输出必须按 OOXML schema 的 L/R/T/B 顺序：

1. `lnL`
2. `lnR`
3. `lnT`
4. `lnB`
5. optional cell fill

None 写 `w="0" cap="flat" cmpd="sng" algn="ctr"` 与 `noFill`。Line 写 quantized width、strict `solidFill` sRGB/theme color、optional `prstDash`（solid 或 sysDash），以及现有 neutral round/head/tail metadata。显式 width 0 仍是 line，不折叠为 none。

没有 explicit border、empty border、all-undefined tuple 或 scalar four-side none 时，输出必须与当前 `NO_BORDERS` 常量逐字节相同。Border 与 fill 同时存在时，fill 仍位于四边之后；纯 fill 与未填充 table 的既有 bytes 不变。

## PptxGenJS 对等与差异

PptxGenJS 4.0.1 cell `options.border` 的公开类型是 scalar `BorderProps` 或 exact TRBL tuple。对显式 none、sRGB line、positive/zero point width、solid/dash 和完整 tuple，native 可生成相同 direct L/R/T/B state、final snapshot、margins、fill 顺序与 geometry。

保留以下 intentional differences：

- Native `border: {}` 表示没有 override并保留 canonical noFill；PptxGenJS 空 `BorderProps` 会应用默认 gray `666666`、1pt、solid。Native 要求 line 明确提供 kind/color/width，避免隐式 defaults。
- Native line style omitted 不写 direct dash；PptxGenJS line type omitted/default 会写 direct solid。需要 direct-state 对等时使用 `style: 'solid'`。
- Native 额外支持 theme color 与 named side object；PptxGenJS public cell border 只声明 hex color 与 scalar/TRBL。
- Native tuple/named 的 undefined side保留 canonical noFill；这与 PptxGenJS runtime partial tuple materialization相同，但 native 不接受 null 或其他 coercion。
- PptxGenJS table-level tuple 的 `pt: 0` 会因 truthy fallback 变为 1pt；本小项只做 cell-level border，native width 0 始终保留。

Negative/non-finite/over-range width、runtime strings/default coercion 和 unsupported line types 不进入 native API。adapter 继续 losslessly 读取公开 PptxGenJS output，strict snapshot 只暴露合法 direct state。

## 测试与验收

实现必须覆盖：

1. Shared normalizer：scalar/TRBL/named、none、sRGB/theme、style omitted/solid/dash、zero/fractional/max width、null prototype、deep detachment和 descriptor safety。
2. Internal creation：mixed string/object、empty options、border only、fill only、border+fill、canonical default overlay 与 input detachment。
3. Exact output：无 border/empty/undefined/empty named/all-undefined tuple/four none byte-identical；L/R/T/B order、width/color/dash/neutral metadata和 border-before-fill 精确。
4. Invalid input：cell/options/border/tuple item/color accessor、class/inherited/array、extra/symbol key、sparse/wrong-length tuple、missing kind/color/width、invalid style/color/width；getter count zero。
5. Existing editor regression：descriptor-safe `setCellBorders()`、whole replacement、no-op、malformed repair、unknown OOXML preservation和 rollback 不变。
6. Public model/SDK：创建后 snapshot、caller detachment、immediate edit/clear、duplicate isolation、outer rollback、geometry、non-target parts和 write/reopen。
7. PptxGenJS 4.0.1：explicit scalar/TRBL/none/zero/solid/dash 的 direct sides 与 final snapshot 对等；empty/default/style omission/table-level zero differences有测试和文档。
8. Packed Node/browser/declaration/CLI smoke：`AddTableCellOptions.border` 可用，border 与 fill 可同时创建、编辑、写出和重开。
9. TypeScript project references、focused/full tests、performance gate和 staged diff review。
10. PowerPoint 2010 profile 对 native source/edited/reopened/PptxGenJS baseline 均 zero error/zero warning。
11. Empty/default 等价 package diff 为 zero changed parts；单 cell border/text edit只改变目标 slide；write/reopen全 part稳定。
12. LibreOffice/Poppler 原图确认 sRGB/theme、solid/dash/none/zero-width、per-side sides、existing fill、row/column geometry与无裁切；两份 overflow checker通过。

## 文档与发布表面

更新 changelog、API README、PptxGenJS compatibility baseline、package README和 packed smoke。文档明确 native creation 现在支持 `options.fill` 与 `options.border`，并解释 creation default overlay、TRBL public order、L/R/T/B XML order、point range、none/line、style omission/direct solid、PptxGenJS defaults差异，以及仍未支持的 cell options。

## 非目标

本小项不改变 `TableCell.borders` snapshot 或 `setCellBorders()` whole-replacement 语义，不增加 table-level border default。它不支持 diagonal、gradient/pattern line fill、border transparency、advanced dash、compound/cap/join/alignment/arrow/effect、shared-edge precedence、effective table style、merge mutation或 layout measurement。它不改变 fill、table geometry、column/row size、stable identity、existing-deck lossless editing、text renderer或其他 cell editors。
