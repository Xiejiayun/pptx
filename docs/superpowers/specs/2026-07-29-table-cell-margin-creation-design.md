# Table Cell Margin Creation Design

## 目标与范围

本小项在已支持的 table cell `options.border` / `options.fill` 基础上增加创建时 `options.margin`。调用方可以继续传 string、`{ text }`、border/fill cell object，也可以在同一矩形 matrix 中使用 `{ text, options: { margin } }` 或同时提供 margin、border 与 fill。

本小项只支持现有 point-based `TextBoxMarginInput` 能表达的四个 physical direct cell margins，并把显式创建值写入 `a:tcPr@marL`、`marR`、`marT`、`marB`。它不同时加入 table-level margin、alignment、vertical alignment、text direction、text fit、hyperlink、merge/span、rich text、auto-page、repeated headers 或内容测量。后续 cell option 继续独立补齐。

## 公共 API

扩展现有创建 option：

```ts
export interface AddTableCellOptions {
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
}
```

`margin` 复用已有文本框与 table-cell editor 的公共 value model：

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
```

所有数值使用 points。Scalar 应用于四边；tuple 使用 `[top, right, bottom, left]`；named object 覆盖指定 sides。`AddTableCell`、`AddTableCellInput` 和 `SlideModel.addTable()` 签名不变，创建、snapshot 与 `setCellMargins()` 继续使用同一 value model，不增加 PptxGenJS-shaped 双单位类型。

创建与编辑的 missing-side 语义有意不同：新建 cell 始终以当前 canonical direct margins 为基线，上/下 `3.6pt = 45,720 EMU`、左/右 `7.2pt = 91,440 EMU`。Scalar/tuple 覆盖四边；partial named object 只覆盖提供边，未提供或显式 `undefined` 的边保留 canonical default。Omitted/undefined margin、`{}` 和全 undefined named object 与相同 text 的既有创建 bytes 相同。现有 `setCellMargins()` 仍是 whole replacement，partial named 中未提供的边会被清除，`{}` / `undefined` 清除四个 direct attributes。

## 方案选择

考虑三个方案：

1. 在 `AddTableCellOptions` 增加 `margin?: TextBoxMarginInput`，共享严格 point normalizer，并把 normalized sides 覆盖到 canonical cell defaults 上后一次渲染。采用此方案；它与现有 native snapshot/editor 一致，保持既有 default bytes，也能表达 PptxGenJS 的最终 direct state。
2. 创建基础表后循环调用 `setCellMargins()`。这会为每个 styled cell 重复 parse/serialize slide、产生多次 mutation，并让 creation output 依赖 editor 的 whole-replacement/direct-absence 语义，不采用。
3. 原样接受 PptxGenJS `margin` 的 legacy dual-unit runtime：第一项 `<1` 按 inches、`>=1` 按 points。相同 number 的单位依赖数值，无法形成稳定双向 API，不采用；adapter 继续忠实读取 PptxGenJS 已生成的 direct EMU。

## 严格输入归一化

Cell `options` own keys 从 `border` / `fill` 扩展为 `border` / `fill` / `margin`，其余严格规则不变。共享 `normalizeTextBoxMargins()` 同时升级为 descriptor-safe：

- `margin` 可省略或为 `undefined`；number 必须 finite，乘以 12,700 后按最近 EMU 量化，quantized raw 必须位于 signed Int32 `-2,147,483,648..2,147,483,647`。
- Tuple 必须是 `Array.prototype` 的 ordinary dense array，长度恰好四项，只含四个 own data indices 与 `length`，每项都是合法 number；sparse、accessor index、array subclass、extra/symbol key 均拒绝。
- Named object prototype 只能是 `Object.prototype` 或 null，只允许 top/right/bottom/left own data properties；class、exotic prototype、accessor、inherited value、array 和 extra/symbol key 均拒绝。
- Named side 省略或显式 `undefined` 表示没有 creation override；scalar/tuple 四边都显式提供。`{}` 与全 undefined named object 合法，并保留四个 creation defaults。
- Normalized tuple/object 与 caller 立即脱离；caller 后续 mutation 不影响 normalized definition、live table 或 write output。
- 所有 getter-free 拒绝必须在 table geometry/rendering 和 package mutation 前完成，accessor invocation count 保持 zero。

共享 normalizer 加固会同时覆盖 `AddTextOptions.margin`、`ShapeModel.textMargins` 和 `TableModel.setCellMargins()`。普通、冻结和 null-prototype inputs 的 normalized 数值/OOXML 不变；此前可能读取 getter 或接受 inherited/class/exotic state 的行为改为严格拒绝。这是输入安全加固，不改变任何合法 value、storage ownership 或 editor mutation target。

## 内部模型与渲染

创建内部 cell 扩展为：

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
}
```

`margins` 只保存显式 normalized overrides。`table-cell-margins.internal.ts` 增加 creation renderer，由它在 canonical defaults 上 overlay overrides，并始终返回四个 direct attributes：

```ts
export function renderTableCellMarginAttributes(
  margins: TextBoxMargins | undefined,
): string;
```

输出顺序保持现有 canonical `marL`、`marR`、`marT`、`marB`，每个值使用 `Math.round(points * 12700)`。例如：

- omitted / `{}`：`marL="91440" marR="91440" marT="45720" marB="45720"`
- scalar `0`：四边 `0`
- tuple `[1, 2, 3, 4]`：`marL="50800" marR="25400" marT="12700" marB="38100"`
- named `{ top: 1, left: 4 }`：`marL="50800" marR="91440" marT="12700" marB="45720"`

Table renderer 用该 helper 替换当前 hard-coded margin attributes，随后仍按 L/R/T/B 输出四条 border，最后输出 optional fill。Margin 只改变 `tcPr` attributes，不改变 border/fill child order、text body、geometry、row/column sizes 或 public `TableModel.rows` shape。

没有 explicit margin、empty margin 或全 undefined named margin 时，不增加 whitespace 或改变任何 token，因此 string、`{ text }`、empty options、margin undefined 与 empty named margin 对相同 text 必须生成 byte-identical native OOXML。

## PptxGenJS 4.0.1 对等与差异

PptxGenJS 4.0.1 的 cell `options.margin` 公共类型是 scalar 或 exact TRBL tuple。省略时仍给每个普通 cell 写 top/bottom 0.05in、left/right 0.1in，与 native canonical defaults 完全相同。Cell-level margin 会覆盖 table-level materialized value，最终文件只保留每个 cell 的 direct attributes。

对 final direct state，native 可以精确对等：

- native `0` 与 PptxGenJS `0` 都写四边 zero；
- native `1` 与 PptxGenJS `1` 都写四边 1pt；
- native `7.2` 与 PptxGenJS `0.1` 都写四边 91,440 EMU；
- native `[3.6, 7.2, 10.8, 14.4]` 与 PptxGenJS `[0.05, 0.1, 0.15, 0.2]` 都写 T/R/B/L 的 45,720/91,440/137,160/182,880 EMU；
- native `-7.2` 与 PptxGenJS `-0.1` 都写四边 -91,440 EMU。

保留以下 intentional differences：

- Native 始终按 points；PptxGenJS 只检查第一项，第一项 `<1` 时整组按 inches、`>=1` 时整组按 points。因此 native `0.1` 是 0.1pt，而 PptxGenJS `0.1` 是 0.1in。
- Native 额外支持 partial named object；PptxGenJS public cell margin 只声明 scalar/TRBL。
- Native 严格拒绝 runtime coercion、non-finite、getter/class/exotic input 和 quantized Int32 越界；PptxGenJS 可能生成负值或无效 token。
- 本小项不增加 table-level margin creation/mutation；需要 per-cell 结果时显式设置 cell option，adapter 仍可读取 PptxGenJS table-level materialized output。

Adapter snapshots 继续把 final direct EMU 除以 12,700 暴露为 points，不猜测原始 input unit。

## 测试与验收

实现必须覆盖：

1. Shared normalizer：scalar/TRBL/named、empty/undefined named sides、zero/negative/fractional/Int32 boundaries、ordinary/frozen/null-prototype、detachment 与 descriptor safety；现有 text-box creation/editor 和 cell editor 合法行为不变。
2. Internal creation：mixed string/object、empty options、margin only、border/fill/margin combined、canonical default overlay 与 caller detachment。
3. Exact output：omitted/undefined/empty/all-undefined named margin byte-identical；scalar/TRBL/named 的 L/R/T/B public mapping、marL/marR/marT/marB token order与 border/fill child order精确。
4. Invalid input：cell/options/margin/named side/tuple index accessor、class/inherited/exotic array、extra/symbol key、sparse/wrong-length tuple、non-number/non-finite/Int32 overflow；getter count zero且失败无 mutation。
5. Existing editor regression：descriptor-safe `setCellMargins()`、whole replacement、clear、no-op、malformed repair、unknown OOXML preservation与 rollback 不变；text-box margin creation/editor 同样回归。
6. Public model/SDK：创建后 margin snapshot、caller detachment、immediate edit/clear、duplicate isolation、outer rollback、geometry、non-target parts和 write/reopen。
7. PptxGenJS 4.0.1：omitted default、zero、one-point、inch-equivalent scalar/TRBL、negative值的 final direct attributes、snapshot、geometry、border/fill 对等；dual-unit difference有测试和文档。
8. Packed Node/browser/declaration/CLI smoke：`AddTableCellOptions.margin` 可用，margin 与 border/fill 可同时创建、编辑、写出和重开。
9. TypeScript project references、focused/full tests、performance gate和 staged diff review。
10. PowerPoint 2010 profile 对 native source/edited/reopened/PptxGenJS baseline 均 zero error/zero warning。
11. Empty/default 等价 package diff 为 zero changed parts；单 cell margin/text edit只改变目标 slide；write/reopen全 part稳定。
12. LibreOffice/Poppler 原图确认 zero/uniform/asymmetric/partial-negative margins、border/fill、row/column geometry与无裁切；两份 overflow checker通过。

## 文档与发布表面

更新 changelog、API README、PptxGenJS compatibility baseline、package README 和 packed smoke。文档明确 native creation 现在支持 `options.border`、`options.fill` 与 point-based `options.margin`，解释 creation default overlay、TRBL public order、signed Int32 quantization、creation/editor missing-side 差异、PptxGenJS dual-unit legacy branch，以及仍未支持的 cell options。

## 非目标

本小项不改变 `TableCell.margins` snapshot 或 `setCellMargins()` whole-replacement 语义，不增加 table-level margin default API。它不支持 effective table style/inheritance、logical merge propagation、layout recomputation、alignment/direction/fit creation、hyperlinks、rich/multi-paragraph text、auto-page、repeated headers 或 content measurement。它不改变 border/fill contract、table geometry、column/row sizes、stable identity、existing-deck lossless editing、text renderer 或其他 cell editors。
