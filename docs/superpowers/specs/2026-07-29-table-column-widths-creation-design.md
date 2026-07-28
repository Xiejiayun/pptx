# Table Column Widths Creation Design

日期：2026-07-29
状态：已批准实施（用户授权实现方持续选择最佳方案）

## 目标与范围

扩展现有 `SlideModel.addTable()`，允许调用方在创建严格矩形字符串表格时为每列指定独立宽度，或用一个 scalar 为全部列指定相同宽度。列宽继续使用本库统一的 EMU 单位，创建后返回同一个 live `TableModel`，并保持基础 table creation 的 validation、transaction、identity、duplicate、write/reopen 与 cell editor 契约。

本小项只补 PptxGenJS `TableProps.colW` 的创建功能对等，不同时加入 row height、cell object、rich/multi-paragraph cell text、merge、table/cell creation style、auto-page、repeated headers、hyperlink、内容测量或已有表格的列宽 mutation。读取和编辑已有 `tblGrid` 将作为后续独立小项，以免把创建输入语义、malformed existing XML 策略和 transform/grid resize 语义混在一个提交中。

## 公共 API

```ts
export interface AddTableOptions {
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
}
```

示例：

```ts
slide.addTable(
  [
    ['Region', 'Revenue', 'Growth'],
    ['East', '$1.2M', '12%'],
  ],
  {
    x: inches(1),
    y: inches(1),
    columnWidths: [inches(2.5), inches(3.5), inches(2)],
    height: inches(1.5),
  },
);

slide.addTable(
  [['A', 'B', 'C']],
  { columnWidths: inches(1.25) },
);
```

字段名使用完整的 `columnWidths`，而不是复制 PptxGenJS 的缩写 `colW`。这与本库已经使用 `width` / `height` 而不是 `w` / `h` 的公共模型一致，同时保持功能对等。scalar 对每列重复；array 必须与矩形表格的列数完全相同。

## 方案选择

考虑过三种组合语义：

1. 直接采用 `colW` 的 inches 输入，并让它无条件覆盖 `width`。这最接近 PptxGenJS 调用表面，但与现有 EMU API 不一致，也会继承整体 width 与 grid width 相互矛盾的问题。
2. 采用 `columnWidths` 的 EMU 输入；它决定 exact grid，总宽由各列求和。若同时给出 `width`，只在 round 后与列宽总和完全相等时接受。采用此方案。它允许调用方保留显式整体 geometry assertion，又不会产生两个真相。
3. 把 `columnWidths` 视为比例，并按 `width` 缩放。这会让名为 width 的值失去 exact 语义，也偏离 PptxGenJS array `colW` 的绝对英寸含义。

方案 2 保持基础 API 的单位和 strict validation，并修复 PptxGenJS runtime 的 permissive fallback。后续 existing-table mutation 可以复用同一 normalized width vector，但需要单独设计 malformed grid 与 proportional resize 行为。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()`、安装包类型与实际 slide XML 校准：

- `colW: number` 声明为每列相同英寸宽度；runtime 却先计算 `Math.floor(colW * columnCount)`，再把取整后的总宽平均分配。例如三列 `colW: 1.25` 最终得到三列各 1 inch，而不是 1.25 inch。
- 单元素 array 在多列表格中走同一 scalar 取整分支。原生 API 已有明确 scalar 形态，因此 array 不复制这一模糊捷径。
- exact-length array 会按项生成 `gridCol@w`。若省略 overall `w`，PptxGenJS 的 xfrm width 使用默认可用页面宽度，并不等于 grid sum；若同时传 `w` 与 array `colW`，两者也可不一致。
- array 长度与 grid 列数不一致时，PptxGenJS 只写 console warning，并回退等宽列。非数字、零、负值及 runtime 类型外输入也存在 truthy/coercion 分支。
- fractional array item 逐项通过 inch-to-EMU rounding；有效 exact-length array 的 grid values 可由 adapter 无损读取。

本库只对有效 public intent 做功能对等：scalar 精确重复，array 精确逐列应用，xfrm width 永远等于 grid sum。它不复制 scalar floor、长度不匹配回退、console warning、string coercion 或 xfrm/grid 不一致。

## 输入归一化与冲突规则

`columnWidths` 在任何 package mutation 前与 rows 和其他 options 一起 descriptor-safe 归一化：

- 省略时保持现有行为：`width` 缺省为 `columnCount * 1 inch`，然后 quotient/remainder 等分到列；显式 `width` 也按现有规则 exact 等分。
- scalar 必须是 finite number，round 到 safe integer EMU 后大于零；每一列获得相同值。
- array 必须是 non-empty dense Array，只允许 `length` 和 exact numeric own data properties；symbol、extra property、sparse item 或 accessor 全部拒绝且不调用 getter。
- array 长度必须等于 column count。每项必须是 finite number，round 到 safe integer EMU 后大于零。
- normalized vector detached 复制，不保留 caller array；调用后修改源 array 不影响文稿。
- 所有 normalized column widths 的和必须是 positive safe integer。总和 overflow 抛 `RangeError`。
- 未给 `width` 时，table width 取 normalized vector 的 exact sum。
- 同时给 `width` 时，先按现有规则 round；只有它等于 vector sum 才接受，否则抛 `RangeError`。不 silently ignore、scale 或改写任一输入。
- `columnWidths: undefined` 等同省略；null、object、typed array、nested array、string、boolean、NaN、Infinity、unsafe、zero 或 negative item 都拒绝。
- 所有失败保持目标 slide bytes、model cache 与 mutation journal 不变。

`height` / auto row height、x/y/name、matrix text validation 与 defaults 完全不变。

## 内部模型与 OOXML

现有 `NormalizedTableDefinition` 增加 exact normalized vector：

```ts
export interface NormalizedTableDefinition {
  readonly rows: readonly (readonly string[])[];
  readonly name?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly autoRowHeight: boolean;
  readonly columnWidths: readonly number[];
}
```

`normalizeTableDefinition()` 先归一化 matrix/options，再选择以下单一路径：

1. 有 `columnWidths`：normalize scalar/array，计算 safe sum，验证 optional `width` 相等。
2. 无 `columnWidths`：沿用当前 default/explicit width validation，并调用 `distributeTableDimension(width, columnCount)`。

`renderTableGraphicFrame()` 不再自行从 overall width 推导列宽，而是直接按 `definition.columnWidths` 顺序生成 `a:gridCol@w`。它仍写 `p:xfrm/a:ext@cx=definition.width`，而 normalizer 保证：

```text
definition.width === sum(definition.columnWidths)
```

除 grid values 外，graphicFrame namespace、shape id/name、x/y/height、row heights、cell text、margins、no-fill borders、schema order 和 deterministic XML 均保持不变。不新增 relationship、part、style、merge 或 layout measurement。

## 原子性、identity 与兼容性

该扩展继续运行在 `SlideModel.addTable()` 的现有 OPC transaction 内。validation 在 parse/ID allocation/write-back 前完成；成功后只改变目标 slide XML 并返回 shape cache 中的同一个 `TableModel`。duplicate、outer rollback、write/reopen、现有 cell setter 与 transform x/y 编辑保持现有行为。

本小项不改变已有表格的 transform/grid mismatch，也不承诺通过 `TableModel.setTransform({ width })` 重分配 grid。后续“table column-width read/edit”小项将定义 strict grid snapshot、exact mutation 与 table-width resize 的协同行为。

## 测试与发布门禁

1. internal tests 覆盖 scalar、exact array、fractional rounding、optional matching width、detached input、safe sum、default/explicit equal distribution 与 exact XML grid order/sum。
2. invalid tests 覆盖 empty/mismatched/sparse/extra/accessor/symbol arrays，typed arrays，wrong item types，non-finite/unsafe/zero/negative item，sum overflow，以及 `width` 与 sum 冲突；accessor call count 保持零。
3. model/SDK lifecycle 覆盖 from-zero unequal columns、stable identity、cell editing、duplicate isolation、outer rollback、write/reopen、PowerPoint 2010 validator 与 only-target-slide mutation。
4. PptxGenJS conformance 使用 exact-length public `colW` array，并给显式 `w=sum`，比较 grid values、xfrm width、row/cell semantics 和 round-trip；另用 scalar 1.25 case 记录并修复其 floor bug。
5. packed Node/browser/declaration smoke 编译并运行 `columnWidths` scalar/array，验证 unequal grid、derived width、cell edit 与 reopen。
6. API、npm README、changelog 与 compatibility matrix 将 column widths 标记为部分支持，并继续列出 row heights、cell objects、merge、styles 与 auto-page 未支持。
7. typecheck、focused/full tests、performance、actual npm tarball smoke、CLI validator、LibreOffice render、Poppler full-size visual inspection 与 overflow checks 全部通过。
