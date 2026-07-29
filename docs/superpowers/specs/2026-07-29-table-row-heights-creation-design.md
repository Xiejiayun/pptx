# Table Row Heights Creation Design

日期：2026-07-29
状态：已批准实施（用户授权实现方按小项持续选择最佳方案）

## 目标与范围

扩展现有 `SlideModel.addTable()`，允许调用方在创建严格矩形字符串表格时为每行指定独立高度，或用一个 scalar 为全部行指定相同高度。行高继续使用本库统一的 EMU 单位，创建后返回同一个 live `TableModel`，并保持基础 table creation 的 validation、transaction、identity、duplicate、write/reopen 与 cell editor 契约。

本小项只补 PptxGenJS `TableProps.rowH` 的创建功能对等，不同时加入 cell object、rich/multi-paragraph cell text、merge、table/cell creation style、auto-page、repeated headers、hyperlink、内容测量或已有表格的行高读取/编辑。已有表格 row snapshot/mutation 及 transform/rows resize 协同语义作为后续独立小项。

## 公共 API

```ts
export interface AddTableOptions {
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly rowHeights?: number | readonly number[];
}
```

示例：

```ts
slide.addTable(
  [
    ['Region', 'Revenue'],
    ['East', '$1.2M'],
    ['West', '$980K'],
  ],
  {
    x: inches(1),
    y: inches(1),
    columnWidths: [inches(2), inches(4)],
    rowHeights: [inches(0.5), inches(1), inches(1.5)],
  },
);

slide.addTable(
  [['A'], ['B'], ['C']],
  { rowHeights: inches(0.75) },
);
```

字段名使用完整的 `rowHeights`，而不是复制 PptxGenJS 的缩写 `rowH`。这与 `width` / `height` / `columnWidths` 的公共模型一致。scalar 对每行重复；array 必须与矩形表格的物理行数完全相同。

## 方案选择

考虑过三种语义：

1. 完整复制 PptxGenJS `rowH` 的 truthy/coercion/fallback 行为。它最接近 runtime，但会静默截断长数组、混合短数组与 overall height 均分值，并生成 transform/rows 总高不一致的表格。
2. 只接受 exact-length `rowHeights` array。它最简单，但缺少 PptxGenJS 公开类型中的 scalar 能力，也与刚完成的 `columnWidths` 不对称。
3. 接受 EMU scalar 或 exact-length array；normalized row vector 决定 exact rows，总高由各行求和。若同时给出 `height`，只在 round 后与总和完全相等时接受。采用此方案。

方案 3 保持 public API、strict validation 与 column-width 语义一致，同时对有效 `rowH` public intent 实现功能对等并修复 PptxGenJS 的 height mismatch。它不把行高解释为比例，也不按 overall height 缩放调用方给出的 exact 值。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()`、安装包类型与实际 slide XML 校准：

- `rowH: number` 声明为应用到每一行的英寸高度，scalar fractional value 会逐行通过 inch-to-EMU rounding；例如三行 `rowH: 1.25` 生成三个 `h="1143000"`。
- exact-length array 会按项生成 `a:tr@h`。单元素 array 在后续行因 array-to-number coercion 被当作 scalar 重复，形成未声明的 shortcut。
- short array 的缺失项、falsy/zero item 会回退到 overall `h / rowCount`；没有 overall `h` 时可能回退为 zero。long array 的多余项被静默忽略。
- 省略 overall `h` 时，graphic-frame xfrm height 固定回退到 1 inch，即使 row heights 的总和不同。显式 `h` 与 `rowH` 也不做相等检查，因此可生成 `xfrm height !== sum(tr@h)`。
- string、negative、non-finite 或 runtime 类型外输入经过 truthy/`Number()` 分支，可能被 coercion、忽略、回退或写出非法值。

本库只对有效 public intent 做功能对等：scalar 精确重复，array 精确逐行应用，显式行高存在时 xfrm height 永远等于 row sum。它不复制 single-array shortcut、长度/falsy 回退、silent truncation、string coercion 或 xfrm/rows 不一致。

## 输入归一化与冲突规则

`rowHeights` 在任何 package mutation 前与 rows 和其他 options 一起 descriptor-safe 归一化：

- 省略 `rowHeights` 且省略 `height` 时，保持现有 auto-row-height 输出：table transform height 为 1 inch，每个 `a:tr@h` 为 zero。
- 省略 `rowHeights` 且提供 `height` 时，保持现有 quotient/remainder exact 均分，并要求每行至少 1 EMU。
- scalar 必须是 finite number，round 到 safe integer EMU 后大于零；每一行获得相同值。
- array 必须是 non-empty dense Array，只允许 `length` 和 exact numeric own data properties；symbol、extra property、sparse item 或 accessor 全部拒绝且不调用 getter。
- array 长度必须等于 row count。每项必须是 finite number，round 到 safe integer EMU 后大于零。
- normalized vector detached 复制，不保留 caller array；调用后修改源 array 不影响文稿。
- 所有 normalized row heights 的和必须是 positive safe integer。总和 overflow 抛 `RangeError`。
- 有 `rowHeights`、未给 `height` 时，table height 取 normalized vector 的 exact sum，并关闭 auto-row-height 标记。
- 同时给出 `height` 时，先按现有规则 round；只有它等于 vector sum 才接受，否则抛 `RangeError`。不 silently ignore、scale 或改写任一输入。
- `rowHeights: undefined` 等同省略；null、object、typed array、nested array、string、boolean、NaN、Infinity、unsafe、zero 或 negative item 都拒绝。
- 所有失败保持目标 slide bytes、model cache 与 mutation journal 不变。

`columnWidths` / width、x/y/name、matrix text validation 与 defaults 完全不变。

## 内部模型与 OOXML

`NormalizedTableDefinition` 增加 exact normalized row vector：

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
  readonly rowHeights: readonly number[];
}
```

`normalizeTableDefinition()` 选择以下路径：

1. 有 `rowHeights`：normalize scalar/array，计算 safe sum，验证 optional `height` 相等，`autoRowHeight=false`。
2. 无 `rowHeights`、有 `height`：沿用当前 height validation，调用 `distributeTableDimension(height, rowCount)`，`autoRowHeight=false`。
3. 两者都无：保留 `height=1 inch` 和每行 zero 的 auto vector，`autoRowHeight=true`。

列宽与行高使用同一个小型 descriptor-safe positive-dimension vector helper 和 overflow-safe sum helper，并传入各自 context；这避免两条严格验证规则以后发生偏移，同时保留现有 column-width 错误语义。

`renderTableGraphicFrame()` 直接按 `definition.rowHeights[rowIndex]` 写 `a:tr@h`，不再在 renderer 中从 overall height 重新推导。显式行高路径保证：

```text
definition.height === sum(definition.rowHeights)
```

auto-row-height 是唯一有意的例外：每行 `h=0` 交给 PowerPoint 测量，而 transform 保持既有 1-inch default。除 row values 与 explicit transform height 外，graphicFrame namespace、shape id/name、x/y/width、grid、cell text、margins、no-fill borders、schema order 和 deterministic XML 均保持不变。

## 原子性、identity 与兼容性

该扩展继续运行在 `SlideModel.addTable()` 的现有 OPC transaction 内。validation 在 parse/ID allocation/write-back 前完成；成功后只改变目标 slide XML 并返回 shape cache 中的同一个 `TableModel`。duplicate、outer rollback、write/reopen、已有 column-width creation、现有 cell setter 与 transform x/y 编辑保持现有行为。

本小项不改变已有表格的 transform/row mismatch，也不承诺通过 `TableModel.setTransform({ height })` 重分配 rows。后续“table row-height read/edit”小项将定义 strict row snapshot、exact mutation 与 table-height resize 的协同行为。

## 测试与发布门禁

1. internal tests 覆盖 scalar、exact array、fractional rounding、optional matching height、detached input、safe sum、auto rows、explicit equal distribution 与 exact XML row order/sum。
2. invalid tests 覆盖 empty/mismatched/sparse/extra/accessor/symbol arrays，typed arrays，wrong item types，non-finite/unsafe/zero/negative item，sum overflow，以及 `height` 与 sum 冲突；accessor call count 保持 zero。
3. model/SDK lifecycle 覆盖 from-zero unequal rows、stable identity、cell editing、duplicate isolation、outer rollback、write/reopen、PowerPoint 2010 validator 与 only-target-slide mutation。
4. PptxGenJS conformance 使用 exact-length public `rowH` array，并给显式 `h=sum`，比较 row values、xfrm height、grid/cell semantics 与 round-trip；另记录省略 `h` 时的 PptxGenJS mismatch，并验证原生 derived height repair。
5. packed Node/browser/declaration smoke 编译并运行 `rowHeights` scalar/array，验证 unequal rows、derived height、cell edit 与 reopen。
6. API、npm README、changelog 与 compatibility matrix 将 row heights 标记为创建时部分支持，并继续列出 existing-table row editing、cell objects、merge、styles 与 auto-page 未支持。
7. typecheck、focused/full tests、performance、actual npm tarball smoke、CLI validator、LibreOffice render、Poppler full-size visual inspection 与 overflow checks 全部通过。
