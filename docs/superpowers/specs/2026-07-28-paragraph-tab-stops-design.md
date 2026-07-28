# 段落制表位设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为普通文本和结构化 rich text 增加段落级 tab stops，使原生创建和已有 PPT 编辑覆盖 PptxGenJS 4.0.1 的 `tabStops`：支持左、居中、右和小数点对齐，支持一个段落的多个有序制表位、文本框创建默认、逐段覆盖、显式空列表，以及包含 `\t` 的 plain/rich run 文本。

本小项不包括 default tab interval、leader 字符、文本框 margin/fit/wrap、普通 paragraph margin/first-line indent、RTL tab 行为或 master/list-style tab stop 编辑。这些继续作为独立小项；读取和非 rich-text mutation 不得破坏原始 `tabLst` XML。

## API

```ts
type ParagraphTabStopAlignment = 'left' | 'center' | 'right' | 'decimal';

interface ParagraphTabStop {
  readonly position: number;
  readonly alignment?: ParagraphTabStopAlignment;
}

interface AddTextOptions extends Partial<Transform> {
  readonly tabStops?: readonly ParagraphTabStop[];
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly tabStops?: readonly ParagraphTabStop[] | false;
}

slide.addText('Label\tValue', {
  tabStops: [{ position: 1.25, alignment: 'right' }],
});

slide.addRichText(
  [
    { runs: [{ text: 'Default\t100' }] },
    {
      runs: [{ text: 'Decimal\t12.50' }],
      tabStops: [{ position: 2.5, alignment: 'decimal' }],
    },
    { runs: [{ text: 'No local stops' }], tabStops: false },
    { runs: [{ text: 'Explicit empty list' }], tabStops: [] },
  ],
  { tabStops: [{ position: 1.5 }] },
);
```

`position` 与 PptxGenJS 一致，以英寸表示，并相对于 paragraph 左侧起点定位。`alignment` 缺失时默认 `left`。数组顺序原样保留，不自动排序、去重或合并相同 position；这与 PptxGenJS 的生成行为及 OOXML sequence 一致。

`AddTextOptions.tabStops` 是创建默认值；rich paragraph 未提供时继承默认值。`tabStops: false` 抑制创建默认并不写本地 `tabLst`，`tabStops: []` 写一个显式空列表。对已有 shape 设置 `richText` 时没有文本框默认：字段缺失或 `false` 都移除该 paragraph 的本地 `tabLst`，空数组仍保留显式空列表。

getter 对没有 `tabLst` 的 paragraph 省略字段，对显式空 `tabLst` 返回 `[]`，并始终返回新的数组和 stop 对象，避免调用方改写模型内部状态。

## 方案选择

考虑过三种方案：

1. 直接暴露 PptxGenJS 的 `'l' | 'r' | 'ctr' | 'dec'`。它便于机械迁移，但把 OOXML 缩写泄漏到公共模型，且与现有 `TextAlignment` 的完整语义名称不一致。
2. 使用语义 alignment 名称，同时保留 PptxGenJS 的英寸 position。API 可读、迁移只需映射四个枚举值，并完整覆盖功能，因此采用。
3. 使用 branded `Emu` position。它与 shape transform 一致且能逐 EMU 表达，但不兼容 PptxGenJS 的公开单位，也让常用的 1.25 英寸制表位变得冗长。

本库追求功能对等而不是复制 PptxGenJS 的运行时缺陷。position 0 不会被偷偷改成 1 英寸；缺失/字符串/NaN/Infinity position、未知 alignment 和额外字段明确失败。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `{ position: 1 }` 写 `pos="914400" algn="l"`。
- 2.25 英寸 right、3.5 英寸 center、4.75 英寸 decimal 分别写 `2057400/r`、`3200400/ctr`、`4343400/dec`。
- `tabStops: []` 写显式空 `<a:tabLst></a:tabLst>`。
- 每个 stop 按调用方数组顺序写入，不排序或去重；rich text 的逐 paragraph options 可以生成不同列表。
- 0 或缺失 position 因 `stop.position || 1` 被错误替换为 1 英寸；负 position 被原样转换；未知 alignment 会写出无效 token。
- 大于 100 的 number 被启发式视为已是 EMU，而不是英寸。

本库保留合法的负 position 和 0，并把所有 number 一致解释为英寸；转换后必须能表示为 OOXML signed 32-bit coordinate。无效 alignment 和畸形 stop 不进入 package。

## OOXML 映射

```xml
<a:pPr>
  <a:buNone/>
  <a:tabLst>
    <a:tab pos="1143000" algn="l"/>
    <a:tab pos="2286000" algn="dec"/>
  </a:tabLst>
</a:pPr>
```

`position` 使用 `Math.round(inches × 914400)` 转为 `a:tab@pos`。结果必须位于 signed 32-bit coordinate 的 `-2147483648..2147483647`。alignment 映射如下：

| API | OOXML |
| --- | --- |
| `left` | `l` |
| `center` | `ctr` |
| `right` | `r` |
| `decimal` | `dec` |

`tabLst` 是 direct `pPr` child，并严格放在 spacing 和 bullet block 之后、`defRPr` / `extLst` 之前。空数组写自闭合或等价的空 `tabLst`；getter 两种形式都规范化为 `[]`。

getter 只在 paragraph 恰有一个 direct `tabLst`，且其中所有 direct child 都是带合法 integer `pos` 和受支持 `algn` 的 `tab` 时暴露值。缺失属性、越界、未知 token、未知 child 或重复 `tabLst` 会让整个公共字段省略，避免把损坏列表伪装成部分成功的 snapshot。合法 signed EMU 除以 914400 返回英寸；再次写入可恢复同一个整数 EMU。

## Lossless patch 与清除

rich text 替换继续按 paragraph 索引选择 `pPr` 模板。实现只移除 direct `tabLst`，并在需要时按 schema 位置插入新列表；alignment、spacing、bullet、level、普通 margin/indent、`defRPr`、`extLst`、未知属性和其他未知 child 保持原顺序与字节内容。替换受支持的 tab stops 时，旧 `tabLst` 内部的未知内容属于被替换值，不继续保留。

未调用 rich-text setter 的读取、plain `.text` 替换、shape transform 和其他非 paragraph mutation 不触碰原始 `tabLst`。plain `.text` 继续以第一个 paragraph 作为样式模板，因此增加或减少换行后，新 paragraph 继承第一个 paragraph 的 tab stops。

所有插入、替换与清除继续使用 `LosslessXmlDocument` source spans，不以正则重写 XML，并覆盖 self-closing `pPr`、empty `tabLst` 和 namespace prefix。

## 验证与错误处理

`tabStops` 必须是数组；每项必须是仅含 `position`、可选 `alignment` 的普通对象。position 必须是 finite number，转为 EMU 后必须落在 signed 32-bit 范围。alignment 必须是四个公共值之一。验证在 package mutation 前完成；任意 stop 失败时，不得留下部分 `tabLst`、package bytes、mutation journal 或 live snapshot 变化。

允许重复 position、非递增顺序、负值、0 和亚 EMU 小数；position 在写入时统一四舍五入到最近 EMU。空数组是合法且有意义的显式值。

## 测试与验收

1. `addText()` 把默认 tab stops 应用到每个 CR/LF paragraph，覆盖四种 alignment、多 stop、重复/非排序 position、0、负值、空行与显式空列表。
2. `addRichText()` 支持文本框默认、逐段 override、`false`、`[]`、空 paragraph，并与 alignment、bullet、level、spacing 和包含 `\t` 的多 run 组合。
3. getter 正确读取 list 和 empty list，返回隔离 snapshot；畸形/越界/重复列表整体省略，读取不产生 mutation。
4. `shape.richText` 可增加、更换或清除 tab stops，同时保持 paragraph 属性和 schema child 顺序；plain `.text` 保留并复制第一个 paragraph 的 tab stops。
5. 非数组、null、缺失/非 number/NaN/Infinity/越界 position、未知 alignment、额外字段在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
6. PptxGenJS 4.0.1 conformance 覆盖四种 alignment、多个 position、empty list 和逐 paragraph list；adapter 按实际 OOXML 读取。
7. write/reopen、duplicate、snapshot isolation、全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
8. LibreOffice headless 无修复打开并导出一页包含 left/center/right/decimal tab 的真实 PPTX；各列对齐，页面无截断、重叠或溢出，兼容矩阵把 `tabStops` 标记为已支持。
