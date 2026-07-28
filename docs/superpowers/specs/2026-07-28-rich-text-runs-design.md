# Rich Text 基础 Run 设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

增加第一个结构化 rich text API，使调用方能够在一个文本框中创建、读取和整体替换多个 paragraph/run，并为每个 run 设置字体、字号、粗体、斜体、sRGB/theme 颜色和 soft break。该能力覆盖 PptxGenJS 4.0.1 `addText(TextProps[])` 最常用的基础 run formatting，同时为后续 paragraph style、hyperlink 和高级 typography 提供稳定值模型。

本小项不包括 underline/strike、highlight、outline/glow、baseline/上下标、character spacing、语言/RTL/亚洲字体、hyperlink/action，以及 alignment、bullet/numbering、indent/spacing 等 paragraph options。读取不会修改或删除这些未知 OOXML；调用方显式整体设置 `richText` 时，只重建本小项公开的 run 语义。

## API

```ts
type RichTextColor =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string };

interface RichTextRunStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number; // points
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: RichTextColor;
}

interface RichTextRun {
  readonly text: string;
  readonly style?: RichTextRunStyle;
  readonly softBreakBefore?: boolean;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
}

const shape = slide.addRichText([
  {
    runs: [
      { text: 'Revenue ', style: { bold: true, fontSize: 24 } },
      { text: '+18%', style: { italic: true, color: { kind: 'srgb', value: '00A651' } } },
    ],
  },
  { runs: [{ text: 'vs. prior year', style: { color: { kind: 'scheme', value: 'tx2' } } }] },
], options);

const snapshot = shape.richText;
shape.richText = updatedParagraphs;
```

`AddTextOptions` 继续提供名称和 transform。`richText` getter 返回新的只读 snapshot，不缓存 XML，也不提供可变内部引用。setter 整体替换 rich content；需要只改纯文字时继续使用 `.text`。

`fontSize` 单位为 point，序列化为 `Math.round(value * 100)` 的 OOXML `sz`，合法范围 1–4000 point。sRGB 接受可选 `#` 的六位十六进制并规范化为大写；scheme color 限于 OOXML theme token（如 `tx1`、`accent1`、`hlink`）。

## 方案选择

考虑过三种方案：

1. 直接让 `addText()` 接受 PptxGenJS 的 `TextProps[]`。迁移短，但会把 PptxGenJS 的大量兼容字段和 paragraph/run 混合语义复制进新 API。
2. 一开始建立可变 `ParagraphModel`/`RunModel` 并承诺稳定 identity。OOXML run 没有稳定 ID，插入、合并和 PowerPoint 保存后很难定义 identity，当前阶段过重。
3. 提供明确的 paragraph/run immutable value tree，shape 负责整体读取与替换；后续可在其上增加窄范围 mutation helper。因此采用。

## OOXML 映射

每个 `RichTextParagraph` 写一个 `a:p`；空 `runs` 写合法空 paragraph。每个 run 的 `softBreakBefore` 先写 `a:br`，非空或带样式的 run 写 `a:r/a:rPr/a:t`。`a:t` 始终带 `xml:space="preserve"`。

run properties 映射为：

- `fontSize` → `a:rPr/@sz`，百分之一 point；
- `bold` / `italic` → 显式 `b="1|0"` / `i="1|0"`；
- `fontFamily` → 同值 `a:latin/a:ea/a:cs/@typeface`；
- sRGB color → `a:solidFill/a:srgbClr`；
- scheme color → `a:solidFill/a:schemeClr`。

未指定字体和颜色时沿用基础文本创建的 theme minor Latin (`+mn-lt`) 和 `tx1`，保证原生新建文稿的视觉默认一致。XML text 和 attribute 均转义，字符串不得包含 CR/LF；paragraph boundary 必须由外层数组表达，soft line break 由 `softBreakBefore` 表达。

getter 只读取 shape direct `txBody` 的 direct paragraph。`a:r` 与 `a:fld` 的可见 text 都成为 run snapshot；local `rPr` 中本小项支持的属性被解析，未设置或继承值不伪造成 local value。`a:br` 映射到后续 run 的 `softBreakBefore`；连续或末尾 break 使用空 text run 表达，确保再次设置时数量不丢失。

## 编辑与保留边界

`richText` setter 在现有 text body 中替换 direct paragraphs，保留 `bodyPr`、`lstStyle` 和未知 direct children。既有第一段的 direct `pPr` 与 `endParaRPr` 作为模板复制到新 paragraphs，因此本小项尚未公开的 alignment/bullet/spacing 不会因基础 run 编辑被无条件清空。

run 内容是调用方显式整体替换的目标，所以旧 run/field/hyperlink OOXML 不保留。仅调用 getter 或进行其他 shape 操作时，原始 rich text 和未知属性继续字节保留。`.text` plain setter 仍使用其已记录的“第一段样式模板”折叠语义。

## 原子性与错误处理

`addRichText()` 和 `richText` setter 都在 package transaction 内完成输入验证、XML 渲染和写回。至少需要一个 paragraph；paragraph 可以没有 runs。非对象/数组结构、非字符串 text、run 内 CR/LF、非法 XML 字符、错误 boolean、空/非法字体、超范围字号、非法颜色在任何 mutation 前失败。

缺失 shape tree、text body 或 paragraph 时抛出 `ModelParseError`。内部或外层 transaction 失败必须恢复 part bytes、mutation journal 和 live shape 内容；成功创建返回稳定的同一 `ShapeModel`。

## 测试与验收

1. 从零创建两个 paragraphs、多 runs、空 paragraph、soft break，并验证 text、richText snapshot、style 与 transform。
2. sRGB `#ff0000` 规范化为 `FF0000`；scheme token、字体特殊字符、12.5pt、显式 false bold/italic 正确往返。
3. write → reopen 后 paragraph/run 分组、soft break 和五类 style 一致，`.text` 给出等价纯文本。
4. 对打开的 rich text 整体替换时保留第一段 `pPr/endParaRPr`、`bodyPr/lstStyle` 和未知 direct child；shape identity 不变。
5. getter 识别 `a:r`、`a:fld`、连续/末尾 `a:br`，且只读不产生 mutation。
6. 所有非法输入、malformed text body 和内外层 rollback 不改变 bytes/journal。
7. PptxGenJS 4.0.1 真实 run 输出作为 conformance fixture；全仓、性能、Node/browser bundle、类型和 tarball smoke 通过。
8. LibreOffice headless 无修复打开并导出包含多字体、字号、颜色和 soft break 的一页 PDF；parity matrix 将基础 rich runs 标记为支持，其余文本 options 保持部分支持。
