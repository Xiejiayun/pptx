# Rich Text 高亮设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加 highlight 创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 的 hex highlight，并支持其运行时已能生成的 theme highlight。该能力与现有 font family、font size、bold、italic、text color、underline、strike、soft break 和 paragraph options 组合。

本小项不包括 pattern/gradient highlight、带 tint/alpha 等 transform 的 highlight color、outline/glow、baseline/上下标、character spacing、语言/RTL/亚洲字体或 hyperlink。读取和非 rich-text mutation 不得破坏原始 run XML。

## API

```ts
interface RichTextRunStyle {
  readonly highlight?: RichTextColor;
}

slide.addRichText([
  {
    runs: [
      {
        text: 'Yellow',
        style: { highlight: { kind: 'srgb', value: 'FFFF00' } },
      },
      {
        text: ' Theme',
        style: { highlight: { kind: 'scheme', value: 'accent2' } },
      },
    ],
  },
]);
```

highlight 复用现有 `RichTextColor`，sRGB 允许六位 hex 或带前导 `#` 的六位 hex，并规范化为大写无 `#`；scheme 使用现有受支持 token。getter 返回新的 color 对象。没有本地 highlight 时省略字段。

## 方案选择

考虑过三种方案：

1. 完全复制 PptxGenJS 的 hex string。迁移最直接，但会让 text color、underline color 和 highlight color 在同一 style 中接受三种不同形状。
2. 只支持 sRGB `RichTextColor`。它覆盖 PptxGenJS 公共类型，却会丢失 PptxGenJS 运行时与 PowerPoint 都能生成的合法 theme highlight。
3. 复用完整 `RichTextColor`。它与现有 API 一致，覆盖 sRGB 和 scheme，并保持明确校验，因此采用。

不增加 `false`：OOXML highlight 没有显式 none token；整体替换时省略 highlight 即清除本地高亮，无法用一个合法 OOXML 值压制继承高亮。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `highlight: 'ffff00'` 写 `<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>`。
- `highlight: '#FF0000'` 去掉 `#` 后写 sRGB。
- 公共类型声明为 `HexColor`，但运行时 `accent1`、`tx2` 等合法 theme token 会写 `schemeClr`。
- 空字符串不写 highlight。
- 非法字符串会打印 warning 并静默写黑色；本库改为在 mutation 前失败，避免调用方请求被悄悄改色。

本库不接受 string shorthand；功能范围与 PptxGenJS 对等，但颜色输入沿用已有结构化公共模型。

## OOXML 映射

```xml
<a:rPr>
  <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
  <a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
  <a:uFill><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:uFill>
  <a:latin typeface="+mn-lt"/>
</a:rPr>
```

highlight 写为 direct `a:highlight`，其中恰好一个 direct `srgbClr` 或 `schemeClr`。它位于普通 text fill 之后、underline fill 和 font elements 之前。

getter 只读取 local `rPr` 中恰好一个 direct `highlight`。其 direct child 必须恰好是一个合法 sRGB 或受支持 scheme color，并且不能带未建模的 color transform；否则不伪造公共 highlight，其他合法 run style 仍正常暴露。读取不产生 mutation。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 highlight XML。plain `.text` 继续保留第一个 paragraph/run 的样式模板，因此原 highlight 随模板复制。

`shape.richText` setter 整体替换 run：旧 run/rPr/field/hyperlink 不逐属性保留。新 run 只写调用方提供的公开 style；未提供 highlight 就没有本地 highlight。paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。

## 验证与错误处理

`highlight` 必须是合法 `RichTextColor` 对象。string、boolean、`null`、array、空对象、额外字段、非六位 sRGB 和未知 scheme token 明确失败。

所有验证在 package mutation 前完成。单个 run 失败时，不得留下部分 paragraph、part bytes、mutation journal、live shape text/style 或 snapshot 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. `addRichText()` 覆盖 sRGB、带 `#` 的小写 sRGB、scheme，以及与 underline/strike/text color 和 paragraph options 的组合。
2. getter 读取两种颜色；孤立、重复、未知、畸形和带 transform 的 highlight 不伪造公共值，只读不产生 mutation。
3. `shape.richText` 增加、更换和清除 highlight，snapshot 隔离，plain `.text` 保留第一 run highlight，write/reopen 与 duplicate 一致。
4. string/boolean/null/array/空对象、额外字段和非法 color 在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
5. PptxGenJS 4.0.1 conformance 覆盖 hex、scheme 与无 highlight 的真实输出导入。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出 sRGB/theme highlight 对照页；用同版本 PptxGenJS 对照文件核对渲染与 OOXML。
