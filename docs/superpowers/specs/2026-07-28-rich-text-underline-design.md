# Rich Text 下划线设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加 underline 创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 的 boolean、有效 underline style、显式 none 和 underline color，同时支持从已有 PPT 读取并重新写入合法 OOXML underline。该能力与现有 font family、font size、bold、italic、text color、soft break 和 paragraph options 组合。

本小项不包括 strike、highlight、outline/glow、baseline/上下标、character spacing、hyperlink 自动下划线、underline line fill/width/dash 的底层 `uLn*` 编辑或 pattern/gradient underline fill。这些继续作为独立小项；读取和非 rich-text mutation 不得破坏原始 run XML。

## API

```ts
type RichTextUnderlineStyle =
  | 'words'
  | 'sng'
  | 'dbl'
  | 'heavy'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dotDashHeavy'
  | 'dotDotDash'
  | 'dotDotDashHeavy'
  | 'wavy'
  | 'wavyHeavy'
  | 'wavyDbl';

interface RichTextUnderline {
  readonly style?: RichTextUnderlineStyle;
  readonly color?: RichTextColor;
}

interface RichTextRunStyle {
  readonly underline?: boolean | RichTextUnderline;
}

slide.addRichText([
  {
    runs: [
      { text: 'Single', style: { underline: true } },
      { text: ' Double', style: { underline: { style: 'dbl' } } },
      {
        text: ' Colored',
        style: {
          underline: {
            style: 'wavyHeavy',
            color: { kind: 'srgb', value: 'FF0000' },
          },
        },
      },
      { text: ' Explicit none', style: { underline: false } },
    ],
  },
]);
```

`underline: true` 规范化为 single (`sng`)；`false` 写显式 none，用于压制继承下划线。对象必须提供 `style` 或 `color`；只提供 color 时默认 single。getter 对合法下划线返回规范化对象并总是带 `style`，对 `u="none"` 返回 `false`，没有本地 underline 时省略字段。snapshot 中的 underline/color 对象都是新值，不暴露模型内部引用。

## 方案选择

考虑过三种方案：

1. 只支持 boolean。它覆盖最常见用法，但会丢失 PptxGenJS 已公开的 double、dash、wavy 和独立 underline color，不能称为功能对等。
2. 直接复制 PptxGenJS 的 style union，包括拼错的 `dotDashHeave`。迁移表面最短，但会生成 OOXML 不支持的 token；其类型还遗漏合法的 `dotDashHeavy` 和 `words`。
3. 保留 PptxGenJS/OOXML 的合法 token 名称，纠正 typo，并用 `false` 表达 none。它覆盖全部有效 PptxGenJS 行为，也能读取 PowerPoint 的 words-only underline，因此采用。

underline color 复用现有 `RichTextColor`，不另建 hex/theme 字符串分支，避免 text color 与 underline color 接受不同格式。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `underline: true` 在预处理后写 `u="sng"`；`false` 不写本地 `u`。
- `{ style: 'sng' }`、`{ style: 'dbl' }`、`{ style: 'wavyDbl' }` 把 style token 原样写入 `a:rPr@u`。
- `{ style: 'dbl', color: 'FF0000' }` 写 `u="dbl"`，并写 `<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>`。
- `{ style: 'none' }` 写 `u="none"`。
- 只给 color 会写 `uFill` 却不写 `u`，结果依赖继承且 snapshot 语义不稳定；本库将其规范化为 single colored underline。
- 公共类型中的 `dotDashHeave` 被原样写成无效 OOXML token；运行时传正确的 `dotDashHeavy` 才生成有效值。本库只接受正确拼写。

PptxGenJS 暴露的其余有效 token 与 OOXML 相同，不需要兼容层翻译。

## OOXML 映射

```xml
<a:rPr u="wavyHeavy">
  <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
  <a:uFill>
    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
  </a:uFill>
  <a:latin typeface="+mn-lt"/>
</a:rPr>
```

所有 style token 直接映射到 `a:rPr@u`。`false` 写 `u="none"`。underline color 写 direct `uFill/solidFill`，其 color choice 是 `srgbClr` 或 `schemeClr`，并位于普通 text `solidFill` 之后、font elements 之前。没有 underline color 时不写本地 `uFill`，由 PowerPoint 使用文字颜色。

getter 严格读取 local `rPr@u`：`none` 返回 false，17 个有效 style 返回对象，未知/畸形 token 不伪造成公共 underline。只有 direct `uFill` 恰好包含一个 direct `solidFill` 且其中是合法 sRGB 或受支持 scheme color 时才暴露 color；fill 畸形时保留 style、忽略 color。`uFillTx` 和 `uLn*` 不在本小项的值模型内。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 underline XML。plain `.text` 继续保留第一个 paragraph/run 的样式模板，因此原 underline 仍随模板复制。

`shape.richText` setter 的目标是整体替换 run，沿用既有边界：旧 run/rPr/field/hyperlink 不逐属性保留。新 run 只写调用方提供的公开 style；未提供 underline 就没有本地 underline，提供 false 则显式写 none。paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。

所有 text、attribute 和 color 值仍使用 XML escaping/normalization，不以正则修改现有 run XML。

## 验证与错误处理

`underline` 必须是 boolean 或仅含 `style`/`color` 的对象。对象不能两项都缺失；style 必须是 17 个合法 token之一；color 使用现有 sRGB 六位 hex 与 scheme token 规则。字符串 shorthand、`null`、array、`none` style、PptxGenJS typo、未知字段和畸形 color 明确失败。

所有验证在 package mutation 前完成。单个 run 失败时，不得留下部分 paragraph、part bytes、mutation journal、live shape text/style 或 snapshot 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. `addRichText()` 覆盖 true、false、17 个 style、sRGB/scheme underline color、color-only 默认 single、空/多 run 和与既有 style/paragraph option 组合。
2. getter 读取 supported style、none、两种 color；未知 token 或无 `u` 的孤立 `uFill` 不伪造 underline，畸形 fill 保留 style 但省略 color，只读不产生 mutation。
3. `shape.richText` 增加、更换、清除和显式压制 underline，snapshot 隔离，plain `.text` 保留第一 run underline，write/reopen 与 duplicate 一致。
4. string/null/array/空对象、未知/`none`/`dotDashHeave` style、额外字段、非法 color 在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
5. PptxGenJS 4.0.1 conformance 覆盖 boolean true、single/double/wavy、none、sRGB color 与正确 `dotDashHeavy` 实际输出；adapter 按 OOXML 读取。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出一页 single/double/dash/wavy/colored underline；样式可见且页面无截断或溢出，兼容矩阵把 underline 标记为已支持。
