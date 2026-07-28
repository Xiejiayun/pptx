# PptxGenJS 4.x conformance baseline

基线版本：`4.0.1`；允许范围：`^4.0.1`。

WP0 只依赖以下公开表面：

- `new PptxGenJS()`。
- `addSlide()` 以及公开的 slide 创建 API。
- `write({ outputType: 'nodebuffer' })` / `writeFile({ fileName })` 产生标准 PPTX。

验证样例包含文本、形状、主题和 slide relationships。PptxGenJS 4.0.1 生成的真实文件可由 `PptxDocument.open()` 导入、局部修改标题，并由 LibreOffice headless 无修复地打开和导出。

adapter 不读取 `_slides` 等私有字段。后续 peer-range conformance test 会对 `^4.0.1` 的最低版本和当前最新兼容版本运行同一组输出导入用例。

## 原生 API 对等进度

| PptxGenJS 4.x surface | 原生 API | 状态 |
| --- | --- | --- |
| `new PptxGenJS()` | `PptxDocument.create()` | 已支持 |
| 默认 `LAYOUT_16x9` | `create({ slideSize: '16:9' })` | 已支持 |
| `LAYOUT_4x3` / `LAYOUT_16x10` / `LAYOUT_WIDE` | `create({ slideSize: '4:3' / '16:10' / 'wide' })` | 已支持 |
| `addSlide()` 空白页 | `document.addSlide()` | 已支持 |
| `defineLayout()` 自定义尺寸 | `create({ slideSize: { width, height } })` | 已支持 |
| 创建后读取或修改页面尺寸 | `document.slideSize` | 已支持 |
| `slide.addText(string)` 单段文本 | `slide.addText(string, options)` | 已支持 |
| string 中的 CR/LF、多段落与空行 | `addText()` / `ShapeModel.text` | 已支持 |
| 文本框 x/y/w/h、旋转、翻转、名称 | `AddTextOptions` + `inches()` / `degrees()` | 已支持 |
| `TextProps[]` 基础 runs、字体、字号、粗斜体、颜色、soft break | `addRichText()` / `ShapeModel.richText` | 已支持 |
| rich run `underline` boolean、17 种有效 style、独立颜色 | `RichTextRunStyle.underline` | 已支持 |
| rich run `strike` boolean、single/double strike | `RichTextRunStyle.strike` | 已支持 |
| rich run `highlight` sRGB/theme color | `RichTextRunStyle.highlight` | 已支持 |
| rich run `outline` solid sRGB/theme color 与 point size | `RichTextRunStyle.outline` | 已支持 |
| rich run `glow` sRGB/theme color、point size 与 opacity | `RichTextRunStyle.glow` | 已支持 |
| rich run `superscript` / `subscript` / custom `baseline` | `RichTextRunStyle.baseline` | 已支持 |
| rich run `charSpacing` 正/负/fractional point | `RichTextRunStyle.characterSpacing` | 已支持 |
| 段落 `align: left/center/right/justify` | `AddTextOptions.align` / `RichTextParagraph.align` | 已支持 |
| 标准/custom bullet、16 种 numbering、startAt、indent | `AddTextOptions.bullet` / `RichTextParagraph.bullet` | 已支持 |
| 段前/段后、exact/multiple line spacing | `AddTextOptions.spacing` / `RichTextParagraph.spacing` | 已支持 |
| 多级列表 `indentLevel` 0–8 | `AddTextOptions.level` / `RichTextParagraph.level` | 已支持 |
| `tabStops`：left/center/right/decimal | `AddTextOptions.tabStops` / `RichTextParagraph.tabStops` | 已支持 |
| 文本框 `margin` scalar/TRBL 与 direct 四边编辑 | `AddTextOptions.margin` / `ShapeModel.textMargins` | 已支持 |
| 文本框 `valign: top/middle/bottom` 与 direct anchor 编辑 | `AddTextOptions.valign` / `ShapeModel.verticalAlignment` | 已支持 |
| 文本框 `wrap: boolean` 与 direct wrapping 编辑 | `AddTextOptions.wrap` / `ShapeModel.textWrap` | 已支持 |
| 文本框 `vert` 七种文本方向与 direct 编辑 | `AddTextOptions.vert` / `ShapeModel.textDirection` | 已支持 |
| 文本框 `fit: none/shrink/resize` 与 direct 编辑 | `AddTextOptions.fit` / `ShapeModel.textFit` | 已支持 |
| paragraph 左右 margin、first-line/hanging indent（非 bullet） | 尚无完整公开 API | 尚未支持，后续逐项补齐 |
| 文本框 RTL、table-cell fit/`textDirection` 与 run lang/transparency | 尚无完整公开 API | 部分支持，后续逐项补齐 |

LibreOffice headless 可无修复打开 underline 文件，但当前会把 double/dash/wavy 和独立 underline color 降级显示为普通单实线；同一 PptxGenJS 4.0.1 对照文件表现一致。OOXML token 与颜色仍保持合法并可由支持这些样式的客户端读取。

LibreOffice headless 当前也不显示 run-level single/double strike；同一 PptxGenJS 4.0.1 对照文件表现一致。三个合法 `strike` token 仍保留在 OOXML 中。

PptxGenJS 4.0.1 的 `margin` tuple 注释声明 `[top, right, bottom, left]`，但 runtime 对非对称值会交换 top/left。本库采用公开文档顺序，并提供具名 `{ top, right, bottom, left }` 输入；adapter 仍会忠实读取 PptxGenJS 已生成的实际 OOXML。

文本框 `valign` 省略时原生 API 与 PptxGenJS 4.0.1 都显式使用 middle；top、middle、bottom 分别映射到 direct `bodyPr@anchor` 的 `t`、`ctr`、`b`。PptxGenJS 放在单个 rich-text run 上的 `valign` 不生效，本库也只在文本框级公开该选项。

文本框 `wrap` 省略或为 true 时原生 API 与 PptxGenJS 4.0.1 都写 `bodyPr@wrap="square"`，false 写 `none`。PptxGenJS 对非法值静默回退 true、并忽略单个 rich-text run 上的 `wrap`；本库只在文本框级接受严格 boolean。

文本框 `vert` 支持 `eaVert`、`horz`、`mongolianVert`、`vert`、`vert270`、`wordArtVert`、`wordArtVertRtl`，并原样映射到 direct `bodyPr@vert`。省略值不写属性，因此与显式 `horz` 不同；`shape.textDirection = undefined` 只清除 direct override。PptxGenJS 4.0.1 会把非法字符串原样写入，但本库创建和编辑严格拒绝，读取未知 token 返回 `undefined` 且保留原始 XML。普通文本框的 `textDirection` 别名以及 run 级 `vert` / `textDirection` 在 PptxGenJS 中不生效；table-cell `textDirection` 留作独立能力。

文本框 `fit` 的 omitted/none 都不写 direct child，shrink 写 `bodyPr/a:normAutofit`，resize 写 `bodyPr/a:spAutoFit`。`shape.textFit` 也读取既有唯一 `a:noAutofit` 为 none；none 或 `undefined` 清除 direct fit choice，同模式 shrink/resize 赋值保留 PowerPoint 已计算的 `fontScale` / `lnSpcReduction`。PptxGenJS 对非法 outer 值和 run-level fit 静默忽略，本库严格拒绝非法原生输入；adapter 仍兼容 deprecated outer `shrinkText` / `autoFit` 输出。最终缩放比例可能要由 PowerPoint 在编辑文字或改变 shape 大小时动态重算。

原生创建会生成可重新打开和验证的 master/layout/theme 关系链，不通过 adapter，也不在运行时安装或调用 PptxGenJS。
