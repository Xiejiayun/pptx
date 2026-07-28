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
| 段落 `align: left/center/right/justify` | `AddTextOptions.align` / `RichTextParagraph.align` | 已支持 |
| 标准/custom bullet、16 种 numbering、startAt、indent | `AddTextOptions.bullet` / `RichTextParagraph.bullet` | 已支持 |
| 段前/段后、exact/multiple line spacing | `AddTextOptions.spacing` / `RichTextParagraph.spacing` | 已支持 |
| 多级列表 `indentLevel` 0–8 | `AddTextOptions.level` / `RichTextParagraph.level` | 已支持 |
| `tabStops`：left/center/right/decimal | `AddTextOptions.tabStops` / `RichTextParagraph.tabStops` | 已支持 |
| 普通段落 margin、first-line indent | 尚无完整公开 API | 尚未支持，后续逐项补齐 |
| rich run outline/glow、上下标、字符间距及其他文本布局 options | 尚无完整公开 API | 部分支持，后续逐项补齐 |

LibreOffice headless 可无修复打开 underline 文件，但当前会把 double/dash/wavy 和独立 underline color 降级显示为普通单实线；同一 PptxGenJS 4.0.1 对照文件表现一致。OOXML token 与颜色仍保持合法并可由支持这些样式的客户端读取。

LibreOffice headless 当前也不显示 run-level single/double strike；同一 PptxGenJS 4.0.1 对照文件表现一致。三个合法 `strike` token 仍保留在 OOXML 中。

原生创建会生成可重新打开和验证的 master/layout/theme 关系链，不通过 adapter，也不在运行时安装或调用 PptxGenJS。
