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
| presentation `pptx.rtlMode` | `CreatePresentationOptions.rtlMode` / `document.rtlMode` | 已支持 |
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
| rich run `transparency` 0–100% | `RichTextRunStyle.transparency` | 已支持 |
| outer/run `lang` 默认、继承与覆盖 | `AddTextOptions.lang` / `RichTextRunStyle.lang` | 已支持 |
| 段落 `align: left/center/right/justify` | `AddTextOptions.align` / `RichTextParagraph.align` | 已支持 |
| 段落 `rtlMode` 默认与逐段覆盖 | `AddTextOptions.rtlMode` / `RichTextParagraph.rtl` | 已支持 |
| 标准/custom bullet、16 种 numbering、startAt、indent | `AddTextOptions.bullet` / `RichTextParagraph.bullet` | 已支持 |
| 段前/段后、exact/multiple line spacing | `AddTextOptions.spacing` / `RichTextParagraph.spacing` | 已支持 |
| 多级列表 `indentLevel` 0–8 | `AddTextOptions.level` / `RichTextParagraph.level` | 已支持 |
| `tabStops`：left/center/right/decimal | `AddTextOptions.tabStops` / `RichTextParagraph.tabStops` | 已支持 |
| 文本框 `margin` scalar/TRBL 与 direct 四边编辑 | `AddTextOptions.margin` / `ShapeModel.textMargins` | 已支持 |
| 文本框 `valign: top/middle/bottom` 与 direct anchor 编辑 | `AddTextOptions.valign` / `ShapeModel.verticalAlignment` | 已支持 |
| 文本框 `wrap: boolean` 与 direct wrapping 编辑 | `AddTextOptions.wrap` / `ShapeModel.textWrap` | 已支持 |
| 文本框 `vert` 七种文本方向与 direct 编辑 | `AddTextOptions.vert` / `ShapeModel.textDirection` | 已支持 |
| 文本框 `fit: none/shrink/resize` 与 direct 编辑 | `AddTextOptions.fit` / `ShapeModel.textFit` | 已支持 |
| paragraph 左右 margin、first-line/hanging indent | `paragraphMarginLeft` / `paragraphMarginRight` / `paragraphIndent` 与 rich paragraph overrides | 已支持 direct 创建、读取、编辑与清除 |
| `slide.addTable(string[][])` + name/x/y/w/h/colW/rowH | `slide.addTable(rows, options)` | 部分支持：strict rectangular single-paragraph strings、EMU geometry 与创建时 scalar/per-axis sizes |
| 已有 table `tblGrid/gridCol@w` | `TableModel.columnWidths` / `TableModel.setColumnWidths()` | 已支持 strict snapshot、direct 无损编辑与 transform width 同步 |
| 已有 table direct `tr@h` | `TableModel.rowHeights` / `TableModel.setRowHeights()` | 已支持 zero 自动行高、direct 无损编辑与条件式 transform height 同步 |
| table-cell `textDirection` | `TableCell.textDirection` / `TableModel.setCellTextDirection()` | 已支持 |
| table-cell `valign: top/middle/bottom` | `TableCell.verticalAlignment` / `TableModel.setCellVerticalAlignment()` | 已支持 direct 编辑 |
| table-cell `margin` scalar/TRBL | `TableCell.margins` / `TableModel.setCellMargins()` | 已支持 direct point snapshot 与编辑 |
| table-cell `border` scalar/TRBL | `TableCell.borders` / `TableModel.setCellBorders()` | 已支持 direct 四边读取、编辑、none 与清除 |
| table-cell `fill` solid/none/transparency | `TableCell.fill` / `TableModel.setCellFill()` | 已支持 direct 读取、编辑与清除 |
| table-cell bodyPr autofit | `TableCell.textFit` / `TableModel.setCellTextFit()` | 原生编辑已支持；PptxGenJS 4.0.1 本身无 table fit API |

基础 table creation 与 PptxGenJS 4.0.1 的 plain string table 输出在 table URI、geometry、grid/row 总尺寸、cell text、direct margins 和 direct no-fill borders 上对等。原生 API 返回 live `TableModel`，x/y 缺省为 0.5 inch、每列缺省 1 inch、总高缺省 1 inch，并修正 PptxGenJS 省略 width 时 xfrm width 为 0 而 grid 非零的不一致。`columnWidths` 以 EMU 接受 scalar 或严格等长数组：scalar 精确重复，数组按列保留，省略 width 时由列宽总和推导，显式 width 必须与总和一致。PptxGenJS 4.0.1 的 scalar `colW` 会先取整、单项数组被当作 scalar shortcut、数组长度不匹配会回退均分，且 array `colW` 可生成 xfrm/grid 总宽不一致；原生 API 对这些情况分别保留精度、要求精确列数并保持 transform/grid 一致。已有表格通过 `TableModel.columnWidths` 暴露 detached strict grid snapshot，并由 `setColumnWidths()` 原子编辑 direct `gridCol@w` 与 `ext@cx`；valid mismatch 会修复为 grid sum，malformed/ambiguous grid 读取为 `undefined`、写入则零 mutation 拒绝。PptxGenJS 没有已有 deck 对象模型，因此这部分是原生 lossless editing 扩展。`rowHeights` 同样以 EMU 接受 scalar 或严格等长数组；省略 height 时由行高总和推导，显式 height 必须相等，而完全省略时仍保留 `tr@h=0` 自动行高。已有表格通过 `TableModel.rowHeights` 暴露 direct `tr@h` detached snapshot，zero 保留为自动行高；`setRowHeights()` 对全正目标同步 `ext@cy` 并修复 valid mismatch，对含 zero 目标只更新 rows 并保留合法 transform height。PptxGenJS `rowH` 会把单项数组隐式广播、对 short/falsy item 混入 overall height 回退、忽略 long array 多余项，并在省略或冲突的 `h` 下生成 xfrm/rows 总高不一致；原生 API 严格拒绝这些模糊输入，且把 existing-deck 自动行高编辑作为 lossless extension。当前只接受严格非空矩形 `string[][]` 和 EMU name/x/y/width/height/columnWidths/rowHeights；cell objects、rich/multi-paragraph text、merge/colspan/rowspan、row insertion/deletion、table/cell creation styles、auto-page/repeated headers、hyperlinks 和内容测量仍未支持。

LibreOffice headless 可无修复打开 underline 文件，但当前会把 double/dash/wavy 和独立 underline color 降级显示为普通单实线；同一 PptxGenJS 4.0.1 对照文件表现一致。OOXML token 与颜色仍保持合法并可由支持这些样式的客户端读取。

LibreOffice headless 当前也不显示 run-level single/double strike；同一 PptxGenJS 4.0.1 对照文件表现一致。三个合法 `strike` token 仍保留在 OOXML 中。

PptxGenJS 4.0.1 的 `margin` tuple 注释声明 `[top, right, bottom, left]`，但 runtime 对非对称值会交换 top/left。本库采用公开文档顺序，并提供具名 `{ top, right, bottom, left }` 输入；adapter 仍会忠实读取 PptxGenJS 已生成的实际 OOXML。

文本框 `valign` 省略时原生 API 与 PptxGenJS 4.0.1 都显式使用 middle；top、middle、bottom 分别映射到 direct `bodyPr@anchor` 的 `t`、`ctr`、`b`。PptxGenJS 放在单个 rich-text run 上的 `valign` 不生效，本库也只在文本框级公开该选项。

文本框 `wrap` 省略或为 true 时原生 API 与 PptxGenJS 4.0.1 都写 `bodyPr@wrap="square"`，false 写 `none`。PptxGenJS 对非法值静默回退 true、并忽略单个 rich-text run 上的 `wrap`；本库只在文本框级接受严格 boolean。

文本框 `vert` 支持 `eaVert`、`horz`、`mongolianVert`、`vert`、`vert270`、`wordArtVert`、`wordArtVertRtl`，并原样映射到 direct `bodyPr@vert`。省略值不写属性，因此与显式 `horz` 不同；`shape.textDirection = undefined` 只清除 direct override。PptxGenJS 4.0.1 会把非法字符串原样写入，但本库创建和编辑严格拒绝，读取未知 token 返回 `undefined` 且保留原始 XML。普通文本框的 `textDirection` 别名以及 run 级 `vert` / `textDirection` 在 PptxGenJS 中不生效。

Table-cell `textDirection` 是独立的四值能力：`horz`、`vert`、`vert270`、`wordArtVert` 映射到选中物理 cell 的 direct `tcPr@vert`。`TableCell.textDirection` 只读取唯一 direct `tcPr` 上唯一、无 namespace 且精确合法的 token；`setCellTextDirection()` 以零基 physical row/cell index 编辑，显式 `horz` 写属性，`undefined` 只清除该属性。PptxGenJS 4.0.1 的 table-level 值会实体化到各 cell，cell-level omitted/explicit `horz` 都不写属性，因此导入时均为 `undefined`；runtime 放行的 `eaVert` 等类型外 token 也严格读为 `undefined` 并原样保留。基础创建后的 cell 可立即使用该编辑器，但 table-level default 仍未支持，并与 cell fit 独立。

Table-cell `verticalAlignment` 复用 `top`、`middle`、`bottom`，但只映射选中物理 cell 的 direct `tcPr@anchor="t/ctr/b"`，不读写 cell `bodyPr@anchor`。getter 要求唯一 direct `tcPr`、唯一无 namespace anchor 和精确 wire token；`setCellVerticalAlignment()` 写 canonical token，`undefined` 清除 direct anchor，不解析 effective default 或支持 `just/dist`。PptxGenJS 4.0.1 完全省略 table/cell `valign` 时不写 anchor；显式 table-level 值会实体化到未覆盖 cells，cell-level 值优先；runtime 类型外 token 原样保留但严格导入为 `undefined`。基础创建后的 cell 可立即使用该编辑器，但 table-level mutation 仍未支持。

Table-cell `margins` 复用 point-based `TextBoxMargins` value shape，但只读取唯一 direct `tcPr` 上独立合法的 `marL/marR/marT/marB` signed-Int32 integer，并以 point 返回 partial snapshot；不读取 `bodyPr@*Ins`、style 或 effective default。`setCellMargins()` 使用零基 physical row/cell index，接受 point scalar、TRBL tuple、partial named object、`{}` 或 `undefined`，整体替换四个受管 direct attributes；合法同数值为 exact no-op，malformed direct token 可由合法值覆盖，或因该边缺省而清除。PptxGenJS 4.0.1 省略 margin 时仍给普通 cell 写 top/bottom 0.05in、left/right 0.1in 的 narrow defaults，table-level 值会实体化到 cells，cell-level 值覆盖；其遗留 runtime 以第一项 `<1` 按 inches、`>=1` 按 points。adapter 忠实读取最终 EMU 并统一暴露 point，不逆推输入单位。基础创建后的 cell 可立即使用该编辑器，但 table-level mutation 与 layout recomputation 仍未支持。

Table-cell `borders` 只管理选中物理 cell 唯一 direct `tcPr` 下同 prefix 的 `lnL/lnR/lnT/lnB`，以 partial `{ top, right, bottom, left }` snapshot 暴露 strict direct state。side value 是 `{ kind: 'none' }` 或 `{ kind: 'line', color, width, style? }`；width 为 `0..1584` point 并量化到 EMU，color 支持 strict sRGB/theme，style omitted/solid/dash 分别写 no direct dash、`solid`、`sysDash`。`setCellBorders()` 接受 scalar、精确 TRBL tuple 或 partial named whole replacement；显式 none 写 zero-width `noFill` line，omitted side 删除 direct element，`{}` / `undefined` 清除四边。getter 独立省略 malformed/unsupported side，不读取或近似 diagonal、其他 dash、gradient/pattern line、compound/cap/join/arrow/effect，也不计算 adjacent shared-edge 或 effective table style；这些状态在无关 mutation 中原样保留。PptxGenJS 4.0.1 omitted border 会给每个 cell 实体化四条 noFill line，table-level scalar/TRBL 会复制到 cells，visible defaults 是 `666666`/1pt/solid，dash 落为 `sysDash`；cell-level `pt: 0` 保留 zero-width line，而 table tuple 的 `pt: 0` 因 truthy fallback 变成 1pt。adapter 只读取最终 L/R/T/B XML，不逆推输入层级。基础创建后的 cell 可立即使用该编辑器，但 table-level mutation、diagonal/advanced line editing 与 layout recomputation 仍未支持。

Table-cell `fill` 只读写选中物理 cell 唯一 direct `tcPr` 下的 fill choice，以 `{ kind: 'none' }` 区分 direct `a:noFill`，以 `{ kind: 'solid', color, transparency? }` 表达 strict sRGB/theme solid fill 和可选 direct alpha。`setCellFill()` 使用零基 physical index；`none` 写 direct noFill，`undefined` 删除 direct choice，两者与 table style fallback 的语义不同。transparency 是 finite `0..100` 百分比，量化到 `0.001%`；omitted 不写 alpha，explicit zero 写完全不透明 alpha。getter 只接受同 prefix、无歧义的 direct noFill/solidFill，不把 border/text descendants 当成 cell fill，也不解析 effective table style；其他 fill 类型在无关 mutation 中原样保留。PptxGenJS 4.0.1 会把 table-level fill 实体化到 cells，把 omitted 与 `type: 'none'` 都折叠为没有 direct fill，显式 zero transparency 折叠为没有 alpha；runtime 越界值可生成非法 alpha，adapter 保留 XML 但 strict snapshot 返回 `undefined`。基础创建后的 cell 可立即使用该编辑器，但 table-level mutation、gradient/pattern/picture fill 与 layout recomputation 仍未支持。

Table-cell `textFit` 复用 `none`、`shrink`、`resize`：只读取唯一 direct `txBody/bodyPr` 中唯一 supported fit child，既有 `noAutofit` 映射为 `none`，缺失或 malformed 状态返回 `undefined`。`setCellTextFit()` 以零基 physical row/cell index 编辑；shrink/resize 分别写 `normAutofit` / `spAutoFit`，none/`undefined` 都移除 direct choice 且不新增 `noAutofit`，同模式赋值保留 PowerPoint 已计算的 scale metadata。它不修改 `tcPr@vert`，也不计算动态缩放、改变 table 尺寸或提供 table-level default。基础创建后的 cell 可立即使用该编辑器；PptxGenJS 4.0.1 的 `TableCellProps` / `TableProps` 没有 fit API，runtime 透传的 `fit`、`autoFit`、`shrinkText` 均被忽略并生成 fit-less `bodyPr`。

LibreOffice headless 可正确显示 table-cell `horz`、`vert` 和 `vert270`，但当前把 `wordArtVert` 显示为水平文字；同一 direct OOXML 的 PptxGenJS 4.0.1 对照文件表现完全一致，`tcPr@vert="wordArtVert"` 仍保持合法并可由 PowerPoint 等支持该模式的客户端读取。

文本框 `fit` 的 omitted/none 都不写 direct child，shrink 写 `bodyPr/a:normAutofit`，resize 写 `bodyPr/a:spAutoFit`。`shape.textFit` 也读取既有唯一 `a:noAutofit` 为 none；none 或 `undefined` 清除 direct fit choice，同模式 shrink/resize 赋值保留 PowerPoint 已计算的 `fontScale` / `lnSpcReduction`。PptxGenJS 对非法 outer 值和 run-level fit 静默忽略，本库严格拒绝非法原生输入；adapter 仍兼容 deprecated outer `shrinkText` / `autoFit` 输出。最终缩放比例可能要由 PowerPoint 在编辑文字或改变 shape 大小时动态重算。

文本语言通过 `AddTextOptions.lang` 提供 plain/rich 创建默认值，`RichTextRunStyle.lang` 可覆盖单个 run。省略时 run 与 `endParaRPr` 使用 `en-US`；显式 outer/run 语言在 run 上同时写 `altLang="en-US"`，`endParaRPr` 只跟随 outer 值。getter 只暴露非空 direct `rPr@lang`，不解析 `altLang` 或 master/layout 继承。原生 API 拒绝非 string、空 string 和非法 XML 控制字符，并安全转义 attribute metacharacters。PptxGenJS 4.0.1 的 falsy fallback 仍由 adapter 输入兼容；未转义字符串产生的无效 XML 不属于兼容承诺。

Rich run `transparency` 使用 `0..100` 百分比：0 完全不透明，100 完全透明，并量化到最近 `0.001%`（`alpha = Math.round((100 - transparency) * 1000)`）。省略字段不写 alpha，显式 0 写 direct `alpha val="100000"`，因此本库可区分 absence 与 explicit opaque；PptxGenJS 4.0.1 因 truthy 判断会把 0 与 omitted 都省略。没有显式 run color 时本库沿用默认 `schemeClr tx1`，PptxGenJS 则实体化 direct black，effective transparency 相同。该字段只控制文字主 `rPr/solidFill`，不影响 outline、glow、highlight、underline、shape 或 table-cell fill。getter 只接受唯一 direct solid fill、唯一合法 sRGB/theme color 和唯一严格 integer alpha，不解析继承或其他/mixed color transform。

Paragraph 左右边距分别通过 `AddTextOptions.paragraphMarginLeft` / `paragraphMarginRight` 提供 plain/rich 创建默认，`RichTextParagraph.marginLeft` / `marginRight` 使用 point 提供本地 number override，`false` 抑制默认或清除对应 direct `pPr@marL` / `marR`。getter 仅读取 `0..51206400` 的 direct integer EMU，不解析继承。普通段落缩进通过 `AddTextOptions.paragraphIndent` 提供创建默认，`RichTextParagraph.indent` 接受 `-4032..4032` point：正值是 first-line indent，负值是 hanging indent，`false` 抑制默认或清除 direct `pPr@indent`。新建普通段落保留 canonical direct zero；setter omission 清除已有 direct 值。getter 仅读取 `-51206400..51206400` 的 direct signed integer EMU。direct bullet/numbering/picture bullet 拥有 `marL` 与 `indent`，因此 getter 不重复暴露 `marginLeft` 或 ordinary `indent`，同 paragraph 的 numeric left margin/ordinary indent 与 active bullet 也在 mutation 前拒绝；`marR` 与列表缩进独立，可在普通、bullet 或 numbering paragraph 上读取和编辑。左右字段与 indent sign 始终映射物理 OOXML，不因 RTL 交换。PptxGenJS 4.0.1 没有普通 paragraph 左右边距或 first-line/hanging indent 公开选项；其普通 paragraph direct `marL="0"` / `indent="0"`、列表 `marL` / negative `indent` 和默认缺失 `marR` 的输出均可正常导入。

Presentation RTL 通过 `CreatePresentationOptions.rtlMode` 创建，并可由 `document.rtlMode` 读取、编辑和清除。getter 只读取 direct `p:presentation@rtl`，接受 `1/true/on` 与 `0/false/off`；未知或 descendant token 不会伪造全局状态。true/false 分别写 `1` / `0`，undefined 清除 direct attribute。PptxGenJS 4.0.1 的 true 与 truthy 非 boolean 都写 `1`，false/omitted 都省略；本库严格拒绝非 boolean，显式 false 写 `0`，effective behavior 相同且 direct intent 可逆。全局 RTL 不改写 paragraph `pPr@rtl`、`bodyPr@rtlCol`、default text style、alignment 或 run 顺序。

段落 RTL 通过 `AddTextOptions.rtlMode` 提供 plain/rich 创建默认值，`RichTextParagraph.rtl` 可逐段覆盖。true/false 分别写 direct `pPr@rtl="1"` / `"0"`；setter 中省略字段会清除 direct override。getter 只接受 `1/true/on` 与 `0/false/off`，未知 token 返回 `undefined` 并在无关编辑中原样保留。该能力不读取或修改 `bodyPr@rtlCol` 或 presentation-level RTL，也不自动交换 alignment。PptxGenJS 4.0.1 的 valid outer true 多段输出可直接导入；其 run-level `rtlMode` 会在同一 paragraph 中插入重复且位置非法的 `pPr`，因此本库不公开 run-level RTL。

原生创建会生成可重新打开和验证的 master/layout/theme 关系链，不通过 adapter，也不在运行时安装或调用 PptxGenJS。
