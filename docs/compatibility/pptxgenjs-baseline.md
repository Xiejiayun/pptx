# PptxGenJS 4.x conformance baseline

基线版本：`4.0.1`；允许范围：`^4.0.1`。

WP0 只依赖以下公开表面：

- `new PptxGenJS()`。
- `addSlide()` 以及公开的 slide 创建 API。
- `write({ outputType: 'nodebuffer' })` / `writeFile({ fileName })` 产生标准 PPTX。

验证样例包含文本、形状、主题和 slide relationships。PptxGenJS 4.0.1 生成的真实文件可由 `PptxDocument.open()` 导入、局部修改标题，并由 LibreOffice headless 无修复地打开和导出。

adapter 不读取 `_slides` 等私有字段。后续 peer-range conformance test 会对 `^4.0.1` 的最低版本和当前最新兼容版本运行同一组输出导入用例。
