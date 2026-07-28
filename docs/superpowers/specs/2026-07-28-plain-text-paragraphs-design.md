# 纯文本换行与多段落设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

让基础文本 API 完整支持字符串换行：`slide.addText()` 可创建多段落文本，`ShapeModel.text` 和 `SlideTitleModel.text` 可按纯文本读取、覆盖并往返换行。该能力对齐 PptxGenJS 4.0.1 对 string 输入的公开行为，并消除当前遇到 CR/LF 即抛错的限制。

本小项仍是纯文本语义。调用方设置 `.text` 时，一个字符串覆盖 shape 的全部可见文本并统一使用第一段的样式模板；逐 run 字体、颜色、超链接和逐 paragraph bullet/spacing 的读写留给后续 rich text model。创建时的字体、颜色和 paragraph 默认值延续基础文本框实现。

## API 与换行语义

```ts
const text = slide.addText('First\nSecond\n\nFourth');
expect(text.text).toBe('First\nSecond\n\nFourth');

text.text = 'Updated\r\nNext';
expect(text.text).toBe('Updated\nNext');
```

输入中的 CRLF 和单独 CR 都规范化为 LF。以 LF 分割后，每个逻辑行对应一个 `a:p`；连续、开头和末尾换行产生空段落，因此 `''` 仍有一个空段落，`'A\n'` 有两个段落。这与 PptxGenJS 4.0.1 的 string 输出结构一致。

getter 按 direct `a:p` 的文档顺序读取。paragraph 之间输出 LF；paragraph 内遇到现有 `a:br` 也输出 LF；多个 `a:r/a:t` 按顺序拼接。这样打开的 rich text shape 虽暂不能逐样式编辑，纯文本视图仍不会吞掉段落或显式 line break。

## 方案选择

考虑过三种方案：

1. 继续拒绝换行，等待完整 rich text model。实现风险最低，但基础 string `addText()` 与 PptxGenJS 仍不对等，也阻塞常见从零创建场景。
2. 把 LF 写成单段落内 `a:br`。实现较小，但 soft line break 与新 paragraph 的 spacing、bullet 和编辑行为不同，也不匹配 PptxGenJS string 输出。
3. 每个逻辑行生成一个 paragraph，getter 同时理解 paragraph 和 `a:br`；plain setter 用第一段作为样式模板重建 paragraph 序列。因此采用。

## 创建与编辑实现

`addText()` 在现有 transaction 中规范化字符串，再为每个逻辑行生成一个 `a:p`。非空行包含默认 `a:r/a:rPr/a:t`，空行只需合法 paragraph properties 与 `a:endParaRPr`。所有 `a:t` 使用 `xml:space="preserve"` 并进行 XML text escaping，保证首尾空格和特殊字符往返。

plain setter 定位 shape 的 direct text body 和 direct paragraph：

1. 缺失 text body 或 paragraph 时抛出 `ModelParseError`，不猜测新的 shape 结构。
2. 选第一段作为模板。模板含 text node 时，克隆其完整 paragraph XML，第一 text node 写入该逻辑行、其余 text node 清空，并确保第一 node 有 `xml:space="preserve"`；因此 paragraph/run properties 和已存在的样式关系保持。
3. 模板没有 text node 或为自闭合 paragraph 时，保留可用的 paragraph/end-paragraph properties，并补一个默认 run。
4. 第一段替换为新的全部 paragraph，删除其余旧 direct paragraph；text body 的 `bodyPr`、`lstStyle` 和未知 direct children 原样保留。

覆盖 rich text 时，第一段的格式成为所有新逻辑行的 plain-text 格式模板；其余旧 run 的可见文字被清空。这个“plain setter 会折叠逐 run 语义”的边界必须在文档中明确，完整无损逐 run 编辑由后续 API 提供。

## 原子性与 identity

创建继续使用 `addText()` transaction。`.text` setter 也在 package transaction 中完成验证、段落渲染和 slide part 写回；任何模板错误、XML patch 冲突或写入异常恢复 bytes 和 mutation journal。

返回的 `ShapeModel` identity 不变。外层 document/package transaction rollback 后，getter 从恢复后的 live XML 读取原文本。换行编辑不新增 relationship 或 package part。

## 测试与验收

1. create → addText 的 LF、CRLF、CR、连续/开头/末尾空行映射到精确 paragraph 数，并由 getter 规范化为 LF。
2. write → reopen 后换行、空段落、首尾空格及 `& < >` 字符保持；package validation 无 error。
3. getter 正确组合多 run、paragraph 与 `a:br`，`SlideTitleModel.text` 使用相同语义。
4. plain setter 覆盖既有多 run/multi-paragraph shape 后返回精确文本，克隆第一段样式，保留 `bodyPr`、`lstStyle` 和未知 text-body child。
5. 缺失 text body/paragraph、非法 XML 控制字符和 setter 内失败不改变 slide bytes 或 mutation journal；外层 rollback 恢复旧文本。
6. 基础单行 addText 的现有 ID、transform、escaping、identity 和 LibreOffice 视觉结果继续通过。
7. 全仓测试、性能、Node/browser bundle、发布类型和 npm tarball smoke 覆盖多段落创建与编辑。
8. LibreOffice headless 无修复打开包含四个逻辑行和空段落的文稿并导出 PDF；PptxGenJS parity matrix 将 string 换行标记为支持，rich run/paragraph options 继续列为部分支持。
