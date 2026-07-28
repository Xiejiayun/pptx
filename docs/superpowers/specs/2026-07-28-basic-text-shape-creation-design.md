# 基础文本框创建设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

在原生创建或打开的 presentation 中增加第一个可视对象创建 API：`SlideModel.addText()`。它创建单段纯文本框，返回稳定的 `ShapeModel`，并立即支持现有的 `.text` 编辑与 `setTransform()`。

本小项覆盖文本内容、对象名称、位置、尺寸、旋转和翻转。rich runs、段落、换行、bullet/numbering、字体、颜色、margin、fit、RTL、hyperlink 和 action 属于后续独立小项。当前遇到 CR/LF 时明确抛错，不能把换行静默写成错误的单段 OOXML。

## API

```ts
interface AddTextOptions extends Partial<Transform> {
  readonly name?: string;
}

const text = slide.addText('Quarterly results', {
  name: 'Heading',
  x: inches(1),
  y: inches(0.75),
  width: inches(8),
  height: inches(1),
});

text.text = 'Updated results';
```

默认值为 `x=0`、`y=0`、`width=1 inch`、`height=1 inch`、`rotation=0`、无翻转。width/height 必须大于零；所有数值必须有限。负 x/y 合法，因为 PowerPoint 允许对象部分位于画布之外。

## 方案选择

考虑过三种方案：

1. 暴露 `appendShapeXml()`。实现最少，但把 shape id、namespace、转义、顺序与事务责任推给调用方，不构成功能对等。
2. 让调用方用 PptxGenJS 创建文本后再 adapter 导入。输出可用，但违反聚合包原生全功能目标。
3. 在 `SlideModel` 上提供语义 `addText()`，复用 OPC transaction、lossless XML patch 和 stable shape identity。该方案与现有对象模型一致，因此采用。

## OOXML 写入

`addText()` 在一个同步 package transaction 中：

1. 解析 slide 并定位直接 `p:spTree`；缺失时抛出 `ModelParseError`。
2. 扫描 slide 内全部 `p:cNvPr/@id`，分配 `max + 1`，至少从 2 开始，避免 group shape property 的 id 1。
3. 创建 `p:sp`，包含 non-visual properties、EMU transform、rect geometry、no fill/no line 和单个 `a:p/a:r/a:t`。
4. 文本和名称分别使用 XML text/attribute escaping。新增 shape 自带 presentation/drawing namespace 声明，不依赖源文件固定使用 `p`/`a` 前缀。
5. 若 `spTree` 有直接 `extLst`，在其前插入；否则追加到末尾，保持 schema child order 和未知扩展。
6. 写回 slide 后通过现有 shape collection 返回同一 `ShapeModel` 实例。

纯文本默认使用 theme `tx1` 颜色和 minor Latin font，默认垂直居中，与 PptxGenJS 4.x 基础文本输出保持视觉等价。创建不增加 relationship 或额外 package part。

## 原子性与 identity

validation、ID 分配、XML patch 和 slide 写回都属于同一 transaction。任何异常必须恢复 slide bytes 和 mutation journal。成功后返回的 model handle 在重复读取 `slide.shapes`、文本编辑和 transform 编辑后保持 `===` identity。

如果外层 document transaction 随后失败，package rollback 仍恢复 slide；已返回但不再存在于文稿中的 handle 访问时按现有 live-model 规则抛出 `ModelParseError`。

## 测试与验收

1. 从 `PptxDocument.create()` 添加 slide 和文本，write → reopen 后文本、名称、kind 和 transform 一致。
2. `& < >` 等内容以及引号名称正确转义，输出 XML 可重新解析。
3. 连续创建对象分配唯一递增 id，并保留既有未知 `extLst` 且插入顺序合法。
4. 返回对象在重复读取、`.text` 更新和 `setTransform()` 后保持 identity。
5. 非有限 transform、非正尺寸、换行或缺失 `spTree` 明确失败，且 transaction 无残留。
6. PptxGenJS parity 文档把单段 string `addText()`、基础位置/尺寸和后续编辑标记为支持；rich text/options 继续标记为部分支持。
7. LibreOffice headless 打开并导出包含文本的一页 PDF；全仓测试、性能、Node/browser bundle、类型和 npm tarball smoke 通过。
