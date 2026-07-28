# 从零创建 Presentation 设计

日期：2026-07-28
状态：已批准实施

## 目标

增加不依赖输入文件或 PptxGenJS 的原生创建入口。调用方可以创建零页文稿、添加空白 slide、写出并重新打开一个具有完整 master/layout/theme 关系链的 OOXML package。

本小项是后续 `addText`、`addShape`、`addImage`、`addTable` 和 `addChart` 的基础，只负责 presentation package、页面尺寸和默认空白 slide。shape 创建、富文本、metadata 编辑、sections、speaker notes 和自定义 layout 分别在后续小项实现。

## API

```ts
type BuiltInSlideSize = '4:3' | '16:9' | '16:10' | 'wide';

interface CreatePresentationOptions {
  readonly format?: PresentationFormat;
  readonly slideSize?: BuiltInSlideSize;
}

const document = PptxDocument.create({
  format: 'pptx',
  slideSize: '16:9',
});
const slide = document.addSlide();
```

`create()` 同步返回 `PptxDocument`。默认 format 为 `pptx`，默认尺寸为 PptxGenJS 4.x 相同的 `16:9`（10 × 5.625 inch）。四种内置尺寸与 PptxGenJS 的公开 layout presets 对齐：

- `4:3`：10 × 7.5 inch；
- `16:9`：10 × 5.625 inch；
- `16:10`：10 × 6.25 inch；
- `wide`：13.333… × 7.5 inch。

六种现有 `PresentationFormat` 均可创建；presentation main content type、`formatProfile`、Blob MIME 和默认下载扩展保持一致。宏启用格式创建为空宏容器，本小项不生成 VBA project。

## 方案选择

考虑过三种方案：

1. 内嵌一个 base64 空白 `.pptx`。实现快，但模板不透明、bundle 更大，part 级变更难以 review。
2. 在 SDK 中直接调用 JSZip 生成全部文件。可行，但绕过现有 OPC content type、relationship、transaction 和 URI 规则，并让 SDK 新增 ZIP 所有权。
3. OPC 提供同步空 package factory，SDK 用现有 `setPart()`/`addRelationship()` 在一个 transaction 中组装可审查的 OOXML parts。该方案复用当前原子写入与图模型，因此采用。

OPC factory 只建立 `[Content_Types].xml` 的合法根结构；它不包含 PowerPoint 语义。Presentation 模板和所有 PowerPoint content type/relationship 由 SDK 创建层负责。

## Canonical package

零页文稿至少包含并正确连接：

- `[Content_Types].xml` 与根 `_rels/.rels`；
- `docProps/core.xml`、`docProps/app.xml`；
- `ppt/presentation.xml`；
- 一个默认 slide master、一个名为 `DEFAULT` 的空白 layout、一个默认 theme；
- `ppt/presProps.xml`、`ppt/viewProps.xml`、`ppt/tableStyles.xml`；
- 一个 notes master 及其 theme relationship，为后续 speaker notes 提供稳定基础。

模板 XML 作为 TypeScript 字符串常量存在，使用正式 OOXML namespace URI，不复制 PptxGenJS 私有状态或在运行时调用 PptxGenJS。theme 包含完整 color/font/format scheme；master 和 layout 都包含合法的 group shape property、color map 和双向 relationship。

创建后的 `slides` 为空。第一次 `addSlide()` 使用 presentation 的第一个 master 的第一个 layout；之后的 slide 沿用同一 layout。现有“从第一张 slide 继承 layout”的路径继续用于打开的文稿，零页 fallback 只在没有模板 slide 时生效。

空白 slide 不自动生成 notes slide；只有后续 speaker-notes API 才创建 notes owned 子图。

## 原子性与错误处理

整个基础 package 在一个 OPC transaction 中创建。无效 `format` 或 `slideSize` 在任何 part 写入前抛出 `TypeError`。模板解析、content type 或 relationship 失败时，factory 不返回半初始化 document。

创建后的 document 使用与打开文件相同的 model、codec、validator、write、Blob/download 和 transaction 路径，不建立第二套 writer。

## 测试与验收

1. 默认 `create()` 返回零页 `pptx`，具有一个 master/layout/theme 和 `16:9` 尺寸。
2. 四种内置尺寸写入精确 `p:sldSz` EMU。
3. 六种 format 写入正确 main content type，并在重新打开后恢复相同 `formatProfile`。
4. 第一次和连续 `addSlide()` 都链接默认 layout，slide/relationship ID 唯一且稳定。
5. write → reopen 后 package validation 为零 error，所有 internal relationship target 存在。
6. 在可用环境中由 LibreOffice headless 打开并导出一页 PDF，不触发修复。
7. Node/browser bundle、发布类型和 npm tarball smoke 通过；发布包不包含 PptxGenJS 运行时依赖。
8. PptxGenJS parity 文档把 constructor、四个内置 layout preset 和空白 `addSlide()` 标记为原生支持；custom layout 仍标记为未支持。
