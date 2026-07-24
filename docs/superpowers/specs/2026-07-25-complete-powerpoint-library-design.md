# 完整 PowerPoint 操作库设计

日期：2026-07-25
状态：已确认

## 1. 目标与边界

本项目的产品定义是“完整 PowerPoint 操作库”，不是只提供少量无损编辑能力的 OOXML 内核。

权威行为基线为最新版 Windows PowerPoint。PowerPoint 2010+、Keynote、LibreOffice Impress 和 Google Slides 不定义核心语义，只通过兼容 profile、诊断和显式降级策略支持。

正式支持以下 OOXML 演示格式：

- `.pptx`
- `.pptm`
- `.ppsx`
- `.ppsm`
- `.potx`
- `.potm`

旧版二进制 `.ppt` 不在本设计范围内。宏、数字签名、OLE、ActiveX 及未公开扩展在缺少安全语义编辑能力时必须原样保留；库不得执行其中的代码或静默删除内容。

正式运行环境为 Node.js 20+ 与现代浏览器。Deno 和 Bun 不作为发布门禁。

## 2. 发布与封装

npm 只发布一个聚合包。本文使用 `pptx` 表示这个唯一公共包；实际 npm 名称遵循单包发布设计及仓库发布配置。仓库现有 `packages/*` 与 `plugins/*` 继续作为内部模块边界，但不要求使用者安装、理解或直接导入。

唯一公共入口提供完整功能：

```ts
import {
  Presentation,
  TextShape,
  Chart,
  Animation,
  SmartArt,
  Transition,
} from 'pptx';
```

构建阶段将内部模块整合进聚合包，发布制品不得包含 `workspace:` 依赖或要求安装内部 `@pptx/*` 包。根入口只导出经过筛选的稳定高层 API，不直接 `export *` 内部 OPC、XML、codec、validator 或 testkit 实现。

Node 与浏览器使用相同导入路径，由 conditional exports 选择环境实现。所有官方功能默认可用，不要求调用方安装插件。当前插件在内部仍可保持独立模块，但在公共 API 中表现为普通 PowerPoint 功能。

## 3. 总体架构

系统分为六层：

```text
Node I/O / Browser I/O
          ↓
统一 SDK、DocumentSession 与事务 API
          ↓
PowerPoint 语义对象模型
          ↓
Feature Codec 与完整功能集合
          ↓
Namespace-aware Lossless XML
          ↓
OPC Package Graph
```

### 3.1 I/O

- 通用核心接受 `Uint8Array`、`ArrayBuffer` 和 `Blob`。
- Node 实现增加文件路径、Node readable/writable stream。
- 浏览器实现增加 `File`、Web `ReadableStream`、`Blob` 和下载输出。
- `write()` 在所有环境返回 `Uint8Array`；环境适配器提供更自然的保存目标。
- 流式入口不得被宣传成恒定内存，除非 ZIP 实现和性能测试能够证明这一点。

### 3.2 内部模块

- OPC 层维护 part、content type、relationship 和入边/出边图。
- Lossless XML 层基于 namespace URI 查询，并保持未修改源片段。
- Model 层提供稳定对象 identity 和完整 PowerPoint 语义。
- Codec 层拥有明确的元素、关系与 part 生命周期。
- Validator 聚合 package、model 和全部已启用 codec 的诊断。
- SDK 是唯一公共门面，隐藏上述实现细节。

## 4. 公共对象模型

### 4.1 Presentation

```ts
const deck = await Presentation.open(input);
const created = Presentation.create({
  format: 'pptx',
  slideSize: 'wide',
});
```

`Presentation` 负责：

- 格式识别与创建；
- slides、masters、layouts、themes、sections、comments 和 document properties；
- 保存、验证、事务、兼容诊断；
- 宏、签名和 opaque 内容状态；
- 受控的 `rawParts` 高级逃生口。

### 4.2 Slides 与 Shapes

`Presentation.slides` 提供 `add`、`insert`、`duplicate`、`move`、`remove` 和 `find`。

`Slide.shapes` 是稳定的有序集合。Slide 同时提供：

- `addText`
- `addShape`
- `addImage`
- `addTable`
- `addChart`
- `addAudio`
- `addVideo`
- `addSmartArt`

所有可视对象统一支持：

- `clone` 与 `remove`
- `name`、`altText`、`visible`、`locked`
- transform、rotation、flip 和 group transform
- z-order 与分组
- hyperlinks、actions 和 accessibility metadata

### 4.3 Rich Text

文本模型固定为 `TextBody → Paragraph → Run`，覆盖字体、字号、颜色、主题字体、语言、粗斜体、下划线、字符间距、段落对齐、缩进、项目符号、编号、行距、超链接、文本框边距、自动适应和占位符继承。

简单的 `shape.text = value` 作为便捷 API 保留，但其格式保留行为必须明确并经过测试。

### 4.4 表格、图表与高级对象

- 表格支持结构和样式 CRUD、合并单元格、边框、填充和文本。
- 图表支持创建、删除、类型、轴、series、category、label、trendline、error bar 与组合图；修改数据时自动同步 cache 和内嵌 workbook。
- 动画、转场、SmartArt、媒体、备注、批注、公式、SVG、OLE、墨迹、3D Model、Zoom 和 Action 通过同一对象模型访问。
- 每项能力必须分别实现并验证 `read/create/edit/delete/preserve/validate`。

## 5. DocumentSession、事务与缓存

每个 `Presentation` 拥有一个 `DocumentSession`，统一管理 OPC 图、XML 文档、语义对象 identity、缓存、ID 分配器、mutation journal 和 codec registry。

所有公共写操作自动运行在 `MutationTransaction` 中：

1. 记录目标对象和依赖图快照；
2. 应用 XML、part 与 relationship 变更；
3. 失效受影响缓存；
4. 运行局部结构验证；
5. 成功则 commit，异常或验证失败则完整 rollback。

调用方可以使用显式批量事务以减少重复解析和验证。事务不得留下半创建 part、悬空关系、重复 ID 或已变更但未同步的 content type。

模型对象具有稳定 identity。读取属性不应每次重新解析整个 part；修改只失效直接受影响的语义缓存。

## 6. 依赖与引用生命周期

每个 codec 必须声明其关系生命周期：

- `owned`：Notes、Comments、独占 Chart 等随拥有者深拷贝和删除。
- `shared`：Theme、Master、Layout 及安全复用的媒体默认共享。
- `clone-on-write`：共享图片、媒体、chart workbook 等在修改前自动复制。
- `external`：保留 URL，绝不自动下载。
- `opaque`：未知内容原样保留，冲突修改产生错误。

Slide 复制必须由依赖图与 codec 协作完成，不能直接复制 `.rels`。删除对象后，只对引用归零且已声明可回收的 owned part 做垃圾回收。

统一 ID 分配器负责 slide、shape、relationship、timing、master、layout、comment 和其他 OOXML ID 空间。

## 7. Codec 契约与验证管线

内部 feature codec 至少实现以下契约：

```ts
interface FeatureCodec<T> {
  detect(context: DecodeContext): boolean;
  decode(context: DecodeContext): T;
  encode(value: T, context: EncodeContext): XmlPatch[];
  validate(value: T, context: ValidationContext): Diagnostic[];
  clone(context: CloneContext): ClonePlan;
  deleteDependencies(context: DeleteContext): DeletePlan;
}
```

保存前统一执行：

1. OPC 和 ZIP 结构检查；
2. content type 与 relationship 检查；
3. Presentation、Slide、Shape 和各 ID 空间检查；
4. Master/Layout/Theme 继承链检查；
5. 全部 codec 的结构与语义检查；
6. 指定兼容 profile 的诊断。

安装到文档中的任何官方能力都必须进入同一验证管线，不能依赖调用方手动调用插件验证方法。

## 8. 格式、宏与签名

`PresentationFormat` 描述六种 OOXML 格式的 presentation content type、主关系、宏能力和模板/放映语义。

- 打开后默认保持原格式。
- 格式转换必须使用显式 API。
- `.pptm/.ppsm/.potm` 默认保留 VBA part、关系和 content type。
- 删除宏必须显式请求，并同步转换为相应无宏格式。
- 修改已签名 package 时必须诊断签名失效风险，不能声称原签名继续有效。
- OLE、ActiveX 和其他可执行负载永不执行。

## 9. 错误与兼容诊断

公共错误统一为 `PptxError`，通过稳定的 `code` 区分：

- `INVALID_PACKAGE`
- `UNSUPPORTED_FORMAT`
- `PARSE_ERROR`
- `INVALID_OPERATION`
- `VALIDATION_FAILED`
- `OPAQUE_CONFLICT`
- `RESOURCE_LIMIT`
- `ABORTED`

错误可携带 `partUri`、对象 ID、XML path 和原始 cause。用户代码不应解析错误文本判断类型。

诊断严重度为：

- `error`：会损坏文件或违反最新版 PowerPoint 结构，默认禁止写出。
- `warning`：最新版 PowerPoint 可用，但目标兼容客户端可能降级。
- `info`：可移植性、性能或建议。

默认保存只以最新版 Windows PowerPoint 正确性作为硬门禁。调用方可以启用 `warningsAsErrors` 或为特定兼容客户端设置门禁。

库不静默删除、不自动栅格化、不自动下载外链。`save({ repair: true })` 只允许执行确定性、安全修复，并返回完整修复清单。

## 10. 完整性矩阵

仓库维护唯一权威功能矩阵。每项最新版 PowerPoint UI 可创建的功能、公开 OOXML 元素和 Microsoft 扩展分别记录：

- read
- create
- edit
- delete
- preserve
- validate
- PowerPoint round-trip

路线图阶段完成不代表产品完成。只有矩阵中对应证据齐全才能标记该项完成。未公开的专有数据至少必须无损保留并产生准确诊断。

功能实施顺序采用门禁式混合路线：

1. 跨运行时核心、格式 profile、namespace-aware XML、事务、引用生命周期和统一验证；
2. Rich Text 与常用 Shape；
3. 图片、SVG、Audio、Video；
4. 表格；
5. 常规与高级图表及 workbook 同步；
6. Master、Layout、Theme、Notes、Comments、Sections；
7. Animation、Transition、SmartArt；
8. 公式、OLE、ActiveX、墨迹、3D Model、Zoom 及其他 PowerPoint 功能。

## 11. PptxGenJS 功能完全覆盖门禁

本库必须覆盖 PptxGenJS 4.x 的全部公开功能，但使用新的高层对象 API；不承诺将现有 PptxGenJS 代码只替换 import 后直接运行。

仓库维护 `PptxGenJS parity matrix`，以当前锁定的最低兼容版本和发布时最新兼容 4.x 版本的公开类型、文档和官方示例为证据。每个 PptxGenJS 功能必须映射到新的公开 API、测试用例和最新版 Windows PowerPoint round-trip 结果。

对等范围至少包括：

- presentation layout、theme、metadata、sections、slide master 和 placeholder；
- slide 创建、背景、隐藏、页码、speaker notes 和 section 归属；
- Rich Text、RTL、亚洲字体、bullet/numbering、fit、margin、hyperlink 和 action；
- 全部公开 shape geometry、line、fill、transparency、shadow、rotation 和 grouping；
- path、data URI、SVG、animated GIF、裁切、contain、cover、sizing 和 hyperlink 图片能力；
- table cell/style、row/column sizing、merge、auto paging 与 HTML table-to-slides；
- PptxGenJS 公开的全部常规和组合 chart type、series 与 chart options；
- audio、video、online media/YouTube 和 poster frame；
- base64、Blob、Buffer、`Uint8Array`、文件、下载和 stream 输出；
- Node、浏览器、React、Angular、Vite、Electron 和 serverless 中与运行时无关的公开能力；
- 官方公开的 image sizing、data conversion 和其他辅助函数的等价能力。

功能等价不要求复刻 PptxGenJS 的参数形状。新的 API 应保持一致的 `Presentation → Slide → Shape` 对象模型，并提供迁移指南说明每项旧能力的新写法。

PptxGenJS adapter 只作为迁移和交叉验证工具；聚合 `pptx` 包自身必须能够完成 parity matrix 中的全部操作，不能把安装或调用 PptxGenJS 作为功能完成证据。

## 12. 测试与发布门禁

测试体系包括：

- XML、OPC、ID、关系与属性边界单元测试；
- 随机对象组合、编辑序列和恶意 ZIP/XML 属性测试；
- 六种 OOXML 格式的真实 PowerPoint 语料；
- no-mutation byte identity 与 mutation isolation；
- 深拷贝、clone-on-write、删除和垃圾回收；
- Windows PowerPoint COM 打开、保存、重新打开、导出和修复错误检查；
- PowerPoint 导出图片/PDF后的视觉回归；
- Keynote、LibreOffice 和 Google Slides 兼容诊断校准；
- Chrome、Edge、Safari 的打开、编辑、下载与大文件内存测试；
- Node 20、22 和当前 LTS，覆盖 Windows、macOS、Linux；
- 单个 npm tarball 在空项目中的安装、类型检查、运行和 CLI smoke test；
- 公共 API snapshot 和弃用策略；
- PptxGenJS parity matrix 的逐项 API、输出和迁移示例测试。

发布制品不得泄漏 workspace 协议、内部包入口或无法安装的依赖。

## 13. 完成定义

`1.0` 完成需要同时满足：

1. 唯一聚合 npm 包提供本文定义的完整高层 API；
2. PptxGenJS parity matrix 的全部公开功能均有新 API、测试、迁移示例和 PowerPoint round-trip 证据；
3. 功能矩阵中目标 PowerPoint 功能具有要求的 CRUD、保留、验证和 round-trip 证据；
4. 六种 OOXML 格式通过真实 PowerPoint 语料门禁；
5. Node 与浏览器发布矩阵通过；
6. 未知和未公开扩展满足无损保留约束；
7. 不存在已知的共享串改、悬空关系、事务残留或静默数据丢失；
8. 兼容客户端的已知差异均有稳定诊断。

在这些条件成立前，版本和文档必须明确标记为开发版或技术预览，不得使用“完整”描述当前实现状态。
