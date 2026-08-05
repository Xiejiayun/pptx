# PPTX 双向编辑库实施方案

> 状态：WP0–WP6 已实现；`0.1.1` `next` 技术预览发布中，stable 仍受独立客户端认证门禁约束
>
> 实施记录：[docs/implementation-progress.md](./docs/implementation-progress.md)
>
> 基线：Node.js / TypeScript，PptxGenJS `^4.0.1` 作为兼容适配层依赖
>
> 兼容目标：Microsoft PowerPoint 2010+；兼顾 Keynote、LibreOffice Impress、Google Slides

## 1. 背景与关键结论

PptxGenJS 擅长“从代码生成新的 PPTX”，但它不是现有 PPTX 的无损读取、编辑和回写引擎。新库不能只在 PptxGenJS 的对象模型上继续堆选项，否则无法可靠解决已有文件编辑、未知 OOXML 保留、动画/SmartArt 等复杂部件的双向处理。

因此采用以下架构：

1. 新建独立的 OOXML 双向内核，负责读取、编辑、校验和保存 PPTX。
2. 对未识别的 XML、关系、扩展命名空间和二进制部件进行原样透传。
3. 将 PptxGenJS 放在独立 adapter 包中，兼容现有创建型代码。
4. 渐变、透明度、母版、媒体等能力由 feature codec 实现；动画、转场、高级图表和 SmartArt 继续以插件形式扩展。
5. 核心包不依赖 PptxGenJS；只有 `pptxgenjs-adapter` 直接依赖 `pptxgenjs:^4.0.1`。

这不是 PptxGenJS 的 fork。PptxGenJS 负责成熟的“新建演示文稿”场景，新内核负责“已有演示文稿的无损双向编辑”和更完整的 OOXML 能力。

## 2. 产品目标

### 2.1 必须实现

- 从 `Buffer`、`Uint8Array`、`ArrayBuffer`、文件和流读取 PPTX。
- 将现有 PPTX 映射为可编辑的语义对象。
- 修改后保存为可被 PowerPoint 正常打开且不触发“修复文件”的 PPTX。
- 未修改内容尽可能保持字节不变；未识别内容必须保留。
- 提供稳定、类型安全、可扩展的 TypeScript API。
- 允许将现有 PptxGenJS 创建流程产出的 Buffer 导入新内核继续加工。
- 支持诊断、兼容性提示、严格校验和受控降级。

### 2.2 v1 范围

- 无损读取、局部编辑、保存。
- 幻灯片、文本、形状、图片、表格、常规图表等常用对象。
- 完整的 Master / Layout / Theme 读取、创建、引用和回写。
- 渐变填充和 OOXML 颜色透明度模型。
- 基础 Audio / Video 的嵌入、外链、封面图和关系管理。
- PptxGenJS adapter。
- OPC、关系、内容类型、XML 和跨客户端兼容性校验。

### 2.3 v1 后续能力

- Transition codec。
- Animation / Timing codec。
- Advanced Chart codec。
- SmartArt / Diagram codec。
- 更完整的批注、墨迹、宏、OLE、ActiveX 等专业对象编辑。

这些后续能力在 v1 中仍需做到无损透传，不能因为不支持编辑而丢失。

## 3. PptxGenJS 能力基线

以下矩阵以 PptxGenJS v4.x 的公开 API 和“程序化生成”定位为基线。具体小版本能力通过 conformance tests 持续校准，不能仅靠文档描述判断。

| 能力 | PptxGenJS 状态 | 新库策略 |
| --- | --- | --- |
| 新建演示文稿与幻灯片 | 已支持 | adapter 直接复用 |
| 文本、形状、图片、表格、链接 | 已支持 | adapter 复用；内核提供双向模型 |
| 常规柱状图、折线图、饼图等 | 已支持 | adapter 复用创建；内核负责读取和回写 |
| 读取现有 PPTX | 不支持 | 新 OOXML 内核实现 |
| 修改并无损回写现有 PPTX | 不支持 | 新 OOXML 内核实现 |
| 未知 OOXML 和扩展内容透传 | 不支持 | lossless XML 与 package graph 实现 |
| 母版、版式、占位符 | 可生成，覆盖不等同于完整双向编辑 | v1 Master/Layout/Theme codec |
| 透明度 | 常见颜色/填充场景部分支持 | v1 统一支持 OOXML alpha transforms |
| 渐变 | 缺少完整 OOXML 双向模型 | v1 Gradient codec |
| Audio / Video | 有创建型媒体能力，但双向编辑和兼容细节有限 | v1 Media codec |
| 转场 | 无完整公开模型 | 后续 Transition codec |
| 动画与时间轴 | 不支持 | 后续 Timing codec |
| 高级及现代图表 | 常见图表可用，复杂组合与现代图表不完整 | 后续 Advanced Chart codec |
| SmartArt | 不支持 | 后续 Diagram codec |

“PptxGenJS 已支持”不代表新内核要复制其内部实现。adapter 应将 PptxGenJS 的最终 PPTX Buffer 交给内核，避免绑定 PptxGenJS 私有对象结构。

## 4. 总体架构

```text
应用 / SDK
   │
   ├── 直接读取已有 PPTX ───────────────┐
   │                                    │
   └── PptxGenJS 创建代码 → PPTX Buffer │
                       │                │
                       ▼                ▼
                pptxgenjs-adapter   OOXML 双向内核
                                          │
                      ┌───────────────────┼───────────────────┐
                      ▼                   ▼                   ▼
                Semantic Model      Feature Codecs      Validator
                      │                   │                   │
                      └──────────────┬────┘                   │
                                     ▼                        │
                         Lossless XML Patch Layer ◄────────────┘
                                     │
                                     ▼
                              OPC Package Graph
                                     │
                                     ▼
                                  PPTX ZIP
```

### 4.1 OPC Package Graph

将 PPTX 视为带类型和关系的有向图，而不是文件名列表：

- 解析 `[Content_Types].xml`、根关系和所有 part relationships。
- 为每个 part 记录 URI、content type、relationship type、目标和反向引用。
- 支持 internal / external relationship。
- 统一分配并校验 `rId`、shape id、slide id、master id 等标识符。
- 防止悬空引用、重复标识符、非法相对路径和关系环问题。
- 新增/删除部件时自动同步 content types 和 `.rels`。
- 未被修改的二进制部件保持原始 payload。

### 4.2 Lossless XML Tree

常规 XML DOM 序列化会改变命名空间前缀、属性顺序、空白和扩展节点，容易产生巨大 diff，甚至破坏 Office 私有扩展。需要可追踪源片段的 XML 表示：

- 保留 XML declaration、命名空间前缀、属性顺序、注释和空白。
- 节点携带原始 source span；未变节点直接复用原始字节。
- 只对发生变化的最小子树重新序列化。
- `mc:AlternateContent`、`extLst` 和未知命名空间默认 opaque-preserve。
- 禁止 DTD、外部实体和网络实体解析。
- 提供 canonical 模式，仅用于测试、diff 和可重复构建，不作为默认保存方式。

### 4.3 Semantic Model

语义模型是 XML/package graph 上的可编辑视图，不是导入后完全重建 PPTX 的中间格式：

- `Presentation`、`Slide`、`Shape`、`TextBody`、`Table`、`Chart`、`Image`。
- `Master`、`Layout`、`Theme`、`Placeholder`。
- `Fill`、`Line`、`Color`、`Effect`、`Transform`。
- `Media`、`Relationship`、`ExtensionNode`。
- 每个语义对象保留 source part 和 source node identity。
- mutation journal 记录增、删、改，保存时生成局部 patch。
- 未有 codec 的对象以 `OpaquePart` / `OpaqueNode` 形式公开，只允许安全移动、复制或原样保存。

### 4.4 Feature Codec

每项复杂能力都实现统一接口：

```ts
interface FeatureCodec<T> {
  readonly id: string;
  detect(context: DecodeContext): boolean;
  decode(context: DecodeContext): T;
  encode(value: T, context: EncodeContext): XmlPatch[];
  validate(value: T, context: ValidationContext): Diagnostic[];
}
```

codec 必须声明所拥有的元素、关系和部件类型；不能删除自己不认识的兄弟节点或 `extLst` 内容。插件冲突按明确的 ownership 和优先级规则报错，禁止静默覆盖。

## 5. PptxGenJS Adapter

### 5.1 依赖边界

- `packages/pptxgenjs-adapter/package.json`：直接依赖 `pptxgenjs:^4.0.1`。
- 核心、model、OPC、codec、validator 均不得依赖 PptxGenJS。
- adapter 只使用 PptxGenJS 公开 API，不读取其私有字段。
- 使用 peer compatibility tests 验证允许的 PptxGenJS 版本范围。

### 5.2 两条使用路径

旧的创建型代码继续工作：

```ts
import PptxGenJS from 'pptxgenjs';
import { importPptxGenJS } from '@pptx/pptxgenjs-adapter';

const generated = new PptxGenJS();
generated.addSlide().addText('Hello', { x: 1, y: 1, w: 4, h: 1 });

const document = await importPptxGenJS(generated);
document.slides[0].background = {
  kind: 'linear-gradient',
  angle: 45,
  stops: [
    { offset: 0, color: '#2563EB', alpha: 1 },
    { offset: 1, color: '#7C3AED', alpha: 0.65 },
  ],
};

await document.writeFile('output.pptx');
```

已有 PPTX 不经过 PptxGenJS：

```ts
import { PptxDocument } from '@pptx/sdk';

const document = await PptxDocument.open('input.pptx');
document.slides[0].title.text = 'Updated';
await document.writeFile('output.pptx');
```

adapter 内部流程固定为：调用 PptxGenJS 公开的输出方法获得 Buffer，再由新内核读取。这样可以兼容未来 PptxGenJS 内部重构。

## 6. v1 Feature Codec 设计

### 6.1 Master / Layout / Theme

需要完整处理以下关系链：

```text
presentation → slide → slideLayout → slideMaster → theme
```

实现范围：

- 读取、创建、复制、删除和重新关联 master/layout/theme。
- 保留 master/layout 上的 shape tree、background、placeholder、header/footer 和扩展内容。
- placeholder 按 type、idx 和继承层级解析，不能只按 shape id 匹配。
- 支持 theme color scheme、font scheme、format scheme 和 theme override。
- 颜色 API 同时保留 resolved color 与原始 `schemeClr` 表达，避免编辑其他字段时把主题色固化为 RGB。
- 复制 slide 时正确复制或复用相关 layout/master/theme，不产生悬空引用。
- 提供 `materializeInheritedStyle()`，但默认保持继承关系。

### 6.2 Gradient / Transparency

统一颜色模型支持：

- `srgbClr`、`scrgbClr`、`schemeClr`、`sysClr`、`prstClr`。
- `alpha`、`alphaMod`、`alphaOff`、`alphaModFix` 等颜色变换。
- 颜色亮度、明暗、饱和度等已有 transforms 的原样保留和有序回写。

渐变模型支持：

- 线性渐变：角度、scaled、rotate-with-shape、flip。
- path 渐变：circle / rect / shape 及 fill rectangle。
- 任意数量 stop，stop position 映射到 OOXML 的 `0..100000`。
- 每个 stop 独立颜色空间、主题引用和 alpha。
- shape、background、line、text 等适用位置按 codec capability 明确声明。

透明度不能只设计成一个顶层 `opacity`。API 可以提供便捷属性，但编码时必须落到 OOXML 的颜色和效果变换，避免误改主题色或渐变 stop。

### 6.3 Audio / Video

v1 支持：

- 从文件、Buffer、stream 或外部 URL 创建媒体。
- 嵌入媒体 part 或创建 external relationship。
- 自动维护 content type、media relationship 和兼容扩展节点。
- 视频 poster frame / thumbnail。
- 基础播放设置：自动/点击播放、循环、隐藏播放图标、音量等；不被目标客户端支持时给出 diagnostic。
- 复制/删除 slide 时正确处理媒体引用计数。
- 相同媒体可按内容哈希去重，默认不跨文档泄漏路径信息。

兼容策略：

- PowerPoint 2010+ 优先使用兼容性最好的媒体关系表达。
- 对 Keynote、LibreOffice、Google Slides 维护独立 capability matrix。
- 外链媒体始终提示可移植性风险。
- 编解码器兼容性属于媒体内容问题，库只做容器/关系验证，不承诺转码；可提供可选 transcoder hook。

## 7. 后续插件设计

### 7.1 Transition Codec

- 处理 `<p:transition>`、速度、持续时间、点击/自动换页和声音。
- 支持 PowerPoint 版本扩展命名空间。
- 对 Morph 等新版本特性保留原始扩展，并提供 capability diagnostic。

### 7.2 Animation / Timing Codec

- 将 `<p:timing>` 建模为 timing tree，而不是扁平动画列表。
- 支持 sequence / parallel、trigger、delay、duration、repeat、target shape、text range 和 motion path。
- 先实现 appear、fade、wipe、fly 等高频效果，再扩展 emphasis、exit、media command。
- shape id 变化时必须同步 timing target；删除被引用 shape 时阻止保存或显式移除动画。

### 7.3 Advanced Chart Codec

- 支持组合图、次坐标轴、趋势线、误差线、复杂 data label。
- 管理 chart XML、style/color parts 和嵌入 workbook 的一致性。
- 现代图表（如 waterfall、histogram、treemap、sunburst、funnel、box-whisker）使用独立扩展 codec。
- 无法编辑的数据结构必须原样透传；可选的 image fallback 只能由调用方显式启用。

### 7.4 SmartArt / Diagram Codec

- 识别并维护 diagram data、layout、quick style、colors、drawing parts 及关系。
- 第一阶段仅支持读取、文本替换、节点增删和样式保留。
- layout engine 不成熟前，不承诺从零生成所有 SmartArt；可通过 PowerPoint 生成的模板进行参数化编辑。
- diagram 与 fallback drawing 必须保持一致性或明确标记需要 PowerPoint 重新布局。

## 8. 兼容性与降级模型

每项能力都声明：

- `native`：目标客户端原生支持且通过测试。
- `preserved`：库不编辑，但可无损透传。
- `degraded`：可受控转换为较低能力表达。
- `unsupported`：无法安全输出，严格模式阻止保存。

默认策略为 preserve-first，不静默 rasterize、不静默删除。调用方可针对目标客户端选择 profile：

```ts
await document.write({
  compatibility: 'powerpoint-2010',
  mode: 'strict',
});
```

计划提供：

- `powerpoint-2010`、`powerpoint-current`。
- `keynote-current`。
- `libreoffice-current`。
- `google-slides-import`。

## 9. 错误、诊断与安全

### 9.1 错误模型

- `PackageError`：ZIP、OPC、content type 或 relationship 错误。
- `ParseError`：已支持部件无法解析。
- `ValidationError`：输出将违反 OOXML 或内部不变量。
- `CompatibilityWarning`：合法但目标客户端可能降级。
- `OpaqueMutationError`：试图修改无 codec 所有权的未知内容。

每条 diagnostic 包含 severity、code、part URI、XML path、对象 id、兼容 profile 和修复建议。

### 9.2 安全约束

- ZIP bomb 限制：总解压大小、part 数量、单 part 大小和压缩比。
- 拒绝 ZIP path traversal 和非法 part URI。
- XML 禁用 DTD、外部实体和网络访问。
- 对媒体、OLE、宏和外链关系提供策略钩子。
- 默认不执行宏、不启动外部程序、不自动抓取外链资源。
- 提供资源预算和 AbortSignal，支持服务端安全取消。

## 10. Monorepo 结构

```text
packages/
  opc/                    # ZIP、Content Types、Relationships、Package Graph
  lossless-xml/           # 带 source span 的解析、最小 patch、序列化
  model/                  # Presentation/Slide/Shape 等语义模型
  codecs/                 # v1 内建 feature codecs
  validator/              # 结构、关系、兼容性和安全校验
  testkit/                # fixtures、diff、render/inspect helpers
  pptxgenjs-adapter/      # 唯一依赖 PptxGenJS 的包
  sdk/                    # 面向用户的统一 API
plugins/
  transitions/
  animations/
  advanced-charts/
  smartart/
fixtures/
  corpus/                 # 授权清晰的测试语料
  expected/
docs/
  architecture/
  compatibility/
```

工具链建议：pnpm workspace、TypeScript strict mode、Vitest、ESLint、API Extractor。ZIP 与 XML 依赖必须经过性能、安全和 lossless 能力验证后再确定，WP0 不预设具体库。

## 11. 测试与质量门槛

### 11.1 测试层级

1. 单元测试：URI、关系、id、颜色变换、EMU/角度换算、XML patch。
2. codec round-trip：decode → encode → decode 的语义等价。
3. package round-trip：打开并直接保存时保持所有 entry payload。
4. mutation isolation：只修改一个属性时，仅允许预期 part 和必要关系变化。
5. schema/relationship validation：无悬空关系、重复 id、缺失 content type。
6. golden files：覆盖 PowerPoint 不同版本及不同客户端创建的文件。
7. fuzz/property tests：ZIP、XML、关系图和未知扩展节点。
8. 性能测试：大文件、大图片、长文本、千级 shape 和大图表 workbook。

### 11.2 跨客户端验证

- 公共 CI：Node LTS、Windows/macOS/Linux；LibreOffice headless 打开和导出检查。
- Windows 私有 runner：安装 PowerPoint，自动打开、保存，并检测 repair dialog/COM 错误。
- macOS 私有 runner：Keynote 打开和导出检查。
- Google Slides：受控测试账号做导入/导出抽样，不放在每次 PR 的阻塞路径。
- 每个 release candidate 执行人工视觉对比和媒体播放抽检。

### 11.3 无损标准

- 文档无 mutation 时，`save()` 直接返回原始文件字节。
- 有 mutation 时，未触及 part 的解压 payload hash 必须不变。
- 未识别节点、关系和扩展命名空间在 round-trip 后必须存在且内容不变。
- ZIP entry 顺序、压缩参数和时间戳不作为 v1 的语义无损要求，但 reproducible-build 模式必须稳定。

## 12. 实施阶段

### WP0：基线与技术验证

交付物：

- ADR：架构、无损定义、插件 ownership、兼容 profile。
- 采集并分类最小测试语料。
- 评估 ZIP/XML 库并完成 source-span patch spike。
- 对 PptxGenJS `^4.0.1` 建立 API/输出 conformance baseline。

退出标准：可以打开一个含未知扩展的 PPTX，并在无修改时字节级返回原文件；局部修改 title 后其余 part payload 不变。

### WP1：OPC 与 Lossless XML

交付物：

- `opc`、`lossless-xml`、基础 validator。
- package graph、relationship/content type 更新器。
- mutation journal 和最小 XML patch writer。
- ZIP/XML 安全限制。

退出标准：语料库 package round-trip、unknown-part preservation 和关系完整性测试全部通过。

### WP2：基础语义模型

交付物：

- Presentation、Slide、Shape、Text、Image、Table、常规 Chart 模型。
- slide 增删、复制、排序和引用更新。
- 统一单位、颜色、transform 与继承模型。

退出标准：常规对象可读取、修改和保存；PowerPoint 打开不修复，mutation isolation 通过。

### WP3：PptxGenJS Adapter

交付物：

- `pptxgenjs-adapter` 及 `pptxgenjs:^4.0.1` 依赖。
- PptxGenJS → Buffer → 新内核导入流程。
- 版本兼容测试和迁移指南。

退出标准：代表性 PptxGenJS 示例可被导入、增加新能力并由各目标客户端打开。

### WP4：v1 高价值 Codec

交付物：

- Master/Layout/Theme codec。
- Gradient/Transparency codec。
- Audio/Video codec。
- 对应兼容 diagnostics 和 fixtures。

退出标准：v1 特性矩阵用例全部通过；PowerPoint 2010+ 无修复；其他客户端按 capability matrix 达到预期。

### WP5：SDK、验证与发布

交付物：

- `sdk` 公共 API、类型声明、错误与 diagnostics。
- CLI inspection 工具和 testkit。
- 性能、安全、fuzz、跨客户端 CI。
- API 文档、示例和 migration guide。

退出标准：达到第 13 节的发布标准，发布 `0.1.0` 技术预览版。

### WP6：扩展插件

按用户价值和技术风险依次推进：Transition → Animation/Timing → Advanced Charts → SmartArt。每个插件独立发布，未安装插件时仍须无损透传对应内容。

## 13. v1 发布标准

- 100% 通过受支持语料的无修改 round-trip 测试。
- 修改场景中，所有非预期 part payload hash 保持不变。
- PowerPoint 测试集无“发现内容有问题，需要修复”提示。
- v1 Master、Gradient/Transparency、Media 用例全部通过。
- PptxGenJS adapter 示例全部通过 `^4.0.1` 版本范围测试。
- 关键 API 有 TypeScript 类型测试和端到端示例。
- 安全预算、异常文件和 fuzz 测试没有未处理崩溃。
- 兼容性差异必须产生 diagnostic，不能静默丢失内容。

## 14. 主要风险与应对

| 风险 | 影响 | 应对 |
| --- | --- | --- |
| OOXML 与 Office 私有扩展非常复杂 | 全量语义模型周期失控 | lossless-first；按 codec 渐进覆盖 |
| 常规 XML 库导致巨大 diff | 未知内容损坏 | source-span tree 与最小 patch |
| PptxGenJS 内部实现变化 | adapter 不稳定 | 只依赖公开输出 API和 conformance tests |
| 动画/SmartArt 隐含引用多 | 文件被 PowerPoint 修复 | 独立插件、关系不变量、模板驱动起步 |
| 各客户端表现不一致 | 用户看到降级 | capability profile、diagnostics、显式 fallback |
| 媒体格式和编解码器差异 | 无法播放 | 容器校验、兼容建议、可选 transcoder hook |
| ZIP/XML 恶意输入 | 服务端资源耗尽或数据泄漏 | 资源预算、XXE 禁用、路径与外链策略 |

## 15. 首个工程里程碑

第一个可演示版本只做一条窄而完整的 vertical slice：

1. 读取一个真实 PPTX。
2. 构建 package graph。
3. 通过 lossless XML 定位第一页标题。
4. 修改标题并保存。
5. 验证未触及 part 的 payload hash 不变。
6. 用 PowerPoint 和 LibreOffice 打开且不触发修复。

完成该里程碑后再扩展对象模型。它将验证整个项目最关键、风险最高的技术前提：局部语义编辑与未知 OOXML 无损共存。
