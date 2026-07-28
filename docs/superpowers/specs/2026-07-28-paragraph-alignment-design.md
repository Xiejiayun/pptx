# 段落水平对齐设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为普通文本和结构化 rich text 增加段落水平对齐，使原生创建和已有 PPT 编辑都能表达 PptxGenJS 4.0.1 的 `align: 'left' | 'center' | 'right' | 'justify'`。普通文本框可提供统一默认值，rich text 可逐段覆盖；读取 PptxGenJS 或 PowerPoint 生成的文件时返回相同语义。

本小项只包含水平段落对齐。不包含垂直对齐、RTL、分散对齐、泰文分散对齐、低字间距两端对齐、bullet/numbering、indent、spacing、tab stop 或文本框布局。这些能力继续作为独立小项实现。

## API

```ts
type TextAlignment = 'left' | 'center' | 'right' | 'justify';

interface AddTextOptions extends Partial<Transform> {
  readonly name?: string;
  readonly align?: TextAlignment;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly align?: TextAlignment;
}

const plain = slide.addText('Left\nRight', { align: 'center' });

const rich = slide.addRichText(
  [
    { runs: [{ text: 'Uses the text-box default' }] },
    { runs: [{ text: 'Overrides the default' }], align: 'right' },
  ],
  { align: 'left' },
);

const snapshot = rich.richText;
rich.richText = snapshot.map((paragraph, index) => ({
  ...paragraph,
  align: index === 0 ? 'justify' : paragraph.align,
}));
```

`AddTextOptions.align` 是创建时的文本框级默认值，应用到没有显式 `align` 的每个 paragraph。`RichTextParagraph.align` 优先于该默认值。`shape.richText` getter 返回每个 paragraph 本地声明的受支持对齐，不把 master/layout/list style 的继承值伪造成本地属性。

`shape.richText` setter 仍是显式整体替换。paragraph 未提供 `align` 时不写本地 `algn`，因此调用方可以清除本地对齐并恢复 OOXML 继承。该 setter 保留同位置既有 `pPr` 中本小项之外的内容，但重建 runs 的既有边界不变。

## 方案选择

考虑过三种方案：

1. 只在 `AddTextOptions` 增加文本框级 `align`。普通文本简单，但无法表达同一文本框中的混合段落，对 PptxGenJS `TextProps[]` 不完整。
2. 只在 `RichTextParagraph` 增加 `align`。值模型清晰，但普通 `addText(string, { align })` 这一常用 PptxGenJS 能力仍缺失。
3. 文本框级默认与逐段覆盖同时提供，并复用一个 `TextAlignment` 类型。它覆盖普通与混合段落，且不把 OOXML token 泄漏给调用方。因此采用。

不直接复制 PptxGenJS 将 run-level `align` 变化隐式切分为新 paragraph 的输入形态。当前结构化 API 已显式建模 paragraph，调用方应通过 paragraph boundary 表达不同对齐；导入 PptxGenJS 输出时，OOXML 中已经完成的 paragraph 分组会被正确读取。

## OOXML 映射

公开值与 `a:pPr/@algn` 的映射为：

| API | OOXML |
| --- | --- |
| `left` | `l` |
| `center` | `ctr` |
| `right` | `r` |
| `justify` | `just` |

普通 `addText()` 的每个 CR/LF paragraph 都写入默认对齐。`addRichText()` 对每段使用 `paragraph.align ?? options.align`；两者都未提供时不写 `algn`。

getter 只识别上述四个 token。OOXML 的 `dist`、`thaiDist`、`justLow` 等未支持 token 不映射成错误语义；只读和其他窄范围 mutation 会原样保留。调用方显式整体设置 `richText` 时，未出现在值模型中的对齐语义与其他未支持 run 语义一样不属于保留目标。

## 编辑与保留边界

rich text 替换按 paragraph 索引选择既有 direct `pPr` 作为模板；新增 paragraph 回退到第一段模板，没有模板时生成合法默认 `pPr`。每个模板只通过 lossless XML patch 增加、替换或移除 `algn`，不使用正则改写 XML，也不重排或重建 bullet、spacing、indent、tab、默认 run properties 及未知扩展 child。

`endParaRPr`、`bodyPr`、`lstStyle` 和 text body 的未知 direct child 保持既有 rich text 规则。paragraph 数量减少时被删除 paragraph 的本地属性随 paragraph 一起删除；数量增加时继承第一段的未公开 paragraph 模板，随后覆盖本小项公开的 `align`。

普通 `.text` setter 的既有语义不扩展：它仍以第一段为样式模板重建纯文本 paragraphs。需要逐段保留或修改对齐时使用 `richText` 值模型。

## 验证与错误处理

`align` 仅接受四个公开字符串。非字符串、大小写变体、OOXML token 和其他值在任何 package mutation 前抛出 `TypeError`。paragraph 的未知字段继续明确失败，避免拼写错误或未来字段被静默忽略。

`addText()`、`addRichText()` 和 `richText` setter 均在 package transaction 内完成验证、渲染和写回。失败必须恢复 part bytes、mutation journal 和 live shape 内容；外层 transaction rollback 也必须恢复 alignment。

## 测试与验收

1. `addText()` 为多个 CR/LF paragraphs 写入统一的 left/center/right/justify 对齐，并在 write/reopen 后保持。
2. `addRichText()` 支持文本框默认和逐段覆盖，空 paragraph 也保留 alignment。
3. getter 正确读取 `l/ctr/r/just`，不把缺失、继承或未支持 token 伪造成受支持值，且只读不产生 mutation。
4. `shape.richText` 替换可修改或清除逐段 alignment，同时保留同位置 `pPr` 的 bullet、spacing、indent、未知属性和 child。
5. 非法 alignment 在创建和替换前失败；内外层 transaction rollback 不改变 bytes/journal/identity。
6. PptxGenJS 4.0.1 真实输出验证文本框级默认和 run alignment 变化产生的四类 paragraph 映射。
7. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
8. LibreOffice headless 无修复打开并导出包含四种水平对齐的一页 PDF；兼容矩阵把 horizontal paragraph alignment 标记为已支持。
