# Paragraph Left Margin 设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为 plain text 与结构化 rich text 增加非列表 paragraph 左边距的创建、读取、整体替换和清除能力，映射 direct `a:pPr@marL`。这是完整 PowerPoint 文本模型中 paragraph indentation 的第一个小项，也让已有 PPT 中的非列表 direct 左边距不再只能被 opaque 保留。

本小项不包括 `a:pPr@marR`、`a:pPr@indent` 表示的首行/悬挂缩进、master/layout/list-style 继承解析、RTL 语义换边，也不改变 bullet/numbering 已有的 hanging-indent 算法。右边距、首行缩进与悬挂缩进继续作为后续独立小项。

## API

```ts
interface AddTextOptions {
  readonly paragraphMarginLeft?: number;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly marginLeft?: number | false;
}

slide.addText('Indented paragraph\nSame default', {
  paragraphMarginLeft: 24,
});

slide.addRichText(
  [
    { runs: [{ text: 'Uses 24pt default' }] },
    { runs: [{ text: 'Uses 12pt override' }], marginLeft: 12 },
    { runs: [{ text: 'No direct left margin' }], marginLeft: false },
  ],
  { paragraphMarginLeft: 24 },
);
```

`paragraphMarginLeft` 是 plain/rich 创建时的 paragraph 默认值；`marginLeft` 是结构化 paragraph 的本地值。两者以 point 为单位，允许 `0..4032`，写入时四舍五入到最近 EMU。rich paragraph 的 number 覆盖外层默认，`false` 抑制外层默认并不写 direct `marL`。

对已有 shape 设置 `shape.richText` 时没有外层默认：number 写入 direct 值，`false` 或字段缺失都清除该 paragraph 的 direct 非列表 `marL`。getter 从不返回 `false`；它对合法 direct 值返回 point number，并且返回的 snapshot 不与模型内部状态共享。

新建的普通非列表 paragraph 在未配置时继续生成 PptxGenJS 风格的 direct `marL="0"` / `indent="0"`，因此 getter 会对该创建结果暴露 `marginLeft: 0`。显式 `false` 则区分“没有 direct 左边距”与“direct 零边距”。

## 方案选择

考虑过三种公共表面：

1. 直接暴露 `marL` 与 EMU number。它与 OOXML 一一对应，但把 schema 缩写和底层单位泄漏到高层文本模型。
2. 一次性公开包含 left/right/first-line/hanging 的嵌套 indentation 对象。形状完整，但会把三个尚未实现、与 bullet 交互不同的字段提前塞进当前小项。
3. 创建默认使用 `paragraphMarginLeft`，结构化段落使用 `marginLeft`，统一使用 point；采用此方案。它将 text-box `margin` 与 paragraph margin 明确区分，改动最小，也允许后续以平行字段逐项增加右边距、首行与悬挂缩进。

PptxGenJS 4.0.1 没有公开非列表 paragraph 左边距选项；它的 `margin` 是 text-box inset，`bullet.indent` 同时控制 bullet 与 `marL`，`indentLevel` 也只对列表生效。本小项因此是完整 PowerPoint 创建/编辑能力，不伪造 PptxGenJS 中不存在的入口；但会保持其普通 paragraph 的 direct zero 默认输出。

## OOXML 映射与读取

```xml
<a:p>
  <a:pPr marL="304800" indent="0">
    <a:buNone/>
  </a:pPr>
  ...
</a:p>
```

24 point 写为 `Math.round(24 * 12700) = 304800` EMU。`a:pPr@marL` 必须是十进制整数且位于 `0..51206400` EMU，对应 `0..4032` point。保存后 getter 用 `EMU / 12700` 返回 point；重新写入可恢复同一整数 EMU。

getter 只读取 direct `pPr@marL`，不解析 `lstStyle`、layout、master 或 theme 继承。缺失、空、非整数、负数或超范围 token 返回 undefined 并保留原 XML。如果该 `pPr` 含有 direct `buChar`、`buAutoNum` 或 `buBlip`，getter 不暴露 `marginLeft`，因为同一 `marL` 此时属于列表 hanging-indent 模型，已通过 `paragraph.bullet.indent` 表达。`buNone` 仍是非列表 paragraph，可正常读取左边距。

## 列表隔离与解析顺序

`marL` 同时被 bullet/numbering 作为文本起点，因此本小项只公开非列表 margin。任何一个 paragraph 在解析外层默认和本地 override 后，如果同时得到 active bullet/numbering 与 number `marginLeft`，则在 package mutation 前抛出 TypeError。`marginLeft: false` 可用于抑制外层 margin，从而允许同一 rich-text shape 中的某些 paragraph 使用 margin，另一些 paragraph 使用 bullet。

渲染时先处理 bullet 及其现有 `marL`/`indent` 对，再仅对非列表 paragraph 更新或删除 `marL`。这保证 bullet 分支不会被后续 margin clear 破坏，也保证从 bullet 改为普通 paragraph 时不会把旧 bullet anchor 误当作新的普通左边距。

`indent`、`lvl`、`marR`、alignment、RTL、spacing、tab stops、run formatting、`defRPr`、`extLst` 和未知 attributes/children 均保持现有语义。本小项不会自动把左边距复制为右边距，也不会根据 paragraph RTL 更名或交换字段。

## Lossless 编辑与事务边界

rich-text 替换继续按 paragraph index 复用 `pPr` template。对非列表 paragraph，number 替换/插入 direct `marL`，`false` 或 setter 中缺失字段只删除 direct `marL`；其他属性与 children 保留原顺序。读取、plain `.text` 替换、transform 和其他非 rich-text mutation 不触碰原 `marL`；plain `.text` 继续复制第一 paragraph template，因此增减换行时保留左边距。

所有验证在添加 shape 或写 part 之前完成。`paragraphMarginLeft` / `marginLeft` number 必须是 finite number，位于 `0..4032`；negative、NaN、Infinity、string、null、object、array、symbol 或与 active bullet 冲突都明确失败。失败不得留下 shape、part bytes、mutation journal 或 live model identity 变化；外层 transaction rollback 同样恢复 exact bytes。

## 文档与兼容状态

PptxGenJS 兼容矩阵把原“paragraph 左右 margin、first-line/hanging indent”行改为部分支持：左边距已支持，其余三项继续待实现。API 文档和聚合包 README 明确区分 text-box margin、paragraph left margin 与 bullet indent。changelog 只宣称当前 direct 左边距能力，不宣称整个 paragraph indentation 已完成。

## 测试与验收

1. plain 创建覆盖 omitted/zero/positive/fractional/boundary、CR/LF 多 paragraph、空 paragraph，并与 alignment、RTL、spacing 和 tab stops 组合。
2. rich 创建覆盖外层默认、number override、`false`、空 paragraph，以及同 shape 中非列表 margin paragraph 与独立 bullet paragraph。
3. getter 严格读取合法 direct integer `marL`，对缺失/空/负数/超范围/非整数和 direct active bullet 省略字段，只读不产生 mutation。
4. `shape.richText` 可增加、替换、设为 zero 和清除左边距；plain `.text` 保留并复制 template margin，其他 paragraph XML 不变。
5. bullet/margin 冲突、非 number、NaN/Infinity、负数、越界值在 mutation 前失败；内外层 rollback 保持 bytes、journal 和 identity。
6. PptxGenJS 4.0.1 对照确认普通 paragraph direct zero 默认与 bullet `marL`/`indent` 不回归；adapter 仅从标准 OOXML 读取，不访问私有字段。
7. duplicate、write/reopen、六种 presentation format、全仓 typecheck/test、独立 performance、Node/browser/declarations/tarball smoke 与 CLI validation 全部通过。
8. LibreOffice headless 无修复打开并导出包含 0/12/24/48 point 左边距的真实 PPTX；文本起点递增，与手工构造的合法 OOXML 基准一致，页面无裁切或溢出。
