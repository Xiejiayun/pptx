# Paragraph Right Margin 设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为 plain text 与结构化 rich text 增加 paragraph 右边距的创建、读取、整体替换和清除能力，映射 direct `a:pPr@marR`。这是 paragraph indentation 能力的第二个小项，与已完成的非列表左边距组成独立的左右边界控制，同时保持既有 lossless 编辑和事务语义。

本小项不包括 `a:pPr@indent` 表示的首行/悬挂缩进、master/layout/list-style 继承解析、RTL 语义换边，也不改变 bullet/numbering 已有的 `marL`/`indent` hanging-indent 算法。首行缩进与悬挂缩进继续作为后续独立小项。

## API

```ts
interface AddTextOptions {
  readonly paragraphMarginRight?: number;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly marginRight?: number | false;
}

slide.addText('Bounded paragraph\nSame default', {
  paragraphMarginRight: 24,
});

slide.addRichText(
  [
    { runs: [{ text: 'Uses 24pt default' }] },
    { runs: [{ text: 'Uses 12pt override' }], marginRight: 12 },
    { runs: [{ text: 'No direct right margin' }], marginRight: false },
    { runs: [{ text: 'Bullet with a right margin' }], bullet: true, marginRight: 18 },
  ],
  { paragraphMarginRight: 24 },
);
```

`paragraphMarginRight` 是 plain/rich 创建时的 paragraph 默认值；`marginRight` 是结构化 paragraph 的本地值。两者以 point 为单位，允许 `0..4032`，写入时四舍五入到最近 EMU。rich paragraph 的 number 覆盖外层默认，`false` 抑制外层默认并不写 direct `marR`。

对已有 shape 设置 `shape.richText` 时没有外层默认：number 写入 direct 值，`false` 或字段缺失都清除该 paragraph 的 direct `marR`。getter 从不返回 `false`；它对合法 direct 值返回 point number，并且返回的 snapshot 不与模型内部状态共享。

新建 paragraph 在未配置右边距时不新增 `marR`，因此 getter 返回 `undefined`。这与左边距现有的 PptxGenJS 风格 canonical `marL="0"` 默认不同；本小项不修改 `marL` 或 `indent` 的创建输出。显式 zero 写 `marR="0"`，从而与 false/absence 区分。

## 方案选择

考虑过三种公共表面：

1. 继续采用平行 scalar 字段：创建默认使用 `paragraphMarginRight`，结构化段落使用 `marginRight`，统一使用 point；采用此方案。它与已发布的左边距 API 对称，改动最小，也不会提前承诺首行/悬挂缩进的模型。
2. 将左右边距合并为 `paragraphMargins: { left, right }`。它能一次表达两侧，但会造成已发布 `paragraphMarginLeft` / `marginLeft` 的重复入口或破坏性迁移。
3. 直接暴露 `marR` 与 EMU number。它与 OOXML 一一对应，但会把 schema 缩写和底层单位泄漏到高层文本模型。

PptxGenJS 4.0.1 没有公开 paragraph `marR` 选项；它的 `margin` 是 text-box inset，`bullet.indent` 只控制 `marL` 与 `indent`。本小项因此补齐完整 PowerPoint 创建/编辑能力，不伪造 PptxGenJS 中不存在的入口；adapter 继续只导入 PptxGenJS 生成的标准 OOXML。

## OOXML 映射与读取

```xml
<a:p>
  <a:pPr marL="0" marR="304800" indent="0">
    <a:buNone/>
  </a:pPr>
  ...
</a:p>
```

24 point 写为 `Math.round(24 * 12700) = 304800` EMU。`a:pPr@marR` 必须是十进制整数且位于 `0..51206400` EMU，对应 `0..4032` point。保存后 getter 用 `EMU / 12700` 返回 point；重新写入可恢复同一整数 EMU。

getter 只读取 direct `pPr@marR`，不解析 `lstStyle`、layout、master 或 theme 继承。缺失、空、非整数、负数或超范围 token 返回 `undefined` 并保留原 XML。与 `marL` 不同，direct `buChar`、`buAutoNum` 或 `buBlip` 不占用 `marR`，所以合法 direct 右边距在这些 paragraph 上都可读取；本 API 可创建和编辑当前公开模型支持的普通、bullet 与 numbering paragraph，不借本小项扩展 picture-bullet 写入能力。

## 列表、RTL 与解析顺序

右边距独立于 bullet/numbering 的文本起点和 hanging indent。numeric `marginRight` 可与 active bullet/numbering 同段存在，不增加冲突错误，也不参与 bullet level 的 `marL = indent × (level + 1)` 计算。渲染先完成现有 bullet 处理，再独立更新或删除 `marR`；右边距 mutation 不读取、重算或清除 `marL`/`indent`。

`marginLeft`、`indent`、`lvl`、alignment、RTL、spacing、tab stops、run formatting、`defRPr`、`extLst` 和未知 attributes/children 均保持现有语义。即使 paragraph 为 RTL，本 API 仍映射物理 direct `marR`，不会自动更名、交换或复制左右边距。

## Lossless 编辑与事务边界

rich-text 替换继续按 paragraph index 复用 `pPr` template。number 替换/插入 direct `marR`，`false` 或 setter 中缺失字段只删除 direct `marR`；其他属性与 children 保留原顺序。读取、plain `.text` 替换、transform 和其他非 rich-text mutation 不触碰原 `marR`；plain `.text` 继续复制第一 paragraph template，因此增减换行时保留右边距。

所有验证在添加 shape 或写 part 之前完成。`paragraphMarginRight` / `marginRight` number 必须是 finite number，位于 `0..4032`；negative、NaN、Infinity、string、boolean、null、object、array 或 symbol 明确失败。失败不得留下 shape、part bytes、mutation journal 或 live model identity 变化；外层 transaction rollback 同样恢复 exact bytes。

## 文档与兼容状态

PptxGenJS 兼容矩阵继续标记 paragraph margin/indent 为部分支持，但更新为 direct 左右边距均已支持，仅 first-line/hanging indent 待实现。API 文档和聚合包 README 明确区分 text-box margin、paragraph 左右边距与 bullet indent。changelog 只宣称当前 direct 右边距能力，不宣称整个 paragraph indentation 已完成。

## 测试与验收

1. plain 创建覆盖 omitted/zero/positive/fractional/boundary、CR/LF 多 paragraph、空 paragraph，并与左边距、alignment、RTL、spacing 和 tab stops 组合。
2. rich 创建覆盖外层默认、number override、`false`、空 paragraph，以及普通、bullet、numbering paragraph 上的右边距。
3. getter 严格读取合法 direct integer `marR`，对缺失/空/负数/超范围/非整数省略字段；active bullet 不抑制合法 `marginRight`，读取不产生 mutation。
4. `shape.richText` 可增加、替换、设为 zero 和清除右边距；plain `.text` 保留并复制 template margin，`marL`/`indent` 与其他 paragraph XML 不变。
5. 非 number、NaN/Infinity、负数和越界值在 mutation 前以 TypeError/RangeError 失败；内外层 rollback 保持 bytes、journal 和 identity。
6. PptxGenJS 4.0.1 对照确认普通与列表 paragraph 默认不生成 `marR`，adapter 不访问私有字段，现有 `marL`/`indent` 行为不回归。
7. duplicate、write/reopen、六种 presentation format、全仓 typecheck/test、独立 performance、Node/browser/declarations/tarball smoke 与 CLI validation 全部通过。
8. LibreOffice headless 无修复打开并导出包含 0/12/24/48 point 右边距的真实 PPTX；右侧换行边界依次向左收缩，与手工构造的合法 OOXML 基准一致，页面无裁切或溢出。
