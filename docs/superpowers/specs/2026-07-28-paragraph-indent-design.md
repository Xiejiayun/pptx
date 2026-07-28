# Paragraph Indent 设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为 plain text 与结构化 rich text 增加普通 paragraph 首行缩进和悬挂缩进的创建、读取、整体替换与清除能力，映射 direct `a:pPr@indent`。OOXML 使用同一个 signed attribute 表达两种语义：正值是首行缩进，负值是悬挂缩进，因此两者必须作为一个原子小项实现。只支持非负 first-line indent 会在 `shape.richText = shape.richText` 时误清除已有负值，破坏 lossless 编辑。

本小项不解析 master、layout 或 `a:lstStyle` 继承，不根据 RTL 交换语义，不改变 bullet/numbering 已有的 `marL`/negative `indent` 算法，也不增加 picture-bullet 写入 API。`indent` 与已完成的 direct 左右 margin 共同覆盖普通 paragraph 的直接边界和缩进属性。

## API

```ts
interface AddTextOptions {
  readonly paragraphIndent?: number;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly indent?: number | false;
}

slide.addText('First-line indent', {
  paragraphIndent: 24,
});

slide.addRichText(
  [
    { runs: [{ text: 'Uses 24pt outer default' }] },
    { runs: [{ text: '18pt first-line indent' }], indent: 18 },
    { runs: [{ text: '18pt hanging indent' }], indent: -18 },
    { runs: [{ text: 'No direct indent' }], indent: false },
  ],
  { paragraphIndent: 24 },
);
```

`paragraphIndent` 是 plain/rich 创建时的 paragraph 默认值；`indent` 是结构化 paragraph 的本地值。两者以 point 为单位，允许 `-4032..4032`。正数表示首行相对 `marL` 向内移动，负数表示首行相对其余行向外形成悬挂缩进，zero 表示 direct 零缩进。写入值为 `Math.round(points * 12700)` EMU。

rich paragraph 的 number 覆盖外层默认，`false` 抑制外层默认并不写 direct `indent`。对已有 shape 设置 `shape.richText` 时没有外层默认：number 写入 direct 值，`false` 或字段缺失都清除该 paragraph 的 direct `indent`。getter 从不返回 `false`；它对合法 direct 值返回 signed point number，并且返回的 snapshot 不共享模型内部状态。

新建的普通非列表 paragraph 在未配置时继续生成 PptxGenJS 风格的 canonical `indent="0"`，因此 getter 返回 `indent: 0`。显式 `false` 则移除该默认 attribute，以区分“没有 direct 缩进”和“direct 零缩进”。新建 bullet/numbering paragraph 继续由列表模型写入自己的 negative `indent`，paragraph getter 不把它暴露为本 API。

## 方案选择

考虑过三种公共表面：

1. 分开公开 `firstLineIndent` 与 `hangingIndent` 两个非负字段。调用语义直观，但两者竞争同一个 OOXML attribute，必须定义额外互斥规则，而且分阶段实现任一字段都会破坏另一种 direct 值。
2. 公开 signed `paragraphIndent` / `indent` scalar；采用此方案。它与 OOXML 一一对应，和现有 `paragraphMarginLeft` / `marginLeft`、`paragraphMarginRight` / `marginRight` 命名保持一致，用正负号无损表达两种模式。
3. 暴露 `{ firstLine, hanging }` 嵌套对象并在内部转为 signed value。它可隐藏负号，但两个字段不能同时生效，增加无收益的冲突状态和迁移成本。

PptxGenJS 4.0.1 没有公开普通 paragraph first-line/hanging indent 选项；`bullet.indent` 和 `indentLevel` 只服务列表。本小项补齐完整 PowerPoint 创建/编辑能力，不伪造 PptxGenJS 中不存在的入口；adapter 仍只从 PptxGenJS 生成的标准 OOXML 导入 direct 状态。

## OOXML 映射与读取

```xml
<a:p>
  <a:pPr marL="304800" marR="152400" indent="-228600">
    <a:buNone/>
  </a:pPr>
  ...
</a:p>
```

`18` point 写为 `228600` EMU，`-18` point 写为 `-228600` EMU。`a:pPr@indent` 必须是完整的 signed 十进制整数且位于 `-51206400..51206400` EMU，对应 `-4032..4032` point。保存后 getter 用 `EMU / 12700` 返回 point；重新写入可恢复同一整数 EMU。

getter 只读取 direct `pPr@indent`，不解析 `lstStyle`、layout、master 或 theme 继承。缺失、空、非整数或超范围 token 返回 `undefined` 并保留原 XML。若同一 `pPr` 含 direct `buChar`、`buAutoNum` 或 `buBlip`，getter 不暴露 `indent`，因为该 attribute 属于列表 hanging-indent 模型，当前 character/number bullet 已由 `paragraph.bullet.indent` 表达。`buNone` 不是 active bullet，可正常读取 ordinary paragraph indent。

## 列表、margin 与 RTL

active bullet/numbering 同时拥有 `marL` 与 `indent`。任一 paragraph 在解析外层默认和本地 override 后，如果同时得到 active bullet 与 numeric paragraph indent，则在 package mutation 前抛出 `TypeError`。`indent: false` 可抑制 outer `paragraphIndent`，使同一个 rich-text shape 中的普通段落使用外层缩进而列表段落继续使用 bullet indent；反向也可用 `bullet: false` 抑制 outer bullet 后使用 numeric paragraph indent。

渲染先完成现有 bullet 处理，再只对非列表 paragraph 更新或删除 ordinary `indent`。active bullet 分支绝不被后续 ordinary clear 破坏；从 bullet 改为普通 paragraph 且未提供 `indent` 时，现有列表清理逻辑把可确认的 bullet `marL`/`indent` pair 归零，随后 existing rich setter 的 omission 清除 direct `indent`，避免把旧 hanging indent 误解释为普通 paragraph 值。

ordinary `indent` 可以与 `marginLeft` 和 `marginRight` 同时存在；其首行位置按 OOXML 相对 direct/inherited `marL` 解释。本 API 不计算最终排版坐标，也不自动新增或修改 margin。即使 paragraph 为 RTL，仍写物理 direct `indent`，不交换 margin、不反转正负号。

## Lossless 编辑与事务边界

rich-text 替换继续按 paragraph index 复用 `pPr` template。number 替换或插入 direct `indent`；`false` 或 setter 中字段缺失只删除 direct `indent`；其他 attributes 与 children 保持原顺序。读取、plain `.text` 替换、transform 和其他非 rich-text mutation 不触碰原 `indent`；plain `.text` 继续复制第一 paragraph template，因此增减换行时保留普通或列表 indent。

所有 validation 与 resolved bullet/indent 冲突检查在添加 shape 或写 part 之前完成。`paragraphIndent` / `indent` number 必须是 finite number 且位于 `-4032..4032`；NaN、Infinity、string、boolean、null、object、array、symbol 或与 active bullet 冲突都明确失败。失败不得留下 shape、part bytes、mutation journal 或 live model identity 变化；外层 transaction rollback 同样恢复 exact bytes。

## 文档与兼容状态

PptxGenJS 兼容矩阵把 paragraph direct margin/indent 行更新为支持 left margin、right margin、first-line indent 和 hanging indent；列表仍通过独立的 `bullet.indent` / `level` API 表达。API 文档和聚合包 README 明确 signed point 语义、zero 与 false/absence 的区别、text-box margin 的层级差异及 bullet 冲突。changelog 只宣称 direct ordinary paragraph indent，不宣称继承解析或全部 PowerPoint text layout 已完成。

## 测试与验收

1. plain 创建覆盖 omitted、zero、positive、negative、fractional、正负边界、CR/LF 多 paragraph 与空 paragraph，并与左右 margin、alignment、RTL、spacing 和 tab stops 组合。
2. rich 创建覆盖 outer default、positive/negative/zero override、`false`、空 paragraph、`bullet: false`，以及 `indent: false` 抑制 outer default 后使用 bullet/numbering。
3. getter 严格读取合法 direct signed integer `indent`，对缺失、空、超范围和非整数省略字段；active character/number/picture bullet 抑制 ordinary `indent`，`buNone` 不抑制，读取不产生 mutation。
4. `shape.richText` 可增加、替换、正负切换、设为 zero 和清除 ordinary indent；plain `.text` 保留并复制 template indent，`marL`/`marR` 与其他 paragraph XML 不变。
5. 非 number、NaN/Infinity 和正负越界值在 mutation 前以 TypeError/RangeError 失败；numeric indent 与 resolved active bullet 冲突以 TypeError 失败；内外层 rollback 保持 bytes、journal 和 identity。
6. PptxGenJS 4.0.1 对照确认普通 paragraph 的 direct zero 与列表 negative `indent` 正确导入；ordinary getter 对列表值保持抑制，adapter 不访问私有字段。
7. duplicate、write/reopen、六种 presentation format、全仓 typecheck/test、独立 performance、Node/browser/declarations/tarball smoke 与 CLI validation 全部通过。
8. LibreOffice headless 无修复打开并导出包含 `-24/-12/0/12/24` point ordinary indent 的真实 PPTX；首行/后续行位置按 signed 值变化，与手工构造的合法 OOXML 基准一致，页面无裁切或溢出。
