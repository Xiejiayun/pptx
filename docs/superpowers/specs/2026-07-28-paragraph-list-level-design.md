# 段落列表层级设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为普通文本和结构化 rich text 增加零基列表层级，使原生创建和已有 PPT 编辑覆盖 PptxGenJS 4.0.1 的 `indentLevel`，并让层级与已经支持的 standard/custom bullet、自动编号、hanging indent、alignment、spacing、空 paragraph 和 rich runs 正确组合。

本小项不包括任意 paragraph 左/右 margin、独立 first-line indent、tab stop、master/list-style level 样式编辑或超过 OOXML 九级限制的列表。这些继续作为独立小项；读取和非 rich-text mutation 不得破坏其原始 XML。

## API

```ts
interface AddTextOptions extends Partial<Transform> {
  readonly level?: number;
}

interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly level?: number;
}

slide.addText('Nested item', {
  bullet: true,
  level: 2,
});

slide.addRichText(
  [
    { runs: [{ text: 'Level one' }] },
    { runs: [{ text: 'Level two' }], level: 2 },
    { runs: [{ text: 'Back to root' }], level: 0 },
  ],
  { bullet: true, level: 1 },
);
```

`level` 是零基整数 0–8。`AddTextOptions.level` 是创建默认值；paragraph 未提供 `level` 时继承默认值，显式 `level: 0` 回到根层级并抑制非零默认。getter 对本地 `lvl=1..8` 返回 `level`，`lvl=0` 或属性缺失规范化为字段缺失。

对已有 shape 设置 `richText` 时没有文本框默认：未提供 `level` 表示目标 paragraph 不写本地 `lvl`，提供 0 同样移除属性。与 alignment、bullet 和 spacing 一致，显式 rich-text setter 替换本小项公开的 paragraph level。

## Bullet indent 语义

`ParagraphBullet.indent` 在 level 0 的行为完全不变：它同时决定 `marL` 的总左 margin 与负的 hanging `indent`。存在非零 level 时，同一个值作为每一级的 margin step：

```text
marL = bullet.indent × (level + 1)
indent = -bullet.indent
```

因此默认 27pt bullet 在 level 2 写 81pt 总 margin 和 -27pt hanging indent。getter 对 level 0 继续从 `marL` 读取 indent；对 level 1–8 从 `marL / (level + 1)` 返回规范化的 per-level indent。这样既保持已发布的根级 snapshot，又与 PptxGenJS 的嵌套列表语义对等。

组合后的总 `marL` 仍必须位于 0–4032pt。单独合法的 bullet indent 若与较深 level 组合后越界，会在 mutation 前失败。

## 方案选择

考虑过三种方案：

1. 直接复制 `indentLevel?: number`。迁移最直接，但名称容易和 first-line indent 混淆，也会延续 4.0.1 接受小数、字符串和 level 9 的缺陷。
2. 暴露零基 `level?: number`，并在内部把它与 bullet hanging step 组合。API 短、getter 稳定、可与未来普通 indent 分开，并完整覆盖 PptxGenJS 功能。因此采用。
3. 暴露原始 `lvl`、`marL`、`indent`。它能表达任意 OOXML margin 组合，但泄漏存储单位和互相约束，调用方很容易生成无效列表。

本库追求功能对等而不是复制 PptxGenJS 的运行时缺陷。adapter 读取 PptxGenJS 实际输出；原生 API 对非整数、字符串、负值、level 9+ 和组合 margin 越界明确失败。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- 默认 bullet、`indentLevel: 1` 写 `lvl="1" marL="685800" indent="-342900"`，即 54pt / -27pt。
- 18pt custom bullet、level 2 写 `lvl="2" marL="685800" indent="-228600"`，即 54pt / -18pt。
- 22pt number、level 3 写 `lvl="3" marL="1117600" indent="-279400"`，即 88pt / -22pt。
- 无 bullet、level 2 写 `lvl="2" indent="0" marL="0"` 和 `buNone`。
- level 0 与负值被当作根级并省略 `lvl`。
- 小数 1.5、字符串 `'2'` 和 9 会被写入 `lvl`；前两种违反公共类型，1.5 与 9 违反 OOXML `ST_TextIndentLevelType` 的整数 0–8 范围。本库全部拒绝。

## OOXML 映射

```xml
<a:pPr lvl="2" marL="1028700" indent="-342900">
  <a:buSzPct val="100000"/>
  <a:buChar char="•"/>
</a:pPr>
```

`level` 映射到 direct `pPr@lvl`。0 规范化为移除属性；1–8 写十进制整数。bullet margin 继续使用 point × 12700 转 EMU，并按上面的 level 公式计算。

getter 严格解析 `lvl`：只接受十进制整数 0–8。畸形值、负值、小数和超范围 token 不伪造公共 level。对受支持的 active bullet，level 会参与 per-level indent 还原；`buNone` 或无 bullet choice 不产生 bullet 值。

## Lossless patch 与清除

rich text 替换继续按 paragraph 索引选择 `pPr` 模板。实现只增加、替换或移除 `lvl` 属性，并复用 bullet choice 的 source-span patch。alignment、spacing、普通 margin/indent、tab、`defRPr`、`extLst`、未知属性和未知 child 保持原顺序与字节内容。

active bullet 写入时，`marL` 使用目标 level 计算，`indent` 始终是负的单级 step。清除既有 bullet 时，若原 margin/hanging 形成 1–9 倍的可确认层级 pair，则归零；无法确认属于 bullet 的自定义 margin/indent 继续保留。显式设置 level 但未设置 bullet 不会凭空创建 bullet。

自闭合 `pPr` 必须安全修改，所有属性和 child 更新继续使用 `LosslessXmlDocument`，不以正则重写 XML。

## 验证与错误处理

`level` 必须是 number、有限整数且位于 0–8。bullet indent 与 level 组合后的总 margin 不得超过 4032pt。所有 validation 和组合检查发生在 package mutation 前或 transaction 的 render 阶段；失败必须恢复 package bytes、mutation journal、live shape 内容和 level/bullet snapshot。

## 测试与验收

1. `addText()` 把默认 level 应用到每个 CR/LF paragraph，level 0–8 与默认/custom/number bullet 的 `lvl/marL/indent` 正确。
2. `addRichText()` 支持文本框默认、逐段 override、显式 0、无 bullet level、空 paragraph，并与 alignment 和 spacing 组合。
3. getter 正确读取 level，按 level 还原 bullet per-level indent，规范化 0，忽略畸形/超范围值，只读不产生 mutation。
4. `shape.richText` 可增加、更换或清除 level，同时保留 spacing、alignment、tab、`defRPr`、`extLst`、未知属性/child；bullet margin 随 level 正确重算。
5. 非 number、NaN/Infinity、小数、负值、9+ 和组合 margin 越界在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
6. PptxGenJS 4.0.1 conformance 覆盖 default/custom/number bullet、level 0–3 和无 bullet level；adapter 按实际 OOXML 读取。
7. write/reopen、duplicate、snapshot isolation、全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
8. LibreOffice headless 无修复打开并导出一页 level 0–3 的 bullet/numbered list PDF；嵌套层级逐级缩进且页面无截断或溢出，兼容矩阵把 `indentLevel` 标记为已支持。
