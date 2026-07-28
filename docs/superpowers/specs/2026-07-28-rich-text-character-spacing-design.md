# Rich Text 字符间距设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加字符间距的创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 `charSpacing` 的正值、负值与 fractional point，并支持显式 0 清除继承间距。该能力与现有 baseline、font、color、glow、highlight、outline、underline、strike、soft break 和 paragraph options 组合。

本小项不公开 kerning threshold、font scale、word spacing、语言/RTL/亚洲字体、hyperlink 或 paragraph spacing。读取和非 rich-text mutation 不得破坏这些原始 XML。

## API

```ts
interface RichTextRunStyle {
  readonly characterSpacing?: number;
}

slide.addRichText([
  {
    runs: [
      { text: 'Expanded', style: { characterSpacing: 2.5 } },
      { text: ' Condensed', style: { characterSpacing: -1.25 } },
      { text: ' Normal', style: { characterSpacing: 0 } },
    ],
  },
]);
```

`characterSpacing` 使用 point，与 `fontSize` 和 PptxGenJS 一致。正值扩展、负值压缩；显式 0 写本地 normal spacing，省略字段才是不指定本地间距。getter 返回量化后的 point number。

## 方案选择

考虑过三种方案：

1. `characterSpacing?: number`。名称明确区分 run 字符间距与 paragraph `spacing`，单位沿用 point；采用此方案。
2. 原样使用 PptxGenJS `charSpacing`。迁移改名最少，但公共 API 缩写不如完整名称清晰。
3. 暴露 `{ spacing, kerning }` 对象。它能控制 `kern`，但会提前引入独立 kerning 语义与继承规则，超出本小项。

不增加 boolean/false：显式 0 已能覆盖继承间距，省略字段表达未指定。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `charSpacing: 2.5` 写 `spc="250" kern="0"`。
- `charSpacing: -1.25` 写 `spc="-125" kern="0"`。
- fractional value 按最近百分之一 point 取整；例如 0.004 写 `spc="0" kern="0"`。
- 精确 0 因 truthy 判断被省略，无本地 `spc`/`kern`。
- 与 baseline 组合时两个属性独立写入。
- 非 finite 或过大值可生成非法属性，runtime 不做边界校验。

本库保留有效输出与 `kern="0"` 的视觉语义，但修正精确 0 无法覆盖继承值的缺陷；非法值在 mutation 前失败。

## OOXML 映射

```xml
<a:rPr spc="250" kern="0">
  <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
  <a:latin typeface="+mn-lt"/>
</a:rPr>
```

number 乘 100 后按最近整数写 `spc`，即百分之一 point。原始属性是 Int32，合法范围为 `-2,147,483,648..2,147,483,647`，对应 point `-21,474,836.48..21,474,836.47`。写入 character spacing 时同时写 `kern="0"`，与 PptxGenJS 一致，避免 kerning 抵消扩展间距。

getter 只根据 direct `rPr@spc` 读取字符间距：属性必须是严格十进制整数并位于 Int32 范围，返回 `raw / 100`。`kern` 是独立、尚未公开的属性；缺少 `spc` 时即使存在 kern 也不返回 character spacing。合法 spc 与任意独立 kern 可同时读取，getter 不修改原 XML。

科学计数、decimal、空值或越界 spc 不伪造公共 character spacing，其他合法 run style 仍正常暴露。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 spc/kern XML。plain `.text` 继续保留第一个 paragraph/run 的样式模板，因此原属性随模板复制。

`shape.richText` setter 整体替换 run。新 run 提供 character spacing 时写量化 spc 与 kern 0；未提供时不写两者。paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。

## 验证与错误处理

character spacing 必须是 finite number，转换后的百分之一 point 必须落在 Int32 范围。string、boolean、null、object、array、非 finite 或超范围值明确失败。

所有验证在 package mutation 前完成。单个 run 失败时，不得留下部分 paragraph、part bytes、mutation journal、live shape text/style 或 snapshot 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. `addRichText()` 覆盖正/负/0、fractional 与 Int32 边界，以及与 baseline/glow/outline/underline/strike 的组合。
2. getter 读取正/负/0、边界和带非零 kern 的合法 spc；decimal、科学计数、空值与越界 spc 不伪造字段，单独 kern 不伪造字段，只读不产生 mutation。
3. `shape.richText` 增加、更换和清除 character spacing，plain `.text` 保留第一 run 属性，write/reopen 与 duplicate 一致。
4. 非法类型、非 finite 和超范围值在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
5. PptxGenJS 4.0.1 conformance 覆盖 positive、negative、fractional、0、与 baseline 组合及无 spacing 的真实输出导入。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出 expanded/condensed 对照页；用同版本 PptxGenJS 对照文件核对渲染与 OOXML。
