# Rich Text 删除线设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加 strike 创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 的 boolean、single strike 和 double strike，同时支持从已有 PPT 读取并重新写入三个合法 OOXML strike token。该能力与现有 font family、font size、bold、italic、text color、underline、soft break 和 paragraph options 组合。

本小项不包括 highlight、outline/glow、baseline/上下标、character spacing、语言/RTL/亚洲字体或 hyperlink。读取和非 rich-text mutation 不得破坏原始 run XML；这些未支持能力继续作为独立小项。

## API

```ts
type RichTextStrikeStyle = 'sngStrike' | 'dblStrike';

interface RichTextRunStyle {
  readonly strike?: boolean | RichTextStrikeStyle;
}

slide.addRichText([
  {
    runs: [
      { text: 'Single', style: { strike: true } },
      { text: ' Double', style: { strike: 'dblStrike' } },
      { text: ' Explicit none', style: { strike: false } },
    ],
  },
]);
```

`strike: true` 规范化为 `sngStrike`；`false` 写显式 `noStrike`，用于压制继承删除线。getter 对合法删除线返回规范化 token，对 `strike="noStrike"` 返回 `false`，没有本地 strike 时省略字段。snapshot 是新值，不暴露模型内部引用。

## 方案选择

考虑过三种方案：

1. 只支持 boolean。它覆盖常见 single strike，但无法表达 PptxGenJS 已公开的 double strike，也无法忠实读取 PowerPoint 双删除线。
2. 使用 `single` / `double` 等新名称。它更口语化，但会制造与 PptxGenJS 和 OOXML 不同的第三套 token，增加迁移与映射成本。
3. 接受 PptxGenJS 公开且 OOXML 合法的 `sngStrike` / `dblStrike`，再用 boolean 表达常用 single 和显式 none。它保持类型对等、无翻译歧义，因此采用。

字符串 `noStrike` 不作为公共输入；它与 `false` 重复。getter 仍能读取现有 OOXML 的 `noStrike` 并返回 `false`。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `strike: true` 写 `a:rPr@strike="sngStrike"`。
- `strike: false` 不写本地 strike。
- `strike: 'sngStrike'` 和 `'dblStrike'` 把 token 原样写入。
- 公共类型只允许 boolean、`sngStrike` 和 `dblStrike`。
- 运行时强行传入 `noStrike` 或未知字符串时会原样写入；本库使用 `false` 表达合法 `noStrike`，拒绝其他字符串，避免无效 OOXML。

本库显式写出 false，与 PptxGenJS 的省略行为不同：省略会继续继承 layout/master 的 strike，而 `noStrike` 能稳定表达调用方要求的关闭状态。

## OOXML 映射

```xml
<a:rPr strike="dblStrike">
  <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
  <a:latin typeface="+mn-lt"/>
</a:rPr>
```

- `true` / `sngStrike` → `a:rPr@strike="sngStrike"`。
- `dblStrike` → `a:rPr@strike="dblStrike"`。
- `false` → `a:rPr@strike="noStrike"`。
- 未提供 strike → 不写本地属性。

getter 只读取 local `rPr@strike`：三个合法 token 分别返回 `sngStrike`、`dblStrike` 和 false；未知或畸形 token 不伪造成公共 strike。其他合法 run style 仍正常暴露。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 strike XML。plain `.text` 继续保留第一个 paragraph/run 的样式模板，因此原 strike 随模板复制。

`shape.richText` setter 整体替换 run，沿用既有边界：旧 run/rPr/field/hyperlink 不逐属性保留。新 run 只写调用方提供的公开 style；未提供 strike 就没有本地 strike，提供 false 则显式写 `noStrike`。paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。

## 验证与错误处理

`strike` 必须是 boolean、`sngStrike` 或 `dblStrike`。`null`、数字、对象、数组、`noStrike` 字符串和未知 token 明确失败。

所有验证在 package mutation 前完成。单个 run 失败时，不得留下部分 paragraph、part bytes、mutation journal、live shape text/style 或 snapshot 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. `addRichText()` 覆盖 true、false、single、double，以及与 underline、两种 text color、其他 run style 和 paragraph options 的组合。
2. getter 读取三个合法 token；未知 token 不伪造 strike，只读不产生 mutation。
3. `shape.richText` 增加、更换、清除和显式压制 strike，plain `.text` 保留第一 run strike，write/reopen 与 duplicate 一致。
4. null、数字、对象、数组、`noStrike` 字符串和未知 token 在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
5. PptxGenJS 4.0.1 conformance 覆盖 boolean true/false、single、double 和真实 `noStrike` OOXML 导入。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出 single/double/none 对照页；若客户端降级显示 double，以同版本 PptxGenJS 对照文件确认一致并记录限制。
