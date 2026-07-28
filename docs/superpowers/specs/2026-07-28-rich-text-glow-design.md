# Rich Text 发光设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加 solid glow 创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 的 glow color、point size 与 opacity，并支持 sRGB/theme color。该能力与现有 font family、font size、bold、italic、text/highlight color、outline、underline、strike、soft break 和 paragraph options 组合。

本小项不包括 outer/inner shadow、reflection、soft edge、blur、effect DAG、gradient/pattern glow、额外 color transform、baseline/上下标、character spacing、语言/RTL/亚洲字体或 hyperlink。读取和非 rich-text mutation 不得破坏这些原始 XML。

## API

```ts
interface RichTextGlow {
  readonly color?: RichTextColor;
  readonly opacity: number;
  readonly size: number;
}

interface RichTextRunStyle {
  readonly glow?: RichTextGlow;
}

slide.addRichText([
  {
    runs: [
      {
        text: 'Glowing',
        style: {
          glow: {
            color: { kind: 'scheme', value: 'accent1' },
            opacity: 0.5,
            size: 8,
          },
        },
      },
    ],
  },
]);
```

`size` 使用 point，`opacity` 使用 `0..1`，与 PptxGenJS 一致。color 复用 `RichTextColor`；省略时规范化为白色 sRGB。getter 返回新的 glow/color 对象；没有本地 glow 时省略字段。

## 方案选择

考虑过三种方案：

1. 直接复制 PptxGenJS 的 hex-string color 对象。它最接近旧参数形状，但不能忠实读取 theme glow，也与现有 text/underline/highlight/outline color 模型不一致。
2. 立即公开通用 DrawingML effect list。它可同时容纳 shadow/reflection/blur，但会把一个 run glow 小项扩大为 shape/run 共用 effect graph，并要求现在就决定尚未验证的 effect 组合语义。
3. 新增窄 `RichTextGlow`，复用 `RichTextColor`，并明确 size/opacity 量化。它覆盖 PptxGenJS 的有效公开能力且保持独立边界，因此采用。

color 按 PptxGenJS 公共类型保留为可选并默认白色；size 与 opacity 按公共类型必填。整体替换时省略 glow 即清除本地 glow。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `{ color: 'ff0000', size: 8, opacity: 0.5 }` 写 `<a:glow rad="101600"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:glow>`。
- runtime 接受 `accent1` 等 theme token 并写 `schemeClr`；fractional size 按最近 EMU 取整。
- 省略 color 默认白色；尽管公共类型要求 size/opacity，runtime 缺字段会默认 8pt/0.75。
- size 0 和 opacity 0 会分别写 `rad="0"` 与 `alpha val="0"`。
- 负 size、越界 opacity 与非 finite 值会生成非法属性；非法 color 会 warning 后静默写黑色。

本库遵循公共必填字段，不采用 runtime 的缺字段默认；保留公开的可选 color 白色默认。负值、越界、非 finite 和非法 color 在 mutation 前失败，不复制非法输出或静默回退。

## OOXML 映射

```xml
<a:rPr>
  <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
  <a:effectLst>
    <a:glow rad="101600">
      <a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr>
    </a:glow>
  </a:effectLst>
  <a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
  <a:latin typeface="+mn-lt"/>
</a:rPr>
```

glow 写为一个 direct `a:effectLst` 内的单一 direct `a:glow`。`rad` 是 `size * 12,700` 后的最近整数；合法 EMU 范围为 `0..27,273,042,316,900`，对应 `ST_PositiveCoordinate` 与 `0..2,147,483,647pt`。opacity 写为 color choice 内单一 `a:alpha val`，按最近十万分之一量化到 `0..100000`。

`effectLst` 位于文字普通 fill 后、highlight 与 underline fill 前。color choice 是 direct `srgbClr` 或 `schemeClr`。getter 也接受没有 alpha transform 的合法 glow，并将其解释为 opacity 1。

getter 要求恰好一个 direct `effectLst` 和其中恰好一个 direct `glow`；glow 除 namespace declaration 外只含 `rad` 属性，并只含一个受支持的 direct color choice。color choice 必须无 transform，或只含一个合法 direct alpha。重复 glow/effect list、effect DAG、缺失/非整数/负/超限 radius、未知 color、重复/非法 alpha 或其他 transform 不伪造公共 glow。合法 glow 可与同一 effect list 中未公开的其他 effect 共存，读取 glow 不删除或重写 sibling effect。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 glow/effect XML。plain `.text` 继续保留第一个 paragraph/run 的样式模板，因此原 glow 随模板复制。

`shape.richText` setter 整体替换 run：旧 run/rPr/field/hyperlink/effects 不逐属性保留。新 run 只写调用方提供的公开 style；提供 glow 时写仅包含 glow 的 effect list，未提供时没有本地 effect list。paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。

## 验证与错误处理

glow 必须是只含 `color`、`opacity` 与 `size` 的 object。size 与 opacity 必须提供且 finite；size 转换后的 EMU 必须在 `ST_PositiveCoordinate` 范围，opacity 必须在 `0..1`。color 省略时使用白色，提供时使用现有严格规则。

string/boolean/null/array、空对象、缺 size/opacity、额外字段、非 finite/负/超范围值和非法 color 明确失败。所有验证在 package mutation 前完成。单个 run 失败时，不得留下部分 paragraph、part bytes、mutation journal、live shape text/style 或 snapshot 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. `addRichText()` 覆盖 size/opacity 的 0、fractional 与 maximum，默认/sRGB/scheme color，以及与 highlight/outline/underline/strike/text color 和 paragraph options 的组合。
2. getter 读取两种颜色、无 alpha 与边界值；重复 effect list/glow、effect DAG、未知 color、额外 transform、非整数/负/超限 radius、重复/非法 alpha 不伪造 glow，只读不产生 mutation；合法 sibling effect 原样保留。
3. `shape.richText` 增加、更换和清除 glow，snapshot 隔离，plain `.text` 保留第一 run glow，write/reopen 与 duplicate 一致。
4. 非 object、空对象、缺字段、额外字段、非 finite/负/超范围 size/opacity 和非法 color 在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
5. PptxGenJS 4.0.1 conformance 覆盖 sRGB、scheme、默认 color、fractional size、0 opacity 与无 glow 的真实输出导入。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出 sRGB/theme glow 对照页；用同版本 PptxGenJS 对照文件核对渲染与 OOXML。
