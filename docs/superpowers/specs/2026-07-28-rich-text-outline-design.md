# Rich Text 轮廓设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加 solid outline 创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 的 outline color 与 point size，并支持 sRGB/theme color。该能力与现有 font family、font size、bold、italic、text/highlight color、underline、strike、soft break 和 paragraph options 组合。

本小项不包括 outline transparency、dash/cap/join/compound/alignment、gradient/pattern outline、glow/shadow、baseline/上下标、character spacing、语言/RTL/亚洲字体或 hyperlink。读取和非 rich-text mutation 不得破坏这些原始 XML。

## API

```ts
interface RichTextOutline {
  readonly color: RichTextColor;
  readonly size: number;
}

interface RichTextRunStyle {
  readonly outline?: RichTextOutline;
}

slide.addRichText([
  {
    runs: [
      {
        text: 'Outlined',
        style: {
          outline: {
            color: { kind: 'srgb', value: 'FF0000' },
            size: 1.5,
          },
        },
      },
    ],
  },
]);
```

`size` 使用 point，与 PptxGenJS 一致；内部转换为 12,700 EMU/point，并按最近 EMU 规范化。color 复用 `RichTextColor`。getter 返回新的 outline/color 对象；没有本地 outline 时省略字段。

## 方案选择

考虑过三种方案：

1. 只暴露 hex color 与 size。它贴近 PptxGenJS 类型，但会让 outline color 与现有 text/underline/highlight color 模型不一致，也无法忠实读取 theme outline。
2. 直接复用完整 line codec。它未来可容纳 dash/cap/join，但会把本小项扩大为通用 DrawingML line API，并引入当前 rich text 不需要的依赖。
3. 新增窄 `RichTextOutline`，color 复用 `RichTextColor`，size 使用 point。它覆盖 PptxGenJS 的有效功能且边界清晰，因此采用。

不增加 boolean 或 `false`：OOXML text outline 没有显式 none token；整体替换时省略 outline 即清除本地轮廓。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `{ color: 'ff0000', size: 1.5 }` 写 `<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`。
- runtime 接受 `accent1` 等 theme token 并写 `schemeClr`。
- 公共类型要求 color 与 size，但运行时缺字段会默认 0.75pt 与白色；size 0 也会回退为 0.75pt。
- 负 size 会写负 EMU，产生非法 OOXML；非法 color 会 warning 后静默写黑色。

本库遵循公开必填字段，不采用运行时缺字段默认；size 0 作为合法 OOXML hairline 明确写 0，负值和非法 color 在 mutation 前失败。

## OOXML 映射

```xml
<a:rPr>
  <a:ln w="19050">
    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
  </a:ln>
  <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
  <a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
  <a:latin typeface="+mn-lt"/>
</a:rPr>
```

outline 写为 direct `a:ln`，仅有 `w` 属性和一个 direct `solidFill`；其 color choice 是一个 direct `srgbClr` 或 `schemeClr`。`a:ln` 位于普通 text fill 之前。

size 必须是 finite number，转换后的 EMU 在 `0..20,116,800`，对应 OOXML `ST_LineWidth` 的 `0..1584pt`。getter 只读取合法整数字段并返回 `w / 12700`。

getter 只暴露恰好一个 direct `a:ln`，并要求除 namespace declaration 外只含 `w` 属性、恰好一个 direct `solidFill`、一个受支持且无 transform 的 color choice。存在 dash/cap/join/compound/alignment、额外 line child、未知 fill、非法 width 或重复 line 时不伪造公共 outline；其他合法 run style 仍正常暴露。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 outline XML。plain `.text` 继续保留第一个 paragraph/run 的样式模板，因此原 outline 随模板复制。

`shape.richText` setter 整体替换 run：旧 run/rPr/field/hyperlink 不逐属性保留。新 run 只写调用方提供的公开 style；未提供 outline 就没有本地 line。paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。

## 验证与错误处理

outline 必须是只含 `color` 与 `size` 的 object，并且两项都必须提供。color 使用现有规则；size 必须 finite 且转换后位于合法 OOXML 范围。string/boolean/null/array、空对象、缺字段、额外字段、负值、超范围和非法 color 明确失败。

所有验证在 package mutation 前完成。单个 run 失败时，不得留下部分 paragraph、part bytes、mutation journal、live shape text/style 或 snapshot 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. `addRichText()` 覆盖 size 0、fractional/maximum size、sRGB、scheme，以及与 highlight/underline/strike/text color 和 paragraph options 的组合。
2. getter 读取两种颜色与边界 width；重复 line、未知/多 fill、color transform、非整数/负/超限 width、额外属性或 line child 不伪造 outline，只读不产生 mutation。
3. `shape.richText` 增加、更换和清除 outline，snapshot 隔离，plain `.text` 保留第一 run outline，write/reopen 与 duplicate 一致。
4. 非 object、空对象、缺字段、额外字段、非 finite/负/超范围 size 和非法 color 在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
5. PptxGenJS 4.0.1 conformance 覆盖 sRGB、scheme、fractional size 与无 outline 的真实输出导入。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出 sRGB/theme outline 对照页；用同版本 PptxGenJS 对照文件核对渲染与 OOXML。
