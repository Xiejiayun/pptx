# Rich Text 基线与上下标设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加自定义 baseline、superscript 与 subscript 的创建、读取和整体替换能力。三者在 DrawingML 中共享同一个 `rPr@baseline` 属性，因此作为一个原子小项实现，并与现有 font、color、glow、highlight、outline、underline、strike、soft break 和 paragraph options 组合。

本小项不包括 character spacing、kerning、font scale、语言/RTL/亚洲字体、hyperlink、vertical text 或 paragraph line layout。读取和非 rich-text mutation 不得破坏这些原始 XML。

## API

```ts
export type RichTextBaseline = number | 'superscript' | 'subscript';

interface RichTextRunStyle {
  readonly baseline?: RichTextBaseline;
}

slide.addRichText([
  {
    runs: [
      { text: 'x', style: { baseline: 'superscript' } },
      { text: '2', style: { baseline: 12.5 } },
      { text: ' normal', style: { baseline: 0 } },
    ],
  },
]);
```

number 表示相对正常基线的百分比：正值上移，负值下移。`superscript` 规范化为 30%，`subscript` 规范化为 -40%。显式 `0` 写 `baseline="0"`，用于清除或覆盖继承的上下标效果；省略字段才是不写本地 baseline。

getter 对标准 `30000`/`-40000` token 返回 named 值，其他合法值返回百分比 number。snapshot 为深拷贝。

## 方案选择

考虑过三种方案：

1. 原样公开 `baseline`、`superscript`、`subscript` 三个字段。它贴近 PptxGenJS，但允许三个字段互相冲突，还需要复制其 truthy 优先级和 baseline 0 丢失行为。
2. 只公开数值 baseline。它最小且覆盖 OOXML，但常见上下标调用需要记忆 30%/-40% 固定值。
3. 单一 `baseline` union，包含百分比 number 与两个 named 值。它既消除冲突又提供常用便利语义，因此采用。

不增加 boolean/false：显式 0 已能表达本地 normal baseline，省略字段表达未指定。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `superscript: true` 写 `baseline="30000"`，即 30%。
- `subscript: true` 写 `baseline="-40000"`，即 -40%。
- 自定义 `baseline` runtime 写 `Math.round(value * 50)`；例如 600 写 30000，1.5 写 75。
- custom baseline 优先于 subscript，subscript 优先于 superscript。
- baseline 0 因 truthy 判断被省略；false 脚本也被省略。
- 非 finite 或过大值可生成非法属性，runtime 不进行公共边界校验。

PptxGenJS 的自定义 baseline 单位没有公开说明且与 OOXML 百分比不直观。本库使用明确百分比，不复制 `*50` 参数语义；迁移时 PptxGenJS 600 对应本库 30 或 `superscript`，-800 对应 -40 或 `subscript`。adapter 读取真实 OOXML 后返回本库的语义值。

## OOXML 映射

```xml
<a:rPr baseline="30000">
  <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
  <a:latin typeface="+mn-lt"/>
</a:rPr>
```

number 乘 1000 后按最近整数写 `baseline`，即 OOXML 千分之一百分比。原始属性是 Int32，因此合法范围为 `-2,147,483,648..2,147,483,647`，对应输入百分比 `-2,147,483.648..2,147,483.647`。

getter 只接受严格十进制整数且位于 Int32 范围内的单一 baseline 属性；缺失时省略字段。`30000` 返回 `superscript`，`-40000` 返回 `subscript`，其他值除以 1000 返回 number。科学计数、decimal、越界或空值不伪造公共 baseline。

## 编辑与 Lossless 边界

仅 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不修改原始 baseline XML。plain `.text` 继续保留第一个 paragraph/run 的样式模板，因此原 baseline 随模板复制。

`shape.richText` setter 整体替换 run。新 run 只写调用方提供的公开 style；未提供 baseline 时没有本地 baseline，提供 0 时明确写 0。paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和其他 paragraph-level XML 继续按现有规则保留。

## 验证与错误处理

baseline 必须是 `superscript`、`subscript` 或 finite number。number 转换后的千分之一百分比必须落在 Int32 范围。其他 string、boolean、null、object、array、非 finite 或超范围值明确失败。

所有验证在 package mutation 前完成。单个 run 失败时，不得留下部分 paragraph、part bytes、mutation journal、live shape text/style 或 snapshot 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. `addRichText()` 覆盖 superscript、subscript、0、正负 fractional/custom 和 Int32 边界，以及与 glow/outline/underline/strike 等样式组合。
2. getter 读取 named、0、custom 与边界 token；decimal、科学计数、空值和越界 token 不伪造 baseline，只读不产生 mutation。
3. `shape.richText` 增加、更换和清除 baseline，plain `.text` 保留第一 run baseline，write/reopen 与 duplicate 一致。
4. 非法类型、未知 string、非 finite 和超范围值在 mutation 前失败；内外层 rollback 不改变 bytes、journal 或对象 identity。
5. PptxGenJS 4.0.1 conformance 覆盖 superscript、subscript、custom positive/negative、0 与无 baseline 的真实输出导入。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出上下标/custom baseline 对照页；用同版本 PptxGenJS 对照文件核对渲染与 OOXML。
