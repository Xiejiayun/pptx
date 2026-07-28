# Paragraph RTL 设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为 plain text 与结构化 rich text 增加 paragraph 级从右到左方向的创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 `TextPropsOptions.rtlMode` 的有效输出，同时提供同一文本框内逐段覆盖和显式 LTR suppression。

本小项只处理 direct `a:pPr@rtl`。不修改 `a:bodyPr@rtlCol`，不实现 presentation-level `pptx.rtlMode` / `p:presentation@rtl`，不自动改变 alignment，不做双向文本重排，也不公开 run-level RTL。语言、文本方向和 RTL 是独立属性。

## API

```ts
interface AddTextOptions {
  readonly rtlMode?: boolean;
}

interface RichTextParagraph {
  readonly rtl?: boolean;
}

slide.addText('مرحبا\nالعالم', { rtlMode: true });

slide.addRichText([
  { rtl: true, runs: [{ text: 'مرحبا' }] },
  { rtl: false, runs: [{ text: 'English override' }] },
], { rtlMode: true });
```

`AddTextOptions.rtlMode` 保留 PptxGenJS 创建入口名称，是每个 plain/rich paragraph 的创建默认值。`RichTextParagraph.rtl` 表达实际 paragraph property；存在时覆盖 outer 默认。`shape.richText` setter 没有 outer 默认：字段省略清除该 paragraph 的 direct RTL override，true/false 分别写显式 RTL/LTR。

## 方案选择

考虑过三种方案：

1. outer 使用 `rtlMode`、paragraph 使用 `rtl`；采用此方案。迁移入口与 PptxGenJS 对齐，结构化模型名称则准确对应 direct paragraph state。
2. 两层都使用 `rtlMode`。命名一致，但会把 PptxGenJS 的 shape-option 命名泄漏到 paragraph snapshot。
3. 使用 `direction: 'ltr' | 'rtl'`。语义清楚，但会制造不必要的迁移改名，且 OOXML 属性本身是 boolean。

不把 RTL 放进 `RichTextRunStyle`。PptxGenJS 对 TextProps runs 的 `rtlMode` 会在同一 `a:p` 中插入多个 `a:pPr`，只有第一个位置合法，outer 值也无法可靠传播；原生 API 以 paragraph 为最小单位，避免复制该无效结构。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- plain outer `rtlMode: true` 对每个 CR/LF paragraph 写 `a:pPr rtl="1"`。
- false 或省略不写 `rtl`。
- truthy 非 boolean 值也会写 `rtl="1"`；runtime 不校验类型。
- TextProps run 上的 true 会尝试写 paragraph properties；多个 runs 可在一个 `a:p` 中产生多个 `a:pPr`，后续节点位置不合法。
- rich TextProps[] 的 outer `rtlMode` 不能可靠传播到 runs，因为 paragraph properties 在 outer options 回填前生成。
- presentation instance 另有 `pptx.rtlMode`，写 `p:presentation@rtl="1"`；它不是本小项的 paragraph API。

本库匹配 boolean true 的有效 paragraph 输出，严格拒绝非 boolean，并提供 PptxGenJS 缺失的逐段 override。显式 false 写 `rtl="0"`，而非静默省略，以便压制 master/layout/presentation 继承；这是有意的 lossless editing 改进。

## OOXML 映射

```xml
<a:p>
  <a:pPr rtl="1"/>
  ...
</a:p>
<a:p>
  <a:pPr rtl="0"/>
  ...
</a:p>
```

创建时每段解析 `paragraph.rtl ?? defaultRtl`。true 写 `1`，false 写 `0`，undefined 不写 direct attribute。属性通过现有 `updateParagraphAttribute()` 更新，保持其他 `pPr` attributes、ordered children 与未知 XML。

getter 只读取 direct `pPr@rtl`：`1`、`true`、`on` 返回 true，`0`、`false`、`off` 返回 false；缺失、空值或未知 token 返回 undefined 并原样保留。getter 不解析 master/layout/list style/presentation 继承，也不根据文本字符或 alignment 猜测方向。

## 编辑与 Lossless 边界

只读、plain `.text`、transform 和非 rich-text mutation 不修改 `rtl`。plain `.text` 继续复制 paragraph template，因此原 direct direction 随模板保留。

`shape.richText` setter 继续整体替换公开 paragraph/run 内容并保留同位置 paragraph 模板。因为 RTL 在本小项后成为公开 paragraph property，setter 的 true/false 替换 direct token，省略字段删除 direct `rtl`；其他 `pPr` attributes/children、`endParaRPr`、`bodyPr`、`lstStyle` 和未知 text-body children 保持现有规则。

`bodyPr@rtlCol` 与 paragraph `rtl` 可同时存在，互不读取或修改。direction 不自动交换 left/right alignment；调用方需要时显式设置 `align`。

## 验证与错误处理

outer `rtlMode` 与 paragraph `rtl` 只接受 boolean。null、number、string、object、array 或 symbol 在任何 package mutation 前抛出 TypeError。false 必须作为有效值保留，不能被 truthy 判断吞掉。

任一 paragraph 失败时，不得留下 shape、part bytes、mutation journal、snapshot 或 identity 变化；外层 transaction rollback 同样恢复。

## 测试与验收

1. plain 创建覆盖 omitted/true/false、多段落、空段落和与 align/bullet/spacing/tabStops 的组合。
2. rich 创建覆盖 outer default、paragraph true/false override、字段省略继承和同文本框混合方向。
3. getter 严格读取六种 boolean lexical forms；缺失/空/未知不伪造值，只读不产生 mutation。
4. rich setter 增加、更换、显式 false 和清除 RTL，保留其他 paragraph XML；plain `.text` 保留模板 RTL。
5. duplicate、write/reopen、transaction rollback 和 invalid-input mutation isolation 全部通过。
6. PptxGenJS 4.0.1 conformance 覆盖 valid outer true、false/omitted、非法 truthy 基线和 run-level bug 隔离。
7. 全仓 typecheck/test、独立 performance、Node/browser/declarations/tarball smoke 与 CLI validate 全部通过。
8. LibreOffice headless 无修复打开并导出 Arabic/Hebrew 与 LTR override 对照页；原生 true 与 PptxGenJS 有效 paragraph XML 一致。
