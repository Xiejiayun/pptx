# Rich Text 字色透明度设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为结构化 rich text run 增加文字主填充透明度的创建、读取和整体替换能力，覆盖 PptxGenJS 4.0.1 `TextPropsOptions.transparency` 的 `0..100` 百分比，并支持 sRGB、theme color 和当前默认文字颜色。该能力映射 run `a:rPr/a:solidFill` 内颜色节点的 direct `a:alpha` transform。

本小项只控制文字 glyph 的主 solid fill，不改变 highlight、outline、underline color、glow、shape fill 或 table-cell fill。它不增加 outer `AddTextOptions.transparency`、plain-text color/style defaults、gradient/pattern text fill、其他 color transform、继承解析或 table-cell API；这些继续按独立小项实现。

## API

```ts
interface RichTextRunStyle {
  readonly transparency?: number;
}

slide.addRichText([
  {
    runs: [
      {
        text: 'Quarter transparent',
        style: {
          color: { kind: 'srgb', value: 'FF0000' },
          transparency: 25,
        },
      },
      {
        text: ' Theme half transparent',
        style: {
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 50,
        },
      },
      {
        text: ' Default color fully transparent',
        style: { transparency: 100 },
      },
    ],
  },
]);
```

`transparency` 使用百分比：`0` 完全不透明，`100` 完全透明。允许 fractional number，并量化到最近 `0.001%`。省略字段表示没有 direct alpha transform；显式 0 写 direct fully-opaque alpha，因此 getter 可区分 explicit zero 与 absence。style 没有显式 `color` 时，透明度作用于本库现有的默认 `schemeClr tx1` 主填充。

## 方案选择

考虑过三种公共表面：

1. 在 `RichTextRunStyle` 增加 `transparency?: number`，范围 `0..100`；采用此方案。它与 PptxGenJS 名称、单位和迁移心智一致，同时只影响 run 主填充。
2. 增加 `opacity?: number`，范围 `0..1`。它与现有 glow opacity 一致，但与 PptxGenJS 文字 API 和 PowerPoint UI 的 transparency 百分比相反，迁移容易出错。
3. 把 alpha 放进 `RichTextColor`。它能复用颜色对象，但会同时改变 outline、highlight、underline 和 glow 的公共含义，并制造同一 color type 在不同位置支持程度不同的问题。

不增加 `false`：整体替换时省略 `transparency` 就不写 direct alpha；显式 0 已能表达本地完全不透明。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- `transparency: 25` 写 `alpha val="75000"`。
- `transparency: 50.5555` 写 `alpha val="49445"`，回读为 `50.555`。
- `transparency: 100` 写 `alpha val="0"`。
- theme `accent1` 与 transparency 组合时，alpha 是 `schemeClr` 的 direct child。
- run 未显式提供 color 时，PptxGenJS 把默认黑色实体化并写 alpha；本库沿用已有默认 `tx1`，透明度语义相同。
- `transparency: 0` 因 PptxGenJS truthy 判断而不写 alpha，和 omitted 输出相同。本库为可逆 direct editing 写 `alpha val="100000"`；导入 PptxGenJS 的 omitted/zero 均返回 `undefined`，effective appearance 仍为完全不透明。
- runtime 不严格拒绝负数、超过 100、string 或 non-finite 值，可能产生 schema-invalid alpha；本库在 mutation 前拒绝。

## OOXML 映射与量化

```xml
<a:rPr>
  <a:solidFill>
    <a:srgbClr val="FF0000">
      <a:alpha val="75000"/>
    </a:srgbClr>
  </a:solidFill>
  ...
</a:rPr>
```

写入公式是 `alpha = Math.round((100 - transparency) * 1000)`，合法 alpha 为 `0..100000`。normalizer 用该整数反算并保存 canonical `100 - alpha / 1000`，因此创建后的 snapshot、保存/reopen 和 PptxGenJS fractional 输出使用相同量化。

渲染只替换主 text fill 的 color choice：无 transparency 时继续输出 self-closing `srgbClr` / `schemeClr`；存在 transparency 时输出同一颜色节点和唯一 direct `a:alpha`。普通颜色、outline、glow、highlight、underline fill、font elements 与 run child 顺序保持现有语义。

## 严格读取

getter 只在 run 有唯一 direct `rPr/a:solidFill`、其中唯一 direct color choice 是合法 `srgbClr` 或受支持 `schemeClr`，并且该颜色恰好有一个合法 direct `alpha` 时返回 transparency。alpha 必须只有 `val` 属性、没有 child，且 val 是 `0..100000` 的完整十进制整数。返回值是 `100 - val / 1000`。

没有 alpha 返回 `undefined`；显式 `alpha val="100000"` 返回 0。重复 fill/color/alpha、缺失或额外属性、decimal/scientific/empty/out-of-range alpha、`alphaMod`/`alphaOff`/tint 等其他或混合 transform 都不伪造公共 transparency。已有 base color 仍按独立 color getter 暴露；读取失败不修改或删除原 XML。

## 编辑与 Lossless 边界

只读 getter、plain `.text`、shape transform 和其他非 rich-text mutation 不触碰原 alpha。plain `.text` 继续保留第一个 paragraph/run template，因此 alpha 随第一 run 样式复制到新增 paragraph。

`shape.richText` setter 整体替换 run：提供 transparency 时写量化后的唯一 alpha，省略时新 run 没有 direct alpha。setter 不逐项保留旧 run 的未知 color transform；paragraph `pPr`、`endParaRPr`、`bodyPr`、`lstStyle` 和 paragraph-level 未知 XML 继续保持现有 lossless 语义。

## 验证与错误处理

transparency 必须是 finite number 且位于 `0..100`。string、boolean、`null`、object、array、symbol、NaN、Infinity、负数或超过 100 都明确失败。合法 fractional value 在 validation 阶段量化到 alpha 整数。

所有验证在添加 shape 或写 part 之前完成。单个 run 失败不得留下部分 paragraph、part bytes、mutation journal、live shape 内容或 model identity 变化；外层 transaction rollback 同样恢复 exact bytes。

## 文档与兼容状态

PptxGenJS 兼容矩阵把现有“table-cell fit/textDirection 与 run transparency”组合行拆开：rich run transparency 标记为已支持，table-cell fit/textDirection 继续部分支持并留作后续小项。API 文档和聚合包 README 说明百分比方向、量化、explicit zero/absence、默认颜色、严格读取以及它与 glow opacity 和其他 fill 的边界。changelog 只宣称 rich text 主填充 transparency。

## 测试与验收

1. `addRichText()` 覆盖 omitted、0、25、fractional、100、sRGB、scheme、默认 color、空 text，以及与其他 run/paragraph style 组合。
2. getter 严格读取 alpha 0/1/49445/75000/100000；missing、empty、decimal、scientific、negative、100001、重复 alpha、其他/mixed transform、重复 fill/color 和非法颜色不伪造 transparency，只读不产生 mutation。
3. `shape.richText` 可增加、替换、量化、设为 0/100 和清除 transparency；plain `.text` 保留并复制第一 run alpha，duplicate、write/reopen 与 snapshot isolation 正确。
4. 非法类型、non-finite 和越界值在 mutation 前以 TypeError/RangeError 失败；内外层 rollback 保持 bytes、journal 和 identity。
5. PptxGenJS 4.0.1 对照覆盖 omitted/0/25/fractional/100、sRGB/theme/default color 的真实输出导入与 reopen，不访问私有字段。
6. 全仓 typecheck/test、独立 performance、Node/browser/declarations/tarball smoke 与 CLI validation 全部通过。
7. native 与手工 alpha baseline package diff 为空；LibreOffice headless 无修复打开并导出 0/25/50/75/100 对照页，透明度逐级变化且没有裁切或溢出。
