# Text Shape Simple Line Creation Design

日期：2026-08-02

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { line })` 对应的原生创建能力：

- `AddTextOptions.line` 接受 strict direct no-line 或 solid line；
- solid line 支持 sRGB/theme color、0–100% transparency、0–1584pt width 与 8 种 preset dash；
- plain text、rich text、layout/master text、declarative master text 与 text placeholder population 共用同一 contract；
- 创建结果立即通过现有 live `ShapeModel.line` 读取、编辑、清除，并保持 duplicate、write、reopen 和六格式生命周期；
- 复用已经完成的 shape simple-line value codec，不引入 text 专用颜色、透明度、宽度或 dash validation。

本小项不加入 text line arrows、shadow、hyperlink、gradient/pattern line fill、custom dash、line join、cap/compound/alignment、preset geometry、`rectRadius`、`isTextBox` 或 `breakLine`。这些继续按 advanced-text 路线逐项完成。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextPropsOptions.line` 与 `addText()` / `write()` 真实输出。审计结果如下：

- omitted、`{ type: 'none' }`、empty line 与 `{ type: 'solid' }` without color 都输出 empty `<a:ln></a:ln>`；
- `{ color: 'AB12CD' }` 输出 sRGB solid line，但省略 `a:ln@w` 与 preset dash；
- theme color 输出 `a:schemeClr`；
- `transparency: 25/100` 分别输出 `a:alpha@val="75000"/"0"`，explicit zero 省略 alpha；
- `width: 0` 通过 falsy fallback 省略 `w`，正 width 按 point 写入 `a:ln@w`；
- `dashType` 的 `solid`、`dash`、`dashDot`、`lgDash`、`lgDashDot`、`lgDashDotDot`、`sysDash`、`sysDot` 原样写入 `a:prstDash@val`；
- nested deprecated `alpha` 仍作为 transparency 生效，nested `lineDash` 被忽略；
- shape fill 位于 geometry 之后，line 紧随 shape fill。

原生 API 采用项目既有 strict、可逆 direct-state 语义，不复制 permissive/falsy fallback：empty/missing-color/unknown/deprecated fields 在 mutation 前拒绝；explicit none 写 direct `a:noFill`；explicit transparency zero 写 `alpha="100000"`；width zero 写 direct `w="0"`；omitted width/dash materialize 为 1pt/solid。Supported color/theme/nonzero transparency/positive width/8 dash 的 final semantics 与 PptxGenJS 对等，上述 direct bytes 差异有独立测试和文档说明。

## 3. 方案比较

### 方案 A：创建后调用 `ShapeModel.line` setter

先创建当前 no-line text shape，再调用 setter。实现很短，但一次 public creation 会产生两次 part mutation，中间状态可观察，outer transaction journal 更嘈杂，也让 renderer 与最终 contract 分离。

### 方案 B：复制 text 专用 line codec

在 text renderer 内重新处理 color/transparency/width/dash。这样会复制已经验证的 descriptor-safe input、EMU 量化与 strict color 规则，之后容易和 preset shape line 漂移。

### 方案 C：复用 simple-line codec并直接渲染（采用）

`validateAddTextOptions()` 调用现有 `normalizeSimpleLine()`，把 detached normalized line 传入 `textShapeXml()`。新增一个只负责 `a:ln` 容器的 text renderer：none 输出 direct no-fill；solid 输出 normalized width、simple solid fill 与 preset dash。现有 text creation transaction、placeholder owner replacement、stable identity 和 shape-property ordering 均不改变。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  // existing fields remain unchanged
}
```

`ShapeLine` 继续使用现有公开 union：

```ts
type ShapeLine =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly transparency?: number;
      readonly width?: number;
      readonly dash?: ShapeLineDash;
    };
```

因为 `AddPlaceholderOptions` 基于 `AddTextOptions`，text placeholder creation/population 自动获得同一 line 类型。`SlideMasterObject` 的 text/placeholder options 也复用该字段；`defineSlideMaster()` 的 strict option-key reader必须显式允许 `line`。

## 5. 归一化与错误语义

- omitted 或 runtime `undefined` 使用 `{ kind: 'none' }` 作为 renderer 默认值，保持当前 text default bytes；
- explicit none 同样写 direct no-fill；
- solid line 必须提供 strict sRGB/supported scheme color；transparency 为 finite `0..100` 并量化到 `0.001%`；
- width 为 finite `0..1584` point 并量化到 1 EMU，omitted materialize 为 1 point；
- dash 必须是 8 个 canonical token之一，omitted materialize 为 `solid`；
- nested object 必须是 ordinary/null-prototype own-data object，unknown keys、symbols、accessors、class instances、missing color 与 invalid range 在任何 package mutation 前拒绝；
- normalized line 和 nested color 与 caller 深度脱离，caller 后续 mutation 不影响 XML；
- invalid plain/rich/placeholder/master definition 保持 parts、relationships、XML、runtime cache 和 mutation journal 不变。

错误继续沿用 shared codec 的 `TypeError` / `RangeError`。不接受 PptxGenJS 的 `type`、`alpha`、`dashType`、`lineDash`、arrow aliases 或顶层 legacy line fields。

## 6. OOXML 与渲染顺序

Default/none output：

```xml
<p:spPr>
  <a:xfrm>...</a:xfrm>
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:noFill/>
  <a:ln><a:noFill/></a:ln>
</p:spPr>
```

Solid output：

```xml
<a:ln w="31750">
  <a:solidFill>
    <a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr>
  </a:solidFill>
  <a:prstDash val="dashDot"/>
</a:ln>
```

Line 必须继续位于 fill 之后、`p:spPr` 结束之前。除目标 line child 外，transform、geometry、fill、name、placeholder identity、body properties、paragraphs、runs、relationships 与相邻 shapes 均保持既有 canonical output。

Placeholder population 仍以 owner geometry/name/identity 为准，line 使用本次调用的 normalized direct value；它不会修改 layout placeholder 或合成 theme-effective line。

## 7. 生命周期与兼容范围

- `addText()`、`addRichText()`、`addPlaceholder()`、layout/master wrappers 与 declarative master definitions 都走同一 renderer；
- returned `ShapeModel.line` 立即反映 normalized direct state；existing setter/clear/no-op contract 不变；
- duplicate、move、write/reopen 与六种 presentation format 保留 line；
- placeholder owner materialization/population 只改变当前 owner，layout/master inherited source 保持隔离；
- simple line 没有 relationship ownership，不引入 GC 分支；
- 旧调用未提供 line 时输出继续为 direct no-fill，不发生默认行为变化。

## 8. 验证策略

### Model/SDK

- plain/rich omitted、none、sRGB、scheme、transparency 0/25/100、width 0/default/positive 与 8 dash；
- immediate detached snapshot、caller detachment、same-value setter no-op 与 clear；
- placeholder、layout/master direct creation和 declarative definition；
- invalid object/accessor/symbol/range zero mutation；
- duplicate、rollback、write/reopen、六格式、stable identity 与 unrelated XML isolation。

### PptxGenJS

- public-only fixtures 覆盖 omitted、none、empty/missing-color、sRGB、scheme、transparency 0/25/100、width 0/positive、8 dash 与 deprecated aliases；
- imported `ShapeModel.line` 与 raw direct line state 对照；
- native supported solid output与 PptxGenJS final semantics 对等；omitted/none/default width/default dash/zero values的 intentional direct-state differences 单独锁定。

### Release gates

- focused model/SDK/adapter/root tests；
- full Vitest、performance、typecheck、build；
- actual tarball Node/types/browser smoke 覆盖 `addText({ line })`；
- PowerPoint profile validation、LibreOffice render 与文档/compatibility matrix 更新。

## 9. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 所有 text creation surfaces 接受并写入 strict none/solid line；
2. validation-before-mutation、detachment、placeholder isolation、identity 与 lifecycle 有测试证明；
3. PptxGenJS public-output supported cases 和 intentional differences 有锁定证据；
4. packed Node/types/browser、full suite、performance、typecheck、build 与 validator 全通过；
5. README、API、compatibility、progress 与 changelog 明确完成范围及下一项。
