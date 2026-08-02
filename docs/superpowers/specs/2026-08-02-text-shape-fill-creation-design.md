# Text Shape Fill Creation Design

日期：2026-08-02

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { fill })` 对应的原生创建能力：

- `AddTextOptions.fill` 接受 strict direct no-fill 或 solid fill；
- plain text、rich text、layout/master text、declarative master text 与 text placeholder population 共用同一 contract；
- 创建结果立即通过现有 live `ShapeModel.fill` 读取、编辑、清除，并保持 duplicate、write、reopen 和六格式生命周期；
- 复用已经完成的 shape simple-fill value codec，不引入第二套颜色、透明度或 descriptor validation。

本小项不加入 text shape line、arrows、shadow、hyperlink、preset geometry、`rectRadius`、`isTextBox`、run hyperlink 或 `breakLine` 组合语义；这些继续按 advanced-text 路线逐项完成。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextPropsOptions.fill` 与 `addText()` / `write()` 输出：

- omitted fill 写 direct `a:noFill`；
- `{ color: 'AB12CD', transparency: 25 }` 写 `a:solidFill/a:srgbClr@val="AB12CD"/a:alpha@val="75000"`；
- scheme color 写 `a:schemeClr`；
- explicit zero transparency 省略 alpha；
- `{ type: 'none' }` 省略 direct fill choice，而不是写 `a:noFill`。

原生 API 保持项目既有 strict direct-state 语义：omitted/`undefined`/`{ kind: 'none' }` 都创建 canonical direct `a:noFill`；explicit transparency zero 写 `a:alpha val="100000"`，从而保留显式 zero；solid 必须提供合法 sRGB 或 scheme color。合法 solid 与非零透明度达到相同最终 OOXML/effective rendering，native 不复制 PptxGenJS 的 falsy collapse 或 permissive fallback。

## 3. 方案比较

### 方案 A：创建后调用 `ShapeModel.fill` setter

先按现状创建 no-fill text shape，再调用 setter 改 fill。实现很短，但一次 public creation 产生两次 part mutation，中间状态可观察，outer transaction journal 更嘈杂，也让 renderer 与最终 contract 分离。

### 方案 B：复制 text 专用 fill normalizer/renderer

在 `slide.ts` 内实现 text-only fill。局部看似独立，但会复制 sRGB/scheme、transparency、descriptor-safe input 与 equality 规则，之后容易和 preset shape/table fill 漂移。

### 方案 C：复用 simple-fill codec 并直接渲染（采用）

`validateAddTextOptions()` 调用现有 `normalizeSimpleFill()`，把 detached normalized fill 传入 `textShapeXml()`，用 `renderSimpleFill()` 替换硬编码 `<a:noFill/>`。现有 text creation transaction、placeholder owner replacement、stable model identity 与 renderer ordering 均不改变；这是最小且完整的 vertical slice。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly fill?: ShapeFill;
  // existing fields remain unchanged
}
```

`ShapeFill` 继续使用现有公开 union：

```ts
type ShapeFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };
```

因为 `AddPlaceholderOptions` 基于 `AddTextOptions`，text placeholder creation/population 自动获得同一 fill 类型。`SlideMasterObject` 的 text/placeholder options 也复用该字段；`defineSlideMaster()` 的 strict option-key reader必须显式允许 `fill`。

## 5. 归一化与错误语义

- omitted 或 runtime `undefined` 归一化为 `{ kind: 'none' }`，保持当前 text default bytes；
- explicit none 同样写 direct no-fill；
- solid color 只接受 strict sRGB/supported scheme，transparency 为 finite `0..100` 并量化到 `0.001%`；
- nested object 必须是 ordinary/null-prototype own-data object，unknown keys、symbols、accessors、class instances、missing color 与 invalid range 在任何 package mutation 前拒绝；
- normalized fill 与 caller 深度脱离，caller 后续 mutation 不影响 XML；
- invalid plain/rich/placeholder/master definition 必须保持 parts、relationships、XML、runtime cache 和 mutation journal 不变。

错误继续沿用 shared codec 的 `TypeError` / `RangeError`。不新增 text-specific fallback。

## 6. OOXML 与渲染顺序

Supported output：

```xml
<p:spPr>
  <a:xfrm>...</a:xfrm>
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:noFill/>
  <a:ln><a:noFill/></a:ln>
</p:spPr>
```

或：

```xml
<a:solidFill>
  <a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr>
</a:solidFill>
```

Fill 必须继续位于 geometry 之后、line 之前。除目标 fill child 外，transform、name、placeholder identity、line、body properties、paragraphs、runs、relationships 和相邻 shapes 均保持既有 canonical output。

Placeholder population 仍以 owner geometry/name/identity 为准，fill 使用本次调用的 normalized direct value；它不会修改 layout placeholder 或合成 theme-effective fill。

## 7. 生命周期与兼容范围

- `addText()`、`addRichText()`、`addPlaceholder()`、layout/master wrappers 与 declarative master definitions 都走同一 renderer；
- returned `ShapeModel.fill` 立即反映 normalized direct state；existing setter/clear/no-op contract 不变；
- duplicate、move、write/reopen 与六种 presentation format 保留 fill；
- placeholder owner materialization/population 只改变当前 owner；layout/master inherited source 保持隔离；
- fill 没有 relationship ownership，不引入 GC 分支；
- 旧调用未提供 fill 时输出继续为 direct no-fill，不发生默认行为变化。

## 8. 验证策略

### Model/SDK

- plain/rich omitted、none、sRGB、scheme、transparency 0/25/100 与 immediate snapshot；
- placeholder、layout/master direct creation 和 declarative definition；
- caller detachment、invalid object/accessor/symbol/range zero mutation；
- duplicate、rollback、write/reopen、六格式、stable identity 与 unrelated XML isolation。

### PptxGenJS

- public-only fixtures 覆盖 omitted、none、sRGB、scheme、25% 与 explicit zero；
- imported `ShapeModel.fill` 与 raw direct fill state 对照；
- native supported solid output与 PptxGenJS final semantics 对等；intentional none/zero direct-state differences 单独锁定。

### Release gates

- focused model/SDK/adapter/root tests；
- full Vitest、performance、typecheck、build；
- actual tarball Node/types/browser smoke 覆盖 `addText({ fill })`；
- PowerPoint profile validation 和文档/compatibility matrix 更新。

## 9. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 所有 text creation surfaces 接受并写入 strict none/solid fill；
2. validation-before-mutation、detachment、placeholder isolation、identity 与 lifecycle 有测试证明；
3. PptxGenJS public-output supported cases 和 intentional differences 有锁定证据；
4. packed Node/types/browser、full suite、performance、typecheck、build 与 validator 全通过；
5. README、API、compatibility、progress 与 changelog 明确完成范围及下一项。
