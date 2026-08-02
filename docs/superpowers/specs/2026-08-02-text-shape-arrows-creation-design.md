# Text Shape Arrows Creation Design

日期：2026-08-02

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { line: { beginArrowType, endArrowType } })` 对应的原生创建能力：

- `AddTextOptions.arrows` 接受现有 strict `ShapeArrows` whole-replacement value；
- begin/end 各支持 `none | arrow | diamond | oval | stealth | triangle` 六种 direct endpoint type；
- plain text、rich text、layout/master text、declarative master text 与 text placeholder creation/population 共用同一 contract；
- 创建结果立即通过现有 live `ShapeModel.arrows` 读取、替换或清除；
- endpoint 与 `ShapeLine` 保持独立 ownership，duplicate、rollback、write/reopen 和六格式生命周期不变；
- 复用已经完成的 shape-arrows normalizer、renderer 和 live editor，不建立 text 专用 arrow type。

本小项不加入 arrow width/length size、shadow、hyperlink、text geometry、gradient/pattern/picture/group line fill、custom dash、cap/compound/alignment/join、`rectRadius`、`isTextBox` 或 `breakLine` 组合语义。这些继续按 advanced-text 路线逐项完成。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextPropsOptions.line` / `ShapeLineProps` 类型，以及 `addText()` / `write()` 真实输出。审计结果如下：

- nested `beginArrowType` 写 `a:headEnd@type`，nested `endArrowType` 写 `a:tailEnd@type`；
- 六种公开 token 均原样输出，两端同时存在时 child order 为 head 后 tail；
- arrow-only text line 不合成 color、width 或 dash，而是输出只有 endpoint 的 `a:ln`；
- 与合法 solid line 组合时，width、solid fill、preset dash 位于 endpoints 之前；
- `{ type: 'none', beginArrowType: 'diamond' }` 输出只有 `headEnd` 的 empty-paint line，不写 `a:noFill`；
- explicit endpoint type `none` 写 direct `type="none"`，不会与 endpoint absence 合并；
- nested deprecated `lineHead` / `lineTail`、top-level deprecated `lineHead` / `lineTail` 和 empty-string endpoint 在 text output 中都被忽略；
- unknown runtime token 被原样写入非法 OOXML；
- PptxGenJS 4.0.1 不公开 arrow size，types 仍将 begin/end size 标记为 future。

原生 API 保持项目既有 strict、可逆 direct-state 语义，不复制 invalid passthrough、empty ignore 或 deprecated aliases。由于 text shape 的已发布默认 line 是 canonical direct no-fill，native arrow-only creation 输出 `a:noFill` 后跟 endpoints；这是与 PptxGenJS empty-paint line 的明确 direct-state 差异。Supported endpoint type/order 以及与显式 solid line 组合的 final semantics 对等。

## 3. 方案比较

### 方案 A：把 begin/end 字段加入 `ShapeLine`

输入形状最接近 PptxGenJS，但会破坏已经发布的独立 ownership：`shape.line = undefined` 当前只清除 width/fill/dash并保留 endpoints。把箭头重新耦合进 line whole replacement 会让 clear 和 same-value 语义变得含糊。

### 方案 B：创建 text 后调用 `ShapeModel.arrows` setter

实现很短，但一次 public creation 产生两次 part mutation，中间状态可观察，outer transaction journal 更嘈杂，也会让 text renderer 的最终 contract 与 create API 分离。

### 方案 C：复用 `ShapeArrows` 并在 text renderer 直接组合（采用）

`validateAddTextOptions()` 调用现有 `normalizeShapeArrows()`，把 detached normalized endpoints 与 normalized line 一起传入 `textShapeXml()`。Renderer 在同一个 `a:ln` 中按 line fill/dash、head、tail 的 schema order 输出。现有 creation transaction、placeholder owner replacement、stable identity 和 live editor 均保持不变。

该方案改动最小、边界清晰，也让 preset shape 与 text shape 共享同一 public endpoint value和 validation contract。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly arrows?: ShapeArrows;
  // existing fields remain unchanged
}
```

`ShapeArrows` 继续使用现有公开类型：

```ts
type ShapeArrowType =
  | 'none'
  | 'arrow'
  | 'diamond'
  | 'oval'
  | 'stealth'
  | 'triangle';

interface ShapeArrows {
  readonly begin?: ShapeArrowType;
  readonly end?: ShapeArrowType;
}
```

`AddPlaceholderOptions` 基于 `AddTextOptions`，因此 text placeholder creation/population 自动获得同一 arrows 字段。`SlideMasterObject` 的 text/placeholder options 同样复用该字段；`defineSlideMaster()` 的 closed option-key reader 必须显式允许 `arrows`。

Getter 和 setter 不新增 API：returned `ShapeModel.arrows` 继续提供 detached frozen snapshot、whole replacement 和 clear。Missing side 表示清除该 side，explicit `'none'` 表示保留 direct endpoint，empty snapshot 与 `undefined` 表示清除两端。

## 5. 归一化与错误语义

- omitted 或 runtime `undefined` 归一化为无 endpoints；
- arrows 必须是 ordinary 或 null-prototype own-data object，只允许 `begin` / `end`；
- 每个 defined endpoint 必须是六个 canonical token 之一；
- missing 与 own runtime `undefined` side 归一化为 absence；
- empty normalized snapshot 合法，creation 时不输出 endpoint；
- symbols、accessors、arrays、class instances、inherited-only fields、unknown keys、aliases、empty string 和 invalid token 在任何 package mutation 前拒绝；
- normalized value 与 caller 脱离，异步 declarative preparation 期间的 caller mutation 不影响最终 XML；
- invalid plain/rich/placeholder/master definition 保持 parts、relationships、XML、runtime cache、shape order和 mutation journal 不变。

错误继续沿用 shared codec 的 `TypeError`。不接受 PptxGenJS 的 `beginArrowType`、`endArrowType`、`lineHead`、`lineTail` 或其他 nested/top-level aliases。

## 6. OOXML 与渲染顺序

Default-line arrow-only output：

```xml
<a:ln>
  <a:noFill/>
  <a:headEnd type="triangle"/>
  <a:tailEnd type="arrow"/>
</a:ln>
```

Solid-line combined output：

```xml
<a:ln w="31750">
  <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
  <a:prstDash val="dashDot"/>
  <a:headEnd type="stealth"/>
  <a:tailEnd type="oval"/>
</a:ln>
```

Text line 继续位于 geometry 与 shape fill 之后、`p:spPr` 结束之前。Endpoint 只追加到同一个 line container：line fill/dash 在前，head 在 tail 前。除目标 endpoint children 外，transform、geometry、shape fill、line width/fill/dash、name、placeholder identity、body properties、paragraphs、runs、relationships 与相邻 shapes 都不改变。

Placeholder population 仍以 owner geometry/name/identity 为准，arrows 使用本次调用的 normalized direct value；它不会修改 layout placeholder 或合成 theme-effective endpoints。

## 7. Line/endpoint ownership 与生命周期

- omitted text line 仍物化既有 canonical direct no-fill；adding arrows 不改变这个 default；
- `line` 与 `arrows` 可同时创建，并在同一 `a:ln` 中输出；
- `shape.line = undefined` 只清除 direct width/fill/dash，保留 head/tail；
- `shape.arrows = undefined` 只清除 endpoints，保留 line width/fill/dash；
- same-value arrows assignment 是 exact bytes/journal no-op；
- explicit endpoint `'none'` 与 endpoint absence 保持可区分；
- `addText()`、`addRichText()`、`addPlaceholder()`、layout/master wrappers 与 declarative master definitions 都走同一 renderer；
- duplicate、move、outer rollback、write/reopen 与六种 presentation format 保留 supported direct state；
- placeholder owner materialization/population 只改变当前 owner，layout/master source 保持隔离；
- arrows 没有 relationship ownership，不引入 relationship 或 GC 分支。

## 8. 验证策略

### Model/SDK

- omitted/runtime-undefined/empty、begin-only、end-only、both、explicit none 与全部六种 token；
- default no-fill、none line、solid line 和 width/dash/transparency 组合的 exact child order；
- immediate detached `ShapeModel.arrows` snapshot、caller detachment、same-value no-op、whole replacement 与 clear；
- line/arrows 双向 ownership isolation；
- plain/rich、placeholder、layout/master direct creation 与 declarative definition；
- invalid object/accessor/symbol/alias/token 的 zero mutation；
- duplicate、rollback、write/reopen、六格式、stable identity 与 unrelated XML isolation。

### PptxGenJS

- public-only fixtures 覆盖 begin/end 的六种 token、both、arrow-only、none-line arrow、solid-line combination、explicit endpoint none、empty string、nested/top-level deprecated aliases和 invalid passthrough；
- imported `ShapeModel.arrows`、`ShapeModel.line` 与 raw direct line state 对照；
- supported endpoint final semantics 对等；native default no-fill、strict aliases和 invalid rejection 的 intentional differences 单独锁定。

### Release gates

- focused model/SDK/adapter/root tests；
- full Vitest、performance、typecheck、build；
- actual tarball Node/types/browser/CLI smoke 覆盖 `addText({ arrows })`；
- real Chrome immediate/detached/reopen state、PowerPoint 2010 profile validation 与文档/compatibility matrix 更新。

## 9. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 所有 text creation surfaces 接受并写入 strict begin/end arrows；
2. validation-before-mutation、detachment、line ownership、placeholder isolation、identity 与 lifecycle 有测试证明；
3. PptxGenJS public-output supported cases和 intentional differences 有锁定证据；
4. packed Node/types/browser/CLI、full suite、performance、typecheck、build 与 validator 全通过；
5. README、API、compatibility、progress 与 changelog 明确完成范围及下一项。
