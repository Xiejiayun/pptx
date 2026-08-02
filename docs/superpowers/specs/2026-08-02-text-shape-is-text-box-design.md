# Text Shape `isTextBox` Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { isTextBox })` 对应的创建语义，并增加同一 direct OOXML state 的严格 existing-deck 编辑能力：

- `AddTextOptions.isTextBox?: boolean` 只控制 `p:cNvSpPr@txBox`；
- omitted、own data property `undefined` 和 `false` 都输出不带 `txBox` 的 canonical `p:cNvSpPr`；
- `true` 输出 canonical `txBox="1"`；
- plain/rich text、`addPlaceholder()`、placeholder population、layout/master wrappers 与 declarative `defineSlideMaster()` 共用一个 contract；
- layout placeholder materialization 保留 source placeholder 的 direct `txBox` semantic；
- 创建或打开后的 `ShapeModel.isTextBox` 可严格读取和切换该状态；
- `isTextBox` 与 preset/custom geometry、adjustments/`rectRadius`、text body、styles、transform、hyperlinks 和 placeholder identity 保持独立 ownership。

本小项不加入 `breakLine` 组合语义，不根据 geometry 自动推断或改写 `isTextBox`，也不把 `txBox` 扩展为 generic non-visual-property object。完整 PptxGenJS parity 仍需后续小项继续完成。

## 2. PptxGenJS 4.0.1 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextPropsOptions.isTextBox?: boolean`、`addText()`、`defineSlideMaster()` 与 `writeFile()` 真实输出。Renderer 的唯一分支是对 `slideItemObj.options.isTextBox` 做 truthiness 判断：

```ts
strSlideXml += '<p:cNvSpPr'
  + (slideItemObj.options?.isTextBox ? ' txBox="1"/>' : '/>');
```

公开输出探针覆盖普通文本、富文本、母版定义中的 text/placeholder、空 placeholder materialization 和 placeholder population，观察到：

- omitted、runtime `undefined`、`false` 和 numeric zero 都输出 `<p:cNvSpPr/>`；
- `true` 输出 `<p:cNvSpPr txBox="1"/>`；
- runtime truthy string 也会写 `txBox="1"`，但这不属于 typed public contract；
- `shape`、`rectRadius`、plain/rich text content 都不参与该判断；
- layout placeholder 的 omitted/false/true 状态在空 slide owner materialization 时分别保持；
- placeholder population 使用 population call 的 `isTextBox`，不会隐式继承 layout source 的值。

Native 对合法 boolean input 比较相同 final semantic。Native 不复制 truthy coercion、inherited-property读取或 accessor执行。

## 3. 方案比较

### 方案 A：只在创建 renderer 增加 PptxGenJS truthiness 分支

改动最小，但 runtime string/object 会被接受，existing deck 仍无法读取或编辑 `txBox`，placeholder materialization 也会继续固定写 true。它只覆盖单一 happy path，不能满足创建与编辑并重的长期目标。

### 方案 B：由 geometry 自动推断 text box 或 shape

可以维持当前 default true，但会把 `txBox` 与 `prstGeom` 错误耦合。PptxGenJS 和 OOXML 都允许相同 geometry 分别带或不带 `txBox`，现有 geometry 设计也已明确两者正交。

### 方案 C：严格 boolean creator + direct live editor（采用）

创建端用 descriptor-safe normalizer 产生明确 boolean，renderer 一次写入 canonical state；existing-deck 端由独立 internal codec 解析、替换同一个 attribute；placeholder materializer复制 source direct semantic。该方案覆盖合法 PptxGenJS final output，同时符合 native strict/lossless editing contract。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly isTextBox?: boolean;
  // existing fields remain unchanged
}

export class ShapeModel extends BaseShapeModel {
  get isTextBox(): boolean | undefined;
  set isTextBox(value: boolean);
}
```

典型调用：

```ts
const shapeText = slide.addText('Shape text');
shapeText.isTextBox; // false

const textBox = slide.addText('Text box', { isTextBox: true });
textBox.isTextBox; // true
textBox.isTextBox = false;
```

`AddPlaceholderOptions` 扩展 `AddTextOptions`，因此自动获得字段。`SlideLayoutModel` / `SlideMasterModel` direct methods 复用同一接口；declarative `defineSlideMaster()` 的 closed option-key reader 显式加入 `isTextBox`。

Getter 返回 `undefined` 只表示 existing OOXML 结构或 token 不可安全解释，不是第三个 writable state。Setter 只接受 boolean；false 以 attribute absence 表示，不公开 `undefined` setter。

## 5. 创建归一化与错误语义

- inherited `isTextBox` 按 absent 处理；
- absent 或 own data property `isTextBox: undefined` 归一化为 `false`；
- own accessor property 不执行 getter并直接拒绝；
- defined value 必须是 primitive boolean；string、number、null、object、symbol、boxed boolean 和 function全部拒绝；
- normalized value 是 detached primitive，在 relationship、part bytes、shape ID/order、live cache或 mutation journal变化前完成验证；
- invalid input使用 `TypeError`，outer transaction 维持 exact zero mutation；
- 本小项不重写整个历史 `AddTextOptions` object contract，只保证新字段使用 own-data读取。

Native 的 default 从历史固定 `txBox="1"` 改为 attribute absence。这是为匹配 PptxGenJS typed default 的有意创建行为修正；已经打开但未编辑的 source deck 保持原 bytes，不做迁移。

## 6. OOXML 读取与编辑

Canonical false：

```xml
<p:nvSpPr>
  <p:cNvPr id="2" name="Shape text"/>
  <p:cNvSpPr/>
  <p:nvPr/>
</p:nvSpPr>
```

Canonical true：

```xml
<p:nvSpPr>
  <p:cNvPr id="3" name="Text box"/>
  <p:cNvSpPr txBox="1"/>
  <p:nvPr/>
</p:nvSpPr>
```

Internal reader只接受当前 `p:sp` 中 namespace-correct、唯一 direct `p:nvSpPr` 和唯一 direct `p:cNvSpPr`：

- attribute absent -> `false`；
- `1` / `true` / `on` -> `true`；
- `0` / `false` / `off` -> `false`；
- invalid token、qualified lookalike、重复 attribute、重复或缺失 required owner结构 -> `undefined`。

Setter先解析安全结构：

- `true` 写 canonical unqualified `txBox="1"`；exact canonical true是 bytes/journal no-op；
- `false` 删除唯一 `txBox` attribute；canonical absence是 bytes/journal no-op；
- 单一 alias或malformed token可由setter canonicalize/recover；
- ambiguous attribute或owner结构在 mutation前抛 `ModelParseError`；
- shape id/name/kind、geometry、transform、text、styles、placeholder和siblings保持不变。

该 codec 只操作 `p:sp/p:nvSpPr/p:cNvSpPr`，不误读 descendant、foreign namespace或图片/graphic-frame的 similarly named nodes。

## 7. Creation surface 与 placeholder lifecycle

- `addText()` / `addRichText()` 将 normalized boolean 传给一次性 text renderer；
- `addPlaceholder()` 的 plain/rich prompt使用同一 boolean；
- placeholder population使用本次 call 的 explicit/default value，并仅从 owner取得name、identity与最终transform；
- layout/master direct text和placeholder source在各自owner part写 direct state；
- declarative text/placeholder definitions在任何异步资源准备前 clone并验证新字段；
- empty layout-placeholder materialization读取 source `txBox` semantic：source true写 1，source false/absence保持 absence；
- source placeholder的invalid/ambiguous `txBox` 阻止materialization，不静默改成 true或false；
- population替换slide owner时不修改layout/master source；
- duplicate保留exact attribute，move不改shape bytes，六格式write/reopen保持semantic，outer rollback恢复bytes/order/journal/live identity。

非 `p:sp` layout placeholder没有text-box attribute，materialized text owner使用false；本小项不为picture或graphic-frame伪造 `txBox`。

## 8. 与 geometry 和其他字段的组合

`isTextBox` 只决定 non-visual attribute。它不选择或重置：

- `shape` / `presetType` / custom geometry；
- `rectRadius` / direct adjustments；
- fill、line、arrows、shadow、outer/run hyperlinks；
- margins、valign、direction、fit、wrap；
- transform/name、placeholder identity、plain/rich paragraphs。

创建后修改 `presetType`、adjustments、transform或任何style不得改动 `isTextBox`；切换 `isTextBox` 也不得改动这些状态。`breakLine` 后续只能消费这一独立状态，不能反向改变其 contract。

## 9. 验证策略

### Model 与 public owners

- omitted/undefined/false/true exact XML和immediate getter；
- own accessor不执行、inherited忽略、invalid primitive/object零 mutation；
- true/false/alias/malformed/ambiguous existing OOXML read/edit；
- same-value no-op、canonicalization、stable model identity；
- plain/rich/empty/multiline、placeholder create/populate、slide/layout/master/declarative owners；
- empty placeholder materialization复制source state，population call覆盖owner state且source隔离；
- geometry/radius/style/text/transform ownership isolation；
- duplicate、move、rollback、六格式和write/reopen；
- root declarations与compile-time invalid cases。

### PptxGenJS conformance

- public omitted/undefined/false/true plain/rich output；
- master text、placeholder、empty materialization和population output；
- false/true controls只差 `txBox` attribute；
- geometry/radius/content不影响该state；
- 锁定upstream truthy runtime coercion，并证明native strict rejection是intentional correction。

### Release 与兼容性

- focused/full Vitest、performance、TypeScript checks和build；
- actual tarball Node/types/browser/CLI 与真实Chrome create/read/edit/reopen；
- PowerPoint 2010 validator、exact part read与mutation isolation；
- LibreOffice round-trip保留true/false direct semantic；
- representative plain/rich/placeholder视觉和overflow检查；
- README、API、compatibility、progress和changelog收尾。

## 10. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 所有text creation owners接受strict boolean并输出PptxGenJS-compatible direct state；
2. `ShapeModel.isTextBox` 可安全读取、canonical edit和拒绝ambiguous state；
3. placeholder materialization复制source state，population call/default与source owner隔离；
4. geometry、radius、styles、text、transform和lifecycle ownership均有永久测试；
5. PptxGenJS public output、六格式、packed Node/types/browser/CLI、validator、LibreOffice与visual QA全部通过；
6. 文档明确default行为修正，以及后续`breakLine`和完整PptxGenJS parity仍未完成。
