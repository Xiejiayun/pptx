# Text Shape Rectangle Radius Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { shape, rectRadius })` 对应的原生创建能力，同时复用已有 preset-adjustment reader/editor：

- `AddTextOptions.rectRadius?: Emu` 为 `shape: 'roundRect'` 选择物理圆角半径；
- 省略或 own data property `undefined` 保持 canonical empty `a:avLst`；
- explicit zero 写 direct `{ name: 'adj', value: 0 }`，不复制 PptxGenJS 的 truthiness 丢失；
- positive radius 根据最终 rounded width/height 计算一个 direct `adj` guide；
- plain/rich text、`addPlaceholder()`、placeholder population、layout/master wrappers 与 declarative `defineSlideMaster()` 共用一个 contract；
- 创建结果立即通过已有 `ShapeModel.adjustments` 读取、whole-replace 和 clear；
- radius 与 fill/line/arrows/shadow/hyperlink、text body、transform 和 placeholder identity 保持独立 ownership。

本小项不增加 `ShapeModel.rectRadius`、`AddTextOptions.adjustments`、`angleRange`、`arcThicknessRatio`、custom geometry 或自动重算机制。`isTextBox` 与 `breakLine` 继续作为后续独立小项。普通 preset shape 的通用 direct adjustments 仍使用既有 `AddShapeOptions.adjustments` / `ShapeModel.adjustments`。

## 2. PptxGenJS 4.0.1 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextPropsOptions.rectRadius?: number`、`addText()` 与 `write()` 真实输出。类型注释将它限定为 rounded rectangle 的 `0.0..1.0` 值；runtime renderer 使用：

```ts
Math.round((rectRadius * 914400 * 100000) / Math.min(cx, cy))
```

其中 `rectRadius` 按 inch 输入，`cx` / `cy` 是最终 EMU extent。已观察行为：

- `shape: 'roundRect'`、w=2in、h=1in、radius=0.5in 写 `adj=50000`；
- w=4in、h=2in、radius=0.5in 写 `adj=25000`；
- omitted 与 numeric zero 都写 empty `a:avLst`；
- positive value 写唯一 `a:gd name="adj" fmla="val N"`；
- runtime 只检查 truthiness，不验证 shape kind、range、finite number 或 input type；
- string 可参与乘法 coercion，negative/out-of-range/non-finite value 可进入非法或不合理的 formula；
- `rectRadius` truthy 时优先于 shape-level `angleRange`，但 text contract 本身没有 `angleRange`；
- renderer 只在创建时计算 guide，后续没有 existing-deck live radius model。

Native 对合法 positive `roundRect` input 比较相同 final guide integer。Native 不复制 falsy loss、string coercion、wrong-shape passthrough、invalid range 或 malformed formula。

## 3. 方案比较

### 方案 A：逐字复制 PptxGenJS number/inch 与 permissive runtime

API 表面最接近 PptxGenJS，但会让 `rectRadius` 成为 native geometry 中唯一隐式 inch number，并复制 zero 丢失、wrong-shape guide、string coercion 和 non-finite formula。与现有 `Transform` / `inches()`、strict validation 和 direct-state preservation 冲突。

### 方案 B：给 `AddTextOptions` 增加通用 `adjustments`

可直接复用 `AddShapeOptions.adjustments`，但没有提供 PptxGenJS 的 `rectRadius` 字段，也把本小项扩大到所有 text preset 的任意 adjustment creation。Existing-deck live `ShapeModel.adjustments` 已覆盖通用编辑，不需要第二个 creation surface。

### 方案 C：严格的 EMU radius shortcut（采用）

公开 branded `Emu` radius，调用方通过 `inches()` 显式换算。创建时在现有 transaction 内把 radius 和最终 extent 转成唯一 direct `adj` guide，之后继续使用统一 `ShapeModel.adjustments`。该方案覆盖 PptxGenJS 合法最终语义，同时保留 native unit、explicit-zero intent 和 strict failure contract。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly shape?: PresetShapeType;
  readonly rectRadius?: Emu;
  // existing fields remain unchanged
}
```

`Emu` 与 `inches()` 已由 root package 公开。典型调用：

```ts
const text = slide.addText('Rounded', {
  shape: 'roundRect',
  rectRadius: inches(0.5),
  width: inches(4),
  height: inches(2),
});

text.adjustments; // [{ name: 'adj', value: 25000 }]
```

`AddPlaceholderOptions` 扩展 `AddTextOptions`，因此自动获得字段。`SlideLayoutModel` / `SlideMasterModel` direct text methods 复用同一接口；declarative `defineSlideMaster()` 的 closed text/placeholder option-key reader 显式加入 `rectRadius`。

不新增 `ShapeModel.rectRadius` alias。Direct live surface 保持：

```ts
text.adjustments = [{ name: 'adj', value: 12500 }];
text.adjustments = [];
```

## 5. 归一化与错误语义

- inherited `rectRadius` 按 absent 处理；
- absent 或 own data property `rectRadius: undefined` 归一化为 `undefined`；
- own accessor property 不执行 getter并直接拒绝；
- defined value 必须是 finite number，rounded EMU 必须是 safe integer；
- allowed range 是 inclusive `0..914400` EMU，即 PptxGenJS documented `0..1` inch；
- `-0` 规范化为 `0`；
- defined radius 要求 normalized `shape === 'roundRect'`，omitted/default `rect`、其他 preset 与 custom token 均拒绝；
- string、boolean、null、object、symbol、NaN、Infinity、negative、over-one-inch 与 unsafe value 全部拒绝；
- radius 在 relationship、part bytes、shape ID/order、live cache 或 mutation journal变化前完成验证和 caller detachment；
- invalid input 使用 `TypeError` 或 `RangeError`，existing malformed adjustment state继续通过已有 reader 返回 `undefined`。

本小项不重写整个历史 `AddTextOptions` object contract；只保证新字段使用 descriptor-safe own-data读取。Declarative master definition 仍由其 closed-key reader先完成 ordinary-object clone。

## 6. Guide 计算与 OOXML

最终 transform 使用 text renderer 已有量化：

```ts
const width = Math.round(options.width ?? inches(1));
const height = Math.round(options.height ?? inches(1));
```

若 radius 省略，adjustments 是 frozen empty list。若 radius存在：

```ts
const value = Math.round(rectRadius * 100000 / Math.min(width, height));
const adjustments = [{ name: 'adj', value }];
```

Radius 已限制在一 inch，final extent 是 positive rounded EMU，因此结果为 safe integer。计算使用 placeholder population 的最终 owner extent，而不是被 owner覆盖的 caller width/height。

Canonical output：

```xml
<a:prstGeom prst="roundRect">
  <a:avLst>
    <a:gd name="adj" fmla="val 25000"/>
  </a:avLst>
</a:prstGeom>
```

Renderer 调用已有 `renderPresetShapeGeometry(type, 'a:', adjustments)`，不建立第二套 XML codec。Transform、fill、line/arrows、effects 与 text body 顺序不变。

Explicit zero 产生 `fmla="val 0"`；omitted/undefined 产生 self-closing empty `a:avLst`。这是可读取、可编辑且可区分的 direct intent。

## 7. Creation surface 与 lifecycle

- `addText()` / `addRichText()` 把 normalized radius传给一次性 text renderer；
- `addPlaceholder()` 的 plain/rich prompt使用同一 radius contract；
- placeholder population使用population call的 explicit `shape` / `rectRadius` 与 owner最终 transform，不修改layout/master source；
- layout/master direct text和placeholder source在各自owner part写guide；
- declarative master definition在任何异步资源准备前 clone并验证新字段；
- duplicate保留exact guide，move不改shape bytes，write/reopen和六格式保持value；
- outer rollback恢复source bytes、relationships、parts、order、journal和live identity；
- radius不拥有relationship或dependent part。

## 8. Live 编辑与组合语义

创建结果由已有 `ShapeModel.adjustments` 暴露 frozen direct snapshot。Same list assignment 是 exact bytes/journal no-op；whole replacement和`[]` clear只修改direct `a:avLst`。

创建后修改 transform 不自动重算 `adj`。Guide 是创建时物化的 direct OOXML state；调用方如需保持同一物理 radius，可根据新 transform显式替换 `adjustments`。这避免普通 resize 暗中修改 geometry bytes，也与现有 lossless editor一致。

赋相同 `presetType = 'roundRect'` 保留guide；赋另一preset清空guide并只替换geometry。Custom geometry conversion与preset conversion沿用现有规则。

Radius 可与 transform/name/placeholder、fill、simple line、arrows、simple shadow、outer/run hyperlinks、margins、valign、direction、fit、wrap和plain/rich paragraphs组合，且任一 unrelated live edit都不得改变guide。

## 9. 验证策略

### Model 与 public owners

- omitted/undefined、zero、boundary one-inch、fractional-EMU rounding，以及2×1/4×2/非方形 extent的exact guide计算；
- own accessor不执行、inherited忽略、invalid type/range/wrong shape零 mutation；
- plain/rich/placeholder create/populate、slide/layout/master/declarative owners；
- immediate `adjustments`、same-value no-op、whole replace/clear、preset reset；
- fill/line/arrows/shadow/hyperlink/text/transform ownership isolation；
- duplicate、move、rollback、六格式、write/reopen和placeholder-source isolation；
- root `Emu` / `AddTextOptions` declarations与invalid compile-time cases。

### PptxGenJS conformance

- public-only 2×1 radius 0.5、4×2 radius 0.5、radius 1.0与representative fractional radius；
- 比较合法positive final `{ name: 'adj', value }`；
- 锁定upstream zero drop、string coercion、negative/out-of-range/non-finite、wrong-shape passthrough；
- 证明native explicit zero、strict range与wrong-shape rejection是intentional correction；
- 证明`rectRadius`只改变adjustment list，不改变geometry token、text或styles。

### Release 与兼容性

- focused/full Vitest、performance、两种 TypeScript check和两套build；
- actual tarball Node/types/browser/CLI 与真实 Chrome create/read/edit/reopen；
- PowerPoint 2010 validator、exact part read与mutation isolation；
- LibreOffice round-trip保留text、roundRect和direct adjustment；
- representative radii visual render、overflow与逐页检查；
- README、API、compatibility、progress和changelog收尾。

## 10. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 全部text creation owners接受strict EMU `rectRadius`并按最终extent输出exact direct guide；
2. omitted/undefined/explicit-zero、range、wrong-shape与descriptor安全语义有永久测试；
3. 创建结果可由统一`ShapeModel.adjustments`立即读取和编辑，且style/transform/placeholder ownership隔离；
4. PptxGenJS 4.0.1合法positive final语义与intentional strict differences有public-output证据；
5. duplicate/rollback/六格式、packed Node/types/browser/CLI、validator、LibreOffice与visual QA全部通过；
6. 文档明确resize不自动重算，以及后续`isTextBox`、`breakLine`和完整PptxGenJS parity仍未完成。
