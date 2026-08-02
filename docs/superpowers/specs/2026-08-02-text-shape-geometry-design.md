# Text Shape Preset Geometry Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { shape })` 对应的原生创建能力，并把既有 preset geometry reader/editor 明确应用到 text shape：

- `AddTextOptions.shape?: PresetShapeType` 选择 text shape 的 direct `p:spPr/a:prstGeom`；
- 省略或 runtime `undefined` 使用 canonical `rect`，保持当前 text-shape 默认输出；
- plain/rich text、`addPlaceholder()`、placeholder population、layout/master wrappers 和 declarative `defineSlideMaster()` 共用同一 contract；
- 创建结果立即通过既有 `ShapeModel.presetType` 严格读取，并可用同一 setter 转换到其他 preset geometry；
- 178 个合法 OOXML preset tokens 全部可创建、读取、编辑、duplicate、write 和 reopen；
- geometry 与 fill/line/arrows/shadow/hyperlink、text body、transform 和 placeholder identity 保持独立 ownership。

本小项只增加 preset geometry 选择，不增加 `rectRadius`、`isTextBox`、text-specific `adjustments` shortcut、custom geometry points 或 WordArt geometry。`rectRadius`、`isTextBox` 和 `breakLine` 组合语义继续作为后续独立小项。Custom geometry 仍通过现有 `ShapeModel.customGeometry` live editor转换，创建 custom shape 继续使用 `addCustomShape()`。

## 2. PptxGenJS 4.0.1 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextPropsOptions.shape?: SHAPE_NAME`、`ShapeType` 和 `addText()` / `write()` 真实输出：

- omitted 和 runtime `undefined` 都输出 `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`；
- 合法 token 逐字进入 `prst`，包括 ellipse、roundRect、line、flowchart、callout、action-button 和 star families；
- `shape` 与 plain/rich text 内容无关，文本仍位于同一个 `p:sp/p:txBody`；
- exported runtime `ShapeType` 含 179 个值：177 个合法且与 native 共同支持的 preset tokens、非法 `folderCorner`，以及不属于 typed `SHAPE_NAME` 的 `custGeom`；
- `folderCorner` 被直接写成无效 OOXML；正确的 `foldedCorner` 没有被 4.0.1 暴露；
- falsy runtime 值（例如 `false` 或空字符串）被当成 omitted，truthy unknown string 和 number 则被无校验地写入 `prst`；
- runtime `custGeom` 会进入另一条 points-based renderer，不是 `TextPropsOptions.shape` 的 typed preset contract；
- `shape: 'line'` 在没有 `line` option 时会触发 4.0.1 的 runtime exception；提供空 line option 后可输出有效 line geometry；
- `rectRadius` 若 truthy 会向 preset geometry 写 adjustment，zero 被忽略；该 shortcut 的范围和优先级留给下一小项；
- `isTextBox` 只决定 `p:cNvSpPr@txBox`，不决定 `prstGeom`。

Native 对 177 个共同合法 tokens 比较最终 geometry 语义。Native 额外支持正确的 `foldedCorner`，拒绝 `folderCorner`；它不会复制 falsy fallback、string/number coercion、unknown token、无 line option exception 或 malformed custom-geometry passthrough。

## 3. 方案比较

### 方案 A：在 text renderer 内维护独立 token set 和 geometry renderer

实现直接，但会复制 preset token 校验、XML escaping 和 canonical geometry 结构。普通 shape 与 text shape 随后可能在新增 token或修正 strict semantics 时漂移。

### 方案 B：创建后调用 `ShapeModel.presetType` 二次修改

可以零改动复用 live setter，但会先写 rect，再进行第二次 package mutation；relationship 或后续解析失败时的边界更复杂，mutation journal 也不再代表一次原子创建。

### 方案 C：抽取并复用 preset geometry primitive（采用）

从现有 preset-shape codec 提供 internal `normalizePresetShapeType()` 和 `renderPresetShapeGeometry()`。普通 preset shape 与 text shape 在首次 renderer 中使用同一 canonical primitive；`ShapeModel.presetType` 继续使用同一 strict reader/replacer。该方案只新增一个 public option，不建立第二套 geometry subsystem，并保持单 transaction 创建。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly shape?: PresetShapeType;
  // existing fields remain unchanged
}
```

`PresetShapeType` 继续由 frozen `PRESET_SHAPE_TYPES` 推导，不新增 text-only enum 或 PptxGenJS-style alias namespace。有效集合恰好是现有 178 个 canonical OOXML tokens：包含 `foldedCorner`，不包含 `folderCorner` 或 `custGeom`。

`AddPlaceholderOptions` 扩展 `AddTextOptions`，因此自动获得字段。`SlideLayoutModel` / `SlideMasterModel` 的 direct text methods复用相同接口；declarative `defineSlideMaster()` 的 closed text/placeholder option-key reader显式加入 `shape`。

创建返回的 `ShapeModel` 不新增 `shape` alias。统一 live surface仍是：

```ts
const text = slide.addText('Hello', { shape: 'ellipse' });
text.presetType;              // 'ellipse'
text.presetType = 'hexagon';
```

## 5. 归一化与错误语义

- absent 或 own data property `shape: undefined` 归一化为 `rect`；
- defined value必须是 `PRESET_SHAPE_TYPES` 中的 exact string，不 trim、不做大小写变换、不 coercion；
- `folderCorner`、`custGeom`、empty/unknown string、number、boolean、null、object 和 symbol全部拒绝；
- own accessor property不执行 getter并直接拒绝；继承的 `shape` 不作为创建输入；
- normalized token是 detached immutable primitive；
- plain/rich/placeholder/declarative owner的 shape value在 relationship、part bytes、shape cache、shape order或 mutation journal变化前完成归一化；
- invalid input使用 `TypeError`，malformed existing OOXML继续使用 `ModelParseError`。

本小项不重写整个历史 `AddTextOptions` top-level object contract；它只保证新 `shape` field的 own-data读取和 strict token校验。所有创建流程仍位于现有 outer package transaction 中，任何后续异常完整回滚。

## 6. OOXML、读取与编辑

Canonical shaped text output：

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Text 2"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>...</a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    ...
  </p:spPr>
  <p:txBody>...</p:txBody>
</p:sp>
```

`renderPresetShapeGeometry(type)` 只负责 escaped `prst` 和 direct empty `a:avLst`。Transform、fill、line/arrows、effects 和 text body仍由 text renderer按原有顺序写入。

本小项不改变当前 `p:cNvSpPr txBox="1"` 默认。PptxGenJS 默认省略 `txBox`，而 `isTextBox: true` 才写 1；该 non-visual semantic与 geometry正交，由 `isTextBox` 小项单独设计，避免 `shape` 暗中改变既有 text-box state。

既有 strict reader只接受 namespace-correct text/shape direct `p:spPr` 中唯一 direct `a:prstGeom`、唯一 unqualified canonical `prst`；missing、unknown、qualified lookalike、repeated或与 `custGeom` ambiguous的状态返回 `undefined`，不改 source bytes。

既有 setter语义保持不变：

- 赋当前 token是 exact bytes/journal no-op，保留现有 adjustments；
- 赋另一 token whole-replace唯一 geometry为 canonical preset geometry和empty adjustment list；
- 给 readable custom geometry赋 preset token会转回 preset geometry；
- malformed/ambiguous existing geometry在 mutation前拒绝；
- shape identity、kind、name、transform、styles、effects、hyperlinks、text body、placeholder identity和sibling bytes保持不变。

## 7. Creation surface 与 lifecycle

- `addText()` 和 `addRichText()` 把 normalized token直接传给一次性 text-shape renderer；
- `addPlaceholder()` 对 plain/rich prompt使用相同 token；
- placeholder population使用population call的 explicit/default geometry，同时继承owner identity/name/transform，不修改layout/master source；
- layout/master direct text和placeholder source在各自owner part写geometry；
- declarative master definitions在异步资源准备前detached clone option，然后走相同strict normalization；
- duplicate保留exact geometry，move不改shape bytes，write/reopen和六种format保持token；
- outer rollback恢复source bytes、relationships、parts、order、journal和live identity；
- `shape` 不拥有relationship或dependent part，不改变 hyperlink clone/cleanup、target deletion或rich-run link ownership。

## 8. 与其他字段的组合

任意 preset token必须能与以下现有字段同时使用而互不改写：

- transform/name/placeholder identity；
- direct none/solid fill；
- simple line、arrows和line geometry；
- outer/inner simple shadow；
- whole-shape和per-run hyperlinks；
- margins、vertical alignment、direction、fit、wrap；
- plain/rich paragraphs、bullets、spacing、tabs和language。

`shape: 'line'` 仍是带text body的preset line geometry；它不等于 `AddTextOptions.line`。前者选择几何，后者选择outline style。Native在省略line style时继续写canonical direct no-fill line，不复制PptxGenJS的exception。

## 9. 验证策略

### Codec 与 model

- helper对178个tokens、default rect、escaped canonical rendering和invalid values；
- plain/rich/empty/multiline text、all-token matrix、immediate `presetType`读取；
- same-value no-op、different-token replacement、adjustment reset、custom-to-preset转换；
- malformed text geometry的undefined read和zero-mutation edit rejection；
- 与fill/line/arrows/shadow/hyperlink/text body组合后的ownership isolation。

### SDK 与 lifecycle

- slide/layout/master direct methods、placeholder create/populate和declarative definitions；
- caller detachment、own accessor不执行、invalid input exact zero mutation；
- duplicate、move、rollback、write/reopen和`pptx/pptm/ppsx/ppsm/potx/potm`。

### PptxGenJS conformance

- public-only fixture覆盖omitted/undefined、representative families、line和177个共同合法tokens；
- 比较direct geometry token和empty adjustment list，不把`txBox`差异混入本小项；
- raw evidence锁定`folderCorner` defect、falsy fallback、unknown/coercion、line exception和runtime `custGeom`，并证明native strict divergences。

### Release gates

- focused model/SDK/adapter/root tests；
- full Vitest、performance、typecheck和build；
- actual tarball Node/types/browser/CLI smoke覆盖create/read/edit/reopen；
- PowerPoint 2010 validator、LibreOffice round-trip和代表性shaped-text visual render；
- README、API、compatibility、progress和changelog更新。

## 10. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 所有text creation surfaces接受178个canonical tokens并默认rect；
2. 创建结果可由统一strict `ShapeModel.presetType`立即读取和编辑；
3. invalid/malformed、ownership isolation、placeholder/master isolation和lifecycle有zero-mutation或round-trip测试证明；
4. PptxGenJS 4.0.1共同合法语义和intentional strict differences有真实public-output证据；
5. packed Node/types/browser/CLI、full suite、performance、build和compatibility validator全部通过；
6. 文档明确`shape`范围及后续`rectRadius`、`isTextBox`和`breakLine`工作。
