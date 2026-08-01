# Slide Default Color Design

## 1. 目标

完成一个可独立验收的 default color 小项，使 native API 与 PptxGenJS 4.0.1 的 `slide.color`
公开行为对等，同时保持现有 OOXML 模型的严格性：

- 每张 slide 可设置、读取和清除一个 transient default text color；
- default 支持 sRGB 与 theme scheme color，输入立即规范化并冻结；
- 只影响设置之后通过 `addText()` / `addRichText()` 创建的文字；
- rich-text run 的显式 color 覆盖 slide default；
- run 只有 transparency、没有显式 color 时，transparency 应作用于 slide default；
- 改变或清除 default 不回改已有 shape；
- duplicate 复制当前 transient default，move 保留，delete 清理；
- 写出时把有效颜色物化到新增 run，因此视觉结果可被 PowerPoint、LibreOffice 和其他客户端保留；
- 覆盖 `pptx/pptm/potx/potm/ppsx/ppsm`、Node、browser、实际 npm tarball、declarations、CLI、
  PptxGenJS public-output conformance 和真实 PPTX 验证。

本项不增加 table default color，不改变 chart、slide number、placeholder、layout 或 master 的颜色语义，不实现
declarative master definition，也不让 `ShapeModel.richText` 后续编辑重新读取 slide default。完整 placeholder/theme
继承与 advanced text 属于后续专项。

用户已明确要求实现方持续决定最佳方案，不设置询问或确认停顿点。本设计完成 self-review 后直接提交、推送，
随后写实施计划，并按每个可独立 review 的小项执行 tests、commit 与 push。

## 2. 当前状态与权威行为审计

### 2.1 Native 当前状态

`SlideModel` 目前没有 default color state。`addText()` 为每个非空 plain-text run 固定写
`a:schemeClr val="tx1"`；`addRichText()` 允许每个 run 通过 `RichTextRunStyle.color` 写 strict sRGB/theme
颜色，未显式设置的 run 同样写 `tx1`。`ShapeModel.richText` 可读取和整体替换这些已经物化的颜色。

现有基础设施可直接复用：

- `RichTextColor` 已表达 strict sRGB/theme discriminated union；
- rich-text normalizer、renderer 和 reader 已处理颜色、transparency 与 canonical DrawingML；
- `PresentationModel` 已维护 stable slide model identity 和 duplicate/move/delete lifecycle；
- OPC transaction、mutation journal、六格式、Node/browser/tarball 和 PptxGenJS evidence 流程均已建立。

OOXML 没有一个合法的 direct slide-level “default text color” 字段。`p:clrMapOvr` 是 theme color mapping，
不是 text default；`p:txStyles` 只属于 master；`a:lstStyle` 属于单个 text body。把该状态写入任一位置都会改变
标准语义或扩大为 master/placeholder 专项。

### 2.2 PptxGenJS 4.0.1 public behavior

审计只使用 public constructor、`addSlide()`、`slide.color`、`slide.addText()`、`slide.addTable()` 和
`write()` bytes，不读取 `_color`、`_slideObjects` 等私有字段。

实际行为如下：

- 新 slide 的 `slide.color` getter 返回 `undefined`；
- setter 原样保存 string；`'ff3399'` getter 仍返回 lowercase，但写出时规范化为 `FF3399`；
- runtime 接受合法 scheme token，例如 `accent1`，尽管 declaration 只写 `HexColor`；
- plain text 和 rich-text 中没有 local color 的 run 在创建时继承当前 slide color；
- rich-text 显式 run color 覆盖 slide color；
- 之后改变或清除 `slide.color` 不改变先前创建的对象；
- 清除后新文字回退到 hard-coded `000000`；
- 非法 string setter 本身不报错，写出时 warning 并回退 `000000`；
- table cell 不继承 `slide.color`，仍使用 table 自己的 `000000` default；
- serialized PPTX 只包含每个 run 的 `a:solidFill`，不存在 slide-level color state。

Native 不复制 delayed warning、非法值静默 fallback 或 lowercase getter。Setter 必须立即严格拒绝非法输入；
getter 返回规范化、detached、deep-frozen value。Native zero-input 继续使用 theme-aware `tx1`，不为了复制
PptxGenJS 的 hard-coded black 而改变既有输出。设置合法 custom default 后，两边的最终 run color 语义一致。

## 3. 方案比较

### A. 写自定义 OOXML extension

把 default 存入 `p:cSld/p:extLst` 可跨 reopen 恢复，但 PptxGenJS 没有这个持久状态，其他客户端不会理解，
还会给普通 slide 引入 private schema 和额外兼容风险。拒绝。

### B. 从已有 run 推断 default

读取一页中相同或最后一个 text color 作为 default 看似无需额外 state，但已有文字只代表历史物化结果，不能证明
创建时 default；混合 run、placeholder、theme inheritance 和 imported deck 会产生错误推断。拒绝。

### C. Transient direct model state + creation-time materialization（采用）

`PresentationModel` 按 active slide part URI 保存规范化 transient color，`SlideModel.color` 暴露 exact
PptxGenJS property name。`addText()` / `addRichText()` 捕获当前 default 并传给现有 renderer；renderer 仅在 run
没有显式 color 时使用它，最终仍写标准 `a:solidFill`。

该方案与 PptxGenJS 的真实生命周期一致，不伪造 OOXML state，并让创建结果在 write/reopen 后保持视觉和语义。
集中由 `PresentationModel` 管理 transient map，可以明确处理 duplicate/move/delete，避免 part URI 重用泄漏旧值。

## 4. 公共 API

复用现有 public type：

```ts
export type RichTextColor =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string };
```

`SlideModel` 新增：

```ts
export class SlideModel {
  get color(): Readonly<RichTextColor> | undefined;
  set color(value: RichTextColor | undefined);
}
```

选择 `color` 而不是 `defaultTextColor`，因为本专项目标是 PptxGenJS public API parity，且 slide 已有
`background`，不会与 background color 混淆。README 必须称其为“default text color”，并明确它不是已有文字的
批量 recolor API。

### 4.1 Normalization

Setter 接受：

```ts
slide.color = { kind: 'srgb', value: '#ff3399' };
// getter -> frozen { kind: 'srgb', value: 'FF3399' }

slide.color = { kind: 'scheme', value: 'accent1' };
slide.color = undefined;
```

规则：

- value 必须是 ordinary 或 null-prototype own-data object；
- 只允许 `kind` 和 `value` 两个 string key，不允许 symbol、accessor、继承值或额外字段；
- sRGB 接受可选 `#` 和恰好六位 hex，规范化为 uppercase、无 `#`；
- scheme 复用 `RichTextColor` 已支持 token 集合；
- input 在返回前复制，normalized value 使用 `Object.freeze()`；
- `undefined` 清除；`null`、string、array、class instance、Proxy/accessor 和 malformed value 立即失败；
- 语义相同的重复赋值不替换 snapshot，也不产生 package/mutation-journal 变化。

为避免颜色规则分叉，rich-text style normalizer 与 slide setter 共用导出的 internal
`normalizeRichTextColor(value, context)`；这不是新的根入口 public helper，只是 `@pptx/model` 内部实现。

## 5. State、lifecycle 与持久化边界

`PresentationModel` 持有：

```ts
readonly #slideDefaultColors = new Map<string, Readonly<RichTextColor>>();
```

并提供 package-internal methods 给 `SlideModel`：

```ts
getSlideDefaultColor(partUri: string): Readonly<RichTextColor> | undefined;
setSlideDefaultColor(partUri: string, value: Readonly<RichTextColor> | undefined): void;
```

这些方法不读取或修改 OPC package。行为如下：

- 新建或打开 presentation：map 为空；
- setter：更新当前 part URI；clear 删除 entry；
- write/writeBlob/download：不清空 map；同一 live document 之后新增文字仍继承；
- reopen bytes：新 model 的 getter 为 `undefined`，因为 serialized runs 已经拥有最终颜色，而 transient default
  本身不是 OOXML state；
- duplicate：package transaction 成功后复制 source 当前 snapshot 到 duplicate；失败时不创建 transient entry；
- move：part URI 不变，default 跟随 slide；
- delete：package transaction 成功后删除该 part URI entry；失败时保留；
- 新 slide 即使复用已删除的 part URI，也不会看到旧 default；
- 一个 slide 的 setter 不影响 sibling、layout、master 或 presentation default。

Duplicate copy 是 native lifecycle 的明确扩展：虽然 PptxGenJS 没有 duplicate API，但本库的 duplicate 表示复制当前
slide state。复制 frozen snapshot 不访问 package，也不改变 serialized content。

## 6. Text inheritance 与 XML

### 6.1 Plain text

`addText()` 的 validation 与 paragraph construction 不变，只把 captured slide color 传入 `textParagraphXml()` /
`defaultTextRunXml()`。每个非空 run 写：

```xml
<a:rPr ...>
  <a:solidFill><a:srgbClr val="FF3399"/></a:solidFill>
  ...
</a:rPr>
```

或 theme variant：

```xml
<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
```

没有 slide default 时继续写现有 `schemeClr tx1`。空 paragraph 沿用当前 behavior，不为不可见 empty run 额外制造
color state。

### 6.2 Rich text

`RenderRichTextOptions` 增加 internal `defaultColor?: Readonly<RichTextColor>`。`renderRun()` 的颜色解析为：

```ts
const color = run.style?.color ?? options.defaultColor ?? TX1;
```

因此：

1. local run color 最高优先；
2. 否则使用 slide default；
3. 否则保持 canonical `tx1`。

如果 run 只有 `transparency`，alpha 写在解析后的 color 上；如果 run 同时有 local color 与 transparency，alpha 写在
local color 上。font、highlight、outline、underline、glow、paragraph options 和 soft break 不变。

### 6.3 非目标 API

- `ShapeModel.richText = ...` 是已有对象的显式整体编辑，不重新读取 slide default；
- `addTable()` 不读取 slide default，与 PptxGenJS 4.0.1 实际输出一致；
- chart labels、slide number、notes、placeholder 和 master/layout text 不读取 slide default；
- 改变 `slide.color` 不扫描或修改现有 XML。

## 7. Transaction、错误与 no-op

- Setter 完成全部 descriptor/type/range validation 后才改变 transient map；失败保持旧 snapshot；
- Setter 和 getter不访问 package，所以 malformed slide XML 不阻止 default state 操作；
- equal assignment 保留 object identity；clear absent 是 exact no-op；
- `addText()` / `addRichText()` 继续在现有 OPC transaction 中完成 validation、render 和 insertion；任何失败同时保持
  package、relationships、shape cache 和 mutation journal 不变，transient default 不变；
- duplicate/delete 只在 package transaction 成功后提交对应 transient map change，防止 rollback 泄漏；
- 默认颜色不是 compatibility diagnostic：它不会生成非标准 package state。

## 8. PptxGenJS conformance 与 intentional differences

Permanent public-output evidence 覆盖：

- sRGB default 对 plain text、多 paragraph 和 rich text omitted-color run；
- theme scheme default；
- rich-text local color override；
- transparency over inherited default；
- 设置新 default 后旧对象保持旧颜色、新对象使用新颜色；
- clear 后 native 回到 `tx1`、PptxGenJS 回到 `000000` 的锁定差异；
- table 不继承；
- write/reopen 后 materialized run colors 可由 native strict reader 读回。

Intentional differences：

- native public value 是 `RichTextColor`，不是 ambiguous string；
- native sRGB getter uppercase，scheme 保持 token；
- native setter 立即拒绝非法值，不延迟到 write warning 并 fallback black；
- native omitted/cleared default 使用 theme-aware `tx1`，不复制 hard-coded black；
- native duplicate 复制 transient default；PptxGenJS 没有对应 lifecycle API。

Production package 不 import 或调用 PptxGenJS；它只出现在 adapter test 和独立控制文件生成流程中。

## 9. 验证策略

### 9.1 Focused tests

- strict normalization：sRGB/theme、freeze、copy、descriptor/accessor/prototype/symbol/extra keys、invalid values；
- state：set/get/clear/equal identity、zero package access、journal exact no-op；
- inheritance：plain/rich/multi-paragraph、local override、transparency、empty run/paragraph；
- non-target isolation：existing shapes、richText setter、table、slide number、layout/master；
- lifecycle：sibling isolation、duplicate copy、move retain、delete/reused URI clear、failed duplicate/delete rollback；
- write/reopen：runs 保持颜色，reopened slide transient getter 为 `undefined`；
- 六格式、strict/permissive compatibility profiles、root exports 和 declaration consumer。

### 9.2 Packed/runtime tests

- actual npm tarball Node ESM consumer 创建 sRGB/theme default，写出并重开；
- TypeScript consumer 编译 `Slide.color` 与 readonly `RichTextColor`；
- real-Chrome browser smoke 覆盖 set、plain/rich override、duplicate、writeBlob/reopen；
- CLI smoke 确认 package 可 inspect/validate；
- 两次 package build manifest 一致，全仓 clean build manifest 一致。

### 9.3 PPTX/client evidence

Native gallery 至少覆盖：zero-input tx1、sRGB、scheme、multiple paragraphs、rich inherited/local override、
transparency、change-over-time、clear、duplicate/move 和 table isolation。独立 PptxGenJS control 覆盖其 public valid cases。

两份 deck 必须：

- strict reopen；
- PowerPoint 2010 profile validation 无 error；
- LibreOffice/Poppler 180-DPI render 无 clipping/overflow，并逐页视觉检查；
- LibreOffice save/reopen 后 materialized run colors 与文字顺序保持；
- PowerPoint automation 只报告实际产生的 open/save/render evidence，不把启动成功误报为 round-trip。

## 10. 实施分解

1. 设计文档：本文件，自审、commit、push。
2. 实施计划：逐文件、逐测试、逐提交计划，自审、commit、push。
3. Public state/lifecycle：strict color normalizer、`SlideModel.color`、presentation transient map、
   duplicate/move/delete semantics。
4. Text materialization：plain/rich renderer inheritance、override/transparency/isolation、六格式与 reopen。
5. PptxGenJS conformance：public-output valid cases与 intentional differences。
6. Packed/runtime/client verification：root package、actual tarball、types、CLI、real Chrome、gallery、validation、
   LibreOffice/PowerPoint factual evidence。
7. Public docs/final closure：README/API/baseline/progress/changelog，下一项切到 master/layout/placeholder。

每项独立 review；focused tests、diff check 通过后立即 commit 与 push。`.pnpm-store/`、临时 deck、render、tarball 和
client round-trip artifacts 永不提交。

## 11. Success criteria

- `slide.color` 可 strict set/read/clear，snapshot normalized、detached、frozen；
- 后续 plain/rich text 正确继承，local run color 与 transparency precedence 正确；
- 已有内容、table、shape edit、layout/master 和 sibling 不受影响；
- lifecycle、rollback、no-op 和 URI reuse 没有 transient state 泄漏；
- serialized package 只含标准 run-level color，没有 custom slide extension；
- valid PptxGenJS public output 与 native final run state 对等，差异均有锁定测试；
- 六格式、全量 tests、typecheck、build、actual package、browser、CLI、gallery、validator 和客户端证据通过；
- 文档把 default color 从 unsupported 移到 supported，并把下一项更新为 master/layout/placeholder。

## 12. Self-review

- Placeholder scan：全部章节均为完整、可执行要求。
- Internal consistency：transient state、creation-time materialization、reopen boundary、duplicate/delete lifecycle 与
  PptxGenJS audit 一致。
- Scope：只处理 slide default text color；table、placeholder/master/layout inheritance 和 advanced text 明确后移。
- Ambiguity：public property、value type、precedence、clear、persistence、rollback、default difference 和 client gate
  均有唯一规则。
