# Text Shape Hyperlink Creation Design

日期：2026-08-02

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { hyperlink })` 对应的原生创建能力：

- `AddTextOptions.hyperlink` 复用现有 strict `Hyperlink` value，支持 external URL、presentation 内部 slide target 与 optional tooltip；
- text shape 创建时写一个 `p:cNvPr/a:hlinkClick`，并在每个实际文本 run 的 `a:rPr` 中写同目标 `a:hlinkClick`；
- 整个 text shape 的 shape-level link 与所有本次创建的 run-level links 共享一个 relationship，与 PptxGenJS 合法输出一致；
- plain/rich text、layout/master text、declarative master text，以及 text placeholder creation/population 共用同一 contract；
- 创建结果立即通过现有 `ShapeModel.hyperlink` 读取 whole-shape direct state；
- duplicate、move、target deletion、rollback、write/reopen 与六种 presentation format 保持 relationship identity 和 package validity；
- hyperlink 与 fill/line/arrows/shadow/text body 保持独立 ownership。

本小项只开放统一 text-shape creation option。它不新增 per-run public hyperlink 字段或 run hyperlink getter/setter，不新增 hover、table-cell、image、chart、media、group、graphic-frame、custom-show、macro、program、sound 或 action-only navigation API。现有 `ShapeModel.hyperlink` 仍只拥有 `p:cNvPr/a:hlinkClick`；run-level 读取、编辑与 rich-text round-trip 是后续独立小项。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextPropsOptions.hyperlink?: HyperlinkProps` 类型，以及 `addText()` / `write()` 的真实输出。审计结果如下：

- plain string URL 会在 `p:cNvPr` 和唯一 `a:rPr` 中各写一个 `a:hlinkClick`，两者共享同一 external hyperlink relationship；
- plain multiline text 会为每个物化 run 写 click link，并与 `p:cNvPr` 共享同一 relationship；
- internal slide link 在 shape 与 runs 中都写 `action="ppaction://hlinksldjump"`，relationship 指向目标 slide part；
- omitted tooltip 被物化为 `tooltip=""`；explicit tooltip 会正确 XML escape；
- hyperlink 默认给 run 增加 single underline，并通过 hyperlink-color extension 保持 text color；
- rich run 自己携带 hyperlink 时只写该 run link，每个 linked run 分配 relationship，不写 whole-shape link；
- rich array 只在 outer options 提供 hyperlink 是 4.0.1 缺陷：输出 `rIdundefined` 且不创建 relationship，shape 与 runs 都形成 dangling references；
- empty/malformed target、coercion、双 target 与越界 slide 沿用 shape hyperlink 审计中已确认的宽松、console-only 或 dangling behavior。

Native 只对合法、可逆的 final-state 能力宣称对等。它不复制 `rIdundefined`、dangling target、coercion、双 click 或 console-only failure；所有不合法输入在 package mutation 前严格拒绝。Native omitted tooltip 保持 attribute absence，explicit empty 保持 `tooltip=""`；两者与 PptxGenJS 的有效显示行为相同。

Native 对 `addRichText(..., { hyperlink })` 采用 PptxGenJS intended semantics 的严格修复：合法 outer link 统一应用于 whole shape 和所有物化 runs，而不是生成损坏引用。Per-run heterogeneous links保留给后续 rich-text hyperlink 小项。

## 3. 方案比较

### 方案 A：只写 whole-shape hyperlink

直接复用现有 preset-shape renderer 和 `ShapeModel.hyperlink`，改动最少，但漏掉 PptxGenJS 在 text runs 中实际写出的 click links。用户点击文字与点击形状空白区域的行为可能不同，不能称为 text hyperlink 对等。

### 方案 B：只写 run-level hyperlink

更接近通常的文本链接概念，但漏掉 PptxGenJS 的 `p:cNvPr` link，也无法让创建结果立即通过现有 `ShapeModel.hyperlink` 读取。它还会使 plain text 与 native preset-shape link 的生命周期割裂。

### 方案 C：一个 normalized value、一个 relationship、双层渲染（采用）

`validateAddTextOptions()` 只归一化一次。创建 surface 在 target 校验后分配一个 relationship；`textShapeXml()` 在 `p:cNvPr` 写 shape click，并把同一 relationship ID 传给 plain/rich run renderer。合法 PptxGenJS 输出、relationship graph 与 immediate native getter由同一个 source of truth驱动。

该方案保持改动集中，也为后续 per-run API 预留边界：本小项的 outer option是 uniform default，未来 run-local value可覆盖或清除该 default，而不改变 whole-shape API。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly hyperlink?: Hyperlink;
  // existing fields remain unchanged
}
```

`Hyperlink` 继续使用现有公开 mutually exclusive union：

```ts
export type Hyperlink =
  | {
      readonly url: string;
      readonly slide?: never;
      readonly tooltip?: string;
    }
  | {
      readonly slide: number;
      readonly url?: never;
      readonly tooltip?: string;
    };
```

`AddPlaceholderOptions` 基于 `AddTextOptions`，因此自动获得同一字段。`SlideLayoutModel` / `SlideMasterModel` 的 direct text methods 同样复用该字段；`defineSlideMaster()` 的 closed text option-key reader显式允许 `hyperlink`。

本小项不改变 `RichTextRun` 或 `RichTextRunStyle`。`ShapeModel.hyperlink` getter只读取创建出的 `p:cNvPr` link，返回 detached frozen snapshot；setter仍只替换该 whole-shape direct click，不隐式改写 run links。

## 5. 归一化、目标解析与错误语义

`validateAddTextOptions()` 调用现有 `normalizeHyperlink()`，保持 shape hyperlink 的 strict contract：

- value 必须是 ordinary 或 null-prototype own-data object，只允许 `url/slide/tooltip`；
- symbols、accessors、arrays、class instances、inherited fields、unknown keys与 aliases拒绝；
- exactly one of URL/slide必须 defined；own runtime `undefined` 与 absence等价；
- URL必须是 non-empty XML-safe string，不 trim、不 coercion、不限制 scheme；
- slide必须是 positive safe integer，并在 assignment 时对应当前 presentation 中的实际 slide；
- tooltip若 defined必须是 XML-safe string，允许 explicit empty；
- normalized value立即与 caller脱离并冻结；
- invalid plain/rich/placeholder/master definition在 relationship、part bytes、shape cache、shape order和 mutation journal变化前失败。

所有 text、paragraph、style、placeholder selector/name/identity、transform、fill/line/arrows/shadow 与 hyperlink value先完成输入校验。内部 slide target也在分配 relationship前解析。之后的 renderer、part replacement和 relationship添加位于同一 outer transaction；任何 package/OOXML异常都完整回滚。

## 6. OOXML 与 relationship 语义

External URL canonical shape output：

```xml
<p:cNvPr id="2" name="Website">
  <a:hlinkClick r:id="rId2" tooltip="Visit"/>
</p:cNvPr>
```

```xml
<a:rPr lang="en-US" u="sng" dirty="0">
  <!-- existing color/font state -->
  <a:hlinkClick r:id="rId2" tooltip="Visit"/>
</a:rPr>
```

```xml
<Relationship Id="rId2"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://example.com/path?x=1&amp;y=2"
  TargetMode="External"/>
```

Internal slide link在两个 click elements上增加 exact `action="ppaction://hlinksldjump"`，relationship type 为 `.../slide`、mode 为 internal、target 为从当前 slide/layout/master part到目标 slide part的相对 URI。

渲染规则：

- 一个 text shape只创建一个 hyperlink relationship；whole-shape click和所有本次物化的 non-empty runs共享其 ID；
- empty paragraph没有 synthetic text run，仍保留 whole-shape click；
- plain multiline 的每个 non-empty paragraph run都链接；
- rich outer hyperlink应用到每个实际 run，包括有独立 font/color/underline style的 run；
- hyperlink在 run中默认补 `u="sng"`；如果 native rich run显式提供 underline state，显式 state优先；
- hyperlink child位于 existing run color/effects/highlight/underline/font children之后，保持合法 DrawingML child order；
- relationship namespace只在 supplied hyperlink时引入；omitted/runtime-undefined hyperlink不改变已发布 text XML或 relationships；
- shape click与 run clicks统一调用现有 canonical hyperlink renderer，tooltip absent/empty与 XML escaping保持一致。

## 7. Ownership 与生命周期

- hyperlink relationship与 shape XML在同一 transaction创建；失败不留下 orphan relationship或消费后的 visible shape ID；
- `ShapeModel.hyperlink` 只拥有 `p:cNvPr/a:hlinkClick`。因为创建时 relationship还被 run links引用，改变 whole-shape target会按现有 reference-aware规则 clone-on-write；run links继续指向原 target；
- whole-shape tooltip-only edit可继续复用 shared relationship；clear whole-shape link不会回收仍被 runs引用的 relationship；
- 后续 `shape.text` / `shape.richText` 编辑按其当前 text-body ownership运行。本小项不把 run hyperlink塞入尚未公开的 rich-text snapshot；run edit parity由后续小项补齐；
- fill/line/arrows/shadow/geometry/transform edits不改 whole-shape或 run hyperlink bytes；hyperlink创建不改这些 style值；
- duplicate slide沿用 dependency clone：external target保持；other-slide link保持 target identity；self-link relationship自动 retarget到 duplicate自身，所有共享 references随 clone保持一致；
- move/insert/delete before target不改 relationship target；`ShapeModel.hyperlink.slide` 按当前 order重新计算 ordinal；
- delete target使用现有 DrawingML cleanup同时删除 `cNvPr` 与 `rPr` 中指向目标的 click/hover elements，再删除 incoming relationship，避免 dangling IDs；
- outer transaction rollback恢复 source/target bytes、relationships、parts、order、journal与 live identity；write/reopen和六格式保持 supported state。

## 8. Placeholder、layout/master 与 declarative definitions

- `addPlaceholder()` 的 plain/rich prompt可直接创建 uniform hyperlink；
- `addText()` / `addRichText()` population到 inherited placeholder时使用 population options的 link，在 materialized slide part创建 relationship，不污染 layout/master source；
- layout/master direct text及 placeholder source的 link relationship属于对应 owner part；
- declarative `defineSlideMaster()` text/placeholder options接受同一 field，并经过相同 descriptor-safe clone、normalization与 transaction流程；
- internal link从 owner part解析到 presentation slide identity，不按猜测的文件名拼接；越界 target在任何 owner mutation前拒绝。

## 9. 验证策略

### Model/SDK

- URL、internal slide、self-link、tooltip absent/empty/Unicode/XML metacharacters；
- exact `cNvPr` + all `rPr` references、single shared relationship、default underline与 namespace binding；
- omitted/runtime-undefined保持基线 bytes；caller detachment、frozen immediate getter与 strict invalid zero mutation；
- plain empty/single/multiline、rich empty/multiple styled runs、placeholder create/populate、layout/master direct与 declarative definitions；
- simultaneous fill/line/arrows/shadow、whole-shape edit/clear后的 run relationship preservation；
- duplicate external/other/self link、move/insert/delete target ordinal、target deletion、rollback、write/reopen和六格式。

### PptxGenJS

- public-only fixture覆盖 plain URL/slide、tooltip omitted/empty/custom、multiline、rich outer bug与 per-run valid baseline；
- supported plain final semantics对照 shape/run references、relationship type/mode/target/action和 underline；
- omitted tooltip direct-state差异、strict rejection、rich outer bug修复与暂未公开 per-run input明确记录为 intentional differences。

### Release gates

- focused model/SDK/adapter/root tests；
- full Vitest、performance、typecheck、build；
- actual tarball Node/types/browser/CLI smoke覆盖 URL/internal links、shape/run XML与 reopen；
- real Chrome immediate/detached/reopen state、PowerPoint 2010 validation、LibreOffice round-trip与 PDF/visual regression；
- README、API、compatibility、progress与 changelog更新。

## 10. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 所有 text creation surfaces接受 strict uniform hyperlink，并写入一个 valid shared relationship及 shape/run click state；
2. validation-before-mutation、placeholder/master isolation、relationship ownership、clone/cleanup/rollback与 lifecycle有测试证明；
3. PptxGenJS 4.0.1 supported semantics和 intentional differences有真实 public-output证据；
4. packed Node/types/browser/CLI、full suite、performance、typecheck、build与 compatibility validator全通过；
5. 文档明确 creation scope、现有 whole-shape editor边界，以及后续 per-run read/edit工作。
