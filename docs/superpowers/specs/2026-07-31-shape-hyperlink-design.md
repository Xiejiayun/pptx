# Shape Hyperlink Design

日期：2026-07-31

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项为原生 preset shape 增加 PptxGenJS 4.0.1 `HyperlinkProps` 的 shape-level vertical slice：

- `slide.addShape()` 创建时可指定 external URL 或 presentation 内部 slide target；
- 支持 optional direct tooltip；
- 读取 existing deck 中唯一、direct、合法的 `a:hlinkClick` 及其 relationship；
- 通过 `ShapeModel.hyperlink` 替换 URL、内部目标、tooltip，或清除整个 direct click hyperlink；
- relationship 更新采用 reference-aware clone-on-write，clear 后清理不再引用的 relationship；
- 内部链接使用实际 slide part identity，因而 move/insert/delete 其他 slide 后仍指向原目标；
- duplicate、target deletion、transaction rollback、write/reopen、六种 presentation format 和 live model identity 保持一致；
- 保留 `cNvPr` 上 `hlinkHover`、extensions、其他合法 hyperlink attributes/children、shape geometry/style/text 与 unrelated relationships/parts。

本设计只开放 `p:sp` 的 whole-shape click hyperlink。它不实现 text-run、table-cell、image、chart、media、group 或 graphic-frame hyperlink creation API，不实现 hover hyperlink、relative/file safety policy、email helper、custom-show/macro/program/action-only navigation、sound action、history/highlight editor，也不改变 `addText()` 的 creation options。这些调用方后续复用同一 public `Hyperlink` value，但各自的 OOXML 容器和 lifecycle 独立实施。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `HyperlinkProps` 类型与 `addShape()` / `write()` 真实输出。类型只声明：

```ts
interface HyperlinkProps {
  slide?: number;
  url?: string;
  tooltip?: string;
}
```

真实 public-output 审计结果如下：

- `{ url }` 在 `p:cNvPr` 内写 `<a:hlinkClick r:id="..." tooltip="..."/>`，relationship type 为 `.../hyperlink`，`TargetMode="External"`；
- `{ slide: 2 }` 写同一 click element，并增加 `action="ppaction://hlinksldjump"`；relationship type 为 `.../slide`，target 为 `slide2.xml`；
- omitted tooltip 也物化为 `tooltip=""`，XML metacharacters 被 escape；
- URL relationship target 的 ampersand被正确 escape，并在 import 后恢复原字符串；
- empty object、empty URL、slide zero 或非-object runtime value只打印 console error或被忽略，最后没有 hyperlink；
- truthy non-string URL被字符串化并输出 external relationship；
- negative、fractional或超出 presentation 的 slide number仍被拼成 `slide-1.xml`、`slide1.5.xml` 或 `slide99.xml`，形成 dangling relationship；
- 同时提供 truthy `url` 与 `slide` 会写两个 `hlinkClick`，并生成错误的 internal target `slidehttps://...xml`；PowerPoint 2010 profile 可直接诊断为 dangling relationship。

Native 只对正常、可移植到本库 direct-state contract 的行为宣称对等。它不复制 console-only failure、coercion、falsy ignore、双 click 或 dangling relationship 缺陷；这些输入在任何 package mutation 前严格拒绝。Native omitted tooltip 保持 attribute absence，explicit empty 保持 `tooltip=""`；二者在 PptxGenJS 4.0.1 中会折叠为 explicit empty，但有效显示行为相同。

## 3. 方案比较

### 方案 A：原样复制三个 optional fields

直接声明 `{ url?: string; slide?: number; tooltip?: string }` 最接近 PptxGenJS 类型，但编译期允许缺少 target或同时提供两个 target。4.0.1 runtime 已证明这些状态会静默忽略或生成损坏关系，因此该形状不适合作为严格 native contract。

### 方案 B：使用 `{ kind: 'url' | 'slide', target, tooltip? }`

Discriminated union 最明确，也方便未来增加 custom show/action，但会让常见调用与 PptxGenJS 不对等，并迫使 adapter 做额外字段映射。当前只有两个 target，不需要额外 `kind`。

### 方案 C：PptxGen-shaped mutually exclusive union（采用）

Public value 保留 `url`、`slide`、`tooltip` 命名，但通过 union 保证恰好一个 target。URL branch 禁止 defined slide，slide branch 禁止 defined URL；runtime normalizer同样执行 descriptor-safe exclusivity、类型和范围检查。

该方案兼顾 PptxGenJS 常规调用形式与 native strictness。未来 text、image、table 等 API可复用同一 `Hyperlink`，而 shape-level adapter只拥有 `p:cNvPr/a:hlinkClick`。

## 4. 公共 API

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

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly arrows?: ShapeArrows;
  readonly hyperlink?: Hyperlink;
}

class ShapeModel {
  get hyperlink(): Hyperlink | undefined;
  set hyperlink(value: Hyperlink | undefined);
}
```

`slide` 是 assignment 时 presentation 中 one-based slide number，与 PptxGenJS 一致。OOXML 不保存 ordinal，而保存目标 slide relationship；getter根据当前 presentation order重新计算 one-based number。因此移动目标 slide、在其前面插入/删除其他 slide时，link identity不变，getter number随当前位置更新。

Getter 返回 detached、frozen direct-state snapshot。Absent或unsupported hyperlink 返回 `undefined`。Tooltip attribute absent时省略 `tooltip` property；direct empty attribute返回 `tooltip: ''`，不会与 absence 合并。

Setter 是 whole replacement：URL 与内部 slide target 可相互切换；omitted tooltip 清除 direct tooltip attribute；explicit empty写 direct empty；`undefined` 清除整个 supported direct `hlinkClick` 并按引用清理 relationship。

## 5. 输入归一化与目标解析

所有 public input 在任何 package mutation 前完成：

- hyperlink 必须是 ordinary 或 null-prototype object，只接受 own data properties；
- symbols、accessors、arrays、class instances、inherited-only fields和 unknown keys全部拒绝；
- 只允许 `url`、`slide`、`tooltip`；own runtime `undefined` 与 field absence等价；
- exactly one of URL/slide必须 defined；两个都 defined或两个都 absent均拒绝；
- URL必须是 non-empty XML-safe string；不 trim、不强制 HTTP scheme，也不对 `mailto:`、fragment、custom URI 或 relative lexical target做隐式改写；
- slide必须是 positive safe integer，并在 assignment 时存在于当前 presentation；zero、negative、fraction、NaN、infinity、numeric string、boolean和越界值拒绝；
- tooltip若 defined必须是 XML-safe string，允许 explicit empty；
- normalized value与 caller立即脱离并冻结。

Public value错误使用 `TypeError` 或 `RangeError`。Native 不接受 `_rId`、`target`、`kind`、deprecated aliases或额外 metadata。Invalid create/edit 在 relationship、slide XML、mutation journal或 live model map变化前失败。

## 6. OOXML 与 relationship 语义

External URL canonical output：

```xml
<p:cNvPr id="2" name="Website">
  <a:hlinkClick r:id="rId2" tooltip="Visit site"/>
</p:cNvPr>
```

```xml
<Relationship Id="rId2"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://example.com/path?x=1&amp;y=2"
  TargetMode="External"/>
```

Internal slide canonical output：

```xml
<p:cNvPr id="3" name="Next">
  <a:hlinkClick r:id="rId3" action="ppaction://hlinksldjump"/>
</p:cNvPr>
```

```xml
<Relationship Id="rId3"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
  Target="slide2.xml"/>
```

Reader只检查 shape 唯一 direct PresentationML `p:cNvPr` 下唯一 direct DrawingML `a:hlinkClick`：

- external branch要求唯一 relationship-namespace `id` attribute、唯一 matching relationship、exact hyperlink relationship type和 `TargetMode="External"`，且 direct `action` absent；
- internal branch要求 exact `action="ppaction://hlinksldjump"`、exact internal slide relationship type、resolved target是当前 presentation 中唯一 slide part；
- tooltip若存在必须是唯一 unqualified attribute；absence与empty分别保留；
- `hlinkHover` 不进入 snapshot，并在 click edit/clear中保留；
- `history`、`highlightClick`、`tgtFrame`、`invalidUrl`、`endSnd`、sound/ext children等不属于 public value，但在 supported click replacement中原 bytes保留；
- alternate legal DrawingML与relationship prefixes可读取和编辑；lexical prefixes不进入 snapshot；
- missing/duplicate relationship、wrong type/mode、dangling/non-presentation slide target、action/type不匹配、multiple click、duplicate/qualified owned attributes、wrong namespace或unsafe child order返回 `undefined`；mutation以 `ModelParseError` zero-change拒绝。

本小项沿用仓库当前 Transitional OOXML relationship namespace；Strict-package relationship URI支持留给统一 strict-package audit，不在单一 shape feature中局部引入。

## 7. Lossless replace、clone-on-write 与 clear

Setter先 normalize和 resolve target，再进入 package transaction：

- fully supported same target与相同 direct tooltip是 exact bytes/journal no-op；
- tooltip-only edit只 patch `hlinkClick@tooltip`，不改 relationship；
- unique relationship只被当前 click引用时，target branch replacement可原位更新 type/target/mode并保留 relationship ID；
- 同一 relationship ID被 slide XML其他 relationship attributes引用时，target replacement分配新 relationship并只 retarget当前 click；其他引用保持旧 target；
- relationship target替换与 click `r:id` / `action` / `tooltip` edits在同一 transaction中完成；失败时全部回滚；
- absent click creation在 `hlinkHover` / `extLst` 前插入 canonical click；self-closing `cNvPr`只在需要时安全展开；
- clear删除整个 direct click，并仅在该 ID 已无任何 slide XML引用时删除 relationship；共享 relationship不会被过早回收；
- target branch replacement后，旧 relationship仅在不再引用时删除；
- edits只 patch owned attributes/element span，保留 `cNvPr` name/descr/title/hidden、hover、extensions、hyperlink extra state、shape properties/text及neighbor bytes。

Relationship reference count按当前 slide XML中 expanded-name-correct relationship ID attributes计算，不只检查 shape links，因此 existing text/image/opaque references也受到保护。Ambiguous relationship IDs或unsafe element/attribute state不会被猜测、合并或部分改写。

## 8. Presentation 生命周期

- Duplicate slide沿用 dependency clone：external target保持不变，link到其他 slide保持同一 target，self-link自动 retarget到 duplicate自身；shape XML relationship ID与raw extra state保持。
- Move slide不改 hyperlink relationship；getter根据新 order返回更新后的 one-based target number。
- 在目标 slide前 insert/delete其他 slide只改变 getter ordinal，不改变 target identity。
- Delete source slide随 source part删除其 relationships，不影响 external资源或其他 slides。
- Delete target slide前，presentation lifecycle会删除其他 slide中引用该 target relationship的 DrawingML `hlinkClick` / `hlinkHover` elements；随后现有 part deletion删除 incoming relationship，避免留下 dangling `r:id`。与目标无关的 link与 non-link bytes保持。
- Target deletion cleanup覆盖 existing text/image containers中的同类 DrawingML link引用，但不因此公开它们的 create/edit API。
- Outer transaction rollback恢复 source/target slide bytes、relationships、parts、presentation order、mutation journal与 live identities。

## 9. 与其他 shape 能力的边界

Hyperlink位于 `p:cNvPr`，不与 `p:spPr` 下 geometry、fill、line、arrows、effects或 `p:txBody` 共享 ownership：

- geometry/fill/line/arrows/text/transform edits必须原样保留 supported或unsupported hyperlink bytes与 relationship；
- hyperlink edits必须原样保留 geometry、adjustments、fill、line/arrows、effects、text、non-visual metadata、hover和extensions；
- AddShape renderer只在 hyperlink存在时改变 `cNvPr` shape与增加 relationship；omitted/runtime-undefined hyperlink保持已发布 shape bytes不变；
- `ShapeModel.hyperlink` 可读取/编辑任何可解析 `p:sp` 的 whole-shape click link，包括 existing text shape；`addText()` creation option仍属于后续 text parity小项。

External relationship在 compatibility validator中会产生 expected portability warning，但不是 package error；内部 slide link与无 hyperlink deck应保持 zero error/zero warning。

## 10. 验证策略

### Internal/value contract

- URL、internal slide、tooltip absent/empty/Unicode/XML metacharacters；
- descriptor-safe、null-prototype、detachment、frozen snapshot与 getter-free rejection；
- missing/both targets、empty/invalid URL、invalid/越界 slide、invalid tooltip、unknown/symbol/accessor字段零 mutation拒绝；
- exact external/internal render/read、alternate prefixes、tooltip direct-state distinction；
- wrong type/mode/action、dangling target、duplicate click/relation/owned attributes与 namespace lookalike严格拒绝。

### Creation and editing

- addShape URL/internal/self link creation、immediate snapshot、stable identity与 caller detachment；
- same-value no-op、tooltip add/replace/empty/clear、URL replace、slide replace、URL↔slide切换、whole clear；
- shared relationship clone-on-write与 reference-aware GC；
- hover/extra attrs/sound/ext、shape style/text/ext、neighbor和 unrelated relationship bytes保留；
- malformed/ambiguous existing state与 outer rollback exact zero-change；
- duplicate external/other-slide/self links、move、insert/delete ordinal变化、target deletion cleanup、write/reopen与六格式。

### PptxGenJS and release evidence

- public-only fixture覆盖 URL、slide、tooltip、omitted/empty、both target、zero/negative/fraction/out-of-range slide、non-string URL和 non-object hyperlink；
- supported final semantics与 PptxGenJS 4.0.1 public output对照，strict divergences和其 malformed outputs单独断言；
- actual tarball Node/browser/types smoke覆盖 create/read/edit/clear、external/internal relationship和 target lifecycle；
- `pnpm check`、performance、build、PowerPoint 2010 validation、LibreOffice click-preserving round-trip、PDF/visual regression、overflow和 artifact-tool import全通过。

## 11. 完成门禁

只有以下条件全部满足，本小项才能把 shape hyperlink标记为已支持：

1. preset shape可 create/read/replace/clear/reopen external URL和内部 slide link；
2. tooltip direct state、target identity、shared relationship clone-on-write/GC与 target deletion cleanup有 exact tests；
3. ambiguity、same-value no-op、rollback、identity、duplicate/move/delete、unknown-byte preservation和 malformed zero-mutation有证明；
4. PptxGenJS public-output、packed Node/browser/types、full suite、validator和真实文件 QA全通过；
5. README、CHANGELOG、API docs和 compatibility matrix明确 supported scope、strict divergences，以及剩余 hover/text-run/table/image/action hyperlink缺口。
