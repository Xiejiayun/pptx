# Rich Text Run Hyperlinks Design

日期：2026-08-02

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `TextPropsOptions.hyperlink` 在 rich-text run 上的原生创建、读取与编辑能力：

- `RichTextRunStyle.hyperlink` 复用公开 strict `Hyperlink` value，支持 external URL、当前 presentation 内部 slide target 与 optional tooltip；
- rich slide text、layout/master rich text、declarative master rich text，以及 rich text placeholder creation/population 共用同一 run contract；plain text继续使用既有outer hyperlink contract；
- `ShapeModel.richText` 读取每个 run 的 direct hyperlink，并通过 whole-replacement setter 创建、替换或清除 run links；
- 显式 linked run 各自拥有 relationship，即使多个 run 使用完全相同的 target；
- outer `AddTextOptions.hyperlink` 继续作为 whole shape 与未覆盖 runs 的 uniform default；run-local value 可覆盖，`false` 可抑制该 default；
- duplicate、move、target deletion、rollback、write/reopen 与六种 presentation format 保持 target identity、引用安全和 package validity；
- hyperlink 与 font/color/underline/outline/glow/highlight、paragraph state、shape fill/line/arrows/shadow 和 whole-shape click 保持独立 ownership。

本小项不新增 hover、table-cell、image、chart、media、group、graphic-frame、custom-show、macro、program、sound 或 action-only navigation API。`ShapeModel.hyperlink` 仍只拥有 `p:cNvPr/a:hlinkClick`；run-local state 只通过 `RichTextRunStyle.hyperlink` 与 `ShapeModel.richText` 公开。

## 2. PptxGenJS 4.0.1 公开行为证据

权威基线是锁定依赖 `pptxgenjs@4.0.1/types/index.d.ts` 的 `TextPropsOptions.hyperlink?: HyperlinkProps`，以及 public `addText()` / `write()` 的真实 OOXML 输出。现有永久 conformance test 与本轮补充探针确认：

- rich array 的 run `options.hyperlink` 只写该 run 的 `a:rPr/a:hlinkClick`，不写 whole-shape click；
- 每个显式 linked run 分配独立 relationship；两个相同 URL 的相邻 runs 仍使用不同 IDs；
- URL run relationship 为 external hyperlink；internal slide run relationship 为 internal slide，并写 exact `action="ppaction://hlinksldjump"`；
- omitted tooltip 被物化为 `tooltip=""`；explicit tooltip 正确 XML escape；
- hyperlink 默认给 run 写 single underline；PptxGenJS 对 boolean `underline: false` 仍回退 single，native 保留既有显式 underline precedence，不复制该 falsy collapse；
- 带 explicit text color 的 URL run 会在 `hlinkClick` 中保留 hyperlink-color extension；reader 必须接受并在无关操作中保留这类合法 child；
- empty linked run 不生成 text run，但仍消耗一个 orphan relationship；native 将这种无可点击内容的输入视为无效并在 mutation 前拒绝；
- rich outer hyperlink 是 4.0.1 缺陷：shape 和 inherited runs 写 `rIdundefined`，但不创建 relationship；outer 与 local run 混用时，local run relationship有效，shape 与其他 runs 仍悬空；
- invalid target、coercion、双 target 与越界 slide 可能被宽松转换、console-only 拒绝或写成 dangling state。

Native 只对合法、可逆的 final-state 能力宣称对等。它不复制 orphan、`rIdundefined`、dangling target、coercion、falsy underline 或 console-only failure；所有不合法 input 在 package mutation 前严格拒绝。Native omitted tooltip 保持 attribute absence，explicit empty 保持 `tooltip=""`。

## 3. 方案比较

### 方案 A：在 `RichTextRun` 顶层增加 hyperlink

`{ text, hyperlink, style }` 看起来直接，但现有所有 run-local formatting 都集中在 `RichTextRunStyle`，PptxGenJS 也把 run hyperlink 放在 `TextProps.options`。该方案会制造第二套 run option 边界，并增加 normalization、snapshot 与 declarative clone 的分叉。

### 方案 B：`RichTextRunStyle.hyperlink`，按 target 去重 relationships

该方案易于实现，也可减少 relationships，但与 PptxGenJS 每个显式 run 独立 relationship 的真实语义不一致。共享关系还会让单 run target edit 必须额外 clone-on-write，并降低 run ownership 的可预测性。

### 方案 C：run-style direct state、显式 run 独立 relationship、outer default 可覆盖（采用）

`RichTextRunStyle.hyperlink?: Hyperlink | false` 与其他 run-local values 一起归一化。显式 hyperlink 每个 run 分配独立 relationship；omitted 值继承 outer default 的既有 shared relationship；`false` 抑制 outer default。Reader 从 direct run click 与 relationship graph 还原 strict value；`ShapeModel.richText` setter 以引用感知方式复用、替换和回收 relationships。

该方案同时满足 PptxGenJS valid per-run semantics、native strictness、现有 outer creation contract，以及后续其他 owner hyperlink 能力的一致 ownership 模型。

## 4. 公共 API

```ts
import type { Hyperlink } from './hyperlink.js';

export interface RichTextRunStyle {
  readonly hyperlink?: Hyperlink | false;
  // existing fields remain unchanged
}
```

Value 语义：

- omitted 或 own runtime `undefined`：没有 run-local override；创建时可继承 `AddTextOptions.hyperlink`；
- `false`：明确不写 run click，用于抑制 outer hyperlink；
- `Hyperlink`：写 direct run click，并覆盖 outer target/tooltip；
- `ShapeModel.richText` getter 只返回 direct legal run hyperlink；direct absence 不合成 outer effective value，也不返回 `false`；
- `ShapeModel.richText` setter 是 whole rich-text replacement，因此 omitted local hyperlink 清除该 replacement run 的 direct link；whole-shape click 保持不变。

`Hyperlink` 继续使用现有 mutually-exclusive `{ url, tooltip? } | { slide, tooltip? }`。不新增 target alias、raw relationship ID、action string 或 mutable relationship handle。

## 5. 归一化与错误语义

`normalizeStyle()` 显式允许 `hyperlink`，并调用现有 `normalizeHyperlink()`：

- style 必须是 descriptor-safe ordinary/null-prototype own-data object；unknown/accessor/symbol/inherited keys、arrays和 class instances 拒绝；
- hyperlink object 只允许 `url/slide/tooltip`，且恰好一个 target；
- URL 必须是 non-empty XML-safe string，不 trim、不 coercion、不限制 scheme；
- slide 必须是 positive safe integer，并在 owning presentation 中解析到唯一 current slide；
- tooltip 若 supplied 必须是 XML-safe string，允许 explicit empty；
- `false` 是唯一 suppression sentinel；`null`、true、numbers和 strings 拒绝；
- explicit hyperlink 只允许在 `run.text.length > 0` 时使用；empty linked run 在关系分配前拒绝，避免复制 PptxGenJS orphan 缺陷；
- normalized nested hyperlink 与 caller 立即脱离并冻结；
- paragraph/run/style/outer options/placeholder selector/transform/shape styles 与所有 run hyperlinks 全部验证并解析 target 后，才可开始 relationship 或 XML mutation。

任何 invalid plain/rich/placeholder/layout/master/declarative input 都必须保持 parts、relationships、shape order/cache、mutation journal 和 caller-visible model identity 零变化。

## 6. 创建期解析与渲染

Rich renderer 按每个 run 解析 effective link：

1. `style.hyperlink === false`：不写 link；
2. `style.hyperlink` 是 value：使用该 local value 与该 run 的独立 relationship；
3. `style.hyperlink` omitted：使用 outer default 与 outer shared relationship；
4. 两者都 absent：不写 link。

Relationship 规则：

- outer hyperlink 仍只分配一个 relationship，由 whole-shape click 和所有 inheriting non-empty runs 共享；
- 每个显式 run hyperlink 分配一个独立 relationship，即使 target/tooltip 与 outer 或其他 run 完全相同；
- `false` 和 empty runs 不分配 relationship；
- internal target 从 owning slide/layout/master part计算相对 URI，不根据 filename 猜测；
- relationship namespace 只在实际写 click 时绑定；
- run click 使用 canonical `renderShapeHyperlink()`，tooltip absence/empty、escaping 与 internal action 与 shape codec一致；
- 没有显式 underline 的 linked run 默认 `u="sng"`；false/true/17-style explicit underline 始终优先；
- `hlinkClick` 写在 run color/effect/highlight/underline/font children之后，保持 DrawingML schema order。

Placeholder layout source、materialized slide owner和 population result各自在自己的 part拥有 relationships；population 不污染 layout/master source。

## 7. Strict reader

`ShapeModel.richText` 读取 run hyperlink 时使用 owning part relationships和当前 presentation slide part order：

- 只检查 direct `a:rPr/a:hlinkClick`，不读取 descendant、hover、paragraph default或 whole-shape link；
- 要求唯一 direct namespace-correct `hlinkClick` 与唯一 namespace-correct relationship ID；
- external URL 要求 hyperlink relationship、External mode与 non-empty target；run reader兼容 PptxGenJS canonical `action=""`，但拒绝其他 non-empty action；
- internal slide 要求 exact `ppaction://hlinksldjump`、slide relationship、Internal mode和唯一 current slide target；
- tooltip absence与 explicit empty保持可区分；
- PptxGenJS `invalidUrl`、`tgtFrame`、`history`、`highlightClick`、`endSnd` 和合法 `extLst` 可被保留，但不进入 public `Hyperlink` value；
- duplicate/malformed/wrong namespace/wrong type或mode/dangling relationship返回 no local hyperlink，不猜测、不修复、不修改 bytes；
- reader 返回 detached normalized value，不暴露 `_rId` 或任何 relationship implementation detail。

Getter 的 direct-state snapshot中，outer-created inherited runs会因为实际存在 direct click而返回 hyperlink；没有 direct click的 run保持 omitted，即使 whole shape可点击。

## 8. 编辑与 relationship 生命周期

`ShapeModel.richText = paragraphs` 延续 whole-replacement text semantics，并新增 hyperlink graph管理：

- 先归一化全部 input并解析全部 internal targets；任何失败发生在 mutation前；
- same normalized snapshot assignment是 exact parts/relationships/journal no-op；
- 对同 paragraph/run index的 strict existing link，同 target edit可复用 relationship并只改变 tooltip XML；
- target改变时，relationship只有该 run引用则可原位更新；若与 whole-shape或其他 run共享则 clone-on-write；
- 新 linked run分配新 relationship；removed/cleared run link只在 package XML中已无引用时回收 relationship；
- whole-shape `ShapeModel.hyperlink` 不被 rich-text setter读取、合并或覆盖；
- imported PptxGenJS per-run IDs保持独立；explicit same-target runs不会在 edit中被合并；
- unsupported/malformed run click在 getter中省略，在 unrelated shape edits中 byte-preserved；显式 rich-text replacement继续按现有 whole-owner contract替换 text body；
- outer transaction rollback恢复 part bytes、relationship rows、target updates、journal与 live model state；
- target slide deletion使用现有 package-wide DrawingML cleanup移除 slide/layout/master run clicks并回收 incoming relationships；
- duplicate external/other-slide links保持 target identity，duplicate self-links retarget duplicate自身，move/insert/delete-before-target只改变 getter ordinal。

`ShapeModel.text` 继续遵循现有 plain-text template editor边界；本小项不重新定义跨 run的 plain-string replacement。用户需要精确 per-run link编辑时使用 `ShapeModel.richText`。

## 9. PptxGenJS conformance 与 intentional differences

永久 public-output tests覆盖：

- URL/internal/tooltip omitted-empty-custom；
- one linked run、mixed linked/unlinked runs、two identical-target runs、styled/color runs；
- per-run only不写 whole-shape click，每个 linked run有独立 relationship；
- outer-only、outer+local override与native strict repair；
- empty linked run orphan、boolean underline collapse、invalid/coercible/dangling inputs。

Native valid per-run output与 PptxGenJS在 target、relationship type/mode、action、tooltip effective state和run selection上对等。Native有意保留以下差异：

- omitted tooltip保持 absence而不是强制 `tooltip=""`；
- explicit underline false/style优先，不复制 PptxGenJS single fallback；
- empty linked run拒绝，不创建 orphan relationship；
- rich outer和outer+local产生全部合法关系，不复制 `rIdundefined`；
- invalid input严格零变更拒绝，不复制 coercion、console-only或dangling behavior；
- native setter提供PptxGenJS没有的existing-deck semantic editing和relationship GC。

## 10. 验证策略

### Model / codec

- strict normalization、false suppression、caller detachment、frozen nested hyperlink和empty-run rejection；
- URL/internal/self/tooltip、default underline与explicit underline precedence；
- outer inherited、local override、false suppression、same-target independent IDs；
- strict imported PptxGenJS run reading，包括 URL `action=""` 与 color extension；
- malformed/duplicate/wrong namespace/type/mode/action/dangling state不被猜测；
- read → edit → clear、same-value no-op、unique update、shared clone-on-write、reference-aware GC和rollback。

### SDK / lifecycle

- slide/layout/master direct text、placeholder create/populate、declarative master definitions；
- duplicate external/other/self、move/insert/delete target、target deletion和source isolation；
- all six formats、write/reopen、stable wrapper identity和validator零 package error；
- fill/line/arrows/shadow/whole-shape click及其他 run styles保持独立。

### Release gates

- focused model/SDK/root/adapter tests；
- full Vitest、performance、TypeScript strict typecheck与project/package/declaration builds；
- actual 57-file tarball的installed Node/types/browser/CLI smoke；
- real Chrome immediate/read/edit/clear/reopen与console/page/network零错误；
- internal-only PowerPoint 2010 profile 0 errors / 0 warnings，external-only portability warning符合预期；
- README、package README、API、compatibility、progress与changelog同步。

## 11. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. `RichTextRunStyle.hyperlink` 在所有 rich-text creation surfaces支持 strict URL/internal/tooltip与false suppression；
2. `ShapeModel.richText` 能严格读取、创建、替换和清除 per-run links，并保持 whole-shape ownership独立；
3. explicit run relationships独立，outer defaults共享，clone/update/GC/delete/duplicate/rollback/write-reopen均有测试证明；
4. PptxGenJS 4.0.1 valid output与 intentional defect corrections有真实 public-output证据；
5. packed Node/types/browser/CLI、full suite、performance、typecheck、build和compatibility validator全部通过；
6. 文档移除 per-run hyperlink gap，并选择下一个真实 parity缺口。
