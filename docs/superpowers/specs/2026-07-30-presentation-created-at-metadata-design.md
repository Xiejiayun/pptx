# Presentation Created-At Metadata Design

## 目标与边界

为 presentation core properties 增加原生创建时间的创建、读取、编辑和清除能力，使 `PptxDocument.create({ createdAt })` 与 `document.createdAt` 同时覆盖从零创建和已有 PPTX 编辑，并能精确导入 PptxGenJS 4.0.1 每次公开 `write()` 生成的 `dcterms:created`。

本小项只拥有 OPC core-properties part 中 direct Dublin Core Terms `created` element及其必需的 `xsi:type="dcterms:W3CDTF"` 类型状态。它不同时实现 `dcterms:modified`、保存时自动刷新、文件系统时间、revision 递增、lastModifiedBy 同步、custom properties或审计策略。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按 PowerPoint-compatible typed property、严格公开 API、preserve-first 编辑、OPC 生命周期和 namespace/QName correctness 定稿。

## 方案选择

考虑三种实现：

1. 复制现有 core-properties relationship、part 和 source-span 逻辑到 created-at helper；不采用。它会在下一项 modified-at 再次复制相同安全边界。
2. 把 created 当作普通无类型文本，只校验 element namespace；不采用。OOXML要求 `xsi:type` 指向 `dcterms:W3CDTF`，忽略该状态会把 malformed 或语义不同的 property 暴露为受支持值。
3. 为现有 descriptor-driven core text helper 增加可选 QName 类型约束，再用 created-at wrapper负责日期时间 lexical validation；采用。既保持 title/author/subject/revision/lastModifiedBy 行为不变，也为 modified-at 提供同一窄复用点。

不公开一次性 metadata object，也不在 `write()` 中读取时钟。Created-at 是稳定、高层、可独立拥有的 live property；自动保存审计策略必须作为后续独立设计处理。

## 公开 API

`CreatePresentationOptions` 新增：

```ts
readonly createdAt?: string;
```

`PresentationModel` / `PptxDocument` 新增：

```ts
get createdAt(): string | undefined;
set createdAt(value: string | undefined);
```

`createdAt` 只接受严格 W3CDTF date-time string 或 `undefined`。API不接受 `Date`、number、bigint、boolean、array、object或symbol，不 trim、不调用 `toString()`、不改变时区，也不把 file mtime、当前时间、modified、revision或 lastModifiedBy 当作 fallback。

显式 `undefined` 清除 direct created property。空字符串不是合法时间，必须在 mutation 前抛 `TypeError`。Getter返回 exact lexical string，不 canonicalize fractional seconds、UTC offset或大小写。

## W3CDTF 子集

原生可创建和编辑的 lexical grammar 是：

```text
YYYY-MM-DDTHH:mm:ss[.fraction](Z|+HH:mm|-HH:mm)
```

约束如下：

- year 是 `0001..9999` 的四位十进制数；month/day必须是有效 Gregorian calendar date，含 leap-year规则；
- hour `00..23`，minute/second `00..59`；不创建 `24:00:00` 或 leap second；
- fractional second可省略；存在时必须含一个或多个 ASCII digits并原样保留；
- timezone必须存在；`Z` 必须大写，offset hour 是 `00..14`，且 `14` 只允许 minute `00`；
- 不接受前后空白、date-only、缺少秒、缺少 timezone、Unicode digits或其他 ISO/RFC 宽松形式。

这个子集覆盖 PptxGenJS 4.0.1 与 PowerPoint 常规输出，并避免由 JavaScript `Date.parse()` 引入实现相关的宽松解析。已有 direct created若不满足该 grammar，getter返回 `undefined`；合法 setter可修复其 simple text lexical state。

## OOXML ownership

目标状态是 package-root core-properties relationship 指向的合法 core-properties part 中唯一 direct：

```xml
<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>
```

Element expanded name必须是：

```text
{http://purl.org/dc/terms/}created
```

Type attribute expanded name和 QName value必须分别是：

```text
{http://www.w3.org/2001/XMLSchema-instance}type
{http://purl.org/dc/terms/}W3CDTF
```

实现不得依赖 `/docProps/core.xml`、`dcterms` / `xsi` lexical prefix、child顺序或 root-level namespace declaration位置。合法 alternate part URI、alternate prefixes和property-local declarations都必须支持。

同 local name但 wrong element namespace、descendant-only created、external/dangling relationship、wrong content type、wrong root、多个 core relationships或多个 direct namespace-correct created不能伪造 snapshot。多个 expanded-name-correct `xsi:type` attributes、element children或CDATA属于 unsafe state；mutation前拒绝而不是删除或猜测。

Unknown attributes、comments、whitespace、wrong-namespace siblings和其他 core children不属于 created-at，必须原样保留。

## Generic typed-property extension

`CoreTextPropertyDescriptor` 增加可选 internal-only type contract：

```ts
export interface CoreQualifiedTypeDescriptor {
  readonly attributeNamespace: string;
  readonly attributePreferredPrefix: string;
  readonly valueNamespace: string;
  readonly valuePreferredPrefix: string;
  readonly valueLocalName: string;
}

export interface CoreTextPropertyDescriptor {
  readonly label: string;
  readonly localName: string;
  readonly namespace: string;
  readonly preferredPrefix: string;
  readonly qualifiedType?: CoreQualifiedTypeDescriptor;
}
```

Created-at descriptor使用 XSI namespace attribute和 DCTERMS `W3CDTF` QName value。未提供 `qualifiedType` 的现有 metadata descriptors经过同一 tests后必须保持 byte-for-byte行为。

读取 typed property时，generic helper先完成既有 relationship、part、root、direct unique element和simple text检查，再按 expanded names解析唯一 type attribute，并按该 attribute所在 element的 in-scope namespace bindings解析 QName value。缺失、错误、unbound prefix或歧义 type都返回 `undefined`且不 mutation。

非 `undefined` mutation必须让最终 element具有一个正确的 type attribute：

- 已正确时保持 exact bytes；
- 唯一正确-expanded-name attribute但 value错误时，只替换 value；
- 缺失时复用 in-scope XSI prefix，或在property上增加一次 canonical `xmlns:xsi`与 `xsi:type`；
- QName value优先复用一个 in-scope non-empty DCTERMS prefix；不存在时在property上增加一次 canonical `xmlns:dcterms`；
- wrong-namespace local `type` 是opaque attribute，保留并另加正确 attribute；
- self-closing property扩展时保留原 attributes并一次性加入所需 bindings/type；
- duplicate expanded type attributes拒绝，避免有歧义地重写。

Missing core-properties part的最小 canonical output只声明一次 `cp`、`dcterms` 和 `xsi`，只写 direct created，不合成 creator、lastModifiedBy、revision、modified或timestamps policy。向已有 part插入时只在缺失处补 namespace declaration，不重排 root attributes或children。

## Read、edit与clear生命周期

Getter是 direct-state snapshot。合法 typed simple text且 lexical-valid时返回 exact string；缺失、不安全、namespace/type错误或 lexical-invalid时返回 `undefined`。读取不得产生 part、relationship、content-type或mutation journal变化。

Setter在 `opcPackage.transaction()` 中完成 normalization、结构检查、typed-property mutation和write-back：

- same lexical value且type状态已正确是 exact bytes/journal no-op；
- same text但缺失/错误的唯一 type状态会只修复type；
- valid replacement可修复 simple lexical-invalid text并确保type；
- `undefined` 删除唯一 direct created；absent clear是 exact no-op；
- complex content、duplicate direct property、duplicate expanded type或unsafe package ownership在任何 mutation前拒绝；
- outer transaction失败时，existing-part edit或新part/relationship/content-type创建全部回滚。

Created-at mutation不修改 modified、creator、lastModifiedBy、revision、title、subject、company、unknown children、unrelated parts、relationships或slide content，也不自动读取当前时间。

## Native creation

`PptxDocument.create({ createdAt })`先建立现有 canonical package，只在值非 `undefined` 时通过 public setter应用。因此：

- omitted和runtime `undefined`保持 canonical bytes且 `createdAt === undefined`；
- valid explicit值立即出现在getter和typed OOXML中；
- invalid input在返回 document前抛出且不泄漏半成品；
- author、lastModifiedBy、revision、subject、title、company和createdAt保持独立ownership，option顺序不影响最终状态；
- 六种presentation formats均使用同一core-property contract。

普通 `write()` 不刷新 createdAt。这个确定性 native default是有意的 existing-deck/editor contract；PptxGenJS的write-time clock行为只作为可导入 final state验证，不复制为隐藏副作用。

## PptxGenJS 4.0.1 conformance

PptxGenJS 4.0.1没有公开created property。它在每次公开 `write()` 生成core.xml时调用当前时钟，写入UTC、秒级、uppercase `Z`的created和modified：

```xml
<dcterms:created xsi:type="dcterms:W3CDTF">YYYY-MM-DDTHH:mm:ssZ</dcterms:created>
```

Adapter tests只通过public constructor、`addSlide()`和`write()`取得真实输出，不读取private fields或mock内部生成器。要求：

- imported `createdAt`符合strict grammar并等于raw core.xml exact text；
- namespace/type状态正确，write/reopen保持同一createdAt；
- native explicit同值可创建相同typed final state；
- native edits不改变PptxGenJS输出中的modified或任何其他part；
- 不断言两次独立PptxGenJS `write()` 时间相同，因为其公开行为会重新读取时钟。

## 文件与职责

- `packages/model/src/presentation-core-properties.internal.ts`：可选qualified type的expanded-name/QName读取、最小渲染与repair。
- `packages/model/src/presentation-created-at.internal.ts`：created-at descriptor、strict lexical/semantic normalization和thin wrappers。
- `packages/model/src/presentation-created-at.internal.test.ts`：typed core-property、calendar grammar、OPC lifecycle和rollback单元测试。
- `packages/model/src/presentation.ts`：live getter/setter与transaction boundary。
- `packages/model/src/model.test.ts`：public model read/edit/clear/reopen与preservation。
- `packages/sdk/src/create.ts`、`packages/sdk/src/index.ts`：create option和显式setter应用。
- `packages/sdk/src/index.test.ts`：Node API、六formats、invalid zero-mutation、duplicate/rollback/write/reopen。
- `packages/pptxgenjs-adapter/src/index.test.ts`：PptxGenJS public-output import、typed state、native parity和reopen。
- `scripts/smoke-npm-package.mjs`：packed Node/browser/declaration smoke。
- `CHANGELOG.md`、`docs/api/README.md`、`docs/compatibility/pptxgenjs-baseline.md`、`packages/pptx/README.md`：contract、差异和剩余缺口。

## 验证与提交边界

实现按可独立review的小项提交并在每次通过focused tests、`git diff --check`和static review后通过SSH 443 push：

1. design；
2. implementation plan；
3. generic typed helper + created-at internal codec/tests；
4. model accessor/tests；
5. create API/tests；
6. PptxGenJS public-output conformance；
7. packed Node/browser/types smoke；
8. docs；
9. full gates与real-deck QA不创建空commit。

真实QA至少生成native omitted/source/edited/cleared/reopened和PptxGenJS output decks。所有文件通过PowerPoint 2010 validation；exact core read确认typed value、offset/fraction preservation和clear；source→edited只允许 `/docProps/core.xml` 变化，edited→reopened零part变化；LibreOffice与PNG hash证明metadata不改变slide rendering。

`.pnpm-store/` 是用户拥有的未跟踪目录，任何阶段都不得修改、删除、stage或commit。
