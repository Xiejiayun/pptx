# Presentation Modified-At Metadata Design

## 目标与边界

为 presentation core properties 增加原生修改时间的创建、读取、编辑和清除能力，使 `PptxDocument.create({ modifiedAt })` 与 `document.modifiedAt` 同时覆盖从零创建和已有 PPTX 编辑，并能精确导入 PptxGenJS 4.0.1 每次公开 `write()` 生成的 `dcterms:modified`。

本小项只拥有 OPC core-properties part 中 direct Dublin Core Terms `modified` element 及其必需的 `xsi:type="dcterms:W3CDTF"` 类型状态。它不刷新 `createdAt`、不在保存时读取时钟、不递增 revision、不同步 lastModifiedBy，也不引入文件系统时间、custom properties 或审计策略。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按 PowerPoint-compatible typed property、strict deterministic API、preserve-first 编辑和与 created-at 一致的 namespace/QName contract 定稿。

## 方案选择

考虑三种实现：

1. 完整复制 created-at 的正则、Gregorian calendar 和 timezone 校验到 modified-at wrapper；不采用。两个公开字段会在错误边界或未来修复上漂移。
2. 在 `write()` 中自动把 modified-at 设置为当前 UTC 时间；不采用。它会让 identical input 产生不同 bytes，破坏 existing-deck 无关写入的 exact preservation，也无法由调用方可靠控制或测试。
3. 抽取一个窄的 internal W3CDTF timestamp helper，让 created-at 与 modified-at 各自保留独立 descriptor、wrapper 和 public property；采用。共享层只拥有完全相同的 lexical/calendar/type contract，不拥有字段选择、OPC relationship、mutation policy 或 save-time clock。

不公开 metadata bag、`Date` 输入、clock provider 或自动审计选项。若将来需要 save-time metadata policy，它必须作为显式、独立的小项设计，不能改变本字段的确定性 setter 语义。

## 公开 API

`CreatePresentationOptions` 新增：

```ts
readonly modifiedAt?: string;
```

`PresentationModel` / `PptxDocument` 新增：

```ts
get modifiedAt(): string | undefined;
set modifiedAt(value: string | undefined);
```

`modifiedAt` 只接受严格 W3CDTF date-time string 或 `undefined`。API 不接受 `Date`、number、bigint、boolean、array、object 或 symbol，不 trim、不 coercion、不改变时区，也不把 created、revision、lastModifiedBy、file mtime 或当前时间当 fallback。

显式 `undefined` 清除 direct modified property。空字符串不是合法时间，必须在 mutation 前抛 `TypeError`。Getter 返回 exact lexical string，不 canonicalize fractional seconds、UTC offset 或大小写。

## 共享 W3CDTF contract

`createdAt` 与 `modifiedAt` 共用以下 lexical grammar：

```text
YYYY-MM-DDTHH:mm:ss[.fraction](Z|+HH:mm|-HH:mm)
```

约束如下：

- year 是 `0001..9999` 的四位 ASCII digits；month/day 必须组成有效 Gregorian calendar date，含 leap-year 规则；
- hour `00..23`，minute/second `00..59`；不创建 `24:00:00` 或 leap second；
- fractional second 可省略；存在时必须含一个或多个 ASCII digits 并原样保留；
- timezone 必须存在；`Z` 必须大写，offset hour 是 `00..14`，且 `14` 只允许 minute `00`；
- 不接受前后空白、date-only、缺秒、缺 timezone、Unicode digits 或 JavaScript 宽松日期形式。

新增 internal-only `presentation-timestamp.internal.ts`，提供：

```ts
export const W3CDTF_QUALIFIED_TYPE: CoreQualifiedTypeDescriptor;

export function isPresentationTimestamp(value: string): boolean;

export function normalizePresentationTimestamp(
  value: unknown,
  propertyName: 'createdAt' | 'modifiedAt',
): string | undefined;
```

该 helper 使用 explicit regex、numeric range 和 Gregorian leap-year 计算，禁止 `Date.parse()`。`normalizePresentationTimestamp()` 只负责 shared input contract和字段准确错误消息；它不读写 package。Created-at wrapper改为使用该 helper，现有 public behavior、accepted lexical values、error class、bytes和 tests必须保持不变。

## OOXML ownership

目标状态是 package-root core-properties relationship 指向的合法 core-properties part 中唯一 direct：

```xml
<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03.456+08:00</dcterms:modified>
```

Element expanded name、type attribute expanded name和 QName value必须分别是：

```text
{http://purl.org/dc/terms/}modified
{http://www.w3.org/2001/XMLSchema-instance}type
{http://purl.org/dc/terms/}W3CDTF
```

Modified-at descriptor复用 generic `CoreTextPropertyDescriptor.qualifiedType` 和共享 `W3CDTF_QUALIFIED_TYPE`。实现不得依赖 `/docProps/core.xml`、lexical prefix、child顺序或 namespace declaration scope。合法 alternate part URI、alternate element/XSI/QName prefixes和property-local declarations都必须支持。

Wrong element namespace、descendant-only modified、external/dangling relationship、wrong content type、wrong root、多个 core relationships、多个 direct namespace-correct modified、duplicate expanded-name-correct type attributes、element children和CDATA都不能伪造 snapshot。Unsafe 或 ambiguous state在 mutation 前拒绝，而不是猜测或删除。

Unknown attributes、comments、whitespace、wrong-namespace siblings、created和其他 core children不属于 modified-at，必须原样保留。

## Read、edit与clear生命周期

Getter 是 direct-state snapshot。唯一合法 typed simple text且 lexical-valid时返回 exact string；缺失、unsafe、namespace/type错误或 lexical-invalid时返回 `undefined`。读取不得产生 part、relationship、content-type或mutation journal变化。

Setter在 `opcPackage.transaction()` 中完成 normalization、结构检查、generic typed-property mutation和write-back：

- same lexical value且type状态正确是 exact bytes/journal no-op；
- same text但缺失/错误的唯一 type状态只修复type；
- valid replacement可修复 simple lexical-invalid text并确保type；
- `undefined` 删除唯一 direct modified；absent clear是 exact no-op；
- complex content、duplicate direct property、duplicate expanded type或unsafe package ownership在任何 mutation前拒绝；
- outer transaction失败时，existing-part edit或新part/relationship/content-type创建全部回滚。

Modified-at mutation不修改 created、creator、lastModifiedBy、revision、title、subject、company、unknown children、unrelated parts、relationships或slide content。`write()` 对 modified-at 是纯序列化，不读取当前时钟，也不根据其他 metadata 自动合成值。

## Native creation

`PptxDocument.create({ modifiedAt })` 先建立现有 canonical package，只在值非 `undefined` 时通过 public setter应用。因此：

- omitted和runtime `undefined`保持 canonical bytes且 `modifiedAt === undefined`；
- valid explicit值立即出现在getter和typed OOXML中；
- invalid input在返回 document前抛出且不泄漏半成品；
- createdAt、author、lastModifiedBy、revision、subject、title、company和modifiedAt保持独立ownership，create option应用顺序不改变字段内容；
- 六种presentation formats使用同一core-property contract。

普通 `write()` 不刷新 modifiedAt。这个 deterministic default优先服务本库从零创建和 existing-deck editor的可重复输出；调用方需要当前时间时必须显式提供经过格式化的 string。

## PptxGenJS 4.0.1 conformance

PptxGenJS 4.0.1没有公开 modified timestamp property。每次 public `write()` 会读取当前时钟并写入UTC、秒级、uppercase `Z`的 created和modified typed properties：

```xml
<dcterms:modified xsi:type="dcterms:W3CDTF">YYYY-MM-DDTHH:mm:ssZ</dcterms:modified>
```

Adapter tests只通过public constructor、`addSlide()`和`write()`取得真实输出，不读取 private fields、不 mock内部生成器。要求：

- imported `modifiedAt`符合shared strict grammar并等于同一 raw core.xml 的 exact modified text；
- namespace/type状态正确，native write/reopen保持同一 modifiedAt；
- native explicit同值可创建相同 typed final state；
- native modified edit不改变同一 PptxGenJS output中的 created或其他part；
- 不断言不同 public `write()` 调用得到相同 timestamp，因为 PptxGenJS会重新读取时钟。

本库只对等 PptxGenJS最终公开输出的可导入状态，不复制其隐藏 clock side effect。

## 文件与职责

- `packages/model/src/presentation-timestamp.internal.ts`：共享 strict W3CDTF qualified-type、lexical/calendar/timezone predicate和normalizer。
- `packages/model/src/presentation-created-at.internal.ts`：改为复用共享 helper，保持 created descriptor和ownership独立。
- `packages/model/src/presentation-modified-at.internal.ts`：modified descriptor与thin read/replace wrappers。
- `packages/model/src/presentation-modified-at.internal.test.ts`：shared extraction regression、typed modified read/edit/repair/clear、OPC lifecycle和rollback。
- `packages/model/src/presentation.ts`：live getter/setter与transaction boundary。
- `packages/model/src/model.test.ts`：public model read/edit/clear/reopen与created isolation。
- `packages/sdk/src/create.ts`、`packages/sdk/src/index.ts`：create option和显式setter应用。
- `packages/sdk/src/index.test.ts`：Node API、六formats、invalid zero-mutation、independence、rollback/write/reopen。
- `packages/pptxgenjs-adapter/src/index.test.ts`：PptxGenJS public-output import、typed state、native parity和reopen。
- `scripts/smoke-npm-package.mjs`：packed Node/browser/declaration smoke。
- `CHANGELOG.md`、`docs/api/README.md`、`docs/compatibility/pptxgenjs-baseline.md`、`packages/pptx/README.md`：contract、PptxGenJS差异和剩余metadata缺口。

## 验证与提交边界

实现按可独立review的小项提交，并在每次通过 focused tests、`git diff --check`和static review后通过SSH 443 push：

1. design；
2. implementation plan；
3. shared timestamp helper + modified-at internal codec/tests；
4. model accessor/tests；
5. create API/tests；
6. PptxGenJS public-output conformance；
7. packed Node/browser/types smoke；
8. docs；
9. full gates与real-deck QA不创建空commit。

真实QA至少生成 native omitted/source/edited/cleared/reopened和PptxGenJS output decks。所有文件通过PowerPoint 2010 validation；exact core read确认typed modified、offset/fraction preservation、created isolation和clear；source→edited只允许 `/docProps/core.xml` 变化，edited→reopened零part变化；LibreOffice与PNG hash证明metadata不改变slide rendering。

`.pnpm-store/` 是用户拥有的未跟踪目录，任何阶段都不得修改、删除、stage或commit。
