# Presentation Revision Metadata Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 presentation core properties 增加原生 revision 创建、读取、编辑和清除能力，使 `PptxDocument.create({ revision })` 与 `document.revision` 覆盖从零创建和已有 PPTX 编辑，并对等 PptxGenJS 4.0.1 顶层 `pptx.revision` 的安全 whole-number string 输出。

本小项只拥有 OPC core-properties part 中 direct core-properties namespace `revision`。它不同时实现 lastModifiedBy、created/modified timestamps、keywords、category、description、content status、custom properties 或自动保存次数。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按公开 API 一致性、PowerPoint-safe 输出、preserve-first 编辑、OPC 生命周期、namespace correctness 和可独立验收定稿。

## 方案选择

考虑三种方案：

1. 复用 `presentation-core-properties.internal.ts`，增加 revision-specific wrapper，并只修正 generic minimal-part creation 对 root core namespace descriptor 的声明；采用。它保留已由 title、author、subject 验证的 relationship、part、namespace 和 lossless patch 边界，新增逻辑只负责 revision lexical contract。
2. 在新的 revision helper 中复制完整 core-properties lifecycle；不采用。它会复制约两百行 ownership 和 source-span 逻辑，并使相同 malformed package 在不同 metadata 字段上出现分叉行为。
3. 把 generic helper 重构成支持任意 parser、validator、schema order 和 typed value 的通用 metadata framework；不采用。当前只有 revision 需要一个读取过滤器，泛化回调和排序策略会扩大本小项审查面，timestamps 与 custom properties 仍需不同 OOXML 类型和 lifecycle。

Revision 是 presentation 的稳定高层属性，不要求调用方使用 raw part API。通用 core-properties helper 和 revision wrapper 均不新增 public internal export。

## 公共 API 与值契约

```ts
export interface CreatePresentationOptions {
  readonly author?: string;
  readonly company?: string;
  readonly format?: PresentationFormat;
  readonly revision?: string;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly subject?: string;
  readonly title?: string;
}

export class PresentationModel {
  get revision(): string | undefined;
  set revision(value: string | undefined);
}
```

Revision 保持 lexical string，而不是 JavaScript number。这样不会丢失合法的 leading zeros，也不会受 `Number.MAX_SAFE_INTEGER`、指数形式、locale 或浮点舍入影响。接受值必须完全匹配：

```text
^[0-9]+$
```

因此 `'0'`、`'1'`、`'7'`、`'01'` 和任意非空 ASCII digit string 合法；`''`、空白、正负号、小数、指数、Unicode digit、number、bigint、boolean、null、array、object 和 symbol 均在 mutation 前抛 `TypeError`。API 不 trim、不 normalize leading zeros、不调用 `toString()`，也不把 revision 当算术计数器自动递增。

原生 canonical package 已包含 direct `<cp:revision>1</cp:revision>`，所以 `PptxDocument.create().revision` 是 `'1'`，与 PptxGenJS 4.0.1 默认值一致。Omitted 或 runtime `undefined` 的 create option 保留该 canonical bytes。显式 `document.revision = undefined` 删除 direct revision；这是一项 existing-deck editing extension。创建选项不能表达 clear，因为 omitted/undefined 按所有 create options 的既有规则表示“不应用 setter”。

Getter 是 direct lexical-state snapshot，不把缺失或 invalid revision 回退为 `'1'`，也不从 application properties、文件版本、保存时间、modifiedBy、custom properties 或 PowerPoint UI 推断。唯一 direct simple-text revision 只有在完全匹配 digit grammar 时返回原 lexical string；缺失、不安全、歧义或 lexical-invalid state 返回 `undefined`，读取不产生 mutation。

## Core-properties internal 边界

新增 `presentation-revision.internal.ts`，声明：

```ts
const REVISION_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'revision',
  localName: 'revision',
  namespace: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  preferredPrefix: 'cp',
};
```

它暴露 internal：

```ts
export function readPresentationRevision(pkg: OpcPackage): string | undefined;

export function replacePresentationRevision(
  pkg: OpcPackage,
  value: string | undefined,
): void;
```

读取先调用 `readCoreTextProperty()`，再过滤 digit grammar；写入先执行 revision-specific strict normalization，再调用 `replaceCoreTextProperty()`。有效同值仍由 generic helper 保证 exact no-op。Malformed nested/CDATA property 仍按 helper 的 unsafe mutation规则拒绝，而 simple lexical-invalid property可由合法 digit string整体替换或由 `undefined` 清除。

`presentation-core-properties.internal.ts` 只做一项必要修正：当 descriptor 的 preferred prefix/namespace 已等于 canonical root `cp` binding 时，minimal part creation 不再生成第二个相同 `xmlns:cp` attribute。Revision missing-part creation必须生成：

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
  <cp:revision>7</cp:revision>
</cp:coreProperties>
```

Title、author 和 subject 使用 Dublin Core descriptor，继续得到既有 `xmlns:cp` 与 `xmlns:dc` declarations，bytes 和行为不得改变。

## OPC 与 OOXML ownership

Core properties 通过 package root 的 exact relationship type 定位：

```text
http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties
```

目标 part content type必须是：

```text
application/vnd.openxmlformats-package.core-properties+xml
```

Revision 只拥有 core-properties root 下 direct core-properties namespace child：

```xml
<c:coreProperties
  xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:d="http://purl.org/dc/elements/1.1/">
  <d:title>Quarterly Review</d:title>
  <c:lastModifiedBy>Editor</c:lastModifiedBy>
  <c:revision>007</c:revision>
</c:coreProperties>
```

实现不得假定 part URI、`cp` lexical prefix 或 child 顺序。Wrong-prefix-but-correct-URI 是合法输入；同 local name 但 wrong namespace、descendant-only revision、external/dangling relationship、wrong content type、wrong root、多个 core-properties relationships或多个 direct revision都不能伪造 snapshot。

插入时复用 root in-scope 的 core-properties prefix；canonical root 使用 `cp:revision`，alternate `c` root 使用 `c:revision`。Getter只接受唯一 direct simple-text revision；XML entities虽然会被 parser 解码，但 digit grammar意味着任何非 digit entity结果均不支持。Empty/self-closing revision是 lexical-invalid state，返回 `undefined`，但可被合法值替换或清除。

读取和 revision mutation不改写 title、subject、creator、lastModifiedBy、timestamps、comments、whitespace、unknown children、unrelated parts或relationships。Generic helper继续承担 source-span replace/insert/clear、same-value no-op、self-closing expansion和 unsafe ownership errors。

## 创建与 part 生命周期

`PptxDocument.create({ revision })` 先建立 canonical package，再仅在 revision 非 `undefined` 时通过 public setter应用值。因此：

- omitted/runtime `undefined` 保留 direct canonical revision `'1'` 及全部 bytes；
- explicit digit string替换 canonical `'1'`，但不改其他 core properties；
- invalid explicit value在 setter 进入 package mutation前拒绝，整个 create调用抛错且不暴露半成品 document。

对没有 core-properties relationship 的合法 existing deck：

- 读取返回 `undefined`；clear 是 exact no-op；
- 设置 digit string创建 minimal core-properties part，优先使用空闲的 `/docProps/core.xml`，否则分配 `/docProps/core1.xml`、`core2.xml` 等安全 URI；
- 创建对应 content-type override和唯一 root relationship；
- 新 part只含 canonical root namespace与direct revision，不合成 title、creator、lastModifiedBy或timestamps。

对唯一 relationship指向的合法 part，setter只 patch direct revision：same lexical value是 exact bytes/journal no-op；digit string replace或insert；`undefined` 删除该 element但保留 part、relationship和所有其他 properties。不能安全确定 ownership 的 malformed/ambiguous package在任何 mutation前抛 `PackageError` 或 `ModelParseError`，并完整 rollback。

## PptxGenJS 4.0.1 基线

锁定版本的 public runtime与真实 `write()` output已确认：

- 未赋值时 public `pptx.revision` 是 `'1'`，并写 direct `<cp:revision>1</cp:revision>`；
- digit strings `'0'`、`'7'`、`'01'` 原样写入 direct revision；
- public TypeScript surface声明 revision为 string，并警告该值必须是没有 `.` 或 `,` 的 whole number，否则 PowerPoint 打开时会报错；
- runtime没有执行该约束，empty、negative、decimal、alphabetic string、number和null都可能被原样或字符串化写出。

Native只对等 PptxGenJS文档契约内的 non-empty ASCII digit strings。PptxGenJS runtime产生的 invalid lexical states由 adapter保留 bytes，但 strict native getter返回 `undefined`；native创建和编辑在 mutation前拒绝，不复制可能导致 PowerPoint错误的 runtime缺陷。Native额外支持 clear、missing-part creation、alternate URI/prefix和 lossless existing-deck editing。

## 文件边界

`packages/model/src/presentation-revision.internal.ts` 负责 strict lexical snapshot和 normalization，并委托 generic core-property lifecycle。`packages/model/src/presentation-core-properties.internal.ts` 只修正 same-namespace minimal part declaration。`packages/model/src/presentation.ts` 公开 revision getter/setter并包裹 transaction。`packages/sdk/src/create.ts` 扩展 option；`packages/sdk/src/index.ts` 只在显式值存在时通过 public setter应用。

后续 adapter、packed smoke和文档阶段只触及其各自表面。不得重构其他 metadata、自动递增 revision、修改 timestamps、同步 lastModifiedBy、删除 invalid existing XML、公开 internal helper，或修改、删除、stage、提交 `.pnpm-store/`。

## 测试与发布门禁

1. Revision internal tests覆盖 absent、canonical/alternate prefix、`0`、leading zeros、long digits、same-value no-op、replace、insert、clear、self-closing root/revision、missing relationship part creation、occupied URI fallback、wrong namespace、descendant、duplicate、invalid lexical state、nested/CDATA state、dangling/external/wrong-content-type/wrong-root/multiple relationships、rollback和其他 core properties保持。
2. Invalid inputs覆盖 empty、whitespace、sign、decimal、exponent、Unicode digits、null、number、bigint、boolean、array、object、symbol和XML control；失败保持part bytes、relationships、content types、journal、slides和model identity不变。
3. Title、author和subject internal regressions证明 generic helper minimal creation修正未改变其 bytes或 lifecycle。
4. Model tests覆盖 getter、digit edit、leading-zero preservation、clear、invalid existing snapshot、repair、rollback、alternate URI/prefix、missing-part creation、same-value/absent-clear no-op、unrelated part isolation和 write/reopen。
5. SDK tests覆盖 native default `'1'`、custom/leading-zero、invalid、immediate getter、edit/clear、rollback、write/reopen和全部六种 presentation formats。
6. Adapter conformance只使用 PptxGenJS public `revision` 与 `write()`，覆盖 default、`'0'`、custom、leading-zero、native strict invalid rejection和 write/reopen；不读取私有字段。
7. Packed Node/browser/declaration/CLI smoke覆盖 create、read、edit、clear和 `string | undefined` 类型。
8. Changelog、API README、compatibility baseline和 package README新增 revision 行，明确 canonical/default `'1'`、digit grammar、leading-zero preservation、clear语义及 PptxGenJS invalid runtime边界。
9. Full TypeScript、Vitest、1000-part performance、actual tarball smoke、PowerPoint 2010 profile validation全部通过。
10. 真实 native/PptxGenJS decks由 `pptx-inspect part read` 确认 exact revision；native source→edit package diff只允许 core-properties part变化，edit→reopen零 diff。LibreOffice导出无修复；metadata无可视内容，因此视觉验收确认页面几何和内容无副作用。
