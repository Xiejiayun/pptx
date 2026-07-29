# Presentation Subject Metadata Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 presentation core properties 增加原生 subject 创建、读取、编辑和清除能力，使 `PptxDocument.create({ subject })` 与 `document.subject` 覆盖从零创建和已有 PPTX 编辑，并对等 PptxGenJS 4.0.1 顶层 `pptx.subject` 的有效字符串输出。

本小项只拥有 OPC core-properties part 中 direct Dublin Core subject。它不同时实现 revision、last-modified-by、created/modified time、keywords、category、description、content status 或 custom properties，也不把 title、slide title、file name、template 或 section name解释为 subject。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按公开 API 一致性、preserve-first 编辑、OPC 生命周期、namespace correctness 和可独立验收定稿。

## 方案选择

考虑三种方案：

1. 在 `PresentationModel` 或新的 subject helper 中复制 core-properties relationship、part、namespace 和 source-span 逻辑；不采用。它会复制已经由 title/author 验证过的安全边界。
2. 复用 `presentation-core-properties.internal.ts`，仅以 subject descriptor 和 subject-specific validation 增加薄包装；采用。它保持 public scope 单一，并让 title、author、subject 使用同一套无损生命周期。
3. 一次公开包含全部 metadata 的对象；不采用。Revision、timestamps、extended/custom properties具有不同值类型、namespace 和 ownership，合并会放大单个审查单元。

Subject 是 presentation 的稳定高层属性，不要求调用方使用 raw part API。通用 core-properties helper 不新增 public export。

## 公共 API

```ts
export interface CreatePresentationOptions {
  readonly author?: string;
  readonly company?: string;
  readonly format?: PresentationFormat;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly subject?: string;
  readonly title?: string;
}

export class PresentationModel {
  get subject(): string | undefined;
  set subject(value: string | undefined);
}
```

`CreatePresentationOptions.subject` 只接受 string。Omitted 与 runtime `undefined` 保留原生 canonical core-properties bytes，zero-input native document 的 immediate subject 是 `undefined`。显式空字符串写 direct empty subject 并由 getter 返回 `''`。`document.subject = undefined` 删除 direct subject，`document.subject = ''` 保留 direct empty subject。Null、number、boolean、array、object、symbol 和非法 XML 控制字符都不被接受；API 不 trim、不 case-fold、不调用 `toString()`。

Getter 是 direct-state snapshot，不回退到 title、first slide title、file name、template、master、section 或 PowerPoint UI 默认值。不存在、安全上歧义或不支持的 direct state 返回 `undefined`，且不产生 mutation。

## Core-properties internal 边界

现有 `presentation-core-properties.internal.ts` 继续拥有 package-root exact core-properties relationship、content type、root namespace、direct simple-text field discovery、minimal part creation、prefix reuse、source-span replace/insert/clear、same-value no-op 和 unsafe ownership errors。

新增 subject wrapper 只声明：

```ts
const SUBJECT_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'subject',
  localName: 'subject',
  namespace: 'http://purl.org/dc/elements/1.1/',
  preferredPrefix: 'dc',
};
```

Wrapper 在进入 generic helper 前执行 subject-specific strict normalization，并暴露 internal `readPresentationSubject()` 与 `replacePresentationSubject()`。Title 和 author wrapper 及其现有测试不得改变行为。任何 metadata internal helper 都不得从 `packages/model/src/index.ts` 或聚合 npm 包导出。

## OPC 与 OOXML ownership

Core properties 通过 package root 的 exact relationship type 定位：

```text
http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties
```

目标 part 的 content type 必须是：

```text
application/vnd.openxmlformats-package.core-properties+xml
```

Subject 只拥有 core-properties root 下 direct Dublin Core namespace child：

```xml
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Quarterly Review</dc:title>
  <dc:subject>Revenue &amp; Forecast</dc:subject>
  <dc:creator>Alice</dc:creator>
  <cp:lastModifiedBy>Editor</cp:lastModifiedBy>
</cp:coreProperties>
```

实现不得假定 part URI、`cp`/`dc` lexical prefix 或 child 顺序。Wrong-prefix-but-correct-URI 是合法输入；同 local name 但 wrong namespace、descendant-only subject、external/dangling relationship、wrong content type、wrong root、多个 core-properties relationships或多个 direct Dublin Core subjects都不能伪造 snapshot。插入时优先复用 root in-scope 的 Dublin Core prefix；没有可复用 binding 时，新 `dc:subject` 在自身声明 canonical `xmlns:dc`，不重写 root namespaces。

Getter 只接受唯一 direct subject，其内容必须是 simple text state；element children或CDATA等不安全结构返回 `undefined`。XML entities由 lossless parser 解码，empty element返回空字符串。读取和 subject mutation不改写 title、creator、lastModifiedBy、revision、timestamps、comments、whitespace、unknown children、unrelated parts或relationships。

## 创建与 part 生命周期

原生 zero-input package 已包含 canonical `/docProps/core.xml` 与 root relationship，但不含 subject。`PptxDocument.create({ subject })` 先创建 package，再通过同一个 public setter 写 direct subject，因此创建与 existing-deck 编辑共享归一化、escaping 和 transaction 语义。Omitted subject 不触碰 canonical core-properties bytes；explicit subject 只更新 core-properties part。

对没有 core-properties relationship 的合法 existing deck：

- 读取返回 `undefined`；clear 是 exact no-op。
- 设置 string 创建 minimal core-properties part，优先使用空闲的 `/docProps/core.xml`，否则分配 `/docProps/core1.xml`、`core2.xml` 等安全 URI。
- 创建对应 content-type override和唯一 root relationship，target使用 root-relative package path。
- 新 part只含 canonical namespace declarations与direct subject，不合成 title、creator、lastModifiedBy、revision或timestamps。

对唯一 relationship 指向的合法 part，setter 只 source-span patch direct subject：same-value 是 exact byte/journal no-op；string replace或insert direct element；`undefined` 删除该 element但保留 part、relationship和所有其他properties。Self-closing root/subject必须安全展开。不能安全确定 ownership 的 malformed/ambiguous package在任何 mutation 前抛 `PackageError` 或 `ModelParseError`，并完整 rollback。

## PptxGenJS 4.0.1 基线

锁定版本的 public runtime 与真实 `write()` output 必须确认：

- 未赋值时 public `pptx.subject` 是 `PptxGenJS Presentation`，并写 direct `<dc:subject>PptxGenJS Presentation</dc:subject>`。
- `pptx.subject = 'Revenue & <Forecast>'` 正确 XML-escape 并写同一 direct field。
- `pptx.subject = ''` 写 direct empty `<dc:subject></dc:subject>`。
- PptxGenJS 没有 existing-deck metadata editor，也不暴露 direct absence。

Native explicit strings与PptxGenJS final subject snapshot对等；adapter导入default/custom/empty outputs后必须分别读到对应字符串。Native omitted保持当前无 direct subject 的 canonical output，而不注入PptxGenJS品牌默认值。Native `undefined` clear与missing-part creation是lossless editing extensions。

## 文件边界

`packages/model/src/presentation-subject.internal.ts` 负责subject descriptor、strict normalization和generic helper委托。`packages/model/src/presentation.ts` 只公开subject getter/setter并包裹package transaction。`packages/sdk/src/create.ts` 扩展 `CreatePresentationOptions`，`packages/sdk/src/index.ts` 在构造后通过同一个 public setter应用explicit subject。

`packages/model/src/presentation-subject.internal.test.ts`、公共model tests、SDK tests、PptxGenJS adapter tests和actual-tarball smoke分别证明internal、public、creation、conformance和发布表面。文档只宣称subject能力，不顺带宣称其他metadata完成。

每个实现阶段只修改该阶段所需文件。不得修改、删除、stage或提交 `.pnpm-store/`。

## 测试与发布门禁

1. Internal/model tests覆盖absent、empty、escaped Unicode/string、same-value exact no-op、replace、clear、self-closing root/subject、alternate prefixes、descendant/wrong namespace、duplicate direct subjects、missing relationship part creation、occupied canonical URI fallback、dangling/external/wrong-content-type/wrong-root/multiple relationships、rollback和其他core properties保持。
2. Invalid inputs覆盖null、number、boolean、array、object、symbol、XML controls；所有失败保持part bytes、relationship XML、content types、journal、slides和model identity不变。
3. SDK tests覆盖native omitted/custom/empty/invalid、与title/author/company共存、immediate getter、edit/clear、transaction rollback、write/reopen和全部六种presentation formats。
4. Adapter conformance只使用PptxGenJS public `subject`与`write()` output，覆盖default/custom/empty import、native omitted差异、exact escaping和write/reopen；不读取私有字段。
5. Packed Node/browser/declaration/CLI smoke覆盖create、read、edit、clear和 `string | undefined` 类型。
6. Changelog、API README、compatibility baseline和package README新增subject行，明确native omitted default、PptxGenJS品牌默认、direct empty与clear差异，并保留其他metadata缺口。
7. Full TypeScript、Vitest、1000-part performance、actual tarball smoke和PowerPoint 2010 profile validation全部通过。
8. 真实native/PptxGenJS decks由 `pptx-inspect part read` 确认exact subject和escaping；native source→edited package diff只允许core-properties part变化，edited→reopened零part变化。LibreOffice导出无修复；metadata无可视内容，因此视觉验收只确认页面几何与内容没有副作用。
