# Presentation Last Modified By Metadata Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 presentation core properties 增加原生 lastModifiedBy 创建、读取、编辑和清除能力，使 `PptxDocument.create({ lastModifiedBy })` 与 `document.lastModifiedBy` 覆盖从零创建和已有 PPTX 编辑。它同时补全 PptxGenJS 4.0.1 公开 `author` 输出中的 creator/lastModifiedBy 镜像状态，但不把两者合并成一个原生 ownership。

本小项只拥有 OPC core-properties part 中 direct core-properties namespace `lastModifiedBy`。它不同时实现 created/modified timestamps、revision 自动递增、keywords、category、description、content status、custom properties或保存审计策略。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按独立 metadata ownership、preserve-first 编辑、OPC lifecycle、namespace correctness 和可独立验收定稿。

## 方案选择

考虑三种方案：

1. 复用 `presentation-core-properties.internal.ts`，以 lastModifiedBy descriptor 和 strict XML-safe string normalization 增加薄 wrapper；采用。Revision 已证明 generic helper可安全处理与 root 相同的 `cp` namespace，title/author/subject 已证明 simple-text lossless lifecycle。
2. 修改原生 `document.author`，每次自动镜像到 lastModifiedBy；不采用。现有设计明确让 creator 与 lastModifiedBy独立，镜像会破坏已有 deck中的审计字段，也会让无关 author edit改写第二个 property。
3. 一次公开完整 document-properties对象并自动维护 revision/timestamps；不采用。动态审计策略需要独立的时间、写出和 opt-in设计，不能由一个 string metadata小项隐式决定。

LastModifiedBy 是 presentation 的稳定高层属性。调用方不需要 raw part API，内部 wrapper不从 model聚合入口导出。

## 公共 API

```ts
export interface CreatePresentationOptions {
  readonly author?: string;
  readonly company?: string;
  readonly format?: PresentationFormat;
  readonly lastModifiedBy?: string;
  readonly revision?: string;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly subject?: string;
  readonly title?: string;
}

export class PresentationModel {
  get lastModifiedBy(): string | undefined;
  set lastModifiedBy(value: string | undefined);
}
```

`lastModifiedBy` 只接受 XML-safe string。API不 trim、不 case-fold、不调用 `toString()`。显式空字符串写 direct empty property并由 getter返回 `''`；assignment `undefined` 删除 direct property。Null、number、bigint、boolean、array、object、symbol和非法 XML控制字符在 mutation前抛 `TypeError`。

原生 canonical package已包含：

```xml
<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>
```

因此 `PptxDocument.create().lastModifiedBy` 是 `'@jiayunxie/pptx'`。Omitted 或 runtime `undefined` create option保持该 bytes。显式 custom/empty通过同一 live setter应用。Create options同时包含 author和lastModifiedBy时，两者分别写 creator与lastModifiedBy；setter应用顺序不改变结果，因为两个字段互不拥有。

Getter是 direct-state snapshot，不回退到 creator、operating-system user、company、application、modified timestamp或PowerPoint UI。缺失、安全上歧义或不支持的 direct state返回 `undefined`，读取不产生 mutation。

## Core-properties internal 边界

新增 `presentation-last-modified-by.internal.ts`：

```ts
const LAST_MODIFIED_BY_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'last modified by',
  localName: 'lastModifiedBy',
  namespace: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  preferredPrefix: 'cp',
};
```

内部接口：

```ts
export function readPresentationLastModifiedBy(
  pkg: OpcPackage,
): string | undefined;

export function replacePresentationLastModifiedBy(
  pkg: OpcPackage,
  value: string | undefined,
): void;
```

Wrapper在进入 generic helper前执行 lastModifiedBy-specific strict normalization。Generic helper继续拥有 exact root relationship、content type、namespace/root validation、direct simple-text discovery、minimal part creation、prefix reuse、source-span replace/insert/clear、same-value no-op和unsafe ownership errors。

Revision已引入的 same-namespace minimal creation规则必须被复用：缺失 core part时只声明一次 canonical `xmlns:cp`，并生成 direct `<cp:lastModifiedBy>`。不得再次修改 generic helper，除非 failing test证明现有行为不满足本设计。

## OPC 与 OOXML ownership

Core properties继续通过 package root exact relationship定位：

```text
http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties
```

目标 content type必须是：

```text
application/vnd.openxmlformats-package.core-properties+xml
```

LastModifiedBy只拥有 core-properties root下 direct core-properties namespace child：

```xml
<c:coreProperties
  xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:d="http://purl.org/dc/elements/1.1/">
  <d:creator>Alice</d:creator>
  <c:lastModifiedBy>Editor</c:lastModifiedBy>
  <c:revision>7</c:revision>
</c:coreProperties>
```

实现不得假定 part URI、`cp` lexical prefix或child顺序。Wrong-prefix-but-correct-URI合法；同local name但wrong namespace、descendant-only field、external/dangling relationship、wrong content type、wrong root、多个core relationships或多个direct lastModifiedBy都不能伪造snapshot。

插入时复用root in-scope core prefix。Getter只接受唯一direct simple text；element child或CDATA返回`undefined`。XML entities由lossless parser解码，empty element返回空字符串。读取和mutation不改写creator、title、subject、revision、timestamps、comments、whitespace、unknown children、unrelated parts或relationships。

## 创建与 part lifecycle

`PptxDocument.create({ lastModifiedBy })`先建立canonical package，再仅在值非`undefined`时通过public setter应用。因此 omitted/runtime-undefined保持canonical author和lastModifiedBy；custom/empty只改lastModifiedBy。普通`write()`不会自动刷新该字段，也不会自动递增revision或修改timestamps。

对没有core-properties relationship的合法existing deck：

- 读取返回`undefined`；clear是exact no-op；
- 设置string创建minimal core part，优先使用空闲`/docProps/core.xml`，否则分配安全的next URI；
- 创建唯一root relationship和content-type override；
- 新part只含canonical root namespace与direct lastModifiedBy，不合成creator、revision或timestamps。

对唯一relationship指向的合法part，setter只patch direct field：same-value exact no-op；string replace/insert；`undefined`删除field但保留part、relationship和其他properties。Malformed/ambiguous ownership在mutation前抛`PackageError`或`ModelParseError`并完整rollback。

## PptxGenJS 4.0.1 基线

PptxGenJS 4.0.1没有公开独立`lastModifiedBy` property；它通过公开`author`同时写`dc:creator`与`cp:lastModifiedBy`：

- default author `PptxGenJS`使两个fields都为`PptxGenJS`；
- custom XML-safe author使两个fields都等于custom string并正确escape；
- explicit empty author使两个fields都为direct empty；
- PptxGenJS没有existing-deck editor，也不能通过public API让creator与lastModifiedBy不同。

Adapter必须通过public`author`与`write()`验证default/custom/empty outputs的lastModifiedBy snapshot和reopen stability。Native不复制镜像行为：`author`只拥有creator，`lastModifiedBy`只拥有lastModifiedBy。调用方可显式设置两者相同以得到PptxGenJS final state，也可在editing场景中独立保留或修改审计字段。Native omitted default使用`@jiayunxie/pptx`而非PptxGenJS品牌值。

## 文件边界

`packages/model/src/presentation-last-modified-by.internal.ts`负责descriptor与strict normalization。`packages/model/src/presentation.ts`公开getter/setter并包裹transaction。`packages/sdk/src/create.ts`扩展option；`packages/sdk/src/index.ts`只在显式值存在时通过public setter应用。

Adapter阶段只增加public-output conformance tests；packed smoke覆盖Node/browser/declaration；文档明确independent ownership和PptxGenJS镜像差异。不得更改author setter、自动更新lastModifiedBy、revision或timestamps，不得修改、删除、stage或提交`.pnpm-store/`。

## 测试与发布门禁

1. Internal tests覆盖absent、empty、escapedUnicode/string、same-value no-op、replace、insert、clear、self-closing root/property、alternate prefix、missing relationship creation、occupied URI fallback、wrong namespace、descendant、duplicate、nested/CDATA、dangling/external/wrong-content-type/wrong-root/multiple relationships、rollback和其他core properties保持。
2. Invalid inputs覆盖null、number、bigint、boolean、array、object、symbol和XML controls；所有失败保持parts、relationships、content types、journal和values不变。
3. Model tests覆盖direct getter、edit/empty/clear、rollback、alternate URI/prefix、missing-part creation、same-value/absent-clear no-op、unrelated part isolation、stable slide identity和write/reopen。
4. SDK tests覆盖native omitted/custom/empty/escaping/invalid、author+lastModifiedBy independent combination、immediate getter、edit/clear、rollback、write/reopen和全部六种formats。
5. Adapter conformance只使用PptxGenJS public`author`与`write()`，覆盖default/custom/empty lastModifiedBy snapshot、creator mirror、native independent ownership和reopen；不读取private fields。
6. Packed Node/browser/declaration/CLI smoke覆盖create/read/edit/empty/clear和`string | undefined`类型。
7. Changelog、API README、compatibility baseline和package README新增lastModifiedBy contract，并从metadata backlog移除该项。
8. Full TypeScript、Vitest、1000-part performance、actual tarball smoke、PowerPoint 2010 validation全部通过。
9. 真实native/PptxGenJS decks由`pptx-inspect part read`确认exact lastModifiedBy；native source→edit diff只允许core-properties part变化，edit→reopen零diff。LibreOffice导出和代表PNG证明metadata edit不影响slide content。
