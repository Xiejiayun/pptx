# Presentation Title Metadata Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 presentation core properties 增加原生 title 创建、读取、编辑和清除能力，使 `PptxDocument.create({ title })` 与 `document.title` 能覆盖从零创建和已有 PPTX 编辑，并对等 PptxGenJS 4.0.1 顶层 `pptx.title` 的有效字符串输出。

本小项只拥有 OPC core-properties part 中 direct Dublin Core title，不同时实现 author、company、revision、subject、created/modified time、keywords、category、description、content status 或 custom properties。Slide title、shape text、file name和 PowerPoint 窗口标题不属于本字段。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按公开 API 一致性、OPC 生命周期、namespace correctness、无损编辑和可独立验收定稿。

## 方案选择

考虑三种方案：

1. 在 `PresentationModel` 内直接解析和修改 `/docProps/core.xml`；不采用。它对 title 足够小，但后续 author、subject 和 revision 会重复关系解析、part 创建和 namespace 规则。
2. 新建聚焦的 core-properties internal helper，本小项只公开 title；采用。它让公共范围保持单一，同时为后续元数据字段复用安全的 part/relationship 生命周期。
3. 一次实现全部 presentation metadata；不采用。它会把 core properties、extended properties 和不同值类型混进同一 review 单元，不符合逐小项 commit 的约束。

不把 metadata 建模为一次性 object snapshot，也不要求调用方使用 raw part API。Title 是 presentation 的稳定高层属性。

## 公共 API

```ts
export interface CreatePresentationOptions {
  readonly format?: PresentationFormat;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly title?: string;
}

export class PresentationModel {
  get title(): string | undefined;
  set title(value: string | undefined);
}
```

`CreatePresentationOptions.title` 只接受 string。Omitted 与 runtime `undefined` 保留当前原生创建 bytes，不写 `dc:title`；显式空字符串写一个 direct empty title 并由 getter 返回 `''`。`document.title = undefined` 删除 direct title，`document.title = ''` 保留 direct empty title。Null、number、boolean、array、object、symbol 和非法 XML 控制字符都不被接受；API 不 trim、不 case-fold、不调用 `toString()`。

Getter 是 direct-state snapshot，不回退到 file name、first slide title、subject、template、master 或 PowerPoint UI 默认值。不存在、安全上歧义或不支持的 direct state 返回 `undefined`，且不产生 mutation。

## OPC 与 OOXML ownership

Core properties 通过 package root 的 exact relationship type 定位：

```text
http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties
```

目标 part 的 content type 必须是：

```text
application/vnd.openxmlformats-package.core-properties+xml
```

Title 只拥有 core-properties root 下 direct Dublin Core namespace child：

```xml
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Quarterly &amp; &lt;Review&gt;</dc:title>
</cp:coreProperties>
```

实现不得假定 part URI、`cp`/`dc` lexical prefix 或 child 顺序。Helper 根据 root relationship 的 resolved target 定位 part，并按 in-scope `xmlns` declarations 解析 root 与 direct child namespace。Wrong-prefix-but-correct-URI 是合法输入；同 local name 但 wrong namespace、descendant-only title、external/dangling relationship、wrong content type、wrong root、多个 core-properties relationships 或多个 direct Dublin Core titles 都不能伪造 snapshot。插入时优先复用 root in-scope 的 Dublin Core prefix；没有可复用 binding 时，新 `dc:title` 在自身声明 canonical `xmlns:dc`，不重写 root namespaces。

Getter 只接受唯一 direct title，其内容必须是 simple text state；element children 或 CDATA 等不安全结构返回 `undefined`。XML entities 由 lossless parser 解码，空 element 返回空字符串。读取不改写 prefix、whitespace、comments、其他 core fields 或 package graph。

## 创建与 part 生命周期

原生 zero-input package 已包含 canonical `/docProps/core.xml` 与 root relationship。`PptxDocument.create({ title })` 先创建 package，再通过同一个 title setter 写 direct state，因此创建与 existing-deck 编辑共享归一化、escaping 和 transaction 语义。Omitted title 不触碰现有 canonical core-properties bytes；显式 title 只更新 core-properties part。

对没有 core-properties relationship 的合法 existing deck：

- 读取返回 `undefined`；清除是 exact no-op。
- 设置 string 创建 minimal core-properties part，优先使用空闲的 `/docProps/core.xml`，否则分配 `/docProps/core1.xml`、`core2.xml` 等安全 URI。
- 创建对应 content-type override 和唯一 root relationship，target 使用 root-relative package path。
- 新 part 只含 canonical namespace declarations 与 direct title，不合成 creator、revision 或 timestamps。

对唯一 relationship 指向的合法 part，setter 只 source-span patch direct title：same-value 是 exact byte/journal no-op；string 替换或插入 direct element；`undefined` 删除该 element但保留 part、relationship 和所有其他 properties。Self-closing root/title 必须安全展开。不能安全确定 ownership 的 malformed/ambiguous package 在任何 mutation 前抛 `PackageError` 或 `ModelParseError`，并完整 rollback。

## PptxGenJS 4.0.1 基线

真实 public-output fixture 确认：

- 未赋值时 PptxGenJS 写 `<dc:title>PptxGenJS Presentation</dc:title>`。
- `pptx.title = 'Quarterly & <Review>'` 正确 XML-escape 并写同一 direct field。
- `pptx.title = ''` 写 direct empty `<dc:title></dc:title>`。
- PptxGenJS 没有 existing-deck metadata editor，也不暴露 direct absence。

Native explicit strings与 PptxGenJS final title snapshot 对等；adapter 导入默认、自定义和 empty title 后必须分别读到对应字符串。Native omitted 保持当前无 direct title 的 canonical output，而不是伪装成 PptxGenJS 品牌默认值；这是 zero-input default 差异。Native `undefined` clear 与 missing-part creation是 lossless editing extension。

## 文件边界

`packages/model/src/presentation-title.internal.ts` 负责：

- strict string normalization；
- root core-properties relationship 与 part resolution；
- namespace-aware direct title snapshot；
- minimal part/relationship creation；
- same-value no-op、replace、insert 和 clear patch。

`packages/model/src/presentation.ts` 只公开 `title` getter/setter并包裹 package transaction。`packages/sdk/src/create.ts` 扩展 `CreatePresentationOptions`，`packages/sdk/src/index.ts` 在构造后通过同一个 public setter应用显式 title。不得把 internal helper从聚合包导出。

每个实现阶段只修改该阶段所需文件。不得修改、删除、stage 或提交 `.pnpm-store/`。

## 测试与发布门禁

1. Internal/model tests覆盖 absent、empty、escaped Unicode/string、same-value exact no-op、replace、clear、self-closing root/title、alternate prefixes、descendant/wrong namespace、duplicate direct titles、missing relationship part creation、occupied canonical URI fallback、dangling/external/wrong-content-type/wrong-root/multiple relationships、rollback和其他 properties保持。
2. Invalid inputs覆盖 null、number、boolean、array、object、symbol、XML controls；所有失败保持 part bytes、relationship XML、content types、journal、slides和model identity不变。
3. SDK tests覆盖 create omitted/custom/empty/invalid、immediate getter、edit/clear、transaction rollback、write/reopen和全部六种 presentation formats。
4. Adapter conformance只使用 PptxGenJS public `write()` output，覆盖 default、自定义、empty title import、native explicit output和write/reopen；不读取私有字段。
5. Packed Node/browser/declaration/CLI smoke覆盖 create、read、edit、clear和 `string | undefined` 类型。
6. Changelog、API README、compatibility baseline和package README新增 title 行，明确 native omitted default、PptxGenJS品牌默认、direct empty与clear差异，并保留其他 metadata缺口。
7. Full TypeScript、Vitest、1000-part performance、actual tarball smoke、PowerPoint 2010 profile validation全部通过。
8. 真实 native/PptxGenJS decks由 `pptx-inspect part read` 确认 exact title和escaping；native source→edit package diff只允许 core-properties part变化，edit→reopen零 diff。LibreOffice导出无修复；metadata无可视内容，因此视觉验收只确认页面几何与内容没有副作用。
