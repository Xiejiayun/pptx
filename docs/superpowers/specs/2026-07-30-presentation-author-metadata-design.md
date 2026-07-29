# Presentation Author Metadata Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 presentation core properties 增加原生 author 创建、读取、编辑和清除能力，使 `PptxDocument.create({ author })` 与 `document.author` 能覆盖从零创建和已有 PPTX 编辑，并对等 PptxGenJS 4.0.1 顶层 `pptx.author` 的有效字符串 author 状态。

本小项只拥有 OPC core-properties part 中 direct Dublin Core creator。它不同时实现 company、revision、subject、last-modified-by、created/modified time、keywords、category、description、content status 或 custom properties。PptxGenJS 会把同一个 public `author` 同时写入 `dc:creator` 与 `cp:lastModifiedBy`；native author 只修改语义准确的 `dc:creator`，把 last-modified-by 保留为独立待实现字段。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按公开 API 一致性、preserve-first 编辑、OPC 生命周期、namespace correctness 和可独立验收定稿。

## 方案选择

考虑三种方案：

1. 在新的 author helper 中复制 title 的 relationship、part、namespace 和 source-span 逻辑；不采用。后续 subject、description 等字段会继续复制同一安全边界。
2. 抽出聚焦的通用 core-properties internal helper，让 title 与 author 通过字段 descriptor 使用同一 OPC 生命周期；采用。现有 title public/internal 接口保持稳定，author 只新增自己的薄包装和测试。
3. 让 author setter 同时覆盖 `dc:creator` 与 `cp:lastModifiedBy`；不采用。它更接近 PptxGenJS 的生成器实现细节，但会在编辑已有文件时破坏独立的“最后修改者”信息，也让 same-value author assignment 可能产生隐藏 mutation。

不把 metadata 建模为一次性 object snapshot，也不一次公开全部 document properties。Author 是 presentation 的稳定高层属性；last-modified-by 后续通过独立小项拥有。

## 公共 API

```ts
export interface CreatePresentationOptions {
  readonly author?: string;
  readonly format?: PresentationFormat;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly title?: string;
}

export class PresentationModel {
  get author(): string | undefined;
  set author(value: string | undefined);
}
```

`CreatePresentationOptions.author` 只接受 string。Omitted 与 runtime `undefined` 保留原生 canonical core-properties bytes，因此 zero-input native document 的 immediate author 是现有 direct `@jiayunxie/pptx`。显式空字符串写 direct empty creator 并由 getter 返回 `''`。`document.author = undefined` 删除 direct creator；调用方若要创建无 author 文件，可在 `create()` 后显式 clear。Null、number、boolean、array、object、symbol 和非法 XML 控制字符都不被接受；API 不 trim、不 case-fold、不调用 `toString()`。

Getter 是 direct-state snapshot，不回退到 last-modified-by、company、file owner、operating-system account、first slide、template、master 或 PowerPoint UI 默认值。不存在、安全上歧义或不支持的 direct state返回 `undefined`，且不产生 mutation。

## 通用 core-properties internal 边界

新增 `presentation-core-properties.internal.ts`，拥有 package root exact core-properties relationship、content type、root namespace、direct simple-text field discovery、minimal part creation、prefix reuse、source-span replace/insert/clear、same-value no-op 和 unsafe ownership errors。

字段通过 internal descriptor 声明：

```ts
interface CoreTextPropertyDescriptor {
  readonly label: string;
  readonly localName: string;
  readonly namespace: string;
  readonly preferredPrefix: string;
}
```

通用 helper 只接受已经过字段 wrapper 严格归一化的 `string | undefined`，并暴露 internal `readCoreTextProperty()` 与 `replaceCoreTextProperty()`。`presentation-title.internal.ts` 保留现有 exports，但改为 title descriptor 的薄包装；其现有测试必须无行为变化。`presentation-author.internal.ts` 定义 creator descriptor 和 author-specific validation/error wording。任何 internal helper 都不得从 `packages/model/src/index.ts` 或聚合 npm 包导出。

## OPC 与 OOXML ownership

Core properties 继续通过 package root exact relationship type 定位：

```text
http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties
```

目标 part 的 content type 必须是：

```text
application/vnd.openxmlformats-package.core-properties+xml
```

Author 只拥有 core-properties root 下 direct Dublin Core namespace child：

```xml
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>Alice &amp; Bob</dc:creator>
  <cp:lastModifiedBy>Editor</cp:lastModifiedBy>
</cp:coreProperties>
```

实现不得假定 part URI、`cp`/`dc` lexical prefix 或 child 顺序。Wrong-prefix-but-correct-URI 是合法输入；同 local name 但 wrong namespace、descendant-only creator、external/dangling relationship、wrong content type、wrong root、多个 core-properties relationships 或多个 direct Dublin Core creators 都不能伪造 snapshot。插入时优先复用 root in-scope 的 Dublin Core prefix；没有可复用 binding 时，新 `dc:creator` 在自身声明 canonical `xmlns:dc`，不重写 root namespaces。

Getter 只接受唯一 direct creator，其内容必须是 simple text state；element children 或 CDATA 等不安全结构返回 `undefined`。XML entities 由 lossless parser 解码，empty element 返回空字符串。读取和 author mutation 不改写 title、subject、lastModifiedBy、revision、timestamps、comments、whitespace、unknown children、unrelated parts 或 relationships。

## 创建与 part 生命周期

原生 zero-input package 已包含 canonical `/docProps/core.xml`、root relationship、direct creator 与 direct lastModifiedBy。`PptxDocument.create({ author })` 先创建 package，再通过同一个 author setter 写 direct creator，因此创建与 existing-deck 编辑共享归一化、escaping 和 transaction 语义。Omitted author 不触碰 canonical core-properties bytes；explicit author 只更新 core-properties part 中 creator。

对没有 core-properties relationship 的合法 existing deck：

- 读取返回 `undefined`；clear 是 exact no-op。
- 设置 string 创建 minimal core-properties part，优先使用空闲的 `/docProps/core.xml`，否则分配 `/docProps/core1.xml`、`core2.xml` 等安全 URI。
- 创建对应 content-type override 和唯一 root relationship，target 使用 root-relative package path。
- 新 part 只含 canonical namespace declarations 与 direct creator，不合成 title、lastModifiedBy、revision 或 timestamps。

对唯一 relationship 指向的合法 part，setter 只 source-span patch direct creator：same-value 是 exact byte/journal no-op；string replace 或 insert direct element；`undefined` 删除该 element但保留 part、relationship 和所有其他 properties。Self-closing root/creator 必须安全展开。不能安全确定 ownership 的 malformed/ambiguous package 在任何 mutation 前抛 `PackageError` 或 `ModelParseError`，并完整 rollback。

## PptxGenJS 4.0.1 基线

真实 public-output fixture 确认：

- 未赋值时 public `pptx.author` 是 `PptxGenJS`，并写 direct `<dc:creator>PptxGenJS</dc:creator>`。
- `pptx.author = 'Alice & <Bob>'` 正确 XML-escape 并写 direct creator。
- `pptx.author = ''` 写 direct empty `<dc:creator></dc:creator>`。
- PptxGenJS 同时把上述值镜像到 `cp:lastModifiedBy`，但没有 existing-deck metadata editor 或独立 last-modified-by API。

Native author snapshot 对 PptxGenJS default/custom/empty creator 对等；adapter 导入三种 public outputs 后必须分别读到对应字符串。Native zero-input default 保持现有 `@jiayunxie/pptx`，而不是伪装成 PptxGenJS 品牌默认值。Native custom author 只修改 creator 并保留 lastModifiedBy；这是 preserve-first existing-deck extension 和有意的生成器差异。Native clear 与 missing-part creation 也是 lossless editing extensions。

## 文件边界

`packages/model/src/presentation-core-properties.internal.ts` 负责：

- root relationship、part、content type 与 namespace resolution；
- generic direct simple-text snapshot；
- minimal part/relationship creation；
- prefix reuse、same-value no-op、replace、insert 和 clear patch；
- malformed/ambiguous ownership rejection。

`packages/model/src/presentation-title.internal.ts` 保留 title-specific strict normalization 和现有 exports，转调 generic helper。`packages/model/src/presentation-author.internal.ts` 负责 author-specific strict normalization 和 descriptor。`packages/model/src/presentation.ts` 只公开 author getter/setter 并包裹 package transaction。`packages/sdk/src/create.ts` 扩展 `CreatePresentationOptions`，`packages/sdk/src/index.ts` 在构造后通过同一个 public setter 应用 explicit author。

每个实现阶段只修改该阶段所需文件。不得修改、删除、stage 或提交 `.pnpm-store/`。

## 测试与发布门禁

1. Generic refactor先运行全部现有 title internal/model/SDK/adapter测试并证明行为与 bytes 不变。
2. Author internal/model tests覆盖 canonical default、absent、empty、escaped Unicode/string、same-value exact no-op、replace、clear、self-closing root/creator、alternate prefixes、descendant/wrong namespace、duplicate direct creators、missing relationship part creation、occupied canonical URI fallback、dangling/external/wrong-content-type/wrong-root/multiple relationships、rollback和 title/lastModifiedBy/其他 properties保持。
3. Invalid inputs覆盖 null、number、boolean、array、object、symbol、XML controls；所有失败保持 part bytes、relationship XML、content types、journal、slides和model identity不变。
4. SDK tests覆盖 native default/custom/empty/invalid、immediate getter、edit/clear、transaction rollback、write/reopen和全部六种 presentation formats。
5. Adapter conformance只使用 PptxGenJS public `author` 与 `write()` output，覆盖 default/custom/empty creator import、PptxGenJS lastModifiedBy镜像、native lastModifiedBy preservation和write/reopen；不读取私有字段。
6. Packed Node/browser/declaration/CLI smoke覆盖 create、read、edit、clear和 `string | undefined` 类型。
7. Changelog、API README、compatibility baseline和package README新增 author 行，明确 native/PptxGenJS default差异与 lastModifiedBy ownership差异，并保留其他 metadata缺口。
8. Full TypeScript、Vitest、1000-part performance、actual tarball smoke、PowerPoint 2010 profile validation全部通过。
9. 真实 native/PptxGenJS decks由 `pptx-inspect part read` 确认 exact creator、escaping和 lastModifiedBy preservation；native source→edit package diff只允许 core-properties part变化，edit→reopen零 diff。LibreOffice导出无修复；metadata无可视内容，因此视觉验收只确认页面几何与内容没有副作用。
