# Presentation Company Metadata Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为 presentation extended properties 增加原生 company 创建、读取、编辑和清除能力，使 `PptxDocument.create({ company })` 与 `document.company` 覆盖从零创建和已有 PPTX 编辑，并对等 PptxGenJS 4.0.1 顶层 `pptx.company` 的合法字符串状态。

本小项只拥有 package-root extended-properties relationship 指向 part 中 direct `Company`。它不修改 `Application`、`AppVersion`、`PresentationFormat`、slide/notes/hidden-slide 计数、heading pairs、titles、链接状态或 custom properties，也不把 company 镜像到 core properties。用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按公开 API 一致性、preserve-first 编辑、OPC 生命周期、namespace correctness 和可独立验收定稿。

## 方案选择

考虑三种方案：

1. 新增专用 extended-properties company helper，集中管理 relationship、part、namespace、source-span patch 和缺失 part 创建；采用。Extended properties 与现有 core-properties helper 的 relationship type、content type、root namespace、合法 child 命名和默认 part bytes 都不同，专用边界最小且易审计。
2. 把 core-properties helper 泛化为跨 core/app/custom part 的通用 metadata engine；不采用。当前只有 company 需要 extended-properties 文本字段，提前引入跨 part descriptor 会扩大 title/author 回归面，并把不同 schema 规则隐藏在过宽抽象中。
3. 把 company 写成 core/custom property，或同时写多个位置；不采用。PptxGenJS 与 OOXML extended properties 都使用 direct `Company`，镜像会制造冲突 ownership，并破坏已有文件中的独立 metadata。

不一次公开完整 document-properties object，也不顺带刷新动态 slide/application statistics。Company 是独立的 presentation 高层属性；其他 metadata 和统计字段继续按小项实现。

## 公共 API

```ts
export interface CreatePresentationOptions {
  readonly author?: string;
  readonly company?: string;
  readonly format?: PresentationFormat;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly title?: string;
}

export class PresentationModel {
  get company(): string | undefined;
  set company(value: string | undefined);
}
```

`CreatePresentationOptions.company` 只接受 string。Omitted 与 runtime `undefined` 不修改原生 canonical app-properties bytes，因此 zero-input native document 的 immediate company 是 `undefined`。显式空字符串写 direct empty Company 并由 getter 返回 `''`。`document.company = undefined` 删除 direct Company。Null、number、boolean、array、object、symbol 和非法 XML 控制字符都不被接受；API 不 trim、不 case-fold、不调用 `toString()`。

Getter 是 direct-state snapshot，不回退到 author、Application、operating-system organization、template、theme、master、custom properties 或 PowerPoint UI 默认值。不存在、安全上歧义或不支持的 direct state返回 `undefined`，且不产生 mutation。

## Extended-properties internal 边界

新增 `presentation-company.internal.ts`，拥有 package-root exact extended-properties relationship、content type、root namespace、direct Company discovery、minimal part creation、namespace-prefix reuse、source-span replace/insert/clear、same-value no-op 和 unsafe ownership errors。

内部接口保持窄且不从 model 聚合入口导出：

```ts
export function readPresentationCompany(pkg: OpcPackage): string | undefined;

export function replacePresentationCompany(
  pkg: OpcPackage,
  value: string | undefined,
): void;
```

Wrapper 内部先执行 company-specific strict normalization，再进入 part lifecycle。该文件不依赖 SDK，不读取 PptxGenJS 对象，也不更新其他 extended properties。

## OPC 与 OOXML ownership

Extended properties 通过 package root 的 exact relationship type 定位：

```text
http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties
```

目标 part 的 content type 必须是：

```text
application/vnd.openxmlformats-officedocument.extended-properties+xml
```

Company 只拥有 extended-properties root 下同 namespace 的 direct child：

```xml
<Properties
  xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <Company>Acme &amp; Partners</Company>
  <AppVersion>16.0000</AppVersion>
</Properties>
```

实现不得假定 part URI、default namespace 或 lexical prefix。以下前缀形式与 default-namespace 形式语义相同：

```xml
<ep:Properties xmlns:ep="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <ep:Company>Acme</ep:Company>
</ep:Properties>
```

同 local name 但 wrong namespace、descendant-only Company、external/dangling relationship、wrong content type、wrong root、多个 extended-properties relationships 或多个 direct Company 都不能伪造 snapshot。插入时复用合法 root 的 namespace lexical form：default namespace root 写 `<Company>`，prefixed root 写相同 prefix 的 `<ep:Company>`，不重写 root namespace declarations。

Getter 只接受唯一 direct Company，其内容必须是 simple text state；element children或 CDATA 等不安全结构返回 `undefined`。XML entities 由 lossless parser 解码，empty element 返回空字符串。读取和 company mutation 不改写 Application、AppVersion、PresentationFormat、统计值、vectors、comments、whitespace、unknown children、unrelated parts 或 relationships。

## 创建、插入顺序与 part 生命周期

原生 zero-input package 已包含 canonical `/docProps/app.xml` 和 root relationship，但不含 Company。`PptxDocument.create({ company })` 先创建 package，再通过同一个 public company setter 写 direct Company，因此创建和 existing-deck 编辑共享归一化、escaping 和 transaction 语义。Omitted company 不触碰 canonical app-properties bytes；explicit company 只更新 extended-properties part。

插入到已有 part 时，在第一个 direct `LinksUpToDate`、`SharedDoc`、`HyperlinksChanged` 或 `AppVersion` 之前写 Company，以维持本库 canonical/PptxGenJS 常见属性顺序；如果这些 following properties 都不存在，则追加为 direct root child。该定位只决定新 Company 的落点，不重排或重写任何已有 child。Self-closing root/Company 必须安全展开。

对没有 extended-properties relationship 的合法 existing deck：

- 读取返回 `undefined`；clear 是 exact no-op。
- 设置 string 创建 minimal extended-properties part，优先使用空闲的 `/docProps/app.xml`，否则分配 `/docProps/app1.xml`、`app2.xml` 等安全 URI。
- 创建对应 content-type override 和唯一 root relationship，target 使用 root-relative package path。
- 新 part 使用 canonical default namespace并只含 direct Company，不合成 Application、AppVersion、PresentationFormat 或统计字段。

对唯一 relationship 指向的合法 part，setter 只 source-span patch direct Company：same-value 是 exact byte/journal no-op；string replace 或 insert direct element；`undefined` 删除该 element但保留 part、relationship和所有其他 properties。不能安全确定 ownership 的 malformed/ambiguous package在任何 mutation 前抛 `PackageError` 或 `ModelParseError`，并完整 rollback。

## PptxGenJS 4.0.1 基线

真实 public-output fixture 确认：

- 未赋值时 public `pptx.company` 是 `PptxGenJS`，并写 direct `<Company>PptxGenJS</Company>`。
- `pptx.company = 'Acme 国际'` 写 direct Company，Unicode 保持。
- `pptx.company = ''` 写 direct empty `<Company></Company>`。
- PptxGenJS 4.0.1 将 company 直接插值到 app.xml，不进行 XML escaping；`A & <B>` 会产生 malformed XML。该 runtime bug 不属于合法对等状态。
- 类型外 `undefined`、null 或 object 会被字符串插值；这些 coercion 不进入 native strict API。

Native company snapshot 对 PptxGenJS default、XML-safe custom 和 empty Company 对等；adapter 导入三种 public outputs 后必须分别读到对应字符串。Native zero-input default 保持现有 `undefined`，不注入 PptxGenJS 品牌值。Native 对 XML metacharacters 正确 escape，clear 与 missing-part creation 是 lossless editing extensions。Adapter 继续保留任意输入 bytes；malformed PptxGenJS app.xml 不得被伪装为受支持的 company snapshot。

## 文件边界

`packages/model/src/presentation-company.internal.ts` 负责：

- root relationship、part、content type 与 namespace resolution；
- strict direct simple-text Company snapshot；
- minimal part/relationship creation；
- prefix reuse、schema-friendly insertion、same-value no-op、replace和 clear patch；
- malformed/ambiguous ownership rejection。

`packages/model/src/presentation.ts` 只公开 company getter/setter并包裹 package transaction。`packages/sdk/src/create.ts` 扩展 `CreatePresentationOptions`，`packages/sdk/src/index.ts` 在构造后通过同一个 public setter应用 explicit company。

每个实现阶段只修改该阶段所需文件。不得修改、删除、stage 或提交 `.pnpm-store/`。

## 测试与发布门禁

1. Company internal tests覆盖 absent、empty、escaped Unicode/string、same-value exact no-op、replace、clear、self-closing root/Company、default/prefixed namespaces、descendant/wrong namespace、duplicate direct Company、Company-before-following-property insertion、missing relationship part creation、occupied canonical URI fallback、dangling/external/wrong-content-type/wrong-root/multiple relationships、rollback和所有其他 app properties保持。
2. Invalid inputs覆盖 null、number、boolean、array、object、symbol、XML controls；所有失败保持 part bytes、relationship XML、content types、journal、slides和 model identity不变。
3. Model tests覆盖 direct getter、edit/empty/clear、transaction rollback、alternate URI/prefix、missing-part creation、same-value/absent-clear no-op、unrelated part isolation和 write/reopen。
4. SDK tests覆盖 native omitted/custom/empty/metacharacter escaping/invalid、immediate getter、edit/clear、rollback、write/reopen和全部六种 presentation formats。
5. Adapter conformance只使用 PptxGenJS public `company` 与 `write()` output，覆盖 default、XML-safe custom、empty、native default差异、native escaping和 write/reopen；不读取私有字段。
6. Packed Node/browser/declaration/CLI smoke覆盖 create、read、edit、clear和 `string | undefined` 类型。
7. Changelog、API README、compatibility baseline和 package README新增 company 行，明确 native/PptxGenJS default差异、PptxGenJS escaping限制和 extended-properties ownership，并保留其他 metadata缺口。
8. Full TypeScript、Vitest、1000-part performance、actual tarball smoke、PowerPoint 2010 profile validation全部通过。
9. 真实 native/PptxGenJS decks由 `pptx-inspect part read` 确认 exact Company、escaping和其他 app properties保持；native source→edit package diff只允许 extended-properties part变化，edit→reopen零 diff。LibreOffice导出无修复；metadata无可视内容，因此视觉验收只确认页面几何与内容没有副作用。
