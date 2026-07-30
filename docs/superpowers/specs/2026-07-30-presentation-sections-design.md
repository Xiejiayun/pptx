# Presentation Sections Design

日期：2026-07-30
状态：已批准实施（用户已授权自主选择最佳方案并持续推进）

## 目标与范围

补齐 PptxGenJS 4.0.1 `pptx.addSection({ title, order? })` 与 `pptx.addSlide({ sectionTitle })` 的原生能力，并为已有 PPTX 增加 section 的严格读取、重命名、排序、删除和 slide 归属编辑。创建、编辑、slide 生命周期、adapter 导入、浏览器/Node 打包和真实 PowerPoint 2010 扩展输出必须保持一致。

本小项只拥有 presentation part direct PowerPoint 2010 section extension 中的 section 顺序、section ID、name 与 slide ID membership。它不实现 named slide master、layout selection、hidden slide、slide number、speaker notes、custom properties、PowerPoint 2012 slide guides 或 UI collapse state，也不把 section title 当作 presentation title、subject 或 slide title。未知 extensions 和 section 容器内未拥有的 XML 必须无损保留。

## 真实 PptxGenJS 4.0.1 基线

公开 runtime 与 `write()` 输出已验证以下行为：

- 没有调用 `addSection()` 时不写 section extension；
- `addSection({ title })` 可先创建 empty section，slide 通过 `addSlide({ sectionTitle })` 按第一个同名 section 归属；
- section name 会正确 XML escape；
- positive `order` 直接传给 JavaScript `splice(order, 0, section)`，因此实际为 zero-based insertion；runtime `order: 0` 因 falsy 分支错误地 append；
- sections 已启用而 `addSlide()` 未指定 title 时，PptxGenJS 创建或延续内部 default section，名称为 `Default-1`、`Default-2` 等；
- section 建立前已经存在的 slide 可以保持 loose，unknown `sectionTitle` 只 warning 并留下 loose slide；
- empty section、loose slide 和 section list 未覆盖全部 presentation slides 的文件均通过当前 PowerPoint 2010 validator；
- 每次 write 为 section 生成新的 brace-wrapped UUID，并按 presentation `p:sldId@id` 写 membership；同时写一个无内容的 PowerPoint 2012 slide-guide extension。

原生实现对等有效、明确的 final section state，不复制 warning-only invalid handling、每次 write 刷新 section ID、`order: 0` append 缺陷，也不为了 section 创建无关的 empty slide-guide extension。Adapter 只通过 PptxGenJS public `write()` 导入最终 XML，不读取 `_sections`、`_type` 或 `_slides`。

## 方案选择

考虑三种方案：

1. 只增加 PptxGenJS-shaped `addSection()` 和 `addSlide({ sectionTitle })`。实现最短，但无法读取或编辑已有 sections，不满足双向库目标。
2. 暴露 `document.sections` 整组 setter。它能表达编辑，却会在单次 rename 时重写完整 section extension，难以保留未知 attributes、children、prefix 和 whitespace。
3. 暴露 detached section snapshots，加上按稳定 section ID 执行的原子 add/rename/move/delete/assign 命令。内部 helper 只 source-span patch 被拥有的节点；PptxGenJS convenience API 与 existing-deck editor 共用同一严格边界。采用此方案。

不增加 live `SectionModel`。Section 数据很小，稳定 OOXML ID 已足以寻址；detached snapshots 避免缓存失效和外部对象 mutation，同时让每个命令在 transaction 内重新验证最新 package state。

## 公共 API

```ts
export interface PresentationSection {
  readonly id: string;
  readonly title: string;
  readonly slideIds: readonly number[];
}

export interface AddSectionOptions {
  readonly title: string;
  readonly order?: number;
}

export interface AddSlideOptions {
  readonly sectionTitle?: string;
}

export class PresentationModel {
  get sections(): readonly PresentationSection[] | undefined;

  addSection(options: AddSectionOptions): PresentationSection;
  renameSection(sectionId: string, title: string): void;
  moveSection(sectionId: string, toIndex: number): void;
  deleteSection(sectionId: string): void;
  assignSlideToSection(slideIndex: number, sectionId: string | undefined): void;

  addSlide(options?: AddSlideOptions): SlideModel;
}
```

`sections` 为 detached direct-state snapshot。没有 section extension 或存在合法 empty section list 时返回 `[]`；extension ownership、section structure、ID 或 membership 歧义时返回 `undefined`。每个 snapshot 的 `slideIds` 使用 presentation `p:sldId@id`，不是 zero-based index、relationship ID 或 part URI。Snapshot 与 nested arrays 每次读取都新建，调用方修改不会改变文稿。

`addSection()` 返回新 section 的 detached snapshot。`order` 是 strict zero-based insertion index，允许 `0..sections.length`；省略时 append。此定义对 positive PptxGenJS values 产生相同顺序，并修复其 `order: 0` falsy bug。

`renameSection()`、`moveSection()`、`deleteSection()` 按 ID 寻址，因此 duplicate titles 合法且不会造成编辑歧义。删除 section 只删除 section metadata，绝不删除其 slides；原成员变为 loose。

`assignSlideToSection()` 使用 zero-based slide index。指定 ID 时先从全部 sections 移除目标 slide，再 append 到目标 section；`undefined` 只取消 section 归属。其他 slide membership 和 section 顺序保持。

`addSlide({ sectionTitle })` 对 duplicate title 选择第一个，和 PptxGenJS 一致。Unknown title 在创建任何 part、relationship 或 slide ID 前抛 `RangeError`。未提供 title 且 sections 已存在时，原生沿用 PptxGenJS final-state 规则：若最后一个 section 是 canonical `Default-N`，新 slide 加入该 section；否则创建下一个未占用的 `Default-N` 并归属。由于 PptxGenJS 的 hidden `_type` 不写入 OOXML，重新打开后只能从 canonical name 恢复这一语义；用户自己创建同名 section 的 hidden intent不可从文件重建，native 以持久 final state 为准。

## 输入契约

Section title 必须是包含至少一个非空白字符的 string，不能包含 XML 1.0 非法字符。合法值保留原字符串，不 trim、不 case-fold，XML 元字符正确转义。Empty、纯空白、非 string、accessor property、symbol property、继承 property、数组、未知 own key 与污染 prototype 在 package mutation 前拒绝。

`AddSectionOptions` 只允许 own data `title`、`order`；`AddSlideOptions` 只允许 own data `sectionTitle`。`undefined` sectionTitle 等同省略。Order 与 move target 必须是 finite safe integer 且位于允许范围，不 floor、不 clamp。Section ID 参数必须是 snapshot 中 exact brace-wrapped GUID；title 与 ID lookup 均区分大小写，GUID validation 接受 upper/lower hex但不去掉 braces。

所有 inputs 在进入 transaction 前 descriptor-safe normalization 并立即与 caller 脱离。Invalid input、unknown section、unknown slide、unsafe section XML 或重复 membership 必须保持 package parts、relationships、content types、journal、slides 和现有 model identity 零变化。

## OOXML ownership 与严格读取

Sections 只位于 presentation root 的 direct extension：

```xml
<p:extLst>
  <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
    <p14:sectionLst
      xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
      <p14:section name="Intro" id="{01234567-89AB-CDEF-0123-456789ABCDEF}">
        <p14:sldIdLst>
          <p14:sldId id="256"/>
        </p14:sldIdLst>
      </p14:section>
    </p14:sectionLst>
  </p:ext>
</p:extLst>
```

Reader 必须找到唯一 PresentationML namespace root、至多一个 direct `extLst`、至多一个 direct `p:ext` whose unqualified `uri` equals section URI case-insensitively，以及该 extension 中唯一 direct PowerPoint 2010 namespace `sectionLst`。不得按任意 local name、固定 `p`/`p14` prefix、descendant 或 first extension 猜测。

每个 direct section 必须有唯一 unqualified `name` 与 `id`、唯一 direct PowerPoint 2010 `sldIdLst`，ID 为 brace-wrapped UUID且 section IDs 不重复。每个 direct member必须有唯一 unqualified decimal `id`，值是合法 presentation slide ID；同一 slide 不可出现在多个 sections。Empty section 与未归属 presentation slides 合法。Wrong namespace、duplicate direct nodes/attributes、unknown member ID、non-decimal/unsafe ID、重复 membership 或多个 owned extensions 使整体 snapshot 为 `undefined`，任何 section mutation 明确失败。

未知 attributes、section/section list 中非 owned children、extension siblings、PowerPoint 2012 slide guides、alternate prefixes、namespace declarations、comments、whitespace 和 root child 顺序均保留。Unknown XML 不被解释为 section；malformed XML 继续由 lossless parser抛出明确 parse error。

## Source-span mutation

新增 `packages/model/src/presentation-sections.internal.ts`，集中负责 strict normalization、namespace-aware discovery、snapshot、UUID generation与最小 patch：

- 无 section extension时，`addSection()` 在已有 direct `extLst` 中 append canonical section extension；没有 `extLst` 时按 schema-last 位置创建；
- 已有合法 section list时 append/insert only one canonical `section`，复用 list lexical prefix；
- rename只替换 direct `name` attribute，same value exact no-op；
- membership editor只增删 direct member节点，保留 list 的其他 children；
- move以原始 section byte slices重排 owned section elements，保留每个 section内部 bytes以及 section list中的foreign children；same index exact no-op；
- delete只移除目标 section。删除最后一个时，仅在 section extension无foreign state时删除该 extension；否则保留 empty list。若 direct `extLst` 因此无任何 children再删除它；
- 所有新 GUID 使用 `globalThis.crypto.randomUUID()` 生成 uppercase brace-wrapped value，在缺少 secure UUID API时于 mutation前失败，不用 `Math.random()`；
- helper 返回是否 changed，PresentationModel 只有在 changed 时 set part，保持 journal no-op。

新 XML 使用 presentation root lexical prefix；section namespace局部声明为 canonical `p14`，即使 root 未声明 p14也不重写 root namespaces。Escaping使用 lossless-xml公共 escape helpers。

## Slide 生命周期一致性

Section membership引用 presentation slide ID，因此现有结构命令必须同步：

- `addSlide(options)` 在同一 outer OPC transaction 中先验证 section state与options，再创建 slide并写 membership；任何后续失败回滚全部新增 part/relationship/content type/XML；
- `duplicateSlide(index)` 若 source属于一个section，则新 slide加入同一section；source loose时duplicate也loose；
- `deleteSlide(index)` 在删除 slide part/relationship前从所有 section lists移除其 ID；empty section保留；unsafe section state使整个 delete在 mutation前拒绝；
- `moveSlide()` 保持 membership，不自动把 slide转移到相邻 section；每个 section内的 member ID顺序按新的全局 presentation slide order稳定排序；
- `assignSlideToSection()` 是唯一显式跨 section移动命令；
- section rename/move/delete不改变 presentation slide order、slide parts或relationships。

以上 native lifecycle 是 PptxGenJS 没有 existing-deck model 的扩展。Slide duplication/deletion现有 dependency clone/GC语义不变。

## 原子性、格式与兼容性

所有 commands 使用 `OpcPackage.transaction()`。Normalization 与 ownership validation在首次 mutation前完成；外层 `document.transaction()` 抛错时 section XML、slide topology和model cache一起回滚。Same-value rename、same-position move、重复 assign到相同 final membership、absent unassign均为 exact byte/journal no-op。

Section extension是 Office 2010+ feature。六种现有 presentation formats `pptx/pptm/ppsx/ppsm/potx/potm` 共用相同 presentation XML模型和 extension；格式与宏 payload不改变。PowerPoint 2010/current应为0 error/0 warning；LibreOffice可以忽略UI section metadata，但必须无修复打开且保留slides与页面几何。

## 文件边界

- `packages/model/src/presentation-sections.internal.ts`：strict direct-state codec与source-span mutation；不从聚合包直接导出。
- `packages/model/src/presentation.ts`：public types、snapshot getter、commands和slide lifecycle协调。
- `packages/model/src/index.ts`：导出 public section types。
- `packages/sdk/src/create.ts`：不新增 create option；sections通过 addSection/addSlide组合创建。
- `packages/pptxgenjs-adapter/src/index.test.ts`：只用PptxGenJS public APIs验证final-state import。
- `scripts/smoke-npm-package.mjs`：packed Node/browser/declaration API证明。
- compatibility、package README与CHANGELOG：精确记录支持范围和strict repairs。

不得修改、删除、stage或提交 `.pnpm-store/`。

## 测试与验收

1. Internal fixtures覆盖absent、empty、single/multiple、empty section、loose slide、duplicate title、alternate prefixes/local namespace、escaped Unicode title、foreign extensions/children与PptxGenJS section+slide-guide输出。
2. Strict read覆盖duplicate extension/list/section IDs/membership、wrong namespace、descendant impostor、missing/repeated attributes/lists、invalid GUID、unknown/non-decimal/unsafe slide IDs；unsafe snapshot为`undefined`且mutation零变化。
3. Commands覆盖add append/order 0/middle/end、rename/escaping/no-op、move/no-op、delete first/middle/last、assign/reassign/unassign/no-op、detached snapshots、invalid input与outer rollback。
4. Slide lifecycle覆盖explicit sectionTitle、duplicate names first-match、unknown title zero mutation、canonical Default-N creation/continuation、loose-before-section、duplicate/delete/move membership和dependency isolation。
5. PptxGenJS 4.0.1 public conformance覆盖none、explicit/escaped、empty section、positive order、default section、loose-before、unknown-title final state；adapter精确读取IDs以外的title/order/membership并在native write/reopen后保持。
6. Packed Node、browser和declaration smoke覆盖all public methods、detached snapshot、write/reopen以及`readonly PresentationSection[] | undefined`类型。
7. `pptx-inspect`确认native/PptxGenJS files为PowerPoint 2010 profile 0 error/0 warning；rename diff只允许presentation part变化，same-value和reopen为零part变化。
8. LibreOffice headless导出native source/edited/reopened与PptxGenJS baseline；页面数、尺寸与可见内容保持，render/overflow检查无异常。Section无可视内容，因此视觉QA只验证无副作用。
9. 全量typecheck、Vitest、performance、browser build、actual tarball smoke、git diff checks和远端`origin/main...HEAD = 0 0`全部通过。
