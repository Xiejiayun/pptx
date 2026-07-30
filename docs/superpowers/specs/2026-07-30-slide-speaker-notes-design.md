# Slide Speaker Notes Design

日期：2026-07-30
状态：已批准（用户已授权实现方持续选择最佳方案并推进，不等待常规设计确认）

## 1. 目标与边界

为原生库增加 PptxGenJS 4.0.1 `slide.addNotes(string)` 的有效输出对等，并把同一 plain-text 能力扩展为已有 PPTX 的严格读取、无损局部编辑、显式清除、复制保持和事务回滚。

本小项拥有 slide 到唯一 notes slide 的 direct relationship、该 per-slide notes part、notes part 中唯一 body placeholder 的文本内容，以及在 presentation 完全没有 notes master 时创建一条 canonical notes master chain 所需的最小拓扑。它不实现 rich notes 公共 API、notes page 自定义排版、header/footer/date、评论、讲义、slide number 显示控制、custom shows 或放映范围。未拥有的 notes shapes、placeholder metadata、body properties、extensions、relationships、notes master/theme bytes 和其他 package parts必须无损保留。

## 2. 公共 API

```ts
class SlideModel {
  get notes(): string | undefined;
  set notes(value: string | undefined);
  addNotes(value: string): this;
}
```

`addNotes()` 是 PptxGenJS 迁移入口，内部等同于 `slide.notes = value` 并返回同一个 live `SlideModel`。Property 同时覆盖从零创建和 existing-deck editing。

- 没有 notes slide relationship 时，getter 返回 `undefined`。
- 安全 notes slide 存在且 body placeholder 可读时返回 plain string；合法空 body 返回 `''`。
- relationship、part、root、shape tree 或 body placeholder ownership 不安全时返回 `undefined`，读取保持 exact no-op。
- Setter 接受 string 或 `undefined`。String 统一 CRLF/CR 为 LF，拒绝 XML 1.0 非法控制字符；`undefined` 表示删除该 slide 的完整 notes slide relationship 与 owned notes part。
- `addNotes()` 只接受 string，不把 `undefined` 当清除，不接受 coercion。

显式删除 per-slide notes part 会移除该 notes page 内的全部 per-slide 内容；shared notes master、theme 和其他 slides 的 notes 不受影响。只想保留 notes page 并清空正文时使用 `slide.notes = ''`。

## 3. 方案取舍

考虑过三种方案：

1. `notes` live property 加 `addNotes()` compatibility method；采用。它同时满足 PptxGenJS creation、existing-deck editing 和清除语义。
2. 只实现 `addNotes(string)`；拒绝。它无法读取或编辑已有 notes，违背双向对象模型目标。
3. 直接公开 rich notes paragraphs、notes page shapes 与 notes master layout；拒绝。本轮基线只有 PptxGenJS plain string，扩大为 rich page editor 会把多个独立能力混入一个 review 单元。

## 4. PptxGenJS 4.0.1 基线

只通过 public constructor、`addSlide()`、public `addNotes()` 和 public `write()` 生成真实输出确认：

- 每个 slide 都生成一个 notes slide part，即使没有调用 `addNotes()`；omitted 与 explicit empty 的 body text 都为空。
- 非空 string 写入 notes slide 唯一 `p:ph type="body"` shape 的 `a:t`；XML metacharacters被转义。
- LF、CRLF 和 CR runtime 输入在最终 `a:t` 中表现为换行文本；native 统一为 LF 以获得确定性输出。
- notes slide 通过 internal relationships 同时引用 source slide 与 shared notes master；slide 通过 internal relationship 引用 notes slide。
- PptxGenJS public type 只接受 string，方法返回 slide。

Native 不复制“所有空 slide 都物化 notes part”的体积副作用：没有 notes 的 native slide 保持无 notes relationship，语义 snapshot 为 `undefined`。一旦写入 empty 或 non-empty string，native 建立完整 notes chain。Adapter 继续只导入 public `write()` 的最终 OOXML，不读取 `_slideObjects`、`_notes` 或其他私有状态。

## 5. OOXML ownership 与严格读取

新增 `packages/model/src/slide-notes.internal.ts`，集中负责 notes relationship/part 定位、plain-text codec 和 canonical part creation。核心接口为：

```ts
readSlideNotes(pkg: OpcPackage, presentationPartUri: string, slidePartUri: string): string | undefined;
replaceSlideNotes(
  pkg: OpcPackage,
  presentationPartUri: string,
  slidePartUri: string,
  value: string | undefined,
): boolean;
```

安全 notes state 要求：

- slide 至多一个 exact notesSlide relationship，且必须为 internal、resolved target 存在、target part content type 为 PresentationML notesSlide；
- notes XML 只有一个 root，root local name 为 `notes`，其 lexical namespace解析为 PresentationML；
- root 下恰好一个 direct `cSld`，其中恰好一个 direct `spTree`；
- body owner 是唯一 direct `sp`，其 direct `nvSpPr/nvPr/ph` 的 exact unqualified `type` 为 `body`；qualified lookalike、descendant placeholder 和 slide-image/slide-number placeholder不参与；
- body shape 至多一个 direct `txBody`，text body 只按 direct paragraph 顺序读取；paragraph中的 direct/descendant `t` 按文档顺序连接，`br` 为 LF，paragraph之间为 LF；
- duplicate relationships、wrong content type、external target、missing part、wrong root namespace、multiple shape trees、multiple body placeholders 或 ambiguous text body 均为 unsafe。

Malformed XML 继续由 lossless parser抛出 parse error。Getter 不猜测由 shape name、index、位置或文本内容定位 notes body。

## 6. 创建、编辑、清除与 notes master

所有输入归一化和 ownership 检查在第一次 package mutation 前完成，并包在一个 `OpcPackage.transaction()` 中。

### String 创建

当 slide 没有 notes relationship 时：

1. 解析 presentation notes master state。存在唯一合法 internal notesMaster relationship 时复用它。
2. 若 presentation 完全没有 notes master，则创建 canonical notes master part、presentation relationship 与 direct `notesMasterIdLst/notesMasterId`；theme 选择顺序是唯一 direct presentation theme，其次 presentation 顺序中第一个 slide master 的唯一 internal theme。没有可安全复用的 theme 或存在歧义时在 mutation 前拒绝，不猜测或复制 opaque theme。
3. 分配 `/ppt/notesSlides/notesSlideN.xml`，写 canonical notes root、group shape tree和唯一 body placeholder。
4. 创建 slide→notesSlide、notesSlide→slide、notesSlide→notesMaster 三条 internal relationship。
5. Body `txBody` 保留 canonical `bodyPr` / `lstStyle`，正文写一个 paragraph/run/text，使用 `xml:space="preserve"` 和 XML escaping。

Canonical notes master只在完全 absent state创建；它包含合法 notes root、group shape tree、color map、header/footer flags 和 notes style，不改变现有 slide master、layout 或 theme。

### String 编辑

已有安全 notes slide 时只替换唯一 body `txBody` 中的 direct paragraphs；body shape的 non-text bytes、properties、placeholder metadata、unknown sibling shapes、extensions、relationships、content type和其他 parts保持。若安全 shape tree没有 body placeholder，setter插入一个新 canonical body shape并使用所有 direct `cNvPr@id` 中下一个安全 ID。多个 body候选或 ambiguous IDs 拒绝。

相同 normalized plain text 是 byte/journal exact no-op。Whole-value replacement有意把 body 的 rich runs和multiple paragraphs折叠为 plain deterministic state；这是调用 plain-string setter 的明确语义，不影响 body shape 以外内容。

### 清除

`slide.notes = undefined` 在 absent state 是 exact no-op。存在唯一安全 notes slide 时，删除 slide relationship并用现有 owned-dependency garbage collection 删除不再被引用的 notes part及其 owned closure；shared notes master/theme保留。Ambiguous relationship或外部入边在 mutation 前拒绝，避免删除其他 slide仍引用的数据。

## 7. Slide lifecycle 与格式

- `addSlide()` 继续不物化 notes；首次 string assignment创建。
- `duplicateSlide()` 使用现有 owned dependency clone，复制 notes part并把 notesSlide→slide relationship retarget 到 duplicate；notes master/theme保持 shared。
- `moveSlide()` 不改变 notes relationship或文本。
- `deleteSlide()` 使用现有垃圾回收删除其 owned notes part；其他 slide notes和shared master保留。
- Outer transaction rollback恢复 parts、relationships、content types、journal、slide model identity和notes snapshot。
- `pptx/pptm/ppsx/ppsm/potx/potm` 六种格式共用相同 notes content type和relationship semantics；写出扩展名/主 presentation content type继续由现有 format profile决定。
- Notes与hidden、sections、title、shapes、layout/master/theme编辑正交。

## 8. 类型、聚合包与 packed surface

`SlideModel.notes` 与 `addNotes()` 从 `@pptx/model` 经 SDK 和单包入口自然导出，不增加 plugin或运行时 PptxGenJS dependency。Node和browser bundle使用同一 OOXML helper。

实际 npm tarball smoke覆盖：

- Node create/addNotes/edit/empty/clear/duplicate/write/reopen；
- browser bundle create/addNotes/write/reopen；
- declaration consumer验证getter为 `string | undefined`、setter接受 `string | undefined`、`addNotes(string)`返回 `SlideModel`；
- CLI/version和现有 packed examples不退化。

## 9. 文档与兼容边界

PptxGenJS parity matrix新增 `slide.addNotes(string)` 行，标记 plain create/read/edit/empty/clear/duplicate/reopen支持，并记录 native lazy materialization。README展示从零添加备注和编辑已有备注；changelog只声明plain speaker notes，不声称rich notes、notes page layout或comments完成。

PptxGenJS omitted/empty均物化empty notes part，而native absent/explicit empty分别保留为`undefined`/`''`；两者对合法 non-empty speaker text的最终用户语义对等。Adapter忠实读取PptxGenJS empty notes part为`''`。

## 10. 测试与验收

1. Helper fixtures覆盖absent、empty、plain、LF/CRLF/CR、XML escaping、multiple paragraphs/runs/br、alternate prefixes和unknown sibling shapes；reads exact no-op。
2. Unsafe fixtures覆盖duplicate/external/missing/wrong-content-type relationship、wrong root namespace、duplicate body placeholder、ambiguous txBody和invalid shape ID；getter为`undefined`，mutation零变化拒绝。
3. Model/SDK覆盖lazy create、non-empty→same no-op→edit→empty→undefined clear、invalid input、outer rollback、write/reopen和六种formats。
4. Lifecycle覆盖duplicate retarget、move、delete、shared notes master保留、其他 slide notes隔离，以及hidden/sections组合。
5. Missing notes master fixture覆盖canonical master/list/relationships创建、theme选择、ambiguous/no-theme zero-mutation拒绝。
6. PptxGenJS public conformance覆盖omitted、empty、plain、multiline和XML metacharacters；只比较public final OOXML和semantic snapshot。
7. Packed Node/browser/types smoke覆盖实际tarball；full typecheck/test、performance、build/pack、PowerPoint 2010 validator、package diff、LibreOffice/render/notes extraction和remote divergence全部通过。

每个可独立review的小项完成后单独commit并push；最终QA若无repository change，不创建空commit。
