# Hidden Slide State Design

日期：2026-07-30
状态：已批准（用户已授权实现方持续选择最佳方案并推进，不等待常规设计确认）

## 1. 目标与边界

为原生库增加 PptxGenJS 4.0.1 `slide.hidden` 的有效输出对等，并把同一能力扩展为已有 PPTX 的严格读取、无损编辑、复制保持和事务回滚。

本小项只拥有 slide part 唯一 PresentationML root 上 unqualified `show` attribute 所表达的隐藏状态。它不实现 slide number、speaker notes、custom show、kiosk mode、section collapse、transition advance、master visibility、placeholder visibility或放映范围。隐藏 slide 仍保留在 presentation slide 顺序、section membership、part graph 与公共 `slides` collection 中。

## 2. 公共 API

```ts
class SlideModel {
  get hidden(): boolean | undefined;
  set hidden(value: boolean);
}
```

从零创建使用现有组合 API：

```ts
const slide = document.addSlide();
slide.hidden = true;
```

不向 `AddSlideOptions` 增加 `hidden`。PptxGenJS 的公开模型同样在返回的 slide 上设置 `hidden`；复用 live slide property 能同时覆盖创建和 existing-deck editing，避免维护两个不同语义入口。

Getter 返回语义隐藏状态，而不是 raw `show` token：

- root 没有 direct unqualified `show` 时返回 `false`；
- `show="0"`、`show="false"`、`show="off"` 返回 `true`；
- `show="1"`、`show="true"`、`show="on"` 返回 `false`；
- root ownership、namespace、attribute 基数或 token 不安全时返回 `undefined`；
- 读取不得改变 bytes、journal、relationships、model identity 或任何 cache。

Setter 只接受 boolean。`true` 生成 canonical `show="0"`；`false` 删除 direct unqualified `show`，恢复默认 visible state。`undefined`、null、number、string、object、array、symbol 和 coercible input 都在 mutation 前抛 `TypeError`。

## 3. 方案取舍

考虑过三种方案：

1. 在 `SlideModel` 暴露 `hidden` live property；采用。它保持 PptxGenJS 命名，同时覆盖从零创建、existing-deck 读取与编辑。
2. 只增加 `addSlide({ hidden })`；拒绝。它不能编辑已有 slide，并会把一个持续可变的 slide state 固化为一次性 create option。
3. 暴露 OOXML-shaped `slide.show`；拒绝。其 boolean 方向与产品概念相反，也会迫使调用者理解 default/override 细节。

## 4. PptxGenJS 4.0.1 基线

只通过 public constructor、`addSlide()`、public `slide.hidden` 与 public `write()` 生成的真实输出确认：

- omitted 和 explicit `false` 都不写 slide root `show`；
- `true` 在 slide root 写 `show="0"`；
- truthy 非 boolean runtime value 也写 `show="0"`，因为 PptxGenJS 不验证 setter input；
- hidden 不改变 presentation `p:sldIdLst`、slide relationship、section membership、visible shapes 或 layout relationship；
- 生成文件通过当前 PowerPoint 2010 validator 的 0 error / 0 warning 门禁。

Native 对等合法 boolean final state，但不复制 truthy coercion。Adapter 继续只调用 PptxGenJS public `write()` 并导入 final OOXML，不读取 `_hidden`、`_slides` 或其他私有字段。

## 5. OOXML ownership 与严格读取

```xml
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  show="0">
  ...
</p:sld>
```

新增 `packages/model/src/slide-visibility.internal.ts`，集中负责：

```ts
readSlideHidden(xml: LosslessXmlDocument): boolean | undefined;
replaceSlideHidden(xml: LosslessXmlDocument, value: boolean): boolean;
```

Reader 必须建立以下 ownership：

- XML 只有一个 root；
- root local name 为 `sld`，其 lexical prefix 通过 in-scope namespace declaration 精确解析到 PresentationML namespace；
- root 至多一个 exact unqualified `show` attribute；qualified `x:show` 不是本能力拥有的 state，必须保留且不得误读；
- descendant、layout、master、presentation 或 extension 中的同名 attribute 不参与读取；
- duplicate namespace declaration、duplicate unqualified `show`、wrong root namespace、multiple roots 或未知 direct token 使 snapshot 为 `undefined`。

Malformed XML 继续由 lossless parser 抛出 parse error。合法 absence 是明确的 visible `false`，不能与 unsafe `undefined` 混淆。

## 6. 最小编辑与 no-op

`replaceSlideHidden()` 在产生任何 patch 前解析完整 root ownership。Unsafe root 或 duplicate `show` 抛 `ModelParseError('Slide visibility is not safely editable')`。

- 设置 `true`：没有 attribute 时在 root start tag 结束前插入 ` show="0"`；已有唯一 attribute 时只替换 value 为 `0`。
- 设置 `false`：没有 attribute 时 exact no-op；已有唯一 attribute 时删除 attribute 与其紧邻的水平空白。
- canonical `show="0"` 再设置 `true` 必须 serialise 为 byte-identical output；absence 再设置 `false` 也必须保持 byte/journal no-op。
- explicit setter 可以 repair 单一 unknown token：`true` replacement 为 `0`，`false` 删除；它不能 repair ambiguous root 或 duplicate attribute。
- 替换保留 attribute 名称、顺序、quote style 以外的 root bytes；删除或插入不改变 root namespace、其他 attributes、children、comments、whitespace、relationships 或 content type。

`SlideModel.hidden` 在现有 `OpcPackage.transaction()` 中解析与写回当前 slide part。Input normalization、ownership validation 和 final bytes comparison 发生在首次 package mutation 之前。Outer `document.transaction()` 抛错时 slide bytes、journal 与 stable model identity 完整恢复。

## 7. Slide lifecycle

- `addSlide()` 继续生成没有 `show` 的 visible slide；调用方可立即设置 `hidden = true`。
- `duplicateSlide()` 复制 source slide part，因此 visible、hidden 和合法 alternate lexical state 都保持；duplicate 不改变 source。
- `moveSlide()` 只改变 presentation order，不改变 slide root visibility。
- `deleteSlide()` 删除 slide part；visibility 不引入新的 dependency 或垃圾回收规则。
- section assignment、duplicate/delete/move membership 同步与 hidden 完全正交；hidden slide 可为 loose，也可属于任意 section。
- hidden slide 的 shapes、background、layout/master/theme chain、notes relationship 和 transition state 均按现有 lifecycle 处理。

## 8. 类型、聚合包与运行时

`SlideModel.hidden` 从 `@pptx/model` 经 SDK 和唯一聚合包自然导出，不新增根级 free function、plugin 或 dependency。Node 与 browser bundle 使用同一 lossless helper；packed declaration consumer 必须读取、设置并窄化 `boolean | undefined` snapshot。

实际 npm tarball smoke 覆盖：

- Node 创建 visible/hidden slides、编辑、duplicate、write/reopen；
- browser bundle 创建 hidden slide 并 write/reopen；
- declaration consumer 的 getter/setter 类型；
- existing CLI/version smoke 不退化。

## 9. 文档与兼容边界

PptxGenJS parity matrix 新增 `slide.hidden` 行，标记合法 boolean final state 的 create/read/edit/duplicate/reopen 支持，并记录 native strict invalid handling。README 展示创建与已有 deck 编辑；changelog 只声明 hidden state，不声称 custom shows、放映范围、slide numbers 或 speaker notes 完成。

Hidden state 通常不改变单页静态渲染内容。视觉 QA 的重点是 visible/hidden slide 的内容、尺寸与顺序不被编辑损坏；放映是否跳过 hidden slide 以 OOXML root state、PowerPoint validator 和可用客户端行为为主要证据，不把 LibreOffice PDF 导出页数作为唯一语义判据。

## 10. 测试与验收

1. Helper fixtures覆盖 absent、六种合法 token、alternate root prefix、foreign qualified `show`、other root attributes/children/comments/whitespace；snapshot 与 mutation均不产生无关变化。
2. Unsafe fixtures覆盖multiple/wrong-namespace root、duplicate namespace declaration、duplicate unqualified `show`、empty/unknown token；getter为`undefined`，ambiguous mutation零变化拒绝，single unknown token可显式repair。
3. Model/SDK覆盖visible→hidden→visible、canonical no-op、invalid input、outer rollback、model identity、write/reopen和全部六种presentation formats。
4. Lifecycle覆盖visible/hidden duplicate、move、delete，以及hidden slide与section membership组合；只允许目标 slide part或既有拓扑 parts变化。
5. PptxGenJS public conformance覆盖omitted、false、true和truthy invalid final output；adapter与native合法 state比较只使用public output。
6. Packed Node/browser/types smoke覆盖实际tarball，不依赖workspace-only import。
7. 全仓typecheck/test、5秒performance、build/pack smoke、PowerPoint 2010 validator、package diff、LibreOffice/render_slides/slides_test与远端同步全部通过。

每个可独立review的小项完成后单独commit并push；最终QA若无repository change，不创建空commit。
