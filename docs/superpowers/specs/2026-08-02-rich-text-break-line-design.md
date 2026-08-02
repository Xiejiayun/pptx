# Rich Text `breakLine` Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 rich-text run `options.breakLine` 对应的段落切分能力，并让同一输入语义覆盖从零创建与 existing-deck 编辑：

- `RichTextRun.breakLine?: boolean` 表示“当前 run 结束后开始新段落”；
- 同一 input paragraph 内带标记的 run 按顺序切成多个 canonical `RichTextParagraph`；
- 最后一个 run 的 `breakLine: true` 不创建尾部空段落；
- 空 run、连续标记、run style、paragraph properties、run hyperlink 和 `softBreakBefore` 均可组合；
- plain slide、layout、master、placeholder 与 declarative master 共用 `normalizeRichText()` 的同一 contract；
- `ShapeModel.richText` getter 继续返回 OOXML 的 canonical explicit paragraphs，不反向伪造 `breakLine`。

本小项不在 outer `AddTextOptions` 增加无效果字段，不接受字符串中的 CR/LF，不改变 `softBreakBefore` 的既有 direct soft-break contract，也不新增尚未公开的 rich-text table-cell API；table rich text继续属于advanced-table专项。本小项也不补齐其余 advanced line/effect/text/table、`tableToSlides`、output/runtime helpers 或 peer-range audit。完成本小项仍不声明完整 PptxGenJS parity。

## 2. PptxGenJS 4.0.1 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `TextBaseProps.breakLine?: boolean`、`slide.addText()` 与 `write()` 真实输出。Renderer 先把 rich-text objects 放入线性数组，再在当前 object 入组后检查 `textObj.options.breakLine`：

```ts
arrTexts.push(textObj);
if (textObj.options.breakLine && idx + 1 < arrTextObjects.length) {
  arrLines.push(arrTexts);
  arrTexts = [];
}
```

public output 探针覆盖普通、空、连续、末尾、false、runtime truthy、`softBreakBefore` 与 paragraph-property 组合，观察到：

- `A(breakLine: true), B` 输出两个 `a:p`；
- 单一或最后一个 run 的 `breakLine: true` 只输出当前 paragraph，不创建尾部空 `a:p`；
- 空 run 上的非末尾标记会创建空 paragraph；两个连续空标记会创建两个连续空 paragraph；
- omitted、`undefined`、`false` 和 numeric zero 不切分；runtime truthy string 会切分，但不属于 typed public contract；
- 切分后的每个 paragraph 使用该 line 首 run 的 PptxGenJS paragraph options；
- `softBreakBefore` 只在 PptxGenJS line 内非首 run 写 `a:br`，成为新 paragraph 首 run 时不写；
- 每个 run 的 text/style/hyperlink 仍属于切分后相同的 run，paragraph boundary 不产生额外 text run。

Native 对合法 boolean input 比较相同 paragraph boundary、run 顺序、visible text 与 hyperlink final semantic。Native 不复制 truthy coercion、prototype读取或 accessor执行；Native paragraph properties 已是显式 owner，因此拆分出的各段复制原 input paragraph properties，而不是采用 PptxGenJS run-local paragraph-option覆盖规则。

## 3. 方案比较

### 方案 A：renderer 遇到 `breakLine` 时直接关闭并重开 `a:p`

初看改动最少，但 paragraph properties、run hyperlink relationship matrix、owner defaults 和 existing-deck replacement 都必须在输出阶段重新索引。归一化结果仍保留一个 paragraph，事务前验证、相等判断与 getter canonical shape 会互相不一致。

### 方案 B：把 `breakLine` 写成 `a:br`

可以复用 `softBreakBefore`，但 `a:br` 是同一 paragraph 内的软换行，不能承载 bullet、alignment、spacing、level 或 tab-stop 段落边界。该方案与 PptxGenJS 输出及 OOXML 语义均不等价。

### 方案 C：归一化阶段拆成 canonical paragraphs（采用）

`normalizeRichText()` 先严格验证全部 input paragraphs/runs，再按标记切分并移除 transient `breakLine`。下游 renderer、relationship allocation、getter、setter与six-format writer只消费现有canonical paragraph model。该方案把兼容语义集中在单一入口，最大限度复用既有正确性边界。

## 4. 公共 API 与 canonical snapshot

```ts
export interface RichTextRun {
  readonly text: string;
  readonly style?: RichTextRunStyle;
  readonly softBreakBefore?: boolean;
  readonly breakLine?: boolean;
}
```

典型调用：

```ts
const shape = slide.addRichText([{
  align: 'center',
  runs: [
    { text: 'First', breakLine: true, style: { bold: true } },
    { text: 'Second' },
  ],
}]);

shape.richText;
// [
//   { align: 'center', runs: [{ text: 'First', style: { bold: true } }] },
//   { align: 'center', runs: [{ text: 'Second' }] },
// ]
```

`breakLine` 是 setter/creator input convenience，不是可从 OOXML 唯一恢复的持久状态。两个显式 input paragraphs 与一个由 `breakLine` 切出的 input paragraph 会产生相同 canonical OOXML，因此 getter 只返回显式 paragraphs，不猜测原调用形式。重新把 getter snapshot 设回 shape 必须 bytes/semantic stable。

## 5. 严格验证与错误语义

- run 仍必须是 ordinary object，symbol/unknown key、prototype object、class instance和accessor data均拒绝；
- own `breakLine` data property 为 omitted/`undefined`/`false` 时不切分，为 `true` 时切分；
- defined value 必须是 primitive boolean；string、number、null、object、symbol、boxed boolean 和 function全部拒绝；
- 验证覆盖全部 paragraphs、runs、styles、paragraph properties 和 hyperlink targets后，才允许 relationship、part bytes、shape ID/order、journal 或 live cache变化；
- invalid input 使用 `TypeError`；无效 internal-slide hyperlink继续使用既有 `RangeError`，两者均保持 outer transaction exact zero mutation；
- 输入对象和数组只读，不删除或改写调用方的 `breakLine`。

PptxGenJS runtime 会对 truthy string/object执行切分；Native 的 strict boolean 是 intentional correction，与项目现有 descriptor-safe input contract一致。

## 6. 拆分算法

对每个已验证 input paragraph：

1. 先一次性归一化 paragraph properties 和全部 runs；
2. 建立当前 segment；每个 run 去掉 transient `breakLine` 后加入当前 segment；
3. 若当前 run 为非末尾且 `breakLine === true`，提交当前 segment并开启新 segment；
4. 遍历结束后始终提交最后一个 segment；原 `runs: []` 仍得到一个空 paragraph；
5. 每个 segment复制同一份 normalized paragraph properties，runs数组保持独立；
6. flatten所有 input paragraphs的segments，保持原 paragraph 和 run 顺序。

示例：

```text
[A!, B!, C]       -> [A] [B] [C]
[empty!, empty!, C] -> [] [] [C]
[A, B!]           -> [A, B]
[]                -> []
```

这里 `!` 表示 `breakLine: true`，`[]` 表示合法空 paragraph。最后一个 run 的标记被消费但不产生尾段，确保与 PptxGenJS loop-exhaustion guard一致。

## 7. Paragraph properties 与 soft break 组合

Native 的 `align`、`rtl`、margins、indent、bullet、level、spacing 和 tabStops属于 `RichTextParagraph`，切分后的每个 segment复制全部 normalized direct properties。outer text defaults仍由现有 renderer对每个 segment解析，不写入 input snapshot。

`softBreakBefore` 与 `breakLine` 保持正交：

- `breakLine` 只决定当前 run 后的 paragraph boundary；
- `softBreakBefore` 仍附着于它原本所在的 run，并由现有 renderer在该 run 前写 `a:br`；
- 如果前一 run 切段使带 `softBreakBefore` 的 run成为新 paragraph首 run，Native仍保留显式 soft break，不静默丢弃调用方 direct intent；
- 这是 Native 已公开的可逆 soft-break语义，相比 PptxGenJS“paragraph首 run忽略 soft break”的 renderer restriction更严格；不影响单独 `breakLine` 的 paragraph parity。

连续或末尾 `a:br` 的 existing-deck getter继续用空 text run + `softBreakBefore` 表达，永不转换为 `breakLine`。

## 8. Hyperlink、owner 与 lifecycle

拆分发生在 `prepareRichTextRunHyperlinks()` 和 relationship创建之前，因此 relationship ID matrix天然使用 canonical `[paragraphIndex][runIndex]`：

- URL/internal-slide、outer inherited hyperlink、per-run override和`false` suppression在拆分后保持目标；
- empty run、连续切分和同目标重复关系不发生错位、泄漏或错误复用；
- setter可在现有 shape上新增/移除/reorder boundaries，过期 relationship按现有 reference-count cleanup删除；
- plain slide、layout、master、placeholder source/population和declarative master共用归一化结果；
- duplicate保留exact generated OOXML，move不改变shape bytes，rollback恢复part bytes/order/relationships/journal/live identity；
- 六格式write/reopen只保留canonical paragraphs，不保留transient marker，这是预期的语义归一化。

Placeholder source与population call的rich text仍遵守现有owner选择规则；`breakLine`不修改placeholder identity、geometry、`isTextBox`、transform、styles或text-body layout options。

## 9. 验证策略

### Model 与 public owners

- omitted/undefined/false/true、middle/trailing、empty/consecutive和多 input paragraphs；
- 全 paragraph properties复制、outer defaults、style与`softBreakBefore`组合；
- URL/internal-slide/outer/per-run hyperlink relationship重索引、cleanup和rollback；
- immediate getter、setter、same-semantic no-op、stable model identity；
- slide/layout/master/placeholder/declarative owners；
- duplicate、move、六格式write/reopen；
- strict invalid values、unknown/accessor/symbol/prototype/class输入零 mutation；
- root declarations和compile-time invalid cases。

### PptxGenJS conformance

- public middle/trailing/empty/consecutive/false输出；
- paragraph count、run grouping、visible text和relationship target比较；
- paragraph options、soft-break组合和最后一个run guard；
- 锁定upstream runtime truthy coercion，并证明native strict rejection是intentional correction。

### Release 与兼容性

- focused/full Vitest、performance、TypeScript checks和build；
- actual tarball Node/types/browser/CLI 与真实Chrome create/read/edit/reopen；
- PowerPoint 2010 validator、exact part read与mutation isolation；
- LibreOffice round-trip保留paragraph boundary、空段落、soft break和hyperlink；
- representative rich/placeholder视觉与overflow检查；
- README、API、compatibility、progress和changelog收尾。

## 10. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. `RichTextRun.breakLine`在全部rich-text owners中产生正确canonical paragraphs；
2. trailing、empty、consecutive、paragraph properties、soft break与hyperlink组合均有永久测试；
3. getter不伪造transient marker，setter/create/edit/reopen保持canonical semantic；
4. invalid input在任何package mutation前严格拒绝，lifecycle与relationship isolation通过；
5. PptxGenJS public output、六格式、packed Node/types/browser/CLI、validator、LibreOffice与visual QA全部通过；
6. 文档把`breakLine`移入支持项，并明确其余full-parity工作仍待继续。
