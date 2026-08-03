# Table-Cell Rich Text Design

日期：2026-08-02
状态：已批准实施

## 目标与范围

为 native table physical cell 增加完整的 structured rich/multi-paragraph text 创建、读取和 whole-replacement 编辑，复用现有 `RichTextParagraph` / `RichTextRun` / `RichTextRunStyle` value model。该能力必须覆盖本库已经支持的 paragraph properties、run styles、soft breaks、`breakLine` paragraph splitting，以及 URL/internal-slide run hyperlink；不得另造一次性的 table-only rich-text 类型。

本小项只拥有 cell 的唯一 direct `a:txBody` 中的 direct paragraphs/runs。它不增加 merge/span、row/column CRUD、auto-page、repeated headers、content measurement、row-height recomputation、table-level text style defaults 或 `tableToSlides`。

用户已授权实现方自主确定后续内容并连续推进，因此本设计按推荐方案直接批准实施，不设置额外等待确认点。

## 公共 API

```ts
export interface AddTableCell {
  readonly text: string | readonly RichTextParagraph[];
  readonly options?: AddTableCellOptions;
}

export interface TableCell {
  readonly text: string;
  readonly richText: readonly RichTextParagraph[];
  // existing borders/fill/hyperlink/alignment/margins/direction/fit remain unchanged
}

export class TableModel {
  setCellRichText(
    rowIndex: number,
    columnIndex: number,
    value: readonly RichTextParagraph[],
  ): void;
}
```

Single-line string cells and `{ text: string }` remain byte-identical for equivalent input. String CR/LF is normalized to `\n` and creates direct paragraphs with empty-line preservation, matching the existing plain text-shape contract. Rich input uses `{ text: RichTextParagraph[] }`; a bare array is not a cell. `rows[].cells[].richText` is a detached structured snapshot; `rows[].cells[].text` becomes a paragraph-aware plain projection with `\n` between direct paragraphs and for direct soft breaks. Existing single-paragraph callers observe no value change.

`TableCell` remains an immutable snapshot, not a live cell model. Indexed edits use zero-based physical row/cell coordinates, consistent with every existing `TableModel.setCell*()` method.

## 方案选择

1. **复用完整 `RichTextParagraph[]` 创建、snapshot 与 replacement；采用。** 它与 text-shape rich text 共用稳定 value model、normalizer、reader、renderer、hyperlink semantics 和 validation，直接满足旧 table-cell alignment 设计预留的最终表面。
2. **先只允许 string 中的 CR/LF。** 改动较少，但 paragraph/run styles、structured reading 和 hyperlink ownership 仍需第二套 API，且会先锁定不完整的 collapsed-text 语义，因此不采用。
3. **新增不含 hyperlink 的 table-only rich types。** 初期 relationship 工作较少，但会复制全部 paragraph/run style 类型并制造后续迁移，违背现有跨 owner value-model 设计，因此不采用。

## 创建语义

`normalizeTableDefinition()` 在任何 relationship、shape ID 或 slide mutation 前完成整个矩阵的 rich-text normalization。String 先完成 XML-safe validation 与 CR/LF normalization，再按 `\n` 拆成 plain paragraphs；没有换行的 string 继续走现有 canonical single-paragraph 路径并保持 bytes。Structured styles、paragraph properties 和 run-local hyperlinks 使用显式 `RichTextParagraph[]`。

每个 rich cell 保存 normalized paragraphs，而不是提前 flatten 成 string。`renderTableCell()` 直接调用既有 `renderRichTextParagraphs()`：

- cell/table `align` 继续作为缺省 paragraph alignment，显式 paragraph `align` 优先；
- `AddTableCellOptions.hyperlink` 作为 cell default hyperlink 应用于没有 explicit run override 的 runs；
- `RichTextRunStyle.hyperlink` 创建独立 run relationship，`false` 抑制 cell default；
- cell default URL/slide relationship 在同一 cell 的 inherited runs 之间共享，explicit run links 各自拥有 relationship；
- omitted/empty tooltip、explicit underline、soft break、`breakLine`、paragraph/run styles 与 text-shape renderer 保持同一语义。

所有 internal slide targets 必须在 mutation 前解析。任一 cell、paragraph、run、style、hyperlink 或 geometry 输入失败时，整个 table 创建保持零 observable mutation，caller input 不得被修改。

## Structured Snapshot

新增窄 `table-cell-rich-text.internal.ts` boundary。Reader 只在 cell 恰好有一个 direct namespace-correct `txBody` 时调用现有 `readRichText()`；missing/repeated/ambiguous direct text body 返回空 snapshot，不扫描 descendant impostor。Paragraph/run 内容和受支持 direct properties 复用现有 tolerant rich-text reader，包括 fields、soft breaks、run-local hyperlinks 与 paragraph styles。

Snapshot 每次从当前 OOXML 重新构造，因此与 live `TableModel` 保持同步且与 caller input 脱离。Plain `text` 只遍历 direct paragraphs，并按 document order 读取 run/field text 与 soft breaks，不扫描 cell 外 descendant。PptxGenJS 或其他客户端的 unsupported/extra run properties 在 unrelated edits 中继续无损保留；structured snapshot 只承诺当前 `RichTextParagraph` 可表达的 direct state。

现有 scalar `TableCell.hyperlink` 和 `horizontalAlignment` reader 继续保持 strict single-paragraph/single-run 边界。Rich cells 通过 `richText` 查看逐 run hyperlink 与逐 paragraph alignment，不把 heterogeneous state 错压成一个 scalar。

## Whole-Replacement Editing

`setCellRichText()` 先调用 `normalizeRichText()` 并解析全部 internal targets，再进入一个 OPC transaction。Transaction 内按 table shape ID 与 physical indexes 解析一个 exact cell，要求 cell 恰好一个 direct `txBody` 且至少一个 direct paragraph；不安全 ownership 抛 `ModelParseError`。

Editor 复用 `readRichTextState()`、`richTextParagraphsEqual()` 与 `replaceRichText()`：

- 相同 normalized value 是 exact slide bytes / relationships / ZIP state / mutation-journal no-op；
- replacement 只替换 direct paragraphs，保留 `bodyPr`、`lstStyle`、`tcPr`、geometry、name、neighbor cells 与 table identity；
- current run 与 requested run 在同一 paragraph/run position 且 target 相同，复用 relationship ID；
- unique relationship target change 原位更新；shared ID target change clone-on-write；
- removed/replaced links 仅在 slide 中最后引用消失后 GC；
- tooltip/style/text-only change 不分配无关 relationship；
- output 使用 canonical rich-text paragraphs，但不重写 text body 或 cell 的非 owned children。

`setCellText()` 保持 plain single-paragraph/single-run local text replacement 和现有 hyperlink/style preservation。为避免 silent partial edit，它在 rich/multi-paragraph/multi-run cell 上改为抛 `ModelParseError`，并要求调用 `setCellRichText()`；合法 plain cell 的输出与错误 contract 不变。

## Relationship 与 Slide 生命周期

Rich cell links 使用现有 drawing hyperlink decoder、relationship target equality、reference counter、relative slide target 和 slide lifecycle codecs。必须覆盖：

- native-created independent explicit run links 与 shared cell-default link；
- imported shared IDs 的 same-target reuse 和 target-change COW；
- cell/shape 间共享 relationship 的 reference-safe clear/replace；
- target 前 insert/delete/reorder 后 getter ordinal 更新；
- duplicate self-link retarget 到 duplicate；
- target deletion 清理 selected source/duplicate 的 incoming run clicks；
- all six formats、write/reopen、outer rollback 和 injected add/update/remove/setXml failure isolation。

Relationship namespace 只在需要新增 direct run click 且当前作用域没有有效声明时补齐。Rich replacement 不创建 graphic-frame click，不改变 existing cell-level scalar hyperlink editor 的 strict ownership。

## PptxGenJS 4.0.1 基线

PptxGenJS 4.0.1 的 public `TableCell.text` 声明为 `string | TableCell[]`，运行时接受 `{ text: [{ text, options }, ...] }`。公开 `write()` 实测确认：

- `breakLine`、CR/LF、alignment change 和 bullet 可拆分 table-cell paragraphs；
- run options 写 bold/italic/font/fontSize/color/underline 等 direct style；
- cell `options.hyperlink` relationship 由未覆盖 runs 共享，run-local hyperlink 分配独立 relationship；
- omitted tooltip 被物化 empty，并写额外 hyperlink compatibility attributes；
- writer 会向 caller cell/run options 注入 `_rId`、inherited options 和内部字段，并可能在一个 paragraph 内重复输出 `pPr`。

Native 对合法可表达 final state 做 semantic import/read/edit/reopen conformance，不复制 caller mutation、truthy fallback、string/number coercion、duplicate paragraph-properties 或 dangling link 缺陷。Native rich input 使用显式 `RichTextParagraph[]`，而不是 PptxGenJS 的 recursive `TableCell[]` 类型。

## 错误与事务边界

- Rich input 必须是 non-empty paragraph array；每段 `runs` 必须是 array，所有 paragraph/run/style keys 继续使用现有 strict validation。
- Bare array cell、unknown/accessor/symbol/inherited fields、invalid style value、missing/both hyperlink target 和 dangling slide target 均在 mutation 前拒绝；string CR/LF 按 plain paragraphs 合法归一化。
- 不存在的 physical row/cell 抛 `RangeError`；ambiguous/missing text-body ownership 或 plain editor 遇到 rich structure 抛带 slide part URI 的 `ModelParseError`。
- Same-value/failed edits 不得改变 part bytes、relationships、ZIP directory/file dates、model identity 或 mutation journal。

## 测试与验收

1. Model/codec tests：single-line string byte parity、CR/LF normalization/empty-line preservation、paragraph-aware plain projection、rich normalize/render/read、multiple paragraphs/runs、empty runs、soft breaks、`breakLine`、全部既有 paragraph/run style families 与 detached input/snapshot。
2. Snapshot/editor tests：exact ownership、same-value no-op、whole replacement、plain `setCellText()` safety、neighbor/bodyPr/lstStyle/tcPr/geometry preservation、invalid structures/inputs/indexes 与 outer rollback。
3. Hyperlink tests：cell default inheritance、false suppression、explicit URL/internal link、ID reuse、unique update、shared COW、GC、self-link、reorder、duplicate、target delete、six formats、write/reopen 与 injected failures。
4. SDK/root/declaration tests：new union、readonly `richText` 与 `setCellRichText()` positive/negative TypeScript cases。
5. PptxGenJS adapter tests：public rich-cell output、paragraph splitting、styles、cell/run hyperlinks、input mutation divergence 与 canonical native output。
6. Package proof：actual tarball Node/TypeScript/browser conditional export/CLI/Inspector、真实 Chrome、PowerPoint 2010 validation、part/relationship inspection 和 zero console/page/network errors。
7. Final gates：focused Vitest、project typecheck/build、full Vitest with two workers、independent performance、docs review、commit、push 与 local/remote `0/0`。

## 后续边界

本小项完成 structured rich/multi-paragraph cell content，不增加 `AddTableOptions` / `AddTableCellOptions` 的 outer font-family/font-size/bold/color/paragraph-spacing defaults，也不增加 per-run indexed live editor。后续小项依次处理 table/cell text style defaults、merge/colspan/rowspan、row/column CRUD、auto-page/repeated headers、`tableToSlides` 与最终 peer/client audit。
