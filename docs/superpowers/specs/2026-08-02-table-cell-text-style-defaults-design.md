# Table-Cell Text Style Defaults Design

日期：2026-08-02
状态：已批准实施

## 目标与范围

为 native `slide.addTable()` 增加 table-level 与 cell-level 的文本样式默认值，使 plain、rich、multi-run、multi-paragraph 和空段落 cell 能在创建时继承 font family、font size、bold、color 与 paragraph spacing，并让最终 direct OOXML 可由既有 `TableCell.richText`、`TableCell.text` 和 indexed editors 继续安全读取或编辑。

本小项只增加创建期默认值及其最终 direct state。它不增加 retained table-style metadata、effective theme-style resolver、单独的 table/cell font getter/editor、per-run indexed editor、merge/span、row/column CRUD、auto-page、repeated headers、content measurement 或 `tableToSlides`。

用户已授权实现方自主确定后续内容并连续推进，因此本设计按推荐方案直接批准实施，不设置额外等待确认点。

## 公共 API

```ts
export interface AddTableOptions {
  // existing fields remain unchanged
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly color?: RichTextColor;
  readonly spacing?: ParagraphSpacing;
}

export interface AddTableCellOptions {
  // existing fields remain unchanged
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly color?: RichTextColor;
  readonly spacing?: ParagraphSpacing;
}
```

Native naming stays aligned with the existing `RichTextRunStyle` and `AddTextOptions`: `fontFamily`, `fontSize`, `bold`, `color`, and `spacing`. PptxGenJS-shaped aliases such as `fontFace`, `paraSpaceBefore`, `paraSpaceAfter`, `lineSpacing`, and `lineSpacingMultiple` are not accepted by the native API; the adapter continues to import their final OOXML state.

Omitted and runtime-`undefined` fields do not change legacy bytes. `fontFamily` is a non-empty XML-safe string; `fontSize` uses the existing strict finite `1..4000` point contract and two-decimal canonicalization; `bold` is a strict boolean; `color` uses the existing strict sRGB/theme `RichTextColor`; `spacing` uses the existing strict partial `ParagraphSpacing` value.

## 方案选择

1. **在两个 option interfaces 增加五个 targeted top-level fields；采用。** 公开表面与现有 native text API 一致，字段优先级清楚，避免开放 hyperlink、underline、glow 等本小项不承诺的 run style，并能复用现有 rich-text renderer。
2. **增加 `textStyle?: RichTextRunStyle` 与独立 paragraph defaults。** 类型复用看似更短，但会立即承诺全部 run-local effects、hyperlink 和 transparency 的 outer-default 语义，扩大验证与兼容范围，因此不采用。
3. **直接复制 PptxGenJS 的 `fontFace` / paragraph spacing aliases。** 可减少迁移拼写，但会在 native API 中制造第二套命名、truthy fallback 与单位边界，和现有 `fontFamily` / `spacing` contract 冲突，因此不采用。

## 归一化与优先级

整个 table matrix 先完成 cell text 与 cell options 的 detached normalization，再解析 table options。每个 physical cell 最终只保存 resolved creation defaults；优先级固定为：

```text
table defaults → cell defaults → explicit paragraph/run fields
```

Font family、font size、bold 和 color 按字段覆盖。Cell 的显式 `bold: false` 必须覆盖 table `bold: true`；run 的显式 `bold: false` 必须覆盖 cell/table true。任何零值、false 或局部字段都不得经过 truthy fallback。

Spacing 按 `before`、`after`、`line` 三个子字段 overlay：cell partial spacing 只覆盖其提供的 table 子字段，paragraph partial spacing 再覆盖 resolved cell defaults；paragraph `spacing: false` 清除全部 outer spacing，paragraph `line: false` 只清除 inherited line spacing。所有 validation 在 relationship、shape ID、part bytes或 mutation journal 发生变化前完成，输入对象保持 detached。

创建完成后不保留 table/cell default metadata，也不在 OOXML extension 中写私有状态。`setCellRichText()` replacement 只使用新 value，不重新继承创建默认；`setCellText()` 继续保留安全 plain run 的现有 direct style，因此可保留已物化的默认值。

## OOXML 渲染

扩展既有 `RenderRichTextOptions`，增加 optional default font family、font size、bold、color 和 paragraph spacing。`renderTableCell()` 将 cell 的 resolved defaults 传入 `renderRichTextParagraphs()`；普通 text shape、placeholder 和既有 callers 在未传值时保持 byte-identical。

每个 non-empty 或 explicitly styled run 的最终 `a:rPr` 直接物化 resolved font face、`sz`、`b` 与 color。Run-local field 逐项覆盖 outer default；没有 run-local 或 outer value 时继续使用现有 theme font与 `tx1` canonical output。Explicit `bold: false` 写 `b="0"`，不折叠成 omission。

有 local 或 inherited hyperlink 的 run 若没有 explicit run color，不继承 table/cell color，并省略 direct main text fill，让 hyperlink color保持 theme-driven；explicit run color 仍优先并写 direct fill。这与 PptxGenJS 4.0.1 的合法 final state一致，但不复制其对其他 falsy fields 的覆盖缺陷。

Paragraph spacing 继续由 `resolveParagraphSpacing()` 与 `renderParagraphProperties()` 输出 schema-ordered `lnSpc`、`spcBef`、`spcAft`。Font family/font size 同时写入 table-cell `endParaRPr`，避免空 paragraph 回退到 Arial/18pt；bold/color 只在存在 direct run 时物化，和 PptxGenJS 的 empty-paragraph behavior 一致。没有任何新 table-level OOXML property 或 custom metadata。

## PptxGenJS 4.0.1 基线

真实 public `write()` 输出与 writer source 确认：

- `TableProps` / `TableCellProps` 都继承 `TextBaseProps`，但 table writer 实际只向 cells 传播 `align`、`bold`、`border`、`color`、`fill`、`fontFace`、`fontSize`、`margin`、`textDirection`、`underline` 和 `valign`；
- table `fontFace` / `fontSize` / `bold` / `color` 被物化到 plain 和 rich runs，cell options 再覆盖；
- cell `paraSpaceBefore` / `paraSpaceAfter` / exact or multiple line spacing 被物化到 paragraphs，table-level spacing fields 不会传播；
- hyperlink runs 跳过 outer color，但继续继承 outer font face、font size 和 bold；
- table-cell `endParaRPr` 特别物化 outer font face/font size；
- writer 使用 truthy fallback，因此 cell/run `bold: false` 会被 outer true 覆盖，且会修改 caller options。

Native table-level `spacing` 是与其他 native table defaults一致的严格扩展；它为每个 physical cell 物化 paragraph state。合法 PptxGenJS cell-spacing、font/size/bold/color、rich overrides、hyperlinks 和 empty paragraphs 必须可导入、读取、编辑和重开。Native 不复制 caller mutation、falsy coercion、重复 `pPr`、宽松 string/number coercion 或 table spacing propagation omission。

## 错误、事务与兼容边界

- Unknown、symbol、accessor、inherited 或 class-instance option fields 继续由 descriptor-safe readers 拒绝。
- Empty/invalid font family、non-finite/out-of-range font size、non-boolean bold、malformed color 和 empty/invalid spacing update 在创建前拒绝。
- Table/cell/paragraph spacing 的 partial overlay 必须先得到完整合法 state；不能因一个 cell 失败而留下部分 table、relationship 或 ID mutation。
- Omitted 与 explicit `undefined` 的 table/cell defaults 对 legacy single-line、rich、hyperlink、border/fill/margin/direction/fit/alignment output保持 byte-identical。
- Existing-deck reads continue to report direct run/paragraph state only；不推断 theme/table style effective values，不新增 ambiguous style ownership。

## 测试与验收

1. Internal TDD：table/cell normalization、partial spacing overlay、run/paragraph precedence、explicit false、detachment、strict invalid inputs、legacy byte parity、rich/empty paragraph rendering 与 hyperlink color suppression。
2. Model lifecycle：plain/rich/multi-paragraph snapshots、`setCellText()` style preservation、`setCellRichText()` no re-inheritance、same-value no-op、duplicate、rollback、six formats 与 write/reopen。
3. SDK/root/declaration：五个 fields 在 `AddTableOptions` 与 `AddTableCellOptions` 的 positive/negative TypeScript contract，Node/browser exports 和 generated declarations。
4. PptxGenJS adapter：合法 table/cell font/size/bold/color、cell spacing、rich overrides、empty paragraph、hyperlink final-state import/edit parity，以及 documented strict differences。
5. Package proof：actual tarball Node/TypeScript/browser/CLI/Inspector、真实 Chrome、PowerPoint 2010 validation、part inspection 和 zero console/page/network errors。
6. Final gates：focused Vitest、project typecheck/build、full Vitest with two workers、independent performance、docs review、commit、push 与 local/remote `0/0`。

## 后续边界

本小项完成 table/cell outer text style defaults。后续小项依次处理 merge/colspan/rowspan、row/column CRUD、auto-page/repeated headers、`tableToSlides` 与最终 peer/client audit。
