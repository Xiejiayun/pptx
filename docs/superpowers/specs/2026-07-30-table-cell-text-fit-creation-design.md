# Table Cell Text-Fit Creation Design

日期：2026-07-30
状态：已批准实施

## 目标与范围

为原生 `slide.addTable()` 增加 strict table-cell text-fit 创建能力，使 plain-text cell object 可通过 `options.fit` 声明 `none`、`shrink` 或 `resize`。创建结果立即由现有 `TableCell.textFit` snapshot 读取，并可继续通过 `TableModel.setCellTextFit()` 编辑，从而补齐“已有 cell 可编辑但不能在创建输入中声明”的缺口。

本小项只增加 cell-level 创建，不增加 `AddTableOptions.fit`、table-level getter/editor、字体缩放计算、内容测量、table/row/cell 尺寸重算、rich/multi-paragraph cells、merge、hyperlink、style、auto-page 或 repeated header。PptxGenJS 4.0.1 没有 table fit API，因此本能力明确是 native creation extension，不虚构 PptxGenJS table option 对等。

用户已授权实现方持续选择最佳方案并逐项推进，不等待常规设计确认；本设计按现有 API 一致性、OOXML ownership、最小独立范围和可验证性定稿。

## 方案选择

考虑三种方案：

1. 只增加 `AddTableCellOptions.fit`；采用。它复用现有 strict 三值类型、direct snapshot/editor 和 body-properties codec，以最小独立改动闭合 cell creation → read → edit → reopen 生命周期。
2. 同时增加 table-level 与 cell-level fit defaults；不采用。PptxGenJS 没有对应层级，且 default propagation 是可单独 review 的后续能力，不应与本小项绑定。
3. 保持现状，要求创建后逐 cell 调用 `setCellTextFit()`；不采用。它虽然能产生相同最终 XML，但没有补齐 declarative table creation surface。

不新增同义类型或 renderer。公共创建字段沿用 PptxGenJS 普通文本框的名称 `fit`，snapshot 继续使用更明确的 `textFit`。

## 公共 API

复用现有类型：

```ts
export type TextBoxFit = 'none' | 'shrink' | 'resize';

export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly fit?: TextBoxFit;
  readonly margin?: TextBoxMarginInput;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}
```

`AddTableCellInput`、`TableCell.textFit` 和 `TableModel.setCellTextFit()` 的现有名称与类型保持不变。String cell 仍表示没有 direct fit 创建输入。`AddTableOptions` 本小项不增加 `fit`。

输入必须是 ordinary 或 null-prototype object 的 own data property；accessor、inherited value、symbol/extra key、array/exotic object 和非法 token 都在任何 package mutation 前拒绝。归一化结果立即与 caller 脱离，并保留 explicit `none`，以便 renderer 精确执行创建语义且不给未来 default propagation 制造歧义。

## 创建语义与 OOXML

Cell fit 只拥有该 physical cell 的 direct `a:tc/a:txBody/a:bodyPr` fit choice：

```xml
<!-- omitted / undefined / none -->
<a:bodyPr/>

<!-- shrink -->
<a:bodyPr><a:normAutofit/></a:bodyPr>

<!-- resize -->
<a:bodyPr><a:spAutoFit/></a:bodyPr>
```

创建不写 `<a:noAutofit/>`。`none`、omitted 与 runtime `undefined` 生成 byte-identical `bodyPr`，因此创建后的 direct snapshot 都是 `undefined`；这是 direct-state 模型，而不是 effective fit 计算。`shrink` 与 `resize` 分别立即 snapshot 为同名值。

实现复用 `normalizeTextBoxFit()` 与 `renderTextBoxFitChild()`。Internal normalized cell 增加 optional `textFit`；`renderTableCell()` 仅在 child 非空时把 self-closing `bodyPr` 展开。`bodyPr` 仍位于 `lstStyle` 和 paragraph 之前，fit choice 不写入 `tcPr`，也不改变 `tcPr` 上 margins、anchor、direction、borders 或 fill 的既有顺序。

创建后只保留 OOXML direct state，不保留 cell input metadata。`setCellTextFit(..., 'none')` 与 `undefined` 继续清除 direct fit choice；same-mode shrink/resize 继续保留 PowerPoint-calculated metadata。清除不会恢复创建输入，且 fit mutation 不改变 `tcPr@vert`、cell text、paragraph state、neighbor cells、grid/rows、relationships 或 table identity。

## PptxGenJS 4.0.1 边界

PptxGenJS 4.0.1 的 `TableCellProps` 与 `TableProps` 不声明 `fit`，runtime 透传的 `fit`、`autoFit` 和 `shrinkText` 均被忽略并生成 fit-less `<a:bodyPr/>`。现有 adapter conformance 已证明 omitted、none、shrink、resize、legacy flags 和冲突组合导入后全部为 `undefined`。

本小项保留并扩展该证据：

- PptxGenJS runtime ignored behavior 不变；adapter 仍只读取公开 `write()` 输出。
- Native omitted/undefined/none 与 PptxGenJS fit-less final state 相同。
- Native shrink/resize 是有意的 strict extension，分别产生合法 direct `normAutofit`/`spAutoFit`，不宣称 PptxGenJS parity。
- Native invalid value 抛 `TypeError` 且零 mutation；PptxGenJS 的未知 runtime property 仍被忽略。

## 实现边界

`packages/model/src/table-create.internal.ts` 负责：

- 将 `fit` 加入 cell option allowlist；
- 用 `normalizeTextBoxFit()` 归一化到 internal `textFit`；
- 用 `renderTextBoxFitChild()` 生成 direct `bodyPr` child；
- 保持 omitted/undefined/none 的原有 bytes。

`packages/model/src/slide.ts` 只扩展 `AddTableCellOptions` 的公开类型。现有 `table-cell-text-fit.internal.ts`、indexed editor、getter 和 `text-box-fit.internal.ts` codec 不重构；它们已覆盖 existing-deck strict read、lossless patch、same-mode no-op 与 schema ordering。

每个实现阶段只修改完成该阶段所需的文件。不得修改、删除、stage 或提交 `.pnpm-store/`。

## 测试与发布门禁

1. Internal normalization/render 覆盖 string、object omitted、empty options、runtime undefined、none、shrink、resize、null prototype、caller detachment，以及和 align/border/fill/margin/textDirection/valign 的组合。Omitted/undefined/none 必须 byte-identical；shrink/resize 必须只展开目标 `bodyPr` 并写 exact child。
2. Invalid runtime 覆盖 null、boolean、number、case/whitespace variant、unknown string、array、object、symbol、accessor、inherited/extra key；所有失败都不调用 getter，并在 public mutation 前保持 slide bytes、journal、shape count 和 identity 不变。
3. Model/SDK 覆盖 live snapshots `[undefined, 'shrink', 'resize']`、cell text/property edits、duplicate、rollback、write/reopen、packed declarations，以及 none/undefined clear 后不恢复创建输入。
4. Adapter test 保留 PptxGenJS table fit-like inputs 全部 ignored 的证据，并对照 native shrink/resize extension；不访问任何私有字段。
5. Actual npm tarball smoke 增加 `tableCellTextFitCreation: true`，同时保留现有 `tableCellTextFit: true`。
6. Changelog、API README、compatibility baseline 与 package README 明确 cell-level creation 已支持，table-level fit creation、字体测量和 layout recomputation 仍未支持。
7. Full TypeScript、Vitest、1000-part performance、Node/browser/types/CLI tarball smoke 全部通过。
8. 真实 native deck 覆盖 omitted/undefined/none/shrink/resize、mixed edit/reopen；PowerPoint 2010 profile 为 0 error / 0 warning，等价输入零 part diff，edit 只改 slide part，LibreOffice 无修复导出且 overflow checker 通过。视觉 QA 只验证合法渲染、文本完整和无裁切；最终 font scale 仍由客户端计算。
