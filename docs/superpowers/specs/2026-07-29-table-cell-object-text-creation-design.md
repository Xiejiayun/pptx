# Plain-Text Table Cell Object Creation Design

## 目标与范围

本小项扩展 native `slide.addTable()`，让每个 cell 除了现有 string 外，还能接受严格的 plain-text object：`{ text: string }`。同一矩形 matrix 内可以混用 string 与 object；两种输入在 text 相同时必须生成完全相同的 table cell OOXML。

这是 PptxGenJS cell-object creation 的基础切片。它只建立公共输入类型、descriptor-safe object 读取和 text normalization，不同时加入 cell `options`、rich-text runs、hyperlinks、merge/colspan/rowspan、row/column insertion、auto-page 或重复表头。后续 option 按独立小项逐个扩展本对象类型。

## 公共 API

在 model/SDK 聚合导出中新增：

```ts
export interface AddTableCell {
  readonly text: string;
}

export type AddTableCellInput = string | AddTableCell;
```

`SlideModel.addTable()` 的 rows 参数改为：

```ts
rows: readonly (readonly AddTableCellInput[])[];
```

`AddTableCell` 将来可增加 optional `options`，但本小项不提前公开空 options 类型，也不接受 `{ text, options: {} }`。这样不会用一个过宽的空 object 类型伪装尚未实现的 cell styling。

## 方案选择

考虑三个方案：

1. 接受 string 或 exact `{ text: string }` ordinary object，归一化后继续使用现有 string renderer。采用此方案；输出稳定，公共类型可增量扩展。
2. 读取任意 object 的 `.text` 并忽略其余字段。它会调用 accessors、接受 class/inherited state，并静默丢弃用户以为生效的 options，不采用。
3. 一次实现 PptxGenJS 全部 cell options、rich text、merge 和 hyperlink。范围过大，无法作为一个可独立 review 和回滚的小项，不采用。

## 严格输入归一化

现有 outer rows 和 inner row 的 dense-array、descriptor-safe、非空、矩形规则保持不变。每个 cell 按以下规则读取：

- string 继续使用当前行为。
- object 必须非 null、非 array，prototype 只能是 `Object.prototype` 或 null。
- `Reflect.ownKeys()` 必须只得到一个 string key `text`；symbol、额外 key、缺失 key 都拒绝。
- `text` 必须是 own data property；accessor、继承 property 或 proxy 暴露的不完整 descriptor 都拒绝，不调用 getter。
- data value 必须是 string、不得包含 CR/LF、不得包含 XML 1.0 禁止字符；empty string 合法。
- normalized definition 只保存 detached string，不保留 caller object、row 或 matrix。调用方随后修改 object.text 不得改变 definition 或渲染结果。

错误继续使用 `TypeError`，并包含准确 `rowIndex,columnIndex`。字符串与 object text 共享同一 paragraph/XML validation helper，避免语义分叉。

## 渲染与兼容性

`NormalizedTableDefinition.rows` 继续是 detached `readonly string[][]`，`renderTableGraphicFrame()` 和 `renderTableCell()` 不需要新的 XML 分支。相同 text 的 string cell 和 object cell 必须产生 byte-identical `a:tc`，包括 paragraphs、default margins 和四边 no-fill borders。

表格 geometry、automatic/explicit row heights、column widths、stable identity、transaction rollback、extension-list insertion 和现有 cell editors 均不改变。创建后返回的 `TableModel.rows` 仍提供当前 snapshot shape；`AddTableCell` 是 creation input，不替代 `TableCell` read model。

## PptxGenJS 对等与差异

PptxGenJS 4.0.1 的 cell object 基本形态为 `{ text, options }`，其中 text 还可包含 rich runs。本小项支持其中 plain string `text` 的 object container，并允许与 string cells 混排；PptxGenJS object 的 `options` 和 rich-text arrays 仍严格拒绝，而不是静默忽略。

PptxGenJS 对非法对象常有 coercion 或晚期 runtime failure。本库保持现有 strict/deterministic contract：不读 inherited values，不调用 accessors，不 stringify 非 string text。差异记录为 intentional safety repair。

## 测试与验收

实现必须覆盖：

1. internal normalization：纯 object、string/object 混排、empty text、null-prototype object、detached mutation 和 equivalent byte output。
2. invalid matrix：missing/extra/symbol keys、accessor text、inherited text、array/date/class object、null、numeric/rich-array text、CR/LF 和 invalid XML character；accessor invocation count 为 zero。
3. public model：创建、stable identity、即时 cell editing、outer rollback 和 reopen 后 text/geometry 不变。
4. SDK：typed mixed matrix、duplicate isolation、write/reopen 和非目标 parts 稳定。
5. PptxGenJS 4.0.1：native `{ text }` objects 与其 public `{ text, options: {} }` basic output 在 cell text、geometry、margins 和 borders 上对等。
6. packed Node/browser/declaration/CLI smoke：`AddTableCell`、`AddTableCellInput` 可导入，mixed cells 可创建、编辑和重开。
7. TypeScript project references、focused/full tests、现有 performance gate。
8. PowerPoint 2010 validation 对 source/native reopen/PptxGenJS baseline 均 zero error/zero warning。
9. string-only 与 equivalent object-input package diff 为 zero changed decompressed parts；native write/reopen 同样稳定。
10. LibreOffice/Poppler 原图确认 object cells、empty cell、纵排与底部对齐正常，overflow checker 无非预期 overflow。

## 文档与发布表面

更新 changelog、API README、PptxGenJS compatibility baseline、package README 和 packed smoke。文档明确 native creation 现在接受 string 或 `{ text: string }`，但 object `options`、rich/multi-paragraph text、hyperlinks 和 merge 仍未支持。

## 非目标

本小项不增加 cell options、fill/border/margin/direction/fit/alignment creation properties、rich text、hyperlinks、merge/colspan/rowspan、row/column mutation、table styles、auto-page、repeated headers 或 content measurement。它不改变现有 table OOXML、editing API、`setTransform()`、`setColumnWidths()` 或 `setRowHeights()` 语义。
