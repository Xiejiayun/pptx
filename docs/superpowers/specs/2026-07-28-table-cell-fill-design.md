# Table Cell Fill Design

日期：2026-07-28
状态：已批准实施

## 目标与范围

为既有 table cell 增加 direct 填充的读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TableCellProps.fill` / `TableProps.fill` 最终写入的 `a:tc/a:tcPr` fill choice。本小项支持显式 no-fill、solid sRGB/theme color 和 direct alpha，并区分“显式无填充”与“没有 direct override”。

该能力只表达选中物理 cell 的 direct `tcPr` fill。它不解析 table style、theme 解析后的 effective color、conditional banding 或 master/layout 继承，不把 border 中的 `solidFill` / `noFill` 误认为 cell fill，也不修改 text/run fill。

本小项不增加 table creation、table-level fill default、gradient/pattern/picture/group fill 编辑、border、merge mutation、row/column sizing、rich cell text、auto paging 或布局计算。已完成的 cell text、direction、fit、vertical alignment 和 margins 与 fill 保持独立。

## 公共 API

复用现有 `RichTextColor` 的 strict sRGB/theme color value，不公开 raw OOXML token：

```ts
export type TableCellFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export interface TableCell {
  readonly text: string;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}

export class TableModel {
  setCellFill(
    rowIndex: number,
    columnIndex: number,
    value: TableCellFill | undefined,
  ): void;
}
```

用法：

```ts
table.rows[0]?.cells[0]?.fill; // direct detached snapshot or undefined

table.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'srgb', value: 'FF0000' },
});
table.setCellFill(0, 1, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
});
table.setCellFill(0, 2, { kind: 'none' }); // write direct a:noFill
table.setCellFill(0, 2, undefined);        // remove the direct fill choice
```

`transparency` 使用 `0..100` 百分比：0 完全不透明，100 完全透明，量化到最近 `0.001%`。字段缺省表示不写 direct alpha；显式 0 写 `alpha val="100000"`，因此两者保持可逆。`kind: 'none'` 写 direct `a:noFill`，`undefined` 删除 direct fill choice，从而与 table style/effective fallback 语义区分。

`rows` 继续返回 detached value snapshots。调用方修改旧 `fill` 或嵌套 `color` 对象不会修改 model。setter 继续使用零基 physical row/cell index，对 merged placeholder 不做 logical-span 折算。

## 方案选择

考虑过三个方案：

1. 直接复制 PptxGenJS `{ color, transparency, type }` value shape。迁移表面最像，但 `type: 'none'` 在 4.0.1 table runtime 中实际被折叠为 omission，无法区分 direct no-fill 和缺省；运行时还会接受 string fill 和越界 transparency，不适合成为新的双向契约。
2. 公开 discriminated `none | solid` value，solid 复用 `RichTextColor` 并使用 strict percentage transparency；采用此方案。它覆盖 PptxGenJS 公开 fill 能力，保留 direct no-fill/absence 差异，与现有颜色和透明度单位一致，并把未支持 fill 类型留给后续小项。
3. 立即把 `OoxmlColor` 和全部 gradient/pattern/blip/group fill 暴露给 table cell。它的 OOXML 覆盖更广，但会把本小项扩展为通用 fill codec 重构，与 PptxGenJS table fill 对等所需的最小范围不匹配。

用户已授权实现方按小项持续选择最佳方案，因此本设计按 strict native API、PptxGenJS 输出兼容、direct-state 可逆和最小可验证改动定稿。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()`、安装包类型声明和实际 slide XML 校准：

- `TableCellProps.fill` 和 `TableProps.fill` 声明为 `ShapeFillProps`。table-level fill 在生成时复制到未覆盖 cells，cell-level fill 优先；最终文件中没有独立 table-level fill 状态。
- table/cell 都省略 fill 时，`tcPr` 不写 direct cell fill child。`{ type: 'none' }` 和 `{ type: 'none', color: ... }` 也被折叠为 omission，不生成 direct `a:noFill`。
- `{ color: 'FF0000' }` 生成 `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`；SchemeColor 生成 `a:schemeClr`。
- `transparency: 25` 生成 `alpha val="75000"`；100 生成 zero alpha。显式 0 因 truthy 分支被折叠为没有 alpha child。fractional percentage 按最近整数 alpha 写入。
- deprecated `alpha` 仍按 transparency percentage 处理。runtime string fill 仍可生成 solid sRGB，但不在公开类型契约内。adapter conformance 只使用公开 typed shape。
- runtime 对越界 transparency 不做 strict validation：负值可生成大于 100000 的 alpha，大于 100 可生成负 alpha。本库 adapter 保留原 XML，strict getter 不伪造合法 snapshot，native setter 在 mutation 前拒绝。
- merged placeholder cells 保持独立 `<a:tcPr/>`，不复制可见起始 cell fill。本库继续按 physical cells 编辑。

adapter 只从公开输出的 OOXML 读取，不访问 `_slides` 等私有字段。它把 direct alpha 还原为 `100 - alpha / 1000` transparency，不推测调用方最初使用 table-level 还是 cell-level fill。

## OOXML 与 strict direct 语义

```xml
<a:tc>
  <a:txBody>...</a:txBody>
  <a:tcPr marL="91440" marR="91440" marT="45720" marB="45720">
    <a:lnL>...</a:lnL>
    <a:lnR>...</a:lnR>
    <a:lnT>...</a:lnT>
    <a:lnB>...</a:lnB>
    <a:solidFill>
      <a:schemeClr val="accent1">
        <a:alpha val="75000"/>
      </a:schemeClr>
    </a:solidFill>
  </a:tcPr>
</a:tc>
```

getter 首先要求 cell 恰好一个 direct `tcPr`，然后检查 `tcPr` 的 direct fill-choice children：`noFill`、`solidFill`、`gradFill`、`blipFill`、`pattFill`、`grpFill`。没有 choice 返回 `undefined`；多个 choice 或唯一 choice 属于未支持类型时返回 `undefined` 并保留 XML。这个 direct-child 边界保证 border 内的 fill 不会被读成 cell fill。

`noFill` 必须是唯一 supported choice，且除 namespace declaration 外没有 attribute 或 element child，才返回 `{ kind: 'none' }`。

`solidFill` 必须只含一个 direct `srgbClr` 或 `schemeClr`：

- `srgbClr@val` 必须是六位 hex，snapshot 统一转大写。
- `schemeClr@val` 必须属于现有 `RichTextColor` 支持的 theme tokens。
- color element 除 namespace declaration 外只允许必需 `val` attribute 和可选的单个 direct `alpha`；额外 attribute、alpha 之外的 color transform、重复 alpha、alpha 的 element child 或非 strict alpha 使整个 fill snapshot 返回 `undefined`，避免忽略会改变最终颜色的部分状态。
- `alpha@val` 必须是 `0..100000` strict decimal integer，以 `100 - value / 1000` 返回 transparency。absence 不伪造显式 zero。

missing/repeated direct `tcPr` 在 snapshot read 中返回 `undefined`，避免 malformed cell 阻断整张 table；setter 对同一结构抛 `ModelParseError`，因为无法安全选择 mutation target。

## 无损编辑与错误边界

setter 在 package mutation 前完整 normalize public value：

- `kind: 'none'` 不允许额外字段。
- `kind: 'solid'` 必须提供 strict `RichTextColor`；sRGB 允许可选 `#` 并规范化为六位大写 hex，scheme token 使用已有集合。
- transparency 必须是 finite `0..100`，并量化为 `Math.round((100 - value) * 1000)` alpha。
- null、boolean、string、array、未知 kind/key、缺少 color、非法 color、non-finite 或越界 transparency 在 mutation 前失败。

随后在 OPC transaction 内定位选中 physical cell，row/cell 不存在时抛 `RangeError`。selected direct `tcPr` 必须唯一，并且不得存在多个 direct fill choices；否则抛 `ModelParseError`。

更新只替换、插入或删除 selected direct fill choice：

- `solid` 写 canonical `solidFill`/color/optional alpha；`none` 写 canonical `noFill`；`undefined` 删除已有 direct fill choice。
- 已有唯一 gradient/pattern/blip/group fill 在无关 mutation 中原样保留；用户显式调用 `setCellFill()` 时可用 supported fill 替换或用 `undefined` 清除它。
- normalized semantic value 相同时 exact no-op，保留小写 sRGB、leading-zero alpha、quote style、whitespace 和 mutation journal。
- malformed 但唯一的 existing fill 可由合法值整体替换或清除；不局部修补其内部节点。
- 新 fill child 使用当前 `tcPr` 的 namespace prefix，插入在 direct `extLst` 之前；自闭合 `tcPr` 安全展开。

既有 margins、anchor、vert、overflow、borders/diagonal borders、cell3D、extensions、unknown attributes/children、text body、fit、text、paragraph/run、merge state、neighboring cells、relationships 和 `TableModel` identity 必须保持不变。`setCellText()`、direction、fit、vertical alignment、margins、transform、duplicate、write/reopen 等非 fill mutation 必须保留 fill。

失败与 outer transaction rollback 必须保持 exact part bytes、mutation journal、slide/table identity 和 fresh rows snapshots。

## 测试与发布门禁

1. model fixture 覆盖 direct noFill、opaque/explicit-zero/partial/full-alpha solid sRGB/theme fill，以及 absent、unsupported fill types、multiple choices、非法 color/alpha、namespaced/duplicate attributes 和 repeated/missing `tcPr`；getter 不产生 mutation。
2. `setCellFill()` 覆盖 solid sRGB/theme、transparency omitted/0/fractional/100、explicit none、clear、same-value no-op、unsupported/malformed replacement、self-closing/expanded `tcPr`、merged placeholder、invalid index/structure/input 和 outer rollback。
3. 交叉能力测试证明 text、direction、fit、vertical alignment、margins、transform、duplicate、write/reopen 保留 fill；只编辑一个 physical cell 不修改 borders、text fill、neighbor 或 relationships。
4. SDK 覆盖 detached nested snapshots、public transaction、invalid runtime values/coordinates、package isolation 和 stable identity。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted/type-none omission、table-level sRGB/theme/transparency materialization、cell override、zero/fractional/full transparency、deprecated alpha 和越界 runtime output；只使用公开输出。
6. API、npm README、changelog 与兼容矩阵明确 direct none/absence、strict native range、physical index、table-level materialization，不声称 table creation、effective style 或 advanced fill 已支持。
7. packed Node/browser/declaration smoke 覆盖 snapshot、sRGB/theme、alpha、none、clear 和 neighbor/cross-property isolation；typecheck、全仓 tests、独立 performance、actual tarball 与 CLI 全部通过。
8. 同源 native/baseline 表格文件覆盖 opaque sRGB、theme color、50% transparency、explicit noFill、clear 和 untouched neighbor；PowerPoint 2010 profile 为 0 error / 0 warning，package diff 为空，LibreOffice 无修复导出，逐图检查颜色/透明度/文字/边框/邻居，两份 overflow 通过且 raster hash 一致。
