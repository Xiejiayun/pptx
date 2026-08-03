# Table Row and Column CRUD Design

日期：2026-08-02
状态：已批准实施

## 目标与范围

为已有 native 或导入 PPTX table 增加严格的 physical row/column 插入与删除。操作必须同步 `a:tblGrid/a:gridCol`、direct `a:tr`、每行 direct `a:tc`、merge topology、row/column size 与 graphic-frame transform；所有幸存 cell 的内容、样式、relationships、unknown attributes/children 和原始字节尽量保持不变。插入的 cell 使用可被现有 `setCellText()` / `setCellRichText()` / style editors 立即编辑的 canonical empty state。

本小项完成 table structural editing，不增加 logical-cell insert/delete overload、retained table creation defaults、style-template inference、row/column object identity、auto-page、repeated headers、content measurement、layout recomputation 或 `tableToSlides`。用户已授权实现方自主决定后续内容并连续推进，因此本设计按推荐方案直接批准，不设置额外等待确认点。

## 公共 API

```ts
export interface InsertTableRowsOptions {
  readonly count?: number;
  readonly rowHeights?: number | readonly number[];
}

export interface InsertTableColumnsOptions {
  readonly count?: number;
  readonly columnWidths?: number | readonly number[];
}

export class TableModel {
  insertRows(rowIndex: number, options?: InsertTableRowsOptions): void;
  deleteRows(rowIndex: number, count?: number): void;
  insertColumns(
    columnIndex: number,
    options?: InsertTableColumnsOptions,
  ): void;
  deleteColumns(columnIndex: number, count?: number): void;
}
```

所有 indexes 都是零基 physical coordinates，与现有 `rows[].cells[]`、merge snapshot 和 indexed cell editors 一致。Insert index 允许等于当前 row/column count，表示 append；delete index 必须指向现有 item。`count` omitted 时为 `1`，只接受 positive safe integer。删除不得移除最后一行或最后一列，插入后的 physical cell 总量不得超过 1,000,000。

`rowHeights` 接受 non-negative safe EMU scalar 或长度等于 `count` 的 dense readonly array；0 表示 automatic height。`columnWidths` 接受 positive safe EMU scalar 或 exact-length array。省略尺寸时，insert-at-middle 复制当前 insertion index 的 direct size，append 复制最后一个 direct size。输入必须是 descriptor-safe ordinary/null-prototype object；未知、accessor、symbol、inherited 或稀疏字段在 mutation 前拒绝。

Insert API 不接受 cell content。新 cells 是 canonical empty plain single-run cells，之后由现有 text/rich-text/hyperlink/alignment/margin/border/fill/direction/fit editors 填充。这样 structural mutation 只有一套职责，也避免把 creation-only table defaults错误地重新应用到 existing direct state。

## 方案选择

1. **Source-span physical splice；采用。** 解析一个严格 direct structure，只插入/删除目标 `gridCol`、`tr`、`tc` 和受影响 merge/transform tokens。幸存 nodes 继续使用原 source slices，符合 lossless editor contract。
2. **从 `TableModel.rows` semantic snapshot 重建整张 table。** 实现较短，但 snapshot 不含全部 unknown XML、lexical form 和 relationship ownership，重建会丢失 opaque state，因此不采用。
3. **引入长期驻留的 logical table matrix 并统一 write-back。** 后续分页可能受益，但会改变当前 live OOXML kernel 的 identity、transaction 和 byte-preservation边界；当前 CRUD 不需要该架构，故不采用。

## 严格结构边界

新增单一 internal table-structure boundary，要求：

- frame 恰有 direct `graphic/graphicData/tbl`、一个 direct `tblGrid` 和一个 direct `xfrm/ext`；
- grid 至少一个 direct `gridCol`，每个有唯一 unqualified positive-safe `w`；
- table 至少一个 direct `tr`，每个有唯一 unqualified non-negative-safe `h`；
- 每行 direct `tc` 数恰等于 grid count；foreign-namespace lookalikes 不计入；
- merge topology 由既有 all-or-nothing reader 完整识别；malformed/ambiguous merge 不允许 structural edit；
- transform `cx` / `cy` 是唯一 unqualified non-negative safe integer；column total 必须可安全求和；
- direct row/grid/cell ownership允许保留 unknown attributes、namespace declarations、opaque children 与合法 `extLst`，但 repeated/missing owned structure拒绝。

Structure reader 返回原 grid/row/cell elements、widths/heights、transform attributes、recognized merge regions 和所有待删除 cell subtree 中 exact non-empty relationship-ID attributes。它不推断 effective table styles、rendered automatic height 或 theme state。

## Canonical 新结构

新增 row 使用 `<a:tr h="H">`，包含当前 column count 个 canonical empty cells；新增 column 为每个现有 row 插入一个 canonical empty cell，并在唯一 direct grid 插入对应 `<a:gridCol w="W"/>`。Empty cell 复用 table creation renderer 的无 options、empty string final state，确保：

- 有唯一 direct `txBody/bodyPr/lstStyle/p/r/t`，`setCellText()` 可立即使用；
- 不创建 hyperlink relationship，不继承或猜测相邻 cell 的 direct style；
- 不携带 merge tokens，除非该 cell 因插入点位于既有 merge 内而成为 continuation；
- canonical output 在 Node/browser、write/reopen 和 PowerPoint 2010 下合法。

插入 row/column 不复制相邻 cell content、hyperlink、opaque XML 或 style。调用方若需要相同样式，可在插入后使用现有 public cell/table editors 显式设置，避免隐式模板选择造成不可预测的 mixed-state行为。

## Merge-Aware 插入

先把所有 recognized regions 映射到 insertion 后的 physical coordinates：

- insertion index 小于或等于 region anchor coordinate 时，整个 region 沿该 axis 后移；在 anchor coordinate 插入被定义为“在 region 前插入”，不扩展 region；
- insertion index 严格位于 region anchor 与 region end 之间时，region span 增加 `count`；每个插入 physical member成为 canonical empty continuation；
- insertion index 等于 region end 或位于其后时，region 不扩展；后方 regions 仅按普通 index shift；
- 同一 insertion 可扩展多个在另一 axis 上互不重叠的 regions；最终 topology 必须仍为 bounds-safe、non-overlapping rectangles。

未受几何影响的 merge region 保留原 cell start-tag bytes。发生 shift 但其 cell nodes整体移动的 region 不需要重写 tokens。只有 span 扩展的 region 才 canonicalize其 anchor/top/left affected merge attributes，并为新 cells写必要的 `rowSpan/gridSpan/vMerge/hMerge`。

## Merge-Aware 删除

删除 axis range 后，每个 recognized region 独立投影：

- 与 range 不相交的 region 保持原 tokens；位于 range 后的 cell nodes随结构自然前移；
- 部分相交时，region span 按幸存 physical members 收缩；2D region 可退化为 horizontal 或 vertical region；
- 如果原 anchor 被删除，最上/最左的首个幸存 cell成为新 anchor。原 anchor content随被删除 cell 一起删除，新的 anchor保留自己的原 hidden content/style/opaque XML并重新可见；
- 收缩到 `1 × 1` 时清除该 survivor 的全部 merge tokens；沿被删 axis 无任何 member 幸存时，整个 region 消失；
- 删除完成后重新验证所有 regions bounds/non-overlap 与每个 physical member的 exact semantic role。

只有几何变化或 anchor promotion 涉及的 survivor start tags允许重写四类 merge attributes。其他 survivor cell bodies、styles、relationships、unknown attributes/children 和 whitespace 保持原 source slice。

## 尺寸与 Transform

Column insertion 把新增 widths 插入 direct grid，并将 `xfrm/ext@cx` 设为新 widths safe sum；delete 删除对应 grid columns 并减去其 widths，结果必须大于 0。不存在 automatic column width。

Row insertion/deletion 更新 direct `tr@h`。如果 mutation 后所有 row heights 都大于 0，则 `xfrm/ext@cy` 同步为 exact safe sum，修复原先合法但不一致的 transform；如果任一 row height 为 0，则 rendered automatic total不可由 tokens 推导，保留现有 valid `cy`。这与现有 `setRowHeights()` contract 一致，不声称 content measurement或 layout recomputation。

Column/row size getters 在 mutation 后立即反映新 vector。Table transform、table ID、shape order、name、placeholder identity、slide wrapper 和 `TableModel` live identity 保持。

## Relationship 生命周期

插入不分配 relationship。删除前收集被删除 cells 内所有 namespace-correct、non-empty relationship-ID attributes，包括普通 rich hyperlink、hover/click 扩展及 hyperlink-owned sound/opaque extension；结构写回后重新解析完整 slide：

- 仍被任意 shape、table cell 或 opaque合法 DrawingML hyperlink引用的 ID 保留；
- reference count 变为 0 的 ID 从 slide relationship part 删除；
- shared relationship 仅在最后引用消失时删除；URL/internal-slide target均适用；
- dangling reference 本身没有可删除 relationship，不阻止安全移除其 owning cell；结构写回后的其他 package reference保持原样，validator仍负责报告任何与本次删除无关的既有 dangling state。

Relationship GC 与 slide XML update 位于同一 OPC transaction。任何 `setPart`、relationship removal、outer transaction 或 ZIP date failure 都完整恢复 bytes、relationships、model identity、dates 和 mutation journal。

## Lossless Mutation 边界

Row insertion只在 table direct rows序列中增加完整 row；row deletion只移除指定 rows。Column insertion/deletion只修改 direct grid sequence和每个 direct row的 cell sequence。Surviving element source slices原样拼接；table/grid/row unknown attributes、合法 unknown children、namespace prefixes/quote styles及 table-level extensions不被重排。

Owned numeric/merge tokens只在其语义确实变化时 canonicalize。操作不调用 XML DOM reserialization来重建整张 table，不修复 unrelated malformed state，也不把 absence/materialized defaults、direct no-fill、explicit zero alpha或其他 existing distinctions折叠。

## PptxGenJS 4.0.1 边界

PptxGenJS 4.0.1 公开 table API 只有 creation-time `TableRow[]`、`rowH`、`colW` 和 auto-page helper，没有 existing-deck row/column object model或 structural editor。因此 CRUD 是为“完整编辑 PPT”目标提供的 native lossless extension，不复制不存在的调用签名。

验收覆盖从 PptxGenJS public `write()` 产生的合法 plain/rich/linked/merged table：native import 后 insert/delete，再 write/reopen；所有未删除 content/style/relationship与 PptxGenJS final state保持，结构 mutation新增部分使用 native canonical state。PptxGenJS malformed lopsided rows、invalid merges或不完整 grid仍保持 preservation-only并拒绝 CRUD。

## 错误与事务

- Index、count、width/height inputs在读取 package前完成 type/range/dense/descriptor-safe validation；结构、merge和relationship validation在第一处 observable mutation前完成。
- Insert 超过 physical cell安全上限、delete 越界、删除全部 rows/columns、dimension sum overflow、unsupported topology或invalid transform均抛明确 error并保持 package零变化。
- Duplicate slide中的 table可独立 CRUD；其他 duplicate保持原 bytes。Move/reorder不改变table内部 state。
- Six-format write/reopen、Node/browser bundle、actual packed artifact、PowerPoint 2010 validation和 real-Chrome lifecycle必须覆盖四个 public methods。

## 测试与验收

1. Internal boundary：exact direct ownership、grid/row/cell counts、sizes/transform、merge state、foreign namespace、opaque children与 malformed rejection。
2. Row CRUD：prepend/middle/append、多行、default/explicit heights、auto-height transform规则、empty-cell editability、delete ranges、minimum row guard与 exact survivor bytes。
3. Column CRUD：prepend/middle/append、多列、default/explicit widths、grid/transform exact sum、每行 physical splice、minimum column guard与 survivor bytes。
4. Merge lifecycle：before/inside/after insertion、span expansion、多 region、partial/full deletion、1D退化、1×1 dissolve、anchor promotion、hidden survivor state与 reopen。
5. Relationships：deleted unique/shared URL/internal links、opaque survivor references、zero orphan relationships、failure injection和 outer rollback。
6. Public model/SDK/root：live rows/sizes/merge snapshots、all four methods、positive/negative TypeScript contracts、six formats、duplicate isolation和 write/reopen。
7. Package/browser proof：actual tarball Node/NodeNext/browser conditional export/CLI/Inspector、real Chrome、PowerPoint 2010、part inspection、deterministic pack与 zero console/page/network errors。
8. Final gates：focused/typecheck/build/full/performance、docs/support matrix、review、commit、push和 local/remote `0/0`。

## 后续边界

本小项完成 existing table row/column physical CRUD。后续依次处理 auto-page/repeated headers、content measurement/layout recomputation、`tableToSlides` 与最终 peer/client audit。
