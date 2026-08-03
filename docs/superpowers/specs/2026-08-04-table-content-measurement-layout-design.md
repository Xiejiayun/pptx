# Table Content Measurement and Layout Recalculation Design

日期：2026-08-04
状态：已批准实施

## 目标与范围

为 native `SlideModel.addTable()` 的 auto-page 路径增加确定性的内容测量与布局重算，使省略或包含 automatic row height 的表格可以按实际 plain/rich text、列宽、colspan、字体大小、字符间距、段落、软换行、cell margin 和 paragraph spacing 物化为 positive EMU row heights，再复用现有分页、重复表头、same-layout slide 插入、section、relationship 和 rollback 生命周期。

本小项同时补齐 PptxGenJS 4.0.1 的 `autoPageCharWeight` / `autoPageLineWeight` table/cell public surface、无 rowspan 超高 text row 的跨页 fragmentation、placeholder auto-page，以及每个 page table transform height 的重算。它保持 explicit fixed-height pagination 的既有精确行为；只有 automatic row、显式 zero row，或明确提供 measurement weight 时才进入内容测量模式。

本小项不导入或解析 HTML，也不创建 HTML source 之外的附加 slide objects；这些属于下一项 `tableToSlides`。它也不承诺成为 PowerPoint 排版引擎的逐像素字体 shaping 替代品，而是提供跨 Node、browser、CLI 和六种输出格式一致的 deterministic estimator。完成本项后，总体剩余项是 `tableToSlides` 和最终 full-surface peer/client audit。

用户已明确授权实现方自主选择最佳方案并连续推进，因此本设计完成方案比较与自审后直接批准，不设置额外确认点。

## 公共 API

```ts
export interface AddTableOptions {
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
  // existing fields unchanged
}

export interface AddTableCellOptions {
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
  // existing fields unchanged
}
```

两个 weight 都必须是 finite number 且位于 inclusive `[-1, 1]`。Table value 是默认值，cell value 是该 anchor cell 的完整 override；explicit zero 保留，不能通过 truthy fallback 丢失。Table weight 只有在 `autoPage: true` 时合法，cell weight 也要求所属 table 启用 auto-page。Accessor、inherited、symbol、class-instance、coercible string 和超范围值在任何 package mutation 前拒绝，不做 silent clamp。

`rowHeights` 的既有 positive scalar/vector contract 保持。仅在 `autoPage: true` 时，scalar zero 或 vector 中的 zero 额外表示该 physical row 需要自动测量；mixed vector 中的 positive value 是该行的 minimum height。包含 zero 的 `rowHeights` 不能再同时提供 `height`，避免一个字段被同时解释为 final transform 和 pre-measure minimum。省略 `rowHeights` 与 `height` 时，全部行自动测量。省略 `rowHeights` 但提供 `height` 时仍按现有逻辑平均分配 fixed positive rows；若同时提供任一 measurement weight，则这些分配值转为 per-row minimum，并在内容需要时向上重算。

以下三种模式明确区分：

1. `rowHeights` 全部 positive 且没有 table/cell weight：fixed mode，沿用现有 exact structural pagination，不测量、不改变任何 row height。
2. 省略 row/height、包含 zero row，或提供任一 weight：measured mode，所有 cell 都测量；positive rows/minimums 不缩小，automatic rows 从内容获得 positive height。
3. `autoPage` 未启用：沿用普通 table creation；automatic `a:tr h="0"` 继续交给 PowerPoint，auto-page weights 严格拒绝。

`SlideModel.newAutoPagedSlides`、repeated headers、continuation Y 和 slide margins 的公共语义不变。

## 方案选择

1. **独立纯测量器先物化 rows/heights，再复用现有分页器；采用。** 新模块只接收 normalized immutable definition 和确定的布局区域，返回 detached/frozen materialized definition。现有 `planTableAutoPages()` 继续只处理 physical rows、positive heights、merge blocks 和 page capacity，因此 slide lifecycle、relationships 与 transaction 不与字体估算耦合。
2. **把断行、row splitting 和 slide insertion 合并到一个 pagination loop。** 这更接近 PptxGenJS 4.0.1 的实现，但会让 caller mutation、line queue、page capacity、header copy、relationship creation 和 package writes共享可变状态，难以证明 pre-mutation failure 与 rollback，故不采用。
3. **加载系统字体并使用 Canvas/native font metrics。** 可改善单一机器上的视觉逼近，但 Node、browser、CI 和 consumer 环境的字体集合、fallback 与 shaping 版本不同，会破坏六种输出的一致性并引入重依赖，故不采用。

## 模块边界与数据流

新增 `packages/model/src/table-content-measurement.internal.ts`，职责限定为四件事：

- 从 normalized cell/rich-text state 生成 immutable measured lines；
- 从 line bands 计算 automatic/minimum row heights和 rowspan lower-bound constraints；
- 把 content-driven oversized、无 rowspan body row 拆成 canonical row fragments；
- 返回 `autoRowHeight: false`、全 positive `rowHeights`、exact summed `height` 的 ordinary normalized table definition。

`table-create.internal.ts` 继续唯一负责 descriptor-safe public input normalization、logical-to-physical merge expansion 和 final cell rendering。它向内部测量器暴露 normalized rich paragraph/run readonly types，不把 measurement metadata写进 public model或 OOXML。

`table-auto-page.internal.ts` 继续负责 slide/page margins、capacity、header boundary、merge blocks和 page slices。共享一个 pure layout-region helper，确保 measurement 与 partition 使用同一组 first/continuation top/bottom limits，不复制略有差异的几何公式。

内部边界固定为 `resolveTableAutoPageLayout()` 返回 frozen `TableAutoPageLayoutRegion { firstY, continuationY, bottomEdge, firstCapacity, continuationCapacity }`，`materializeTableAutoPageContent(definition, region)` 返回 ordinary `NormalizedTableDefinition`。Placeholder auto-page preparation在 definition上设置 internal `placeholderAutoPage: true`；该 flag由 page slices复制，`commitNormalizedTable()`据此替换本地 placeholder owner但保留 page的 row heights/height，不走 legacy stretch-to-placeholder分支。它们都不是 public SDK字段。

`slide.ts` 的顺序变为：

1. normalize public rows/options；
2. preflight placeholder owner和 effective layout region；
3. pure content measurement、row fragmentation和 final page partition；
4. pre-resolve every page hyperlink target与 slide insertion plan；
5. 在一个 existing outer OPC transaction中提交 source table和 continuation slides。

步骤 1–4 不改变 package、slide cache、shape cache或 `newAutoPagedSlides`。测量器不持有 `SlideModel`、`OpcPackage`、relationship ID、XML document或 wall-clock state。

## 确定性字符测量

固定常量与 PptxGenJS public weighting model 对齐：

```ts
const DEFAULT_TABLE_FONT_SIZE_PT = 12;
const BASE_CHAR_DIVISOR = 2.3;
const BASE_LINE_MODIFIER = 1.67;
const EMU_PER_POINT = 12_700;
const EMU_PER_INCH = 914_400;
```

每个可见 text cluster 的 advance 为：

```ts
max(1, round(
  fontSizePt * EMU_PER_POINT * clusterUnits / (BASE_CHAR_DIVISOR + charWeight)
  + characterSpacingPt * EMU_PER_POINT
))
```

Font size按 run style → cell default → 12pt解析，character spacing按 run style解析，char weight按 cell override → table default → 0解析。Font family、bold、italic、baseline、outline、glow和 hyperlink完整保留，但不改变估算宽度；测量不能依赖本机是否安装某个字体。

Cluster算法固定在源码，不调用 locale-sensitive word segmenter：

- `Array.from()` 保证不拆 UTF-16 surrogate pair；
- combining ranges `0300–036F`、`1AB0–1AFF`、`1DC0–1DFF`、`20D0–20FF`、`FE20–FE2F`、variation selectors `FE00–FE0F` / `E0100–E01EF`、emoji modifiers `1F3FB–1F3FF` 和 ZWJ-linked code points附着到前一个 base cluster；
- tab advance到第一个大于 current inline position的 paragraph tab stop；不存在该 stop时前进到下一个 four-space grid；
- 其他 ECMAScript WhiteSpace/LineTerminator是 `0.5` unit；ASCII punctuation ranges `0021–002F`、`003A–0040`、`005B–0060`、`007B–007E` 是 `0.6` unit；ASCII digits `0030–0039`、Latin `0041–005A` / `0061–007A` / `00C0–024F`、Greek `0370–03FF` 和 Cyrillic `0400–052F` base是 `1` unit；其余 base code point是 `2.3` units。

Wrap先在 whitespace boundary断行；单个 non-whitespace token超宽时才按 cluster断开。Whitespace、空 run 和原始 run boundary都不删除。`softBreakBefore` 强制在 run 前断行；由 `breakLine` normalization产生的 paragraph boundary也强制断行；empty paragraph始终产生一条 line。

Cell inline width是 anchor colspan覆盖的 exact column-width sum，减去 left/right cell margins和 paragraph占用。未显式设置的 cell margins使用 renderer同一 canonical defaults：top/bottom `3.6pt`、left/right `7.2pt`。Paragraph `marginLeft`、`marginRight`、`indent`、bullet indent和 tab stops都使用 normalized point/EMU values：first line使用 non-negative `marginLeft + indent`，后续 line使用 non-negative `marginLeft`，active bullet的 first-line reserve至少为 bullet indent，right reserve使用 non-negative `marginRight`；`false`等价于 clear。任一 anchor最终没有 positive inline width时在 mutation前拒绝。

`textDirection` 与 `fit` 保留在 fragment cells。为与 PptxGenJS estimator和跨环境确定性一致，所有 text directions使用 logical horizontal inline estimate；`shrink`/`resize`不反向猜测 viewer执行后的 font scale。该规则可能保守估计 vertical/shrink cell，但不会丢失其最终 OOXML state。

## Line Height 与 Paragraph Spacing

一个 run在未指定 paragraph line spacing时的 natural line height是：

```ts
round(fontSizePt * (BASE_LINE_MODIFIER + lineWeight) * EMU_PER_INCH / 100)
```

Line weight按 cell override → table default → 0解析。一个 measured line取其全部 run natural heights的最大值；空 line使用 cell/default font size。Paragraph resolved `line` spacing覆盖 natural rule：`exact` 使用 exact points，`multiple` 将 natural height乘以 factor。Explicit exact spacing优先于 line weight；multiple spacing在已应用 line weight的 natural height上乘算。Paragraph `before`只加到该 paragraph第一条 line，`after`只加到最后一条 line。所有 point-to-EMU转换只 round一次并检查 safe integer overflow。

每个 cell产生 ordered line boxes。一个 physical row把各 anchor cell的第 N 条 line组成第 N 个 line band，band height取该 band内最大 line-box height；较短 cell在后续 band不贡献高度。Row content height等于 `max effective top margin + sum(all band heights) + max effective bottom margin`。Continuation cells不独立测量；colspan anchor只计算一次。

这种 band model有意比“只取最大 cell总高度”略保守，但它让不同 font size的多列 row具有确定、可拆、与 PptxGenJS round-robin line model相容的垂直边界。

## Rowspan 与 Height Materialization

先用所有 `rowspan === 1` anchor得到每行基础高度。每个 rowspan anchor形成约束：其覆盖 rows的 height sum必须至少达到该 anchor的 measured content height。约束按 span length、anchor row、physical column升序处理；若当前 sum不足，把 deficit以 quotient/remainder平均加到覆盖 rows，remainder从第一行开始每行加一 EMU。约束只增加高度，因此一次有序 pass不会破坏已经满足的短 span。

Automatic row至少为 1 EMU，positive explicit/minimum row不缩小。若 measured mode由 weight触发，所有 positive input row heights都是 lower bound；未触发 measured mode的全-positive vector保持 exact fixed behavior。

最终 materialized definition满足：

- 每个 `rowHeights[i]` 是 positive safe integer；
- `height` 是 row-height vector的 safe exact sum；
- `autoRowHeight` 为 false；
- rows、columns、merge topology、text/style/link values都与 normalized source detached；
- caller objects和 normalized source不被写入 `_lines`、`_lineHeight`或其他缓存字段。

## 超高 Text Row Fragmentation

Repeated header rows永不拆分。包含任意 vertical merge anchor/continuation的 body block也永不拆分；它测量后必须完整放入 continuation body capacity，否则拒绝。

只有同时满足以下条件的 body row可拆：

- block只含这一条 physical row；
- 没有 `rowspan` anchor、`vMerge` continuation或跨 row merge ownership；
- 超高由 measured content bands驱动，而不是仅由 fixed/minimum blank height驱动；
- 至少一条 content band能与该 row的 effective top/bottom margins一起放入 continuation body capacity。

Fragmenter按 ordered line bands贪心选择每个最大 non-empty prefix，使 `topMargin + selectedBands + bottomMargin <= continuationBodyCapacity`。每个 fragment保留原 row的 physical cell count、colspan/gridSpan/hMerge topology、borders、fill、margins、alignment、valign、text direction、fit、default hyperlink和所有 run styles。Top/bottom margins在每个 fragment重复并计入 height，避免跨 slide row的文字贴边。

每个 anchor cell只得到该 fragment band range内的 text slices；没有 slice时写合法 empty paragraph。相邻且来自同一原 paragraph的 wrap slices重新合并，不注入人为 newline。Fragment从原 paragraph中部开始时清除 `spacing.before` 和 bullet/number marker，结束于中部时清除 `spacing.after`；原有 explicit soft break只有在 fragment内部仍有前置内容时保留。Run在 cluster boundary切开时复制完整 style/hyperlink并重算 `text` projection，不能拆 surrogate、combining cluster或 ZWJ sequence。

Fragment heights使用同一 band算法重算，不从原 row height按比例猜测。一个 single line band仍超过 body capacity时拒绝，不缩放、截断或生成超出 slide的 table。Materialized fragments在分页前成为 ordinary physical rows；existing planner不需要知道它们来自同一 logical row。

## Page Capacity 与 Repeated Headers

非 placeholder table沿用现有区域：source `[definition.y, slide.height - bottomMargin)`，continuation `[autoPageSlideStartY ?? topMargin, slide.height - bottomMargin)`。

Measurement先物化 header rows，得到 exact `headerHeight`。`continuationBodyCapacity = continuationCapacity - headerHeight`必须 positive；fragmenter以此作为 single body row fragment上限。Source可以只含 repeated headers，body在第一张 generated slide开始；没有 repeated headers时 source必须仍包含至少一个 body row。普通 row/merge block只因当前 page剩余不足而整体移到下一页，不为填满零碎空间而拆分；只有大于 continuation body capacity的 content row进入 fragmentation。

Page slice继续设置 `height = sum(page.rowHeights)`，因此 `p:xfrm/a:ext@cy`、每个 `a:tr@h`和当前 page rows精确同步。Header fragment prohibition、header/body rowspan boundary和 continuation non-empty规则与现有 planner一致。

## Placeholder Auto-Page

Placeholder selector在 measurement之前通过 source slide的 existing strict owner resolver preflight。Effective source placement使用 layout placeholder的 `x`、`y`、`width`；column widths按该 width等比例缩放后才测量。Placeholder bottom limit是 `min(layoutPlaceholder.y + layoutPlaceholder.height, slide.height - bottomMargin)`，因此表格不会越过 placeholder区域或 slide bottom margin。Zero-width、zero-height、missing、ambiguous、filled、wrong-domain或 unsafe owner在 mutation前拒绝。

Placeholder auto-page把 placeholder height解释为每页最大布局区域，不把每页 measured rows拉伸到该高度。Source page使用 placeholder Y；continuation page使用 `autoPageSlideStartY ?? placeholder.y`，bottom limit保持 layout placeholder bottom。显式 continuation Y必须位于该 bottom limit之前。

Generated slides继承 same layout并 materialize同一 empty placeholder。每个 page table替换其本地 owner，保留该 owner的 identity、shape ID和 name；source与 generated definitions使用已经确定的 fixed/materialized row heights与 exact page height，commit path不再对 auto-page table执行 placeholder height scaling。Non-auto-page placeholder table保持现有 fill-placeholder/scaled-row behavior，不受本小项影响。

`placeholder`、auto-page控制和 measured-placement flag都只存在于 normalized runtime；最终 OOXML仍是普通 placeholder-associated `p:graphicFrame/a:tbl`，不写自定义 extension。

## Relationship、Slide Lifecycle 与 Transaction

Rich run或 plain cell在 fragmentation后可能出现在多个 pages。每个 page继续独立创建自己的 URL/internal-slide relationships；run split复制 semantic hyperlink value而不复制 source `rId`。Default cell hyperlink、run-local override/false、tooltip、underline和 repeated header relationships保持现有 ownership contract。

Internal slide indexes在插入 continuation slides之前解析成 stable target part URIs。Placeholder owner、layout region、measurement、fragmentation、page partition、layout/section insertion plan和全部 hyperlink target也都在第一处 mutation之前完成。

Source table、generated slides、materialized placeholders、relationships、presentation order、sections、slide numbers和 caches继续位于一个 outer OPC transaction。任一 injected failure必须恢复 exact package parts/relationships/dates、mutation journal、existing shape identities和先前成功的 `newAutoPagedSlides`。失败产生的 detached generated `SlideModel`从 cache清除；成功后 getter才原子替换。

## PptxGenJS 4.0.1 对等边界

本小项覆盖 PptxGenJS `TableProps` / `TableCellProps` 声明的 `autoPageCharWeight`、`autoPageLineWeight`，以及 automatic content-based pagination、single text-row continuation和 repeated headers组合。Table default + cell override是 native完整实现；PptxGenJS 4.0.1虽然在两个 type层都声明 weights，但其 paginator会把 table char weight写回 cell并主要使用 table line weight，native不会复制该丢失 override的缺陷。

共同基线使用 12pt default、`2.3 + charWeight`字符 divisor和`1.67 + lineWeight` line modifier。Native在合法 plain ASCII基线 fixtures上要求相同 wrap/page boundaries；在 rich font sizes、CJK、paragraph spacing、colspan、cell override和 placeholder场景，native以本文确定规则为准并要求内容完整、无 overlap/overflow和跨运行时一致。

Native不复制以下 PptxGenJS runtime行为：

- 超范围 weight silent clamp、numeric coercion或 truthy fallback丢失 explicit zero；
- 修改 caller table/cell/run objects并注入 `_lines`、`_lineHeight`或 header caches；
- colspan width filter错误、rowspan height直接置零、margin truthiness和 paragraph spacing遗漏；
- 按 word重建内容时丢失 rich run/soft break/hyperlink boundary；
- 复用并修改 source后已有 slide；
- single line无法容纳时继续生成 overflow table或无限 page loop。

字体 shaping、kerning、ligature、theme fallback和 PowerPoint版本差异不会被虚报为 exact；这些不改变 public feature availability，且所有 estimator差异必须是 documented deterministic behavior，而不是环境随机结果。

## 测试与验收

1. **Normalization：** table/cell weight omitted/zero/bounds/override、autoPage cross-field、getter/symbol/inherited/class/coercion rejection、caller detachment、zero/mixed row vectors和 mode selection。
2. **Character wrapping：** ASCII words、long token、whitespace/tab/tab stop、CJK、emoji surrogate/combining/variation/ZWJ、soft break、breakLine-normalized paragraph、empty paragraph和 exact boundary。
3. **Rich measurement：** run/cell/default font sizes、character spacing、exact/multiple/before/after spacing、paragraph margins/indent/bullet、cell default/custom margins、colspan width、vertical direction/fit preservation和 safe overflow。
4. **Row materialization：** automatic/minimum/fixed rows、multi-cell bands、empty cells、nested/adjacent rowspan lower bounds、detached/frozen output、positive heights与 exact transform sum。
5. **Fragmentation：** multi-page plain/rich rows、unequal cell line counts/sizes、colspan/hMerge、mid-paragraph spacing/bullet cleanup、soft break ownership、default/local URL/internal links、single-band failure、header prohibition和 rowspan block rejection。
6. **Placeholder：** effective width changes wrapping、source/continuation region、placeholder bottom vs slide margin、per-page local replacement、identity/name/shape ID/layout/section、non-auto unchanged和所有 unsafe owner preflight failures。
7. **Integration：** repeated headers、header-only source、exact page edge、source middle/end、following-slide isolation、stable identities、move/delete/duplicate/reopen、`newAutoPagedSlides` reset和 relationship zero-orphan checks。
8. **Public surface：** model/SDK/root/adapter declarations、PptxGenJS plain ASCII boundary comparisons、table/cell override cases、six output formats和 deterministic repeated write。
9. **Failure isolation：** measurement/preflight zero mutation、injected source/generated/placeholder/relationship/section failures、outer rollback、cache cleanup、prior runtime result preservation和 byte-identical retry。
10. **Release proof：** focused/typecheck/build/full/performance、clean dist manifests、two byte-identical actual tarballs、installed Node/NodeNext/browser/CLI/Inspector、real Chrome、PowerPoint 2010 validation、documentation/review/commit/push和 remote divergence `0/0`。

## 完成标准与后续

本项完成时，automatic/zero/minimum table rows、strict table/cell weights、plain/rich/CJK/soft-break/paragraph-aware measurement、merge-aware materialization、无 rowspan超高 text row fragmentation、placeholder auto-page和 per-page transform recomputation全部通过 public、lifecycle、package与真实浏览器证明。Explicit fixed rows继续保持此前 exact structural pagination结果。

完成证明后，把 automatic content measurement/layout recomputation、weights、text-row fragmentation和 placeholder auto-page从所有兼容性文档的 unsupported列表移到 supported；总体仍不宣称 100%，下一项是 `tableToSlides`，最后是 full-surface peer/client audit。
