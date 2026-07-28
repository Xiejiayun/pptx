# Basic Table Creation Design

日期：2026-07-29
状态：已批准实施（用户授权实现方持续选择最佳方案）

## 目标与范围

为原生 `PptxDocument.create()` 与已打开的 presentation 增加第一个 table 创建垂直切片：`SlideModel.addTable()` 接受严格二维字符串矩阵和基础 geometry，创建可立即读取、编辑、duplicate、write 与 reopen 的 `TableModel`。该能力不调用 PptxGenJS，也不创建额外 relationship 或 package part。

本小项完成 table creation 的稳定骨架，而不是一次性 XML 注入接口。新建 physical cells 必须立即兼容现有 `setCellText()`、`setCellTextDirection()`、`setCellTextFit()`、`setCellVerticalAlignment()`、`setCellMargins()`、`setCellFill()` 与 `setCellBorders()`；table 本身立即兼容现有 transform 编辑和 stable model identity。

本小项只接受非空、严格矩形、单段 `string[][]`。rich cell text、换行/多段落、cell object options、table/cell text style、table style、独立 `colW` / `rowH`、colspan/rowspan、merge mutation、row/column insertion/deletion、auto page、repeated headers、hyperlink、percent coordinates 与内容测量分别作为后续独立小项扩展。后续把 cell union 从 `string` 扩展为 object 不改变本小项的调用代码。

## 公共 API

```ts
export interface AddTableOptions {
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
}

export class SlideModel {
  addTable(
    rows: readonly (readonly string[])[],
    options?: AddTableOptions,
  ): TableModel;
}
```

示例：

```ts
const table = document.addSlide().addTable(
  [
    ['Region', 'Revenue'],
    ['East', '$1.2M'],
    ['West', '$980K'],
  ],
  {
    name: 'Revenue table',
    x: inches(1),
    y: inches(1.25),
    width: inches(8),
    height: inches(2.25),
  },
);

table.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
});
table.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: 'FFFFFF' },
  width: 1,
});
```

公开 geometry 与现有 `Transform` 一样使用 EMU；调用方用 `inches()` 获得 inch ergonomics。x/y 缺省为 `0.5 inch`，width 缺省为 `columnCount * 1 inch`，height 缺省为 `1 inch`。显式或缺省 x/y 允许负值；width/height round 到 EMU 后必须大于零。`name` 缺省为 `Table <shape id>`。

## 方案选择

考虑过三种方案：

1. 直接复制完整 PptxGenJS `TableProps` 与 `TableCell` union。迁移表面最短，但会在一个提交中混合 text engine、merge grid、row/column sizing、auto page、style inheritance 和 layout measurement，无法为每个语义建立可验证的 strict contract。
2. 先提供严格 `string[][] + geometry` 的语义创建核心，返回现有 `TableModel`，并预留 cell/object 与 sizing union 的无破坏扩展；采用此方案。它形成从零创建的最小完整 table，同时复用全部既有 cell 编辑能力。
3. 暴露 `appendGraphicFrameXml()` 或 table XML builder。实现量少，但把 namespace、schema order、shape id、escaping、grid consistency 和 transaction 责任推给调用方，不构成 PptxGenJS 功能对等。

方案 2 的关键约束是：首个 API 必须是最终 API 的子集，不引入以后需要废弃的 method、单位或返回类型。后续完整 `addTable()` 继续扩展 `rows` 的 cell union 与 `AddTableOptions`，而不是新增第二套创建入口。

## PptxGenJS 4.0.1 基线

通过公开 `addTable()` / `write()`、安装包类型与实际 slide XML 校准：

- `addTable()` 接受 plain string cell 或 `{ text, options }` cell；返回 slide。本库遵循自身对象模型，返回 live `TableModel`，因此创建后无需重新搜索 shape。
- 缺省 x/y 为 `0.5 inch`。缺省 column grid width 为每列 `1 inch`；PptxGenJS 在省略 overall width 时会留下 `p:xfrm/a:ext@cx="0"`，但 grid 总宽仍为列数英寸。本库修正这个内部不一致，使 xfrm width 始终等于 grid width。
- 缺省 xfrm height 为 `1 inch`，每个 row 的 `a:tr@h` 为 `0`，交给 PowerPoint 自动计算；显式 height 且未提供 row heights 时平均分配。
- 缺省 plain cell text 为 12pt black，cell margins 是 left/right `0.1 inch`、top/bottom `0.05 inch`，四边 materialize 为 direct no-fill lines。最终文件没有 table-level margin 或 border state。
- `p:graphicFrame` 使用 table graphicData URI，`a:tbl` 包含 `tblPr`、`tblGrid` 与 physical rows/cells。table 不需要 relationship。
- PptxGenJS 接受 ragged rows、number cells、cell objects、rich text、merge 与大量 permissive runtime value。本小项只宣称 strict rectangular string subset；adapter 继续能读取其合法公开输出。

PptxGenJS conformance 使用显式 geometry 和 plain string matrix 比较最终语义与 XML invariants，不复制 `cx=0` 的缺省 bug、非确定性 PowerPoint 2010 `modId` 或 permissive coercion。

## 输入归一化与错误边界

`addTable()` 在 package mutation 前完整归一化输入：

- `rows` 必须是 non-empty Array；每一 row 必须是 non-empty dense Array，且所有 row length 与第一行完全相同。
- row/outer array 不允许 symbol、named extra property、sparse index 或 index accessor；创建使用 detached normalized strings，调用方随后修改原数组不影响文稿。
- 每个 cell 必须是 string，不能包含 CR/LF 或 XML 1.0 禁止的控制字符。空字符串合法并生成可编辑 empty paragraph。
- `options` 必须是 prototype 为 `Object.prototype` 或 `null` 的 ordinary non-array object；不能含 accessor 或 symbol key，只接受 `name/x/y/width/height` own data keys。name 必须是 string 且不能含非法 XML 控制字符；空名称合法并被原样保留。
- geometry 必须是 finite number，round 后必须是 safe integer。width/height 必须大于零；width 至少允许每个 grid column 分到一个正 EMU，显式 height 至少允许每个 row 分到一个正 EMU。
- unknown key、wrong type、ragged/empty/sparse matrix、line break、invalid XML text/name、non-finite/unsafe geometry 或过小 dimensions 抛 `TypeError` / `RangeError`，并且 package bytes 与 mutation journal 不变。

归一化结果复制所有 string 与 option scalar。创建不保留用户数组、row 或 options object 的引用。

## OOXML 创建

实现层新增窄 `table-create.internal.ts`，只负责输入归一化、deterministic integer distribution 与 canonical table XML。`SlideModel.addTable()` 负责 transaction、解析 slide、定位 direct `p:spTree`、分配 shape id、在 direct `extLst` 前插入、写回和解析 live model。

```xml
<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvGraphicFramePr>
    <p:cNvPr id="2" name="Revenue table"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="914400" y="1143000"/>
    <a:ext cx="7315200" cy="2057400"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr/>
        <a:tblGrid>...</a:tblGrid>
        <a:tr h="685800">...</a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

创建规则：

1. shape id 扫描 slide 全部 `cNvPr@id` 并取有限 numeric max + 1，至少从 2 开始，与 `addText()` 共用同一 allocator，避免 text/table id collision。
2. 新 graphicFrame 自带 canonical `p`/`a` namespace，不依赖源 slide 的 lexical prefix。name 使用 attribute escaping，cell text 使用 text escaping。
3. `p:cNvGraphicFramePr/a:graphicFrameLocks@noGrp="1"` 阻止 UI accidental grouping；不生成随机/non-deterministic `p14:modId`。
4. xfrm width 必须等于所有 `gridCol@w` 之和。整数平均分配使用 quotient + remainder，按左到右把一个额外 EMU 分给前面的列，保证每列正数且总和 exact。
5. 显式 height 时所有 `tr@h` 也用 quotient + remainder 从上到下 exact 分配，总和等于 xfrm height。省略 height 时 xfrm height 为 1 inch，而所有 `tr@h="0"`，保留 PowerPoint auto-size 语义。
6. `tblPr` 初始为空，不伪造 table style、banding 或 header flag。后续 table-style 小项可在同一节点增量编辑。
7. 每个 cell 有 `txBody/bodyPr/lstStyle`、一个 plain paragraph、一个 direct `tcPr`。plain paragraph 复用现有 rich-text normalization/rendering，保持 XML escaping、default theme text color、theme fonts、language 与 empty-run 规则一致。
8. cell direct margins materialize 为 7.2pt left/right 与 3.6pt top/bottom；direct `lnL/lnR/lnT/lnB` materialize 为 canonical zero-width noFill。这与 PptxGenJS effective defaults一致，并让现有 strict margin/border snapshots 立即可读。
9. cell 不写 direct fill、anchor、vert 或 fit；现有 getters 对这些状态返回 `undefined`，setter 可立即增量写入。
10. slide 的 direct `cSld/spTree` 必须恰好一个；其 direct `extLst` 至多一个。如果有 `extLst`，graphicFrame 插在它之前；否则 append。缺失、重复或层级错误时抛 `ModelParseError`，不猜测 descendant target。

## 原子性、identity 与交叉能力

validation、ID allocation、XML rendering、slide patch、write-back 和 created-model resolution 属于同一 OPC transaction。任何异常或 outer document transaction rollback 恢复 exact slide bytes、mutation journal 和既有 model identity。

成功后 `addTable()` 返回 slide shape cache 中的 live `TableModel`。重复访问 `slide.shapes`、读取 rows、编辑任一 cell、table transform、duplicate slide 与 write/reopen 都保持语义；同一 live slide 中 model identity 为 `===`。外层 rollback 删除刚创建 table 后，stale handle 按现有 live-model 规则在 resolve 时抛 `ModelParseError`。

创建 table 只改变目标 slide XML。它不改变 presentation、master、layout、theme、relationships、其他 slide、既有 shape bytes 或 package content types。

## 测试与发布门禁

1. model/internal tests 覆盖 1x1、2x2、multi-row/column、special XML text/name、empty string、default/explicit geometry、integer remainder distribution、ID allocation、extLst insertion、detached input 与 missing shape tree。
2. invalid tests 覆盖 null/non-array/empty/ragged/sparse/extra-property/accessor matrices、non-string cells、line breaks、invalid XML characters、unknown/wrong options、non-finite/unsafe/zero/negative/too-small geometry；失败保持 exact bytes/journal。
3. SDK lifecycle 从 `PptxDocument.create()` 创建 table，立即执行全部现有 cell setters与 table transform，验证 stable identity、duplicate isolation、outer rollback、write/reopen 和 package validator 0 errors。
4. PptxGenJS 4.0.1 conformance 使用 plain string matrix + explicit geometry，验证 table URI、shape geometry、grid/row totals、physical cells、text、direct margins/no-border 与 round-trip；只使用公开 API/输出。
5. packed npm Node/browser/declaration smoke 编译并运行 `addTable()`，验证返回类型、rows、geometry、cell mutation、write/reopen；CLI smoke 保持通过。
6. API、npm README、changelog 与 compatibility matrix 新增 `addTable(string[][])` 基础创建，明确 strict rectangular/single-paragraph、EMU geometry、defaults、direct cell defaults，以及 cell objects/merge/sizing/style/auto-page 仍待后续小项。
7. typecheck、全仓 tests、独立 performance、npm tarball smoke 全部通过；`git diff --check` 无错误。
8. native 与 PptxGenJS baseline 文件使用相同 explicit geometry/plain matrix；PowerPoint 2010 profile 0 error/0 warning，LibreOffice 无修复打开/导出，两份 overflow 检查通过。渲染逐图检查 table bounds、row/column geometry、文字、空 cell 与 special characters；如默认字体产生可解释差异，不以 raster hash 相等冒充功能等价。
