# 文本框内边距设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为文本 shape 增加四边内边距的原生创建、读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TextPropsOptions.margin` 的标量与四元组输入。该能力属于文本框级布局，映射到 `a:bodyPr`；它不属于 paragraph `a:pPr@marL`、`marR` 或 `indent`，也不改变现有 bullet hanging indent。

本小项只包含 left、top、right、bottom 四个 direct inset。不包含 paragraph 左右 margin/首行 indent、文本 autofit、wrap、vertical alignment、text direction、RTL columns、column count/spacing 或 shape geometry。这些能力继续拆成独立小项；读取和非 margin mutation 不得破坏这些原始 XML。

## API

```ts
export interface TextBoxMargins {
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly left?: number;
}

export type TextBoxMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | TextBoxMargins;

export interface AddTextOptions {
  readonly margin?: TextBoxMarginInput;
}

export class ShapeModel {
  get textMargins(): TextBoxMargins | undefined;
  set textMargins(value: TextBoxMarginInput | undefined);
}
```

`margin` 与 `textMargins` 都使用 point。创建时：

```ts
slide.addText('Uniform', { margin: 8 });
slide.addRichText([{ runs: [{ text: 'TRBL' }] }], {
  margin: [4, 8, 12, 16],
});
slide.addText('Named', {
  margin: { top: 3.5, right: 7, bottom: 3.5, left: 7 },
});
```

标量应用到四边。四元组使用 PptxGenJS 公共注释声明的 `[top, right, bottom, left]` 顺序。具名对象可以只提供需要 direct override 的边；缺少的边不写 direct attribute，继续使用 OOXML 默认值或既有继承语义。

`shape.textMargins` getter 返回新的具名对象，只包含当前 shape 的 `a:bodyPr` 上严格合法的 direct margin attributes；没有任何合法 direct attribute 时返回 `undefined`。setter 是四个受支持属性的整体替换：标量/四元组写全四边，具名对象只保留提供的边，`{}` 或 `undefined` 移除四个 direct margin attributes。getter 对象或 tuple 输入的后续外部修改不影响 model。

## 方案选择

考虑过三种方案：

1. 只复制 PptxGenJS 的 `number | [number, number, number, number]`。迁移最短，但 tuple 在 4.0.1 类型注释和真实 runtime 中存在 top/left 顺序冲突，getter 也无法清楚表示部分 direct attributes。
2. 只公开 `{ top, right, bottom, left }`。语义最清楚，但会让现有 PptxGenJS 标量与 tuple 调用无法直接迁移。
3. 创建输入接受标量、文档顺序 tuple 和具名对象；编辑 getter 始终返回具名 direct snapshot。它覆盖 PptxGenJS 的有效功能、消除读取歧义并支持部分 existing OOXML，因此采用。

不把这项能力放到 `RichTextParagraph`。paragraph margin/indent 与 text-body inset 是两个独立的 OOXML 层级，混用会破坏 bullet、list level 和已有段落模板。

## PptxGenJS 4.0.1 基线

真实生成结果确认：

- 省略 `margin` 时不写 `lIns`、`tIns`、`rIns`、`bIns`。
- `margin: 0` 显式写四个 `0`。
- `margin: 10` 将 10pt 乘 12,700，四边写 `127000` EMU。
- `margin: [1, 2, 3, 4]` 的公共注释声明顺序为 top/right/bottom/left。
- 4.0.1 runtime 对非对称 tuple 实际写成 left=1、right=2、bottom=3、top=4；这是类型注释与实现不一致的缺陷。
- fractional point 按最近 EMU 取整，例如 `0.125` 写 `1588`。
- runtime 接受负数和远高于注释所述 UI 范围的值，也不拦截 non-finite/OOXML 越界结果。

本库采用公开文档的 tuple 顺序，并提供具名对象作为无歧义路径；不复制 top/left 交换缺陷。负值属于 DrawingML signed coordinate 的有效范围，因此保留。non-finite 和量化后越界值在 mutation 前失败。

## OOXML 映射

```xml
<a:bodyPr
  wrap="square"
  lIns="203200"
  tIns="50800"
  rIns="101600"
  bIns="152400"
  rtlCol="0"
  anchor="ctr"/>
```

四边映射为：

| 公共字段 | OOXML attribute |
| --- | --- |
| `left` | `lIns` |
| `top` | `tIns` |
| `right` | `rIns` |
| `bottom` | `bIns` |

point 使用 `Math.round(value * 12700)` 转为 EMU。每个 raw attribute 必须位于 signed Int32 `-2,147,483,648..2,147,483,647`；公共 point 范围由该量化边界决定。显式 `0` 必须保留，省略字段才表示没有 direct override。

getter 只读取严格十进制整数且位于 signed Int32 的 direct attributes，并返回 `raw / 12700` point。科学计数、decimal、空值或越界边值不伪造公共字段；同一 `bodyPr` 上其他合法边仍可读取。读取不解析或改写 OOXML 默认 margin，因为这会丢失 direct 与继承/默认语义的区别。

## 编辑与 Lossless 边界

创建的 `bodyPr` 保持现有 `wrap="square"`、`rtlCol="0"` 和 `anchor="ctr"`，仅在提供 margin 时增加 `lIns/tIns/rIns/bIns`。plain 与 rich text 共用同一 text-shape 创建路径，必须产生相同 margin XML。

setter 定位目标 shape 的 direct `txBody/bodyPr`，只增加、替换或移除四个 margin attributes。原 attribute 的 quote style/order 在替换时沿用 lossless XML 行为；未涉及的 `wrap`、`anchor`、`vert`、`numCol`、`spcCol`、autofit child、extension、namespace 和未知内容保持原字节与相对顺序。缺少 direct `txBody` 或 `bodyPr` 时抛出 `ModelParseError`，不猜测或创建结构。

`.text`、`.richText`、transform 和其他非 margin mutation 保留 `bodyPr` 及其 margin attributes。设置 margins 不重建 paragraphs/runs，不改变 shape identity、text snapshot、relationships 或 mutation journal 之外的 package state。

## 验证与错误处理

scalar、tuple 和对象中的每个提供值都必须是 finite number，量化后的 EMU 必须落在 signed Int32。tuple 必须恰好四项；对象只接受 `top/right/bottom/left` keys。string、boolean、null、数组长度错误、未知 key、非 finite 或越界值明确失败。

所有输入在 package mutation 前完整规范化；一个边失败时不得留下其他边的部分更新。创建或 setter 内失败、以及外层 transaction rollback，都必须恢复 part bytes、mutation journal、live `ShapeModel` identity、text 和 margin snapshot。

## 测试与验收

1. `addText()` 和 `addRichText()` 覆盖省略、scalar、显式 0、文档顺序 tuple、部分/完整对象、negative、fractional 与 signed Int32 边界。
2. getter 读取完整和部分 direct attributes；严格忽略 decimal、科学计数、空值与越界单边，同时暴露同一 `bodyPr` 上其他合法边；只读不产生 mutation。
3. `shape.textMargins` 可增加、更换、部分替换和清除 margins，保留 `bodyPr` 的已知/未知 attributes、children、paragraphs、runs 和 stable identity；write/reopen 与 duplicate 一致。
4. plain/rich text 替换与 transform 编辑保留 margins；缺失 `txBody/bodyPr`、非法输入及内外层 rollback 不改变 bytes、journal 或 live snapshots。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted、zero、scalar、对称 tuple、fractional 和 negative 的真实输出导入；非对称 tuple 缺陷单独记录，不作为本库预期行为。
6. 全仓 typecheck/test、独立 performance、Node/browser bundle、发布 declarations、tarball smoke 和仓库 CLI validate 全部通过。
7. LibreOffice headless 无修复打开并导出 zero/uniform/asymmetric inset 对照页；用同版本 PptxGenJS 对称 margin 文件核对边距方向和可见布局。
