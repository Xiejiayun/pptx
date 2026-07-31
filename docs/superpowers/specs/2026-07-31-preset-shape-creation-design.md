# Preset Shape Creation Design

日期：2026-07-31

状态：已确认

## 1. 目标与范围

本小项为原生模型增加完整的 preset-shape vertical slice：

- 从零创建 PptxGenJS 4.0.1 公开 `SHAPE_NAME` 所表达的全部合法 preset shapes；
- 读取 existing deck 中唯一、direct、合法的 preset geometry；
- 把 existing 或新建 shape 更换为另一种 preset geometry；
- 保持 live model identity、shape order、transaction rollback、duplicate、write 和 reopen；
- 为后续 fill、line、shadow、hyperlink、adjustment 和 custom geometry 提供稳定边界。

本设计不实现 shape fill、line style、shadow、hyperlink、adjustment handles、custom geometry 或 shape text。这些能力分别进入后续小项，不能把未实现字段静默接受到本 API。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的公开类型声明和 public `addShape()` / `writeFile()` 输出。代表性输出确认：

- omitted options 与 `{}` 都使用 `x=914400`、`y=914400`、`cx=914400`、`cy=914400`；
- omitted fill 写 direct `a:noFill`，omitted line 写 empty direct `a:ln`；
- `rotate: 45` 写 `rot="2700000"`，`flipH/flipV` 写 `1`；
- `objectName` 安全出现在 `p:cNvPr@name`；
- preset geometry 使用 `a:prstGeom@prst` 和 direct empty `a:avLst`；
- `rect`、`ellipse`、`line`、`lineInv`、flowchart、star、action-button 和其他公开 preset 共用同一 shape skeleton。

公开 `ShapeType`/`SHAPE_NAME` 声明列出 178 个 preset token。PptxGenJS 4.0.1 的 `folderCorner` 是运行时缺陷：它输出的 `a:prstGeom prst="folderCorner"` 不是合法 OOXML `ST_ShapeType`，Open XML SDK 拒绝解析；canonical token `foldedCorner` 可正常导入和渲染。原生 API 因此提供 `foldedCorner`，拒绝 `folderCorner`，但 lossless 打开和无关编辑仍保留 existing malformed token。PptxGenJS runtime 额外暴露但类型未声明的 `custGeom` 不属于 preset 集合，留给 custom-geometry 小项。

## 3. 方案比较

### 方案 A：直接扩展 `slide.ts`

把 token 列表、输入检查和 XML renderer 全部放进现有 `slide.ts`。改动最少，但该文件已经同时承担 text、table、notes、visibility 和 shape-tree mutation；继续加入 178-token contract 与 geometry editor 会使边界更难审查，也不利于后续 shape style 复用。

### 方案 B：克隆模板 shape

从 canonical presentation 或隐藏模板中复制一个 shape，再替换 geometry 和 transform。它能继承模板格式，但引入额外 package 依赖、默认样式和 relationship 清理，难以保证 zero-slide/zero-template 创建，也无法产生确定的最小 OOXML。

### 方案 C：独立 public types + internal codec（采用）

`preset-shape.ts` 只定义 canonical token 集合、public types 和 creation options；`preset-shape.internal.ts` 负责 descriptor-safe normalization、direct geometry read/replace 和 deterministic XML rendering。`SlideModel` 只分配 ID、把 XML 插到 shape tree 的正确位置、保存并解析回 live `ShapeModel`。该方案职责清楚，并允许后续 style 小项在不改变 geometry contract 的情况下组合。

## 4. 公共 API

```ts
export const PRESET_SHAPE_TYPES: readonly PresetShapeType[];

export type PresetShapeType = (typeof PRESET_SHAPE_TYPES)[number];

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
}

class SlideModel {
  addShape(type: PresetShapeType, options?: AddShapeOptions): ShapeModel;
}

class ShapeModel {
  get presetType(): PresetShapeType | undefined;
  set presetType(value: PresetShapeType);
}
```

`PRESET_SHAPE_TYPES` 在运行时冻结并保持稳定顺序，可用于 feature discovery、遍历测试和输入验证。它采用 OOXML canonical spelling，不复制 PptxGenJS 的 invalid `folderCorner`。本库不提供 PptxGenJS namespace-compatible `ShapeType` facade。

`AddShapeOptions` 使用 native EMU/OOXML angle units，与现有 `Transform`、`inches()` 和 `degrees()` 一致。defaults 为 1 英寸的 x/y/width/height、0 rotation、无 flip；默认名称为 `Shape ${shapeId}`，保证 existing deck 中可预测且不依赖已删除 shape 的历史计数。自定义名称是 selection-pane metadata，不影响 visible content。

## 5. 输入归一化

所有输入在任何 package mutation 之前完成归一化：

- `type` 必须是 canonical frozen set 的 exact string；
- options 必须是 ordinary 或 null-prototype object；
- 只接受 `name/x/y/width/height/rotation/flipHorizontal/flipVertical` own data properties；
- symbols、accessors、array、class instance 和未知字段全部拒绝；
- `name` 必须是 XML-safe string；
- x/y/width/height/rotation 必须是 finite number，并量化为 safe integer；
- width/height 量化后必须大于零；
- rotation 必须落在 `-21600000..21600000`，即 `-360..360` 度；
- flip values 必须是 boolean。

归一化结果只含 primitive values，立即与 caller 脱离。PptxGenJS 可能对 unknown token、truthy flip、越界 rotate 或 getter-bearing object 继续生成输出；native API 均严格拒绝，不复制这些不安全 runtime 行为。

## 6. OOXML 与数据流

新建 shape 使用最小 deterministic skeleton：

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Shape 2"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="914400" y="914400"/>
      <a:ext cx="914400" cy="914400"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln/>
  </p:spPr>
</p:sp>
```

零 rotation/false flip 不写 attributes；非零/true 值写到 direct `a:xfrm`。Shape 没有 `p:txBody`、relationship、extension 或 compatibility metadata。`SlideModel.addShape()` 在 presentation package transaction 中执行：normalize → parse slide → require exactly one direct shape tree → allocate next shape ID → render → insert before direct `p:extLst` or append → save → resolve live model。解析回的对象必须是同一 ID、`kind === 'shape'` 的 `ShapeModel`，否则整个 transaction rollback。

读取只接受 shape direct `p:spPr` 中唯一 direct `a:prstGeom`、唯一 unqualified `prst` attribute 和 canonical token；missing、unknown、qualified lookalike、duplicate 或 ambiguous state 返回 `undefined`，绝不猜测 effective/master geometry。

更换 preset 时：

- 相同 canonical type 是 exact no-op，保留现有 `a:avLst` 和未知 geometry-local content；
- 不同 type 把唯一 direct `a:prstGeom` whole-replace 为 canonical `prstGeom + empty avLst`，从而清除只对旧 geometry 有意义的 adjustments；
- missing/ambiguous geometry 在 mutation 前拒绝；
- transform、fill、line、effects、text、non-visual properties、extensions、relationships 和其他 parts 保持不变。

## 7. 生命周期与无损约束

Preset geometry 没有 owned relationship 或 dependent part。现有 slide duplicate/move/delete 机制可直接处理它，但必须以测试证明：

- duplicate 保留 exact shape geometry 和 order；
- move 不改变 shape bytes；
- write/reopen 保留 type、transform 和 name；
- outer transaction rollback 恢复 slide bytes、mutation journal 和 model state；
- invalid input、unsafe shape tree 或 unsafe direct geometry 均为 zero mutation；
- 插入前后的 unrelated shape XML、direct `p:extLst`、unknown package parts 和 relationships 保持不变。

Live identity 继续使用 `SlideModel.#shapeModels` 的 `(slide, shapeId)` cache；creation 和 preset replacement 不创建第二个 model object。

## 8. 错误策略

Public input 错误使用 `TypeError` 或 `RangeError`，与现有 strict table/text creation 一致。Malformed 或 ambiguous existing OOXML 使用 `ModelParseError` 并携带 slide part URI。所有错误在 mutation 前发生，或由 package transaction 完整 rollback；不返回 partially-created shape。

## 9. 验证策略

### Internal contract

- frozen token set 恰好 178 个、无重复、包含 `foldedCorner`、不包含 `folderCorner/custGeom`；
- every token 可 normalize 和 render 为 parseable direct `a:prstGeom`；
- defaults、custom transform/name 和 XML escaping 有 exact assertions；
- unknown keys、accessors、symbols、class instances、invalid numbers/booleans/names 全部在调用 getter 或修改 package 前拒绝；
- read/replace 对 missing、duplicate、unknown、qualified lookalike 和 same-value no-op 有独立测试。

### Model/SDK lifecycle

- blank slide 创建 representative shapes 和全 token matrix；
- returned `ShapeModel` identity、kind、name、transform 和 `presetType` 立即正确；
- direct edit、duplicate、rollback、write/reopen 保持语义；
- insertion 位于 direct `p:extLst` 之前；
- invalid operations 对 slide bytes、relationships、parts 和 journal 为 exact no-op。

### PptxGenJS conformance

Public-output fixtures覆盖 omitted/empty options、rect、ellipse、line、lineInv、rotate/flip/objectName、flowchart、star、action button 和 `folderCorner` defect。对合法值比较 final geometry/transform/name/no-fill/line state；对 defect 证明 native `foldedCorner` 生成合法输出而 `folderCorner` 被拒绝。

### Package and visual QA

- packed Node/browser/type smoke 能调用 `addShape()` 并读回 `presetType`；
- `pptx-inspect package validate --profile powerpoint-2010` 为 0 errors/0 warnings；
- exact part read 验证只有目标 slide XML 变化；
- 用 existing gradient fill 给代表性和 full-token shapes 提供可见填充，经 LibreOffice/Poppler 和 presentation renderer 分别导出；逐页检查 geometry、rotation/flip、裁切和 overflow；
- artifact-tool 对 `foldedCorner` 成功导入，PptxGenJS malformed `folderCorner` 失败的行为作为已知 upstream defect 证据保留。

## 10. 完成门禁

只有以下条件全部满足，本小项才能在兼容矩阵标记“已支持”：

1. 178 个 canonical preset types 均可 create/read/replace/reopen；
2. invalid/ambiguous cases 有 zero-mutation 证明；
3. duplicate、rollback、unknown XML preservation 和 live identity 均通过；
4. PptxGenJS public-output conformance、packed smoke、full test suite、PowerPoint 2010 profile validation 和 visual QA 均通过；
5. public README 和 compatibility matrix 记录 supported scope、strict divergences 和后续 styling/custom-geometry 缺口。
