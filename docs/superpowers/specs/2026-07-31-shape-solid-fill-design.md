# Shape Solid/No-Fill Design

日期：2026-07-31

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项为原生 shape 模型增加 PptxGenJS 4.0.1 `ShapeFillProps` 的 solid/no-fill vertical slice：

- preset shape 创建时接受 strict solid 或 none fill；
- 读取 existing deck 中唯一、direct、合法的 no-fill 或 solid fill；
- 对新建或 existing shape 替换、清除 direct fill；
- 保持 live identity、shape order、transaction rollback、duplicate、write 和 reopen；
- 与现有 gradient fill 共存，并为后续 `addText()` outer fill 提供可复用内核。

本设计不实现 line/dash/arrows、shadow、hyperlink、adjustments、custom geometry、shape text、pattern/picture/group fill，也不扩展 slide background。现有 `ShapeModel.gradientFill` 保持独立公开表面。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `ShapeFillProps` 类型与 `addShape()` / `write()` 输出：

- omitted `fill` 写 direct `a:noFill`；
- `{ color: 'FF0000' }` 写 `a:solidFill/a:srgbClr@val="FF0000"`；
- theme color 写 `a:schemeClr`；
- `transparency: 50` 写 `a:alpha@val="50000"`；
- explicit `transparency: 0` 与 omitted transparency 都省略 alpha；
- deprecated `alpha: 40` 写 60% opacity；
- `type: 'none'` 会省略 fill choice，而不是写 direct `a:noFill`；
- empty fill 或 `type: 'solid'` without color 会输出 warning 并回退 black。

原生 API 采用可逆 direct-state 语义，不复制这些 permissive/falsy 行为：solid 必须提供合法颜色，explicit transparency zero 写 `alpha="100000"`，none 写 direct `a:noFill`，unsupported/deprecated fields 在 mutation 前拒绝。合法显式 color、theme color 和非零 transparency 与 PptxGenJS final OOXML/effective rendering 对等；adapter 继续忠实读取 PptxGenJS 已生成的 bytes。

## 3. 方案比较

### 方案 A：复用 `GradientCodec`

把 no-fill/solid fill 加进 `GradientCodec`。改动看似集中，但该 codec 当前处理 slide background、复杂颜色与 gradients，shape lookup/read 又不是 strict direct-state contract；继续扩展会混合 codec ownership、model transaction 和 public shape editing 三类职责。

### 方案 B：复制 table-cell fill 实现

新增完全独立的 shape fill normalizer/parser/renderer。它最容易局部落地，但 color、transparency、descriptor-safe object 和 equality 规则会出现第二份实现，后续 table、shape、text fill 容易漂移。

### 方案 C：共享 simple-fill 内核 + shape 容器适配（采用）

新增 internal `simple-fill`，只负责 descriptor-safe normalize、strict fill-choice decode、deterministic render 和 equality；table-cell 与 shape 各自保留 direct container/navigation/ordering 逻辑。Public `ShapeFill` 与既有 `TableCellFill` 保持语义命名，不暴露 generic internal 类型。该方案让本次改动可审查，也为后续 text shape fill 直接复用。

## 4. 公共 API

```ts
export type ShapeFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
  readonly fill?: ShapeFill;
}

class ShapeModel {
  get fill(): ShapeFill | undefined;
  set fill(value: ShapeFill | undefined);
}
```

`ShapeFill` 使用与 table-cell fill 相同的 native value shape。`RichTextColor` 只接受 strict sRGB 或 supported scheme color。Getter 返回 detached snapshot；unsupported/malformed/ambiguous direct state 返回 `undefined`，不计算 theme/master effective fill。

Creation 中 omitted、runtime `undefined` 和 `{ kind: 'none' }` 都生成 canonical direct `a:noFill`，保持 preset-shape 默认 contract。Editing 中 `undefined` 表示删除唯一 direct fill choice；它不会恢复创建输入或合成 inherited value。

## 5. 输入归一化

所有 public input 在任何 package mutation 前完成：

- fill 必须是 ordinary 或 null-prototype object；
- 只接受 own data properties；symbols、accessors、arrays、class instances 和 unknown keys 全部拒绝；
- `kind: 'none'` 只允许 `kind`；
- `kind: 'solid'` 必须提供 `color`，只允许 `kind/color/transparency`；
- color 必须是 ordinary/null-prototype `{ kind, value }`，只接受六位 sRGB 或 supported scheme token；
- transparency 必须是 finite `0..100`，量化到最近 `0.001%`；
- normalized value 深度脱离 caller，后续 caller mutation 不影响 XML。

Creation options 继续拒绝 inherited/unknown/accessor fields；`fill` 加入唯一允许键集合。`fill: undefined` 与 omitted 等价。Native 不接受 PptxGenJS deprecated `alpha`，也不把 missing color 静默变成 black。

## 6. OOXML 与 direct-state 语义

Supported simple fills：

```xml
<a:noFill/>

<a:solidFill>
  <a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>
</a:solidFill>
```

Reader 只检查 shape 唯一 direct `p:spPr` 下 direct fill choice：`a:noFill`、`a:solidFill`、`a:gradFill`、`a:blipFill`、`a:pattFill`、`a:grpFill`。规则如下：

- exactly one empty/unattributed DrawingML `noFill` → `{ kind: 'none' }`；
- exactly one unattributed `solidFill` with exactly one strict sRGB/scheme color and optional sole direct alpha → solid snapshot；
- absence、gradient/picture/pattern/group、unknown color transforms、namespace lookalikes、duplicate choices 或 malformed state → `undefined`。

Setter 先 normalize，再进入 package transaction：

- unique supported same value 是 exact no-op；
- `undefined` 删除唯一 direct fill choice，包括 unique gradient/picture/pattern/group fill；
- none/solid whole-replace unique existing choice；
- absent choice 时，在唯一 direct `a:prstGeom` 或 `a:custGeom` 后、line/effects/3D/extLst 前插入；
- repeated direct fill choice、repeated/missing direct `spPr`，或 insertion 时缺少唯一 direct geometry，均以 `ModelParseError` zero-mutation 拒绝。

Renderer 使用 geometry 的 in-scope DrawingML prefix，不硬编码 source prefix。除目标 fill choice 外，transform、geometry/adjustments、line、effects、text、non-visual properties、extensions、relationships、neighbor shapes 和其他 parts 保持原 bytes。

## 7. 与 gradient fill 的关系

`ShapeModel.gradientFill` 继续负责 `a:gradFill` 的 typed API。Simple getter 遇到 gradient 返回 `undefined`；这不修改 XML。显式设置 simple fill 会替换 unique gradient choice；`shape.fill = undefined` 会清除 unique gradient choice。反方向由现有 gradient setter whole-replace simple choice。

本小项不把 simple 与 gradient 合并成一个大 union，也不重写 GradientCodec。后续如统一 public fill facade，必须单独设计，不改变本次 direct-state contract。

## 8. 生命周期、错误与无损约束

- Public value errors 使用 `TypeError` / `RangeError`；malformed existing OOXML 使用 `ModelParseError` 并携带 slide part URI。
- Add/edit 都运行在 existing OPC transaction 中；任何失败或 outer rollback 恢复 exact slide bytes、mutation journal 和 live model state。
- Same-value assignment 不写 part、不改变 journal。
- Duplicate、move、write/reopen 和六种格式保持 fill snapshot 与 raw XML。
- Existing unsupported fill 在无关 geometry/text/transform mutation 中原样保留。
- Shape fill 无 relationship ownership；picture fill relationship lifecycle 留给 image/fill 后续小项。

## 9. 验证策略

### Shared/internal contract

- 从 table-cell fill 提取 simple normalize/decode/render/equality 时，existing table tests 保持全绿；
- none、sRGB、scheme、transparency 0/50/100、fractional quantization 有 exact assertions；
- invalid objects/colors/transparency/accessors/symbols 在读取 getter或 mutation 前拒绝；
- alternate prefixes、namespace lookalikes、unsupported transforms 和 duplicate choices 有严格测试。

### Creation and editing

- omitted/undefined/none creation、solid sRGB/scheme/alpha creation；
- immediate live `ShapeModel.fill` snapshot、detachment 和 stable identity；
- same-value no-op、replace、clear、unsupported unique choice replacement、unsafe ambiguity zero mutation；
- fill/geometry/line/effects/text/ext ordering和 isolation；
- duplicate、rollback、write/reopen、six-format lifecycle。

### PptxGenJS and release evidence

- public-only fixtures覆盖 omitted、none、sRGB、scheme、50% transparency 和 explicit zero；
- supported final fill semantics 与 PptxGenJS 对照，strict divergences 单独断言；
- actual tarball Node/browser/types smoke 覆盖 creation/read/edit；
- `pnpm check`、performance、build、PowerPoint 2010 profile validation、LibreOffice visual render 和 artifact-tool import 全部通过。

## 10. 完成门禁

只有以下条件全部满足，本小项才能把 shape fill 标记为已支持：

1. preset shape 可 create/read/replace/clear/reopen simple fill；
2. direct-state ambiguity、same-value no-op、rollback、identity 和 unknown-byte preservation 有证明；
3. table-cell fill 共享内核 refactor 无语义回归；
4. PptxGenJS public-output、packed Node/browser/types、full suite、validator 和 visual QA 全通过；
5. README、CHANGELOG、compatibility matrix 明确 supported scope、strict divergences 和剩余 line/shadow/hyperlink/advanced fills 缺口。
