# Custom Geometry Text Rectangle Design

## 1. 目标

在现有 custom geometry paths、guide formulas、adjustment handles 与 connection sites API 上增加
DrawingML `a:rect` custom text rectangle。支持 blank presentation 原生创建、existing-deck strict
direct-state 读取、通过 `ShapeModel.customGeometry` whole-replace 编辑、preset/custom 转换、duplicate、
rollback、六种 presentation formats、write/reopen 和实际 npm tarball。

本项只公开 custom geometry 的文字布局矩形元数据，不实现 formula/geometry evaluation、resolved
rectangle、文字测量/自动适配、connector snapping/creation 或 handle drag evaluation。调用方可引用已存在的
adjustment/shape guide 或 DrawingML built-in token，但 codec 不求值。

用户已明确授权后续方案和执行均由实现方决定，因此本设计完成方案比较与 self-review 后直接进入实施
计划，不设置交互式停顿。

## 2. 当前状态与问题

现有 `CustomGeometry` 已公开：

- `adjustments` / `guides` 与全部 17 个 DrawingML formula operators；
- ordered XY/polar `handles`；
- ordered duplicate-preserving `connectionSites`；
- 六种 path commands、multiple subpaths/paths 与 path flags；
- numeric/token coordinate、radius、angle、bound、position 和 site values；
- strict create/read/whole-replace、detachment、deep freeze、semantic no-op 和转换生命周期。

Renderer 当前固定输出：

```xml
<a:rect l="l" t="t" r="r" b="b"/>
```

Reader 接受 absent rect 或 exact default rect，但任何 non-default `a:rect` 都使 `customGeometry`
返回 `undefined`。因此 LibreOffice 插入的 `textAreaLeft/textAreaTop/textAreaRight/textAreaBottom`、
PowerPoint 生成的 guide-backed rect，以及直接 numeric inset rect 都只能无损保留，不能通过 strict API
读取或编辑。

`a:rect` 是 `a:custGeom` root 上的 optional direct child，位于 `a:cxnLst` 之后、`a:pathLst` 之前。
它有 required `l/t/r/b` 四个 attributes，值使用 custom geometry adjustment-coordinate grammar，允许
direct integer 或 guide/built-in token。它描述形状内部文字可用区域，不是 shape transform、text-body
margin，也不是 resolved bounding box。

## 3. 方案比较

### A. Raw `a:rect` XML

公开 raw XML string 能表达 schema，但会绕过 namespace/attribute ownership、XML escaping、
descriptor-safe input、deep freeze、semantic equality 和 transaction isolation。拒绝。

### B. 两个 point：`topLeft` / `bottomRight`

可复用 `CustomGeometryPoint`，但 public tree 会增加与 OOXML 不一致的嵌套层，并把 text rectangle
误导成已解析 Cartesian bounds；`left > right` 或 token values 也不一定具有普通 point/bounds 语义。
拒绝。

### C. Root-level flat `textRectangle`（采用）

`CustomGeometry.textRectangle` 是 optional `CustomGeometryTextRectangle`，四个 required 字段与
`a:rect@l/t/r/b` 一一对应。名称明确区分 geometry text region、shape transform 与 text-body margins，
同时可直接复用 `CustomGeometryValue` 和现有 strict lexical contract。

## 4. 公共 API

```ts
export interface CustomGeometryTextRectangle {
  readonly left: CustomGeometryValue;
  readonly top: CustomGeometryValue;
  readonly right: CustomGeometryValue;
  readonly bottom: CustomGeometryValue;
}

export interface CustomGeometry {
  readonly adjustments?: readonly CustomGeometryGuide[];
  readonly guides?: readonly CustomGeometryGuide[];
  readonly handles?: readonly CustomGeometryHandle[];
  readonly connectionSites?: readonly CustomGeometryConnectionSite[];
  readonly textRectangle?: CustomGeometryTextRectangle;
  readonly paths: readonly CustomGeometryPath[];
}
```

Public-to-OOXML mapping：

| Public field | OOXML |
| --- | --- |
| `textRectangle` | unique direct `a:rect` |
| `left` | required unqualified `a:rect@l` |
| `top` | required unqualified `a:rect@t` |
| `right` | required unqualified `a:rect@r` |
| `bottom` | required unqualified `a:rect@b` |

`AddCustomShapeOptions` 不增加第二个入口；text rectangle 只属于 `geometry` 参数。现有
`ShapeModel.customGeometry` 仍是唯一 live editor。

## 5. Default 与 source compatibility

Canonical default 是：

```ts
{ left: 'l', top: 't', right: 'r', bottom: 'b' }
```

为保持现有 path/formula/handle/connection snapshot 完全兼容：

- omitted `textRectangle` 不产生 own snapshot property；
- explicit canonical default 同样 normalize 为没有 own property；
- absent source `a:rect` 与 canonical default source `a:rect` 都读取为没有 own property；
- renderer 对没有 own property 的 geometry 继续写现有 canonical default `a:rect`；
- 任一非默认字段存在时，snapshot 保留完整四字段对象；
- 把 non-default geometry whole-replace 为 omitted/default 会重置成 canonical default rect。

因此旧 snapshot、合法 PptxGenJS output 和 existing tests 不增加 empty/default optional state；同时
non-default direct intent 可稳定读取、编辑与重开。

## 6. Direct-state contract

四个字段都复用 `CustomGeometryValue`：

- number 必须是 finite safe integer，按 custom-geometry coordinate space 直接写入；
- string 必须 non-empty、无 XML whitespace、XML-safe 且不是 signed decimal integer string；
- string 作为 guide 或 DrawingML built-in token 原样保留，不验证 guide 是否存在；
- API 不隐式执行 inch、point、percentage 或 shape-transform conversion；
- API 不执行 `left <= right`、`top <= bottom`、rect 位于 path/shape 内或非空区域验证；
- API 不解析 token、dependency、cycle、arithmetic domain 或最终坐标。

这与 path/handle/connection 的 direct-state contract 一致，并把 resolved semantics 留给下一
geometry evaluator 子项。

## 7. Strict input 与 normalization

Root 新增唯一 allowed key `textRectangle`。当 own property 存在时：

1. value 必须是 ordinary 或 null-prototype object；
2. 必须恰有 own data `left/top/right/bottom`；
3. unknown、inherited、symbol、accessor 或 runtime `undefined` field 在 mutation 前拒绝；
4. 每个 value 通过现有 `normalizeCustomGeometryValue(value, context, false)`；
5. result 与 nested root 立即 detached 并 `Object.freeze()`；
6. exact canonical default 不保留 own property；其他值保留完整 frozen object；
7. normalization 不读取 getter，不在失败时分配 shape ID 或修改 package/journal。

Caller 后续修改原 object 不影响 model。Getter 每次返回 detached deep-frozen snapshot。

## 8. Deterministic OOXML

Renderer 保持 `custGeom` child order：

```xml
<a:custGeom>
  <a:avLst>...</a:avLst>
  <a:gdLst>...</a:gdLst>
  <a:ahLst>...</a:ahLst>
  <a:cxnLst>...</a:cxnLst>
  <a:rect l="textAreaLeft" t="25000" r="textAreaRight" b="75000"/>
  <a:pathLst>...</a:pathLst>
</a:custGeom>
```

Attribute order 固定为 `l/t/r/b`。所有 string value 通过 `escapeXmlAttribute()`；number 使用已
normalize 的 canonical integer lexical form。Absent/default property 继续输出
`<a:rect l="l" t="t" r="r" b="b"/>`。Alternate DrawingML prefix 与必要 namespace declaration
沿用现有 codec。

## 9. Strict reader

Reader 继续要求唯一 namespace-correct direct `p:spPr/a:custGeom`、合法 child order、guide/handle/
connection lists 与 supported path tree。Rect 规则：

- absent 或恰好一个 direct same-namespace `rect`；
- rect 无 child、无 non-whitespace text、无 namespace 以外的 extra attributes；
- 恰好 required unqualified `l/t/r/b`；qualified lookalike、missing 或 repeated attribute 拒绝；
- 四值按 numeric-or-token grammar 解析；
- parsed canonical default 不传 root property，non-default 传完整 `textRectangle`；
- 构造 root 后统一调用 `normalizeCustomGeometry()`。

Malformed namespace/attribute/value/child/text/repetition 返回 `undefined` 且不修改 package。Lexical
integer 差异可 semantic normalize；把等价 snapshot 赋回仍识别 exact bytes/journal no-op。

## 10. Equality、whole replacement 与生命周期

`customGeometryEqual()` 在 connection sites 之后、paths 之前比较 optional `textRectangle`：

- 两边 own-property presence 必须一致；
- present 时按 `left/top/right/bottom` 比较 normalized values；
- canonical default 已在 normalization 时折叠为 absent；
- 相同 snapshot 是 exact no-op，任一字段变化触发 whole-geometry replacement。

现有生命周期保持不变：

- preset → custom 与 custom → preset 只替换 geometry choice；
- live identity、name/transform、fill/line/arrows/shadow/hyperlink、shape text/effects/ext、relationships、
  connector sibling shapes 和 unrelated parts 保持；
- duplicate/source isolation、outer transaction rollback、move/delete、六格式 write/reopen 与 opaque
  byte preservation 沿用已有路径；
- 不新增 `setTextRectangle()` 或 partial-field patch API。调用方从 detached snapshot 构造新 root，
  再 whole-replace。

## 11. PptxGenJS 4.0.1 边界

PptxGenJS 4.0.1 public `ShapeType.custGeom + points` 固定生成 canonical default `a:rect`，没有 arbitrary
custom text-rectangle input。本功能是完整 DrawingML 创建/编辑所需的 native extension，不声称对应
PptxGenJS option。

合法 PptxGenJS custom-path output 继续导入为没有 own `textRectangle` property 的相同 snapshot；
malformed runtime output 仍无损保留并返回 `undefined`。Adapter production code 不读取 PptxGenJS
private fields。

## 12. 测试与兼容门禁

### Internal codec

- numeric/token/mixed values、negative/zero/large integers 与 XML escaping；
- explicit default folding、absent/default/non-default readback；
- caller detachment、recursive freeze、semantic equality/no-op；
- unknown/accessor/symbol/missing/unsafe/invalid lexical input 零 mutation 拒绝；
- malformed rect namespace/attribute/child/text/repetition fixtures；
- paths/formulas/handles/connection sites 既有 snapshot 零回归。

### Public lifecycle

- blank create/read/edit/reset、preset/custom conversion；
- identity/style/text/relationship/sibling preservation；
- duplicate/source isolation、rollback、move/delete、六格式 write/reopen；
- SDK zero-input create/reopen/edit/validate；
- PptxGenJS public-output default-rect regression boundary。

### Packed/runtime/real PPTX

- actual tarball Node/browser/types/CLI 覆盖 rect create/read/edit/reset/reopen；
- 真实 gallery 使用已有文字的 shape 转成 custom geometry，覆盖 default、numeric inset、guide/token、
  asymmetric rect 与 lifecycle；
- LibreOffice render/round-trip、逐页 visual review、overflow，并比较 text wrapping/layout 与 raw rect
  normalization；
- 重新检查 path/formula/handle/connection galleries 的 LibreOffice copies，使此前因 non-default rect
  返回 `undefined` 的 shape 在本项后可由 strict reader 打开；
- 原始与 round-trip 文件均用 PowerPoint 2010 profile 要求 0 errors / 0 warnings；
- full typecheck/test/performance/build/dist reproducibility/package smoke。

## 13. 明确剩余边界

完成本项后仍未支持：

- formula dependency/cycle/domain evaluation；
- guide/built-in token 求值与 resolved path/handle/site/text-rectangle coordinates；
- handle drag/constraint evaluation；
- connector snapping、connector creation 与 relationship orchestration；
- custom geometry 与文字内容的自动测量、auto-fit 或 layout recomputation。

## 14. 文档与完成条件

更新 changelog、root/package README、API README 与 PptxGenJS baseline，明确 type、OOXML mapping、
default folding、coordinate units、token、strictness、freeze/no-op、PptxGenJS native-extension status、
LibreOffice normalization 和 remaining evaluator boundary。

完成条件：

1. Public type、codec、reader/editor、lifecycle、packed runtimes 与真实 PPTX 门禁全部通过；
2. Non-default rect 的四值、optional presence 与 exact no-op 不丢失；
3. Existing paths/formulas/handles/connection sites snapshots 与 PptxGenJS default output 无回归；
4. Previous LibreOffice non-default-rect copies 可在不放宽其他 strict ownership 的前提下读取；
5. 每个实施 task review 后独立 commit/push，远端 divergence 为 `0 0`。

## 15. Self-review

- Placeholder scan：没有占位符或未决命名。
- Scope：只增加 direct `a:rect` typed state，不包含 evaluator、text layout engine 或 connector API。
- Type/OOXML：`CustomGeometryTextRectangle` 与 required `l/t/r/b` 一一对应。
- Compatibility：explicit default folding 保持所有旧 snapshot；renderer canonical bytes 不变。
- Strictness：input/reader 均拒绝 unknown、missing、qualified、repeated 和 unsafe state。
- Lifecycle：复用现有 whole-replacement transaction，不新增 partial editor 或 ownership surface。
- Verification：internal/public/packed/real-file/client validation 均有独立完成条件。
