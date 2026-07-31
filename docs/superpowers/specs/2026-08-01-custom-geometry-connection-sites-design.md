# Custom Geometry Connection Sites Design

## 1. 目标

在现有 custom geometry paths、guide formulas 与 adjustment handles API 上增加 DrawingML
`a:cxnLst` connection sites。支持 blank presentation 原生创建、existing-deck strict direct-state
读取、通过 `ShapeModel.customGeometry` whole-replace 编辑、preset/custom 转换、duplicate、rollback、
六种 presentation formats、write/reopen 和实际 npm tarball。

本项不实现非默认 `a:rect` custom text rectangle、formula/geometry evaluation、connector snapping
计算、connector relationship 建立、resolved positions 或 angle normalization。它们继续无损保留；
非默认 text rectangle 的 strict getter 返回 `undefined`，编辑在 package mutation 前拒绝。

用户已授予持续方案决策和执行权，因此本设计自行完成方案比较、边界选择与 self-review，不增加
交互式停顿。

## 2. 当前状态与问题

现有 `CustomGeometry` 已公开：

- `adjustments` / `guides` 与全部 17 个 DrawingML formula operators；
- ordered XY/polar `handles`，包括 position、guide refs 与独立 optional bounds；
- point、arc 与 handle fields 对 guide 或 built-in token 的引用；
- 六种 path commands、multiple subpaths/paths 与 path flags；
- strict create/read/whole-replace、detachment、deep freeze、semantic no-op 和转换生命周期。

Codec 仍只接受 absent 或 empty `a:cxnLst`。任何 non-empty connection list 会使
`customGeometry` 返回 `undefined`，因此无法从零描述 shape 的 connector attachment points，也无法
安全编辑已有 connection-site geometry。

DrawingML `a:cxn` 有一个 required `ang` attribute 和一个 required direct `a:pos` child。`ang`
与 position 的 `x/y` 都允许 direct integer 或 guide/built-in token；`a:cxnLst` 保留 source order。
Connection sites 是 shape-level custom geometry metadata，不属于单条 path，也不是 presentation
relationship 或 connector shape 本身。

## 3. 方案比较

### A. Raw connection XML

允许调用方传入 `a:cxnLst` string。表达力高，但绕过 namespace/schema order、XML escaping、
descriptor-safe input、deep freeze、semantic equality 和 transaction isolation，也无法形成稳定类型。
拒绝。

### B. 把 connection sites 放进每条 path

视觉上位置与 path 相关，但 OOXML 的 `cxnLst` 是 `custGeom` root child，不属于 `a:path`。多 path
geometry 会产生虚假的 ownership 和合并顺序。拒绝。

### C. Root-level ordered `connectionSites`（采用）

`CustomGeometry.connectionSites` 是 readonly ordered list，每项用
`CustomGeometryConnectionSite` 表达 required angle 与 position。名称明确区分 connector
attachment metadata 和 OPC relationships，并与 OOXML root ownership 一一对应。

## 4. 公共 API

```ts
export interface CustomGeometryConnectionSite {
  readonly position: CustomGeometryPoint;
  readonly angle: CustomGeometryValue;
}

export interface CustomGeometry {
  readonly adjustments?: readonly CustomGeometryGuide[];
  readonly guides?: readonly CustomGeometryGuide[];
  readonly handles?: readonly CustomGeometryHandle[];
  readonly connectionSites?: readonly CustomGeometryConnectionSite[];
  readonly paths: readonly CustomGeometryPath[];
}
```

Public-to-OOXML mapping：

| Public field | OOXML |
| --- | --- |
| `connectionSites` | ordered direct `a:cxnLst` children |
| `connectionSites[i]` | direct `a:cxn` |
| `angle` | required unqualified `a:cxn@ang` |
| `position` | required direct `a:pos` child |
| `position.x` / `position.y` | required unqualified `a:pos@x` / `@y` |

`connectionSites` omitted 或 `[]` normalize 为没有 own snapshot property，保持已有 geometry source
compatibility。`AddCustomShapeOptions` 不增加第二个入口；connection sites 只属于 geometry 参数。

## 5. Direct-state contract

`position` 复用 `CustomGeometryPoint`。Numeric `x/y` 是 shape coordinate-space direct safe integer；
string 是现有 non-empty、no-XML-whitespace、XML-safe、non-decimal token。API 不隐式执行 inch 或
percentage conversion。

`angle` 复用 `CustomGeometryValue`。Numeric angle 是 OOXML direct `1/60000°` safe integer，调用方可
显式使用 `degrees()`；string 使用同一 token grammar。API 不 clamp 到 `0..<360°`，不解析 guide，
也不把 negative/large angle normalization 成等价角度。

每项必须同时有 own data `position` 与 `angle`，不接受其他 own/inherited/symbol/accessor field。
List 顺序与重复项原样保留；本项不验证 site 位于 path 上、angle 指向外侧、token 已定义，或多个
site 是否重叠。这些属于 evaluator 与 connector-snapping semantic layer。

## 6. Normalization 与 snapshot

Normalization 顺序：

1. descriptor-safe 读取 root、adjustments、guides、handles；
2. 读取 optional dense ordinary `connectionSites` array；
3. 每项要求 exact own data `position` / `angle`；
4. normalize position 与 angle；
5. normalize paths；
6. deep-freeze site、position、list 和整个 root；
7. empty list 不保留 own property。

Caller 后续修改原数组、site 或 position 不影响 model；getter 每次返回 detached deep-frozen
snapshot。Semantic equality 比较 root property presence、ordered list length、每项 angle 与 position；
相同 snapshot 赋回保持 exact package bytes 与 mutation journal no-op。

## 7. Deterministic OOXML

Renderer 保持 `custGeom` child order：

```xml
<a:custGeom>
  <a:avLst>...</a:avLst>
  <a:gdLst>...</a:gdLst>
  <a:ahLst>...</a:ahLst>
  <a:cxnLst>
    <a:cxn ang="0">
      <a:pos x="hc" y="t"/>
    </a:cxn>
    <a:cxn ang="cd4">
      <a:pos x="r" y="vc"/>
    </a:cxn>
  </a:cxnLst>
  <a:rect l="l" t="t" r="r" b="b"/>
  <a:pathLst>...</a:pathLst>
</a:custGeom>
```

每个 `a:cxn` 先写唯一 `ang`，再写唯一 self-closing `a:pos x="..." y="..."/>`。所有 string
values 使用 `escapeXmlAttribute()`。Absent/empty list 写 canonical self-closing `a:cxnLst/>`；
alternate DrawingML prefix 与最小 namespace declaration 沿用现有 codec。

## 8. Strict reader

Reader 继续要求唯一 namespace-correct direct `p:spPr/a:custGeom`、合法 child order、guide/handle
lists、default rect 和 supported path tree。`cxnLst` 规则：

- absent、empty 或恰好一个 list；list 无 non-namespace attributes 和 non-whitespace text；
- direct children 只能是 same-namespace `a:cxn`，按原顺序读取；
- 每个 `cxn` 只允许且必须有一个 unqualified `ang` attribute；qualified lookalike、unknown、
  missing 或 repeated attribute 拒绝；
- 每个 `cxn` 必须有且只有一个 direct same-namespace `a:pos` child，无额外 child 或
  non-whitespace text；
- `pos` 恰有 unqualified `x/y`；angle 与 coordinates 按 numeric-or-token grammar 读取；
- 构造 root 后统一调用 `normalizeCustomGeometry()`。

Malformed list/child/namespace/attribute/position/value 或 non-default rect 返回 `undefined`，且不修改
package。Lexical integer 差异可 semantic normalize；把 snapshot 赋回仍识别 exact no-op。

## 9. Whole replacement 与生命周期

现有 `ShapeModel.customGeometry` 继续是唯一 editor：

- supported custom state 同值赋值 exact no-op；
- 变化值 whole-replace 唯一 direct geometry；
- preset → custom 与 custom → preset 只替换 geometry choice；
- live identity、name/transform、fill/line/arrows/shadow/hyperlink、text/effects/ext、relationships、
  connector sibling shapes 和 unrelated parts 保持不变；
- duplicate/source isolation、outer transaction rollback、move/delete、六格式 write/reopen 与 opaque
  byte preservation 沿用已验证路径。

本项不新增 `setConnectionSite()`、connector creation、snap 或 evaluator API。调用方从 detached
snapshot 构造新的 ordered list，再 whole-replace geometry。

## 10. PptxGenJS 4.0.1 边界

PptxGenJS 4.0.1 public `ShapeType.custGeom + points` 只生成 empty `a:cxnLst`，没有 arbitrary
connection-site input。本功能是完整 DrawingML 创建/编辑所需的 native extension，不声称对应
PptxGenJS option。

现有合法 PptxGenJS custom-path output 继续导入为没有 own `connectionSites` property 的相同
snapshot；malformed runtime output 仍无损保留并返回 `undefined`。Adapter production code 不读取
PptxGenJS private fields。

## 11. 测试与兼容门禁

### Internal codec

- numeric/token angles 与 positions、zero/negative/large angle、ordered/duplicate sites；
- caller detachment、recursive freeze、empty-list normalization、semantic equality/no-op；
- unknown/accessor/symbol/sparse/subclass/missing/unsafe input 零 mutation 拒绝；
- malformed list/child/namespace/attribute/position/value fixtures；
- non-default custom rect 继续保持 unsupported boundary。

### Public lifecycle

- blank create/read/edit/reorder、preset/custom conversion；
- identity/style/relationship/sibling preservation；
- duplicate/source isolation、rollback、move/delete、六格式 write/reopen；
- SDK zero-input create/reopen/edit/validate；
- PptxGenJS public-output regression boundary。

### Packed/runtime/real PPTX

- actual tarball Node/browser/types/CLI 覆盖 site create/read/edit/reopen；
- connection-site gallery 保存 source snapshots 与 structural comparison；
- LibreOffice render/round-trip、逐页 visual review、overflow；客户端若求值 token、插入 non-default
  rect 或扁平化 metadata，只记录 normalization，不放宽 strict reader；
- 原始与 round-trip 文件均用 PowerPoint 2010 profile 要求 0 errors / 0 warnings；
- full typecheck/test/performance/build/dist reproducibility/package smoke。

## 12. 文档与完成条件

更新 changelog、root/package README、API README 与 PptxGenJS baseline，明确 type、OOXML mapping、
coordinate/angle units、token、order、strictness、freeze/no-op、PptxGenJS native-extension status 和
remaining text-rectangle/evaluator boundary。

完成条件：

1. Public types、codec、reader/editor、lifecycle、packed runtimes 与真实 PPTX 门禁全部通过；
2. Connection-site source order、angle/position direct values 与 exact optional root presence 不丢失；
3. Existing supported paths/formulas/handles 无回归；
4. Non-default text rect 继续安全拒绝，不为客户端 normalization 扩大 ownership；
5. 每个实施 task review 后独立 commit/push，远端 divergence 为 `0 0`。

## 13. Self-review

- Scope 单一：只增加 root-level connection sites，不包含 text rectangle、evaluator 或 connector API。
- Type 与 OOXML 一一对应：required angle + required position，没有模糊 optional state。
- Empty-list、order、duplicate、numeric/token、freeze 与 exact no-op contract 明确。
- Strict reader、malformed-state、transaction 和 client-normalization 边界无矛盾。
- PptxGenJS 4.0.1 无对应公开输入，parity claim 只覆盖其 empty-list final output。
