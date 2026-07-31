# Custom Geometry Adjustment Handles Design

## 1. 目标

在现有 custom geometry paths 与 guide-formula API 上增加 DrawingML `a:ahLst`
adjustment handles，完整覆盖 `a:ahXY` 与 `a:ahPolar`。本项支持从 blank
presentation 创建、existing-deck strict direct-state 读取、通过
`ShapeModel.customGeometry` whole-replace 编辑、preset/custom 转换、duplicate、rollback、
六种 presentation formats、write/reopen 和打包产物。

本项不实现 `a:cxnLst` connection sites、非默认 `a:rect` custom text rectangle、
formula/handle evaluation、拖拽求值或 resolved bounds。它们继续无损保留；strict getter
返回 `undefined`，编辑在 package mutation 前拒绝。

用户已要求由实现方连续决定方案并执行，因此本设计完成方案比较和 self-review 后直接进入
实施计划，不设置交互式停顿。

## 2. 当前状态与问题

现有 `CustomGeometry` 已公开：

- `adjustments` / `guides` 和全部 17 个 DrawingML formula operators；
- point、arc radius/angle 对 guide 或 built-in token 的引用；
- 全部六种 path commands、multiple subpaths/paths 与 path flags；
- strict create/read/whole-replace、detachment、deep freeze、semantic no-op 和转换生命周期。

但 codec 仍要求 `a:ahLst` absent 或 empty。任何 non-empty handle list 会使
`customGeometry` 返回 `undefined`，因此无法从零描述一个真正可调的 custom shape，也无法
安全编辑已有 handle geometry。

OOXML handle 有两种：

- `a:ahXY`：可引用 horizontal/vertical adjustment guide，并以 `minX/maxX`、
  `minY/maxY` 表达 direct movement bounds；
- `a:ahPolar`：可引用 radial/angular adjustment guide，并以 `minR/maxR`、
  `minAng/maxAng` 表达 direct movement bounds。

两种 handle 都要求一个 direct `a:pos` child，其 `x/y` 与现有
`CustomGeometryPoint` 使用相同 numeric-or-token grammar。`a:ahLst` 是两类 child 的有序
choice list，因此 API 必须保留 cross-kind order。

## 3. 方案比较

### A. Raw handle XML

允许调用方直接传 `a:ahLst` 字符串。表达力最大，但绕过 namespace/schema order、XML
escaping、descriptor-safe input、deep freeze、semantic equality 和 transaction isolation，
也会给 connection/evaluator 引入第二套表示。拒绝。

### B. 分离 `xyHandles` 与 `polarHandles`

每类结构更简单，但无法保留 existing `ahXY` / `ahPolar` 交错顺序。重新写出时会发生无理由
重排，破坏 direct-state snapshot 和 exact no-op contract。拒绝。

### C. 单一 ordered discriminated union（采用）

`CustomGeometry.handles` 是 readonly ordered list；每项以 `kind: 'xy' | 'polar'`
区分字段。它忠实对应 OOXML choice list，复用 `CustomGeometryValue`、
`CustomGeometryPoint` 与当前 whole-replacement 路径，并为后续 evaluator 提供稳定类型边界。

## 4. 公共 API

```ts
export interface CustomGeometryXyHandle {
  readonly kind: 'xy';
  readonly position: CustomGeometryPoint;
  readonly xGuide?: string;
  readonly minX?: CustomGeometryValue;
  readonly maxX?: CustomGeometryValue;
  readonly yGuide?: string;
  readonly minY?: CustomGeometryValue;
  readonly maxY?: CustomGeometryValue;
}

export interface CustomGeometryPolarHandle {
  readonly kind: 'polar';
  readonly position: CustomGeometryPoint;
  readonly radiusGuide?: string;
  readonly minRadius?: CustomGeometryValue;
  readonly maxRadius?: CustomGeometryValue;
  readonly angleGuide?: string;
  readonly minAngle?: CustomGeometryValue;
  readonly maxAngle?: CustomGeometryValue;
}

export type CustomGeometryHandle =
  | CustomGeometryXyHandle
  | CustomGeometryPolarHandle;

export interface CustomGeometry {
  readonly adjustments?: readonly CustomGeometryGuide[];
  readonly guides?: readonly CustomGeometryGuide[];
  readonly handles?: readonly CustomGeometryHandle[];
  readonly paths: readonly CustomGeometryPath[];
}
```

Public names follow the existing ergonomic model instead of exposing abbreviated XML names:

| Public field | OOXML |
| --- | --- |
| `position` | direct `a:pos` child |
| `xGuide` / `yGuide` | `gdRefX` / `gdRefY` |
| `radiusGuide` / `angleGuide` | `gdRefR` / `gdRefAng` |
| `minRadius` / `maxRadius` | `minR` / `maxR` |
| `minAngle` / `maxAngle` | `minAng` / `maxAng` |

`handles` omitted or `[]` normalize to no own property，保持已有 geometry snapshot 与源码兼容。
`AddCustomShapeOptions` 不增加第二个 handle 入口；handle 只属于第一个 geometry 参数。

## 5. Direct-state contract

`position` 必填并使用现有 `CustomGeometryPoint`：每个 `x/y` 是 direct safe integer 或
single token。Numeric point 和 XY/radial bounds 使用 shape coordinate-space direct value；
numeric angular bounds 使用 OOXML `1/60000°` direct value。API 不隐式执行 inch、degree 或
percentage 转换；调用方可显式使用 `inches()` / `degrees()`。

Guide-reference fields 必须是现有 non-empty、no-XML-whitespace、XML-safe、non-decimal
token string。Min/max fields 复用 `CustomGeometryValue`，允许 negative/zero/positive safe
integer 或 token。

所有 XML attributes 都是 schema-optional，API 保留独立 omission：

- 不要求 guide ref 与 min/max 同时出现；
- 不要求 min 和 max 成对出现；
- 不比较 numeric min <= max；
- 不要求 guide ref 当前解析到 `adjustments`；
- 不拒绝 forward reference、unknown built-in、duplicate position 或 identical handles。

这些属于 handle/evaluator semantic layer。本项只保证明确、可回写的 strict direct state，
避免在没有 OOXML evaluator 时猜测有效移动域。Handle list 顺序具有语义并原样保留。

## 6. Normalization 与 snapshot

Normalization 顺序：

1. descriptor-safe 读取 root、adjustments、guides；
2. 读取 optional dense ordinary `handles` array；
3. 按 `kind` 使用 exact required/allowed key set；
4. normalize required position、optional guide refs 和 optional bounds；
5. normalize paths；
6. deep-freeze handle object、position、list 和整个 root；
7. 空 handle list 不保留 own property。

Handle object 必须是 ordinary 或 null-prototype object，只允许 own data properties；
accessor、inherited/unknown/symbol keys、array subclass、sparse array、非法 token 或 unsafe number
都在 mutation 和 shape ID allocation 前拒绝。Caller 后续修改原数组/object/position 不影响
model；getter 每次返回 detached deep-frozen snapshot。

Semantic equality 比较 handle property presence、kind、position、optional guide/bound values 和
ordered list position。相同 snapshot 赋回保持 exact package bytes 与 mutation journal no-op。

## 7. Deterministic OOXML

Renderer 保持现有 `custGeom` child order，并只替换 `ahLst` 的当前 empty materialization：

```xml
<a:custGeom>
  <a:avLst>
    <a:gd name="adjX" fmla="val 25000"/>
    <a:gd name="adjR" fmla="val 30000"/>
  </a:avLst>
  <a:gdLst/>
  <a:ahLst>
    <a:ahXY gdRefX="adjX" minX="0" maxX="100000">
      <a:pos x="adjX" y="vc"/>
    </a:ahXY>
    <a:ahPolar gdRefR="adjR" minR="0" maxR="50000"
               gdRefAng="adjAng" minAng="0" maxAng="cd">
      <a:pos x="x1" y="y1"/>
    </a:ahPolar>
  </a:ahLst>
  <a:cxnLst/>
  <a:rect l="l" t="t" r="r" b="b"/>
  <a:pathLst>...</a:pathLst>
</a:custGeom>
```

Canonical attribute order：

- XY：`gdRefX`, `minX`, `maxX`, `gdRefY`, `minY`, `maxY`；
- polar：`gdRefR`, `minR`, `maxR`, `gdRefAng`, `minAng`, `maxAng`。

只有 present public property 才写 attribute。Position 使用 direct `a:pos x="..." y="..."/`；
所有 string values 通过 `escapeXmlAttribute()`。Absent/empty handles 继续写 self-closing
`a:ahLst`，alternate DrawingML prefix 与最小 namespace declaration 沿用现有 codec 规则。

## 8. Strict reader

Reader 继续要求唯一 namespace-correct direct `p:sp/p:spPr/a:custGeom`、合法 child order、
guide lists、default rect 和 supported path tree。`ahLst` 规则：

- absent、empty 或恰好一个 non-empty list；list 无 non-namespace attributes 和
  non-whitespace text；
- direct children 只能是 same-namespace `ahXY` / `ahPolar`，按原顺序读取；
- 每个 handle 只允许其 exact unqualified optional attributes；qualified lookalike 和 unknown
  attribute 拒绝；
- 每个 handle 必须有且只有一个 direct same-namespace `pos` child，无额外 child 或
  non-whitespace text；
- `pos` 恰有 unqualified `x/y`，两者按 numeric-or-token grammar 读取；
- guide refs 按 token grammar，bounds 按 `CustomGeometryValue` grammar；
- 构造 root 后统一调用 `normalizeCustomGeometry()`，使 detachment/freeze/strictness 只有一个
  source of truth。

Malformed list、wrong namespace、unknown child、missing/repeated position、extra attributes、unsafe
value、non-empty connection list 或 non-default rect 都返回 `undefined`，且不修改 package。
Lexical integer 差异可 semantic normalize；把 snapshot 赋回仍必须识别 exact no-op。

## 9. Whole replacement 与生命周期

现有 `ShapeModel.customGeometry` 继续是唯一 editor：

- supported custom state 同值赋值 exact no-op；
- 变化值 whole-replace 唯一 direct geometry；
- preset → custom 与 custom → preset 只替换 geometry choice；
- live shape identity、name/transform、fill/line/arrows/shadow/hyperlink、text/effects/ext、
  relationship 和 sibling shape bytes 保持不变；
- duplicate/source isolation、outer transaction rollback、move/delete、六格式 write/reopen、
  unknown XML/opaque parts 沿用已验证路径。

本项不新增 `setHandle()`、drag 或 guide mutation API。调用方通过 detached snapshot 构造新的
ordered `handles` list，再 whole-replace geometry，保持 API 与 paths/formulas 一致。

## 10. PptxGenJS 4.0.1 边界

PptxGenJS 4.0.1 public `ShapeType.custGeom + points` 只生成 empty `ahLst`，没有 arbitrary
adjustment-handle input。本功能是完整 DrawingML 创建/编辑所需的 native extension，不声称有
对应 PptxGenJS option。

现有 PptxGenJS 合法 custom-path output 必须继续导入为没有 own `handles` property 的相同
snapshot；malformed runtime output 仍无损保留并返回 `undefined`。Adapter production code 不因
本项增加私有字段读取。

## 11. 测试与兼容门禁

### Internal codec

- XY/polar、mixed order、multiple/identical handles；
- all optional fields absent、single field、all fields；
- numeric/token position 和 bounds、escaped refs/tokens、alternate prefix；
- caller detachment、recursive freeze、empty-list normalization、semantic equality/no-op；
- unknown/accessor/symbol/sparse/subclass/wrong-kind/unsafe input 零 mutation拒绝；
- malformed list/child/namespace/attribute/position/value fixtures；
- connection/custom rect 继续保持 unsupported boundary。

### Public lifecycle

- blank create/read/edit、preset/custom conversion；
- identity/style/relationship/sibling preservation；
- duplicate/source isolation、rollback、move/delete、六格式 write/reopen；
- SDK zero-input create/reopen/edit/validate；
- PptxGenJS public-output regression boundary。

### Packed/runtime/real PPTX

- typecheck、full suite、performance、build、dist reproducibility；
- actual tarball Node/browser/types/CLI 覆盖 handle create/read/edit/reopen；
- 创建 XY/polar handle gallery，并保存 source snapshots/structural comparison；
- LibreOffice render/round-trip、逐页 visual review、overflow；
- PowerPoint 2010 validator 对原始和 round-trip PPTX 均要求 0 errors / 0 warnings；
- 每个实施 task review 后独立 commit/push，确认 `HEAD...origin/main` 为 `0 0`。

Handle 本身是 PowerPoint 编辑 UI metadata，不保证在 slideshow/render 中可见；gallery 通过
guide-driven geometry 显示当前 adjustment 值，并通过 raw OOXML/snapshot 验证 handle metadata。
Round-trip client normalization 必须与 library 行为分开记录，不为接受含糊 client output 而
放宽 strict reader。

## 12. 明确剩余边界

完成本项后仍未支持：

- connection sites；
- custom text rectangle；
- formula dependency/domain evaluation；
- handle drag evaluation、constraint resolution、resolved coordinates/bounds；
- PowerPoint UI automation 对 yellow adjustment handle 的真实拖拽验证。

后续 connection 与 evaluator 复用本项的 ordered direct-state、`CustomGeometryValue` 和
guide-reference model，不引入 raw XML 旁路。
