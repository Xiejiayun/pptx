# Custom Geometry Paths Design

## 1. 目标

为 DrawingML `a:custGeom/a:pathLst` 增加严格的原生创建、读取、whole-replace
编辑和 preset/custom 双向转换能力。首个 custom-geometry 纵切面覆盖 PptxGenJS
4.0.1 公开 `ShapeType.custGeom + ShapeProps.points` 能产生的全部合法最终路径命令，
同时提供现有 PPTX 的安全 direct-state 编辑、多个 path、path fill/stroke/extrusion
属性和完整生命周期保证。

本小项不实现 arbitrary guide formulas、non-empty `avLst/gdLst`、adjustment
handles、connection sites、自定义 text rectangle 公式或 geometry evaluation。这些状态继续
无损保留，严格 getter 返回 `undefined`，编辑在 package 变化前拒绝；它们作为紧接本小项
之后的 custom-geometry formulas/handles 子项，而不是被 raw XML 旁路。

## 2. 兼容性证据

PptxGenJS 4.0.1 的公开类型把 custom geometry 表达为一组 `points`：普通 point
生成首个 `moveTo` 或后续 `lnTo`，`moveTo: true` 开始新 subpath，`curve.type`
分别生成 `arcTo`、`quadBezTo` 或 `cubicBezTo`，`{ close: true }` 生成 `close`。

公开输出 probe 确认：

- `ShapeType.custGeom` 是 runtime/public enum 值，但不属于 OOXML preset token；
- 输出固定 materialize empty `avLst/gdLst/ahLst/cxnLst`、default text rect、一个
  `pathLst/path`，path 的 `w/h` 等于 shape extents；
- 第一个普通 point 无论 `moveTo` 是否省略都会成为 `moveTo`；后续普通 point 成为
  `lnTo`，显式 `moveTo` 可创建更多 subpath；
- arc point 的 `x/y` 被 runtime 忽略，因为 OOXML `arcTo` 没有 endpoint attribute；
- number 小于 100 时按 inch 转 EMU，大于等于 100 时被当作已转换数值，percentage
  string 按整张 slide 而不是 shape/path 计算；
- empty `points` 仍生成一个 empty path；
- representative all-command、empty 和 percentage/EMU 文件通过 PowerPoint 2010
  profile validation，0 error、0 warning。

Native 对等目标是这些合法最终 OOXML path snapshots，而不是复制 PptxGenJS 的单位
heuristic、string coercion、ignored arc endpoints 或 malformed runtime passthrough。

## 3. 方案比较

### A. 公开 raw `a:custGeom` XML

表达力最大，但调用方可绕过 namespace、schema order、XML escaping、safe integer、
detachment、semantic no-op 和 mutation isolation。它也无法提供稳定的 TypeScript API。
拒绝。

### B. 原样复制 PptxGenJS `points`

迁移成本低，但 arc entry 暴露最终不会写出的 `x/y`，coordinate 同时混合 inch、EMU、
numeric string 和 slide-relative percentage，且只能表达一个 path。它不适合作为现有文件的
direct-state snapshot。拒绝作为原生模型；只用于 public-output conformance fixture。

### C. Typed direct path model（采用）

公开与 OOXML 一一对应的 path/command discriminated unions，所有 coordinate 使用 native
EMU integer、所有 angle 使用 OOXML angle integer。它可无歧义导入 PptxGenJS 的合法最终
输出，支持多个 path 和 existing-deck editing，并给 formulas/handles 留出独立扩展边界。

## 4. 公共 API

```ts
export interface CustomGeometryPoint {
  readonly x: number;
  readonly y: number;
}

export type CustomGeometryCommand =
  | { readonly kind: 'moveTo'; readonly point: CustomGeometryPoint }
  | { readonly kind: 'lineTo'; readonly point: CustomGeometryPoint }
  | {
      readonly kind: 'arcTo';
      readonly widthRadius: number;
      readonly heightRadius: number;
      readonly startAngle: number;
      readonly sweepAngle: number;
    }
  | {
      readonly kind: 'quadraticBezierTo';
      readonly control: CustomGeometryPoint;
      readonly end: CustomGeometryPoint;
    }
  | {
      readonly kind: 'cubicBezierTo';
      readonly control1: CustomGeometryPoint;
      readonly control2: CustomGeometryPoint;
      readonly end: CustomGeometryPoint;
    }
  | { readonly kind: 'close' };

export type CustomGeometryPathFill =
  | 'none'
  | 'norm'
  | 'lighten'
  | 'lightenLess'
  | 'darken'
  | 'darkenLess';

export interface CustomGeometryPath {
  readonly width: number;
  readonly height: number;
  readonly fill?: CustomGeometryPathFill;
  readonly stroke?: boolean;
  readonly extrusionOk?: boolean;
  readonly commands: readonly CustomGeometryCommand[];
}

export interface CustomGeometry {
  readonly paths: readonly CustomGeometryPath[];
}

export type AddCustomShapeOptions = Omit<AddShapeOptions, 'adjustments'>;

class SlideModel {
  addCustomShape(
    geometry: CustomGeometry,
    options?: AddCustomShapeOptions,
  ): ShapeModel;
}

class ShapeModel {
  get customGeometry(): CustomGeometry | undefined;
  set customGeometry(value: CustomGeometry);
}
```

`addShape()` 继续只接受 178 个 `PresetShapeType`；`custGeom` 不进入
`PRESET_SHAPE_TYPES`。单独的 `addCustomShape()` 让 creation options 保持静态无歧义，
也避免把 preset-only `adjustments` 静默接受到 custom geometry。

`customGeometry` getter 对 supported custom path state 返回 detached deep-frozen
snapshot；对 preset 或 unsupported custom state 返回 `undefined`。Setter 只接受完整
`CustomGeometry`，不存在 `undefined` clear，因为合法 shape 必须保留一个 geometry choice。

## 5. 输入 contract

所有 arrays 必须是 dense ordinary arrays，无 symbol、accessor、extra own property 或 array
subclass。所有 objects 必须是 ordinary 或 null-prototype object，只含 union branch 明确列出的
own data properties。Normalization 绝不调用 caller getter，并在任何 shape ID allocation、
relationship 或 package mutation 前完成深复制和 deep freeze。

具体约束：

- `paths` 至少一个；每个 path 的 `commands` 可为空，以支持 PptxGenJS empty points；
- path `width/height` 为正 safe integer EMU；
- point `x/y` 为 finite safe integer EMU，允许负值；
- arc radii 为正 safe integer EMU，angles 为 safe integer OOXML angle；
- path fill 必须是六个 `ST_PathFillMode` token 之一；stroke/extrusion 必须是 boolean；
- command 必须有一个已知 `kind`，字段集合与该 branch 完全匹配；
- `close` 不能是 path 的第一个 drawing command；non-empty path 的第一个非-close
  command 必须是 `moveTo`，从而拒绝无当前点的 line/curve；
- 后续 `moveTo` 合法并开始新 subpath；`close` 后可用新的 `moveTo` 继续。

Native 不接受 inch/percentage/string coordinate shortcut；调用方用现有 `inches()` 和
`degrees()` 明确转换。这样 final stored unit 只有一种，不复制 PptxGenJS `<100` heuristic。

## 6. Supported OOXML direct state

Reader 只从 namespace-correct direct `p:sp/p:spPr/a:custGeom` 开始，并要求
`p:spPr` 中只有一个 geometry choice。Supported state 允许 absent 或 empty same-namespace
`avLst/gdLst/ahLst/cxnLst`；一旦任一列表含 guide、handle、connection 或其他 child，
snapshot 为 `undefined`。

`a:rect` 可 absent，或必须是唯一、same-namespace、无 child 且 exact unqualified
`l="l" t="t" r="r" b="b"`。其他 text-rectangle formulas 留给下一子项。

`a:pathLst` 必须唯一并至少包含一个 direct same-namespace `a:path`。每个 path：

- 必须有唯一 unqualified positive integer `w/h`；
- 只允许 optional unqualified `fill/stroke/extrusionOk`；
- boolean lexical forms 按 OOXML 解析为 boolean，非法 token 拒绝；
- 只允许 direct `moveTo/lnTo/arcTo/quadBezTo/cubicBezTo/close`；
- point/curve child count、顺序、namespace 和 owned attributes 必须 exact；
- coordinate/radius/angle attribute 必须是 signed safe decimal integer；guide-name、
  formula token、qualified lookalike、extra child/attribute 或 non-whitespace text 均拒绝。

Reader 不修改 bytes。Unsupported state 继续由 lossless kernel 在无关编辑、duplicate、write
和 reopen 中原样保留。

## 7. Deterministic rendering

Creation 和 replacement 使用 compact canonical order：

```xml
<a:custGeom>
  <a:avLst/>
  <a:gdLst/>
  <a:ahLst/>
  <a:cxnLst/>
  <a:rect l="l" t="t" r="r" b="b"/>
  <a:pathLst>
    <a:path w="3657600" h="2743200">
      <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
      <a:lnTo><a:pt x="914400" y="0"/></a:lnTo>
      <a:arcTo wR="914400" hR="457200" stAng="1800000" swAng="7200000"/>
      <a:close/>
    </a:path>
  </a:pathLst>
</a:custGeom>
```

Omitted path fill/stroke/extrusion attributes remain omitted so DrawingML defaults apply。渲染使用
当前 geometry 的 in-scope DrawingML prefix；如果 parent 没有正确 binding，replacement 在
root geometry 上补最小 namespace declaration。

## 8. 创建、编辑和转换

`SlideModel.addCustomShape()` 复用 preset creation 已有 transform、name、fill、line、arrows、
shadow、hyperlink strict normalizers 和 transaction flow。Custom geometry normalization 与
options normalization 全部先完成，再分配 shape ID。Renderer 只把 preset skeleton 中的
`prstGeom` 换成 canonical `custGeom`；其他 direct style semantics 完全一致。

Setter 对 supported custom state 赋相同 snapshot 是 exact bytes/journal no-op。不同值
whole-replace 唯一 direct `a:custGeom`，保留 transform、non-visual identity、name、fill、line、
arrows、effects、hyperlink、text、shape order、extensions、relationships 和 live object identity。

Preset/custom conversion规则：

- 对唯一 supported preset shape 设置 `customGeometry`，只把 `prstGeom` whole-replace 为
  canonical `custGeom`；
- 对 supported custom shape 设置 `presetType`，只把 `custGeom` whole-replace 为 canonical
  `prstGeom + empty avLst`；
- 对 unsupported、missing、repeated、wrong-namespace 或 mixed geometry choice，两个方向均在
  package 变化前抛 `ModelParseError`；
- conversion 不尝试保留 preset adjustments 或 custom paths，因为它们属于被替换 geometry。

## 9. 生命周期和错误策略

Public input 错误使用 `TypeError` 或 `RangeError`；unsafe existing OOXML 使用带 slide part
URI 的 `ModelParseError`。所有失败保持 slide bytes、relationships、parts、shape ID sequence、
mutation journal 和 cached model identity 不变。

能力必须覆盖：caller detachment、deep-frozen snapshots、source/duplicate isolation、outer
transaction rollback、六种 presentation formats、write/reopen、move/delete、unknown XML 和
opaque package part preservation。Custom geometry 不拥有 relationship 或 dependent part。

## 10. 验证策略

### Internal codec

- normalize/render/read 全六种 command、多个 subpath、多个 path、empty commands、negative
  coordinate、zero angle、path flags、alternate prefix；
- reject sparse/exotic arrays、symbols、accessors、unknown/missing fields、non-safe integers、
  invalid radii/extents、bad sequence、duplicate/malformed/mixed geometry；
- supported list/rect absence 与 empty canonical forms 语义相同；non-empty formulas/handles/
  connections 和 custom rect 均 read `undefined`、edit zero-mutation；
- same snapshot exact no-op、prefix-preserving whole replacement、preset/custom conversion 和
  unrelated byte preservation。

### Public model and SDK

- 从 blank presentation 创建 custom shapes，并组合 transform、fill、line、arrows、shadow、
  hyperlink；
- live `ShapeModel.customGeometry` create/read/edit、preset↔custom、duplicate、rollback、
  six-format write/reopen、stable identity；
- invalid input 在 ID allocation 和 relationship creation 前拒绝；
- TypeScript declarations、Node/browser bundle、CLI 和 actual tarball 均可使用公开 types/API。

### PptxGenJS conformance and release QA

- public-output fixtures 覆盖 ordinary/multiple move、line、arc、quadratic、cubic、close、empty、
  inch、direct numeric 和 percentage final state；
- imported PptxGenJS final snapshots 与 native explicit-EMU creation 相同；
- 记录 arc endpoint ignored、slide-relative percentage、`<100` numeric heuristic、coercion 和
  malformed runtime passthrough，但 native 不复制这些缺陷；
- full check、performance、build、packed smoke、PowerPoint 2010 profile、LibreOffice render/
  round-trip、Poppler PDF、overflow 和 artifact-tool import 全部通过；
- visual gallery 至少包含 open/closed polyline、arc、quadratic/cubic、multiple subpaths、multiple
  paths、style coexistence 和 reopen copy。

## 11. 完成门禁

只有以下条件全部满足，custom geometry paths 才标记完成：

1. 六种 command、multiple subpaths/paths 和 path flags 可 create/read/replace/reopen；
2. PptxGenJS 4.0.1 所有合法 custom-points final output 可导入为相同 snapshot；
3. malformed、formula/handle/connection/custom-rect state 无损保留且不能被误编辑；
4. preset/custom 双向 conversion、no-op、detachment、identity、duplicate、rollback、六格式和
   unrelated-byte preservation 有直接测试；
5. packed public surfaces、validator、round-trip 和逐页视觉 QA 通过；
6. formulas、adjustment handles、connection sites、custom text rectangle 与 geometry evaluation
   明确列为下一 custom-geometry 子项。
