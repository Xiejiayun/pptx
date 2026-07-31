# Custom Geometry Evaluator Design

## 1. 目标

在已完成的 custom geometry direct-state API 上增加一个纯、只读、确定性的
DrawingML geometry evaluator。它使用 shape-local `width/height` 上下文，求值 37 个
DrawingML built-in guides、全部 17 个 guide-formula operators、ordered adjustment/shape
guides，并把 handles、connection sites、text rectangle 与 path command values 转换为
detached deep-frozen numeric snapshot。

本项提供两个入口：

- `evaluateCustomGeometry(geometry, context)` 用于从零创建或独立快照的纯求值；
- `ShapeModel.evaluateCustomGeometry()` 使用 live shape transform 的 width/height 求值当前
  strict custom geometry。

用户已要求实现方连续决定后续内容，因此本设计完成 self-review 后直接
进入实施计划，不设置交互式确认点。

## 2. 当前状态与问题

`CustomGeometry` 已能 strict create/read/edit 下列 direct state：

- `a:avLst` / `a:gdLst` 与全部 17 个 typed formula operators；
- point、radius、angle、handle bound、connection site 和 text-rectangle edge 中的
  safe integer 或 guide/built-in token；
- XY/polar handles、connection sites、custom text rectangles 和全部 path commands。

当前 strict codec 只验证 lexical/schema ownership，故意不判断 unknown built-in、前向引用、
cycle、算术 domain 或最终坐标。这使 direct state 可无损读写，但下游无法得到
connector snapping、geometry inspection 或后续 layout 所需的数值树。

Evaluator 必须保持这个责任边界：求值失败不应让 direct-state getter 失去无损读写
能力，求值成功也不应 canonicalize 或覆写原 OOXML。

## 3. 方案比较

### A. 把 token/formula 原位替换为 number

调用后可直接使用原 `CustomGeometry` 类型，但会丢失 guide 名称、原 formula、token
意图与 exact-byte no-op，还会把本来只读的解析变成 mutation。拒绝。

### B. 在 `customGeometry` getter/codec 内隐式求值

API 表面最少，但会把严格 direct-state ownership 与算术语义绑在一起。一个可无损保留、
但含 unknown token 的文件会突然无法读取，也无法区分 parse failure 与 evaluation
failure。拒绝。

### C. 纯 evaluator + live shape 便捷方法（采用）

在独立文件中实现纯求值内核，输入 direct geometry 与 exact width/height context，
输出新的 numeric tree。`ShapeModel` 只负责读取 live transform 并调用纯函数。该方案
保留 codec 的无损边界，易于单元测试全部算术语义，也为后续 connector
snapping/creation 提供稳定输入。

## 4. 公开 API

```ts
export interface CustomGeometryEvaluationContext {
  readonly width: number;
  readonly height: number;
}

export type CustomGeometryEvaluationErrorCode =
  | 'unknown-token'
  | 'forward-reference'
  | 'cyclic-reference'
  | 'invalid-domain'
  | 'non-finite-result';

export class CustomGeometryEvaluationError extends Error {
  readonly code: CustomGeometryEvaluationErrorCode;
  readonly guideName?: string;
  readonly token?: string;
}

export interface EvaluatedCustomGeometryGuide {
  readonly name: string;
  readonly value: number;
}

export interface EvaluatedCustomGeometryPoint {
  readonly x: number;
  readonly y: number;
}

export interface EvaluatedCustomGeometryTextRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export type EvaluatedCustomGeometryCommand =
  | { readonly kind: 'moveTo'; readonly point: EvaluatedCustomGeometryPoint }
  | { readonly kind: 'lineTo'; readonly point: EvaluatedCustomGeometryPoint }
  | {
      readonly kind: 'arcTo';
      readonly widthRadius: number;
      readonly heightRadius: number;
      readonly startAngle: number;
      readonly sweepAngle: number;
    }
  | {
      readonly kind: 'quadraticBezierTo';
      readonly control: EvaluatedCustomGeometryPoint;
      readonly end: EvaluatedCustomGeometryPoint;
    }
  | {
      readonly kind: 'cubicBezierTo';
      readonly control1: EvaluatedCustomGeometryPoint;
      readonly control2: EvaluatedCustomGeometryPoint;
      readonly end: EvaluatedCustomGeometryPoint;
    }
  | { readonly kind: 'close' };

export interface EvaluatedCustomGeometryXyHandle {
  readonly kind: 'xy';
  readonly position: EvaluatedCustomGeometryPoint;
  readonly xGuide?: string;
  readonly minX?: number;
  readonly maxX?: number;
  readonly yGuide?: string;
  readonly minY?: number;
  readonly maxY?: number;
}

export interface EvaluatedCustomGeometryPolarHandle {
  readonly kind: 'polar';
  readonly position: EvaluatedCustomGeometryPoint;
  readonly radiusGuide?: string;
  readonly minRadius?: number;
  readonly maxRadius?: number;
  readonly angleGuide?: string;
  readonly minAngle?: number;
  readonly maxAngle?: number;
}

export type EvaluatedCustomGeometryHandle =
  | EvaluatedCustomGeometryXyHandle
  | EvaluatedCustomGeometryPolarHandle;

export interface EvaluatedCustomGeometryConnectionSite {
  readonly position: EvaluatedCustomGeometryPoint;
  readonly angle: number;
}

export interface EvaluatedCustomGeometryPath {
  readonly width: number;
  readonly height: number;
  readonly fill?: CustomGeometryPathFill;
  readonly stroke?: boolean;
  readonly extrusionOk?: boolean;
  readonly commands: readonly EvaluatedCustomGeometryCommand[];
}
```

Evaluated handles 保留 guide-reference names，将 position 和所有 present bounds 求值为 number。
Connection sites 将 position/angle 数值化。Evaluated paths 保留 source `width/height` 与
可选 fill/stroke/extrusion flags，commands 改用 numeric union。

```ts
export interface EvaluatedCustomGeometry {
  readonly context: CustomGeometryEvaluationContext;
  readonly adjustments?: readonly EvaluatedCustomGeometryGuide[];
  readonly guides?: readonly EvaluatedCustomGeometryGuide[];
  readonly handles?: readonly EvaluatedCustomGeometryHandle[];
  readonly connectionSites?: readonly EvaluatedCustomGeometryConnectionSite[];
  readonly textRectangle: EvaluatedCustomGeometryTextRectangle;
  readonly paths: readonly EvaluatedCustomGeometryPath[];
}

export function evaluateCustomGeometry(
  geometry: CustomGeometry,
  context: CustomGeometryEvaluationContext,
): EvaluatedCustomGeometry;

export class ShapeModel {
  evaluateCustomGeometry(): EvaluatedCustomGeometry | undefined;
}
```

Context 必须是 ordinary/null-prototype object，只含 own data `width/height`，两者都必须是
positive safe integer。Pure API 继续用现有 `normalizeCustomGeometry()` 严格验证并脱离 caller。

`ShapeModel.evaluateCustomGeometry()` 对 preset、malformed 或 absent custom geometry 返回
`undefined`；对 safe direct state 使用 local `{ width: transform.width, height: transform.height }`。Shape
`x/y/rotation/flip` 不进入 local geometry 求值。语义求值失败抛
`CustomGeometryEvaluationError`，不修改 package 或 mutation journal。

## 5. Built-in guide environment

Evaluator 支持 DrawingML 的 37 个 built-in tokens：

```text
3cd4 3cd8 5cd8 7cd8
b cd2 cd4 cd8 hc h hd2 hd3 hd4 hd5 hd6 hd8
l ls r ss ssd2 ssd4 ssd6 ssd8 ssd16 ssd32
t vc w wd2 wd3 wd4 wd5 wd6 wd8 wd10 wd32
```

- `l = t = 0`，`r = w = width`，`b = h = height`；
- `hc = width / 2`，`vc = height / 2`；
- `ss = min(width, height)`，`ls = max(width, height)`；
- `wdN` / `hdN` / `ssdN` 按 token 中的 N 直接除法；
- `cd2/cd4/cd8` 分别为 180°/90°/45°，`3cd4/3cd8/5cd8/7cd8` 分别为
  270°/135°/225°/315°，全部使用 OOXML `1/60000°`。

Custom guide name 可与 built-in 同名：已求值 custom value 对后续 formula 优先，尚未求值
的同名 guide 不遮蔽 built-in。因此 `name="w" fmla="val w"` 先读取 built-in width，
随后的 `w` 引用读取 custom value。

## 6. Formula 语义与数值策略

Guide 按 `adjustments` source order、再按 `guides` source order求值。每个 operand 的解析优先级是：

1. direct number；
2. 已求值 custom guide；
3. built-in guide；
4. 后续 custom guide 则抛 `forward-reference`；
5. 参与 dependency cycle 则抛 `cyclic-reference`；
6. 其他 token 抛 `unknown-token`。

Evaluator 在开始算术前预扫描 dependency graph，以区分 cycle、forward reference 与真正
unknown token，但不通过拓扑排序放宽 source-order 语义。

| Operator | Result |
| --- | --- |
| `val x` | `x` |
| `abs x` | `abs(x)` |
| `sqrt x` | `sqrt(x)` |
| `at2 x y` | `atan2(y, x)` 转 OOXML angle |
| `cos x y` | `x * cos(y)` |
| `sin x y` | `x * sin(y)` |
| `tan x y` | `x * tan(y)` |
| `max x y` / `min x y` | `max(x, y)` / `min(x, y)` |
| `*/ x y z` | `z === 0 ? 0 : x * y / z` |
| `+- x y z` | `x + y - z` |
| `+/ x y z` | `z === 0 ? 0 : (x + y) / z` |
| `?: x y z` | `x > 0 ? y : z` |
| `cat2 x y z` | `x * cos(atan2(z, y))` |
| `sat2 x y z` | `x * sin(atan2(z, y))` |
| `mod x y z` | `sqrt(x² + y² + z²)` |
| `pin x y z` | `max(x, min(y, z))` |

算术使用 JavaScript IEEE-754 double，不进行整数截断或四舍五入；有限 fractional result 是
合法输出，`-0` 归一化为 `0`。Division by zero 按 DrawingML/Apache POI 行为返回
`0`。Negative sqrt 或任何产生 `NaN`/infinity 的 operation 抛
`invalid-domain` 或 `non-finite-result`，绝不向 snapshot 泄漏非有限 number。

## 7. 完整 geometry tree 求值

求值顺序是 context → dependency audit → adjustments → guides → handles → connection
sites → text rectangle → paths。

- Omitted `adjustments/guides/handles/connectionSites` 不产生 own output property；source order 保留。
- Omitted/canonical-default text rectangle 物化为 `{ left: 0, top: 0, right: width, bottom:
  height }`，因为 evaluated tree 表达 effective geometry，不再表达 direct omission。
- Handle guide-reference names 保留，position 与 present min/max fields 求值；不执行 drag 或 clamp。
- Connection-site order/duplicate 保留，angle/position 求值；不判断它是否位于 path 上。
- Path `width/height` 与 flags 保留，所有 command values 求值为 number。Evaluator 不将
  path coordinate space 缩放到 shape transform，不计算 arc endpoint，也不 tessellate curves。
- Evaluated `arcTo.widthRadius/heightRadius` 必须严格大于 0；guide 产生 zero 或
  negative radius 时抛 `invalid-domain`。
- Root、context、arrays、guides、points、handles、sites、rectangle、paths 与 commands 全部递归
  freeze；每次调用返回新的 detached snapshot。

## 8. 错误与无损边界

Structural input 错误继续使用现有 strict normalizer 的 `TypeError`。只有通过 direct-state
normalization 后的语义失败使用 `CustomGeometryEvaluationError`。Error message 必须包含
guide/field context，`code` 稳定，可选 `guideName/token` 用于机器处理。

Pure evaluator 不持有 package；shape method 只读取 current XML/transform。任何成功或失败调用都必须
保持 part bytes、relationships、shape identity 与 mutation journal 不变。Direct `customGeometry`
snapshot 的可读性不受 evaluator 失败影响。

## 9. 验证策略

### 纯单元测试

- 37 个 built-ins 在 portrait、landscape 和 odd safe-integer size 下的 exact/fractional 结果；
- 17 个 operators 的 positive/zero/negative/quadrant 与 division-by-zero 结果；
- custom-over-built-in precedence、adjustment → guide 顺序、duplicate-free name environment；
- unknown token、forward reference、self/multi-node cycle、negative sqrt、overflow/non-finite 错误码；
- full tree 的 optional-property presence、order/duplicate preservation、default text rectangle 物化、recursive
  freeze 与 caller detachment。

### Model/SDK 纵向测试

- blank deck 中 create/evaluate，transform width/height edit 后重新求值反映 live size；
- write/reopen、duplicate、move 后求值一致，无 mutation journal 增量；
- preset/malformed custom geometry 返回 `undefined`，semantic-invalid safe snapshot 抛 typed error；
- PptxGenJS 4.0.1 legal numeric `custGeom + points` 导入后求值为相同 numeric path；
- actual npm tarball 的 Node/browser/types smoke 公开 evaluator、error class 与 shape method。

### 真实 PPTX 兼容验证

生成一份包含全部 operators、built-ins、handles、sites、text rectangles 和 multi-path 的
gallery。记录 source/evaluated JSON，对 native 与 LibreOffice round-trip 分别 strict reopen/evaluate，
并执行 PowerPoint 2010 validation、180 DPI render、overflow 与逐页视觉检查。

## 10. 不在本项范围

- preset geometry catalog 展开为 custom formula tree；
- external adjustment overrides 与 handle drag behavior；
- connector snapping/creation、path attachment 与 relationship lifecycle；
- path scaling、arc endpoint 计算、curve tessellation、bounds/hit-testing 和 renderer；
- 编辑或 canonicalize source OOXML；
- PptxGenJS 未公开的 arbitrary formula/built-in evaluator API 仿写。

这些边界使本项保持为一个可独立 review、可打包验证的纯求值纵向切片，后续
connector 子系统可直接消费 `EvaluatedCustomGeometry`。
