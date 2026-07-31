# Shape Shadow Design

日期：2026-07-31

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项为原生 preset shape 增加 PptxGenJS 4.0.1 `ShadowProps` 的 shape-level vertical slice：

- `slide.addShape()` 创建时可指定 outer 或 inner shadow；
- 支持 sRGB/theme color、opacity、blur、angle、distance，以及 outer-only rotate-with-shape；
- 读取 existing deck 中唯一、direct、合法的 `a:outerShdw` 或 `a:innerShdw`；
- 通过 `ShapeModel.shadow` whole-replace shadow kind/parameters，或用 `undefined` 清除 direct shadow；
- 保留同一 `a:effectLst` 中其他合法 effect、shape geometry/fill/line/arrows/hyperlink/text、unknown unrelated bytes、live identity 和 transaction isolation；
- duplicate、rollback、write/reopen、六种 presentation format 与 packed Node/browser/types 保持一致。

本设计只开放 `p:sp/p:spPr/a:effectLst` 中的 simple inner/outer shadow。它不实现 preset shadow、reflection、soft edge、blur、fill overlay、effect DAG、custom shadow scale/skew/alignment、group/image/text/table/chart/media shadow API，也不改变 `addText()` 的 creation options。这些调用方可复用 public value 的参数语义，但其 OOXML owner 和 lifecycle 后续独立实施。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `ShadowProps` 类型与 `addShape()` / `write()` 真实输出。公开类型声明：

```ts
interface ShadowProps {
  type: 'outer' | 'inner' | 'none';
  opacity?: number;
  blur?: number;
  angle?: number;
  offset?: number;
  color?: HexColor;
  rotateWithShape?: boolean;
}
```

真实 shape public-output 审计结果如下：

- omitted shadow 与 `{ type: 'none' }` 都不写 direct effect；
- `{ type: 'outer' }` 写 outer shadow，实际 defaults 为 black、opacity `0.75`、blur `8pt`、angle `270°`、offset `4pt`、`rotWithShape="0"`；
- custom outer 的 blur/offset 以 point 转 EMU、angle 以 `1/60000°` 转换、opacity 以 `1/100000` 转换；
- outer 固定写 `sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="0"`，因此 public `rotateWithShape: true` 在 shape serializer 中被忽略；
- runtime zero opacity/blur/angle/offset 被 `||` fallback 折叠为上述非零 defaults；
- shape 路径未执行共享的 shadow correction：hash-prefixed color、unknown type、negative/out-of-range blur/offset/angle/opacity可直接写入非法 OOXML且没有 warning；
- `{ type: 'inner' }` 开始写 `<a:innerShdw>`，却错误地以 `</a:outerShdw>` 闭合，导致整个 slide XML not well-formed；
- `type` 缺失的 runtime empty object被当作 outer，但 public TypeScript type要求 `type`。

Native 只对合法、可逆的 final-state 能力宣称对等。它不复制 falsy fallback、invalid passthrough、ignored rotate flag或 malformed inner closing-tag缺陷。Native `undefined` 对应 omitted/`type: 'none'`；zero 值保持为 direct zero。Outer defaults选择 PptxGenJS 4.0.1 的真实 shape output，而不是其源文件中未被 shape serializer使用的 `DEF_SHAPE_SHADOW` 常量。

## 3. 方案比较

### 方案 A：原样复制 `ShadowProps` 并宽松写出

该方案调用形式最接近 PptxGenJS，但会把 `none`、invalid type、hash color、coercible value和 out-of-range number带入 native API，并可能复制 malformed inner shadow。它不能满足 strict mutation-before-validation或合法 PPTX 门禁。

### 方案 B：一次实现 generic DrawingML effect stack

通用 `effects[]` 可以覆盖 glow/reflection/soft-edge/preset-shadow/effect-DAG，但远超当前 PptxGenJS shape shadow小项。它会让 ownership、schema order、partial editing和 public types一次膨胀，难以独立验证。

### 方案 C：focused strict `ShapeShadow`（采用）

Public value使用 `kind: 'outer' | 'inner'`，并公开当前 simple shadow真实需要的参数。`undefined` 表示无 direct shadow；setter只拥有一个 direct inner/outer child，保留其他 effect siblings。Outer/inner共享 color/opacity/geometry，outer额外支持 `rotateWithShape`。

该方案与现有 `ShapeFill` / `ShapeLine` / `ShapeArrows` 的 immutable snapshot和 whole-replacement模式一致，同时完整覆盖 PptxGenJS合法 shape-shadow能力。Generic effect stack留给独立高级小项。

## 4. 公共 API

```ts
export interface ShapeShadowBase {
  readonly color?: RichTextColor;
  readonly opacity?: number;
  readonly blur?: number;
  readonly angle?: number;
  readonly distance?: number;
}

export type ShapeShadow =
  | (ShapeShadowBase & {
      readonly kind: 'outer';
      readonly rotateWithShape?: boolean;
    })
  | (ShapeShadowBase & {
      readonly kind: 'inner';
      readonly rotateWithShape?: never;
    });

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly arrows?: ShapeArrows;
  readonly hyperlink?: Hyperlink;
  readonly shadow?: ShapeShadow;
}

class ShapeModel {
  get shadow(): ShapeShadow | undefined;
  set shadow(value: ShapeShadow | undefined);
}
```

Native使用 `distance` 而不是 PptxGenJS legacy `offset`，与 OOXML `dist` 和 PowerPoint UI术语一致。`opacity` 保持 `0..1`，与现有 `RichTextGlow.opacity` 一致；`blur`/`distance` 使用 points，`angle` 使用 degrees，color复用 strict `RichTextColor`。

Getter返回 detached、deep-frozen、fully normalized direct-state snapshot。Absent、unsupported或 ambiguous shadow返回 `undefined`。Setter是 whole replacement；`undefined` 清除唯一 supported direct shadow但保留 effect container和其他 effects。

## 5. 输入归一化与 defaults

所有 public input 在任何 package mutation 前完成：

- shadow必须是 ordinary 或 null-prototype object，只接受 own data properties；
- symbols、accessors、arrays、class instances、inherited-only fields和 unknown keys全部拒绝；
- `kind` 必须是 exact `outer` 或 `inner`；不接受 `none`、case variants或 PptxGenJS `type` alias；
- inner禁止 defined `rotateWithShape`；outer若提供则必须是 boolean；
- color若提供必须是合法 strict sRGB/theme `RichTextColor`；
- opacity必须是 finite number `0..1`，量化到 `1/100000`；
- blur必须是 finite point number `0..100`，distance必须是 finite point number `0..200`，均量化到 1 EMU；
- angle必须是 finite degree number `0 <= angle < 360`，量化到 `1/60000°`；
- own runtime `undefined` 与 field omission等价；nested color和最终 value立即复制并冻结。

创建 defaults 对 outer/inner均为：

```ts
{
  color: { kind: 'srgb', value: '000000' },
  opacity: 0.75,
  blur: 8,
  angle: 270,
  distance: 4,
}
```

Outer另默认 `rotateWithShape: false`。Defaults与 PptxGenJS 4.0.1真实 shape output一致。Explicit zero不触发 fallback。Public value错误使用 `TypeError` / `RangeError`，并在 slide XML、shape ID、mutation journal或 live model map变化前失败。

## 6. Canonical OOXML

Outer canonical output：

```xml
<a:effectLst>
  <a:outerShdw sx="100000" sy="100000" kx="0" ky="0"
      algn="bl" rotWithShape="0" blurRad="101600"
      dist="50800" dir="16200000">
    <a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr>
  </a:outerShdw>
</a:effectLst>
```

Inner canonical output：

```xml
<a:effectLst>
  <a:innerShdw blurRad="101600" dist="50800" dir="16200000">
    <a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr>
  </a:innerShdw>
</a:effectLst>
```

Native outer保留 PptxGenJS的 neutral scale/skew和 bottom-left alignment，使合法 outer final state与其输出语义一致。Inner使用 matching legal closing tag。Theme color以 `a:schemeClr` 写出；PptxGenJS只接受 hex，这是 native strict extension。

## 7. Reader 与 supported direct state

Reader只检查 shape唯一 direct PresentationML `p:spPr`：

- `a:effectDag` 存在时 direct simple shadow不可安全判定，返回 `undefined`；
- direct DrawingML `a:effectLst` 必须至多一个，且不能同时出现 inner与outer；
- 唯一 shadow必须是 namespace-correct `a:innerShdw` 或 `a:outerShdw`，child order符合 effect-list schema；
- `blurRad` / `dist` / `dir` 若省略按 OOXML zero defaults读取，若存在必须是唯一 unqualified合法整数并分别满足 `0..100pt`、`0..200pt`、`0..<360°`；
- outer `rotWithShape` 若省略按 OOXML true读取，只接受唯一 unqualified `0/1/true/false`；
- outer neutral `sx/sy/kx/ky/algn` 可全部省略，或存在时只能使用 PptxGenJS-compatible `100000/100000/0/0/bl`；这些兼容 attributes不进入 public snapshot，但在同-kind编辑中逐字节保留；
- inner不允许 outer-only `sx/sy/kx/ky/algn/rotWithShape`；
- shadow必须恰有一个 sRGB或theme color child；color element只允许唯一 `val`，并可有零个或一个合法 `a:alpha`；alpha absent按 opacity 1读取；
- unsupported color transform、preset/system color、extra/qualified/repeated owned attribute、child、wrong namespace、multiple shadow或 schema-order violation返回 `undefined`；mutation以 `ModelParseError` zero-change拒绝；
- 同一 effect list中合法 `blur/fillOverlay/glow/prstShdw/reflection/softEdge` siblings不进入 snapshot，但必须原样保留。

Reader接受 alternate legal DrawingML prefix，lexical prefix不进入 public value。Public snapshot永远包含 normalized `kind/color/opacity/blur/angle/distance`，outer另外包含 `rotateWithShape`。

## 8. Lossless replace、insert 与 clear

Setter先 normalize，再进入 package transaction：

- fully supported same kind/values是 exact bytes/journal no-op，包括省略 defaults、alternate prefix、missing alpha和 compatible outer attributes；
- same-kind edit只 patch changed owned attribute或 color/alpha span，保留 lexical attribute order、quote style、namespace declaration和 neutral outer attributes；
- outer↔inner切换 whole-replace shadow child，并按目标 kind输出 canonical child；
- absent shadow但 existing effect list安全时，按 effect-list schema order插入 inner/outer，保留所有 sibling bytes；
- absent effect list时，在 `a:ln` 后、`a:scene3d/a:sp3d/extLst` 前插入 canonical `a:effectLst`；存在 `effectDag` 或 unsafe property order时拒绝；
- clear只删除 direct inner/outer child；即使 effect list变空也保留 container，避免改变 unrelated direct-effect ownership或暴露 inherited effect；
- absent shadow clear是 exact no-op；multiple/unsafe state绝不猜测或部分修改。

Effect-list child stages固定为 `blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge`。新增和 kind切换必须保持该顺序。Shadow edit不拥有 effect siblings、shape-level style reference或任何 relationship。

## 9. 与其他 shape 能力的边界

Shadow位于 `p:spPr/a:effectLst`，与 geometry/fill/line/arrows和 `p:cNvPr` hyperlink分离 ownership：

- geometry/fill/line/arrows/hyperlink/text/transform edits必须原样保留 supported或unsupported effect bytes；
- shadow edit必须原样保留 geometry、fill、advanced line、arrows、hyperlink、text、non-visual metadata、scene3d/sp3d/extLst和neighbor bytes；
- AddShape renderer只在 supplied shadow存在时新增 effect list；omitted/runtime-undefined shadow保持已发布 default shape bytes完全不变；
- `ShapeModel.shadow` 可读取/编辑任何可解析 `p:sp`，包括 existing text shape；`addText()` shadow creation仍属于后续 text parity小项。

`effectDag`、multiple effect lists、multiple inner/outer shadow或 unsupported shadow transform可以在 unrelated edits中无损保留，但本 setter不会修改它们。

## 10. 生命周期、错误与无损约束

- Creation/edit都在 existing OPC transaction中完成；任何失败或 outer rollback恢复 exact slide bytes、parts、mutation journal、shape order与 live identity。
- Same-value assignment不写 part、不改变 journal。
- Duplicate、move、write/reopen与六种格式保持 supported snapshot、sibling effects和raw unrelated XML。
- Shadow没有 relationship；future picture/effect relationships不进入本小项。
- Existing malformed shadow getter返回 `undefined`；setter抛包含 slide part URI的 `ModelParseError`且零 mutation。

## 11. 验证策略

### Internal/value contract

- outer/inner defaults、all custom fields、zero/min/max/fractional quantization；
- sRGB/theme color、alpha absent/zero/full和 alternate prefix；
- descriptor-safe、null-prototype、detachment、deep freeze与 getter-free rejection；
- invalid kind/color/opacity/blur/angle/distance/rotate、unknown/symbol/accessor field零 mutation拒绝；
- optional OOXML defaults、compatible outer attributes、effect siblings与 canonical render/read/equality。

### Creation and editing

- omitted/runtime-undefined/outer/inner creation与 immediate snapshot；
- same-value no-op、每个 field replacement、color type switch、outer↔inner、whole clear；
- effect-list absent insertion、existing sibling insertion、alternate prefix和 container preservation；
- fill/line/arrows/hyperlink/shadow bidirectional ownership isolation；
- duplicate、rollback、write/reopen、six-format lifecycle和 stable identity；
- effectDag、multiple lists/shadows、invalid attributes/color/alpha/order/namespace zero-mutation rejection。

### PptxGenJS and release evidence

- public-only fixture覆盖 omitted/none/outer defaults/custom/zero、inner、rotateWithShape、hash/invalid color、unknown type和 out-of-range numbers；
- supported outer final semantics与 PptxGenJS 4.0.1 public output对照；zero preservation、theme color、valid inner和 strict rejection差异单独断言；
- PptxGenJS malformed inner fixture由 validator精确诊断，native inner必须是 well-formed并通过 PowerPoint 2010 profile；
- actual tarball Node/browser/types smoke覆盖 create/read/edit/clear和 effect sibling preservation；
- `pnpm check`、performance、build、PowerPoint 2010 validation、LibreOffice/Poppler visual comparison、overflow和 artifact-tool import全通过。

## 12. 完成门禁

只有以下条件全部满足，本小项才能把 shape shadow标记为已支持：

1. Preset shape可 create/read/replace/clear/reopen合法 outer和inner shadow；
2. 全参数、zero/default direct state、other-effect preservation、same-value、rollback和 identity有 exact tests；
3. Ambiguous/malformed state、PptxGenJS invalid runtime与 inner closing-tag缺陷有 zero-mutation/validator证据；
4. PptxGenJS public-output、packed Node/browser/types、full suite、validator和真实文件视觉 QA全通过；
5. README、CHANGELOG、API docs和 compatibility matrix明确 supported scope、strict divergences，以及剩余 advanced effects和非-shape shadow缺口。
