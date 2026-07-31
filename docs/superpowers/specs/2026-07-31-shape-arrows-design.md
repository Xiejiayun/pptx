# Shape Arrows Design

日期：2026-07-31

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项为原生 preset shape 增加 PptxGenJS 4.0.1 `ShapeLineProps` 的 begin/end arrow type vertical slice：

- shape 创建时可独立指定 begin、end 或两端 arrow type；
- 支持 `none | arrow | diamond | oval | stealth | triangle` 六种公开类型；
- 读取 existing deck 中唯一、direct、合法的 `a:headEnd` / `a:tailEnd`；
- 对新建或 existing shape 整体替换、单端替换或清除 direct arrow children；
- 与 `ShapeLine` 分离 ownership，使 line width/fill/dash 与 arrows 可以独立编辑和清除；
- 保持 advanced line fill/dash、join、合法 arrow size、extensions、live identity、shape order、transaction rollback、duplicate、write 和 reopen。

本设计不实现 arrow width/length size、line cap、compound/alignment、join/miter、custom dash、advanced line fill creation、shadow、hyperlink、adjustments、custom geometry、shape text options 或 percentage positions。PptxGenJS 4.0.1 的公开类型也把 begin/end arrow size 标为 future。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `ShapeLineProps` 类型与 `addShape()` / `write()` 真实输出。审计结果如下：

- `beginArrowType` 写 `a:headEnd@type`，`endArrowType` 写 `a:tailEnd@type`；
- 六种公开 token 均原样输出：`none`、`arrow`、`diamond`、`oval`、`stealth`、`triangle`；
- 两端同时存在时，child order 为 `headEnd` 后 `tailEnd`；
- 指定 color 的 line 会先输出 width、solid fill、preset dash，再输出 arrow children；
- 只指定 arrow 时，PptxGenJS 隐式补齐 `333333`、1pt、solid line；
- `{ type: 'none', beginArrowType: 'triangle' }` 只输出 `<a:ln><a:headEnd type="triangle"/></a:ln>`，不输出 `a:noFill`；
- runtime empty string arrow 被忽略，unknown token 被原样写入非法 OOXML；
- nested `line.lineHead` / `line.lineTail` 虽在 deprecated type 中声明，但 4.0.1 runtime 忽略；top-level deprecated `lineHead` / `lineTail` 会映射到 head/tail；
- PptxGenJS 不公开 arrow size；合法 OOXML `w` / `len` 使用 `sm | med | lg`。

原生 API 采用严格、可逆的 direct-state 语义，不复制隐式 line default、empty-string ignore、invalid passthrough 或 deprecated aliases。需要复现 PptxGenJS arrow-only final state时，调用方显式组合 default `ShapeLine` 与 `ShapeArrows`；需要复现 PptxGenJS `type: 'none'` 加 arrow 的 final state时，只传 `arrows`，不传 native line。

## 3. 方案比较

### 方案 A：把 begin/end 加进 `ShapeLine`

这最接近 PptxGenJS 的 input shape，但会破坏现有承诺：`shape.line = undefined` 当前只清除 width/fill/dash 并保留 arrows。扩展同一 whole-replacement value 后，旧调用方无法判断 clear 是否应删除 arrows，也会把 line paint 与 arrow endpoint ownership 再次耦合。

### 方案 B：分别暴露 `beginArrow` / `endArrow` 两个属性

单端操作直观，但缺少原子 two-end snapshot；同时修改两端需要两次 transaction，容易出现中间状态，也不符合库内 fill、line 等不可变快照的编辑模式。

### 方案 C：独立 `ShapeArrows` whole-replacement snapshot（采用）

新增 `ShapeArrows`，包含可选 `begin` / `end`。`AddShapeOptions.arrows` 与 `ShapeModel.arrows` 共享同一 detached value；setter 是 whole replacement，missing side 表示清除该 side，`undefined` 或空 snapshot 清除两端。内部 adapter 只拥有 direct `headEnd` / `tailEnd`，保留 line 的其他 state。

该方案保持已有 `ShapeLine` public contract 完全不变，两套 editor 可组合、可独立清除，也为未来 arrow size 扩展保留明确边界。

## 4. 公共 API

```ts
export type ShapeArrowType =
  | 'none'
  | 'arrow'
  | 'diamond'
  | 'oval'
  | 'stealth'
  | 'triangle';

export interface ShapeArrows {
  readonly begin?: ShapeArrowType;
  readonly end?: ShapeArrowType;
}

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly arrows?: ShapeArrows;
}

class ShapeModel {
  get arrows(): ShapeArrows | undefined;
  set arrows(value: ShapeArrows | undefined);
}
```

Getter 返回 detached direct-state snapshot。两端都不存在时返回 `undefined`；explicit `type="none"` 返回对应的 `'none'`，不会与 child absence 合并。

Setter 使用 whole-replacement 语义：`{ begin: 'triangle' }` 写 begin 并清除 existing end；`{ end: 'arrow' }` 写 end 并清除 existing begin；`{}`、`{ begin: undefined, end: undefined }` 与 `undefined` 都清除两端。单独替换一端且保留另一端时，调用方传入 getter snapshot 展开后的完整目标值。

## 5. 输入归一化

所有 public input 在任何 package mutation 前完成：

- arrows 必须是 ordinary 或 null-prototype object，只接受 own data properties；
- symbols、accessors、arrays、class instances、inherited-only fields与 unknown keys 全部拒绝；
- 只允许 `begin` / `end`，每个 defined value 必须是六个 canonical token 之一；
- missing 与 own runtime `undefined` side 都归一化为 absence；
- empty normalized snapshot 合法，表示 clear both；
- normalized value 冻结并与 caller 脱离。

Native 不接受 `beginArrowType`、`endArrowType`、`lineHead`、`lineTail` aliases，不接受 empty string、case variants、numeric enum、null 或 unknown token。Invalid creation/edit 在 mutation 前抛 `TypeError`，不读取 accessor。

## 6. OOXML 与 direct-state 语义

Native arrows 的 canonical output：

```xml
<a:ln>
  <a:headEnd type="triangle"/>
  <a:tailEnd type="arrow"/>
</a:ln>
```

与 simple line 组合时：

```xml
<a:ln w="31750">
  <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
  <a:prstDash val="dashDot"/>
  <a:headEnd type="stealth"/>
  <a:tailEnd type="oval"/>
</a:ln>
```

Reader 只检查 shape 唯一 direct `p:spPr` 下唯一 direct DrawingML `a:ln`：

- line absent、empty 或无两端 child → `undefined`；
- 唯一 direct `headEnd` / `tailEnd` 的 unique unqualified `type` 必须是六个 canonical token之一；
- endpoint 可带合法 optional unqualified `w` / `len`，值只能是 `sm | med | lg`；它们不进入 snapshot；
- `w="med" len="med"` 与省略 size 均表示默认-size compatible state；`sm` / `lg` 被无损识别和保留，但本 API 不宣称可创建或读取 arrow size value；
- endpoint 不允许 child elements、重复/qualified/unknown attributes、wrong namespace、unknown type 或 schema-order violation；上述状态 getter 返回 `undefined`，mutation 以 `ModelParseError` zero-change 拒绝；
- alternate legal DrawingML prefix 可读取和编辑，lexical prefix 不进入 public snapshot。

Arrow type replacement在 existing endpoint 上只改 `type` attribute value，保留合法 `w` / `len` 与 lexical form。新增 endpoint 只写 `type`，不物化 size attributes。

## 7. Lossless replace、clear 与 line 边界

Setter 先 normalize，再进入 package transaction：

- fully supported same arrow types 是 exact bytes/journal no-op，即使 existing endpoint 带合法 `sm` / `lg` size；
- target missing side 删除对应整个 endpoint child；
- existing target side 原位只替换 `type` attribute；
- absent target side 按 schema order 插入：join 后、对应 peer 前后、`extLst` 前；两端同时创建固定为 head 后 tail；
- direct line absent且 target 非空时，在 shape fill choices 后、effect/scene/3D/extLst 前插入 canonical `a:ln`；
- direct line absent且 target empty，或 existing line已无 endpoints，clear 是 exact no-op；
- repeated `spPr` / `ln` / endpoint、wrong namespace、invalid endpoint attributes/content、head/tail reversed order 或 unsafe insertion order 以 `ModelParseError` zero-mutation 拒绝。

Arrow edit 保留 `a:ln@w`、fill choice、dash choice、cap/compound/alignment attributes、round/bevel/miter join、extensions 与 unknown bytes。Existing gradient/pattern/picture/group line fill或 custom dash 不阻止合法 arrow type read/edit。

反向边界保持不变：`shape.line = undefined` 仍只清除 line width/fill/dash并保留 endpoints；`shape.arrows = undefined` 只清除 endpoints并保留 width/fill/dash。`{ kind: 'none' }` line 可与 endpoints 共存，这是 direct state，不保证 arrow 有可见 stroke。

Creation renderer 在同一 `a:ln` 中按 fill、dash、join、head、tail、extension schema order输出。Arrows-only creation不合成 color、width或 dash；line-only creation bytes 与已发布实现完全一致。

## 8. 生命周期、错误与无损约束

- Public value errors 使用 `TypeError`；malformed existing OOXML 使用 `ModelParseError` 并携带 slide part URI。
- Creation/edit 都在 existing OPC transaction中完成；任何失败或 outer rollback恢复 exact slide bytes、relationships、parts、mutation journal、shape order与 live identity。
- Same-value assignment不写 part、不改变 journal。
- Duplicate、move、write/reopen与六种格式保持 supported snapshot、合法 size attrs和 raw unrelated XML。
- Existing unsupported line paint/dash在 arrow mutation中原样保留；existing unsupported arrow在 geometry/fill/line/text/transform等无关 mutation中原样保留。
- Shape arrows不拥有 relationships；line picture fill 的 relationship lifecycle不进入本小项。

## 9. 验证策略

### Internal/value contract

- undefined、empty、begin-only、end-only、both与全部六种 token；
- descriptor-safe、null-prototype、detachment与 getter-free rejection；
- exact head/tail render、decode、equality与 alternate prefix；
- size omitted/med/sm/lg preservation；
- invalid token、attribute、child、namespace、duplicate和 wrong order严格拒绝。

### Creation and editing

- omitted/runtime-undefined/empty/begin/end/both creation；
- 与 none/solid line、8 dash、advanced line fill/custom dash、join/ext组合；
- immediate detached `ShapeModel.arrows` snapshot、stable identity与 caller detachment；
- same-value no-op、whole replacement、single-side replacement、clear、absent-line insertion；
- line/arrows bidirectional ownership isolation；
- duplicate、rollback、write/reopen、six-format lifecycle；
- malformed/ambiguous input与 existing XML zero mutation。

### PptxGenJS and release evidence

- public-only fixtures覆盖 begin/end每个 token、both、arrow-only defaults、none-line arrow、empty string、nested/top-level aliases与 invalid passthrough；
- supported endpoint final semantics 与 PptxGenJS 对照，implicit line defaults和 strict divergences单独断言；
- actual tarball Node/browser/types smoke覆盖 creation/read/edit/clear和 line组合；
- `pnpm check`、performance、build、PowerPoint 2010 validation、LibreOffice visual render、overflow与 artifact-tool import全通过。

## 10. 完成门禁

只有以下条件全部满足，本小项才能把 shape arrow types标记为已支持：

1. preset shape可 create/read/replace/clear/reopen两端六种 arrow types；
2. line/arrows independent ownership、explicit none与 absence、size preservation有 exact tests；
3. ambiguity、same-value no-op、rollback、identity、unknown-byte preservation和 malformed zero-mutation有证明；
4. PptxGenJS public-output、packed Node/browser/types、full suite、validator和 visual QA全通过；
5. README、CHANGELOG、compatibility matrix明确 supported scope、strict divergences，以及剩余 arrow size/shadow/hyperlink/advanced line缺口。
