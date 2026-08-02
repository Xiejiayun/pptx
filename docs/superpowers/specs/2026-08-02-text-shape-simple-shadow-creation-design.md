# Text Shape Simple Shadow Creation Design

日期：2026-08-02

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `slide.addText(text, { shadow })` 对应的原生创建能力：

- `AddTextOptions.shadow` 接受现有 strict `ShapeShadow` whole-replacement value；
- 支持 outer/inner、sRGB/theme color、opacity、blur、angle、distance，以及 outer-only `rotateWithShape`；
- plain text、rich text、layout/master text、declarative master text 与 text placeholder creation/population 共用同一 contract；
- 创建结果立即通过现有 live `ShapeModel.shadow` 读取、替换或清除；
- shadow 与 fill/line/arrows 保持独立 ownership，duplicate、rollback、write/reopen 和六格式生命周期不变；
- 复用已经完成的 simple-shadow normalizer/renderer 和 shape-shadow live editor，不建立 text 专用 shadow type。

本小项不加入 generic effect stack、preset shadow、reflection、soft edge、effect DAG、custom shadow scale/skew/alignment、hyperlink、text geometry、`rectRadius`、`isTextBox` 或 `breakLine` 组合语义，也不扩展 image/table/chart/media shadow API。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `ShadowProps` 类型，以及 `addText()` / `write()` 的真实输出。审计结果如下：

- omitted shadow 与 `{ type: 'none' }` 都不写 direct effect；
- `{ type: 'outer' }` 写 outer shadow，实际 defaults 为 black、opacity `0.75`、blur `8pt`、angle `270°`、offset `4pt`、`rotWithShape="0"`；
- 合法 custom outer 的 blur/offset 以 point 转 EMU、angle 以 `1/60000°` 转换、opacity 以 `1/100000` 转换；
- `rotateWithShape: true` 在 text serializer 中被忽略，outer 始终写 `rotWithShape="0"`；
- runtime zero opacity/blur/angle/offset 被 falsy fallback 折叠为上述 defaults；
- missing/invalid type 会被 warning-correct 为 outer，hash-prefixed color 会被改写，angle/opacity 接受部分 coercion；blur/offset 仍可经过宽松路径；
- `{ type: 'inner' }` 开始写 `<a:innerShdw>`，却错误地以 `</a:outerShdw>` 闭合，使 slide XML 不合法。

Native 只对合法、可逆的 final-state 能力宣称对等。它不复制 coercion、warning correction、falsy fallback、ignored rotate flag 或 malformed inner closing-tag 缺陷。Native `undefined` 对应 omitted/`type: 'none'`；explicit zero 保持 direct zero；theme color 与 rotate true 是严格扩展。Outer defaults 与 PptxGenJS 4.0.1 真实 text output 一致。

## 3. 方案比较

### 方案 A：在 `AddTextOptions` 复制 PptxGenJS `ShadowProps`

调用形式接近 PptxGenJS，但会为 text 引入第二套 `type/offset` value，并与已经公开的 `ShapeShadow.kind/distance`、live getter/setter 和 strict validation 分裂。它还容易复制 PptxGenJS 的 permissive fallback。

### 方案 B：创建 text 后调用 `ShapeModel.shadow` setter

实现较短，但一次 public creation 会产生两次 part mutation，中间无 shadow 状态可观察，outer transaction journal 更嘈杂，也让 placeholder replacement 与最终 renderer contract 分离。

### 方案 C：复用 `ShapeShadow` 并在 text renderer 直接组合（采用）

`validateAddTextOptions()` 调用现有 `normalizeShapeShadow()`，把 detached normalized shadow 与 fill/line/arrows 一起传入 `textShapeXml()`。Renderer 在同一次 shape creation 中、line 之后写唯一 canonical `a:effectLst`。现有 transaction、placeholder owner replacement、stable identity 与 live editor 均保持不变。

该方案改动最小，让 preset shape 与 text shape 共享同一 public shadow value、defaults、strict validation 和 direct-state lifecycle。

## 4. 公共 API

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly arrows?: ShapeArrows;
  readonly shadow?: ShapeShadow;
  // existing fields remain unchanged
}
```

`ShapeShadow` 继续使用现有公开类型：

```ts
interface ShapeShadowBase {
  readonly color?: RichTextColor;
  readonly opacity?: number;
  readonly blur?: number;
  readonly angle?: number;
  readonly distance?: number;
}

type ShapeShadow =
  | (ShapeShadowBase & {
      readonly kind: 'outer';
      readonly rotateWithShape?: boolean;
    })
  | (ShapeShadowBase & {
      readonly kind: 'inner';
      readonly rotateWithShape?: never;
    });
```

`AddPlaceholderOptions` 基于 `AddTextOptions`，因此 text placeholder creation/population 自动获得同一 `shadow` 字段。`SlideMasterObject` 的 text/placeholder options 同样复用该字段；`defineSlideMaster()` 的 closed option-key reader 必须显式允许 `shadow`。

Getter 和 setter 不新增 API：`ShapeModel.shadow` 继续提供 detached deep-frozen fully normalized snapshot、whole replacement 和 clear。`undefined` 表示没有 supported direct shadow或清除当前 direct shadow。

## 5. 归一化与错误语义

所有规则直接复用 `normalizeShapeShadow()`：

- omitted 或 runtime `undefined` 归一化为无 shadow；
- value 必须是 ordinary 或 null-prototype own-data object；只接受 `kind/color/opacity/blur/angle/distance/rotateWithShape`；
- kind 必须是 exact `outer` 或 `inner`；inner 禁止 defined `rotateWithShape`；outer rotate 若提供必须是 boolean；
- color 复用 strict sRGB/theme `RichTextColor`；opacity 为 finite `0..1`；blur 为 finite `0..100pt`；distance 为 finite `0..200pt`；angle 为 finite `0 <= value < 360°`；
- defaults 为 black、opacity `0.75`、blur `8pt`、angle `270°`、distance `4pt`，outer rotate false；
- explicit zero 保留；number 按现有 codec 量化到 OOXML units；nested color 与最终 value 立即复制并冻结；
- symbols、accessors、arrays、class instances、inherited-only fields、unknown keys、`type/offset` aliases、coercible strings 和 invalid ranges 在任何 package mutation 前拒绝；
- invalid plain/rich/placeholder/master definition 保持 parts、relationships、XML、runtime cache、shape order和 mutation journal 不变。

## 6. OOXML 与渲染顺序

Outer canonical output：

```xml
<p:spPr>
  <!-- transform, geometry, fill, line -->
  <a:effectLst>
    <a:outerShdw sx="100000" sy="100000" kx="0" ky="0"
        algn="bl" rotWithShape="0" blurRad="101600"
        dist="50800" dir="16200000">
      <a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr>
    </a:outerShdw>
  </a:effectLst>
</p:spPr>
```

Inner canonical output：

```xml
<a:effectLst>
  <a:innerShdw blurRad="101600" dist="50800" dir="16200000">
    <a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr>
  </a:innerShdw>
</a:effectLst>
```

Effect list 只在 supplied shadow 存在时创建，并位于现有 line/endpoints 之后、`p:spPr` 结束之前。Omitted/runtime-undefined shadow 必须保持当前 text XML byte-for-byte 不变。Theme color 使用 `a:schemeClr`；inner 必须使用 matching legal closing tag。

除目标 effect list 外，transform、geometry、shape fill、line width/fill/dash/endpoints、name、placeholder identity、body properties、paragraphs、runs、relationships 与相邻 shapes 都不改变。

## 7. Ownership 与生命周期

- `fill`、`line`、`arrows` 与 `shadow` 可同时创建，输出顺序为 geometry → fill → line/endpoints → effect list；
- `shape.fill = undefined`、`shape.line = undefined` 与 `shape.arrows = undefined` 均保留 shadow；
- `shape.shadow = undefined` 只删除 direct inner/outer child并保留安全 effect-list container及其他 effect siblings；
- same-value shadow assignment 是 exact bytes/journal no-op；same-kind edit保留可保留的 lexical bytes，kind switch只替换 shadow child；
- unsupported/ambiguous imported effect state只做无损保留，getter 返回 `undefined`，setter zero-change拒绝；
- `addText()`、`addRichText()`、`addPlaceholder()`、layout/master wrappers 与 declarative master definitions 都走同一 renderer；
- placeholder owner materialization/population 只改变当前 owner，layout/master source 保持隔离；
- duplicate、move、outer rollback、write/reopen 与六种 presentation format 保留 supported direct state；
- shadow 没有 relationship ownership，不引入 relationship 或 GC 分支。

## 8. 验证策略

### Model/SDK

- omitted/runtime-undefined、outer/inner defaults、all custom fields、explicit zero、sRGB/theme 与 rotate true/false；
- exact effect-list order、line/arrows/shadow ownership isolation、immediate live snapshot、caller detachment、same-value no-op、whole replacement 与 clear；
- plain/rich、placeholder、layout/master direct creation 与 declarative definition；
- invalid object/accessor/symbol/alias/value/range 的 zero mutation；
- duplicate、rollback、write/reopen、六格式、stable identity、effect siblings 与 unrelated XML isolation。

### PptxGenJS

- public-only fixtures覆盖 omitted、none、outer defaults、custom outer、zero fallback、rotate ignored、inner malformed、missing/invalid type、hash color 与 coercion；
- imported legal outer `ShapeModel.shadow` 与 raw direct effect state对照；inner malformed package只记录生成缺陷，不作为合法 import fixture；
- supported outer final semantics对等；native legal inner、explicit zero、theme color、rotate true与 strict rejection作为 intentional differences锁定。

### Release gates

- focused model/SDK/adapter/root tests；
- full Vitest、performance、typecheck、build；
- actual tarball Node/types/browser/CLI smoke覆盖 `addText({ shadow })`；
- real Chrome immediate/detached/reopen state、PowerPoint 2010 profile validation与文档/compatibility matrix更新。

## 9. 完成门禁

只有以下条件全部满足，本小项才标记完成：

1. 所有 text creation surfaces 接受并写入 strict simple shadow；
2. validation-before-mutation、detachment、effect ownership、placeholder isolation、identity与 lifecycle有测试证明；
3. PptxGenJS public-output supported cases和 intentional differences有锁定证据；
4. packed Node/types/browser/CLI、full suite、performance、typecheck、build与 validator全通过；
5. README、API、compatibility、progress与 changelog明确完成范围及下一项。
