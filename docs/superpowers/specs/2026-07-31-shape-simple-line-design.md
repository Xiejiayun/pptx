# Shape Simple Line Design

日期：2026-07-31

状态：已确认（用户已授权自主决策）

## 1. 目标与范围

本小项为原生 preset shape 增加 PptxGenJS 4.0.1 `ShapeLineProps` 的 simple-line vertical slice：

- shape 创建时接受 strict none 或 solid line；
- solid line 支持 sRGB/theme color、0–100% transparency、0–1584pt width 与 8 种 preset dash；
- 读取 existing deck 中唯一、direct、合法的 simple line；
- 对新建或 existing shape 替换、清除 direct line paint/width/dash；
- 保持 arrowhead、join、扩展节点、live identity、shape order、transaction rollback、duplicate、write 和 reopen；
- 为后续 shape arrows、outer text line 与其他 drawing object line 提供经过验证的 shape 容器语义。

本设计不实现 begin/end arrows、arrow size、line cap、compound/alignment、join/miter、custom dash、gradient/pattern line fill、shadow、hyperlink、adjustments、custom geometry、shape text 或 percentage positions。Arrowheads 是紧随本小项之后的独立对等项。

## 2. PptxGenJS 公开行为证据

权威基线为锁定依赖 `pptxgenjs@4.0.1` 的 public `ShapeLineProps` 类型与 `addShape()` / `write()` 真实输出。审计结果如下：

- omitted `line` 与 `{ type: 'none' }` 都输出 empty `<a:ln></a:ln>`；
- empty line 或 `{ type: 'solid' }` without color 回退为 `333333`、1pt、solid；
- `{ color: 'FF0000' }` 输出 1pt sRGB solid line；
- theme color 输出 `a:schemeClr`；
- `transparency: 50` 输出 `a:alpha@val="50000"`；explicit zero 省略 alpha；
- `width: 0` 通过 falsy fallback 变成 1pt；正 width 按 point 写入 `a:ln@w`；
- `dashType` 的 `solid`、`dash`、`dashDot`、`lgDash`、`lgDashDot`、`lgDashDotDot`、`sysDash`、`sysDot` 原样写入 `a:prstDash@val`；
- deprecated `alpha` 和 `lineDash` 在 shape line 上被静默忽略；
- `beginArrowType` / `endArrowType` 分别写入 `a:headEnd` / `a:tailEnd`，与 fill/width/dash 共用一个 `a:ln` 容器。

原生 API 采用严格、可逆的 direct-state 语义，不复制 permissive/falsy fallback：empty/missing-color/unknown/deprecated fields 在 mutation 前拒绝；explicit none 写 direct `a:noFill`；explicit transparency zero 写 `alpha="100000"`；width zero 写 direct `w="0"`。Supported explicit color/theme/nonzero transparency/positive width/8 dash 的 final semantics 与 PptxGenJS 对等；direct bytes 有意保留上述差异。

## 3. 方案比较

### 方案 A：直接复用 table-cell border

`TableCellBorder` 已有 none/line、color、width 与两种 style，但它只接受 solid/sysDash、没有 transparency，且 renderer 强制写 cap/compound/alignment/round/headEnd/tailEnd。Shape line 的 8 种 dash、PptxGenJS defaults 与 direct child preservation 不同。直接复用会让 table contract 泄漏到 shape，并引入无关 bytes。

### 方案 B：一次实现 line + arrows + join + shadow

把整个 PptxGenJS `ShapeLineProps` 与邻近样式一次建模，可减少未来 public type 变更，但会把 paint、几何端点、line join 和 effect tree 的独立 ownership 绑成一个大 mutation，难以证明 lossless isolation，也不符合逐小项 review/commit 的执行方式。

### 方案 C：shape 专用 simple-line vertical slice，复用 simple-fill（采用）

新增 internal simple-line value codec 与 shape line container adapter。Color/transparency 的 strict normalize/decode/render 复用现有 `simple-fill`；width/dash/defaults 与 `a:ln` 子节点次序由 line codec 负责；`p:spPr` 导航、safe insert 和 package transaction 由 shape adapter 负责。Arrow/join 等 sibling state 不进入 snapshot，但在 replace/clear 中原样保留。

该方案不改 table border 行为，能独立完成 PptxGenJS simple-line 对等，并为下一项 arrows 预留同一 `a:ln` 容器的无损扩展点。

## 4. 公共 API

```ts
export type ShapeLineDash =
  | 'solid'
  | 'dash'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot';

export type ShapeLine =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly transparency?: number;
      readonly width?: number;
      readonly dash?: ShapeLineDash;
    };

export interface AddShapeOptions extends Partial<Transform> {
  readonly name?: string;
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
}

class ShapeModel {
  get line(): ShapeLine | undefined;
  set line(value: ShapeLine | undefined);
}
```

Input 中 omitted `width` 与 `dash` 分别 materialize 为 1 point 与 `solid`，因此 native configured line 与 PptxGenJS canonical output 对等。Getter 只返回 fully supported direct line，并始终包含 normalized `width` 与 `dash`；类型保持 optional 是为了让 creation/edit input 可使用 documented defaults。

Omitted/runtime-`undefined` creation 保持现有 canonical `<a:ln/>` skeleton 和 `line === undefined`。Explicit `{ kind: 'none' }` 写 direct no-fill。Editing `undefined` 清除本 API 拥有的 width/fill/dash state，并保留 `a:ln` 容器及 arrow/join 等无关 state；若无无关 state，结果是 canonical `<a:ln/>`。

## 5. 输入归一化

所有 public input 在任何 package mutation 前完成：

- line 必须是 ordinary 或 null-prototype object，只接受 own data properties；
- symbols、accessors、arrays、class instances、inherited-only fields 与 unknown keys 全部拒绝；
- `kind: 'none'` 只允许 `kind`；
- `kind: 'line'` 必须提供 `color`，只允许 `kind/color/transparency/width/dash`；
- color 使用 simple-fill 的 strict 六位 sRGB 或 supported scheme token；
- transparency 必须是 finite `0..100`，量化到最近 `0.001%`；
- width 必须是 finite `0..1584` point，量化到最近 1 EMU；omitted materialize 为 1 point；
- dash 必须是 8 个 canonical token 之一；omitted materialize 为 `solid`；
- normalized value 与 nested color 深度脱离 caller。

Native 不接受 PptxGenJS 的 `type`、`alpha`、`dashType`、`lineDash`、`lineHead`、`lineTail` aliases，也不对 empty/missing color 或 width zero做 falsy fallback。Public native property 使用 `dash`，compatibility matrix 明确映射 PptxGenJS `dashType`。

## 6. OOXML 与 direct-state 语义

Native simple line 的 canonical 输出：

```xml
<a:ln><a:noFill/></a:ln>

<a:ln w="31750">
  <a:solidFill>
    <a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr>
  </a:solidFill>
  <a:prstDash val="dashDot"/>
</a:ln>
```

Reader 只检查 shape 唯一 direct `p:spPr` 下唯一 direct DrawingML `a:ln`：

- absent 或 empty line → `undefined`；
- unique direct `noFill`、无 owned width/dash → `{ kind: 'none' }`；
- unique strict solid fill + strict integer `w` + unique canonical `prstDash` → detached line snapshot；
- solid color/alpha 规则与 simple-fill 一致；
- line width 必须是 safe integer `0..20116800` EMU；
- arrowhead、join 与 extension children 可共存且不进入 simple snapshot；
- unsupported fill/custom dash、namespace lookalike、conflicting/repeated fill/dash/line/spPr 或 malformed owned state → `undefined`，读取不修改 bytes。

Canonical/default `cap="flat"`、`cmpd="sng"`、`algn="ctr"` 可被 reader 接受并由 setter 保留；其他 direct line attributes 不伪装成 fully supported simple line。

## 7. Lossless replace、clear 与 arrows 边界

Setter 先 normalize，再进入 package transaction：

- fully supported same value 是 exact bytes/journal no-op；
- `undefined` 删除 owned `w`、line fill choice 与 preset/custom dash choice；
- none 删除 owned width/dash 并写 direct no-fill；
- solid line 写 normalized width、whole-replace line fill choice 与 preset dash；
- line absent 时，在 shape fill choices 后、effect/scene/3D/extLst 前插入 canonical `a:ln`；
- existing unique gradient/pattern line fill 或 custom dash 可被 explicit simple replace/clear；
- repeated direct `spPr`/`ln`、repeated/conflicting line fill/dash choice，或 unsafe insertion order 以 `ModelParseError` zero-mutation 拒绝。

`headEnd`、`tailEnd`、round/bevel/miter join、canonical unrelated attributes 和 extensions 在 line replace/clear 时保持原 bytes、relative order 与 namespace prefix。下一 arrows 小项只拥有 head/tail children；它不需要改写本次 public line contract。Explicit no-line 可能与 preserved arrow nodes 共存，这是 direct-state preservation，不代表 arrow 有可见 stroke。

Renderer 使用 source in-scope DrawingML prefix，不硬编码 imported prefix。除 owned line fields外，geometry/adjustments、shape fill、effects、text、non-visual properties、extensions、relationships、neighbor shapes 和其他 parts 保持原 bytes。

## 8. 生命周期、错误与无损约束

- Public value errors 使用 `TypeError` / `RangeError`；malformed existing OOXML 使用 `ModelParseError` 并携带 slide part URI。
- Creation/edit 都在 existing OPC transaction 中完成；任何失败或 outer rollback 恢复 exact slide bytes、relationships、parts、mutation journal、shape order 和 live identity。
- Same-value assignment 不写 part、不改变 journal。
- Duplicate、move、write/reopen 与六种格式保持 supported snapshot 和 raw XML。
- Existing unsupported line 在 geometry/fill/text/transform 等无关 mutation 中原样保留。
- Shape simple line 不拥有 relationships；image/pattern line fill 的 relationship lifecycle 不进入本小项。

## 9. 验证策略

### Internal/value contract

- none、sRGB、scheme、transparency 0/50/100 与 fractional quantization；
- width omitted/0/fractional/max 与 EMU quantization；
- 8 种 dash exact render/decode；
- invalid objects/colors/transparency/width/dash/accessors/symbols 在 mutation 前拒绝；
- alternate prefixes、namespace lookalikes、unsupported line fill/custom dash 与 duplicate choices 有 strict tests；
- simple-fill/table-cell fill/border regression 保持全绿。

### Creation and editing

- omitted/undefined/none creation，以及 line color/theme/transparency/width/dash defaults 与 explicit values；
- immediate detached `ShapeModel.line` snapshot、stable identity 和 caller detachment；
- same-value no-op、replace、clear、unsupported unique state replacement；
- arrow/join/ext preservation、unsafe ambiguity zero mutation；
- geometry/fill/line/effects/text/ext ordering 与 mutation isolation；
- duplicate、rollback、write/reopen、six-format lifecycle。

### PptxGenJS and release evidence

- public-only fixtures覆盖 omitted、none、empty/missing-color、sRGB、scheme、transparency 0/50、width 0/positive、8 dash、deprecated aliases 与 arrows coexistence；
- supported final semantics 与 PptxGenJS 对照，strict divergences 单独断言；
- actual tarball Node/browser/types smoke 覆盖 creation/read/edit/clear；
- `pnpm check`、performance、build、PowerPoint 2010 validation、LibreOffice visual render、overflow 与 artifact-tool import 全通过。

## 10. 完成门禁

只有以下条件全部满足，本小项才能把 shape simple line 标记为已支持：

1. preset shape 可 create/read/replace/clear/reopen none 或 solid line；
2. color/transparency/width/8 dash defaults 与 direct state 有 exact tests；
3. ambiguity、same-value no-op、rollback、identity、arrow preservation 和 unknown-byte preservation 有证明；
4. PptxGenJS public-output、packed Node/browser/types、full suite、validator 和 visual QA 全通过；
5. README、CHANGELOG、compatibility matrix 明确 supported scope、strict divergences，以及剩余 arrows/shadow/hyperlink/advanced line fills 缺口。
