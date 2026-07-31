# Custom Geometry Guide Formulas Design

## 1. 目标

在现有 custom geometry paths API 上增加 DrawingML `a:avLst` adjustment guides、
`a:gdLst` shape guides，以及 path coordinate/radius/angle 对 guide token 的引用。
本小项覆盖从零创建、existing-deck strict direct-state 读取、whole-replace 编辑、
preset/custom 转换、duplicate/rollback/write/reopen 和六种 presentation formats。

本小项不实现 `a:ahLst` adjustment handles、`a:cxnLst` connection sites、非默认
`a:rect` custom text rectangle 或 geometry evaluation。它们继续无损保留；strict getter
返回 `undefined`，编辑在 package mutation 前拒绝，并按后续小项依次开放。

用户已经授权由实现方连续决定最佳方案并执行，因此本设计完成 self-review 后直接进入计划与
实现，不设置交互式确认点。

## 2. 当前状态与问题

现有 `CustomGeometry` 只包含 numeric path tree：

- `a:avLst/a:gdLst` 只能 absent 或 empty；任一非空 list 会使 snapshot 返回
  `undefined`；
- point `x/y`、arc `wR/hR/stAng/swAng` 只接受 direct safe integer；
- renderer 固定输出 empty `avLst/gdLst`；
- handles、connections 与 custom text rectangle 同样被 strict reader 拒绝。

只开放 guide list 而不开放 path attribute 的 guide token 没有可用价值，因为公式结果无法参与
几何路径。反之，只允许 string coordinate 而不公开 formula tree 会迫使调用方引用不存在或只能
依赖 built-in 的 token。因此两者必须作为一个可独立验收的纵切面落地。

PptxGenJS 4.0.1 的 public `ShapeType.custGeom + points` 只生成 empty
`avLst/gdLst`，没有 arbitrary guide API。本功能是完整 DrawingML 创建/编辑所需的 native
extension；已有 PptxGenJS 合法输出必须继续得到原 snapshot。

## 3. 方案比较

### A. Raw `a:custGeom` XML

表达力最大，但绕过 namespace/schema order、XML escaping、detachment、deep freeze、
semantic no-op 和 transaction isolation，也无法为后续 handles/evaluation 提供稳定类型边界。
拒绝。

### B. Raw `fmla` string + arbitrary coordinate string

代码量最少，也能忠实表达 OOXML，但把 operator、arity、numeric coercion 和 whitespace
歧义推给调用方；malformed formula 只能到客户端或 validator 才暴露。它不符合现有 strict
API 的失败前置原则。拒绝。

### C. Typed operator/operand tree + direct numeric-or-token values（采用）

公式使用 17 个 OOXML operator token 的 discriminated union，并用 readonly tuple 在
TypeScript 层表达 1/2/3 operand arity。Coordinate、radius、angle 和 formula operand 统一
为 direct safe integer 或单个 guide/built-in token。该方案无 raw XML 旁路，能严格读取已有
文件，也为 handles、connections、custom rect 和 evaluator 复用同一 value/formula model。

## 4. 公共 API

```ts
export type CustomGeometryValue = number | string;

export type CustomGeometryFormula =
  | {
      readonly operator: 'val' | 'abs' | 'sqrt';
      readonly operands: readonly [CustomGeometryValue];
    }
  | {
      readonly operator: 'at2' | 'cos' | 'max' | 'min' | 'sin' | 'tan';
      readonly operands: readonly [CustomGeometryValue, CustomGeometryValue];
    }
  | {
      readonly operator:
        | '*/'
        | '+-'
        | '+/'
        | '?:'
        | 'cat2'
        | 'mod'
        | 'pin'
        | 'sat2';
      readonly operands: readonly [
        CustomGeometryValue,
        CustomGeometryValue,
        CustomGeometryValue,
      ];
    };

export interface CustomGeometryGuide {
  readonly name: string;
  readonly formula: CustomGeometryFormula;
}

export interface CustomGeometryPoint {
  readonly x: CustomGeometryValue;
  readonly y: CustomGeometryValue;
}

export interface CustomGeometry {
  readonly adjustments?: readonly CustomGeometryGuide[];
  readonly guides?: readonly CustomGeometryGuide[];
  readonly paths: readonly CustomGeometryPath[];
}
```

`CustomGeometryCommand` 中 `arcTo.widthRadius/heightRadius/startAngle/sweepAngle`
同步改为 `CustomGeometryValue`。Path `width/height` 保持 positive safe integer，因为
`a:path@w/h` 是 coordinate-space extent，不是 formula reference。

`adjustments` 映射 `a:avLst`，`guides` 映射 `a:gdLst`。两者省略或传 empty array
语义等价；normalized snapshot 只在 list 非空时包含对应 optional property，以保持现有
numeric-only snapshot 与源码兼容。`AddCustomShapeOptions` 继续排除 preset-only
`adjustments`，custom adjustment guides 只属于 geometry 第一个参数。

## 5. Formula 与 token contract

支持的 operator 和 arity 固定为 DrawingML guide formula grammar：

| Arity | Operators |
| --- | --- |
| 1 | `val`, `abs`, `sqrt` |
| 2 | `at2`, `cos`, `max`, `min`, `sin`, `tan` |
| 3 | `*/`, `+-`, `+/`, `?:`, `cat2`, `mod`, `pin`, `sat2` |

每个 formula object 必须是 ordinary 或 null-prototype object，只含 own data
`operator/operands`；operand tuple 必须是 exact-length dense ordinary array。未知
operator、错误 arity、accessor、symbol、extra key 或 array subclass 在 mutation 前拒绝。

`CustomGeometryValue` 的 number branch 必须是 finite safe integer，`-0` 归一化为
`0`。String branch 表达一个 guide name 或 DrawingML built-in token：

- 必须非空；
- 不得含 XML whitespace；
- 不得含 XML 1.0 非法字符；
- 不得是 signed decimal integer string，避免 `"1"` 与 `1` 双重表示。

本小项不解析 token dependency graph，也不拒绝 forward reference、cycle、unknown
built-in 或 domain-invalid arithmetic。OOXML schema 把 `fmla` 作为 required string，
而完整语义验证属于后续 evaluator；本层只保证明确的 operator/arity/operand lexical tree。

Guide `name` 使用同一 non-empty single-token/XML-safe contract，并在
`adjustments + guides` 的合并命名空间中要求唯一。重复名称会使创建/编辑拒绝、existing
snapshot 返回 `undefined`，避免后续 handle/evaluator 的引用歧义。

## 6. Normalization 与 snapshot

Normalization 顺序：

1. descriptor-safe 读取 root；
2. normalize optional adjustment guides；
3. normalize optional shape guides，同时检查全局 unique names；
4. normalize paths，并允许 point/arc attributes 使用 `CustomGeometryValue`；
5. deep-freeze guides、formula objects、operand tuples、paths 和 commands；
6. 返回 detached root snapshot。

空 `adjustments/guides` 不保留 own property。Caller 后续修改原数组、guide、formula、
operand 或 path 不影响 model。每次 getter 返回新的 detached deep-frozen snapshot；
same semantic assignment 是 exact package bytes/mutation journal no-op。

## 7. Deterministic OOXML

Renderer 保持 canonical child order：

```xml
<a:custGeom>
  <a:avLst>
    <a:gd name="adj1" fmla="val 25000"/>
  </a:avLst>
  <a:gdLst>
    <a:gd name="x1" fmla="*/ w adj1 100000"/>
    <a:gd name="y1" fmla="+- h 0 x1"/>
  </a:gdLst>
  <a:ahLst/>
  <a:cxnLst/>
  <a:rect l="l" t="t" r="r" b="b"/>
  <a:pathLst>
    <a:path w="100000" h="100000">
      <a:moveTo><a:pt x="x1" y="0"/></a:moveTo>
      <a:lnTo><a:pt x="r" y="y1"/></a:lnTo>
      <a:arcTo wR="x1" hR="hd2" stAng="0" swAng="cd2"/>
    </a:path>
  </a:pathLst>
</a:custGeom>
```

Absent/empty guide arrays render self-closing lists，保持当前输出。非空 formula 以单个 ASCII
space 连接 operator 与 operands；number 使用 canonical decimal，string/name 通过
`escapeXmlAttribute()`。Geometry 使用当前 in-scope DrawingML prefix，必要时只在 replaced
`custGeom` root 补最小 namespace declaration。

## 8. Strict reader

Reader 继续要求唯一 namespace-correct direct `p:sp/p:spPr/a:custGeom`、schema child
order、唯一 path list 和 supported path tree。对 `avLst/gdLst`：

- list 可 absent、empty 或各出现一次；必须无 non-namespace attribute、无 non-whitespace
  text；
- direct children 只能是 same-namespace `a:gd`；
- `a:gd` 必须 leaf，且恰有 unqualified `name/fmla` 两个 attribute；
- formula 用 XML whitespace tokenize，operator/arity 必须精确；
- signed decimal operand 转为 safe integer，其他 operand 走 token contract；
- 两个 lists 合并检查 guide name 唯一性。

Path point 与 arc attributes 同样先识别 signed decimal integer，再识别 token。Malformed、
unsafe integer、qualified lookalike、whitespace token、extra child/attribute、duplicate guide、
unsupported handle/connection 或非默认 rect 均返回 `undefined`，不修改 bytes。

Reader 对 lexical 差异做 semantic normalization，例如 `fmla="  val   +1 "` snapshot 为
`{ operator: 'val', operands: [1] }`。把该 snapshot 原样赋回必须通过 semantic equality
识别为 no-op，而不是 canonical rewrite。

## 9. Whole replacement 与转换

对 supported custom state：

- 同值赋值保持 exact bytes/journal no-op；
- 变化值 whole-replace 唯一 `custGeom`；
- shape identity、name、transform、fill、line、arrows、shadow、hyperlink、text、effects、
  extensions、order 和 relationships 不变。

对唯一 supported preset geometry 设置含 formulas 的 `customGeometry`，只替换 geometry
choice；对 supported custom geometry 设置 `presetType`，仍只生成 canonical preset geometry
和 empty preset adjustment list。Unsupported/mixed/malformed target 在 mutation 前抛
`ModelParseError`。

## 10. 测试与兼容门禁

### Internal codec

- 17 operators、1/2/3 arity、numeric 与 token operands；
- non-empty/empty/omitted `avLst/gdLst`；
- guide reference 覆盖 move/line/quadratic/cubic/arc fields；
- caller detachment、recursive freeze、global unique names；
- formula/name XML escaping、prefix preservation、semantic no-op；
- invalid objects、arrays、operator、arity、integer/token、XML state 零 mutation 拒绝；
- handles/connections/custom rect 仍保持 unsupported boundary。

### Public model

- create/read/edit、preset/custom conversion、identity/style/relationship preservation；
- duplicate/source isolation、outer transaction rollback、move/delete、write/reopen；
- 六种 presentation formats 和 unknown/opaque part preservation。

### PptxGenJS 4.0.1

- 原有 legal `custGeom + points` snapshot 不增加 empty optional properties；
- all-command、multiple-subpath、empty-path 和 malformed-output contract 继续通过；
- 明确记录 guide formulas 是 native extension，PptxGenJS 无对应 public input。

### Packaged/runtime/real PPTX

- typecheck、full unit suite、performance gate、build、dist diff；
- actual tarball Node/browser/types/CLI smoke 覆盖 guide formula 创建、读取、编辑；
- 生成 formula gallery，检查 snapshot、PPTX reopen、LibreOffice render/round-trip、overflow；
- PowerPoint 2010 validator 对原始和 round-trip 文件均要求 0 errors / 0 warnings；
- review 后按小项独立 commit、push，并确认 `HEAD...origin/main` 为 `0 0`。

## 11. 明确剩余边界

完成本项后仍未支持：

- adjustment handles；
- connection sites；
- custom text rectangle；
- formula evaluation、resolved bounds 和 handle constraint evaluation。

这些能力依次复用本项的 `CustomGeometryValue`、`CustomGeometryGuide` 和 formula model，
不再需要 raw string 或第二套 formula 表达。
