# Existing Table Row Height Editing Design

## 目标与范围

本小项为 `TableModel` 增加已有表格行高的严格读取和无损编辑，使调用方既能读取 PowerPoint direct `a:tr@h` 的 exact EMU 值，也能原子更新全部 physical rows。显式正行高会同步表格 transform 高度；包含自动行高的向量会保留现有 transform 高度，因为 `h="0"` 不提供可求和的实际渲染高度。

这是已有表格 geometry 编辑的第二个小项，沿用列宽编辑已经验证的 strict-read、descriptor-safe input、validation-before-write 和 minimal XML patch 模式。它只处理行高，不同时加入比例缩放、内容测量、auto-fit、分页、重复表头、行增删、merge 编辑或 cell creation 扩展。

## 公共 API

`TableModel` 新增：

```ts
get rowHeights(): readonly number[] | undefined;
setRowHeights(value: number | readonly number[]): void;
```

所有值均为 EMU。getter 每次返回 detached snapshot；调用方修改返回数组不得影响模型。setter 使用方法，因为 scalar 需要按当前 direct physical row 数广播，而数组需要做严格结构和长度验证。

`0` 是合法值，明确表示 OOXML 自动行高。getter 的 `undefined` 只表示无法从现有 table XML 得到唯一、完整、可信的行高向量。setter 不接受 `undefined`。

## 方案选择

考虑三个方案：

1. 允许 non-negative 行高；目标全部为正数时同步 `ext@cy`，任一目标为 `0` 时保留已验证的现有 `cy`。采用此方案。
2. setter 只接受正行高并总是同步 `cy`。它无法保留或主动设置 PowerPoint 常见的 `tr@h="0"` 自动行高，已有表格编辑能力不完整，不采用。
3. 允许 `0` 但总是把 `cy` 设置为向量数值之和。全自动向量会把表格 transform 高度写成 zero，混合自动/显式向量也会低估实际布局高度，不采用。

采用方案 1 后，调用方可把所有行切换为自动，也可混合自动和显式最小行高。`setTransform({ height })` 的现有行为不改变：它仍只修改 transform；要修改 row tokens，应使用 `setRowHeights()`。

## 严格读取

读取从当前 live `TableModel` 对应的 `p:graphicFrame` 开始，只沿 direct-child 路径解析：

```text
graphicFrame -> graphic -> graphicData -> tbl -> tr@h
```

每一级 `graphic`、`graphicData` 和 `tbl` 都必须唯一。节点按 direct `localName` 匹配，以兼容合法 alternate prefixes；descendant 中的同名 fake 节点不参与解析。

`tbl` 必须至少包含一个 direct `tr`。其他未知 direct child、注释、namespace declaration、属性、`tblPr`、`tblGrid` 和 extension data 不影响读取，并在编辑时原样保留。每个 direct row 都是一个 physical row；`gridSpan`、`rowSpan`、`hMerge` 和 `vMerge` 不改变行数，也不从 cell 数或 merge markup 猜测行数。

每个 direct `tr` 必须恰好包含一个 unqualified `h` attribute。token 必须只包含 ASCII 十进制数字，转换后为 non-negative safe integer。zero 和带 leading zeros 的 zero/positive token 合法；缺失、重复、namespaced-only、空值、小数、指数、带正负号或超过 safe integer 的 token 都使整个 getter 返回 `undefined`。getter 不返回部分向量，也不 fallback 到 transform height、平均分配或内容测量。

读取不检查 transform，因此 valid row vector 即使与 `a:ext@cy` 不一致也会返回。getter 不 mutation package bytes 或 journal。

## 输入归一化

setter 首先把输入复制为 detached normalized representation，再解析和修改 package：

- scalar 必须是 finite number，`Math.round()` 后为 non-negative safe integer；随后按当前 valid physical row 数精确广播。scalar `0` 把所有行设为自动。
- array 必须通过 `Array.isArray()`、非空且 dense；只允许 `length` 和从 `0` 到 `length - 1` 的 own data properties。hole、accessor、symbol key、额外字符串 key、typed array 或继承取值都拒绝，且不得调用 getter。
- array 每项使用与 scalar 相同的 finite、round、non-negative safe-integer 规则；数组可混合 zero 与正值。
- array 长度必须严格等于当前 table 的 direct `tr` 数；不截断、不补齐、不把单项数组隐式广播。
- 当 normalized target 全部为正数时，exact sum 必须是 safe integer；overflow 在任何 package write 前以 `RangeError` 拒绝。
- 当 target 包含 zero 时不推导总高，也不要求正值子集的和可表示，因为 transform height 不由该向量求出。

创建 API 的语义不改变：`addTable({ rowHeights })` 继续只接受显式正行高；创建时完全省略 `rowHeights` 才生成自动 zero rows。允许显式 zero 是 existing-table editing 对 OOXML 自动行高的必要控制，不反向扩大创建输入。

## 原子无损编辑

`setRowHeights()` 在 `OpcPackage.transaction()` 内重新 resolve live shape，验证 rows 和 transform，再修改同一个 slide XML document。setter 始终要求以下 transform 路径唯一，即使目标包含自动行高且不会改变 `cy`：

```text
graphicFrame -> xfrm -> ext@cy
```

`xfrm` 和 `ext` 必须是 direct child；`cy` 必须恰好有一个 unqualified attribute，且为 non-negative safe-integer decimal token。缺失、重复、namespaced-only、negative、小数、指数或 unsafe `cy` 都视为不可安全编辑，抛 `ModelParseError`，不在 opaque malformed transform 上做部分 row mutation。

通过全部验证后：

1. 每个 direct `tr@h` 只在当前 numeric value 与 normalized target 不同时替换 attribute value。
2. 若 target 全部为正数，`ext@cy` 只在当前 numeric value 与 exact row-height sum 不同时替换；这也会修复 valid rows/transform mismatch。
3. 若 target 包含 zero，`ext@cy` 完全保留，包括其原始 token 和 leading zeros。
4. 不替换整个 `tbl`、`tr`、`xfrm` 或 `ext` element；未知 attributes、children、prefixes、whitespace、comments、row/cell content 和 merge XML 保持原样。
5. 只要有一个 row token 或 eligible `cy` token 改变，序列化后的 slide part 才在 transaction 内写回；任何异常由 transaction 恢复 bytes 和 mutation journal。

若 target 全部为正数且 rows 已匹配但 `cy` 不等于 sum，调用会只修复 `cy`。若 target 包含 zero，则任何 rows/transform 数值 mismatch 都不是可推导的不一致，因此只按目标修改 rows，不推测 `cy`。

若所有 target row heights 与现有值在 numeric 上相等，且全正目标的 `cy` 也等于 sum，则方法是 semantic no-op：不调用 slide write，原始 tokens、完整 slide bytes 和 mutation journal 都保持不变。含 zero 的 numeric row no-op 同样保留 `cy`，不因其与数值和不同而写入。

## 错误模型

- 非法 scalar/array 结构或 item type：`TypeError`。
- negative dimension、unsafe rounding 或需要同步时的 sum overflow：`RangeError`。
- array 与现有 physical row 数不等：`TypeError`。
- 无法唯一、安全解析现有 rows 或 transform：带 slide `partUri` 的 `ModelParseError`。
- shape 在 rollback、删除或替换后已无法 resolve：沿用现有 live model 的 `ModelParseError`。

所有失败都必须保持 slide bytes、其他 package parts、shape identity 和 mutation journal 不变。getter 对 row XML 歧义返回 `undefined`；setter 对同一歧义抛错，保持“宽容读取、严格写入”。

## PptxGenJS 对等与差异

PptxGenJS 4.0.1 能在创建时通过 `rowH` 写 `tr@h`，但没有读取和编辑已有 PPTX 的对象模型。本 API 覆盖其 absolute row-height 能力，并扩展到已有 deck 的 lossless editing 和自动行高控制。

编辑输入保持本库已经确定的修复语义：scalar 不 floor，数组必须 exact length，不复制 PptxGenJS 的单项数组广播、short/falsy item fallback、long array truncation、字符串 coercion 或 transform/rows mismatch。文档将这些差异标记为 intentional compatibility repairs。

## 测试与验收

实现必须覆盖：

1. focused internal tests：valid unequal/zero rows、detached getter snapshot、positive/zero scalar broadcast、exact mixed array、rounding、all-positive safe sum、semantic no-op、positive mismatch repair、zero-path `cy` preservation 和 unknown XML preservation。
2. malformed matrix：missing/repeated/nested-only `tbl`，zero direct rows，missing/repeated/namespaced-only `h`，negative/decimal/exponent/signed/unsafe `h`，missing/repeated/nested-only `xfrm` 或 `ext`，以及 invalid/repeated/namespaced-only `cy`。
3. descriptor safety：hole、accessor、extra own key、symbol、typed array、wrong length、single-item array 和 all-positive overflow；accessor invocation count 必须为 zero。
4. public model integration：live identity、created auto/explicit table immediate edit、imported table edit、outer transaction rollback、no-op bytes/journal isolation 和 `setTransform({ height })` 行为不变。
5. merge coverage：含 `rowSpan`、`hMerge`、`vMerge` 或其他 merge XML 的表格仍按 direct `tr` 数读取和编辑，不修改 cells 或 merge markup。
6. SDK 与 packed Node/browser/declaration/CLI smoke：读取、scalar/array/zero 编辑、write/reopen 后保持 rows；全正目标同步 transform，含 zero 目标保留 transform。
7. PptxGenJS 4.0.1 fixture：读取其真实 row tokens，执行正值和自动行高编辑后重开验证。
8. 全仓 TypeScript project references、focused/full tests 和现有性能门禁。
9. PowerPoint 2010 package validation 对 native edit、reopen 和 PptxGenJS-derived edit 均为 zero error/zero warning。
10. package diff 证明仅目标 slide part 改变；native write/reopen 后 row heights、merge XML 和非目标 parts 稳定。
11. LibreOffice/Poppler render 目检 unequal rows、mixed auto/explicit rows、纵排、底部对齐和空 cell 正确，且 presentations overflow checker 无非预期 overflow。

## 文档与发布表面

更新 `CHANGELOG.md`、API README、PptxGenJS compatibility baseline、package README 和 packed smoke 输出。文档明确：

- `rowHeights` 是 exact EMU snapshot，zero 表示自动行高；使用 `inches()` 可提供正的英寸值。
- `setRowHeights()` 只在全部目标为正时同步总高；包含 zero 时保留 transform height。
- 不要用 `setTransform({ height })` 代替 row token 编辑。
- malformed row XML 的 getter 返回 `undefined`，setter 拒绝 mutation。
- 创建 API 仍不接受显式 zero `rowHeights`。

## 非目标

本小项不实现比例 resize、单行增删、行重排、merge/unmerge、cell count 修复、布局继承、主题样式、auto-fit、内容测量、分页、重复表头或重排文本。它不 canonicalize 整个 table XML，不修复 rows/transform 以外的 malformed OOXML，不改变 `addTable()` 或 `setTransform()` 的既有行为。
