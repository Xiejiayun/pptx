# Existing Table Column Width Editing Design

## 目标与范围

本小项为 `TableModel` 增加已有表格列宽的严格读取和无损编辑，使调用方既能读取 PowerPoint `a:tblGrid/a:gridCol@w` 的 exact EMU 值，也能在一次原子操作中更新所有列宽并同步表格 transform 宽度。

这是从“可创建表格”推进到“可精确编辑已有表格”的第一个 geometry 小项。它只处理列宽，不同时加入行高编辑、按比例缩放、合并单元格编辑、自动布局、内容测量或分页。后续行高编辑沿用本设计的严格解析和同步 mutation 模式，作为独立小项 review、commit 和 push。

## 公共 API

`TableModel` 新增：

```ts
get columnWidths(): readonly number[] | undefined;
setColumnWidths(value: number | readonly number[]): void;
```

所有值均为 EMU。getter 每次返回 detached snapshot；调用方修改返回数组不得影响模型。setter 使用方法而不是属性 setter，因为 scalar 需要按当前 physical grid column 数广播，而数组需要做严格结构验证。

`undefined` 只表示 getter 无法从现有 XML 得到唯一、完整、可信的列宽向量；setter 不接受 `undefined`，因为合法 `tblGrid` 的每一列都必须有正宽度。

## 方案选择

考虑三个方案：

1. 增加独立、同步的 `columnWidths` getter 与 `setColumnWidths()`。setter 同时更新 grid 和 transform，保持一个明确 invariant。采用此方案。
2. 只更新 `tblGrid`，保留 `p:xfrm/a:ext@cx`。这会主动生成 PowerPoint geometry 不一致的文件，不采用。
3. 改变继承的 `setTransform({ width })`，自动按比例缩放 grid。它会引入比例语义、舍入余数分配和零宽/overflow 边界，也会悄然改变现有 API 行为。本小项不采用；将来若需要，应作为独立 resize API 设计。

采用方案 1 后，调用方仍可使用现有 `setTransform()` 修改 x/y/height 等字段；要改变表格总宽和各列，应使用 `setColumnWidths()`。本小项不破坏或重新解释现有 `setTransform({ width })`。

## 严格读取

读取从当前 live `TableModel` 对应的 `p:graphicFrame` 开始，只沿 direct-child 路径解析：

```text
graphicFrame -> graphic -> graphicData -> tbl -> tblGrid -> gridCol
```

每一级 `graphic`、`graphicData`、`tbl` 和 `tblGrid` 都必须唯一。节点按 direct `localName` 匹配，以沿用当前 lossless XML 模型对合法 alternate prefixes 的兼容；descendant 中的同名 fake 节点不参与解析。

`tblGrid` 必须至少包含一个 direct `gridCol`。其他未知 direct child、注释、namespace declaration、属性和 extension data 不影响读取，并在编辑时原样保留。合并单元格的 `gridSpan`、`hMerge` 或 `vMerge` 不改变 physical grid column 数；`tblGrid` 是唯一权威来源，不从行或 cell 数猜测列数。

每个 direct `gridCol` 必须恰好包含一个 unqualified `w` attribute。token 必须只包含 ASCII 十进制数字，转换后为大于 zero 的 safe integer。缺失、重复、namespaced-only、空值、小数、指数、带正负号、zero 或超过 safe integer 的宽度都使整个 getter 返回 `undefined`，不返回部分向量，也不 fallback 到 transform width 或平均列宽。

读取不检查 transform，因此 valid grid 即使与 `a:ext@cx` 不一致，getter 仍返回 grid snapshot。getter 不 mutation package bytes 或 journal。

## 输入归一化

setter 首先把输入复制为 detached normalized representation，再解析和修改 package：

- scalar 必须是 finite number，`Math.round()` 后为大于 zero 的 safe integer；随后按当前 valid grid column 数精确广播。
- array 必须通过 `Array.isArray()`、非空且 dense；只允许 `length` 和从 `0` 到 `length - 1` 的 own data properties。hole、accessor、symbol key、额外字符串 key、typed array 或继承取值都拒绝，且不得调用 getter。
- array 每项使用与 scalar 相同的 finite、round、positive safe-integer 规则。
- array 长度必须严格等于当前 `tblGrid` 的 direct `gridCol` 数；不截断、不补齐、不把单项数组隐式广播。
- normalized widths 的 exact sum 必须是 safe integer；overflow 在任何 package write 前以 `RangeError` 拒绝。

这与创建时 `columnWidths` 的公共语义一致，但 editing helper 保持独立的 XML 解析/mutation 职责。实现可抽取或复用纯 dimension normalization helper，前提是创建路径的输出保持字节和错误语义不变。

## 原子无损编辑

`setColumnWidths()` 在 `OpcPackage.transaction()` 内重新 resolve live shape，验证 grid 和 transform，再修改同一个 slide XML document。setter 要求以下 transform 路径唯一：

```text
graphicFrame -> xfrm -> ext@cx
```

`xfrm` 和 `ext` 必须是 direct child；`cx` 必须恰好有一个 unqualified attribute，且为 non-negative safe-integer decimal token。zero 被接受，因为 PptxGenJS 和其他生成器可能产生 `cx="0"` 而 grid 为正；setter 会把它修复为新 grid sum。缺失、重复、namespaced-only、negative、小数、指数或 unsafe `cx` 都视为不可安全编辑，抛 `ModelParseError`，不猜测或重建 opaque transform。

通过验证后：

1. 每个 direct `gridCol@w` 只在当前 numeric value 与 normalized target 不同时替换 attribute value。
2. `ext@cx` 只在当前 numeric value 与 normalized width sum 不同时替换。
3. 不替换整个 `tblGrid`、`gridCol`、`xfrm` 或 `ext` element；未知属性、children、前缀、空白、注释和顺序保持原样。
4. 只要有一个 token 改变，序列化后的 slide part 在 transaction 内写回；任何异常由 transaction 恢复 bytes 和 mutation journal。

setter 接受 valid 但 grid/transform mismatch 的文件，并把 `cx` 规范化为新 grid sum。若 normalized grid 已与现有 widths 相等但 `cx` 不等于 sum，调用仍会修复 `cx`。

若 widths 与 `cx` 在 numeric 上都已完全匹配，则方法是 semantic no-op：不调用 slide write，原始 token（包括 leading zero）、完整 slide bytes 和 mutation journal 都保持不变。

## 错误模型

- 非法 scalar/array 结构或 item type：`TypeError`。
- 非正 dimension、unsafe rounding 或 sum overflow：`RangeError`。
- array 与现有 physical grid column 数不等：`TypeError`。
- 无法唯一、安全解析现有 grid 或 transform：带 slide `partUri` 的 `ModelParseError`。
- shape 在 rollback、删除或替换后已无法 resolve：沿用现有 live model 的 `ModelParseError`。

所有失败都必须保持 slide bytes、其他 package parts、shape identity 和 mutation journal 不变。getter 对 grid 内容歧义返回 `undefined`；setter 对同一歧义抛错，这是现有“宽容读取、严格写入”模式。

## PptxGenJS 对等与差异

PptxGenJS 4.0.1 能在创建时通过 `colW` 写 grid，但没有读取和编辑已有 PPTX 的对象模型。本 API 覆盖其 exact absolute column-width 能力，并扩展到已有 deck 的 lossless editing。

编辑输入保持本库已经确定的修复语义：scalar 不 floor，数组必须 exact length，不复制 PptxGenJS 的单项数组广播、长度不匹配 fallback、字符串 coercion 或 transform/grid mismatch。文档会把这些差异标记为 intentional compatibility repairs，而不是遗漏。

## 测试与验收

实现必须覆盖：

1. focused internal tests：valid/unequal grid、detached getter snapshot、scalar broadcast、exact array、rounding、safe sum、semantic no-op、transform mismatch repair、unknown XML preservation。
2. malformed matrix：missing/repeated/nested-only `tblGrid`，zero/negative/decimal/exponent/unsafe/repeated/namespaced-only `w`，missing/repeated/nested-only `xfrm` 或 `ext`，以及 invalid/repeated/namespaced-only `cx`。
3. descriptor safety：hole、accessor、extra own key、symbol、typed array、wrong length、single-item array 和 overflow；accessor invocation count 必须为 zero。
4. public model integration：live identity、created table immediate edit、imported table edit、outer transaction rollback、no-op bytes/journal isolation。
5. merge coverage：含 `gridSpan`、`hMerge`、`vMerge` 的表格仍按 `tblGrid` 列数读取和编辑，不修改 merge XML。
6. SDK 与 packed Node/browser/declaration/CLI smoke：读取、scalar/array 编辑、write/reopen 后保持 widths 与 transform sum。
7. PptxGenJS 4.0.1 fixture：读取其真实 table grid，修复可能的 `cx=0`/mismatch，再重开验证。
8. 全仓 TypeScript project references、focused/full tests 和现有性能门禁。
9. PowerPoint 2010 package validation 对 native edit、reopen 和 PptxGenJS-derived edit 均为 zero error/zero warning。
10. package diff 证明仅目标 slide part 改变；native write/reopen 后列宽、merge XML 和非目标 parts 稳定。
11. LibreOffice/Poppler render 目检 unequal columns 正确，且 presentations overflow checker 无非预期 overflow。

## 文档与发布表面

更新 `CHANGELOG.md`、API README、PptxGenJS compatibility baseline 和 package smoke 输出。文档明确：

- `columnWidths` 是 exact EMU snapshot；使用 `inches()` 可提供英寸。
- `setColumnWidths()` 同步总宽；不要用 `setTransform({ width })` 代替列宽编辑。
- malformed grid 的 getter 返回 `undefined`，setter 拒绝 mutation。
- row-height reading/editing 仍是下一独立小项，不在本提交伪装为已支持。

## 非目标

本小项不实现 row-height getter/setter、比例 resize、单列增删、列重排、cell merge/unmerge、cell count 修复、布局继承、主题样式、auto-fit、内容测量、分页或重排文本。它也不 canonicalize 整个 table XML，不修复 grid 以外的 malformed OOXML，不改变 `setTransform()` 的既有行为。
