# Text Box Fit Design

## 目标与范围

为普通文本 shape 增加文本框自动适配的原生创建、读取、编辑和清除能力，覆盖 PptxGenJS 4.0.1 `TextPropsOptions.fit` 的三个公开值：不自动适配、溢出时缩小文字、按文字调整 shape。该能力只管理 direct `a:bodyPr` 下的 text-autofit choice child，不计算字体缩放比例，也不主动改变 shape transform。

本小项只包含普通文本 shape 的 direct fit 状态。不包含 table-cell fit、placeholder/master inheritance 编辑、overflow、word breaking、columns、paragraph/run font-size 重算或 PowerPoint 的动态 layout engine。PptxGenJS 文档也明确说明 `shrink` / `resize` 的最终缩放由 PowerPoint 在编辑文字或改变 shape 大小时动态计算，本库只生成和保留相同 OOXML 意图。

## 公共 API

公共值与 PptxGenJS 4.0.1 的现代 `fit` union 一致：

```ts
export type TextBoxFit = 'none' | 'shrink' | 'resize';

interface AddTextOptions {
  readonly fit?: TextBoxFit;
}
```

live shape 使用完整语义名：

```ts
class ShapeModel {
  get textFit(): TextBoxFit | undefined;
  set textFit(value: TextBoxFit | undefined);
}
```

用法：

```ts
const shape = slide.addText('Long text', {
  fit: 'shrink',
});

shape.textFit = 'resize';
shape.textFit = 'none';      // PptxGenJS-compatible no-autofit wire form
shape.textFit = undefined;   // remove any direct supported fit child
```

`addText()` 和 `addRichText()` 共用相同 text-body 创建路径。`shape.textFit` 只读取 direct fit child，不解析继承，也不通过 text length 或 shape size 猜测适配状态。

## 方案选择

考虑过三个方案：

1. `none` 写 `<a:noAutofit/>`。语义显式，但不匹配 PptxGenJS 4.0.1；其源码还明确记录该 child 会导致 PowerPoint 2013 问题。
2. 只公开 `shrink` / `resize`。实现最小，但遗漏 PptxGenJS 的公开 `none`，不能形成完整现代 `fit` 表面对等。
3. 采用 PptxGenJS wire semantics：省略和 `none` 都不写 fit child，`shrink` 写 `<a:normAutofit/>`，`resize` 写 `<a:spAutoFit/>`；同时读取既有 `<a:noAutofit/>` 为 `none`。采用此方案。

创建参数沿用 PptxGenJS 的 `fit` 名称，live property 使用 `textFit`，与现有 `valign` / `verticalAlignment`、`wrap` / `textWrap`、`vert` / `textDirection` 模式一致。实现层新增窄 `text-box-fit.internal.ts` codec；不把 child choice 强行塞进 attribute helper，也不提前抽象 RTL、columns 或 overflow。

用户已授权实现方持续选择最佳方案并直接推进，因此本设计按 PptxGenJS 兼容、PowerPoint 2013 安全、direct-state 可解释性和最小职责边界定稿。

## PptxGenJS 4.0.1 对照

使用公开 `addText()` / `writeFile()` 生成真实文件并检查 slide XML：

- 省略 `fit` 与显式 `fit: 'none'` 都生成不含 fit child 的空 `bodyPr`；PptxGenJS raw XML 使用 expanded empty form，本库保留现有 self-closing creation form，两者都不写 `noAutofit`。
- `fit: 'shrink'` 生成 direct `<a:normAutofit/>`。
- `fit: 'resize'` 生成 direct `<a:spAutoFit/>`。
- truthy 非法字符串不会生成 fit child；本库不复制这种静默回退，严格拒绝非法输入。
- 单个 `TextProps[]` run 上的 `fit`、`shrinkText` 和 `autoFit` 都被忽略。
- deprecated outer `shrinkText: true` 仍生成 `<a:normAutofit/>`，`autoFit: true` 仍生成 `<a:spAutoFit/>`。本库 adapter 必须正确读取这些公开输出，但原生创建 API 不新增 deprecated alias，因为现代 `fit` 已覆盖相同功能。

adapter 仍只导入 PptxGenJS 写出的 OOXML，不读取 `_slides` 或其他私有字段。冲突的现代/deprecated 参数可能产生多个 fit children；这种非规范输入保留原始 XML，getter 不猜测单一状态。

## OOXML 映射与 direct 语义

三个语义值的 wire 映射：

```xml
<!-- omitted / none -->
<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>

<!-- shrink -->
<a:bodyPr wrap="square" rtlCol="0" anchor="ctr">
  <a:normAutofit/>
</a:bodyPr>

<!-- resize -->
<a:bodyPr wrap="square" rtlCol="0" anchor="ctr">
  <a:spAutoFit/>
</a:bodyPr>
```

getter 对 direct children 使用严格 choice 规则：恰好一个 `noAutofit` 返回 `none`，恰好一个 `normAutofit` 返回 `shrink`，恰好一个 `spAutoFit` 返回 `resize`。没有 direct supported child 返回 `undefined`；多个 supported children 返回 `undefined`，因为该结构违反 choice 约束，不能安全选一个。descendant、case variant、相似未知 child 都不生效，只读访问不增加 mutation。

显式 `fit: 'none'` 创建后 getter 返回 `undefined`，这是有意保留的 wire-level 行为：PptxGenJS 将 none 表示为 direct child 缺席，本库不伪造继承或默认快照。既有显式 `<a:noAutofit/>` 仍可读取为 `none`，且所有不相关编辑必须原样保留它。

## 创建与无损编辑

`validateAddTextOptions()` 在 package mutation 前严格验证 `fit`。normalized creation options 持有可选 `textFit`：

- omitted / `none` 不渲染 child，保持当前 self-closing `bodyPr` 输出。
- `shrink` / `resize` 把 `bodyPr` 展开，并在 attributes 后、paragraphs 前写唯一 canonical fit child。

live setter 定位 direct `txBody/bodyPr`，只管理 direct `noAutofit`、`normAutofit`、`spAutoFit` children：

- `undefined` 或 `none` 移除所有 direct supported fit children，不新增 `noAutofit`。
- `shrink` / `resize` 产生唯一对应 child。
- 设置 `shrink` / `resize` 且当前已经是唯一对应 child 时执行 no-op，保留 `normAutofit@fontScale`、`lnSpcReduction`、quote style、namespace 与扩展内容。
- 切换模式或规范化多个冲突 children 时，移除旧 supported choice，在第一个旧 choice 的位置写 canonical child；没有旧 choice 时插入到 `scene3d` / `sp3d` / `extLst` 之前，否则放在 `bodyPr` 末尾，以符合 CT_TextBodyProperties 顺序。

wrap、margins、anchor、vert、RTL/column attributes、`prstTxWarp`、3D、extension、namespace 与未知 children 保持原字节和相对顺序。self-closing/expanded `bodyPr` 都保持合法；清除后不要求把 expanded XML 压回 self-closing，以免无关格式重写。缺少 direct `txBody` 或 `bodyPr` 时抛出 `ModelParseError`，不创建推测结构。

`.text`、`.richText`、`textMargins`、`verticalAlignment`、`textWrap`、`textDirection`、transform 和其他非 fit mutation 保留原 fit XML。设置 fit 不重建 paragraphs/runs，不改变 shape identity、relationships 或其他 package state。

## 输入失败与事务

创建和 live setter 对 null、boolean、number、空字符串、case/whitespace variant、未知字符串、array、object 和 symbol 抛出 `TypeError`。所有 validation 在 package mutation 前完成。

失败不得改变 part bytes、mutation journal、live `ShapeModel` identity、text、margin、vertical-alignment、wrapping、direction 或 fit snapshot；外层 transaction rollback 同样完整恢复。读取 malformed direct text body 遵循现有 `ModelParseError` 行为。

## 测试与发布门禁

验收覆盖：

1. `addText()` / `addRichText()` 对 omitted、none、shrink、resize 生成正确 bodyPr 结构，并与 margin、valign、wrap、vert、paragraph/run style 组合。
2. getter 只接受唯一 direct `noAutofit` / `normAutofit` / `spAutoFit`；absent、descendant、case variant、unknown 和 multiple choice 返回 `undefined`，只读不产生 mutation。
3. `shape.textFit` 在 self-closing/expanded `bodyPr` 上切换、清除或规范化 fit choice，保留 calculated shrink attributes、其他 bodyPr metadata、paragraphs/runs 与 stable identity；write/reopen 和 duplicate 一致。
4. plain/rich text、margin、vertical alignment、wrapping、direction 与 transform 编辑保留 fit；缺失 `txBody/bodyPr`、非法输入及 rollback 不改变 bytes、journal 或 live snapshots。
5. PptxGenJS 4.0.1 conformance 覆盖 omitted、none、shrink、resize、invalid ignored、run ignored、legacy `shrinkText` / `autoFit` 的真实输出导入。
6. Node/browser/declaration tarball smoke 覆盖创建、读取、模式切换与 clear assignment。
7. 原生与 PptxGenJS 对照文件均通过 `powerpoint-2010` CLI profile 的 0 error / 0 warning；LibreOffice 无修复导出并视觉核对 none/shrink/resize 的一致表现，同时记录 PowerPoint 动态重算限制。

完成后更新 changelog、公共 API、npm README 与兼容矩阵，把普通文本框 fit 从剩余 partial row 中移除，并明确 table-cell fit 仍未覆盖。全仓测试、独立性能、真实 tarball、CLI 和 LibreOffice 对照全部通过后，才允许 implementation commit 与 push。
