# Presentation Vertical Alignment Runtime Catalog Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `AlignV` 对应的 runtime discovery 能力，同时保持 native 既有 API 命名和 catalog 设计：

```ts
export const TEXT_VERTICAL_ALIGNMENTS = Object.freeze([
  'top',
  'middle',
  'bottom',
] as const);

export type TextBoxVerticalAlignment =
  typeof TEXT_VERTICAL_ALIGNMENTS[number];
```

- Node 与 browser 可在运行时枚举全部垂直文字对齐值；
- `TextBoxVerticalAlignment` 从唯一 tuple 派生，不再维护独立手写 union；
- SDK/root 通过现有 model export chain 公开 constant 与 type；
- catalog 顺序稳定、无重复且 frozen；
- 现有 text box、slide number、table/table-cell 创建和编辑继续接受相同三值；
- 不改变任何 OOXML bytes、normalization、fallback、write 或 reopen 语义。

本项不实现 output types、scheme colors 或其他 runtime helpers，也不增加 document-instance property。它只完成垂直对齐 runtime catalog。

## 2. 当前状态与 PptxGenJS 权威行为

Native 已有 `TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom'`，并在以下公开入口复用：

- `AddTextOptions.valign`、text shape `verticalAlignment` 与 slide-number style；
- table/table-cell `valign` 创建；
- `TableCell.verticalAlignment` 与 `setCellVerticalAlignment()`；
- slide/layout/master/placeholder/declarative text owners。

`normalizeTextBoxVerticalAlignment()` 与现有 codecs 已严格覆盖 text-body `top → t`、`middle → ctr`、`bottom → b`，table-cell codec 复用同样的 direct token。非法 runtime input 在 mutation 前拒绝。缺口只是调用方无法从 runtime export 发现三种值。

PptxGenJS 4.0.1 declaration 公开 `readonly AlignV: typeof PptxGenJS.AlignV`，enum keys/values 依次为 `top`、`middle`、`bottom`。Public runtime probe 显示 prototype getter 没有 setter，两个实例返回同一个可扩展、未冻结 object。Native 对等 runtime values 和顺序，但不复制 mutable alias。

## 3. 方案比较

### A. 增加 `PptxDocument.AlignV` getter 和 enum-shaped object

迁移表面最接近 PptxGenJS，但常量与文稿 package 无关；每个实例暴露同一 mutable helper 会扩大状态表面，并偏离项目现有 root catalog 模式。拒绝。

### B. 只保留 `TextBoxVerticalAlignment` 类型

零代码变化，但 JavaScript 调用方仍不能枚举值，也没有 PptxGenJS `AlignV` 的 runtime 能力。拒绝。

### C. Frozen tuple + derived type（采用）

在 `packages/model/src/text.ts` 直接新增 `TEXT_VERTICAL_ALIGNMENTS`，并从 tuple 派生现有 `TextBoxVerticalAlignment`。SDK 与 aggregate root 已 `export *`，无需 facade 或重复状态。该模式与 `TEXT_ALIGNMENTS`、`CHART_TYPES`、`PLACEHOLDER_TYPES`、`PRESET_SHAPE_TYPES` 一致。

`TEXT_VERTICAL_ALIGNMENTS` 比 `VERTICAL_ALIGNMENTS` 更明确，避免与对象、表格布局或几何方向混淆；比 `TEXT_BOX_VERTICAL_ALIGNMENTS` 更贴近该类型已跨 text box、slide number 和 table cell 复用的实际范围。

## 4. 公共 API 与兼容性

```ts
import {
  TEXT_VERTICAL_ALIGNMENTS,
  type TextBoxVerticalAlignment,
} from '@jiayunxie/pptx';

for (const alignment of TEXT_VERTICAL_ALIGNMENTS) {
  const value: TextBoxVerticalAlignment = alignment;
  console.log(value);
}
```

精确 contract：

- value 为 readonly tuple `['top', 'middle', 'bottom']`；
- `Object.isFrozen(TEXT_VERTICAL_ALIGNMENTS) === true`；
- `TextBoxVerticalAlignment` 仍是完全相同的三值 union；
- 不增加 generic `AlignV` export，避免复制 PptxGenJS namespace/instance shape；
- 不改变现有 source compatibility 或 emitted declarations 中 `TextBoxVerticalAlignment` 可赋值范围。

## 5. Runtime、错误与 OOXML 语义

- catalog 在模块初始化时创建一次并冻结；所有 import 共享同一 immutable reference；
- runtime 对 tuple 的 `push`、index assignment 或 extension 在 strict mode 失败；
- catalog 本身不参与 document mutation、OPC graph、diagnostics 或 output；
- text-body 与 table-cell 的 `Record<TextBoxVerticalAlignment, string>` 继续让 TypeScript 检查三值 mapping 完整性；
- 无效 alignment 的既有 `TypeError`、rollback 与 mutation-isolation 行为不变；
- text box、slide number 和 table/table-cell create/edit/write/reopen final state 不变。

## 6. PptxGenJS 对等边界

对等范围是三个 runtime token、稳定枚举顺序、公开可导入/读取，以及 token 可直接用于对应文字和表格 API。Native 有意不复制：

1. document instance 上的 `AlignV` getter；
2. enum-shaped key/value object；
3. 可修改同一个内部对象的 alias 行为。

Native tuple 更适合 tree-shaking、JavaScript iteration 和 derived literal type；compatibility 文档继续说明不要求 PptxGenJS 调用只替换 import 后原样运行。

## 7. 验证策略

### Source 与 public API

- exact tuple、顺序、唯一性、frozen 与 mutation rejection；
- `TextBoxVerticalAlignment` 对三项接受，对未知值拒绝；
- SDK/root runtime 与 type export identity；
- 每个 catalog token 通过 text box 和 table-cell existing creation path 并 write/reopen；
- PptxGenJS public `AlignV` keys/values 与 native tuple 集合/顺序一致；
- catalog 读取不产生 package mutation，既有 slide-number 与 vertical-alignment suites 不回归。

### Packed runtime

- actual tarball declarations 含 readonly tuple 与 derived type；
- installed Node 和 browser conditional export 检查 exact values、frozen、shared identity；
- installed TypeScript consumer 迭代 tuple，拒绝 push、index assignment 与 unknown token；
- real Google Chrome 检查 runtime catalog，并用三值创建/reopen text shapes 与 table cells；
- CLI package inspection 继续验证生成 PPTX 可读。

## 8. 完成门禁

1. model/SDK/root/actual tarball Node/browser/types 均公开同一 frozen catalog；
2. PptxGenJS public AlignV 三值有永久对照测试；
3. 现有 vertical alignment create/edit/write/reopen 行为不回归；
4. focused/full Vitest、performance、两套 TypeScript、bundles、declarations、pack、CLI 与 real Chrome 全绿；
5. 六份 release 文档把 `AlignV` runtime helper 从缺口移入支持项，保留其他 helpers；
6. 每个实施任务独立 review、commit、push，并确认远端同步。
