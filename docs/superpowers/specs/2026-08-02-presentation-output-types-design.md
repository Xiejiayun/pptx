# Presentation Output Type Runtime Catalog Design

日期：2026-08-02

状态：已确认（按持续任务约定自主执行）

## 1. 目标与范围

本小项补齐 PptxGenJS 4.0.1 `OutputType` 对应的 runtime discovery 能力，同时保持 native 的输出层边界和 frozen catalog 设计：

```ts
export const OUTPUT_TYPES = Object.freeze([
  'arraybuffer',
  'base64',
  'binarystring',
  'blob',
  'nodebuffer',
  'uint8array',
] as const);

export type OutputType = typeof OUTPUT_TYPES[number];
```

- Node 与 browser 可在运行时枚举六种标准 presentation output type；
- `OutputType` 从唯一 tuple 派生，不维护重复手写 union；
- SDK 直接拥有 catalog，aggregate root 复用同一导出；
- catalog 顺序稳定、无重复且 frozen；
- PptxGenJS 4.0.1 public instance 的六个 `OutputType` keys/values 与顺序有永久对照；
- 不改变当前 `write(): Promise<Uint8Array>`、`writeBlob()`、`writeFile()` 或 `download()` 行为。

本项不实现六种返回值转换、Node readable stream、compression policy 或新的 write overload。它只完成后续输出语义所依赖的 runtime catalog 与类型源。

## 2. 当前状态与权威行为

Native 当前输出 API 位于 `@pptx/sdk`：

- `write(options?)` 验证文稿后返回 `Uint8Array`；
- `writeBlob(options?)` 包装相同 bytes 为当前格式 content type 的 `Blob`；
- `writeFile(path, options?)` 在 Node 写入路径并返回 `void`；
- `download(fileName?, options?)` 在 browser 触发下载并返回 `void`；
- `WriteOptions` 只包含 compatibility profile 与 strict/permissive mode。

因此 native 已支持 `uint8array` 和独立的 blob/path/download convenience surface，但尚无统一 `outputType` 选择器。缺口分为两个独立层次：先公开稳定六值 catalog，再让 write API 消费该 type 并实现精确返回语义。本设计只处理第一层。

锁定依赖 PptxGenJS 4.0.1 的 public declaration 公开 instance getter `readonly OutputType: typeof PptxGenJS.OutputType`。Runtime keys/values 顺序均为：

1. `arraybuffer`
2. `base64`
3. `binarystring`
4. `blob`
5. `nodebuffer`
6. `uint8array`

真实 Node public `write({ outputType })` probe 分别返回 `ArrayBuffer`、base64 string、binary string、`Blob`、`Buffer` 与 `Uint8Array`。`stream()` 使用额外的 internal/public write token `STREAM`，但 `STREAM` 不属于 instance `OutputType` enum 的六值，因此不进入本 catalog。

PptxGenJS 两个实例返回同一个可扩展、未冻结 enum-shaped object；prototype getter 无 setter。Native 对等 values 和顺序，不复制 shared mutable alias。

## 3. 方案比较

### A. 增加 `PptxDocument.OutputType` getter 与 enum-shaped object

调用形状最接近 PptxGenJS，但 output types 与任何 document package state 无关。实例 getter 会扩大每个文稿对象的非文稿状态表面，并重复项目已用 root catalog 取代 PptxGenJS mutable helper 的决定。拒绝。

### B. 只公开当前已能直接产生的 `uint8array` 与 `blob`

短期看更贴近现状，但 catalog 将不是 PptxGenJS `OutputType` 对等表面；后续增加四值会改变 tuple、type 与 enumeration contract，并迫使调用方分两次迁移。拒绝。

### C. 完整 frozen 六值 tuple + derived type（采用）

在 `packages/sdk/src/output-type.ts` 定义 `OUTPUT_TYPES` 与 `OutputType`。SDK `index.ts` 显式导出，aggregate root 通过既有 `export * from '@pptx/sdk'` 复用同一对象。后续六种 write 返回值直接消费这一 type，不需要重命名或新增第二套 catalog。

`OUTPUT_TYPES` 与 `CHART_TYPES`、`TEXT_ALIGNMENTS`、`TEXT_VERTICAL_ALIGNMENTS`、`PLACEHOLDER_TYPES` 和 `PRESET_SHAPE_TYPES` 的公开发现模式一致。`OutputType` 延续 PptxGenJS public concept 和项目 `ChartType`/`PlaceholderType` 命名，不增加 `PptxGenJSOutputType` 或 generic alias。

## 4. 公共 API 与兼容性

```ts
import { OUTPUT_TYPES, type OutputType } from '@jiayunxie/pptx';

for (const outputType of OUTPUT_TYPES) {
  const value: OutputType = outputType;
  console.log(value);
}
```

精确 contract：

- value 为 readonly tuple `['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']`；
- `Object.isFrozen(OUTPUT_TYPES) === true`；
- `OutputType` 精确等于 `(typeof OUTPUT_TYPES)[number]`；
- SDK 与 aggregate root runtime import 指向同一个对象；
- 不增加 generic `OutputType` value、instance getter、enum-shaped object 或 mutable alias；
- 本项不在 `WriteOptions` 增加 `outputType`，避免 catalog 落地时顺带引入未实现返回语义；
- 当前无参数或 compatibility/mode-only `write()` typing 和 runtime 完全不变。

新增 export 是 additive change。仓库中当前没有其他 `OutputType` public type，因此不存在重名迁移。

## 5. Runtime、错误与数据语义

- catalog 在模块初始化时创建一次并冻结；所有 SDK/root imports 共享 immutable reference；
- runtime 对 tuple 的 `push`、index assignment 或 extension 在 strict mode 失败；
- 读取与迭代 catalog 不访问 document、OPC package、JSZip、Node module 或 browser global；
- catalog 可安全进入 Node 与 browser bundle，不引入 `Buffer`、`Blob` 或 conditional-module dependency；
- catalog 自身不改变 diagnostics、mutation journal、write bytes、ZIP compression、MIME 或 file name；
- `STREAM` 被明确排除，因为它不是 PptxGenJS instance `OutputType` 的 public enum member，后续由独立 stream API 处理；
- 无新 runtime validation 或 error message。

## 6. 后续输出阶段边界

下一小项让 `WriteOptions` 和 `write()` 消费 `OutputType` 时必须另行设计并验证：

1. 六种输入 token 的环境可用性与精确返回类型；
2. 默认仍返回 native `Uint8Array`，还是迁移到 PptxGenJS browser-default blob；
3. `nodebuffer` 在 browser 的拒绝时机和错误；
4. base64 与 binary string 的无损编码；
5. generic/overload typing 对现有 `write(options)` call sites 的 source compatibility；
6. `writeBlob()`、`writeFile()`、`download()` 与统一 selector 的职责；
7. compression policy、Node readable stream 与返回值语义。

这些决定不会被 catalog 实现偷偷锁定。本项唯一承诺是 token 集合、顺序与 derived type。

## 7. 验证策略

### Source 与 public API

- exact tuple、顺序、唯一性、frozen 与 mutation rejection；
- `OutputType` 接受全部六项，拒绝 `STREAM`、未知值和大小写变体；
- SDK/root runtime identity 与 type export；
- import/iteration 不产生 document package mutation；
- PptxGenJS public `OutputType` keys/values 与 native tuple 精确同序；
- public probe 确认六值实际返回种类，作为下一阶段基线，但不改变 native write。

### Packed runtime

- actual tarball declarations 含 readonly tuple 与 derived type；
- installed Node 与 browser conditional export 检查 exact values、frozen 与 shared identity；
- installed TypeScript consumer 迭代 tuple，并拒绝 push、index assignment、`STREAM` 和 unknown token；
- real Google Chrome 检查 catalog，不要求调用尚未实现的统一 output selector；
- installed CLI package inspection 与现有 create/write/reopen 继续通过。

## 8. 完成门禁

1. SDK、aggregate root、actual tarball Node/browser/types 均公开同一 frozen catalog；
2. PptxGenJS public `OutputType` 六值和顺序有永久对照测试；
3. 当前 write/writeBlob/writeFile/download behavior 与 declarations 无回归；
4. focused/full Vitest、performance、两套 TypeScript、bundles、declarations、pack、CLI 与 real Chrome 全绿；
5. 发布文档把 `OutputType` runtime catalog 从缺口移入支持项，但继续明确六种 write semantics、stream 与 compression 尚未完成；
6. 每个实施任务独立 review、commit、push，并确认远端同步。
