# OPC Mutation Transaction 基础设计

日期：2026-07-28
状态：已批准实施

## 目标

为 package graph 增加同步、可嵌套的 mutation transaction，使一次高层编辑涉及的 part、content type、relationship、ZIP entry 和 mutation journal 要么全部提交，要么完整回滚。首批接入现有的 slide、master/layout/theme 和 media 组合写操作，消除异常路径留下孤立 part、悬空关系或半更新 XML 的风险。

本小项只建立事务与回滚边界，不同时实现稳定语义对象 identity、完整引用所有权/垃圾回收或统一 codec 验证；这些能力在后续独立小项中接入事务钩子。

## 方案选择

考虑过三种方案：

1. 每次事务把当前 PPTX 写成 ZIP bytes，失败后重新打开。实现直观，但事务前已有 mutation 时会重复压缩、破坏对象引用，并把同步编辑强制变成异步。
2. 为每个 mutation 记录反向命令。内存开销较小，但 part、content type、relationship 和 ZIP 状态之间的逆操作容易遗漏，初期正确性风险高。
3. 在 `OpcPackage` 内创建内存 savepoint，深拷贝可变 package 状态；提交时丢弃 savepoint，失败时原位恢复。该方案保持现有同步对象模型、支持嵌套事务，并能用明确状态集合验证完整性，因此采用此方案。

## API 与语义

`OpcPackage` 新增同步事务：

```ts
const result = pkg.transaction(() => {
  // one or more package/model mutations
  return value;
});
```

- callback 正常返回即提交，并原样返回结果。
- callback 抛出时恢复事务开始时的全部 package 状态，并重新抛出原错误。
- callback 返回 Promise/thenable 时立即回滚并抛出 `TypeError`；异步资源读取必须在进入事务前完成。
- 显式嵌套事务各自形成 savepoint。内层失败可由外层捕获，且只回滚内层变更；外层失败回滚整个外层范围。
- package 的单次公开 mutation 在没有活动事务时自动使用事务保护；在活动事务中直接加入当前事务，避免为每个内部步骤重复复制状态。

`PptxDocument` 新增同样的同步 `transaction(callback)` 门面，callback 接收当前 document。成功返回前运行 package 结构验证；出现 error diagnostic 时抛出 `ValidationError` 并回滚。codec 与兼容性诊断的统一聚合留给后续验证小项。

TypeScript 签名不接受异步 callback。运行时 thenable 检查只是防止已经发生的同步 mutation 被提交，无法取消调用方自行启动的异步 continuation；因此异步 callback 不属于受支持用法，也不得在返回后继续修改同一 document。

## Savepoint 内容与恢复

Savepoint 必须包含：

- 每个 part 的 URI、content type、独立 bytes 和 relationship 数组；
- content type defaults 与 overrides；
- mutation journal 的完整顺序；
- 与上述 part 集合一致的 JSZip 文件 entry。

恢复在现有 `OpcPackage` 实例上完成，不能替换 package 对象。这样 `PresentationModel`、codec 和已取得的 model 对象仍引用同一个 package。事务开始前没有 mutation 时，回滚后 `write()` 继续返回原始输入 bytes；开始前已有 mutation 时，回滚保留那些既有变更和 journal。

## 首批接入范围

- `PresentationModel` 的 add、duplicate、move、delete slide。
- `MasterLayoutThemeCodec` 中会同时写多个 part/relationship/XML list 的 create、copy、delete、relink 操作。
- `MediaCodec` 的 add/delete。异步 source、transcode、poster 和 hash 解析先完成；实际 part、relationship 和 slide XML 写入放入一个同步事务。
- OPC 的 part 与 relationship 公开 mutation 自动具备单操作原子性。

只修改一个既有 part 的 shape/text/table/gradient 写操作继续依赖 OPC 单操作原子性，不增加重复的 model 层 savepoint。

## 错误处理

- 回滚不包装业务错误，调用方仍收到原始 `PackageError`、`ValidationError`、`RangeError` 或 codec 错误。
- 如果 callback 返回 thenable，错误明确说明事务仅支持同步 callback，并建议先完成异步准备工作。
- rollback 自身不得向 mutation journal 添加记录。
- 事务成功后不保留 savepoint 或暴露内部可变快照。

## 测试与验收

1. OPC 事务提交同时保留 part、content type、relationship 和 journal 变更。
2. 异常回滚恢复 part payload、content type、关系图、ZIP 输出和 journal；无 mutation 输入恢复字节 identity。
3. 嵌套事务的内层回滚只撤销内层，外层回滚撤销全部外层变更。
4. Promise/thenable callback 被拒绝且无残留 mutation。
5. malformed slide 上的 add/duplicate/media 等故障注入不会留下孤立 part 或 relationship。
6. `PptxDocument.transaction()` 在 package 结构验证失败时回滚。
7. 既有 slide、master/layout/theme、media 成功路径保持通过。
8. 全仓 typecheck、测试、Node/browser bundle 和 npm tarball smoke 通过。
