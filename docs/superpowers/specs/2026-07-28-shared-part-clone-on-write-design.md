# Shared Part Clone-on-Write 设计

日期：2026-07-28
状态：已批准实施

## 目标

为现有 `ImageModel.replaceData()` 和 `ChartModel.setXml()` 增加 clone-on-write。编辑独占 target 时继续原位修改；target 被多个 relationship 引用，或同一 slide relationship id 被多个 shape 引用时，只克隆被编辑 shape 的 target 并重定向该 shape，其他对象保持不变。

本小项覆盖当前公开的 image payload 与 raw chart XML 编辑。完整 chart data/workbook 同步 API 属于后续 chart 功能小项，但 chart clone 必须同时复制其 owned workbook/style/color 等依赖子图。

## 方案选择

考虑过三种方案：

1. 永远原位修改 target。最快，但 duplicate slide 后共享图片或外部导入的共享 chart 会发生跨对象串改。
2. 每次编辑都 clone。隔离简单，但独占 part 也会产生无用副本和 relationship churn。
3. 根据 package graph incoming 和 slide XML 内 rId 引用数决定：独占时原位写，共享时 clone 并最小重定向。该方案同时保证隔离和文件大小，因此采用此方案。

## 共享判定

一个 internal target 在以下任一条件成立时视为 shared：

- package graph 中有两个或更多 incoming relationships；
- 当前 slide XML 中，同一个 relationship id 被两个或更多 shape 引用。

只有 graph incoming 为一且 slide 内 rId 只被当前 shape 使用时允许原位修改。external target 不支持 payload 替换，继续明确报错。

## Image 写入

`replaceData(bytes, contentType?)` 在一个同步 transaction 中：

1. 定位当前 shape 的 `a:blip/@r:embed` 和 relationship。
2. 独占时对原 target `setPart()`。
3. shared 时在原目录分配 sibling URI，写入新 payload/content type。
4. relationship id 只被当前 shape 使用时，更新原 relationship target。
5. relationship id 也被其他 shape 使用时，新建 relationship，并只 patch 当前 shape 的 `r:embed`。

## Chart 写入

`setXml(value)` 先验证 XML。独占 chart 继续原位写；shared chart 使用 slide dependency lifecycle 的 owned-subgraph clone，复制 chart 及 workbook/style/color 等 owned 依赖，保留 shared/external 子依赖，再按 image 相同规则更新或新建 relationship。最后只把新 XML 写入 clone chart part。

relationship 重定向与 payload 写入属于同一个 transaction，失败时不得留下 clone part、relationship 或 slide XML patch。

## Identity 与保留

- Shape model identity 仍由 slide + shape id 决定，clone-on-write 不替换 model handle。
- 未编辑对象的 target URI、bytes、relationships 和 XML 保持不变。
- 新 relationship 使用正常 rId 分配器；现有 relationship 只在它由当前 shape 独占时更新。
- content type 与 cloned part 同步；未知 chart 子关系按 owned 子图规则保留。

## 测试与验收

1. duplicate slide 后 image 初始共享；编辑 duplicate image 后 target 分离，源 bytes 不变。
2. 独占 image 编辑保持原 URI，不创建 relationship/part。
3. 同一 rId 被两个 image shape 使用时，只重定向被编辑 shape。
4. 两个 slide 引用同一 chart 时，编辑一个会 clone chart/workbook 子图，另一 chart XML/target 不变。
5. 独占 chart `setXml()` 保持原 URI。
6. XML/relationship 故障注入触发 transaction rollback，无孤立 clone。
7. 编辑前后 shape model `===` identity 保持。
8. 全仓 typecheck、测试、Node/browser bundle 和 npm tarball smoke 通过。
