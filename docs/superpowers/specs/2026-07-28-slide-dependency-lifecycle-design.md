# Slide 依赖引用生命周期设计

日期：2026-07-28
状态：已批准实施

## 目标

修复 slide duplicate/delete 对 relationship 依赖的生命周期处理。复制 slide 时不再直接复制 `.rels` 文件；独占依赖深拷贝、共享依赖继续引用原 part、external 原样保留。删除 slide 时，只回收引用归零的 owned 依赖子图，不删除共享资源或 opaque 目标。

本小项聚焦 slide 依赖图。Image/Chart handle 对任意已共享 part 的通用 clone-on-write 属于下一小项；本小项通过 duplicate 时深拷贝 chart 等 owned 依赖，先消除现有最直接的共享串改风险。

## 方案选择

考虑过三种方案：

1. 继续复制 slide XML 和原 `.rels` bytes。relationship id 保持容易，但所有内部 target 都被共享，chart/notes/comments 的修改会串改原 slide。
2. 对 slide 的所有内部 target 全部递归复制。隔离最强，但会复制 layout、theme、图片和媒体，破坏有意共享、显著膨胀文件，并可能形成错误的 master 链。
3. 根据 relationship lifecycle 重建关系图：owned 深拷贝，shared 复用，external 保留，根级 opaque 保守复用；owned 子图内部默认随 owner 克隆，但明确 shared/back-reference 例外。该方案符合总体架构并能渐进扩展，因此采用此方案。

## 生命周期分类

Slide 根关系按 type suffix 分类：

- `owned`：chart、notesSlide、comments、SmartArt data/layout/style/colors/drawing，以及独占 OLE/package/control 等对象 part。
- `shared`：slideLayout、image、audio、video、media、theme、slideMaster、notesMaster，以及其他明确可复用的资源。
- `external`：任何 `TargetMode="External"`，保留 URL 和 relationship id，绝不下载。
- `opaque`：尚未声明策略的根关系。复制时保留指向同一 target，删除时不做 target GC；后续统一验证会提供诊断。

进入 owned part 子图后，内部依赖默认跟随 owner 深拷贝；明确 shared 类型、external 关系以及已知 back-reference 除外。这样 chart 的 workbook/style/color、notes/comments 的附属内容和 SmartArt part set 可以形成独立闭包。

## 深拷贝算法

`duplicateSlide()` 在一个 mutation transaction 中：

1. 分配并写入新 slide part，先记录 `source slide URI → clone slide URI`。
2. 遍历源 slide relationships，保留每个原 relationship id。
3. shared/external/opaque relationship 指向原 target；owned relationship 递归 clone target part。
4. 递归 clone 使用 source→clone map 处理环和 back-reference；已映射的源 slide 自动指向新 slide。
5. 每个 cloned part 保留 payload 与 content type，在原目录按 basename stem/extension 分配新 URI。
6. 用 `addRelationship()` 重建 `.rels` 并重新计算相对 target，不直接复制 relationship XML。
7. 最后把新 slide 挂到 presentation。

任何 part、relationship 或 presentation XML 写入失败时，外层 transaction 回滚完整 clone 子图。

## 删除与垃圾回收

删除 slide 前记录其 owned root 及可回收子图。移除 presentation entry、presentation relationship 和 slide part 后，从 owned root 开始检查 package graph incoming：

- incoming 为零才删除该 owned part；
- 删除前先记录其 owned/default-owned 子依赖，随后递归检查；
- shared、external 和根级 opaque target 永不由 slide GC 删除；
- 任意 target 仍被其他 part 引用时立即保留，并停止向其子图传播删除。

GC 与 slide 删除处于同一个 transaction，失败不得留下半删除子图。

## 兼容与保留

- Slide XML 中引用的 `rId` 保持不变，避免改写已知或未知 XML 节点。
- 未识别 relationship 不丢失、不下载、不猜测所有权。
- 共享 image/media payload 在 duplicate 后仍只保留一份。
- owned chart/workbook、notes/comments 和 SmartArt payload 在 duplicate 后使用独立 part URI。
- no-mutation write 和未触及 opaque payload 保持现有保留保证。

## 测试与验收

1. duplicate 后 slide XML 相同且所有 relationship id 保持；`.rels` 由 graph API 重建。
2. image/layout 等 shared target URI 与源 slide 相同。
3. chart target URI 与源 slide 不同，series/XML 修改 clone 不影响源 chart。
4. owned chart 的 embedded workbook 等内部依赖递归 clone，external 和 shared 子依赖不 clone。
5. 删除 duplicated slide 回收其无引用 owned 子图，但保留源 chart、共享 image/layout/media。
6. 根级 opaque relationship 在 duplicate/delete 中保留 target 且不被 GC。
7. cycle/back-reference clone 不递归失控，notes slide back-reference 指向 clone slide。
8. 故障注入验证 clone/GC transaction 无孤立 part、悬空关系或 journal 残留。
9. 全仓 typecheck、测试、Node/browser bundle 和 npm tarball smoke 通过。
