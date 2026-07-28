# 稳定语义对象 Identity 设计

日期：2026-07-28
状态：已批准实施

## 目标

让当前公开的 class-based 语义对象在 document 生命周期内具有稳定 identity。重复读取同一 slide、shape、master、layout 或 theme 时返回同一个 JavaScript 对象实例；移动 slide、编辑文本/transform/theme 等不改变 OOXML identity 的操作也不得替换实例。

本小项不缓存 XML DOM，也不把 `MediaModel`、placeholder、table row/cell、chart series 等当前值快照接口伪装成持久对象。它们会在后续引用生命周期和完整对象模型小项中升级。

## 方案选择

考虑过三种方案：

1. 继续每次读取创建快照对象，只保证 `partUri`/`id` 相同。实现最简单，但调用方无法可靠保存对象引用，也不满足总体设计的稳定 identity。
2. 用一个以 `OpcPackage` 为键的全局 `WeakMap` 保存所有对象。可以跨 codec 共享，但会把 model、codec 和 package 层耦合到全局状态，测试隔离和自定义 codec 较差。
3. 由对象集合的直接 owner 维护 registry：`PresentationModel` 管 slide，`SlideModel` 管 shape，`MasterLayoutThemeCodec` 管 master/layout/theme；`PptxDocument` 复用同一个内建 codec 实例。边界清晰、无全局状态，也为后续 `DocumentSession` 汇总这些 registry 保留迁移路径，因此采用此方案。

## Identity Key 与集合语义

- Slide：规范化 slide part URI。
- Shape：所属 slide identity 加 `p:cNvPr/@id`。
- Master、Layout、Theme：各自 part URI，并限定在同一个 codec/document owner 内。

集合每次读取仍从当前 presentation XML、relationship graph 或 slide XML 计算顺序和成员，不缓存旧数组。registry 只复用成员对象，因此 slide move 会改变数组顺序但保持每个 slide 的 `===` identity。

同一个 key 的 registry entry 在 owner 生命周期内保留。part/shape 暂时不存在时，持有的 handle 通过现有缺失 part/shape 错误失败；事务回滚或同 key 内容恢复后，同一 handle 可继续解析。这样 cache 不需要在嵌套事务中维护第二套回滚日志。

## Live 属性与类型变化

语义对象继续在 getter 调用时从当前 OOXML 解析数据：

- `SlideModel.relationshipId` 与 `slideId` 在 collection reconciliation 时同步当前值。
- Shape `name` 改为 live getter，不把首次 decode 的名称永久缓存。
- transform、text、rows、series、theme colors/fonts 和 relationship target 保持现有 live 读取行为。

如果 raw/opaque mutation 在保留同一个 shape id 的同时把 shape 改成不兼容的语义 class（例如 image 变成 table），registry 替换该 entry；稳定 identity 只覆盖保持 OOXML identity 和语义类型的受支持编辑。

## Owner Registry

`PresentationModel.slides` 根据当前 `p:sldId` 顺序取得或创建缓存 `SlideModel`，并同步 relationship/slide id 元数据。`attachSlide()` 也必须通过同一 registry 返回对象，保证 create/duplicate 的返回值与随后 collection 读取一致。

`SlideModel.shapes` decode 当前元素描述后，按 shape id 和语义 class 复用缓存对象。返回数组是新数组，元素对象稳定。

`MasterLayoutThemeCodec` 为三类 part 分别维护 registry。所有 getter 和 create/copy 返回路径都经过 registry。`MasterModel.layouts/theme` 也复用所属 codec registry。

`PptxDocument` 构造时只创建一个 `MasterLayoutThemeCodec`，同时用于 codec registry、`masters`/`layouts`/`themes` 和 `masterLayoutTheme` getter，避免相同 document 出现多个 identity 域。

## 错误与内存边界

- 已删除对象的 live getter 使用现有 `PackageError` 或 `ModelParseError`，不返回陈旧快照。
- registry 不保存 XML document 或 part bytes，只保存轻量 model handle。
- registry 大小受当前 document 生命周期内观察过的 OOXML identity 数量限制；关闭对 document/codec 的引用后可整体回收。
- 只读 collection 访问不得产生 OPC mutation，也不得改变无修改 write 的字节 identity。

## 测试与验收

1. 重复读取 `slides`、`shapes`、`masters`、`layouts`、`themes` 的对应成员满足 `===`。
2. slide move、标题/shape transform/表格编辑和 theme color 编辑后 identity 保持。
3. create/duplicate 返回的 slide/master/layout/theme 与随后 collection 中对象相同。
4. transaction rollback 后，原 slide/shape handle 仍是 collection 返回对象并读取回滚后的内容。
5. shape name 修改后已有 handle 读取最新值；保持 id 但改变语义 class 时创建正确的新类型对象。
6. `document.masterLayoutTheme` 重复读取稳定，且与 `document.codecRegistry` 中注册实例相同。
7. 只读 identity 测试保持 no-mutation byte identity。
8. 全仓 typecheck、测试、Node/browser bundle 和 npm tarball smoke 通过。
