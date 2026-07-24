# ADR 0002：Feature codec ownership

- 状态：Accepted
- 日期：2026-07-25

每个 codec 必须声明其拥有的元素、关系类型和 part content type。一个节点在同一优先级只能有一个 owner；冲突产生 `CODEC_OWNERSHIP_CONFLICT` diagnostic。codec 只能重写自己拥有的最小子树，不得删除未知兄弟节点或 `extLst` 内容。

内建 codec 优先级为 100，显式注册的应用 codec 默认为 200，透传 codec 为 0。覆盖内建 codec 必须由调用方明确配置，不允许安装插件后静默改变输出。

