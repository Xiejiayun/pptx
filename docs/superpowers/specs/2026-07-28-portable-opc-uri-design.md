# 可移植 OPC Part URI 设计

日期：2026-07-28  
状态：已批准实施

## 目标

移除 `@pptx/opc`、`@pptx/model` 和内建 codecs 对 `node:path` 的依赖，使用一套只处理 OPC 正斜杠 part URI 的可移植工具。该小项不改变公开演示文稿语义，目的是让核心 package graph、model 和 codecs 可以进入后续浏览器构建。

## API 与职责

`@pptx/opc` 继续拥有 URI 规范化，并新增以下纯函数：

- `partUriDirname(uri)`：返回规范化 part URI 的父目录。
- `partUriBasename(uri, suffix?)`：返回最后一段，可选择移除确定后缀。
- `partUriExtension(uri)`：返回包含点号的扩展名，没有扩展名时为空字符串。
- `joinPartUri(base, ...segments)`：连接并规范化绝对 OPC URI。
- `relativeRelationshipTarget(sourcePartUri, targetPartUri)`：从 source part 所在目录生成内部 relationship target。

`normalizePartUri()` 改为纯字符串 segment 归约：折叠重复 `/` 和 `.`，解析 `..`，一旦越过 package root 就抛出 `PackageError`。它不把反斜杠视作合法分隔符，也不做 URL decode。

## 迁移范围

- OPC content type、relationship part URI、source part URI 和 target resolution 改用新工具。
- PresentationModel 的 slide part 分配和 relationship target 改用新工具。
- Master/Layout/Theme codec 的目录、basename 和 relationship target 改用新工具。
- Media codec 的 relationship target 改用新工具；本地文件扩展名只做分隔符无关的字符串提取。
- Transition 插件删除自己的相对路径实现，复用 OPC 工具。

Node 文件系统、Node stream 和加密依赖不属于本小项，将在后续浏览器 I/O/媒体小项分别处理。`@pptx/testkit` 保持 Node-only。

## 兼容与安全

现有合法 package URI 的输出保持一致。ZIP entry traversal 继续在加载时拒绝；公共 URI API 对 `/../x`、`../../x` 等越界输入也必须拒绝。relationship target 始终使用 `/`，不得受宿主操作系统路径规则影响。

## 测试与验收

1. URI normalize、dirname、basename、extension、join、relative 和 relationship part/source 映射都有边界测试。
2. traversal、反斜杠和根目录边界测试通过。
3. model、master、media、transition 的既有 mutation 测试保持通过。
4. `rg "node:path"` 在 OPC、model、codecs 和功能插件源码中无结果。
5. 全仓 typecheck、测试和聚合包构建通过。
