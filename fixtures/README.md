# 测试语料

语料按风险分类：

- `corpus/minimal`：最小 OPC、关系、标题编辑和未知扩展。
- `corpus/clients`：PowerPoint、Keynote、LibreOffice 和 Google Slides 产生的授权文件。
- `corpus/features`：master/layout/theme、渐变、媒体、图表和插件功能。
- `corpus/security`：ZIP traversal、压缩炸弹、DTD/ENTITY 和悬空关系。
- `expected`：语义快照、diagnostic 和允许变化的 part 清单。

二进制 fixture 必须记录来源和许可。可程序化构造的 fixture 默认在测试内生成，避免把无来源文件提交到仓库。

