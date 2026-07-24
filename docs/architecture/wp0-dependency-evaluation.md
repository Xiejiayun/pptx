# WP0 ZIP/XML 依赖评估

## ZIP：JSZip 3.10

选用 JSZip 作为 0.1 基线。它能读取和重写常规 PPTX、访问原始 entry 名、验证 CRC，并保留未触及 entry 的解压 payload。缺点是保存时会重建 ZIP 容器，且大文件以内存为主；因此默认无改动时绕过 ZIP 生成直接返回原始 bytes，并通过资源预算限制输入。

后续在性能语料上同时评估流式 ZIP 后端。只要满足 `OpcPackage` 边界和无损测试，可替换实现而不影响语义模型。

## XML：自有 source-span parser

常规 DOM/SAX 库无法同时提供原始属性顺序、空白、命名空间前缀、注释位置和最小区间回写。WP0 实现只读 source-span tree 与无重叠 patch writer；它拒绝 DTD/ENTITY，不执行实体或网络访问。

当前 parser 有意保持小范围：负责可靠定位与局部替换，不尝试规范化 OOXML。后续能力通过 helper/codec 增量增加，未知 XML 保持 opaque。

## 已验证结论

- 无修改 package 输出与输入字节相同。
- 标题局部修改只重写 slide part。
- 未知 XML 节点仍在原始位置。
- 未触及二进制/XML part 的 SHA-256 不变。
- ZIP entry 数、单 part 大小、总解压大小、压缩比和路径均受检查。

