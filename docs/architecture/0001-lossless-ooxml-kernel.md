# ADR 0001：采用无损 OOXML 双向内核

- 状态：Accepted
- 日期：2026-07-25

## 决策

PPTX 以 OPC package graph 管理，XML 使用保留源跨度的树和局部 patch。无 mutation 时直接返回输入字节；有 mutation 时只重写发生变化的 part，其他 entry 的解压 payload 必须保持一致。

PptxGenJS 只存在于 adapter 的依赖边界内。核心不能读取其私有对象结构，而是接收其公开 API 生成的 PPTX bytes。

## 原因

DOM 全量序列化会改变命名空间前缀、属性顺序、空白和 Office 扩展节点。以语义模型重建整个文件也无法安全保留未知部件。source-span patch 可以把改动限定在最小 XML 区间，OPC 图则负责关系和 content type 不变量。

## 无损定义

1. 无 mutation：输出字节与输入完全相同。
2. 有 mutation：未触及 part 的解压 payload SHA-256 相同。
3. 未识别节点、关系、扩展命名空间和二进制 part 原样保留。
4. ZIP entry 顺序、压缩参数和时间戳不属于默认语义无损范围。

## 后果

已识别功能由 codec 拥有并生成局部 patch；未知内容默认 opaque-preserve。删除、移动或覆盖未知内容必须显式失败，不能静默丢失。

