# PPTX 双向编辑库实施进度

最后更新：2026-07-25

## WP0：基线与技术验证

状态：完成

### 本阶段 change

- 建立 pnpm/TypeScript strict/Vitest monorepo。
- 新增 `@pptx/lossless-xml`：保留源跨度、属性顺序、空白和未知节点；仅重写目标区间；拒绝 DTD/ENTITY。
- 新增 `@pptx/opc`：读取 content types、内部/外部 relationships、规范化 part URI、建立 package graph，并实施 ZIP 资源预算。
- 新增基础 validator diagnostic 模型。
- 新增 `@pptx/sdk` 首条竖切：从 Buffer、Uint8Array、ArrayBuffer、文件或 stream 打开，读取/修改第一页标题并保存。
- 无 mutation 时原字节返回；有 mutation 时未触及 part 的 payload SHA-256 保持一致。

### 新增功能演示

下面的文件由真实 PPTX 经过 `PptxDocument` 修改标题后，用 LibreOffice headless 打开并导出。页面可正常渲染，证明输出未触发结构修复。

![WP0 标题无损编辑演示](./images/wp0-title-edit.png)

### 验证结果

- TypeScript strict typecheck：通过。
- Vitest：3 个测试文件、8 个测试全部通过。
- 无修改 round-trip：字节级相同。
- 标题 mutation isolation：通过。
- 未知 XML 节点保留：通过。
- LibreOffice 打开/导出：通过。

### 相关设计记录

- [ADR 0001：无损 OOXML 内核](./architecture/0001-lossless-ooxml-kernel.md)
- [ADR 0002：Codec ownership](./architecture/0002-codec-ownership.md)
- [ADR 0003：兼容 profile](./architecture/0003-compatibility-profiles.md)
- [WP0 依赖评估](./architecture/wp0-dependency-evaluation.md)

