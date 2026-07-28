# 自定义幻灯片尺寸创建设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

让原生 `PptxDocument.create()` 接受任意 OOXML 合法幻灯片尺寸，覆盖 PptxGenJS 4.x `defineLayout()` 加 `layout` 选择后创建文稿的输出能力。调用方无需建立只在内存中存在的命名 layout registry，直接把最终宽高交给创建入口。

本小项只负责创建时的 presentation/notes page size。读取和修改已打开文稿的尺寸、根据尺寸重排既有对象、命名 slide master/layout，以及 shape 的百分比坐标分别属于后续小项。

## API

```ts
interface CustomSlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

interface CreatePresentationOptions {
  readonly format?: PresentationFormat;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
}

const document = PptxDocument.create({
  slideSize: {
    width: inches(11.7),
    height: inches(8.3),
  },
});
```

现有 `'4:3' | '16:9' | '16:10' | 'wide'` 完全兼容。自定义尺寸使用库统一的 EMU 类型；`inches()` 和 `points()` 可生成该类型，避免让不同 API 混用英寸和 EMU。

宽高序列化前取最接近的整数 EMU，并分别限制在 OOXML `ST_SlideSizeCoordinate` 的闭区间 914,400–51,206,400 EMU，即 1–56 英寸。非对象、数组、缺失宽高、非有限值或范围外数值明确抛出 `TypeError` 或 `RangeError`，不能生成 PowerPoint 需要修复的 package。

## 方案选择

考虑过三种方案：

1. 原样复制 `defineLayout({ name, width, height })`、可变 `layout` 属性和命名 registry。迁移表面最相似，但 registry 不写入 PPTX；同时引入与文稿模型无关的临时状态。
2. 只增加 `customWidth`/`customHeight` 两个 create 参数。实现简单，但两个字段可与内置 `slideSize` 冲突，类型无法保证成对出现。
3. 把 `slideSize` 扩展为内置名称或 `{ width, height }` 的判别联合。调用点只有一个尺寸来源，保持 create factory 原子且与现有单位系统一致，因此采用。

PptxGenJS 功能对等不要求参数形状相同。它的 layout 名称只用于在写出前选择最终宽高；本 API 直接表达最终宽高，生成的 `p:sldSz` 和 `p:notesSz` 语义相同。

## OOXML 与数据流

`createPresentationPackage()` 在创建空 `OpcPackage` 前解析尺寸：

1. 字符串必须命中四个现有 preset，并复用其精确 EMU 常量。
2. 对象必须只提供可读取的合法 width/height；两者各自四舍五入并验证范围。
3. `p:sldSz` 写入 `cx=width`、`cy=height`。
4. 与 PptxGenJS 和现有模板一致，`p:notesSz` 写入交换后的 `cx=height`、`cy=width`。

尺寸对象只在调用期间读取，不保存引用，也不修改调用方对象。后续创建 slide 不复制尺寸，因为页面尺寸属于 presentation 根级属性。

## 原子性与兼容性

格式和尺寸都在任何 package part 写入前验证。失败时不创建或返回半初始化文稿。成功后的自定义尺寸走现有 write、validation、Blob/download、format profile 和 LibreOffice/PowerPoint round-trip 路径。

不改变现有 preset 的数值、默认 `16:9` 行为或创建出的 master/layout/theme 关系链。`CustomSlideSize` 从 SDK 和聚合包公开导出，Node 与 browser bundle 使用同一代码路径。

## 测试与验收

1. 使用 `inches(11.7)` × `inches(8.3)` 创建文稿，`p:sldSz` 和交换后的 `p:notesSz` 为精确整数 EMU。
2. write → reopen 后 format、slides 和尺寸 XML 保持一致，package validation 无 error。
3. 内置四种 preset 和默认尺寸的现有断言继续通过。
4. 1 英寸与 56 英寸边界合法；零、低于 1 英寸、高于 56 英寸、NaN、Infinity、缺失字段和错误类型在创建前失败。
5. 传入冻结对象不会被修改；后续修改原对象不影响已创建 package。
6. 发布声明包含 `CustomSlideSize` 和扩展后的 `CreatePresentationOptions`；Node/browser/type/tarball smoke 覆盖自定义尺寸。
7. LibreOffice headless 无修复打开自定义尺寸的一页文稿并导出 PDF，PDF page size 与输入尺寸一致。
8. PptxGenJS parity matrix 将 `defineLayout()` 的自定义页面尺寸映射到 create 的尺寸对象并标记为支持；运行时尺寸编辑仍明确列为未支持。
