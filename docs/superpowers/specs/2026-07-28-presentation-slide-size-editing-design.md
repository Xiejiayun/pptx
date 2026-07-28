# Presentation 页面尺寸读写设计

日期：2026-07-28
状态：已批准实施

## 目标与范围

为原生和打开的文稿提供可读写 `slideSize`，让页面尺寸不再只在 `PptxDocument.create()` 时决定。调用方可以读取现有 `p:sldSz`，也可以在同一 lossless model 上修改画布宽高后继续编辑并写出。

本小项只改变 presentation 画布。它不缩放、移动或裁剪既有 shape，不改变 notes page size，也不重建 master/layout。自动“确保适合”或“最大化”式的内容重排、notes size 独立 API 和 preset 名称识别分别属于后续小项。

## API

```ts
interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

const current = document.slideSize;
document.slideSize = {
  width: inches(11.7),
  height: inches(8.3),
};
```

`SlideSize` 属于 model 的公共单位类型。SDK 的现有 `CustomSlideSize` 扩展该类型，因此创建和编辑使用相同字段与 `inches()`/`points()` 单位，不破坏已发布的 create 调用。

getter 每次从当前 presentation XML 返回新的值对象，不缓存或暴露内部 XML。setter 立即读取传入值并写入整数 EMU，不保留调用方对象引用。

## 方案选择

考虑过三种方案：

1. 只提供 `getSlideSize()`/`setSlideSize()` SDK helper。改动局部，但尺寸是 presentation 的核心语义，放在 SDK 会让 model 与高层 API 能力不一致。
2. 提供 `PresentationModel.slideSize` 属性，只修改画布。API 与现有 model 属性一致，打开和创建的文稿走同一路径，且不会替用户猜测内容缩放策略，因此采用。
3. 设置尺寸时按宽高比例自动缩放全部 shape。表面方便，但 group、placeholder、master object、table/chart 和裁剪图片需要不同规则；静默修改会破坏 lossless 编辑边界。

## 读取与 OOXML 写入

getter 定位 presentation 根的直接 `p:sldSz`，解析 `cx`/`cy` 为 width/height。元素缺失、属性缺失、非整数、非有限或超出 OOXML 1–56 英寸范围时抛出带 presentation part URI 的 `PackageError`，不伪造默认尺寸。

setter 在 package transaction 中：

1. 在任何 XML mutation 前验证对象、width/height 类型、四舍五入后的整数 EMU 和 914,400–51,206,400 范围。
2. 若直接 `sldSz` 已存在，只替换或补齐 `cx`/`cy`，保留 `type`、命名空间、未知属性、相邻空白和所有其他 children。
3. 若 `sldSz` 缺失但直接 `notesSz` 存在，在 `notesSz` 前插入自带 presentation namespace 的 `p:sldSz`，满足 schema child order。
4. 若 presentation 根或必需的 `notesSz` 缺失，抛出 `PackageError` 并恢复 part bytes 与 mutation journal。

`notesSz` 刻意保持原样。它描述 notes 页面而非 slide canvas；原生 create 仍按模板生成交换方向的 notes size，但编辑 slide 不应静默覆盖已有 notes 排版。

## 原子性与兼容性

setter 使用现有 OPC transaction，可嵌套在 `document.transaction()` 中。内部失败或后续外层 rollback 都必须恢复原始 presentation bytes、mutation journal 和 getter 结果。

现有 slide/master/layout/shape model identity 不受影响。设置尺寸后已存在和新增 shape 的 EMU transform 不变。未知根属性、`extLst` 和 namespace prefix 原样保留，除目标 `sldSz` 的两个坐标外不重写语义。

## 测试与验收

1. 创建和打开的文稿都能读取精确 width/height，返回对象不是内部可变引用。
2. 设置 A4 尺寸后 write → reopen 的 getter 和 `p:sldSz` 一致；已有 slide shape transform 不变。
3. 现有 `sldSz/@type`、未知属性、`notesSz`、根 `extLst` 和其他 opaque 内容保持不变。
4. 缺少可选 `sldSz` 时在 `notesSz` 前合法插入；缺失 presentation 根或必需 `notesSz` 时无 mutation 失败。
5. 非对象、数组、缺失字段、NaN、Infinity、低于 1 英寸或高于 56 英寸的 setter 输入抛出明确错误且不改变 bytes/journal。
6. 外层 transaction rollback 恢复尺寸；随后 getter 仍反映旧 XML。
7. 全仓测试、性能、Node/browser bundle、发布类型和 npm tarball smoke 通过，公开声明同时包含 `SlideSize` 与兼容的 `CustomSlideSize`。
8. LibreOffice headless 打开修改尺寸后的现有文稿并导出 PDF，页面尺寸与新画布一致且不报告修复。
