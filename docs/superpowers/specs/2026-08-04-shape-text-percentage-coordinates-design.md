# Shape and Text Percentage Coordinates Design

日期：2026-08-04
状态：已批准实施

## 目标

补齐 PptxGenJS 4.0.1 `Coord` 在最基础绘图对象上的合法百分比能力：
`SlideModel.addShape()`、`addCustomShape()` 和 `addText()` 接受相对当前幻灯片尺寸的
百分比坐标，同时保留 native API 已有的绝对 EMU 输入、创建事务和可编辑
`Transform` 快照。

本小项直接关闭以下 12 个审计原子：

- `PositionProps.x/y/w/h`；
- `ShapeProps.x/y/w/h`；
- `TextPropsOptions.x/y/w/h`。

PptxGenJS 使用 `w`/`h` 和“小于 100 的 number 视为英寸”的启发式；native API
继续使用明确的 `width`/`height`、`Emu` 与 `inches()`。两者覆盖相同合法最终状态，
但 API 形状和数值单位不同，因此这些原子记为 `deliberate-difference`，而不是声称
调用级兼容。

## 已锁定的 PptxGenJS 行为

在 10 × 8 英寸自定义布局中，PptxGenJS 4.0.1 将：

- `{ x: '10%', y: '20%', w: '30%', h: '40%' }` 序列化为
  `x=914400, y=1463040, cx=2743200, cy=2926080`；
- `{ x: '12.5%', y: '25%', w: '37.5%', h: '50%' }` 序列化为
  `x=1143000, y=1828800, cx=3429000, cy=3657600`。

上游实现对任意包含 `%` 的字符串使用 `parseFloat()`，会吞掉尾随垃圾并可能生成
`NaN` 或非法负 extent。这些 runtime 缺陷不属于合法能力；native 只接受完整、有限、
可安全解析的百分比字符串，并保证宽高为正。

## 方案比较

### 方案一：在形状和文本实现中分别解析字符串

改动局部，但会马上在图片、媒体、图表、表格和 master/layout 中复制相同规则，
错误边界容易漂移，不采用。

### 方案二：只在 PptxGenJS adapter 中转换

可以隐藏差异，但 native 从零创建 API 仍然缺少百分比能力，也不能服务后续对象，
不采用。

### 方案三：共享 coordinate 输入类型和内部解析器

采用。公开类型表达“绝对 EMU 或百分比”，内部解析器根据水平/垂直轴和当前
`SlideSize` 得到一个确定的 EMU。形状与文本先接入，后续对象逐项复用。

## 公共类型

`packages/model/src/units.ts` 新增：

```ts
export type SlideCoordinate = Emu | `${number}%`;

export interface TransformInput {
  readonly x: SlideCoordinate;
  readonly y: SlideCoordinate;
  readonly width: SlideCoordinate;
  readonly height: SlideCoordinate;
  readonly rotation: OoxmlAngle;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}
```

`AddShapeOptions` 与 `AddTextOptions` 改为 `Partial<TransformInput>`。
`Transform` 仍然只含已解析的 EMU；`BaseShapeModel.transform` 和
`setTransform()` 不接受百分比。百分比不是 OOXML 中可保留的语义，它只是在调用时
计算最终位置，因此编辑 API 继续暴露明确、可逆的绝对状态。

## 解析规则

内部 `resolveSlideCoordinate()` 接受值、轴、slide size、默认值和错误上下文：

- `undefined` 返回调用方给定的默认 EMU；
- finite number 按现有 native 规则视为 EMU，四舍五入并要求 safe integer；
- 百分比必须是无首尾空白、以单个 `%` 结尾、数值主体可由 `Number()` 完整解析的
  finite 字符串；
- `x`/`width` 乘以 slide width，`y`/`height` 乘以 slide height，并四舍五入；
- 解析结果必须是 safe integer；`-0` 规范化为 `0`；
- position 可为负或超过 100%，因为合法对象可以位于画布外；extent 必须大于 0；
- accessor、对象、空字符串、裸数值字符串、重复 `%`、尾随字符、`NaN` 和
  `Infinity` 均抛出，不进入序列化。

PptxGenJS 的普通 number 英寸能力由 native 的 `inches()` 明确覆盖。测试同时证明
`inches(1)` 与 `'10%'` 可以混用，避免把新增字符串支持误写成隐式单位转换。

## 数据流与事务

`SlideModel` 在创建事务内读取一次当前 `presentation.slideSize`，传给形状或文本
normalize/render 路径。所有百分比在写入 XML 前完成解析；任何错误都在 package
mutation 前抛出，或由现有 transaction 完整回滚。

Placeholder owner 的 transform 仍然优先于调用方坐标。`AddPlaceholderOptions` 复用
`AddTextOptions`，因此同一解析器不会破坏其类型与创建路径；但
`PlaceholderProps.x/y/w/h` 的 PptxGenJS 审计原子留到 master/layout placeholder
小项，用独立 control 证明继承和 override 语义后再关闭。

## 证据与测试

1. 内部单元测试锁定水平/垂直百分比、decimal/exponent、absolute EMU、默认值、
   负 position、超 100% 和严格非法矩阵。
2. shape/text 创建测试在 10 × 8 英寸布局中断言精确 `a:off`/`a:ext`，并验证
   preset shape、custom shape、text、reopen transform 和失败回滚。
3. adapter control 使用锁定的 PptxGenJS 4.0.1 生成相同 shape/text，比较最终 EMU，
   同时记录 `w/h`、英寸 number 与 native `width/height`、`Emu` 的 API 差异。
4. packed-package smoke 从公开声明编译百分比输入，创建并重开 package，断言精确
   transform。
5. OOXML/validator 证据证明输出在 PowerPoint 2010 profile 下没有 error。

该能力只改变最终几何，不需要额外关系、内容类型或视觉客户端专项；已有 OOXML
结构、重开和 validator 证据足以关闭本小项。

## 文件边界

- `packages/model/src/units.ts`：公开 input/output 类型边界；
- `packages/model/src/slide-coordinate.internal.ts`：唯一百分比解析实现；
- `packages/model/src/preset-shape.internal.ts`：shape/custom-shape 接入；
- `packages/model/src/slide.ts`：text 接入和当前 slide size 传递；
- 对应 model/adapter/packed tests：直接行为与发布证据；
- `scripts/pptxgenjs-surface-manifest.mjs`：12 个原子的直接映射；
- 生成的 compatibility JSON/Markdown：更新可信进度。

## 完成标准

- 三类对象在内建和自定义 slide size 下精确解析百分比；
- 绝对 EMU、placeholder owner、transform 读取/编辑和事务语义不回归；
- 非法输入不产生 package mutation；
- PptxGenJS control、native OOXML、packed consumer 和 validator 证据全部通过；
- 审计计数新增 12 个 `deliberate-difference`，没有 stale/diagnostic；
- focused tests、全量测试、typecheck、package smoke 和 diff review 全部通过。
