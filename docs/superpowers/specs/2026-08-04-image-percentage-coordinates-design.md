# Image Percentage Coordinates Design

日期：2026-08-04
状态：已批准实施

## 目标

让 raster 与 SVG 图片的所有 native 创建入口接受相对当前幻灯片尺寸的百分比
`x`、`y`、`width` 和 `height`，同时保留绝对 EMU、图片 sizing、placeholder owner、
最终可编辑 `Transform` 以及现有 source loader 的事务边界。

本小项关闭 PptxGenJS 4.0.1 的 4 个直接图片坐标原子：

- `interface:ImageProps@property:x`；
- `interface:ImageProps@property:y`；
- `interface:ImageProps@property:w`；
- `interface:ImageProps@property:h`。

PptxGenJS 使用 `w`/`h` 和隐式英寸 number；native 继续使用 `width`/`height`、
`SlideCoordinate`、`Emu` 与 `inches()`。两者覆盖相同合法最终几何，但调用形状和数值
单位不同，因此 4 个原子记为 `deliberate-difference`。

## 已锁定的 PptxGenJS 行为

在 10 × 8 英寸自定义布局中，PptxGenJS 4.0.1 将一个 raster 图片的
`{ x: '10%', y: '20%', w: '30%', h: '40%' }` 序列化为：

- `x=914400`；
- `y=1463040`；
- `cx=2743200`；
- `cy=2926080`。

native control 使用相同布局、同一 PNG payload 和
`{ x: '10%', y: '20%', width: '30%', height: '40%' }`，必须产生完全相同的最终
transform。PptxGenJS 对普通 number 的隐式英寸解释由 native 的 `inches()` 明确覆盖。

## 范围边界

本小项覆盖：

- `SlideModel.addImage()` 的 raster 图片；
- `SlideModel.addSvgImage()` 的 SVG + fallback 图片；
- `PptxDocument.addImage()` 的 raster/SVG source loader；
- 复用上述入口的 master/layout 图片创建；
- 顶层显式图片 rectangle 的百分比坐标。

本小项不改变 `ImageSizing`。`contain`、`cover` 和 `crop` 的 frame `width`/`height`
继续使用绝对 EMU，crop source 继续使用源图片像素；对应
`inline:interface:ImageProps@property:sizing...` 原子留到独立 sizing 小项。提供 `sizing`
时，计算出的具体 frame 宽高继续覆盖顶层宽高，顶层 `x`/`y` 百分比仍然有效。

图片 transform 的读取和编辑继续返回、接受绝对 EMU。百分比只在创建时解析，不是
OOXML 中可保留的语义。图片 hyperlink、rounding、shadow 等其他未验证原子不在本项
范围内。

## 方案比较

### 方案一：raster 与 SVG 分别解析百分比

改动局部，但两条路径会复制相同解析和错误规则，后续 source loader 与
master/layout 也容易出现不同边界，不采用。

### 方案二：只在高层 source loader 转换

可以覆盖 `PptxDocument.addImage()`，但低层 `SlideModel.addImage()`、
`addSvgImage()` 和 master/layout 仍缺少百分比能力，不采用。

### 方案三：共享图片 appearance 正规化接入 coordinate resolver

采用。`normalizeEmbeddedRasterImage()` 是 raster appearance 的唯一正规化边界，
SVG 路径已经通过它复用 transform 处理。让这条路径接收当前 `SlideSize` 并调用现有
`resolveSlideCoordinate()`，即可覆盖所有图片创建入口且不复制解析器。

## 公共类型与数据流

`AddImageOptions` 与 `AddSvgImageOptions` 从 `Partial<Transform>` 改为
`Partial<TransformInput>`。高层 `AddImageSourceOptions` 在未使用 `sizing` 的分支中把
`width`/`height` 从 `number` 改为 `SlideCoordinate`；`x`/`y` 已从底层图片 options
继承同一类型。

数据流为：

1. 高层 source loader 严格读取、复制 options 并解析图片来源；
2. `commitPreparedImage()` 把顶层 coordinate 输入或具体 sizing 结果交给 slide；
3. `SlideModel.addImage()` / `addSvgImage()` 把当前
   `this.presentation.slideSize` 传给图片正规化器；
4. raster 与 SVG 共享 appearance 路径，把百分比解析为具体 EMU；
5. placeholder owner 若存在，继续以 owner transform 覆盖调用方 geometry；
6. renderer 只接收具体整数并写入 `a:off` / `a:ext`。

内部正规化函数允许 numeric-only 测试调用不传 `SlideSize`，但任何百分比在没有
slide size 时必须抛出。所有真实 slide 创建路径都必须传入当前尺寸。

## 验证与错误语义

百分比字符串复用共享 coordinate 规则：水平字段按 slide width，垂直字段按 slide
height，并对最终 EMU 四舍五入。既有 numeric 图片输入继续要求输入本身就是 safe
integer，不采用 resolver 对 numeric 小数的四舍五入，以免放宽现有图片契约。
position 可为负或超过 100%，width/height 必须解析为正 safe integer；空白、尾随字符、
非有限值、对象、数组与 accessor 均不接受。

低层入口在任何关系、media part 或 XML 写入前完成正规化。高层异步 source loader
可能先读取图片来源再发现合法容器内的错误 coordinate，但仍必须在 package mutation
前失败并保持 parts、relationships、shape IDs、mutation journal 与输出 bytes 不变。
这与现有高层 `width: 0` 和非法 crop 的失败边界一致。

## 证据与测试

1. raster normalizer 单测在 10 × 8 英寸尺寸中断言 decimal 百分比和绝对 EMU，拒绝
   无 slide size、零/负 extent 与严格非法输入。
2. SVG normalizer 单测证明同一 shared appearance 路径解析百分比。
3. SDK 端到端测试通过高层 source loader 创建 raster 和 SVG 百分比图片，断言即时
   transform、精确 `a:off` / `a:ext`、PowerPoint 2010 diagnostics、reopen 结果、
   placeholder precedence 和失败隔离。
4. adapter control 使用真实 PptxGenJS 4.0.1 自定义布局比较相同 PNG 的最终 EMU，
   并用 `inches(1)` 锁定 numeric-unit 差异。
5. packed-package smoke 从发布声明编译 `SlideCoordinate` 图片 options，创建并重开图片，
   报告 `imagePercentageCoordinates: true`。
6. manifest 为 4 个 `ImageProps` 原子登记 code、test、package、OOXML 与 control 证据，
   重新生成稳定 JSON/Markdown。

## 文件边界

- `packages/model/src/image.ts`：图片创建输入类型；
- `packages/model/src/image-create.internal.ts`：共享 raster/SVG appearance coordinate
  解析；
- `packages/model/src/svg-image-create.internal.ts`：传递 slide size；
- `packages/model/src/slide.ts`：把当前 slide size 传给两类图片正规化器；
- `packages/sdk/src/raster-image-source.ts`：高层显式 rectangle 类型；
- model/SDK/adapter/packed tests：行为、PptxGenJS 与发布包证据；
- `scripts/pptxgenjs-surface-manifest.mjs` 和生成的 compatibility 产物：关闭 4 个原子。

## 完成标准

- raster、SVG、高层 source loader 与 master/layout 图片可使用顶层百分比 rectangle；
- 严格 safe-integer 绝对 EMU、默认一英寸、sizing、crop、placeholder、图片编辑和
  source lifecycle 不回归；
- 非法输入不产生 package mutation；
- PptxGenJS control、OOXML、reopen、PowerPoint 2010 profile 与 packed declarations 全部通过；
- 审计计数变为 `supported=7`、`deliberate-difference=16`、
  `defect-excluded=1`、`unverified=1750`、`unsupported=0`、`stale=0`，诊断为零；
- focused/full tests、typecheck、package smoke、生成物确定性和 diff review 全部通过。
