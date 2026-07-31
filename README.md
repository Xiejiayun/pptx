# PPTX OOXML

一个面向现有 PPTX 的 TypeScript 双向编辑内核。它把演示文稿作为 OPC package graph 读取，通过 source-span XML patch 做局部修改，并默认保留未知 OOXML、扩展节点和二进制部件。

当前实施进度和功能截图见 [docs/implementation-progress.md](./docs/implementation-progress.md)。完整路线图见 [PLAN.md](./PLAN.md)。

```sh
npm install @jiayunxie/pptx@next
```

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = await PptxDocument.open('input.pptx');
document.slides[0].title.text = 'Updated';
await document.writeFile('output.pptx');
```

## 从路径、URL、data URI、Blob、stream 或 bytes 添加图片

```ts
import { readFile } from 'node:fs/promises';
import {
  degrees,
  inches,
  inspectRasterImage,
  PptxDocument,
  type AddImageSourceOptions,
} from '@jiayunxie/pptx';

const options: AddImageSourceOptions = {
  name: 'Quarterly chart',
  altText: 'Revenue by quarter',
  x: inches(1),
  y: inches(1.5),
  width: inches(5),
  height: inches(3),
  rotation: degrees(10),
};

const document = PptxDocument.create();
document.addSlide();
const image = await document.addImage(0, 'chart.png', options);
const info = inspectRasterImage(new Uint8Array(await readFile('chart.png')));
console.log(info); // { contentType: 'image/png', width, height }
image.setTransform({ x: inches(1.5) });
image.replaceData(new Uint8Array(await readFile('chart-updated.png')), 'image/png');
await document.writeFile('images.pptx');
```

`PptxDocument.addImage()` 接受 Node 本地路径、HTTP/HTTPS URL、浏览器相对 URL、strict base64 data URI、`Uint8Array`、`ArrayBuffer`、`Blob`/`File`、Web `ReadableStream` 和 async byte iterable。格式只以 bytes signature 为准，支持 PNG/JPEG/GIF；`AddImageSourceOptions.contentType` 是可选 assertion，不一致会在 package mutation 前拒绝。`signal` 可中止文件、Fetch、Blob 和 stream 加载。文件扩展名、`Blob.type`、HTTP `Content-Type` 和 `File.name` 不参与格式判断。

`inspectRasterImage()` 返回 signature 检测出的 canonical content type 与 raw pixel width/height；PNG 读取 direct IHDR，GIF 读取 logical screen，JPEG 安全遍历所有尺寸型 SOF marker。当前仍保留 PptxGenJS 的 1-inch 默认 transform，不会自动把 intrinsic pixels 转为布局尺寸。底层同步 `SlideModel.addImage(bytes, { contentType, ... })` 仍用于调用方已持有 bytes 的严格原子创建。

每次创建原子地拥有一个唯一 media part、一个 internal image relationship 和一个 canonical rectangular picture；任何验证或写入失败都会回滚 package、关系、slide XML 与 mutation journal。复制页面先共享图片 part，`ImageModel.replaceData()` 对独占 target 原位更新，对共享 target clone-on-write；六种 presentation format 都可创建、写出和重开。

锁定 PptxGenJS 4.0.1 的公开 path/data PNG/JPEG/GIF 输出与高层 loader 形成相同最终语义。实际 npm tarball 的 Node/browser/types/CLI smoke 与 4 页、32 shapes、12 图片 gallery 已通过；原件和 LibreOffice 回存件均为 PowerPoint 2010 validation 0 errors / 0 warnings，2400×1350 渲染无 overflow，并逐页检查。LibreOffice 保留 12/12 图片的 payload SHA、content type、名称、非空 alt text、顺序和 internal relationship，transform 最大量化 432 EMU；它会把 12 个重复 payload target 去重为 3 个并重写全部 picture markup。

当前仍未提供 contain/cover/crop sizing、SVG、rounding/transparency、alt-text 编辑、图片 hyperlink/shadow/placeholder，以及单图片删除与 media GC。

## 创建和编辑预设形状、调整值与样式

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();
const shape = slide.addShape('roundRect', {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 20,
  },
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '1F4E78' },
    transparency: 10,
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
  shadow: {
    kind: 'outer',
    color: { kind: 'srgb', value: '000000' },
    opacity: 0.35,
    blur: 6,
    angle: 45,
    distance: 4,
  },
  hyperlink: {
    url: 'https://example.com/docs',
    tooltip: '打开文档',
  },
});

shape.fill = { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } };
shape.fill = { kind: 'none' };
shape.fill = undefined;
shape.line = { kind: 'line', color: { kind: 'scheme', value: 'accent2' } };
shape.line = { kind: 'none' };
shape.line = undefined;
shape.arrows = { begin: 'diamond' }; // 同时清除省略的 end
shape.arrows = { begin: 'none', end: 'oval' };
shape.arrows = undefined; // 清除两端，保留线条样式
shape.shadow = {
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0.5,
  blur: 3,
  angle: 270,
  distance: 2,
};
shape.shadow = undefined;
document.addSlide(); // 创建第 2 页作为内部链接目标
shape.hyperlink = { slide: 2, tooltip: '前往详情' };
shape.hyperlink = { url: 'mailto:team@example.com', tooltip: '' };
shape.hyperlink = undefined;
```

`ShapeModel.fill` 支持 direct solid/no-fill 的创建、读取、编辑与清除。`{ kind: 'none' }` 写入明确的 direct no-fill，`undefined` 只清除 direct fill state；gradient、pattern、picture 和 group fill 不属于这个 simple-fill API。

`ShapeModel.line` 支持 direct none/solid line 的创建、读取、编辑与清除，包括 sRGB/theme color、0–100% transparency、0–1584pt width 和 8 种 preset dash。省略 width/dash 默认 1pt/solid；`undefined` 只清除 line 的 width/fill/dash，同时保留 line 容器、箭头、join 和扩展节点。

`AddShapeOptions.arrows` 与 `ShapeModel.arrows` 支持 begin/end 的 `none | arrow | diamond | oval | stealth | triangle`。快照与输入脱离；赋值采用 whole replacement，缺失的一端会被清除，显式 `none` 则保留对应 direct endpoint。`undefined` 只清除两端而保留 line，反向的 `shape.line = undefined` 也保留 arrows。只创建 arrows 不会隐式生成颜色、宽度或 dash；已有合法 `w` / `len` size 会在类型编辑中无损保留，但 size 创建/读取/编辑尚未公开。

`AddShapeOptions.shadow` 与 `ShapeModel.shadow` 支持 direct outer/inner shadow 的创建、读取、whole replacement 与清除，包括 sRGB/theme color、`0..1` opacity、`0..100pt` blur、`0..<360°` angle、`0..200pt` distance，以及 outer-only `rotateWithShape`。默认值为 black、0.75、8pt、270°、4pt 和 outer rotate false；显式 zero 会保留。输入在 mutation 前深度脱离，getter 的嵌套快照会 deep-freeze；同值赋值是 exact no-op，`undefined` 只移除 direct shadow 并保留 `effectLst` 与 glow/reflection 等 sibling effects。Generic/advanced effects、custom shadow transforms，以及 text/image/table/chart/media 等非 preset-shape shadow API 仍待后续小项。

`AddShapeOptions.hyperlink` 与 `ShapeModel.hyperlink` 支持整个 preset shape 的 click URL 或内部页链接。输入必须恰好包含一个非空 `url` 或一个当前文稿内的一基 `slide`；`tooltip` 可省略，也可显式为空。Getter 返回 detached frozen snapshot，setter 采用 whole replacement，同值赋值为 exact no-op，`undefined` 清除 click link。内部关系按目标页 identity 保存，移动或在目标前插删页面只更新 getter ordinal；复制 self-link 会指向副本自身，删除目标页会清理相关 click/hover，shared relationship 则按引用 clone-on-write 与回收。外部链接产生 validator 的预期可移植性 warning。Hover 编辑、text-run/table/image/chart/media 链接创建、action navigation、advanced line fill/custom dash 和 percentage positions 仍待后续小项。

### 创建和编辑自定义几何路径

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();
const custom = slide.addCustomShape({
  paths: [{
    width: inches(4),
    height: inches(3),
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: inches(4), y: 0 } },
      { kind: 'lineTo', point: { x: inches(2), y: inches(3) } },
      { kind: 'close' },
    ],
  }],
}, {
  name: 'Custom triangle',
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
});

const updatedPaths = [{
  ...custom.customGeometry!.paths[0]!,
  commands: [
    { kind: 'moveTo' as const, point: { x: 0, y: 0 } },
    { kind: 'lineTo' as const, point: { x: inches(4), y: inches(3) } },
    { kind: 'lineTo' as const, point: { x: 0, y: inches(3) } },
    { kind: 'close' as const },
  ],
}];
custom.customGeometry = { ...custom.customGeometry!, paths: updatedPaths };
custom.presetType = 'triangle';

const formulaShape = slide.addCustomShape({
  adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [25_000] } }],
  guides: [{ name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } }],
  handles: [{
    kind: 'xy',
    position: { x: 'x1', y: 'vc' },
    xGuide: 'adj1',
    minX: 0,
    maxX: 'r',
  }],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'b' } },
      { kind: 'close' },
    ],
  }],
});
formulaShape.customGeometry = {
  ...formulaShape.customGeometry!,
  adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [50_000] } }],
};
const evaluated = formulaShape.evaluateCustomGeometry();
console.log(evaluated?.guides); // [{ name: 'x1', value: ... }]
```

`SlideModel.addCustomShape()` 与 `ShapeModel.customGeometry` 使用 direct OOXML values：path extent 固定为 positive safe-integer EMU；point、arc radius、arc angle 和 formula operand 使用 `CustomGeometryValue`，可为 safe integer 或单个 guide/built-in token。数值坐标和半径按 EMU，数值角度按 `1/60000°`；可用 `inches()` / `degrees()` 显式换算。路径支持 `moveTo`、`lineTo`、`arcTo`、`quadraticBezierTo`、`cubicBezierTo`、`close`，也支持多个 subpath、多个 path、empty path，以及 `fill`、`stroke`、`extrusionOk` path flags。

`CustomGeometryGuide` 将 `adjustments` / `guides` 分别映射到 `a:avLst` / `a:gdLst`；`CustomGeometryFormula` 支持全部 17 个 DrawingML operators：一元 `val/abs/sqrt`，二元 `at2/cos/max/min/sin/tan`，三元 `*/`、`+-`、`+/`、`?:`、`cat2`、`mod`、`pin`、`sat2`。Guide 名称在两个列表中全局唯一；string value 必须是非空、无 XML whitespace、XML-safe 且非十进制整数的 token。Strict codec 负责 lexical/arity/tree validation；独立 evaluator 会先按 source order 计算全部 adjustments，再按 source order 计算 guides，拒绝 cycle、forward reference 和 unknown token。已完成的 custom guide 会覆盖同名 built-in，但定义该 shadow guide 的公式仍可引用原 built-in。

`CustomGeometryXyHandle` 与 `CustomGeometryPolarHandle` 组成有序的 `CustomGeometryHandle` union，并由 `CustomGeometry.handles` 映射到 `a:ahLst` 中的 `a:ahXY` / `a:ahPolar`。两类都要求 `position`（direct `a:pos`）；XY 的 `xGuide/yGuide/minX/maxX/minY/maxY` 对应 `gdRefX/gdRefY/minX/maxX/minY/maxY`，polar 的 `radiusGuide/angleGuide/minRadius/maxRadius/minAngle/maxAngle` 对应 `gdRefR/gdRefAng/minR/maxR/minAng/maxAng`。位置、XY/radius bounds 接受 shape coordinate-space safe integer 或 token，angle bounds 的数字使用 direct `1/60000°`；每个 optional 字段都可独立出现，跨 kind 顺序原样保留，省略或 `[]` 不产生 own `handles` property。

`CustomGeometryConnectionSite` 通过有序的 `CustomGeometry.connectionSites` 映射 `a:cxnLst`；每项要求 `angle`（`a:cxn@ang`）与 `position`（direct `a:pos`）。数值 position 是 custom-geometry coordinate-space direct safe integer，数值 angle 是 direct `1/60000°`，两者也接受 guide/built-in token；codec 保留 direct token，evaluator 可将其解析成 number，但不归一化角度或判断 site 是否位于 path 上。顺序和重复项原样保留，省略或 `[]` 不产生 own `connectionSites` property。

`CustomGeometryTextRectangle` 由 `CustomGeometry.textRectangle?` 公开，其 required `left/top/right/bottom` 分别映射 `a:rect@l/t/r/b`。每个边都接受 direct safe integer 或 guide/DrawingML built-in token。省略该属性或显式传入 `{ left: 'l', top: 't', right: 'r', bottom: 'b' }` 都折叠为 canonical default，snapshot 不产生 own `textRectangle` property。LibreOffice 保存时常生成的 `textAreaLeft/Top/Right/Bottom` guides 及对应 rect token 可严格读取。

输入会立即脱离 caller，getter 返回 detached deep-frozen snapshot；setter whole-replace 整个 geometry，同值赋值是 exact bytes/journal no-op，不接受 `undefined` 清除。Strict reader 要求 namespace、属性、child order、handle/site 唯一合法 direct `position`，以及 optional text rectangle 在存在时为唯一 direct `a:rect` 且具有完整合法四边；malformed state 保留原包 bytes，但返回 `undefined` 并在任何更改前拒绝 replacement。给 `presetType` 赋值会转成 preset geometry；给 preset shape 设置 `customGeometry` 会转回 custom geometry，并保留 shape identity 与样式。

`evaluateCustomGeometry(geometry, { width, height })` 是不修改输入或 package 的 pure API；`ShapeModel.evaluateCustomGeometry()` 使用当前 live transform 的 width/height，非 custom 或 strict snapshot 不可读时返回 `undefined`。Evaluator 支持 37 个 DrawingML built-ins：`3cd4/3cd8/5cd8/7cd8/b/cd2/cd4/cd8/h/hc/hd2/hd3/hd4/hd5/hd6/hd8/l/ls/r/ss/ssd2/ssd4/ssd6/ssd8/ssd16/ssd32/t/vc/w/wd2/wd3/wd4/wd5/wd6/wd8/wd10/wd32`，并把 guide、handle、connection site、text rectangle 与全部 path command token 解析为 number。合法有限小数会保留，`-0` 规范为 `0`，`*/` 与 `+/` 除零按 DrawingML 返回 `0`，缺省 text rectangle 会实体化为 `{ left: 0, top: 0, right: width, bottom: height }`。结果与所有嵌套分支都 detached 且 recursively frozen。

公开类型包括 `CustomGeometryEvaluationContext`、`CustomGeometryEvaluationErrorCode`、`CustomGeometryEvaluationError`、`EvaluatedCustomGeometryGuide/Point/TextRectangle/Command/XyHandle/PolarHandle/Handle/ConnectionSite/Path` 与 `EvaluatedCustomGeometry`。求值失败通过 `unknown-token`、`forward-reference`、`cyclic-reference`、`invalid-domain`、`non-finite-result` 区分，并在适用时提供 `guideName` / `token`。

PptxGenJS 4.0.1 的合法 `ShapeType.custGeom` points 最终输出可导入为相同 native snapshot，包括后续 `moveTo`、arc/quadratic/cubic 与 close。其 `<100` 数字和数字字符串按 inch、`>=100` 数字按 direct value、百分比按整张 slide 计算，arc point 的 `x/y` 被忽略；native API 不复制这些启发式或 coercion，只接受显式 direct values。PptxGenJS 4.0.1 没有公开的 guide-formula、arbitrary adjustment-handle、connection-site 或 text-rectangle 输入，只生成 empty `a:cxnLst` 与 canonical default `a:rect`；formulas、handles、connection sites 与 arbitrary text rectangles 是完整 DrawingML 创建/编辑所需的 native extensions。

PptxGenJS 对等范围是其合法最终 numeric path 与 canonical default text rectangle；native formula/guide/handle/site 求值属于扩展。Evaluator 不进行 path coordinate scaling，不计算 arc endpoint 或 resolved bounds，也不实现 handle dragging、connector snapping/creation。实际 tarball 生成的 4 页 gallery 含 22 个 evaluator 目标，原件与 LibreOffice round-trip 均为 22/22 strict evaluable、PowerPoint 2010 validation 0 errors/0 warnings；LibreOffice 会把 22/22 direct expressions 改写成 numeric paths/text-area guides，其中 21/22 numeric path/text-rectangle 匹配，唯一差异是 `sqrt` 样例终点从 `600000` 改为 `0`，且它会重写全部 guide arrays，并改写一项 handle/site metadata。

### 预设形状调整值

```ts
const arc = slide.addShape('blockArc', {
  adjustments: [
    { name: 'adj1', value: 16_200_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 25_000 },
  ],
});
arc.adjustments = [{ name: 'adj1', value: 10_800_000 }];
arc.adjustments = [];
```

`ShapeAdjustment.value` 是 `a:gd@fmla="val N"` 的 direct safe integer，不执行形状专属单位换算。列表有序且名称唯一；输入会立即脱离 caller，getter 返回 detached deep-frozen snapshot，赋相同列表是 exact bytes/journal no-op。赋值采用 whole replacement，`[]` 清空 `a:avLst`，setter 不接受 `undefined`。复杂公式、重复或歧义结构读取为 `undefined`，编辑会在 package 变化前拒绝；改变 `presetType` 会重置调整值，同类型赋值保留原 bytes。

PptxGenJS 4.0.1 的合法 `rectRadius`、`angleRange` 与 `arcThicknessRatio` 最终输出可直接导入。原生 API 接受最终整数 guide 列表，因此保留显式 zero，也不会复制 PptxGenJS 的 zero truthiness 丢失、字符串转换、`rectRadius` 快捷字段优先级、无 angles 时忽略 thickness 或 malformed/unsafe passthrough。Custom geometry paths、guide formulas、handles、connection sites、text rectangle 与 numeric evaluation 已由独立 API 支持。

## 开发

```sh
pnpm install
pnpm check
pnpm build
```

Node.js 20+ 或现代浏览器；TypeScript strict mode。浏览器支持 `Blob`、`File`、Web `ReadableStream` 输入，以及 `writeBlob()` 和 `download()` 输出。

## Workspace packages

npm 用户只需安装 `@jiayunxie/pptx`。以下是仓库内部模块边界，不需要分别安装：

- `@pptx/sdk`：统一的打开、编辑、验证和保存 API。
- `@pptx/opc`：ZIP、content types、relationships 和 package graph。
- `@pptx/lossless-xml`：source-span XML tree 与最小 patch。
- `@pptx/model`：slide 和常规对象语义模型。
- `@pptx/codecs`：Master/Layout/Theme、Gradient/Transparency、Media。
- `@pptx/pptxgenjs-adapter`：PptxGenJS 公开输出适配。
- `@pptx/testkit`：part hash diff、mutation isolation 和 LibreOffice helper。
- `@pptx/cli`：`pptx-inspect` 离线 inspection CLI。
- `@pptx/plugin-transitions`：转场、自动换页与声音。
- `@pptx/plugin-animations`：动画/媒体 timing tree。
- `@pptx/plugin-advanced-charts`：组合/现代图表及高级 series 功能。
- `@pptx/plugin-smartart`：SmartArt part set 与节点编辑。

## CLI

```sh
npx @jiayunxie/pptx@next --json doctor
pptx-inspect --json package validate deck.pptx
pptx-inspect --json slides list deck.pptx
```

CLI 写操作要求明确输出路径，并支持 `--dry-run`。完整命令见 [packages/cli/README.md](./packages/cli/README.md)。
