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
  calculateRasterImageSizing,
  degrees,
  inches,
  inspectRasterImage,
  PptxDocument,
  type AddImageSourceOptions,
  type RasterImageSizing,
} from '@jiayunxie/pptx';

const sizing: RasterImageSizing = {
  type: 'cover',
  width: inches(5),
  height: inches(3),
};
const options: AddImageSourceOptions = {
  name: 'Quarterly chart',
  altText: 'Revenue by quarter',
  x: inches(1),
  y: inches(1.5),
  sizing,
  rotation: degrees(10),
};

const document = PptxDocument.create();
document.addSlide();
const image = await document.addImage(0, 'chart.png', options);
const info = inspectRasterImage(new Uint8Array(await readFile('chart.png')));
console.log(info); // { contentType: 'image/png', width, height }
console.log(calculateRasterImageSizing(info, sizing));
image.setTransform({ x: inches(1.5) });
image.sourceRectangle = { left: 10, top: -5, right: 5, bottom: 0 };
image.sourceRectangle = undefined;
image.replaceData(new Uint8Array(await readFile('chart-updated.png')), 'image/png');
await document.writeFile('images.pptx');
```

`PptxDocument.addImage()` 接受 Node 本地路径、HTTP/HTTPS URL、浏览器相对 URL、strict base64 data URI、`Uint8Array`、`ArrayBuffer`、`Blob`/`File`、Web `ReadableStream` 和 async byte iterable。格式只以 bytes signature 为准，支持 PNG/JPEG/GIF；`AddImageSourceOptions.contentType` 是可选 assertion，不一致会在 package mutation 前拒绝。`signal` 可中止文件、Fetch、Blob 和 stream 加载。文件扩展名、`Blob.type`、HTTP `Content-Type` 和 `File.name` 不参与格式判断。

`inspectRasterImage()` 返回 signature 检测出的 canonical content type 与 raw pixel width/height；PNG 读取 direct IHDR，GIF 读取 logical screen，JPEG 安全遍历所有尺寸型 SOF marker。`calculateRasterImageSizing()` 是不读取 package 的纯计算器，接受 EMU target frame，并为 `contain`、`cover` 或 source-pixel `crop` 返回同尺寸 transform 与四边 `sourceRectangle`。高层 `sizing` 与 top-level `width`/`height` 互斥；省略 `sizing` 时仍保留 PptxGenJS 的 1-inch 默认 transform，不会自动把 intrinsic pixels 转为布局尺寸。

底层同步 `SlideModel.addImage(bytes, { contentType, sourceRectangle, ... })` 用于调用方已持有 bytes 的严格原子创建。`ImageSourceRectangle` 的 `left/top/right/bottom` 单位为百分比，`1` 表示 1%，精度为 0.001%；负值用于 contain 扩展可见源区域。`ImageModel.sourceRectangle` 返回 detached frozen snapshot，可 whole-replace，赋 `undefined` 清除 direct `a:srcRect`。高层 `PptxDocument.addImage()` 不接受 direct `sourceRectangle`，调用方应使用 `sizing`；options/sizing 会在任何异步 source I/O 前脱离 caller，计算也在 package mutation 前完成。

每次创建原子地拥有一个唯一 media part、一个 internal image relationship 和一个 canonical rectangular picture；任何验证或写入失败都会回滚 package、关系、slide XML 与 mutation journal。复制页面先共享图片 part，`ImageModel.replaceData()` 对独占 target 原位更新，对共享 target clone-on-write；六种 presentation format 都可创建、写出和重开。

锁定 PptxGenJS 4.0.1 的公开 path/data PNG/JPEG/GIF loader 语义，并对 contain/cover/equal-ratio/crop 的 6 个 sizing case 精确匹配最终 transform 与 direct `srcRect`。实际 npm tarball 的 Node/browser/types/CLI smoke 通过，连续两次构建的 38 个 dist 文件 SHA-256 完全一致。4 页、40 shapes、12 图片 sizing gallery 的原件和 LibreOffice 回存件均可严格重开，PowerPoint 2010 validation 为 0 errors / 0 warnings，overflow 为 0，并已逐页检查。

LibreOffice 保留 12/12 图片的 payload SHA、content type、名称、非空 alt text、顺序和 internal relationship；它把 12 个重复 payload target 去重为 3 个并重写全部 picture markup。最大 transform 量化为 360 EMU，最大 `srcRect` 量化为 0.007%，双翻转被等价规范化为旋转增加 180°。原件 180 DPI 输出为 2400×1350；回存件把页面宽度规范化后直接输出为 2401×1350，逐页检查使用等比例 2400×1350 raster。

## 添加和编辑 SVG 图片

```ts
import { readFile } from 'node:fs/promises';
import {
  calculateImageSizing,
  inches,
  inspectSvgImage,
  PptxDocument,
} from '@jiayunxie/pptx';

const svgBytes = new Uint8Array(await readFile('architecture.svg'));
const fallbackPng = new Uint8Array(await readFile('architecture.png'));
const svgInfo = inspectSvgImage(svgBytes);
const sizing = calculateImageSizing(svgInfo, {
  type: 'contain',
  width: inches(6),
  height: inches(4),
});

const document = PptxDocument.create();
const slide = document.addSlide();
const image = await document.addImage(0, svgBytes, {
  contentType: 'image/svg+xml', // 可选的 MIME assertion
  fallback: fallbackPng,
  name: 'Architecture',
  altText: 'System architecture diagram',
  sizing: { type: 'contain', width: inches(6), height: inches(4) },
});
console.log(image.isSvg, image.sourcePartUri, image.fallbackPartUri, image.svgPartUri);
console.log(sizing.sourceRectangle);

// 已持有严格 bytes 时也可使用同步底层 API。
const strictImage = slide.addSvgImage(svgBytes, fallbackPng, {
  x: inches(6.5),
  width: inches(4),
  height: inches(3),
});
strictImage.replaceSvgData(
  new Uint8Array(await readFile('architecture-updated.svg')),
  new Uint8Array(await readFile('architecture-updated.png')),
);
await document.writeFile('svg-images.pptx');
```

高层 `PptxDocument.addImage()` 对 SVG 支持与 raster 相同的 path、URL、data URI、bytes、Blob/File、Web stream 和 async iterable 来源；`inspectImage()` / `inspectSvgImage()` 返回 canonical `image/svg+xml` 与 intrinsic dimensions，通用 `calculateImageSizing()` 支持 contain/cover/crop。`contentType: 'image/svg+xml'` 只是可选 assertion，真实类型仍由严格 SVG XML 检查决定。

每张 SVG 图片始终包含一个 SVG part 和一个 PNG fallback part。高层 API 会验证 fallback 的真实 PNG signature，优先级是调用方显式 PNG、浏览器 Canvas 自动栅格化、最后是内建透明 PNG；同步底层 `addSvgImage()` 只接受并复制 non-empty bytes，由调用方保证 payload 符合声明。需要 fallback-only 客户端保持完整视觉时，应显式提供高质量 PNG。库不执行 SVG script、不抓取 SVG 外部资源，也不承诺任意 SVG 的跨客户端栅格化一致性。

`ImageModel.isSvg` 只在 namespace/relationship/part 均安全且无歧义时为 true；此时 `sourcePartUri === fallbackPartUri`，`svgPartUri` 指向矢量 part。`replaceData()` 会拒绝 SVG，必须用 `replaceSvgData(svgBytes, fallbackPngBytes)` 原子替换一对 payload。复制页面最初共享两类 target；替换共享图片时两侧同时 clone-on-write。任何 malformed paired state 或替换失败都保持 package、关系、XML、identity 与 mutation journal 不变。

PptxGenJS 4.0.1 的 data-contain、path-cover、data-crop 3 个公开 case 已精确匹配 picture 结构、transform、direct `srcRect`、extension URI/namespace、关系角色和 SVG payload。Native 明确修复了 PptxGenJS path SVG 把 SVG bytes 写进 `.png` fallback part 的缺陷，始终写入可验证的 PNG。

实际 tarball 的 Node/browser/declaration/CLI smoke 通过，两次 clean build 的 38 个 dist 文件 SHA-256 完全一致。5 页 gallery 含 13 个 shapes、8 张 SVG picture、7 个 SVG parts、7 个 PNG fallbacks 和 16 条 image relationships；原件和 LibreOffice 回存件均 strict reopen，PowerPoint 2010 validation 为 0 errors / 0 warnings。LibreOffice 保留 shape order、名称、alt text、SVG hashes、关系角色和 7+7 targets，将 MIME 从 `image/svg+xml` 规范化为 `image/svg`；最大 position/size 量化 360 EMU、最大 `srcRect` 量化 0.003%，flip/rotation 采用等价规范化。它回存后可能选择 PNG fallback 渲染，所以旧客户端视觉仍以 fallback 质量为准。

当前仍未提供 external SVG relationship、SVG DOM 局部编辑、rounding/transparency、alt-text 编辑、图片 hyperlink/shadow/placeholder，以及单图片删除与 media GC。嵌入媒体创建能力见下一节。

## 创建嵌入式音频与视频

```ts
import { readFile } from 'node:fs/promises';
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
document.addSlide();
const poster =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAACXBIWXMAAAABAAAAAQBPJcTWAAAAEElEQVR4nGP8ywACLGCSAQANEQED1LYyQAAAAABJRU5ErkJggg==';

const audio = await document.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
  name: 'Opening narration',
  altText: 'Opening narration audio',
  poster,
  x: inches(1),
  y: inches(1),
  width: inches(3),
  height: inches(2),
  play: 'auto',
  loop: true,
  hideWhenStopped: true,
  volume: 0.5,
});

const videoBytes = new Uint8Array(await readFile('overview.mp4'));
await document.addVideo(0, videoBytes, {
  contentType: 'video/mp4',
  fileName: 'overview.mp4',
  poster,
});

audio.name = 'Opening narration edited';
audio.altText = undefined;
audio.settings = { play: 'click', loop: false, volume: 0.75 };
audio.setTransform({ x: inches(2), y: inches(1.5) });
await audio.replaceSource('https://example.com/narration.wav');
await audio.replaceSource(new Uint8Array(await readFile('narration.wav')), {
  contentType: 'audio/wav',
  fileName: 'narration.wav',
});
await audio.replacePoster(new Uint8Array(await readFile('poster.gif')), {
  contentType: 'image/gif',
});
await audio.replacePoster(); // reset to the built-in PNG
await document.writeFile('media.pptx');
```

`PptxDocument.addAudio()` 与 `addVideo()` 接受 Node path、strict base64 data URI、`Uint8Array`、`ArrayBuffer`、`Blob`/`File`、Web `ReadableStream` 和 async byte iterable。音频支持 `audio/mpeg` (`.mp3`)、`audio/mp4` (`.m4a`)、`audio/wav` (`.wav`) 与 `audio/ogg` (`.ogg`)；视频支持 `video/mp4` (`.mp4`/`.m4v`)、`video/quicktime` (`.mov`) 与 `video/webm` (`.webm`)；海报支持 `image/png`、`image/jpeg` 与 `image/gif`。省略海报时使用内建 PNG。HTTP/HTTPS 媒体保持 external relationship，库不会下载；HTTP/HTTPS 海报会被拒绝。

创建返回 live `MediaModel`；同一会话内 `document.media(slideIndex)`、`slide.media` 与 `slide.shapes` 返回同一个对象。`name`、`altText`、`settings` 和继承的 transform 可原位编辑；`replaceSource()` 在不改变 audio/video kind 与对象 identity 的前提下支持 embedded↔external，`replacePoster()` 支持 PNG/JPEG/GIF，省略 source 会重置为内建 PNG。`replaceSource()` 只接受 `contentType`、`fileName`、`transcode`，`replacePoster()` 只接受 `contentType`、`fileName`；两者返回原对象。`media.remove()` 与 `slide.deleteMedia(shapeId)` 删除对象。

MIME/扩展名解析优先级是显式 `contentType` assertion → data URI 声明 → `fileName` 或 path/`File.name` 的已知扩展名 → audio/video 默认值。Assertion 与 data URI MIME 不一致、已知扩展名与最终 MIME 不一致都会在 package mutation 前拒绝；未知扩展名不作为格式事实，输出使用所选 MIME 的 canonical extension。Data URI 必须包含受支持 MIME 和标准、完整 padding、canonical padding bits 的 base64，不能含空白、URL-safe alphabet 或 percent encoding。`Blob.type` 不参与判断。

Options 和内存 bytes 会在异步读取前脱离 caller；path、Blob、stream、可选 `transcode` 与海报全部解析后，才进入一个同步 package transaction。创建、source/poster 替换与删除的任何验证、I/O、transcode、hash、relationship、part 或 XML 失败都会保持 parts、content types、关系、slide XML、ZIP state、对象 identity 与 mutation journal 不变。相同 SHA-256 且 MIME 完全相同的 media/poster payload 会复用同一 `/ppt/media` part；duplicate 初始共享 target，首次不同写入 clone-on-write，只 retarget 被编辑的 picture。对象或幻灯片删除只 GC package graph 中无 incoming reference 的 target。

创建会写 canonical `a:audioFile` / `a:videoFile`、kind relationship、Microsoft media relationship、poster image relationship、media click action 与矩形海报 picture。`name`、`altText`、EMU transform 和 playback preferences 均受严格验证；使用 `inches()` 把布局尺寸转为 EMU。`play`、`loop`、`hideWhenStopped` 与 `volume` 当前保存在库的私有 playback extension 中，尚未自动生成 PowerPoint native timing tree。

PptxGenJS 4.0.1 的 4/4 个公开有效 data/path、audio/video、cover、`extn`、`objectName` 与 transform 用例已达到最终语义对等。Reader 兼容其 audio `a:videoFile`、`audio/mp3` 与 duplicate-audio relationship 缺陷；只读、metadata、settings 与 transform 编辑不会重写这些 legacy primary roles，`replaceSource()` 才把当前 picture 的 primary roles canonicalize 为 native audio/video 结构。Native 创建始终使用 `a:audioFile`、canonical `audio/mpeg` 和标准 audio relationship。

实际 npm tarball 的 Node、browser、declaration 与 CLI smoke 全部通过，覆盖 live identity、全部编辑面、embedded↔external、poster replacement/reset、duplicate COW、对象/幻灯片删除、GC 与 reopen；连续两次 clean build 的 44 个 dist 文件 SHA-256 manifest 完全一致。4 页全格式 lifecycle gallery 包含 6 个音频、4 个视频、30 条媒体角色关系、7 个唯一媒体载荷、4 个 poster payload 和 11 个 `/ppt/media` parts，零孤儿；原件 strict reopen、180 DPI 渲染、overflow 与逐页视觉检查通过。全格式文件的 PowerPoint 2010 profile 为 0 errors，只有 OGG/WebM 两条预期 warning；排除这两种格式的 8-object 可移植子集为 0 errors / 0 warnings。External 对照文件产生 4 条预期 portability warnings。

LibreOffice 26.8 当前会在 save/reopen 时保留 4 页顺序与 wide 画布，但删除全部 10 个媒体对象、30 条媒体角色关系和 11 个媒体/海报 parts；回存四页为空白，仍可 strict reopen、0 errors / 0 warnings、零 overflow。这是已记录的客户端降级，不是 native 写出或 round-trip 保留承诺。

下一媒体小项是 native PowerPoint timing tree。仍未支持 online video、remote-fetch embedding、captions/subtitles、crop/rounding/shadow/hyperlink/placeholder styles、内建转码引擎与更广泛 PowerPoint/Keynote/Google Slides 认证；因此整体 PptxGenJS 全功能对等路线尚未完成。

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
