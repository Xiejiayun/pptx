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

当前仍未提供 external SVG relationship、SVG DOM 局部编辑、rounding/transparency、alt-text 编辑、图片 hyperlink/shadow 与高级 placeholder 样式，以及单图片删除与 media GC。嵌入媒体创建能力见下一节。

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

创建会写 canonical `a:audioFile` / `a:videoFile`、kind relationship、Microsoft media relationship、poster image relationship、media click action、矩形海报 picture，以及可直接被 PowerPoint 执行的原生 `p:timing` 播放树。`play: 'click' | 'auto'`、`loop`、`hideWhenStopped` 与 `volume: 0..1` 同时写入精确偏好/ownership `px:playback` 和 native `cMediaNode`/play/pause command；无需安装动画插件。`settings` whole-replace 会原子同步两处状态，`settings = undefined` 只清除该媒体的私有偏好和库拥有的原生播放 graph，不改普通动画或未知 timing 分支。

读取顺序固定为：合法 `px:playback` 精确偏好 → 没有私有偏好时严格识别唯一、直接、完整的 native media graph → 空设置。Native-only PowerPoint 文件因此可读取并在首次设置时安全 adoption；已有 ownership 过期时会重建当前媒体 graph。库只认领唯一匹配的 direct graph，unsupported/ambiguous timing 保持原字节并拒绝危险编辑。创建、设置、清除、复制、删除与 timing ID 分配都在同一 OPC transaction 内；普通动画和媒体共用全页 ID 空间。

`MEDIA_TIMING_MISSING`、`MEDIA_TIMING_STALE`、`MEDIA_TIMING_UNSUPPORTED`、`MEDIA_TIMING_AMBIGUOUS`、`MEDIA_TIMING_DANGLING_TARGET` 与 `MEDIA_TIMING_KIND_MISMATCH` 会报告缺失、陈旧、不支持、歧义或悬空状态。动画插件仍兼容旧版只含 `px:playback` 的文件，并通过共享 codec 幂等补写 native graph；健康 native timing 不会被重复生成。

PptxGenJS 4.0.1 的 4/4 个公开有效 data/path、audio/video、cover、`extn`、`objectName` 与 transform 用例已达到最终语义对等。Reader 兼容其 audio `a:videoFile`、`audio/mp3` 与 duplicate-audio relationship 缺陷；只读、metadata、settings 与 transform 编辑不会重写这些 legacy primary roles，`replaceSource()` 才把当前 picture 的 primary roles canonicalize 为 native audio/video 结构。Native 创建始终使用 `a:audioFile`、canonical `audio/mpeg` 和标准 audio relationship。

实际 45-file npm tarball 的 Node、Chrome browser、declaration 与 CLI smoke 全部通过，`nativeMediaTiming: true`，覆盖 clear、duplicate isolation、delete、唯一 ID、目标隔离、诊断和 reopen；连续两次 clean build 的 42 个 dist 文件 SHA-256 manifest 完全一致。六种 presentation format 均通过 native timing 创建/编辑/复制/删除/重开。9 页真实媒体 gallery 含 12 个媒体对象、7 种媒体 MIME、3 种 poster MIME、10 个去重后的 `/ppt/media` parts 和零孤儿；原件 strict reopen、180 DPI 渲染、overflow 与逐页视觉检查通过。PowerPoint 2010 profile 为 0 errors，仅 OGG/WebM 两条预期 warning。

LibreOffice 26.8 当前会在 save/reopen 时保留 9 页顺序与文案，但删除全部媒体、poster、media relationships 和 timing；回存件仍可 strict reopen 且为 0 errors / 0 warnings。这是已记录的客户端降级，不是 native 写出或 round-trip 保留承诺。本机 PowerPoint 16.112 自动打开对 gallery、LibreOffice 回存件与最小控制文件都返回同一 `-9074`，因此没有把该环境的 PowerPoint 往返误记为通过。

媒体 timing 的下一层仍未支持 trim/bookmarks、有限重复、narration/cross-slide audio、captions/subtitles、online video、remote-fetch embedding、crop/rounding/shadow/hyperlink 与高级 placeholder 样式、内建转码引擎与更广泛 PowerPoint/Keynote/Google Slides 认证。

## 创建和语义编辑原生图表

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create({ slideSize: 'wide' });
const slide = document.addSlide();
const revenue = [{
  name: 'Revenue',
  categories: ['Q1', 'Q2', 'Q3'],
  values: [100, 130, 160],
}];

const chart = await slide.addChart('bar', revenue, {
  name: 'Quarterly revenue',
  altText: 'Revenue by quarter',
  x: inches(1),
  y: inches(1),
  width: inches(8),
  height: inches(4.5),
});
await chart.replaceDefinition({
  groups: [{
    type: 'bar',
    series: revenue,
    options: {
      grouping: 'clustered',
      dataLabels: { showValue: true, position: 'outsideEnd' },
    },
  }],
  options: {
    title: { text: 'Quarterly revenue' },
    legend: { position: 'bottom' },
    valueAxis: { minimum: 0, numberFormat: '#,##0' },
  },
});

await slide.addChart([
  { type: 'bar', series: revenue },
  {
    type: 'line',
    axis: 'secondary',
    series: [{ name: 'Margin', categories: ['Q1', 'Q2', 'Q3'], values: [24, 28, 31] }],
  },
], { x: inches(1), y: inches(1), width: inches(8), height: inches(4.5) });

await chart.replaceSeries([{ ...revenue[0], values: [105, 136, 172] }]);
console.log(chart.definition, await chart.diagnostics());
await document.writeFile('native-charts.pptx');
```

`CHART_TYPES` 覆盖 `area`、`bar`、`bar3D`、`bubble`、`doughnut`、`line`、`pie`、`radar`、`scatter`。分类图使用 `categories`/`values`，scatter 使用 `xValues`/`values`，bubble 另带正数 `sizes`；bar/area/line 可组成主轴/次轴组合图。每个原生图表都同步创建内嵌 XLSX、A1 formulas 和 display caches，workbook bytes 在同步 OPC transaction 前生成。

`ChartModel.definition` 是 detached frozen 语义快照；`replaceDefinition()` 可替换类型、组合、数据和受支持选项，`replaceSeries()` 更新单组图表数据，`remove()` / `slide.deleteChart()` 删除并按引用回收 chart/workbook/style/color 子图。导入的共享目标在首次编辑时采用 clone-on-write；`setXml()` 保留为显式 raw escape hatch。标题、图例、chart/plot area、主/次轴、gridlines、labels、data table、series fill/line/marker、颜色以及各类型 grouping/gap/hole/angle/radar/scatter/bubble/3D 选项均有 strict read/create/edit。Diagnostics 区分关系、结构、cache、axis、workbook 缺失/分歧和 modern chart。

PptxGenJS 4.0.1 的九种公开图表和 bar+line 主/次轴组合已通过真实输出导入、编辑、公式/cache/XLSX/relationship 和选项对照。实际 npm tarball 的 Node、real-Chrome、declaration 与 installed CLI smoke 均报告 `nativeCharts: true`；11 页 gallery 包含 10 个 chart parts、10 个 XLSX、零孤儿，PowerPoint 2010 profile 为 0 errors / 0 warnings，180 DPI overflow 为 0，并已逐页检查。

LibreOffice 26.8 能显示八种 2D 图表及组合图；`bar3D` 在 native 与独立 PptxGenJS 控制文件中都只显示标题。保存时 LibreOffice 保留全部 10 个图表的类型和 cache 数据，但移除内嵌 workbook 并把公式改成客户端占位符；reader 将其识别为可编辑的 `cache-only` 状态并报告 `CHART_WORKBOOK_MISSING` warning，首次语义替换会重新生成同步 XLSX。本机 PowerPoint 16.112 对 gallery 与独立控制文件自动打开均返回同一 `-9074`，本轮不声明 PowerPoint 往返通过。

仍未支持 Office 2016 `cx:*` modern chart 创建/语义编辑、external workbook 编辑、chart animations、内建趋势线/error bar 创建以及更广泛 Keynote/Google Slides 认证；高级插件继续处理 modern inspection、trendline、error bar 与显式 fallback。Slide background、slide number、default text color 与 master/layout/placeholder 均已完成；PptxGenJS 全功能对等的下一项是 advanced text。

## 创建和编辑页面背景

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();

slide.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
};

await document.setSlideBackgroundImage(0, './background.png');
slide.background = { kind: 'none' }; // 显式写入合法 a:noFill
slide.background = undefined;        // 删除 direct p:bg，恢复 layout/master 继承
```

`SlideModel.background` 支持 direct `p:cSld/p:bg/p:bgPr` 的 `none`、sRGB/theme solid、linear/path gradient 和 PNG/JPEG/GIF image。读取严格且不修复原包；返回的颜色、stops、rectangles 与图片 bytes 都会脱离 caller。相同值赋值是 exact no-op。图片背景拥有 internal image relationship；duplicate 初始共享 target，首次不同写入 clone-on-write，替换、清除和删除页面只回收 package graph 中已无 incoming reference 的媒体部件。

高层 `setSlideBackgroundImage()` 接受与 raster loader 相同的 Node path、HTTP/HTTPS 或 browser-relative URL、strict data URI、bytes/ArrayBuffer、Blob/File、Web stream 和 async iterable，并在 source、signature 与可选 MIME assertion 全部通过后才进入同步 transaction。`undefined` 与 `{ kind: 'none' }` 不等价：前者恢复继承，后者保留显式无填充意图。Layout/master 现在也可通过语义 wrapper 读写相同的 direct background；仍未提供 `p:bgRef`、pattern/group fill 与图片 crop/tile/effects，这些状态在无关编辑中保持原字节。

PptxGenJS 4.0.1 的合法 solid、transparency 与 PNG background 最终结构已对等。它的 `{ type: 'none' }` 实际不写 direct background，而 `{ type: 'none', color }` 会产生空 `p:bgPr`；native 不复制这个缺陷，显式 none 始终写合法 `a:noFill`。实际 npm tarball 的 Node、real-Chrome、declaration 与 installed CLI smoke 均报告 `slideBackgrounds: true`。两次 clean build 的 48 个 dist 文件 SHA-256 manifest 完全一致；11 页 native gallery 含 41 parts、39 relationships 和 3 个背景媒体部件，PowerPoint 2010 profile 为 0 errors / 0 warnings，全部 11 页与 7 页 PptxGenJS 对照均已逐页渲染检查且 overflow 为 0。

LibreOffice 26.8 回存后保留 11 页顺序、solid/gradient/image 类型和全部图片 payload hash，但把 explicit noFill 规范化为继承、把两种 gradient 的 `rotateWithShape` 改为 false，并为 path gradient 写入全幅 fill rectangle；回存件仍为 0 errors / 0 warnings。本机 PowerPoint 16.112 自动打开返回 `-9074`，因此本轮不声明 PowerPoint 往返通过。

## 创建、编辑和同步页码

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create({ firstSlideNumber: 5 });
const slide = document.addSlide();

slide.slideNumber = {
  x: inches(8.1),
  y: inches(5),
  width: inches(1.4),
  height: inches(0.35),
  align: 'center',
  rtl: true,
  valign: 'middle',
  margin: [1, 2, 3, 4],
  style: {
    fontFamily: 'Aptos',
    fontSize: 18,
    lang: 'zh-CN',
    bold: true,
    italic: true,
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
};

document.layouts[0].slideNumber = { x: inches(0.5), align: 'left' };
document.masters[0].slideNumber = { x: inches(4.5), align: 'center' };
const duplicate = document.duplicateSlide(0);
document.moveSlide(document.slides.indexOf(duplicate), 0);
document.firstSlideNumber = 10; // 同步所有安全识别的 direct slide cache
slide.slideNumber = undefined; // 只清除该页的 direct 页码
```

`SlideModel.slideNumber`、`LayoutModel.slideNumber` 与 `MasterModel.slideNumber` 分别只拥有对应 part 的 direct `p:ph type="sldNum"` 字段；master setter 还同步 direct `p:hf@sldNum`。直接设置 slide 不会暗写唯一 master 或默认 layout。位置和尺寸使用 EMU，margin/font size 使用 point，transparency 使用百分比。Getter 返回 detached、deep-frozen 的 `SlideNumber`；相同值赋值和 absent clear 是 exact no-op。严格 reader 对 wrong namespace、重复 owner、普通 field、无效 style/geometry、shape-id collision 或歧义结构返回 `undefined`，危险 setter 在 mutation 前拒绝。

`CreatePresentationOptions.firstSlideNumber` / `document.firstSlideNumber` 接受 signed Int32 safe integer；`undefined` 删除 `firstSlideNum` 并恢复 OOXML 默认 1。直接 slide cache 始终为起始值加当前零基索引，layout/master cache 为 `‹#›`；修改起始值以及 duplicate/move/delete 都在同一 transaction 内同步安全识别的 cache。Diagnostics 报告 fixed-id collision、disabled master 和 noncanonical cache。PptxGenJS 4.0.1 的公开 slide-number variants 可严格导入；native 不复制它固定 shape id 25、layout/master 随机 cache、disabled master 或 zero-size fallback 等缺陷。

实际 54-file tarball（51 个 `dist` 文件）的 Node、real-Chrome、browser conditional export、declaration 与 installed CLI smoke 均报告 `slideNumbers: true`；两次 package build 的 51-file manifest 完全一致，SHA-256 为 `3d77e6f56b8f299f2d580112fd0ebe77d0a98c38c07764259a3735064d5f9bea`。Focused suite 为 448/448；全量 Vitest 为 1194 passed、1 performance 默认 skipped，独立 performance 1/1。16 页 native gallery 为 48 parts、45 relationships、PowerPoint 2010 profile 0/0；16 页 PptxGenJS control 为 82 parts、95 relationships、0 errors 和 4 条已锁定 warning。32 页均以 180 DPI 渲染、逐页检查，最小非空像素边距分别为 50px/81px。

LibreOffice 26.8 渲染全部 direct slide 页码，但按自身页序显示 1..16；保存后保留 16 页标题顺序、15 个 direct owner、clear 状态、field type、对齐和主要显式样式，同时移除 `firstSlideNum`、重写 cache/default font/language 与 layout/master placeholder。需要跨 LibreOffice 可见性时应优先创建 direct slide field，而不是只依赖 layout/master placeholder。回存件可重开且为 0 errors、15 条 cache-normalization warning。本机 PowerPoint 16.112 自动化启动了应用但未形成 active presentation、PDF 或回存 PPTX，因此不声明 PowerPoint 往返通过。

## 设置页面默认文字颜色

```ts
slide.color = { kind: 'scheme', value: 'accent1' };
slide.addText('Uses accent1');
slide.addRichText([{
  runs: [
    { text: 'Inherited' },
    { text: ' override', style: { color: { kind: 'srgb', value: '00AA00' } } },
  ],
}]);
slide.color = undefined;
```

`SlideModel.color` 是 strict、transient 的 default text color，支持六位 sRGB 与 DrawingML theme token。它只影响设置之后新建的 `addText()` / `addRichText()` run，显式 run color 优先；改变或清除 default 不会重染已有 shape 或 table，table、master、layout 与 placeholder 也不继承该状态。Getter 返回 normalized、detached、frozen 快照；duplicate 复制当前 default，move 保留，delete 清理，且 sibling slide 彼此隔离。

OOXML 没有合法的 direct slide-level default text color 字段，因此 transient default 本身不序列化；写出时颜色已物化到每个 run 的标准 `a:solidFill`。重开后文字颜色与显示保持，但 `slide.color === undefined`。该语义覆盖 `pptx/pptm/potx/potm/ppsx/ppsm`，并通过 PptxGenJS 4.0.1 合法输出、Node、真实 Chrome、TypeScript declarations 与 installed CLI 验证。

Default Color 定向验证为 10 passed / 409 skipped，全量 Vitest 为 1205 passed / 1 skipped，独立 performance gate 为 1/1（998ms），typecheck 与 build 通过。Actual tarball 含 54 个文件、51 个 `dist` 文件，dist manifest SHA-256 为 `467d87ffea6994355c357dbad3b1ea18afa8538b1bacb85b6de43de90ad16829`，tarball SHA-256 为 `6812000a83247fdf2d63eddf81ec6ffb43c721d478e4cdcbbf4c4a3ce2b65ad1`。Native/PptxGenJS gallery 分别为 11/9 slides、38/52 parts、35/58 relationships，PowerPoint 2010 profile 均为 0 errors / 0 warnings；20 页 180-DPI 渲染全部逐页检查，overflow 为 0，最小边距均为 106px。LibreOffice 26.8 回存保留页序、文字顺序、自定义 sRGB/theme/override/40% transparency，仅将 native `tx1` 规范化为等价 `dk1`；回存件仍为 0/0。本机 PowerPoint 16.112 对 native/control 都返回 `-9074`，未加载 presentation 也未产生输出，因此不声明 PowerPoint 往返通过。

## 创建、选择和编辑 master/layout/placeholder

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create({ slideSize: 'wide' });
const layout = await document.defineSlideMaster({
  title: 'BRAND',
  background: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
  },
  margin: [inches(0.5), inches(0.5), inches(0.5), inches(0.5)],
  objects: [{
    kind: 'placeholder',
    text: 'Presentation title',
    options: {
      name: 'title_box',
      type: 'title',
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
    },
  }],
});

const slide = document.addSlide({ masterName: layout.name });
slide.addText('Quarterly results', { placeholder: 'title_box' });
await document.writeFile('branded.pptx');
```

PptxGenJS 的 `masterName` 拼写在这里保持兼容，但它严格选择的是唯一 `SlideLayoutModel.name`。`defineSlideMaster()` 在真实 parent master 下创建 named layout，支持 direct background、transient margin、slide number，以及按顺序的 rect、line、plain/rich text、placeholder、image 和 chart objects。普通 objects 留在 layout 中由 slide 继承；placeholder prompt 也留在 layout，新页仅物化可填充的 owner。

`PLACEHOLDER_TYPES` 固定为 `title` / `body` / `pic` / `chart` / `tbl` / `media`。创建方法中的 `placeholder` 可按 layout 内唯一 name，或 `{ type, index }` identity 定位；text/rich text、shape、image/SVG、chart、table 和 audio/video 会在原 owner 的继承 geometry 上填充对应 domain。重复名称、重复 identity、错误 domain、未知 `masterName` 和已填充 owner 都会在危险 mutation 前拒绝。

`document.masters` / `layouts` 返回 stable live `SlideMasterModel` / `SlideLayoutModel`，可读写 background、slide number，列出 shapes/placeholders，并添加 placeholder、text/rich text、shape、raster/SVG image 和 chart。`replaceSlideMaster()` 原子整体替换 layout definition，保留 layout part URI、wrapper identity 和 incoming slide relationships；`deleteSlideMaster()` 在 layout 被使用时要求同文档 replacement，并先重定向 slides 再安全回收 owned dependencies。Background、content 与 placeholder relationships 会持久化；layout `margin` 是供 table auto-page 与后续 `tableToSlides` 使用的 runtime state，重开后为 `undefined`。

定向 master/layout/placeholder suite 为 45 passed / 434 skipped；全量 Vitest 为 1256 passed / 1 skipped，独立 performance gate 为 1/1（578ms），typecheck、build 与 package build 通过。Actual tarball 含 57 个文件、54 个 `dist` 文件；两次 clean build 的 sorted dist-hash manifest 与 tarball 逐字节相同，SHA-256 分别为 `0a8e958ccde379ae071a7388dc4c29278ac5033a8641976324fcd5820339ad27` 和 `8362a3af38a4a7e8316a7e49e8cb3f4fb405753bd20cc935db609441819ca5e8`。Packed Node、TypeScript、CLI 与真实 Chrome 均报告 `masterLayouts: true`；Chrome 精确重开六类 placeholder、master/layout background、关系 target、payload hash 和 chart definition，validation/console/page/network errors 均为 0。

2 页 native gallery 含 32 parts / 29 relationships / 2 layouts / 1 master，PowerPoint 2010 profile 为 0 errors / 0 warnings；2 页 PptxGenJS control 为 36 parts / 34 relationships。原件与 LibreOffice 回存件共 8 页均以 2400×1350、180 DPI 渲染并逐页检查；全幅背景使 minimum non-white margin 按预期为 0px。Fixture 的 background 和 image 是 1×1 黑色 PNG，native 第二页按测试意图重定向空白 default layout，黑/空白输出不代表丢失继承。LibreOffice 26.8 保留两页、两个 layouts 和一个 master，但会改写 placeholder identity/slide-number cache，并移除 audio 与内嵌 chart workbooks；这是降级记录，不声明完整 round-trip。PowerPoint 16.112 对 native 与 control 均返回 `-9074` 且没有产生 PPTX/PDF，因此不声明 PowerPoint 往返通过。

尚未实现完整 theme text cascade、percentage coordinates、高级 text/table/media/chart 样式和更广泛客户端认证。Advanced text 已完成文本框 direct fill、simple line、begin/end arrows、simple shadow、outer hyperlink、per-run rich-text hyperlink、preset geometry、`roundRect` 绝对圆角半径、direct `isTextBox` 状态与 rich-text `breakLine` 段落拆分。

## 创建和编辑文本框填充

```ts
const plain = slide.addText('Solid text box', {
  fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: 'D9EAF7' },
    transparency: 25,
  },
});
const rich = slide.addRichText([{
  runs: [{ text: 'Theme-filled rich text' }],
}], {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  },
});
const placeholder = slide.addPlaceholder('Filled placeholder', {
  name: 'filled_title',
  type: 'title',
  fill: { kind: 'none' },
});

plain.fill = { kind: 'solid', color: { kind: 'scheme', value: 'accent3' } };
plain.fill = { kind: 'none' };
plain.fill = undefined; // 只清除 direct fill choice
```

`AddTextOptions.fill` 复用 strict `ShapeFill`：值只能是 `{ kind: 'none' }`，或带有合法六位 sRGB / DrawingML theme color 的 `{ kind: 'solid', color, transparency? }`。`transparency` 必须是 finite `0..100`，并量化到 `0.001%`。创建时省略、runtime `undefined` 和 explicit none 都写 canonical direct `a:noFill`；solid 的显式 zero transparency 会保留为 `a:alpha val="100000"`。空对象、缺少 color、PptxGenJS 风格的 `{ color: 'FF0000' }`、unknown key、accessor、symbol、class instance 和越界值都会在 package mutation 前拒绝。

同一 contract 覆盖 plain/rich text、`addPlaceholder()`、title/body placeholder population、`SlideLayoutModel` / `SlideMasterModel` 的 text 创建，以及 `defineSlideMaster()` 的 declarative text/placeholder objects。创建结果可立即通过 `ShapeModel.fill` 读取、替换或清除；输入与快照保持 detached，same-value 是 exact no-op，duplicate、outer transaction rollback、六格式 write/reopen 和 placeholder source isolation 均已覆盖。

本项跨 package focused gate 为 5/5，SDK/root 与 adapter suites 分别为 188/188、76/76；最终全量 Vitest 为 1262 passed / 1 skipped，独立 performance 为 1/1（560ms），TypeScript typecheck 与 project build 均通过。Actual 57-file tarball 的 Node、declarations、browser export 与 installed CLI 均报告 `textShapeFills: true`；真实 Chrome validation/console/page/network errors 全为 0，CLI PowerPoint 2010 profile 为 0 errors / 0 warnings。

PptxGenJS 4.0.1 的 omitted fill 同样写 direct no-fill，但 `{ type: 'none' }` 会省略 direct fill choice，显式 zero transparency 也会省略 alpha；本库保留 explicit none/zero 的 direct intent。合法 solid 与非零透明度在最终语义上对等。Gradient/pattern/picture/group text fill 仍只做无损保留，不在 simple-fill 创建范围内。Text outer simple line、arrows、simple shadow、hyperlink、preset geometry、`rectRadius`、`isTextBox` 与 rich-text `breakLine` 已在后续小节支持。

## 创建和编辑文本框线条

```ts
const outlined = slide.addText('Outlined text box', {
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '2F5597' },
    transparency: 25,
    width: 2.5,
    dash: 'dashDot',
  },
});
const themed = slide.addRichText([{
  runs: [{ text: 'Theme outline' }],
}], {
  line: { kind: 'line', color: { kind: 'scheme', value: 'accent2' } },
});
const placeholder = slide.addPlaceholder('No outline', {
  name: 'outlined_title',
  type: 'title',
  line: { kind: 'none' },
});

outlined.line = { kind: 'line', color: { kind: 'scheme', value: 'accent3' } };
outlined.line = { kind: 'none' };
outlined.line = undefined; // 只清除 direct width/fill/dash
```

`AddTextOptions.line` 复用 strict `ShapeLine`。值只能是 `{ kind: 'none' }`，或 `{ kind: 'line', color, transparency?, width?, dash? }`；color 支持合法六位 sRGB/theme，transparency 是量化到 `0.001%` 的 finite `0..100`，width 是量化到 1 EMU 的 finite `0..1584` point，dash 是 `solid/dash/dashDot/lgDash/lgDashDot/lgDashDotDot/sysDash/sysDot`。省略 width/dash 会物化为 1pt/solid，zero width 与 explicit zero transparency 都保留 direct intent。省略、runtime `undefined` 与 explicit none 创建保持既有 canonical `<a:ln><a:noFill/></a:ln>`。

Plain/rich text、`addPlaceholder()`、placeholder population、layout/master wrappers 与 declarative `defineSlideMaster()` text/placeholder objects 共用同一 normalizer/renderer。创建结果可立即通过 live `ShapeModel.line` read/replace/clear；caller detachment、same-value bytes/journal no-op、duplicate isolation、outer rollback、stable identity、六格式 write/reopen 与 placeholder-source isolation 均已覆盖。PptxGenJS-shaped `type`/`dashType`/`alpha`/`lineDash`、missing color、invalid dash/range、unknown/accessor/symbol/class input 会在 mutation 前拒绝。

PptxGenJS 4.0.1 的 omitted/none/empty/missing-color text line 都输出 empty `a:ln`，省略 width/dash 时依赖隐式 1pt/solid，width zero 和 transparency zero 也被 falsy collapse；native 写明确可逆的 no-fill、默认 width/dash 与 zero direct state。合法 sRGB/theme、非零 transparency、正 width 和全部八种 dash 的 final semantics 对等。Nested deprecated `alpha` 在 PptxGenJS text line 中仍生效而 `lineDash` 被忽略；native 不接受两者。Gradient/pattern/picture/group line fill、custom dash 与 cap/compound/alignment/join 仍待后续；text arrows、simple shadow、hyperlink 与 preset geometry 已在后续小节支持。

本项跨 package focused gate 为 5/5，model、SDK、root 与 adapter suites 分别为 189/189、182/182、9/9、77/77。最终全量为 1268 passed / 1 skipped，独立 performance 为 1/1（553ms），两种 TypeScript build 与两套 package build 通过。Actual 57-file tarball 的 Node、declarations、browser export 与 CLI 均报告 `textShapeLines: true`；真实 Chrome immediate/detached/reopen state 完全匹配且 console/page/network 为 0，CLI PowerPoint 2010 profile 为 0 errors / 0 warnings。

## 创建和编辑文本框箭头

```ts
const arrowed = slide.addText('Flow', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
});
const rich = slide.addRichText([{
  runs: [{ text: 'Theme endpoints' }],
}], {
  arrows: { begin: 'none', end: 'stealth' },
});
const placeholder = slide.addPlaceholder('Arrow placeholder', {
  name: 'arrow_title',
  type: 'title',
  arrows: { end: 'diamond' },
});

arrowed.line = undefined; // 清除 line paint，保留 endpoints
arrowed.arrows = { begin: 'oval' }; // whole replacement，清除省略的 end
arrowed.arrows = undefined; // 清除 endpoints，保留当前 line state
```

`AddTextOptions.arrows` 复用 strict `ShapeArrows`，可选 `begin` / `end` 只接受 `none | arrow | diamond | oval | stealth | triangle`。省略、runtime `undefined` 或空对象不创建 endpoint；显式 `none` 保留 direct endpoint，和 absence 可区分。输入在 package mutation 前完成 descriptor-safe normalization 与脱离；alias、空字符串、非法 token、unknown/inherited/accessor/symbol key 和非普通对象都会零变更拒绝。

Plain/rich text、`addPlaceholder()`、placeholder population、layout/master wrappers 与 declarative `defineSlideMaster()` text/placeholder objects 共用同一 renderer。Arrow-only native 创建保留 canonical direct `a:noFill`；和 `line` 组合时固定按 line fill/dash → `headEnd` → `tailEnd` 输出。Live `ShapeModel.arrows` 继续使用 detached whole-replacement/clear 语义，line 与 endpoints 独立 ownership；duplicate、rollback、六格式 write/reopen、stable identity、异步 declarative detachment 与 placeholder-source isolation 均已覆盖。

PptxGenJS 4.0.1 对 omitted text line 写 empty `a:ln`，arrow-only 与 `{ type: 'none' }` + endpoint 也不写 no-fill；native 保留 canonical direct no-fill。PptxGenJS 会忽略 empty endpoint 与 nested/top-level `lineHead` / `lineTail`，并可把非法 endpoint token 原样写入 OOXML；native 不接受这些 alias 或非法值。六种合法 begin/end、单端/双端、显式 `none` 与 solid line + arrows 的 final endpoint semantics 对等，非法既有 endpoint 只做无损保留。

本项跨 package focused gate 为 4/4，model、SDK、root 与 adapter suites 分别为 191/191、184/184、10/10、78/78。最终全量为 1274 passed / 1 skipped，独立 performance 为 1/1（581ms），两种 TypeScript build 与两套 package build 通过。Actual 57-file tarball 的 Node、declarations、browser export 与 installed CLI 均报告 `textShapeArrows: true`；真实浏览器 immediate/detached/reopen state 完全匹配且错误为 0，CLI PowerPoint 2010 profile 为 0 errors / 0 warnings。Text-shape simple shadow 已在下一节支持。

## 创建和编辑文本框阴影

```ts
const shadowed = slide.addText('Shadowed heading', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
  shadow: {
    kind: 'outer',
    color: { kind: 'scheme', value: 'accent4' },
    opacity: 0.4,
    blur: 2,
    angle: 45,
    distance: 3,
    rotateWithShape: true,
  },
});
const rich = slide.addRichText([{
  runs: [{ text: 'Inner shadow' }],
}], {
  shadow: {
    kind: 'inner',
    color: { kind: 'srgb', value: '667788' },
    opacity: 0,
    blur: 0,
    angle: 0,
    distance: 0,
  },
});
const placeholder = slide.addPlaceholder('Shadow placeholder', {
  name: 'shadow_title',
  type: 'title',
  shadow: { kind: 'outer' },
});

shadowed.shadow = { kind: 'inner', opacity: 0.25 };
shadowed.shadow = undefined; // 只清除 direct shadow
```

`AddTextOptions.shadow` 复用 strict `ShapeShadow`。Outer/inner 都支持合法六位 sRGB/theme color、finite `0..1` opacity、`0..100pt` blur、`0 <= angle < 360°` 与 `0..200pt` distance；只有 outer 可提供 boolean `rotateWithShape`。省略字段归一化为 black、0.75、8pt、270°、4pt 与 outer rotate false，所有显式 zero 都保留。输入和嵌套 color 在 mutation 前脱离 caller；`type`/`offset` alias、coercible string、非法 range、unknown/inherited/accessor/symbol key、数组和 class instance 均以零 package 变化拒绝。

Plain/rich text、`addPlaceholder()`、placeholder population、layout/master wrappers 与 declarative `defineSlideMaster()` text/placeholder objects 共用同一 renderer。Shadow 在 line/endpoints 后写入唯一 canonical `a:effectLst`；fill、line、arrows 与 effect ownership 相互独立。Live `ShapeModel.shadow` 立即返回 detached deep-frozen snapshot，并支持 whole replacement/clear；same-value 是 exact bytes/journal no-op，`undefined` 只移除 direct inner/outer child，保留安全的 effect-list container 与 sibling effects。Duplicate、outer rollback、六格式 write/reopen、stable identity、异步 declarative detachment 与 placeholder-source isolation 均已覆盖。

PptxGenJS 4.0.1 对 omitted shadow 与 `{ type: 'none' }` 不写 direct effect，合法 outer final semantics 与 native 对等，legacy `offset` 对应 native `distance`。它会把 runtime zero 回退为 defaults、忽略 text `rotateWithShape: true`、纠正或宽松转换部分非法 type/color/number，并为 inner 写出不匹配的 closing tag；native 不复制这些缺陷，而是保留 zero、支持 theme color 和 rotate true、写合法 inner XML，并在 mutation 前严格拒绝非法输入。

本项 model/codec、SDK/root 与 adapter suites 分别为 234/234、197/197、79/79，跨 package focused gate 为 6/6。最终全量为 1280 passed / 1 skipped，独立 performance 为 1/1（607ms），两种 TypeScript build、两套 package build 与 declaration build 全部通过。Actual 57-file tarball 的 Node、declarations、browser export 与 installed CLI 均报告 `textShapeShadows: true`；真实浏览器 immediate/detached/reopen state 完全匹配且 console/page/network 为 0，CLI PowerPoint 2010 profile 为 0 errors / 0 warnings。

## 创建文本框超链接

```ts
const plain = slide.addText('打开产品页', {
  hyperlink: { url: 'https://example.com/product', tooltip: '查看产品' },
});
const rich = slide.addRichText([{
  runs: [
    { text: '跳转到' },
    {
      text: '详情页',
      style: {
        hyperlink: { url: 'https://example.com/details', tooltip: '' },
        underline: false,
      },
    },
    { text: '（不链接）', style: { hyperlink: false } },
  ],
}], {
  hyperlink: { slide: 2, tooltip: '' },
});

plain.hyperlink = { url: 'mailto:team@example.com' };
plain.hyperlink = undefined; // 只清除整个 shape 的 click link
```

`AddTextOptions.hyperlink` 接受与 preset shape 相同的 strict `Hyperlink`：恰好一个非空 URL 或当前文稿内的一基 slide target，并可带 `tooltip`。Plain/rich text、`addPlaceholder()`、placeholder population、layout/master wrappers 与 declarative `defineSlideMaster()` text/placeholder objects 全部支持；输入立即脱离 caller，非法、可强制转换或悬空的 target 会在任何 mutation 前拒绝。

创建时会在 non-visual shape click 与每个未覆盖的非空 text run 上写入 outer relationship；空段落和空 run 不伪造链接。`RichTextRunStyle.hyperlink?: Hyperlink | false` 可为每个 run 显式设置独立 URL/内部页 relationship，省略时继承 outer hyperlink，`false` 抑制 outer hyperlink。链接 run 默认写单下划线，显式 `RichTextRunStyle.underline` 始终优先。

`ShapeModel.richText` getter 只返回合法 direct run link；whole replacement setter 支持同值 no-op、relationship ID 复用、shared clone-on-write、clear/GC 与 rollback。`ShapeModel.hyperlink` 仍只管理 whole-shape click，不覆盖 run-local state。URL/internal-slide/tooltip 可跨 slide/layout/master/placeholder/declarative owners、duplicate、move、delete、六格式与 write/reopen 保存；self-link 复制后指向副本，删除 target 会清理 incoming run click。

PptxGenJS 4.0.1 对省略 tooltip 固定写 `tooltip=""`，rich outer hyperlink 会错误输出悬空的 `rIdundefined`；其合法 per-run 形式只写 run links，并与本库一样为每个显式 run 分配独立 relationship。本库接受其 external run 上的 `action=""`，但不复制 dangling/orphan/falsy-underline/console-only 缺陷。

最终全量 Vitest 为 1303 passed / 1 skipped，独立 performance 为 1/1（624ms）；model、SDK、root 与 adapter suites 分别为 199/199、191/191、13/13、80/80，两种 TypeScript build、两套 tsup 与 declaration build 全部通过。Actual 57-file tarball 的 installed Node/declarations/browser/CLI 与真实 Chrome 均报告 `richTextRunHyperlinks: true`；Chrome validation/console/page/network errors 均为 0。外链 smoke deck 为 24 parts / 32 relationships / 3 slides、0 errors 与 8 条预期 portability warnings；纯内链 deck 为 20 parts / 19 relationships / 2 slides、0 errors / 0 warnings。

## 创建和编辑预设文本框几何

```ts
import { type PresetShapeType } from '@jiayunxie/pptx';

const geometry: PresetShapeType = 'ellipse';
const text = slide.addText('Shaped text', { shape: geometry });
console.log(text.presetType); // 'ellipse'

text.presetType = 'hexagon';
```

`AddTextOptions.shape?: PresetShapeType` 接受 `PRESET_SHAPE_TYPES` 中全部 178 个 canonical OOXML token；省略或显式 `undefined` 时默认 `rect`。Plain/rich text、`addPlaceholder()`、named placeholder population、slide/layout/master wrappers 与 declarative `defineSlideMaster()` text/placeholder objects 共用同一严格 contract。创建结果可立即通过 live `ShapeModel.presetType` 读取和编辑；同值赋值保持 exact bytes 与 adjustments，换值只替换 direct geometry 并清空旧 adjustments。

Native 使用正确的 `foldedCorner`，并拒绝 PptxGenJS 4.0.1 的无效 `folderCorner`。空字符串、`false`、unknown string、number/string coercion、`custGeom` 与 accessor input 都会在 package mutation 前拒绝；custom geometry 继续使用 `addCustomShape()` / `ShapeModel.customGeometry`。`shape: 'line'` 选择文本 shape 的几何，`AddTextOptions.line` 选择 outline style，两者互不替代。

Geometry 与 fill、line、arrows、shadow、whole-shape/run hyperlink、transform、text body 和 placeholder identity 独立。Duplicate、move、rollback、六格式 write/reopen、layout/master source isolation 均已覆盖。最终全量为 1313 passed / 1 skipped，performance 为 1/1（578ms）；57-file tarball 的 Node/types/browser/CLI 与真实 Chrome 均报告 `textShapePresetGeometry: true`，Chrome validation/console/page/network errors 为 0。代表性 2 页文件在原件与 LibreOffice 回存后均保持全部 17 个 `(text, presetType)`，无 overflow 且逐页视觉一致；原件 PowerPoint 2010 validation 为 0 errors / 0 warnings，回存件为 0 errors 与 2 条 placeholder-owner warnings。

该小项不代表已实现完整 PptxGenJS 对等。`isTextBox` 与 rich-text `breakLine` 已在后续小节完成；其余 advanced text/table、`tableToSlides`、output/runtime helpers 与更广泛 peer/client audit 仍待完成。

## 创建圆角文本框半径

```ts
import { inches } from '@jiayunxie/pptx';

const rounded = slide.addText('Rounded text', {
  shape: 'roundRect',
  rectRadius: inches(0.5),
  width: inches(4),
  height: inches(2),
});

console.log(rounded.adjustments); // [{ name: 'adj', value: 25000 }]
rounded.adjustments = [{ name: 'adj', value: 12500 }];
```

`AddTextOptions.rectRadius?: Emu` 只用于 `shape: 'roundRect'`。输入必须是 finite `0..914400` EMU，按最近 EMU 取整；创建时写入 `adj = round(rectRadius × 100000 / min(finalWidth, finalHeight))`。省略或显式 `undefined` 保留 canonical empty `a:avLst`，显式 zero 保留 `adj=0`。Named placeholder population 使用最终 owner extent，而不是调用方临时 width/height。

该快捷字段仅参与创建。后续统一通过 `ShapeModel.adjustments` 读取、whole replacement 或 `[]` 清空；resize 不自动重算已有 guide。Plain/rich text、slide/layout/master、`addPlaceholder()`、placeholder population 与 declarative `defineSlideMaster()` text/placeholder objects 均受支持，且与 fill、line、arrows、shadow、hyperlink、transform 和 text body 独立。

PptxGenJS 4.0.1 的合法正值输出与本库语义一致。本库不会复制其 zero/NaN truthiness 丢失、string coercion、wrong-shape passthrough、负值/超范围输出或 `Infinity` formula；这些输入会在 package mutation 前拒绝。

最终全量为 1320 passed / 1 skipped；model、SDK、root 与 adapter suites 分别为 203/203、195/195、15/15、85/85。57-file actual tarball 的 Node/types/browser/CLI 与真实 Google Chrome 均报告 `textShapeRectRadius: true`，Chrome validation/console/page/network errors 为 0。三页 QA deck 通过 PowerPoint 2010 0 errors / 0 warnings、单部件 mutation isolation、无 overflow 与逐页视觉检查；LibreOffice 渲染和回存保留全部显式 guides，仅将 omitted 默认值物化为 `16667`。

## 创建和编辑文本框标记

```ts
const shapeText = slide.addText('Shape text');
console.log(shapeText.isTextBox); // false

const textBox = slide.addText('Text box', { isTextBox: true });
console.log(textBox.isTextBox); // true
textBox.isTextBox = false;
```

`AddTextOptions.isTextBox?: boolean` 只管理当前 `p:sp/p:nvSpPr/p:cNvSpPr@txBox`。Omitted、own data property `undefined` 与 `false` 都不写该属性；`true` 写 canonical `txBox="1"`。Inherited property 按 absent 处理，accessor 不会执行，defined value 只接受 primitive boolean；string、number、boxed boolean、object 与其他 truthy 值均在 package mutation 前拒绝。

Live `ShapeModel.isTextBox` 对 attribute absence 返回 `false`，接受 `1/true/on` 与 `0/false/off`，对 malformed、qualified lookalike、重复 attribute 或歧义 owner 结构返回 `undefined`。Setter 只接受 boolean：`true` canonicalize 为 `txBox="1"`，`false` 删除唯一 direct attribute；canonical 同值是 exact bytes/journal no-op，单一 alias 或 malformed token 可修复，歧义结构在零变更下拒绝。

Plain/rich text、`addPlaceholder()`、layout/master direct methods、declarative `defineSlideMaster()` text/placeholder objects 与六种 presentation format 共用这一状态。Layout placeholder materialization 保留 source 的 direct boolean；named placeholder population 会先验证 call-site 值，但最终由 layout source 覆盖，且不会修改 layout/master source。`isTextBox` 与 preset/custom geometry、adjustments/`rectRadius`、transform、text、fill、line、arrows、shadow、hyperlink 与 placeholder identity 独立。

PptxGenJS 4.0.1 的合法 boolean public contract 与 native 最终状态对等；其 runtime 会把任意 truthy 值写成 true，本库不复制该宽松行为。最终全量 Vitest 为 1337 passed / 1 skipped，独立 performance 为 1/1（704ms）；57-file actual tarball 的 SHA-256 为 `2c6afd9bdb1f4c076ff0d0eb8bc8e8711793ae46dc320b2978d08f5e3a44b41a`，Node、browser conditional export、declarations、installed CLI 与真实 Google Chrome 均报告 `textShapeIsTextBox: true`，Chrome validation/console/page/network errors 为 0。

两页 QA deck 的 PowerPoint 2010 profile 为 0 errors，并只有 fixture 外部超链接产生的 1 条预期 portability warning；单项切换只改变 `/ppt/slides/slide1.xml`，目标 `<p:cNvSpPr/>` 仅增加 ` txBox="1"`，其余 21 parts byte-identical。原件与 LibreOffice 渲染均为 0 overflow 并通过逐页视觉检查。LibreOffice 回存会移除 native 与独立 PptxGenJS 4.0.1 文件中的所有 true `txBox` 状态，因此这是已确认的客户端重写边界，不作为库内 workaround 或完整 round-trip 声明。

## 创建富文本段落换行

```ts
const rich = slide.addRichText([{
  align: 'center',
  runs: [
    { text: '第一段', breakLine: true },
    { text: '', breakLine: true }, // 保留中间空段
    { text: '第三段', softBreakBefore: true },
  ],
}]);

console.log(rich.richText.length); // 3；getter 返回规范段落，不返回 breakLine 标记
```

`RichTextRun.breakLine?: boolean` 是创建和 whole-rich-text replacement 的输入语法。非末尾 run 的 `true` 会在该 run 后结束当前段落；middle、empty 与 consecutive flags 会保留相应空段，末尾 flag 被消费但不会额外生成尾部空段。每个拆分段复制原输入段落的 align、RTL、margin、indent、bullet、level、spacing 与 tab stops；`softBreakBefore` 仍附着于原 run，即使拆分后成为新段首 run。它也可与 run-local URL/内部页 hyperlink 组合，关系索引会按规范段落重新分配。

该字段覆盖 slide/layout/master、placeholder prompt/population、declarative master、live edit、duplicate/move/rollback/reopen 与六种 presentation format。Getter 只返回显式 `RichTextParagraph[]`，不猜测或回放私有 marker；把 getter 快照重新赋值仍保持相同语义。`breakLine` 不属于 outer `AddTextOptions`，run text 中的 CR/LF 也不会被当作该字段的快捷方式。

PptxGenJS 4.0.1 合法 boolean 输入的段落、属性与 hyperlink 最终语义对等；本库严格拒绝 string、number、null、object、boxed boolean 等 truthy/falsy runtime 值。PptxGenJS 会抑制拆分后段首 run 的 soft break，本库保留现有可逆 `softBreakBefore` contract。最终 release gates 为 1350 passed / 1 skipped，performance 1/1；57-file tarball SHA-256 为 `d06b84c0c3b8ff8e610c87c55b0fe9b67de6b41e59b5ec7fad62b206fdbe2699`，installed Node/types/browser/CLI 与真实 Chrome 均报告 `richTextBreakLine: true`，Chrome validation/console/page/network errors 为 0。

四页 source deck 在 PowerPoint 2010 profile 下为 0 errors / 0 warnings；单项 edit 只改变 `slide1.xml` 与对应 relationships，其他 24 parts byte-identical。五页 source/LibreOffice 视觉件均无 overflow、裁切或意外换行。LibreOffice 保留可见段落、空行、软换行与内部页链接，但会合并相邻 runs、省略空 tooltip、下推 master 内容、重命名 placeholder，并丢弃 master placeholder prompt；因此不声明 owner identity 完整往返。

完整 PptxGenJS 对等仍需继续完成其余 advanced text/table、`tableToSlides`、output/runtime helpers 与 peer/client audit。

## 读取库运行时版本

```ts
import {
  PPTX_VERSION,
  PptxDocument,
  type PptxVersion,
} from '@jiayunxie/pptx';

const current: PptxVersion = PPTX_VERSION; // '0.1.0'
const document = PptxDocument.create();
console.log(document.version === current); // true
```

`PPTX_VERSION` 是 browser-safe 的编译期字面量常量，`PptxVersion` 是对应字面量类型；只读 `PptxDocument.version` 在 create/open/write/reopen 全生命周期始终返回同一当前库版本，不读取 package 路径，也不修改 presentation。它不是 OOXML extended-properties 中的 `AppVersion`，也不表示输入文件的 producer 或 PowerPoint 版本。

PptxGenJS 4.0.1 实例返回它自己的 `'4.0.1'`，本库当前返回自己的 `'0.1.0'`；两个值不相等是正确行为，对等点是公开可读、稳定且各自与 manifest 同步。三份 manifest 由测试防漂移，CLI `--version` 与 JSON doctor 共用同一常量。最终 release gates 为 1354 passed / 1 skipped，performance 1/1（617ms），两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 58-file tarball 的 SHA-256 为 `ce300d3c5da10a8fbdb9910b10497d02af496532b99329e18a314c9604e6f9a8`；installed Node/types/browser/CLI 与真实 Google Chrome 均报告 `presentationVersion: true`，Chrome validation/console/page/network errors 为 0。

该历史检查点仍不声明完整 PptxGenJS 对等；`OutputType` runtime catalog 已在后续专项完成，仍待六种实际 `write({ outputType })` 返回语义、stream、compression、其余 runtime helper constants、advanced text/table、`tableToSlides` 与最终 peer/client audit。

## 读取当前演示文稿布局

```ts
import {
  inches,
  PptxDocument,
  type PresentationLayout,
  type PresentationLayoutName,
} from '@jiayunxie/pptx';

const document = PptxDocument.create({ slideSize: '16:9' });
const layout: PresentationLayout = document.presLayout;
const name: PresentationLayoutName = layout.name; // 'screen16x9'

document.slideSize = { width: inches(11.7), height: inches(8.3) };
console.log(document.presLayout); // { name: 'custom', width: 10698480, height: 7589520 }
```

Getter-only `PptxDocument.presLayout` 从唯一的 `p:sldSz` / `slideSize` 状态即时投影 `{ name, width, height }`，宽高单位为 EMU。10×7.5、10×5.625、10×6.25 inch 精确映射为 `screen4x3`、`screen16x9`、`screen16x10`；其他合法尺寸（包括 `wide`）映射为 `custom`。每次读取返回 detached plain-object snapshot；修改旧快照不会改变文稿，读取也不产生 OPC mutation。修改 `slideSize` 后下一次读取立即反映新值，write/reopen 后保持。

PptxGenJS 4.0.1 的公开 getter 对默认、四种内建和自定义尺寸使用相同 EMU 数值；native 不暴露其未声明的 `_sizeW` / `_sizeH`，也不返回内部 mutable alias。`defineLayout()` 的自定义名称只存在于 PptxGenJS 进程内且不写入 PPTX，所以 native 打开自定义尺寸时使用可恢复的 canonical `custom`，而不伪造命名 registry。

最终 release gates 为 1363 passed / 1 skipped，performance 1/1（1.01s），两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 59-file tarball 的 SHA-256 为 `a07a11156840071f0945289c0a48fdd9741549d2003ca21006e6efab28104b3d`；installed Node/types/browser/CLI 与真实 Google Chrome 均报告 `presentationLayouts: true`，Chrome validation/console/page/network errors 为 0。

该历史检查点仍不声明完整 PptxGenJS 对等；`OutputType` runtime catalog 已在后续专项完成，仍待六种实际 `write({ outputType })` 返回语义、stream、compression、其余 runtime constants、advanced text/table、`tableToSlides` 与最终 peer/client audit。

## 枚举水平文字对齐值

```ts
import { TEXT_ALIGNMENTS, type TextAlignment } from '@jiayunxie/pptx';

for (const alignment of TEXT_ALIGNMENTS) {
  const value: TextAlignment = alignment;
  console.log(value);
}
```

`TEXT_ALIGNMENTS` 是 frozen readonly tuple，顺序固定为 `left`、`center`、`right`、`justify`；`TextAlignment` 直接由该 tuple 派生，因此 runtime discovery 与 TypeScript union 不会漂移。四个 token 可直接用于 plain/rich text 与 table/table-cell 水平对齐 API，既有 `left → l`、`center → ctr`、`right → r`、`justify → just` OOXML 映射、非法值拒绝和 write/reopen 语义均未改变。

PptxGenJS 4.0.1 的实例 `AlignH` 公开相同的四个 keys/values。Native 对等这些 runtime values 和稳定枚举顺序，但采用现有 root catalog 模式，不增加 `PptxDocument.AlignH`、enum-shaped object 或 mutable alias。最终 release gates 为 1368 passed / 1 skipped tests，performance 1/1（1.17s），两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 59-file tarball 的 SHA-256 为 `46a88495acbffb2f81eb99f290bd15928974ddad86f507941c1ea5bfb65eaa90`；installed Node/types/browser/CLI 与真实 Google Chrome 均报告 `horizontalAlignments: true`，Chrome validation/console/page/network errors 为 0。

该检查点总体 PptxGenJS 对等进度约 96%；`AlignV` 已在下一节完成。

## 枚举垂直文字对齐值

```ts
import {
  PptxDocument,
  TEXT_VERTICAL_ALIGNMENTS,
  type TextBoxVerticalAlignment,
} from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();
const alignment: TextBoxVerticalAlignment = TEXT_VERTICAL_ALIGNMENTS[1];

slide.addText('Centered vertically', { valign: alignment });
slide.slideNumber = { valign: alignment };
slide.addTable([[{ text: alignment, options: { valign: alignment } }]]);
```

`TEXT_VERTICAL_ALIGNMENTS` 是 frozen readonly tuple，顺序固定为 `top`、`middle`、`bottom`；`TextBoxVerticalAlignment` 直接由该 tuple 派生，因此 runtime discovery 与 TypeScript union 不会漂移。三个 token 可用于 text box、slide number 与 table/table-cell 垂直对齐，既有 `top → t`、`middle → ctr`、`bottom → b` OOXML 映射、非法值拒绝、默认值和 write/reopen 语义均未改变。

PptxGenJS 4.0.1 的实例 `AlignV` 公开相同的三个 keys/values。Native 对等这些 runtime values 和稳定枚举顺序，但沿用 root catalog 设计，不增加 `PptxDocument.AlignV`、enum-shaped object 或 mutable alias。最终 clean gates 为 72 passed / 1 skipped test files、1373 passed / 1 skipped tests，performance 1/1（939ms），两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 59-file tarball 的 SHA-256 为 `aaaa5e0ceb053a472af49732784e0ea5babb00968734ec5093cd4f80afc34095`；installed Node/types/browser/CLI 与真实 Google Chrome 均报告 `verticalAlignments: true`，Chrome values、text/table reopen、frozen catalog 检查通过，validation/console/page/network errors 为 0。

该检查点总体 PptxGenJS 对等进度约 97%；`OutputType` runtime catalog 已在下一节完成。

## 枚举演示文稿输出类型

```ts
import { OUTPUT_TYPES, type OutputType } from '@jiayunxie/pptx';

const outputTypes: readonly OutputType[] = OUTPUT_TYPES;
for (const outputType of outputTypes) {
  console.log(outputType);
}
```

`OUTPUT_TYPES` 是 frozen readonly tuple，顺序固定为 `arraybuffer`、`base64`、`binarystring`、`blob`、`nodebuffer`、`uint8array`；`OutputType` 直接由该 tuple 派生。Catalog 由 SDK 输出层拥有，aggregate root 复用同一对象；读取或遍历不访问或修改任何文稿 package。

PptxGenJS 4.0.1 的实例 `OutputType` 公开相同六个 keys/values。Native 对等 runtime values 和稳定顺序，但不增加 `PptxDocument.OutputType`、enum-shaped object 或 mutable alias；`STREAM` 不属于该 public enum，因此留给独立 stream API。该历史检查点只发布 catalog；六种实际返回语义已在下一节完成。

最终 release gates 为 73 passed / 1 skipped test files、1378 passed / 1 skipped tests，performance 1/1（1.10s），两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 60-file tarball 的 SHA-256 为 `31a38643c8c851ae24a381a68cd225972b76dbf7b37758c16efd2fe27248df0d`；installed Node/types/browser/CLI 与真实 Google Chrome 均报告 `outputTypes: true`。Chrome 六值、frozen catalog 和 mutation isolation 检查通过，页面、bundle 与 blob 请求均为 200，console/page/network errors 为 0。

总体 PptxGenJS 对等进度仍约 97%。六值 `write({ outputType })` 返回语义已在下一节完成。

## 选择 `write()` 输出类型

```ts
import {
  PptxDocument,
  type OutputType,
  type WriteOptions,
  type WriteOutput,
} from '@jiayunxie/pptx';

const document = PptxDocument.create();
document.addSlide().addText('Hello');

const bytes = await document.write(); // Uint8Array
const base64 = await document.write({ outputType: 'base64' }); // string
const blob = await document.write({ outputType: 'blob' }); // Blob

async function encode<T extends OutputType>(
  options: WriteOptions<T>,
): Promise<WriteOutput<T>> {
  return document.write(options);
}
```

`write({ outputType })` 现在支持全部六个公开 token：`arraybuffer` 返回独立 `ArrayBuffer`，`base64` 返回不带 data-URI 前缀的 base64 字符串，`binarystring` 返回每个 code unit 表示一个 byte 的字符串，`blob` 返回 MIME 为 `application/zip` 的 `Blob`，`nodebuffer` 在 Node 返回 `Buffer`，`uint8array` 返回普通 `Uint8Array`。浏览器请求 `nodebuffer` 会明确拒绝并报告 `nodebuffer is not supported by this platform`。

未传 options、传空对象或只传 validation options 时继续返回 `Uint8Array`。`WriteOutput<T>` 与泛型 `WriteOptions<T>` 保留 literal output type 的精确返回类型；为避免 browser declarations 引入 `node:buffer`，`nodebuffer` 的公开结构类型是 `Uint8Array`，Node runtime 值仍是 `Buffer`。所有转换都基于同一 canonical ZIP bytes，不修改 package、diagnostics 或 mutation journal。

显式 `outputType: 'blob'` 遵循 PptxGenJS/ZIP 输出语义，MIME 为 `application/zip`；便捷 `writeBlob()` 继续返回 `application/vnd.openxmlformats-officedocument.presentationml.presentation`。`writeFile()` 与 `download()` 的既有契约也不变。

最终 clean gates 为 74 passed / 1 skipped test files、1383 passed / 1 skipped tests，performance 1/1（966ms），两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 61-file tarball 的 SHA-256 为 `26bbc7eb7c33eb194388576db2c2eaab33c80d0d99b19ed7a9b4a7375c3f9f37`；installed Node/types/browser/CLI 与真实 Google Chrome 均报告 `writeOutputTypes: true`。六种 Node 输出和五种浏览器输出均 byte-identical、可重开；Chrome validation/console/page/network errors 为 0。

总体 PptxGenJS 对等进度约 97%。Node readable stream 已在下一节完成。

## 输出 Node Readable stream

```ts
import { createWriteStream } from 'node:fs';
import { PptxDocument, type PptxNodeReadableStream } from '@jiayunxie/pptx';

const document = PptxDocument.create();
document.addSlide().addText('Stream output');

const readable: PptxNodeReadableStream = await document.stream();
for await (const chunk of readable) {
  console.log(chunk.byteLength);
}

const file = createWriteStream('output.pptx');
(await document.stream()).pipe(file);
```

`stream(options?: WriteBaseOptions)` 只在 Node.js 提供真正的 non-object-mode `Readable`，支持 `pipe()`、async iteration、`data/end/error` events、pause/resume/read/destroy。每个输出 chunk 最大 64 KiB，按顺序拼接后与同一状态的 `write()` byte-identical，并可直接重开；strict/permissive 与 compatibility diagnostics 继续复用同一 `#writeBytes()` 路径。

当前 ZIP writer 仍会先在内存中生成完整 canonical `Uint8Array`，再通过 Readable 进行 backpressure-aware delivery；因此这里不宣称 constant-memory ZIP generation 或更早的 time-to-first-byte。浏览器调用会在 validation/ZIP write 前明确拒绝并报告 `PptxDocument.stream() is only supported in Node.js`，应改用 `write()`、`writeBlob()` 或 `download()`。

PptxGenJS 4.0.1 的 public `stream()` 实际返回 Buffer，而不是真正的 Readable；该 byte-result 已由 `write({ outputType: 'nodebuffer' })` 对等。本库把 `stream()` 用于真实 Node stream，并保留 `STREAM` 不进入 `OUTPUT_TYPES`。最终 clean gates 为 75 passed / 1 skipped test files、1390 passed / 1 skipped tests，performance 1/1（682ms）。实际 61-file tarball SHA-256 为 `37b1d6bec7b5a144d577c57b61c0777f2aad8515015e9cbee05abd55f8e067d2`；installed Node/types/browser/CLI 与真实 Chrome 均报告 `nodeReadableStream: true`，Chrome validation/console/page/network errors 为 0。

总体 PptxGenJS 对等进度约 97%。Compression policy 已在下一节完成。

## 控制演示文稿 ZIP 压缩

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
document.addSlide().addText('Compression policy');

const compressed = await document.write({
  outputType: 'uint8array',
  compression: true,
});
const readable = await document.stream({ compression: true }); // Node.js
await document.writeFile('compressed.pptx', { compression: true }); // Node.js
const uncompressedBlob = await document.writeBlob({ compression: false }); // browser
```

`compression` 是严格的 primitive boolean。省略、`undefined` 或 `false` 对新建/已修改文稿使用 ZIP STORE；`true` 使用 DEFLATE level 6。该策略在六种 `write({ outputType })`、`stream()`、`writeFile()`、`writeBlob()` 和 `download()` 上一致，压缩只改变 ZIP 表示，不改变 OOXML 语义、diagnostics 或 mutation journal。字符串、数字、对象和 boxed Boolean 会在 OPC 写入前以 `PptxDocument output compression must be a boolean` 拒绝。

唯一的 fast path 是打开后完全未修改的文稿：省略或传 `undefined` 会返回原始 bytes，从而保留原文件的压缩、entry metadata 与无关字节；显式 `false` 或 `true` 总会重新打包为 STORE 或 DEFLATE。PptxGenJS 4.0.1 的合法 boolean 意图保持对等，但本库不会复制其显式 `outputType` 路径丢失 compression 或对 truthy 非 boolean 值进行 coercion 的行为。

最终 clean gates 为 1400 passed / 1 skipped tests，performance 1/1（749.5ms），两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 61-file tarball SHA-256 为 `4bbaa25b83a0d20dd3d2239708c628afec79bcff19c69faca6fe67b03e3bd990`；packed Node 的代表性 STORE/DEFLATE 输出为 149,598/9,347 bytes，真实 Chrome 为 84,062/9,270 bytes。Installed Node/types/browser/CLI 与 Chrome 均报告 `compressionPolicy: true`，Chrome download 使用 ZIP method 8、可重开，console/page/network errors 为 0。

总体 PptxGenJS 对等进度仍约 97%。`SchemeColor` runtime helper 已在下一节完成。

## 使用 PptxGenJS 对等的主题色 helper

```ts
import {
  PptxDocument,
  SCHEME_COLORS,
  type SchemeColor,
} from '@jiayunxie/pptx';

const accent: SchemeColor = SCHEME_COLORS.accent1;
const document = PptxDocument.create();
document.addSlide().addRichText([{
  runs: [
    { text: 'Theme text', style: { color: { kind: 'scheme', value: SCHEME_COLORS.text1 } } },
    { text: ' accent', style: { color: { kind: 'scheme', value: accent } } },
  ],
}], {
  fill: { kind: 'solid', color: { kind: 'scheme', value: SCHEME_COLORS.background1 } },
});
```

`SCHEME_COLORS` 按稳定顺序公开 `text1→tx1`、`text2→tx2`、`background1→bg1`、`background2→bg2` 和 `accent1..accent6` 十项映射，`SchemeColor` 直接由这些 values 派生。Catalog 在 model 中创建并冻结，SDK 与 aggregate root 复用同一对象；读取它不访问或修改 presentation package，Node 与浏览器结果一致。

PptxGenJS 4.0.1 的 `SchemeColor` getter 公开相同 keys、values 和顺序，但返回共享 mutable object；native 采用 frozen root mapping，不增加 `PptxDocument.SchemeColor`、instance getter 或第二份 catalog。`SchemeColor` 只表示 PptxGenJS helper 的十个值；native 既有 color APIs 仍接受经验证的更广 DrawingML scheme-color 子集，因此它不是全部 OOXML 主题色 token 的穷举类型。

最终 clean full Vitest 为 77 passed / 1 skipped test files、1404 passed / 1 skipped tests，performance 1/1（736ms）；两种 TypeScript check、Node/browser bundle 与 declaration build 全部通过。实际 62-file tarball SHA-256 为 `5d7096b0347d605c105dff15bb357781c4dcaa1cb7c3eff69f89ea6baa70e742`；installed Node/types/browser/CLI 门禁全部通过，Node 与真实 Google Chrome 均报告 `schemeColors: true`。Write/reopen、frozen/shared identity、mutation isolation 通过，Chrome validation/console/page/network errors 为 0。证据保存在 `/tmp/pptx-scheme-color-artifacts.AOU1Qb`。

## 表格级垂直对齐读取与批量编辑

```ts
import { PptxDocument, type TextBoxVerticalAlignment } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { valign: 'middle' });

const uniform: TextBoxVerticalAlignment | undefined = table.verticalAlignment; // middle
table.setCellVerticalAlignment(0, 1, 'top');
console.log(table.verticalAlignment); // undefined: mixed direct cell state
table.verticalAlignment = 'bottom';  // atomically overwrites every physical cell
table.verticalAlignment = undefined; // clears every direct tcPr@anchor
```

`TableModel.verticalAlignment` 是全部物理 cell direct `tcPr@anchor` 的共识投影：一个或多个 cell 都具有同一合法值时返回 `top`、`middle` 或 `bottom`；mixed、absent、empty 或 unsafe state 返回 `undefined`。赋值会原子覆盖全部 physical cells（包括 merge continuation），`undefined` 会清除全部 direct anchor；同值和全 absent clear 是 exact no-op。DrawingML 没有在这里保存 table creation default，本属性也不会合成或记忆默认值；需要查看 mixed 细节时使用 `rows[].cells[].verticalAlignment`。结构不安全的写入以 `ModelParseError` 零 mutation 拒绝，并保留文本、边框、填充、margin、方向、fit、grid、row 与 transform。

PptxGenJS 4.0.1 只提供创建期 table/cell `valign`，生成的最终 direct cell anchors 可被该共识 getter 精确读取；native bulk editor 是对相同 OOXML state 的 lossless existing-deck 扩展。Focused 为 4 files / 521 tests，最终 full 为 78 passed / 1 skipped test files、1411 passed / 1 skipped tests，performance 1/1（885ms）。实际 62-file tarball SHA-256 为 `6ce48d8bb73d59148754f14dc379b9cd11ba34d358dd8e7ebba7b72cf8208f1e`；installed Node/types/browser/CLI 与真实 Chrome 均报告 `tableVerticalAlignment: true`，Chrome validation/console/page/network errors 为 0。证据位于 `/tmp/pptx-table-vertical-alignment-artifacts.1kZjyy`。

## 表格级文本方向读取与批量编辑

```ts
import { PptxDocument, type TableCellTextDirection } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { textDirection: 'vert' });

const uniform: TableCellTextDirection | undefined = table.textDirection; // vert
table.setCellTextDirection(0, 1, 'vert270');
console.log(table.textDirection); // undefined: mixed direct cell state
table.textDirection = 'horz';     // 显式写入每个 physical cell 的 vert="horz"
table.textDirection = undefined;  // 清除全部 direct tcPr@vert
```

`TableModel.textDirection` 是全部物理 cell direct `tcPr@vert` 的严格共识投影：只有每个 cell 都存在同一个合法 direct token 时才返回 `horz`、`vert`、`vert270` 或 `wordArtVert`；absent、mixed、empty 或 unsafe state 都返回 `undefined`。它不会把属性缺失合成为 `horz`。赋值会在单一事务中覆盖全部 physical cells（包括 merge continuation），其中 `horz` 写显式属性，只有 `undefined` 才清除属性；同值和全 absent clear 是 exact no-op。需要查看 mixed 明细时使用 `rows[].cells[].textDirection`。不安全结构会以 `ModelParseError` 零 partial mutation 拒绝，并保留所有无关 cell/table 状态。

PptxGenJS 4.0.1 创建期会把 resolved `horz` 折叠为属性缺失，因此这类文件导入后的 `table.textDirection` 是 `undefined`，而不是 `horz`；三个非水平值可按最终 direct state 精确读取。显式 native `horz` 与 existing-deck bulk edit 是相同 OOXML direct state 上的 lossless 扩展。Focused 为 5 files / 529 tests，最终 full 为 79 passed / 1 skipped test files、1419 passed / 1 skipped tests，performance 1/1（1118ms）。实际 62-file tarball SHA-256 为 `5f427a8ff77cf64f6dda593ec02fdbe405c44d22481f0357bf05fa39b63ec92d`；installed Node/types/browser/CLI 与真实 Chrome 均报告 `tableTextDirection: true`，Chrome validation/console/page/network errors 为 0。证据位于 `/tmp/pptx-table-text-direction-artifacts.BksCOP`。

## 表格级水平对齐读取与批量编辑

```ts
import { PptxDocument, type TextAlignment } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { align: 'center' });

const uniform: TextAlignment | undefined = table.horizontalAlignment; // center
table.setCellHorizontalAlignment(0, 1, 'right');
console.log(table.horizontalAlignment); // undefined: mixed direct cell state
table.horizontalAlignment = 'justify'; // 覆盖全部 physical cells
table.horizontalAlignment = 'left';    // 显式写入每个 pPr@algn="l"
table.horizontalAlignment = undefined; // 清除全部 direct pPr@algn
```

`TableModel.horizontalAlignment` 是全部物理 cell 唯一单段落 direct `pPr@algn` 的严格共识投影：只有每个 cell 都存在同一个安全合法值时才返回 `left`、`center`、`right` 或 `justify`；absent、mixed、empty、multi-paragraph 或 unsafe state 都返回 `undefined`，且属性缺失绝不合成为 `left`。赋值会在单一事务中覆盖全部 physical cells（包括 merge continuation），其中 `left` 写显式 `algn="l"`，只有 `undefined` 才清除属性；同值和全 absent clear 是 exact no-op。需要查看 mixed 明细时使用 `rows[].cells[].horizontalAlignment`。不安全结构会以 `ModelParseError` 零 partial mutation 拒绝，并保留所有无关 cell/table 状态。

PptxGenJS 4.0.1 的四个合法 table alignment 创建值可按最终 direct state 精确导入；omitted alignment 保持 `undefined`，table center 加单 cell right override 会投影为 mixed `undefined`。Native existing-deck bulk edit 是相同 OOXML direct state 上的 lossless 扩展。Focused 为 6 files / 537 tests，最终 full 为 80 passed / 1 skipped test files、1427 passed / 1 skipped tests，performance 1/1（1.34s）。实际 62-file tarball SHA-256 为 `03b376861aeb799fa21a99dd105871b8943e29bd4fe51c875a508ff295b9f9c0`；installed Node/types/browser/CLI 与真实 Chrome 均报告 `tableHorizontalAlignment: true`，CLI 确认最终 slide 恰有四个 direct `pPr@algn="r"` 且无 `tcPr/bodyPr@algn`，Chrome validation/console/page/network errors 为 0。证据位于 `/tmp/pptx-table-horizontal-alignment-artifacts.oe2f5A`。

## 表格级 margins 读取与批量编辑

```ts
import { PptxDocument, type TextBoxMargins } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { margin: [3.6, 7.2, 10.8, 14.4] });

const uniform: TextBoxMargins | undefined = table.margins;
table.setCellMargins(0, 1, { top: 9 });
console.log(table.margins);       // undefined: mixed direct cell state
table.margins = 6;               // 四边均为 6pt
table.margins = { top: 2, left: 4 }; // whole-replace，并清除 right/bottom
table.margins = undefined;       // 清除全部 direct marL/marR/marT/marB
```

`TableModel.margins` 是全部 physical cells（包括 merge continuation）唯一 direct `tcPr@marL/marR/marT/marB` 的严格共识投影。只有每个 cell 都具有同一组非空、安全 complete 或 partial 边和值时才返回 detached `TextBoxMargins`；absent、mixed keys/values、empty、malformed、repeated 或 ambiguous state 返回 `undefined`。它不会从 canonical defaults、table style 或创建输入合成值；mixed 明细继续通过 `rows[].cells[].margins` 读取。

赋值接受 point scalar、TRBL tuple、partial named object、`{}` 或 `undefined`，并在单一事务中 whole-replace 全部 physical cells。Scalar/TRBL 写四边，partial object 清除遗漏边，`{}`/`undefined` 清除四边；同数值和全 absent clear 是 exact bytes/journal no-op。不安全写入以 `ModelParseError` 零 package partial mutation 拒绝，文本、border、fill、alignment、direction、fit、grid、rows、transform 与 relationships 均保留。

PptxGenJS 4.0.1 只提供创建期 margin，最终只留下 direct cell attributes。Omitted creation 会导入其显式 canonical `{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }`；合法 uniform table margin 可投影，单 cell override 会得到 mixed `undefined`。其 legacy runtime 把首项 `<1` 解释为 inches、`>=1` 解释为 points；native 始终使用 points，所以 PptxGenJS `[0.05, 0.1, 0.15, 0.2]` 对等 native `[3.6, 7.2, 10.8, 14.4]` 的最终 EMU，而不是复制输入单位歧义。

Focused 为 7 files / 547 tests，最终 full 为 81 passed / 1 skipped test files、1437 passed / 1 skipped tests，performance 1/1（1.65s）。实际 62-file tarball SHA-256 为 `428f47de86cebb89ae19a59b4b5500f3c67c116f63107253e2bf997b04008e37`；installed Node/types/browser/CLI 与真实 Chrome 均报告 `tableMargins: true`。CLI 确认最终四个 direct `tcPr` 都恰有 `marL="50800" marR="25400" marT="12700" marB="38100"`，Chrome validation/console/page/network errors 为 0。证据位于 `/tmp/pptx-table-margins-artifacts.gPmz7V`。

## 表格级 fill 读取与批量编辑

```ts
import { PptxDocument, type TableCellFill } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
});

const uniform: TableCellFill | undefined = table.fill;
table.setCellFill(0, 1, { kind: 'none' });
console.log(table.fill); // undefined：mixed direct cell state
console.log(table.rows[0]!.cells[1]!.fill); // { kind: 'none' }
table.fill = { kind: 'none' }; // 覆盖全部 physical cells
table.fill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'D9EAF7' },
  transparency: 0, // 显式 direct alpha=100000
};
table.fill = undefined; // 清除全部 direct fill choice
```

`TableModel.fill` 只在全部 physical cells（包括 merge continuation）都具有同一个安全、受支持的 direct `tcPr` fill choice 时返回 detached `TableCellFill`。Absent、mixed、malformed、advanced 或 ambiguous state 返回 `undefined`；它不会解析 table style、effective default 或创建期输入。Mixed 明细始终保留在 `rows[].cells[].fill`。赋值会在单一事务中 whole-replace 全部 physical cells；`{ kind: 'none' }` 写 direct `a:noFill`，solid 支持 strict sRGB/theme color 与可选 `0..100` transparency，`undefined` 清除 direct choice。同值与全 absent clear 是 exact no-op，不安全写入以 `ModelParseError` 零 partial mutation 拒绝。

PptxGenJS 4.0.1 的合法 uniform table solid fill 可按最终 direct state 投影，omitted fill 为 `undefined`，table fill 加不同 cell override 为 mixed `undefined`。PptxGenJS 会把 `type: 'none'` 与 omitted 都折叠为没有 direct fill，并把 transparency 0 折叠为没有 alpha；native direct none、omitted、显式零透明度三种 direct intent 保持可区分。Existing-deck 共识读取和批量编辑是 native lossless extension。

Focused 为 8 files / 557 tests；最终 full 为 1447 passed / 1 skipped tests，performance 1/1（1.17s）。实际 62-file tarball SHA-256 为 `ae7f09233b2ff596c21ec0dab891d5069d810d64921bce9ba96cd08771c6cfdc`；installed Node/types/browser/CLI、`pptx-inspect` 与真实 Google Chrome 均报告 `tableFill: true`，最终四个 physical cells 各有且仅有一个 direct `a:noFill`，PowerPoint 2010 validation 与 Chrome validation/console/page/network errors 均为 0。证据位于 `/private/tmp/pptx-table-fill-artifacts.5MqrXK`。

## 表格级 borders 读取与批量编辑

```ts
import { PptxDocument, type TableCellBorders } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], {
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 1.5,
    style: 'dash',
  },
});

const uniform: TableCellBorders | undefined = table.borders;
table.setCellBorders(0, 1, { kind: 'none' });
console.log(table.borders); // undefined：mixed direct cell state
console.log(table.rows[0]!.cells[1]!.borders); // 四边 direct none
table.borders = {
  top: {
    kind: 'line',
    color: { kind: 'srgb', value: 'D9EAF7' },
    width: 2,
  },
  bottom: { kind: 'none' },
}; // partial whole-replace，并清除 right/left
table.borders = { kind: 'none' }; // 四边广播到全部 physical cells
table.borders = undefined; // 清除全部 direct L/R/T/B
```

`TableModel.borders` 只在全部 physical cells（包括 merge continuation）都具有同一个非空、安全、受支持的 direct complete 或 partial L/R/T/B vector 时返回 detached `TableCellBorders`。All-absent、mixed、malformed、advanced、repeated 或 ambiguous state 返回 `undefined`；它不会解析 table style、shared edge、effective border、默认值或创建期输入。Mixed 明细始终保留在 `rows[].cells[].borders`。赋值接受 scalar、精确 TRBL、partial named object、`{}` 或 `undefined`，并在单一事务中 whole-replace 全部 physical cells：scalar 写四边，TRBL/named 清除遗漏边，empty/`undefined` 清除四边。同值与全 absent clear 是 exact no-op，不安全写入以 `ModelParseError` 零 partial mutation 拒绝。

PptxGenJS 4.0.1 省略 border 时会物化 uniform four-side direct none；合法 uniform scalar/TRBL 可投影为同一个共识，合法不同 cell override 投影为 mixed `undefined`。Native bulk edit 是 final direct cell state 上的 existing-deck lossless extension。Native 继续区分 direct absence、direct none、zero-width line、omitted style 与 explicit solid；PptxGenJS 的 empty object 默认灰色 1pt solid、omitted type 物化 solid、TRBL zero 可能变 1pt、short tuple 自动补 none 等宽松行为不会被复制。

Focused 为 9 files / 567 tests；最终 full 为 1457 passed / 1 skipped tests，performance 1/1（1.20s）。实际 62-file tarball SHA-256 为 `47d9666c1dac8454524a87f7ca1898af0442c6faa7816e39cec34ff42dbf0d48`；installed Node/types/browser/CLI、`pptx-inspect` 与真实 Google Chrome 均报告 `tableBorders: true` / `tableBordersInspect: true`。最终四个 physical cells 各自恰有一组 direct `lnL/lnR/lnT/lnB` no-fill，PowerPoint 2010 validation 为 0 errors / 0 warnings，Chrome validation/console/page/network errors 均为 0。设计、计划、修正、core 与 package-proof commits 分别为 `8d6dfae`、`d4bf9c9`、`faca4ba`、`9195e57`、`f8a9d0d`；证据位于 `/tmp/pptx-table-borders-artifacts.vyi1yo`。

## 创建、读取和编辑表格单元格超链接

```ts
import { PptxDocument, TableModel } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const source = document.addSlide();
const target = document.addSlide();
const table = source.addTable([[
  { text: '官网', options: { hyperlink: { url: 'https://example.com' } } },
  { text: '详情', options: { hyperlink: { slide: 2, tooltip: '' } } },
]], { name: '单元格链接' });

console.log(table.rows[0]!.cells[0]!.hyperlink); // { url: 'https://example.com' }
console.log(table.rows[0]!.cells[1]!.hyperlink); // { slide: 2, tooltip: '' }
document.moveSlide(document.slides.indexOf(target), 0);
console.log(table.rows[0]!.cells[1]!.hyperlink); // { slide: 1, tooltip: '' }
table.setCellText(0, 1, '打开详情'); // 保留该单元格链接
table.setCellHyperlink(0, 0, {
  url: 'https://example.com/docs',
  tooltip: '查看文档',
});
table.setCellHyperlink(0, 1, undefined); // 清除 click；保留现有下划线/样式

const reopened = await PptxDocument.open(await document.write());
const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
const reopenedTable = reopenedSource.shapes.find(
  ({ name }) => name === '单元格链接',
) as TableModel;
console.log(reopenedTable.rows[0]!.cells[0]!.hyperlink);
// { url: 'https://example.com/docs', tooltip: '查看文档' }
console.log(reopenedTable.rows[0]!.cells[1]!.hyperlink); // undefined
```

`AddTableCellOptions.hyperlink` 与 `TableModel.setCellHyperlink(rowIndex, columnIndex, value)` 支持 URL 或当前文稿内部的一基页码，`tooltip` 的省略与显式 `''` 保持可区分；`undefined` 清除选中 physical cell 的 direct run click。行列索引从零开始，创建或编辑后都可立即通过只读 `TableCell.hyperlink` 获得 detached frozen direct-state snapshot。链接属于普通 plain cell 的唯一 direct run，`setCellText()` 和其他已支持的 cell 属性编辑不会移除它。内部链接按目标页 identity 保存，页面移动后 snapshot 的 ordinal 会同步变化；duplicate、delete、rollback、六格式与 write/reopen 共用既有 relationship lifecycle。

Native 创建时每个 linked cell 都独占 relationship，即使多个 cells 使用相同 URL；导入文件若共享 ID，则 target 变化只对所选 cell clone-on-write。相同值是 exact part/relationship/journal no-op；仅改 tooltip 会复用 ID，唯一 relationship 的 target 切换原位更新，clear/replace 只回收最后一个引用。添加链接只在没有 direct underline 时补 single underline；清除只删除 click，保留现有下划线和其他 run properties。当前没有 table-level hyperlink default。该 scalar reader/editor 只接受唯一 direct text body、paragraph、run、run properties 与 text，不猜测 rich/multi-run/multi-paragraph 状态；这些 cell 的 default/local run links 通过 `TableCell.richText` 读取并由 `setCellRichText()` whole-replace。

PptxGenJS 4.0.1 的合法 URL/内部页输出可读取并编辑，包括 run click 上的 `invalidUrl=""`、`action=""`、`history="1"` 等额外兼容属性；它没有 existing-deck table-cell hyperlink editor。其 omitted tooltip 会物化 empty，并会向 caller hyperlink object 写入 `_rId`；native 保留更精确的 omitted/empty 差异、从不修改输入，也不复制宽松 coercion 或悬空 relationship 行为。

最终 full Vitest 为 84 passed / 1 skipped test files、1476 passed / 1 skipped tests，performance 1/1（核心测试 1.63s）。实际 62-file tarball SHA-256 为 `2d06b955b48a25fc6f1e06accf2bd059045a7b3db04a6f3640c5bdca987ea816`；installed Node/types/browser/CLI、`pptx-inspect` 与真实 Google Chrome 均报告 `tableCellHyperlinkEditing: true`。最终证据文件为 22 parts、23 relationships、3 slides；首个表格页包含 6 cells、2 个 click、2 个匹配 relationship 与 6 个保留下划线。PowerPoint 2010 validation 为 0 errors，并只有 2 条预期 `OPC_EXTERNAL_RELATIONSHIP`，Chrome validation/console/page/network errors 均为 0。设计、计划、rollback、core 与 package-proof commits 分别为 `4fe0c43`、`0e566bd`、`dca33ba`、`99a6d3b`、`93b6b09`；证据位于 `/tmp/pptx-table-cell-hyperlink-editing-UphgYg`。

## 创建、读取和编辑表格单元格富文本

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();
const table = slide.addTable([[
  {
    text: [
      {
        align: 'center',
        runs: [{
          text: '季度摘要',
          style: {
            bold: true,
            fontSize: 18,
            color: { kind: 'scheme', value: 'accent1' },
          },
        }],
      },
      {
        runs: [
          { text: '默认链接' },
          {
            text: ' · OpenAI',
            style: { italic: true, hyperlink: { url: 'https://openai.com' } },
          },
          {
            text: '软换行且不继承链接',
            softBreakBefore: true,
            style: { hyperlink: false },
          },
        ],
      },
    ],
    options: { hyperlink: { url: 'https://example.com/report' } },
  },
  '第一行\r\n\r第三行',
]], { name: '富文本单元格' });

const cell = table.rows[0]!.cells[0]!;
console.log(cell.text);     // 段落和 soft break 均投影为 \n
console.log(cell.richText); // detached readonly RichTextParagraph[]

table.setCellRichText(0, 0, [{
  runs: [
    { text: '已替换', style: { bold: true } },
    { text: ' · 文档', style: { hyperlink: { url: 'https://example.com/docs' } } },
  ],
}]);
await document.writeFile('rich-table-cells.pptx');
```

`AddTableCell.text` 现在接受普通 string 或 `readonly RichTextParagraph[]`；富文本数组必须放在 `{ text, options? }` cell 对象中，不能把裸数组当作 cell。String 中的 CRLF/CR 会先规范化为 LF，再按段落拆分，连续或结尾换行产生的空段落会被保留；结构化输入直接表达多段落，`softBreakBefore` 表示同一段落内的软换行，`breakLine` 可在创建归一化时拆出新段落且不会出现在 snapshot。`TableCell.text` 将段落边界与 soft break 都投影为 `\n`，`TableCell.richText` 则保留 paragraph、run、alignment、list/spacing/tab 以及 font、size、bold/italic、color、underline、strike、highlight、outline、glow、baseline、character spacing、transparency 和 hyperlink 等已支持样式。

Cell `options.hyperlink` 是默认 run link：未提供局部设置的非空 runs 继承它，显式 run hyperlink 使用独立 relationship，`hyperlink: false` 抑制继承。Rich cell 的链接通过 `richText[].runs[].style.hyperlink` 读取，并用 `setCellRichText()` whole-replace；plain `TableCell.hyperlink` / `setCellHyperlink()` 仍只负责安全的单段落单 run 状态。链接目标按 slide identity 保存，replacement 复用相同目标 ID、对共享 relationship clone-on-write，并只回收最后一个引用。

`setCellRichText(rowIndex, columnIndex, value)` 使用从零开始的 physical row/cell index，保留 `bodyPr`、`lstStyle`、`tcPr`、表格 geometry、相邻 cells 与 live model identity；与当前结构和值完全相同的 replacement 是 exact bytes/relationships/journal no-op。输入与返回 snapshot 都会立即和 caller 脱离。`setCellText()` 只适合恰好一个安全 direct paragraph + run 的普通 cell，并保留该 run 的样式与链接；rich、multi-run、multi-paragraph、field 或 break cell 必须使用 `setCellRichText()`，否则在 mutation 前抛 `ModelParseError`。

PptxGenJS 4.0.1 的合法 CR/LF、`breakLine`、多段落、paragraph alignment/bullet、run font/size/bold/italic/color/underline、cell-default link 与 run-local link 输出均可读取并继续编辑。Native 使用 canonical `RichTextParagraph[]`，不修改 caller、不写重复 `pPr`，保留 omitted 与 explicit-empty tooltip 的差异，并在 mutation 前拒绝 coercible、dangling 或 descriptor-unsafe input。

最终 full Vitest 为 85 passed / 1 skipped test files、1487 passed / 1 skipped tests（143.22s），1000-part performance 为 1648ms；TypeScript project references、Node/browser bundle 与 declaration build 均通过。实际 62-file tarball SHA-256 为 `7de2354ac691ad09b58e0e103fd07ff1428caa799b548d2a65d9d19a4e0fd79f`；installed Node、NodeNext types、browser、CLI、Inspector 与 PptxGenJS 4.0.1 import/edit 均通过。最终 edited package 为 20 parts / 19 relationships，只有 slide XML 与其 relationships part 变化，2 个 click 无 dangling ID。PowerPoint 2010 全部为 0 errors，仅有预期 external-link warnings；Google Chrome 150.0.7871.188 的 create/snapshot/edit/link 均为 true，validation/console/page/network errors 均为 0。实现与复核 commits 为 `3eb6f37`、`6b40fc5`、`fd6fc43`、`d0aa76d`、`0e3c36a`；证据位于 `/tmp/pptx-table-cell-rich-text-DA3x3Z`。

## 表格与单元格文本样式默认值

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();
const table = slide.addTable([[
  '继承表格默认值',
  {
    text: [{
      spacing: { after: 10 },
      runs: [
        { text: '继承单元格默认值' },
        {
          text: ' · 局部覆盖',
          style: { fontSize: 12, bold: false },
        },
        {
          text: ' · 链接',
          style: { hyperlink: { url: 'https://example.com' } },
        },
      ],
    }],
    options: {
      fontFamily: 'Courier New',
      bold: false,
      spacing: { before: 3 },
    },
  },
]], {
  fontFamily: 'Aptos',
  fontSize: 18.25,
  bold: true,
  color: { kind: 'scheme', value: 'accent1' },
  spacing: {
    before: 6,
    after: 8,
    line: { kind: 'multiple', factor: 1.5 },
  },
});

table.setCellText(0, 0, '保留已物化的 direct run 样式');
table.setCellRichText(0, 1, [{ runs: [{ text: '全量替换，不重新继承创建默认值' }] }]);
```

`AddTableOptions` 与 `AddTableCellOptions` 现在都公开 `fontFamily`、`fontSize`、`bold`、`color` 和 `spacing`。创建优先级固定为 table → cell → explicit paragraph/run；字体、字号、粗体与颜色逐字段覆盖，spacing 则按 `before`、`after`、`line` 子字段合并。显式 `bold: false`、paragraph `spacing: false` 与 `line: false` 会阻断对应继承。Font family 必须是非空 XML-safe string，font size 使用 strict 1–4000pt，bold 只接受 boolean，颜色与 paragraph spacing 复用既有 strict native value。Native 名称固定为 `fontFamily` 和 structured `spacing`，不接受 PptxGenJS 的 `fontFace`、`paraSpaceBefore`、`paraSpaceAfter`、`lineSpacing` 或 `lineSpacingMultiple` alias。

Resolved 值只物化为每个 physical cell 的 direct paragraph/run OOXML，不保留 table/cell 创建 metadata。`TableCell.richText` 可立即读取最终 direct state；`setCellText()` 保留安全 plain run 的当前样式模板，`setCellRichText()` whole-replace 后不会重新应用创建默认值。空段落的 `endParaRPr` 会携带 resolved font family/font size。Cell-default hyperlink 保留 outer color；run-local hyperlink 若没有 explicit run color，则不继承 outer color，显式 run color 始终优先。

PptxGenJS 4.0.1 的合法 table/cell `fontFace`、font size、bold、color、cell paragraph spacing、rich override、empty paragraph 和 hyperlink final state 均可导入、编辑并重开。Native 另外支持 table-level spacing 传播，并修正其 truthy fallback 会覆盖 cell/run `bold: false` 以及 writer 修改 caller options 的行为。表格合并、physical row/column CRUD 与 explicit-row-height auto-page/repeated headers 已在后续专项完成；当前总体 PptxGenJS 对等进度约 99.3%，剩余 automatic content measurement/layout recomputation、`tableToSlides` 与最终 peer/client audit。

最终 full Vitest 为 85 passed / 1 skipped test files、1497 passed / 1 skipped tests（167.50s），1000-part performance 为 1565ms；TypeScript、Node/browser bundles 与 declarations 均通过。实际 62-file tarball SHA-256 为 `79ed789e6d4f218cc5c838af9e5965e96bd7e35f132d2a630a85ac5dd39ed222`；installed Node、NodeNext types、browser conditional export、CLI 与 Inspector 均报告 table text defaults 通过。最终 evidence deck 为 18 parts / 15 relationships、1 slide / 1 table / 3 cells，PowerPoint 2010 为 0 errors / 0 warnings；Google Chrome 150.0.7871.188 的 create/snapshot/plain edit/rich replacement/reopen 均为 true，validation/console/page/network errors 均为 0。实现、复核与发布证明 commits 为 `e8cd0c7`、`0fc1567`、`af4e419`、`d6f3fd9`、`6e5df9a`、`145148b`、`2eb1a5f`；证据位于 `/tmp/pptx-table-text-defaults-proof.ViSdTX`。

## 创建、读取和编辑表格单元格合并

```ts
const table = slide.addTable([
  [{ text: 'Summary', options: { colspan: 2, rowspan: 2 } }, 'Total'],
  ['42'],
]);

console.log(table.mergeRegions);
table.unmergeCell(1, 1); // region 中任意 physical cell 都可定位整个合并区域
table.mergeCells(0, 0, 2, 2);
```

创建输入使用 logical rows。首行所有 cell 的 `colspan` 总和定义 physical column count；后续行会跳过仍被上方 `rowspan` 占用的列，因此完全被覆盖的行可以写成 `[]`。`colspan` / `rowspan` 只接受正 safe integer，省略或 `1` 表示该维不合并。输入会在 package mutation 前完成展开和校验；空洞、重叠、越界、非矩形或超过安全上限的布局直接拒绝。

读取和编辑统一使用从零开始的 physical row/column coordinates。`TableModel.mergeRegions` 对合法未合并表返回 `[]`，对合法合并表返回按 row-major 排序的 detached deep-frozen `TableMergeRegion[]`，对不安全或歧义拓扑返回 `undefined`；每个合法区域内的 physical `TableCell.merge` 都指向同一 anchor，并以 `isAnchor` 区分左上角。`mergeCells()` 要求至少覆盖两个 cells，拒绝与现有区域发生非完全相同的相交；`unmergeCell()` 可接收 anchor 或任意 continuation。

合并与拆分只修改 physical cell start tag 的 `rowSpan`、`gridSpan`、`vMerge`、`hMerge`。Continuation 中原有的隐藏文本、样式、关系和未知 XML 均被保留，拆分后重新可见；既有 cell editor 仍可按 physical coordinate 显式编辑这些状态。重复同一合并、拆分未合并 cell 都是 exact bytes/relationships/journal no-op；malformed 或不能安全识别的拓扑保留原 bytes，并拒绝语义编辑。

PptxGenJS 4.0.1 的合法 horizontal、vertical、rectangular 和 offset span 输出可读取、编辑并由 native 创建为相同最终合并语义。Native 不复制其 lopsided non-span row、negative/fractional span 或 out-of-bounds rowspan 可产生非法 OOXML 的缺陷，而是在任何可观察 mutation 前严格拒绝。

最终 focused gate 为 5/5 test files、594/594 tests（28.11s）；全量为 86 passed / 1 skipped test files、1512 passed / 1 skipped tests（73.26s），独立 1000-part performance 为 1/1（709ms）。TypeScript project references、Node/browser bundles 与 declarations 全部通过。两次构建的 59-file dist manifests 完全一致，两份 62-file actual tarball byte-identical，SHA-256 均为 `0c85afa9bed6a04faa5d3dab6934a3974cea731091dc673ab2ff6e92cb83343d`。Installed Node、NodeNext types、browser conditional export、CLI 与 Inspector 均报告 `tableCellMerges: true` / `tableCellMergesInspect: true`。

真实 Google Chrome 150.0.7871.188 的 create/read/frozen snapshot/unmerge/edit/remerge/reopen 全部为 true，validation/console/page/network errors 均为 0。Browser evidence deck 为 18 parts / 15 relationships、1 slide / 1 table、2×3 physical cells 与 1 个 2×2 merge region；四种 anchor/continuation token 均存在，slide relationship 只有合法 layout owner，PowerPoint 2010 validation 为 0 errors / 0 warnings。识别、创建、snapshot/editor、SDK/adapter、文档与 package proof commits 为 `688f9f6`、`3d93f07`、`db01937`、`b2f6846`、`5832399`、`7073eae`、`f174519`；完整证据位于 `/tmp/pptx-table-cell-merges-artifacts.B7ZhGQ`。Physical row/column CRUD 与 explicit-row-height auto-page/repeated headers 已在下节完成；当前总体 PptxGenJS 对等进度约 99.3%，之后依次是 automatic content measurement/layout recomputation、`tableToSlides` 与最终 peer/client audit。

## 插入和删除表格行列

```ts
const table = slide.addTable([
  ['A0', 'A1', 'A2'],
  ['B0', 'B1', 'B2'],
  ['C0', 'C1', 'C2'],
], {
  columnWidths: [inches(1), inches(2), inches(3)],
  rowHeights: [inches(0.5), inches(1), inches(1.5)],
});

table.mergeCells(0, 0, 2, 2);
table.insertRows(1, { count: 2, rowHeights: [inches(0.25), inches(0.5)] });
table.insertColumns(1, { columnWidths: inches(0.75) });
table.setCellText(1, 1, '新增的隐藏 continuation');
table.deleteRows(4);
table.deleteColumns(3);
```

`insertRows()`、`deleteRows()`、`insertColumns()` 和 `deleteColumns()` 统一使用从零开始的 physical coordinates，与 `rows[].cells[]`、`mergeRegions` 和现有 cell editor 一致。Insert index 可以等于当前数量以 append；delete 必须指向现有项，且不能删除最后一行或最后一列。`count` 默认为 1，只接受正 safe integer；结构变化后的 physical cell 总量上限为 1,000,000。

`rowHeights` 接受非负 EMU scalar 或长度恰好等于 `count` 的 dense array，其中 0 表示自动行高；`columnWidths` 接受正 EMU scalar 或 exact-length array。省略尺寸时，中间插入复制 insertion index 处的 direct size，append 复制最后一项。Column CRUD 总是让 transform width 等于 grid width 总和；row CRUD 只在所有行高都大于 0 时同步 transform height，存在自动行高时保留现有 height，不声称进行内容测量或 layout recomputation。

在 merge anchor 坐标插入表示在区域之前插入；严格插入区域内部会扩展 rowspan/colspan。删除会收缩区域，在原 anchor 被删除时提升最上/最左的 survivor，并在区域退化为 1×1 时解除合并。新 cell 是可立即由 text/rich-text/hyperlink/style editor 填充的 canonical empty plain cell，不复制相邻内容或样式。幸存 cell 的 source bytes、隐藏内容、样式、relationship 与未知 XML 保持；删除只回收整张 slide 中最后引用已经消失的 relationship。四个方法与 relationship GC 都在同一个 package transaction 中执行。

PptxGenJS 4.0.1 只有创建期 table rows、`rowH` / `colW` 和 auto-page helper，没有 existing-deck row/column editor。合法 PptxGenJS plain/rich/linked/merged/sized table 可以通过 `importPptxGenJS()` 导入后使用上述 native lossless CRUD。Explicit-row-height auto-page 与 repeated headers 已在下节支持；logical content insertion、automatic content measurement/layout recomputation 与 `tableToSlides` 仍未支持。

最终 focused gate 为 5 个文件 / 611 项测试（29.96s）；全量为 87 passed / 1 skipped test files、1535 passed / 1 skipped tests（64.27s），独立 1000-part performance 为 1/1（核心测试 1204ms，test file 1207ms，total 2.52s）。TypeScript project references、root build、Node/browser bundles 与 declarations 全部通过。两次 clean build 的 59-file dist manifests 完全一致，manifest SHA-256 为 `51d0c19da69fbd81682933d4a5418ff58ef2a805b4164d624f150d1674924e41`；两份 62-file、660,178-byte actual tarball byte-identical，SHA-256 均为 `17d43a887a9871fd4910bcf33415d985b4d8f1968b4020670a64166c148aeaa4`。Installed Node、NodeNext declarations、browser conditional export、CLI 与 Inspector 均报告 `tableStructureEditing: true`。

真实 Google Chrome 150.0.7871.188 的 create/row insert/column insert/new-cell edit/row delete/column delete/dimensions/merge/survivor/relationships/reopen 全部为 true，validation/console/page/network errors 均为 0。Node 与 browser evidence deck 均为 18 parts / 16 relationships、1 slide / 1 table、4×4 physical matrix；column widths `[914400, 457200, 1828800, 2743200]` 与 transform width 同为 5943600，row heights `[457200, 228600, 914400, 1371600]` 与 transform height 同为 2971800。最终 3×3 merge 含 1 anchor、2 top、2 left、4 interior continuation，隐藏 inserted text 与 styled survivor 均存在；两个 click 共用 `rId2`，只保留 1 个 external relationship，orphan hyperlinks 为 0。PowerPoint 2010 validation 为 0 errors / 1 个预期 `OPC_EXTERNAL_RELATIONSHIP` portability warning。实现、契约、文档与 package proof commits 为 `ee68731`、`d70c2af`、`099f345`、`89f9b1b`、`2250826`、`1ab602f`；完整证据位于 `/tmp/pptx-table-structure-editing-proof.S1rVAZ`。CRUD 专项 8/8 完成；explicit-row-height auto-page/repeated headers 已在下节完成。

## 跨幻灯片自动分页与重复表头

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
document.addSlide();
const source = document.addSlide();
const following = document.addSlide();

source.addTable([
  ['Region', 'Revenue'],
  ['Unit', 'USD'],
  ['North', '$120'],
  ['South', '$95'],
  ['East', '$110'],
  ['West', '$105'],
], {
  autoPage: true,
  autoPageRepeatHeader: true,
  autoPageHeaderRows: 2,
  autoPageSlideStartY: inches(0.75),
  slideMargin: [inches(0.5), inches(0.4), inches(0.5), inches(0.4)],
  y: inches(5),
  columnWidths: [inches(3), inches(2)],
  rowHeights: [
    inches(0.3),
    inches(0.3),
    inches(0.55),
    inches(0.55),
    inches(0.55),
    inches(0.55),
  ],
});

console.log(source.newAutoPagedSlides.length); // 1
console.log(document.slides.indexOf(following)); // 3，generated slide 插在 source 与 following 之间
```

当前支持的是 deterministic structural pagination：`autoPage: true` 必须同时提供每个 physical row 的正 EMU `rowHeights`。Source page 从 table 自身的 `y` 开始；continuation page 从 `autoPageSlideStartY`、当前 runtime named-layout top margin 或 canonical 0.5 英寸 margin 开始，并在 bottom margin 前结束。显式 `slideMargin` 可为单个非负 EMU 或严格 `[top, right, bottom, left]` 四元组，并优先于 layout margin。

`autoPageRepeatHeader` 为 true 时，`autoPageHeaderRows` 默认为 1，也可指定多个 header rows。Rowspan block 不会被拆开，且 merge 不得跨越 header/body 边界。每张输出页都是普通、可继续编辑并可重开的 table；rich text、cell styles、URL/internal-slide hyperlinks、column widths、merge tokens 与 relationship ownership 都按页保留。内部页链接会在插入 continuation slides 前解析，所以目标 identity 不会因页序变化而漂移。

Continuation slides 总是紧邻 source 连续新建，使用同一 layout 并保持 source 的 section membership；不会复用或修改原本的 following slide。Source table、所有生成页、relationships、presentation order、sections 与 runtime state 位于同一 transaction，失败时整体恢复。

`source.newAutoPagedSlides` 是最近一次成功 `addTable()` 产生的 continuation slides 的 frozen readonly runtime snapshot，不包含 source。成功的普通 table 或无溢出 auto-page 会把它重置为空；失败调用保留上一次成功结果；删除已生成页后 getter 会过滤 detached identity；duplicate 与 write/reopen 后均为空，因为该 metadata 不写入 PPTX。

当前尚不支持省略/automatic `rowHeights`、`autoPageCharWeight`、`autoPageLineWeight`、超高文字行 fragmentation、placeholder auto-page、content measurement/layout recomputation 或 `tableToSlides`。这些输入会被拒绝，而不是静默估算。Native 也不复制 PptxGenJS 4.0.1 修改 caller objects、夹取权重、coerce 非法值或复用既有后续 slide 的行为。当前总体 PptxGenJS 对等进度约 99.3%，仍不声明完整 parity。

最终 focused gate 为 7 个文件（6 passed / 1 skipped）、39 passed / 629 skipped，Vitest 5.44s；full gate 为 89 个文件（88 passed / 1 skipped）、1579 passed / 1 skipped，Vitest 40.17s；独立 1000-part performance gate 为 1/1，核心测试 682ms、Vitest 1.89s。TypeScript typecheck、root build 与 `@jiayunxie/pptx` package build 分别为 1.40s、1.10s、7.00s。两次 59-file dist manifest 完全一致，manifest SHA-256 为 `f5fd12f308fc360fb49edd0c4d35e4e43aaf5e3bd7ad1ed3c3caebe9d4a25a8e`；两份 62-file actual tarball byte-identical，SHA-256 均为 `3db5ca1dcf61b81ac6072639e342d8b5e1bbca9e35c2025476cc9f4d781432d3`。Installed Node、NodeNext declarations 与 browser conditional export 均报告 `tableAutoPage: true`，installed CLI/Inspector 另报告 `tableAutoPageInspect: true`。

Google Chrome 150.0.7871.188 的 create/edit/move/delete/relationship/reopen 阶段全部为 true，validation/console/page/network errors 均为 0。Node 与 browser evidence deck 均为 26 parts / 32 relationships、5 slides / 2 tables，生成阶段把 7 行分为 4/4/3 三页并重复两行表头；三页保持同一 layout 与连续 section membership，header/body merge、rich style 与目标 slide identity 均保留。最终 move/edit/delete 后两张 table slide 都是 4×3 physical cells、row height 457200、column width 914400、transform 2743200×1828800；source 的 4 个 click 对应 3 external + 1 internal relationship，continuation 的 5 个 click 对应 3 external + 2 internal relationships，且全部 page-local IDs 精确匹配。Node/browser deck SHA-256 分别为 `8665affbdbc2e96eecb7683770685d08cf6761ee4d12c4d72ee057cea9943cb7` 与 `a6a8b4458af12dd95635e6ea3443346910de07cbbd5a91ee68b5f86e8d82e862`；两者 PowerPoint 2010 validation 均为 0 errors / 6 个预期 `OPC_EXTERNAL_RELATIONSHIP` warnings。Commit chain 为 `0b63c07`、`a52708a`、`062e0b3`、`d30a4e6`、`3ce9cc7`、`d66f884`、`c377799`、`0ab580f`，证据保留在 `/tmp/pptx-table-auto-page-proof.IQxIFi`。本专项 8/8 完成；总体 PptxGenJS 对等进度约 99.3%，下一项是 content measurement/layout recomputation，之后为 `tableToSlides` 与最终 peer/client audit。

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

`AddShapeOptions.shadow`、`AddTextOptions.shadow` 与 `ShapeModel.shadow` 支持 preset/text shape direct outer/inner shadow 的创建、读取、whole replacement 与清除，包括 sRGB/theme color、`0..1` opacity、`0..100pt` blur、`0..<360°` angle、`0..200pt` distance，以及 outer-only `rotateWithShape`。默认值为 black、0.75、8pt、270°、4pt 和 outer rotate false；显式 zero 会保留。输入在 mutation 前深度脱离，getter 的嵌套快照会 deep-freeze；同值赋值是 exact no-op，`undefined` 只移除 direct shadow 并保留 `effectLst` 与 glow/reflection 等 sibling effects。Generic/advanced effects、custom shadow transforms，以及 image/table/chart/media 等其他 owner 的 shadow API 仍待后续小项。

`AddShapeOptions.hyperlink`、`AddTextOptions.hyperlink` 与 `ShapeModel.hyperlink` 支持整个 preset/text shape 的 click URL 或内部页链接。输入必须恰好包含一个非空 `url` 或一个当前文稿内的一基 `slide`；`tooltip` 可省略，也可显式为空。Getter 返回 detached frozen snapshot，setter 采用 whole replacement，同值赋值为 exact no-op，`undefined` 清除 click link。内部关系按目标页 identity 保存，移动或在目标前插删页面只更新 getter ordinal；复制 self-link 会指向副本自身，删除目标页会清理相关 click/hover，shared relationship 则按引用 clone-on-write 与回收。Text outer 与 `RichTextRunStyle.hyperlink` 分别管理 whole-shape/default run 和显式 run-local 链接，ownership 相互独立；run hyperlink 可与 `RichTextRun.breakLine` 组合并按规范段落重新索引。外部链接产生 validator 的预期可移植性 warning。Plain single-run table-cell scalar hyperlink API 与 rich/multi-paragraph cell default/local run links 均已支持；hover、table graphic-frame/image/chart/media 链接、action navigation、advanced line fill/custom dash 和 percentage positions 仍待后续小项；`isTextBox` 与 rich-text `breakLine` 已完成。

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
