# @jiayunxie/pptx

Lossless bidirectional PPTX OOXML editing for Node.js and TypeScript.

## Install

This release is a technical preview published under the `next` tag.

```sh
npm install @jiayunxie/pptx@next
```

## Create a presentation

```ts
import {
  degrees,
  evaluateCustomGeometry,
  inches,
  PRESET_SHAPE_TYPES,
  PptxDocument,
} from '@jiayunxie/pptx';

const document = PptxDocument.create({
  author: 'Presentation Team',
  company: 'Acme & Partners',
  createdAt: '2024-02-29T12:34:56.123+05:30',
  lastModifiedBy: 'Presentation Team',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
  revision: '7',
  rtlMode: true,
  slideSize: { width: inches(11.7), height: inches(8.3) },
  subject: 'Revenue & Forecast',
  title: 'Quarterly Review',
});
document.author = 'Updated Author';
document.company = 'Updated Company';
document.company = '';
document.company = undefined;
document.createdAt = '2026-07-30T00:00:00Z';
document.createdAt = undefined;
document.lastModifiedBy = 'Updated Editor';
document.lastModifiedBy = '';
document.lastModifiedBy = undefined;
document.modifiedAt = '2026-07-30T01:02:03Z';
document.modifiedAt = undefined;
document.subject = 'Updated Subject';
document.subject = '';
document.subject = undefined;
document.revision = '008';
document.revision = undefined;
document.title = 'Updated Review';
const slide = document.addSlide();
const shape = slide.addShape('roundRect', {
  x: inches(1),
  y: inches(1),
  width: inches(3),
  height: inches(2),
  rotation: degrees(15),
  name: 'Feature card',
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
    tooltip: 'Open documentation',
  },
});
shape.arrows = { begin: 'diamond' };
shape.arrows = { begin: 'none', end: 'oval' };
shape.arrows = undefined;
shape.shadow = {
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0.5,
  blur: 3,
  angle: 270,
  distance: 2,
};
shape.shadow = undefined;
document.addSlide(); // creates slide 2 as an internal-link target
shape.hyperlink = { slide: 2, tooltip: 'Go to details' };
shape.hyperlink = { url: 'mailto:team@example.com', tooltip: '' };
shape.hyperlink = undefined;
const arc = slide.addShape('blockArc', {
  adjustments: [
    { name: 'adj1', value: 16_200_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 25_000 },
  ],
});
arc.adjustments = [{ name: 'adj1', value: 10_800_000 }];
arc.adjustments = [];
shape.presetType = 'hexagon';
shape.fill = { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } };
shape.fill = { kind: 'none' };
shape.fill = undefined;
shape.line = { kind: 'line', color: { kind: 'scheme', value: 'accent2' } };
shape.line = { kind: 'none' };
shape.line = undefined;
console.log(shape.presetType); // 'hexagon'
console.log(PRESET_SHAPE_TYPES.length); // 178
slide.addText('Quarterly results\nQ4 forecast', {
  x: inches(1),
  y: inches(1),
  width: inches(6),
  height: inches(1),
  align: 'center',
  fit: 'shrink',
  lang: 'en-US',
  paragraphIndent: 12,
  paragraphMarginLeft: 18,
  paragraphMarginRight: 18,
  valign: 'middle',
  vert: 'vert270',
  wrap: true,
});
slide.addRichText([
  { indent: -12, runs: [{ text: 'مرحبا' }] },
  { indent: 18, marginLeft: 12, marginRight: 12, rtl: false, runs: [{ text: 'English override' }] },
  { bullet: true, indent: false, marginLeft: false, runs: [{ text: 'List-owned indent' }] },
  { runs: [{ text: 'Quarter transparent', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 25 } }] },
], { paragraphIndent: 12, paragraphMarginLeft: 24, paragraphMarginRight: 24, rtlMode: true });
const table = slide.addTable([
  [{ text: 'Region', options: {
    align: 'center',
    fit: 'shrink',
    textDirection: 'vert',
    valign: 'top',
    margin: { top: 4, left: 8 },
    border: { kind: 'line', color: { kind: 'scheme', value: 'accent2' }, width: 1, style: 'dash' },
    fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
  } }, { text: 'Revenue', options: { fill: { kind: 'none' } } }],
  ['East', { text: '$1.2M' }],
  [{ text: 'West' }, '$980K'],
], { name: 'Revenue table', x: inches(1), y: inches(3), height: inches(2), columnWidths: [inches(2.5), inches(3.5)], rowHeights: [inches(0.5), inches(0.75), inches(0.75)], align: 'left', border: { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'solid' }, fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent4' }, transparency: 20 }, margin: { top: 9, left: 18 }, textDirection: 'vert270', valign: 'middle' });
table.setColumnWidths([inches(2), inches(4)]);
table.setRowHeights([inches(0.5), inches(1), inches(0.5)]);
table.setCellVerticalAlignment(0, 0, 'bottom');
table.setCellBorders(0, 0, { bottom: { kind: 'line', color: { kind: 'srgb', value: '70AD47' }, width: 2 } });
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } });
await document.writeFile('created.pptx');
```

## Load and create embedded raster images

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
const imageDocument = PptxDocument.create();
imageDocument.addSlide();
const image = await imageDocument.addImage(0, 'chart.png', options);
const info = inspectRasterImage(new Uint8Array(await readFile('chart.png')));
console.log(info); // { contentType: 'image/png', width, height }
console.log(calculateRasterImageSizing(info, sizing));
image.setTransform({ x: inches(1.5) });
image.sourceRectangle = { left: 10, top: -5, right: 5, bottom: 0 };
image.sourceRectangle = undefined;
image.replaceData(new Uint8Array(await readFile('chart-updated.png')), 'image/png');
await imageDocument.writeFile('images.pptx');
```

`PptxDocument.addImage()` accepts Node file paths, HTTP/HTTPS URLs, browser-relative URLs, strict base64 data URIs, `Uint8Array`, `ArrayBuffer`, `Blob`/`File`, Web `ReadableStream`, and async byte iterables. Byte signatures are the only source of truth for PNG/JPEG/GIF format; optional `AddImageSourceOptions.contentType` is an assertion rejected before package mutation when it disagrees. `signal` aborts file, Fetch, Blob, and stream loading. File extensions, `Blob.type`, HTTP `Content-Type`, and `File.name` are not trusted for format detection.

`inspectRasterImage()` returns the canonical content type plus raw pixel width and height detected from the signature and format structure. Pure `calculateRasterImageSizing()` accepts an EMU target frame and returns the same transform size plus a four-edge source rectangle for `contain`, `cover`, or a source-pixel `crop` region. High-level `sizing` is mutually exclusive with top-level `width`/`height`; when sizing is omitted, the loader retains the PptxGenJS-compatible one-inch default transform and does not infer layout size from intrinsic pixels.

The synchronous lower-level `SlideModel.addImage(bytes, { contentType, sourceRectangle, ... })` remains available when callers already hold strict bytes and a canonical content type. `ImageSourceRectangle.left/top/right/bottom` use percent units where `1` means 1%, quantized to 0.001%; negative values extend the visible source area for contain. Live `ImageModel.sourceRectangle` returns a detached frozen snapshot, whole-replaces on assignment, and clears direct `a:srcRect` when assigned `undefined`. High-level `PptxDocument.addImage()` rejects direct `sourceRectangle`; callers use `sizing`. Options detach before asynchronous source I/O, and sizing finishes before package mutation.

Creation atomically owns one unique media part, one internal image relationship, and one canonical rectangular picture with aspect lock and stretch fill. Validation or write failure rolls back the part, content type, relationship, slide XML, and mutation journal. Slide duplication initially shares media targets; `ImageModel.replaceData()` updates an exclusive target in place and redirects a shared shape to a private clone. Creation, transform editing, replacement, duplication, rollback, write/reopen, and all six presentation formats are covered.

Valid PptxGenJS 4.0.1 public path/data PNG/JPEG/GIF output reaches the same supported loader semantics, and six contain/cover/equal-ratio/crop cases match final transforms and direct source rectangles exactly. The actual npm tarball passes Node, browser, declaration, and CLI smoke; two clean builds produce identical SHA-256 manifests for all 38 dist files. The sizing gallery has four slides, 40 shapes, and 12 images. Source and LibreOffice round-trip packages strictly reopen, validate with 0 errors and 0 warnings, report zero overflow, and were reviewed slide by slide.

LibreOffice preserves 12/12 payload hashes, content types, names, non-empty alt text, ordering, and internal relationships; it deduplicates 12 repeated payload targets to three and rewrites all 12 picture blocks. Its largest transform quantization is 360 EMU, its largest source-rectangle quantization is 0.007%, and two flips normalize equivalently to an added 180-degree rotation. The source renders at 2400×1350 at 180 DPI. LibreOffice normalizes the round-trip page width, so its direct 180-DPI raster is 2401×1350; the inspected proportional raster is 2400×1350.

## Create, load, and edit SVG images

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
console.log(calculateImageSizing(svgInfo, {
  type: 'cover',
  width: inches(6),
  height: inches(4),
}));

const svgDocument = PptxDocument.create();
const svgSlide = svgDocument.addSlide();
const svgImage = await svgDocument.addImage(0, svgBytes, {
  contentType: 'image/svg+xml', // optional assertion
  fallback: fallbackPng,
  name: 'Architecture',
  altText: 'System architecture diagram',
  sizing: { type: 'contain', width: inches(6), height: inches(4) },
});
console.log(svgImage.isSvg); // true
console.log(svgImage.sourcePartUri === svgImage.fallbackPartUri); // true
console.log(svgImage.svgPartUri); // /ppt/media/imageN.svg

const lowLevelSvg = svgSlide.addSvgImage(svgBytes, fallbackPng, {
  x: inches(6.5),
  width: inches(4),
  height: inches(3),
});
lowLevelSvg.replaceSvgData(
  new Uint8Array(await readFile('architecture-updated.svg')),
  new Uint8Array(await readFile('architecture-updated.png')),
);
await svgDocument.writeFile('svg-images.pptx');
```

SVG uses the same path, URL, data URI, bytes, Blob/File, Web Stream, and async-iterable sources as raster images. `inspectImage()` and `inspectSvgImage()` return canonical `image/svg+xml` plus intrinsic dimensions, and generic `calculateImageSizing()` supports contain, cover, and crop. The optional `contentType: 'image/svg+xml'` is an assertion; strict SVG XML inspection remains the source of truth.

Every SVG picture owns an SVG part and a PNG fallback part. The high-level API verifies the fallback's actual PNG signature; priority is an explicit caller-supplied PNG, browser Canvas rasterization, then a built-in transparent PNG. The synchronous low-level `addSvgImage()` copies non-empty bytes and leaves declared-payload correctness to its caller. Supply a high-fidelity PNG when clients that consume only the fallback must preserve the full appearance. The library does not execute SVG scripts, fetch SVG external resources, or promise arbitrary cross-client SVG rasterization fidelity.

`ImageModel.isSvg` becomes true only for a safe, unambiguous namespace/relationship/part pair. In that state, `sourcePartUri` and `fallbackPartUri` identify the PNG while `svgPartUri` identifies the vector payload. `replaceData()` rejects SVG pictures; `replaceSvgData(svgBytes, fallbackPngBytes)` atomically replaces both sides. Duplicated slides initially share both targets, and replacement clone-on-writes both sides for only the selected picture. Malformed pair state or any failed replacement leaves package parts, relationships, XML, identity, and the mutation journal unchanged.

Three PptxGenJS 4.0.1 public conformance cases cover data-contain, path-cover, and data-crop, matching picture structure, transform, direct `srcRect`, extension URI/namespace, relationship roles, and SVG bytes. Native intentionally fixes PptxGenJS's invalid path-SVG fallback, which stores SVG bytes in a `.png` part: every native fallback validates as PNG.

The actual tarball passes Node, browser, declaration, and CLI SVG smoke, and two clean builds produce identical SHA-256 manifests for all 38 dist files. The five-slide gallery contains 13 shapes, eight SVG pictures, seven SVG parts, seven PNG fallbacks, and 16 image relationships. Source and LibreOffice round-trip packages strictly reopen and validate against PowerPoint 2010 with 0 errors and 0 warnings. LibreOffice preserves shape order, names, alt text, SVG hashes, relationship roles, and all 7+7 targets, while normalizing `image/svg+xml` to `image/svg`; maximum position/size quantization is 360 EMU and maximum `srcRect` quantization is 0.003%. After save it may render the PNG fallback, so fallback fidelity remains observable client behavior.

External SVG relationships, SVG DOM-level editing, image rounding/transparency, alt-text editing, hyperlinks/shadows, advanced placeholder styling, and public per-image deletion/media garbage collection remain pending. Embedded media creation is documented next.

## Create embedded audio and video

```ts
import { readFile } from 'node:fs/promises';
import { inches, PptxDocument } from '@jiayunxie/pptx';

const mediaDocument = PptxDocument.create();
mediaDocument.addSlide();
const poster =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAACXBIWXMAAAABAAAAAQBPJcTWAAAAEElEQVR4nGP8ywACLGCSAQANEQED1LYyQAAAAABJRU5ErkJggg==';

const audio = await mediaDocument.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
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
await mediaDocument.addVideo(0, videoBytes, {
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
await mediaDocument.writeFile('media.pptx');
```

`PptxDocument.addAudio()` and `addVideo()` accept Node paths, strict base64 data URIs, `Uint8Array`, `ArrayBuffer`, Blob/File, Web `ReadableStream`, and async byte iterables. Supported audio is `audio/mpeg` (`.mp3`), `audio/mp4` (`.m4a`), `audio/wav` (`.wav`), and `audio/ogg` (`.ogg`). Supported video is `video/mp4` (`.mp4`/`.m4v`), `video/quicktime` (`.mov`), and `video/webm` (`.webm`). Posters support `image/png`, `image/jpeg`, and `image/gif`; omission uses a built-in PNG. HTTP/HTTPS media remains external and is never fetched, while an HTTP/HTTPS poster is rejected.

Creation returns a live `MediaModel`; within one document, `document.media(slideIndex)`, `slide.media`, and `slide.shapes` return the same object. `name`, `altText`, `settings`, and the inherited transform are editable. `replaceSource()` preserves audio/video kind and object identity across embedded↔external transitions. `replacePoster()` accepts PNG/JPEG/GIF, and an omitted source resets the built-in PNG. Source replacement accepts only `contentType`, `fileName`, and `transcode`; poster replacement accepts only `contentType` and `fileName`. Both return the original model. Remove media with `media.remove()` or `slide.deleteMedia(shapeId)`.

Descriptor priority is explicit `contentType` assertion, data-URI declaration, a known `fileName` or path/`File.name` extension, then the audio/video default. Conflicting assertions, declarations, or known extensions reject before package mutation. An unknown extension is not format evidence and the selected MIME receives its canonical extension. Data URIs require a supported MIME plus standard, fully padded canonical base64; whitespace, URL-safe alphabets, percent encoding, and noncanonical padding bits are rejected. `Blob.type` is not trusted.

Options and memory byte sources detach before asynchronous work. Paths, Blobs, streams, optional transcoding, posters, descriptor resolution, and hashing all finish before one synchronous package transaction. Any create, source/poster replacement, or delete failure leaves parts, content types, relationships, slide XML, ZIP state, model identity, and the mutation journal unchanged. Payloads deduplicate only when both SHA-256 and exact MIME match. Duplicates initially share targets; the first different edit uses clone-on-write and retargets only the selected picture. Object and slide deletion collect only targets with no incoming package-graph reference.

Creation emits canonical `a:audioFile` or `a:videoFile`, kind/Microsoft-media/poster relationships, the media click action, a rectangular poster picture, and executable native PowerPoint `p:timing`. `play: 'click' | 'auto'`, `loop`, `hideWhenStopped`, and `volume: 0..1` are synchronized between exact preference/ownership state in `px:playback` and native `cMediaNode` plus play/pause commands; no plugin is required. Whole-setting replacement updates both atomically. `settings = undefined` removes only that media's preference extension and library-owned native graph while preserving ordinary animation and unknown timing branches.

Effective reads use a valid `px:playback` preference first, otherwise a strict unique/direct/complete native media graph, otherwise empty settings. This supports native-only PowerPoint imports and safe adoption on the first edit. Only an unambiguous graph is owned; unsupported or ambiguous imports remain byte-preserved and unsafe edits are rejected. Create, set, clear, duplicate, delete, target repair, and page-wide timing-ID allocation share one OPC transaction. Diagnostics distinguish missing, stale, unsupported, ambiguous, dangling-target, and kind-mismatch states through the six `MEDIA_TIMING_*` codes.

Four of four valid public PptxGenJS 4.0.1 data/path, audio/video, cover, `extn`, `objectName`, and transform cases match final semantics. The reader accepts its audio `a:videoFile`, `audio/mp3`, and duplicate-audio relationship defects. Read-only, metadata, settings, and transform work preserves those legacy primary roles; `replaceSource()` canonicalizes only the edited picture. Native creation always uses `a:audioFile`, canonical `audio/mpeg`, and the standard audio relationship.

The actual 45-file npm tarball passes Node, real-Chrome browser, declaration, and installed-CLI smoke with `nativeMediaTiming: true`, including clear, duplicate isolation, delete, unique IDs, target isolation, diagnostics, and reopen. Two clean builds produce identical SHA-256 manifests for all 42 dist files. All six presentation formats pass timing create/edit/duplicate/delete/reopen. A nine-slide playable gallery contains 12 media objects, seven media MIME types, three poster MIME types, ten deduplicated `/ppt/media` parts, and zero orphans. It strictly reopens, renders at 180 DPI with no overflow, and passed slide-by-slide visual inspection. PowerPoint 2010 validation has 0 errors and only the expected OGG/WebM warnings.

LibreOffice 26.8 preserves the nine-slide order and text but removes every media object, poster, media relationship, and timing branch when it saves the gallery; the package still strictly reopens and validates with 0 errors and 0 warnings. This is a documented client degradation, not a native round-trip guarantee. Local PowerPoint 16.112 automation returned the same `-9074` for the gallery, the LibreOffice output, and a minimal control file, so no PowerPoint round-trip pass is claimed from this environment.

Trim/bookmarks, finite repeats, narration/cross-slide audio, captions/subtitles, online video, remote-fetch embedding, crop/rounding/shadow/hyperlink and advanced placeholder styles, a built-in transcoding engine, and broad PowerPoint/Keynote/Google Slides certification remain pending.

## Create and semantically edit native charts

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const chartDocument = PptxDocument.create({ slideSize: 'wide' });
const chartSlide = chartDocument.addSlide();
const revenue = [{
  name: 'Revenue',
  categories: ['Q1', 'Q2', 'Q3'],
  values: [100, 130, 160],
}];

const chart = await chartSlide.addChart('bar', revenue, {
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

await chartSlide.addChart([
  { type: 'bar', series: revenue },
  {
    type: 'line',
    axis: 'secondary',
    series: [{ name: 'Margin', categories: ['Q1', 'Q2', 'Q3'], values: [24, 28, 31] }],
  },
], { x: inches(1), y: inches(1), width: inches(8), height: inches(4.5) });

await chart.replaceSeries([{ ...revenue[0], values: [105, 136, 172] }]);
console.log(chart.definition, await chart.diagnostics());
await chartDocument.writeFile('native-charts.pptx');
```

`CHART_TYPES` contains `area`, `bar`, `bar3D`, `bubble`, `doughnut`, `line`, `pie`, `radar`, and `scatter`. Categorical charts use `categories` and `values`; scatter uses `xValues` and `values`; bubble additionally requires positive `sizes`. Compatible bar/area/line groups can share primary axes or add secondary axes. Every native chart owns an embedded XLSX whose cells, A1 formulas, and display caches are generated from one normalized definition before the synchronous OPC transaction.

`ChartModel.definition` is a detached frozen semantic snapshot. `replaceDefinition()` replaces supported types, groups, data, and options; `replaceSeries()` updates a single-group chart; `remove()` and `slide.deleteChart()` remove the frame and garbage-collect only unreferenced chart/workbook/style/color dependencies. Duplication initially shares nothing mutable and semantic edits retain stable identity with relationship-aware clone-on-write. `setXml()` remains the explicit raw escape hatch. Strict options cover title, legend, chart/plot areas, primary/secondary axes, gridlines, labels, data tables, series fill/line/marker, colors, and type-specific grouping/gap/hole/angle/radar/scatter/bubble/3D state. Diagnostics distinguish relationship, structure, cache, axis, missing/divergent workbook, and modern-chart states.

All nine PptxGenJS 4.0.1 public chart types plus a bar+line primary/secondary combination pass real-output import, semantic edit, formula/cache/XLSX/relationship, and representative option conformance. The actual package tarball passes Node, real-Chrome, declaration, and installed-CLI smoke with `nativeCharts: true`. Its 11-slide gallery contains ten chart parts, ten embedded workbooks, and no orphan owned parts; PowerPoint 2010 validation reports 0 errors / 0 warnings, 180-DPI overflow is zero, and every page was inspected.

LibreOffice 26.8 renders the eight 2D types and the combination; `bar3D` is title-only for both the native gallery and an independent PptxGenJS control. On save, LibreOffice retains every chart object's group types and cached data but removes embedded workbooks and rewrites formulas to client placeholders. The reader treats those charts as editable `cache-only` state with `CHART_WORKBOOK_MISSING` warnings, and the first semantic replacement recreates a synchronized XLSX. Local PowerPoint 16.112 automation returned the same `-9074` for the gallery and independent control, so this environment does not establish a PowerPoint round-trip pass.

Office 2016 `cx:*` chart creation/semantic editing, external-workbook editing, chart animations, built-in trendline/error-bar creation, and broad Keynote/Google Slides certification remain pending. The advanced-charts plugin continues to cover modern inspection, trendlines, error bars, and explicit fallback. Slide backgrounds, slide numbers, default text color, and master/layout/placeholder support are complete; the next PptxGenJS parity slice is advanced text.

## Create and edit slide backgrounds

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const backgroundDocument = PptxDocument.create();
const backgroundSlide = backgroundDocument.addSlide();

backgroundSlide.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
};

await backgroundDocument.setSlideBackgroundImage(0, './background.png');
backgroundSlide.background = { kind: 'none' }; // explicit legal a:noFill
backgroundSlide.background = undefined;        // remove direct p:bg and inherit again
```

`SlideModel.background` supports direct `p:cSld/p:bg/p:bgPr` none, sRGB/theme solid, linear/path gradient, and PNG/JPEG/GIF image state. Reads are strict and non-repairing; returned colors, stops, rectangles, and image bytes are detached. Equal assignments are exact no-ops. Image backgrounds own an internal image relationship. Duplicated slides initially share its target, the first different write uses clone-on-write, and replacement, clear, or slide deletion collects only media parts with no remaining package-graph incoming reference.

High-level `setSlideBackgroundImage()` accepts the raster loader's Node path, HTTP/HTTPS or browser-relative URL, strict data URI, bytes/ArrayBuffer, Blob/File, Web stream, and async iterable sources. Source resolution, signature inspection, and an optional MIME assertion complete before the synchronous transaction. `undefined` restores inheritance; `{ kind: 'none' }` preserves explicit no-fill intent. Layout/master semantic wrappers now read and write the same supported direct background forms. `p:bgRef`, pattern/group fill, and image crop/tile/effects remain outside this API and stay byte-preserved during unrelated edits.

Valid PptxGenJS 4.0.1 solid, transparency, and PNG backgrounds reach the same final supported structure. Its `{ type: 'none' }` writes no direct background, while `{ type: 'none', color }` emits an empty `p:bgPr`; native intentionally corrects that defect and always writes legal `a:noFill` for explicit none. The actual npm tarball passes Node, real-Chrome, declaration, and installed-CLI smoke with `slideBackgrounds: true`. Two clean builds produce an identical SHA-256 manifest for all 48 dist files. The 11-slide native gallery contains 41 parts, 39 relationships, and three background media parts; PowerPoint 2010 validation is 0 errors / 0 warnings. All 11 native and seven PptxGenJS control slides render without overflow and were reviewed individually.

LibreOffice 26.8 preserves the 11-slide order, solid/gradient/image kinds, and every image payload hash on save. It normalizes explicit no-fill to inheritance, changes both gradient `rotateWithShape` flags to false, and adds a full fill rectangle to the path gradient; the saved package still validates 0/0. Local PowerPoint 16.112 automation returned `-9074`, so this environment does not establish a PowerPoint round-trip pass.

## Create, edit, and synchronize slide numbers

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const numberedDocument = PptxDocument.create({ firstSlideNumber: 5 });
const numberedSlide = numberedDocument.addSlide();

numberedSlide.slideNumber = {
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

numberedDocument.layouts[0].slideNumber = { x: inches(0.5), align: 'left' };
numberedDocument.masters[0].slideNumber = { x: inches(4.5), align: 'center' };
const numberedDuplicate = numberedDocument.duplicateSlide(0);
numberedDocument.moveSlide(numberedDocument.slides.indexOf(numberedDuplicate), 0);
numberedDocument.firstSlideNumber = 10;
numberedSlide.slideNumber = undefined;
```

`SlideModel.slideNumber`, `LayoutModel.slideNumber`, and `MasterModel.slideNumber` own only the direct `p:ph type="sldNum"` field in their respective parts; the master setter also synchronizes direct `p:hf@sldNum`. Assigning a slide never writes hidden fields into the unique master or default layout. Position and size use EMU, margins and font size use points, and transparency uses percent. Getters return detached, deeply frozen `SlideNumber` snapshots. Equal assignment and clearing an absent field are exact no-ops. The strict reader returns `undefined` for wrong namespaces, duplicate owners, ordinary fields, invalid geometry/style, shape-id collisions, or ambiguous structures, and unsafe setters fail before mutation.

`CreatePresentationOptions.firstSlideNumber` and `document.firstSlideNumber` accept signed Int32 safe integers; `undefined` removes `firstSlideNum` and restores the OOXML default of 1. Direct slide caches equal the start plus the current zero-based index, while layout/master caches are `‹#›`. Start changes and duplicate/move/delete operations synchronize safely recognized caches transactionally. Diagnostics cover fixed-id collisions, disabled masters, and noncanonical caches. PptxGenJS 4.0.1 public slide-number variants import strictly; native intentionally avoids its fixed shape id 25, random layout/master caches, disabled master, and zero-size fallback defects.

The actual 54-file tarball contains 51 `dist` files and passes Node, real-Chrome, browser conditional-export, declaration, and installed-CLI smoke with `slideNumbers: true`. Two package builds produce an identical 51-file manifest with SHA-256 `3d77e6f56b8f299f2d580112fd0ebe77d0a98c38c07764259a3735064d5f9bea`. The focused suite is 448/448. Full Vitest reports 1194 passed with one performance test skipped by default; its separate gate is 1/1. The 16-slide native gallery has 48 parts and 45 relationships and validates 0/0 against PowerPoint 2010. The 16-slide PptxGenJS control has 82 parts and 95 relationships with zero errors and four locked warnings. All 32 pages render at 180 DPI, were reviewed individually, and have minimum ink margins of 50px/81px.

LibreOffice 26.8 renders all direct slide fields but displays its own 1..16 sequence. On save it preserves 16-slide title order, 15 direct owners, the cleared field, field types, alignment, and principal explicit styles, while removing `firstSlideNum` and rewriting caches, default fonts/languages, and layout/master placeholders. Prefer direct slide fields when LibreOffice-visible portability matters instead of relying only on layout/master placeholders. The reopened package has zero errors and 15 cache-normalization warnings. Local PowerPoint 16.112 automation launched the application but produced no active presentation, PDF, or saved PPTX, so no PowerPoint round-trip pass is claimed.

## Set a slide default text color

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

`SlideModel.color` is a strict, transient default text color supporting six-digit sRGB and DrawingML theme tokens. It applies only to `addText()` / `addRichText()` runs created after the assignment, and an explicit run color wins. Changing or clearing the default never recolors existing shapes or tables; tables, masters, layouts, and placeholders do not inherit it. Getter snapshots are normalized, detached, and frozen. Duplication copies the current default, moving retains it, deletion cleans it up, and sibling slides remain isolated.

There is no legal direct slide-level default-text-color field in OOXML, so the transient default itself is never serialized. Each inherited color is materialized as standard run-level `a:solidFill` when the run is created. After write and reopen, the materialized colors remain but `slide.color === undefined`. The behavior is covered across `pptx/pptm/potx/potm/ppsx/ppsm`, valid PptxGenJS 4.0.1 output, Node, real Chrome, TypeScript declarations, and the installed CLI.

The focused Default Color run reports 10 passed / 409 skipped; full Vitest reports 1205 passed / 1 skipped, the separate performance gate is 1/1 at 998ms, and typecheck/build pass. The actual tarball has 54 files, 51 under `dist`, manifest SHA-256 `467d87ffea6994355c357dbad3b1ea18afa8538b1bacb85b6de43de90ad16829`, and tarball SHA-256 `6812000a83247fdf2d63eddf81ec6ffb43c721d478e4cdcbbf4c4a3ce2b65ad1`. Native/PptxGenJS galleries contain 11/9 slides, 38/52 parts, and 35/58 relationships; both validate 0/0 under the PowerPoint 2010 profile. All 20 pages were inspected at 180 DPI with zero overflow and 106px minimum margins. LibreOffice 26.8 retains slide/text order, custom sRGB, theme, override, and 40% transparency, normalizing only native `tx1` to equivalent `dk1`; its saved package still validates 0/0. Local PowerPoint 16.112 returned `-9074` for both native and control files without loading or producing output, so no PowerPoint round-trip pass is claimed.

## Create, select, and edit masters, layouts, and placeholders

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

The PptxGenJS spelling `masterName` is retained, but it strictly selects one unique `SlideLayoutModel.name`. `defineSlideMaster()` creates a named layout under a real parent master and supports a direct background, transient margin, slide number, and ordered rect, line, plain/rich text, placeholder, image, and chart objects. Ordinary objects remain on the layout and are inherited. Placeholder prompts also remain on the layout; slide creation materializes only empty owners that can be filled.

`PLACEHOLDER_TYPES` is the frozen `title`, `body`, `pic`, `chart`, `tbl`, and `media` tuple. Creator `placeholder` options select by the unique layout name or `{ type, index }` identity. Text/rich text, shapes, images/SVG, charts, tables, and audio/video fill their matching domains using inherited geometry. Duplicate names or identities, a domain mismatch, an unknown `masterName`, and an already-filled owner reject before unsafe mutation.

`document.masters` and `document.layouts` return stable live `SlideMasterModel` and `SlideLayoutModel` wrappers. Both expose direct background and slide-number editing, shape/placeholder collections, and placeholder, text/rich-text, shape, raster/SVG-image, and chart creation. `replaceSlideMaster()` atomically whole-replaces a layout definition while retaining its part URI, wrapper identity, and incoming slide relationships. `deleteSlideMaster()` requires a same-document replacement for a used layout, retargets slides first, and then collects only unreferenced owned dependencies. Background, content, and placeholder relationships persist. Layout `margin` is runtime-only state used by table auto-page and future `tableToSlides` work, and reopens as `undefined`.

The focused master/layout/placeholder run reports 45 passed / 434 skipped. Full Vitest reports 1256 passed / 1 skipped, the separate performance gate is 1/1 at 578ms, and typecheck, build, and package build pass. The actual tarball contains 57 files and 54 `dist` files. Two clean builds have byte-identical sorted dist-hash manifests and tarballs, with SHA-256 `0a8e958ccde379ae071a7388dc4c29278ac5033a8641976324fcd5820339ad27` and `8362a3af38a4a7e8316a7e49e8cb3f4fb405753bd20cc935db609441819ca5e8`. Packed Node, TypeScript, CLI, and real-Chrome checks all report `masterLayouts: true`; Chrome exactly reopens all six placeholder domains, master/layout backgrounds, selected relationship targets, payload hashes, and the chart definition with zero validation, console, page, or network errors.

The two-slide native gallery contains 32 parts, 29 relationships, two layouts, and one master and validates with 0 errors / 0 warnings under the PowerPoint 2010 profile. The two-slide PptxGenJS control contains 36 parts and 34 relationships. Eight source and LibreOffice-round-trip pages render at 2400×1350 and 180 DPI and were inspected individually; their full-bleed backgrounds give the expected 0px minimum non-white margin. The fixtures deliberately use 1×1 black PNG background/image payloads, and the second native slide is deliberately retargeted to the blank default layout, so black/blank output is not evidence of lost inheritance. LibreOffice 26.8 retains two slides, two layouts, and one master, but rewrites placeholder identities and slide-number caches and removes audio plus embedded chart workbooks. This is a degradation record, not a complete round-trip pass. PowerPoint 16.112 returned the same `-9074` for native and control inputs and produced no PPTX/PDF output, so no PowerPoint round-trip pass is claimed.

Full theme text cascade, percentage coordinates, advanced text/table/media/chart styling, and broader client certification remain pending. Advanced text now includes text-shape direct fill, simple line, begin/end arrows, simple shadow, outer and per-run hyperlinks, preset geometry, rounded-rectangle radius, direct `isTextBox` state, and rich-text `breakLine` paragraph splitting.

## Create and edit text-shape fills

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
plain.fill = undefined; // clears only the direct fill choice
```

`AddTextOptions.fill` reuses the strict `ShapeFill` union. It accepts only `{ kind: 'none' }` or `{ kind: 'solid', color, transparency? }`, where color is a valid six-digit sRGB or DrawingML scheme color and transparency is a finite `0..100` percentage rounded to `0.001%`. Omitted, runtime-`undefined`, and explicit none creation all write canonical direct `a:noFill`; explicit zero transparency on a solid fill remains direct `a:alpha val="100000"`. Empty or missing-color values, PptxGenJS-shaped `{ color: 'FF0000' }`, unknown keys, accessors, symbols, class instances, and out-of-range values reject before package mutation.

The same contract covers plain and rich text, `addPlaceholder()`, title/body placeholder population, `SlideLayoutModel` / `SlideMasterModel` text creation, and declarative `defineSlideMaster()` text/placeholder objects. The returned shape exposes the direct fill immediately through `ShapeModel.fill` for read, replacement, or clear. Inputs and snapshots stay detached, same-value assignment is an exact no-op, and duplicate, outer-transaction rollback, all six formats, write/reopen, and placeholder-source isolation are covered.

The cross-package focused gate is 5/5; SDK/root and adapter suites are 188/188 and 76/76. Final full Vitest is 1262 passed / 1 skipped, the separate performance gate is 1/1 at 560ms, and the TypeScript typecheck plus project build pass. The actual 57-file tarball reports `textShapeFills: true` from installed Node, declarations, browser export, and CLI checks. Real Chrome has zero validation, console, page, or network errors, and installed CLI PowerPoint 2010 validation is 0 errors / 0 warnings.

PptxGenJS 4.0.1 also writes direct no-fill for an omitted text fill, but `{ type: 'none' }` omits the direct fill choice and explicit zero transparency omits alpha. Native preserves explicit none/zero intent; supported solid and non-zero transparency cases are semantically equivalent. Gradient, pattern, picture, and group text fills remain preservation-only outside this simple creator. Text outer simple line, arrows, simple shadow, hyperlink, preset geometry, rounded-rectangle radius, direct `isTextBox` state, and rich-text `breakLine` are supported below.

## Create and edit text-shape lines

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
outlined.line = undefined; // clears only direct width/fill/dash
```

`AddTextOptions.line` reuses the strict `ShapeLine` union. It accepts only `{ kind: 'none' }` or `{ kind: 'line', color, transparency?, width?, dash? }`: color is strict six-digit sRGB or a supported scheme color, transparency is finite `0..100` rounded to `0.001%`, width is finite `0..1584` points rounded to one EMU, and dash is one of `solid/dash/dashDot/lgDash/lgDashDot/lgDashDotDot/sysDash/sysDot`. Omitted width/dash materialize as 1pt/solid; zero width and explicit zero transparency retain direct intent. Omitted, runtime-`undefined`, and explicit none creation preserve the existing canonical `<a:ln><a:noFill/></a:ln>` text-box default.

Plain/rich text, `addPlaceholder()`, placeholder population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects share the same normalizer and renderer. The returned live `ShapeModel.line` is immediately readable, replaceable, and clearable. Caller detachment, exact same-value bytes/journal no-op, duplicate isolation, outer rollback, stable identity, all six formats, write/reopen, and placeholder-source isolation are covered. PptxGenJS-shaped `type`/`dashType`/`alpha`/`lineDash`, missing color, invalid dash/range, and unknown/accessor/symbol/class input reject before mutation.

PptxGenJS 4.0.1 emits an empty `a:ln` for omitted/none/empty/missing-color text lines, relies on implicit 1pt/solid when width/dash are omitted, and collapses zero width and zero transparency. Native writes explicit reversible no-fill, default width/dash, and zero direct state. Supported sRGB/theme, non-zero transparency, positive width, and all eight dash values reach equivalent final semantics. Nested deprecated `alpha` still affects PptxGenJS text lines while `lineDash` is ignored; native accepts neither. Gradient/pattern/picture/group line fill, custom dash, and cap/compound/alignment/join remain pending. Text arrows and simple shadow are supported below; hyperlink and geometry continue as separate slices.

The cross-package focused gate is 5/5; model, SDK, root, and adapter suites are 189/189, 182/182, 9/9, and 77/77. Final full Vitest is 1268 passed / 1 skipped, the separate performance gate is 1/1 at 553ms, and both TypeScript builds plus both package builds pass. The actual 57-file tarball reports `textShapeLines: true` from installed Node, declarations, browser export, and CLI checks. Real Chrome immediate/detached/reopen state matches with zero console/page/network errors, and installed CLI PowerPoint 2010 validation is 0 errors / 0 warnings.

## Create and edit text-shape arrows

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

arrowed.line = undefined; // clears line paint while preserving endpoints
arrowed.arrows = { begin: 'oval' }; // whole replacement clears the omitted end
arrowed.arrows = undefined; // clears endpoints while preserving current line state
```

`AddTextOptions.arrows` reuses strict `ShapeArrows`. Optional `begin` / `end` accept only `none | arrow | diamond | oval | stealth | triangle`. Omitted, runtime-`undefined`, or empty values create no endpoints; explicit `none` remains a direct endpoint distinguishable from absence. Inputs normalize and detach before package mutation. Aliases, empty strings, invalid tokens, unknown/inherited/accessor/symbol keys, and non-ordinary objects reject with no package change.

Plain/rich text, `addPlaceholder()`, placeholder population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects share one renderer. Arrow-only native creation keeps canonical direct `a:noFill`; combined output is ordered line fill/dash, `headEnd`, then `tailEnd`. Live `ShapeModel.arrows` retains detached whole-replacement/clear semantics and independent line/endpoint ownership. Duplicate, rollback, all six formats, write/reopen, stable identity, asynchronous declarative detachment, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 writes empty `a:ln` for omitted text lines and omits no-fill for arrow-only or `{ type: 'none' }` plus endpoints, while native retains canonical direct no-fill. PptxGenJS ignores empty endpoints and nested/top-level `lineHead` / `lineTail`, and can emit invalid endpoint tokens verbatim; native rejects aliases and invalid tokens. All six legal begin/end types, one- and two-sided endpoints, explicit `none`, and solid line plus arrows reach equivalent final endpoint semantics. Invalid imported endpoints remain preservation-only.

The cross-package focused gate is 4/4; model, SDK, root, and adapter suites are 191/191, 184/184, 10/10, and 78/78. Final full Vitest is 1274 passed / 1 skipped, performance is 1/1 at 581ms, and both TypeScript plus package builds pass. The actual 57-file tarball reports `textShapeArrows: true` from installed Node, declarations, browser export, and installed CLI checks. Real-browser immediate/detached/reopen state matches with zero errors, and installed CLI PowerPoint 2010 validation is 0 errors / 0 warnings. The following section documents text-shape simple shadow support.

## Create and edit text-shape shadows

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
shadowed.shadow = undefined; // clears only the direct shadow
```

`AddTextOptions.shadow` reuses strict `ShapeShadow`. Outer and inner both accept a valid six-digit sRGB or theme color, finite `0..1` opacity, `0..100` point blur, `0 <= angle < 360` degrees, and `0..200` point distance; only outer accepts boolean `rotateWithShape`. Omitted fields normalize to black, 0.75, 8pt, 270°, 4pt, and outer rotate false, while every explicit zero remains zero. Inputs and nested colors detach before mutation. `type` / `offset` aliases, coercible strings, invalid ranges, unknown/inherited/accessor/symbol keys, arrays, and class instances reject without package changes.

Plain/rich text, `addPlaceholder()`, placeholder population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects share one renderer. The shadow is written in one canonical `a:effectLst` after the line/endpoints; fill, line, arrows, and effect ownership remain independent. Live `ShapeModel.shadow` immediately returns a detached deep-frozen snapshot and supports whole replacement or clear. Same-value assignment is an exact bytes/journal no-op, while `undefined` removes only the direct inner/outer child and preserves a safe effect-list container plus sibling effects. Duplicate, outer rollback, all six formats, write/reopen, stable identity, asynchronous declarative detachment, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 omits a direct effect for omitted shadow and `{ type: 'none' }`; supported legal outer output reaches the same final semantics, and legacy `offset` corresponds to native `distance`. PptxGenJS falls back from runtime zero to defaults, ignores text `rotateWithShape: true`, warning-corrects or loosely converts some invalid type/color/number inputs, and emits a mismatched closing tag for inner shadow. Native does not copy those defects: it retains explicit zero, supports theme color and rotate true, emits legal inner XML, and rejects invalid input before mutation.

The model/codec, SDK/root, and adapter suites are 234/234, 197/197, and 79/79; the cross-package focused gate is 6/6. Final full Vitest is 1280 passed / 1 skipped, performance is 1/1 at 607ms, and both TypeScript builds, both package builds, and the declaration build pass. The actual 57-file tarball reports `textShapeShadows: true` from installed Node, declarations, browser export, and installed CLI checks. Real-browser immediate/detached/reopen state matches with zero console/page/network errors, and installed CLI PowerPoint 2010 validation is 0 errors / 0 warnings.

### Text-shape hyperlink creation

```ts
const plain = slide.addText('Open product', {
  hyperlink: { url: 'https://example.com/product', tooltip: 'View product' },
});
const rich = slide.addRichText([{
  runs: [
    { text: 'Go to ' },
    {
      text: 'details',
      style: {
        hyperlink: { url: 'https://example.com/details', tooltip: '' },
        underline: false,
      },
    },
    { text: ' (unlinked)', style: { hyperlink: false } },
  ],
}], {
  hyperlink: { slide: 2, tooltip: '' },
});

plain.hyperlink = { url: 'mailto:team@example.com' };
plain.hyperlink = undefined; // clears only the whole-shape click link
```

`AddTextOptions.hyperlink` accepts the same strict `Hyperlink` as preset shapes: exactly one non-empty URL or one-based current-presentation slide target, plus an optional tooltip. Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects are covered. Inputs detach immediately; invalid, coercible, or dangling targets reject before package mutation.

Creation writes the outer relationship to the non-visual shape click and each non-empty run that does not override it; empty paragraphs and runs do not synthesize links. `RichTextRunStyle.hyperlink?: Hyperlink | false` assigns an independent URL/internal-slide relationship to one run, inherits the outer hyperlink when omitted, and suppresses that default when `false`. Linked runs default to single underline, while an explicit `RichTextRunStyle.underline` always wins.

`ShapeModel.richText` returns only valid direct run links; whole replacement supports exact no-op, relationship-ID reuse, shared clone-on-write, clear/garbage collection, and rollback. `ShapeModel.hyperlink` remains whole-shape-only and does not overwrite run-local state. URL/internal-slide targets and tooltip state survive slide/layout/master/placeholder/declarative owners, duplicate, move, delete, all six formats, and write/reopen. Duplicate self-links retarget to the duplicate, while deleting a target cleans incoming run clicks.

PptxGenJS 4.0.1 materializes omitted tooltip as `tooltip=""` and emits broken `rIdundefined` references for rich outer links. Its legal per-run form writes only run links and, like native, allocates one relationship per explicit run. Native accepts its external-run `action=""` but rejects dangling, orphaned, falsy-underline, and console-only defects before mutation.

Final full Vitest is 1303 passed / 1 skipped and performance is 1/1 at 624ms. Model, SDK, root, and adapter suites are 199/199, 191/191, 13/13, and 80/80; both TypeScript builds, both tsup builds, and declaration build pass. The actual 57-file tarball plus installed Node/declarations/browser/CLI and real Chrome report `richTextRunHyperlinks: true`; Chrome has zero validation/console/page/network errors. The external smoke deck is 24 parts / 32 relationships / 3 slides with 0 errors and 8 expected portability warnings; the internal-only deck is 20 parts / 19 relationships / 2 slides with 0 errors / 0 warnings.

## Create rounded text-shape corners

```ts
const rounded = slide.addText('Rounded text', {
  shape: 'roundRect',
  rectRadius: inches(0.5),
  width: inches(4),
  height: inches(2),
});

console.log(rounded.adjustments); // [{ name: 'adj', value: 25000 }]
rounded.adjustments = [{ name: 'adj', value: 12500 }];
```

`AddTextOptions.rectRadius?: Emu` is a creation-only shortcut for `shape: 'roundRect'`. The input must be finite and within `0..914400` EMU, then rounds to the nearest EMU. Creation writes `adj = round(rectRadius * 100000 / min(finalWidth, finalHeight))`; omission or runtime `undefined` keeps the canonical empty `a:avLst`, while explicit zero writes `adj=0`. Named-placeholder population uses the final owner extent.

After creation, use `ShapeModel.adjustments` for read, whole replacement, or `[]` clear. Resizing deliberately does not recalculate the stored guide. Plain/rich text, slide/layout/master wrappers, `addPlaceholder()`, placeholder population, and declarative master text/placeholder objects share the same behavior without changing fill, line, arrows, shadow, hyperlink, transform, or text ownership.

Valid positive PptxGenJS 4.0.1 output is semantically equivalent. Native rejects its zero/NaN truthiness loss, string coercion, wrong-shape passthrough, negative/out-of-range values, and infinite formulas before package mutation. Final verification is 1320 passed / 1 skipped tests; model, SDK, root, and adapter suites are 203/203, 195/195, 15/15, and 85/85. The actual 57-file tarball and installed Node/types/browser/CLI plus real Chrome report `textShapeRectRadius: true`, with zero Chrome validation/console/page/network errors. A three-slide QA deck validates at PowerPoint 2010 0/0, changes only the intended slide part, has no overflow, and passed slide-by-slide visual review. LibreOffice preserves every explicit guide and materializes only the omitted default as `16667`.

## Create and edit text-box state

```ts
const shapeText = slide.addText('Shape text');
console.log(shapeText.isTextBox); // false

const textBox = slide.addText('Text box', { isTextBox: true });
console.log(textBox.isTextBox); // true
textBox.isTextBox = false;
```

`AddTextOptions.isTextBox?: boolean` owns only the current `p:sp/p:nvSpPr/p:cNvSpPr@txBox`. Omission, an own data property set to `undefined`, and `false` omit the attribute; `true` writes canonical `txBox="1"`. An inherited property is treated as absent, accessors are not invoked, and any defined value must be a primitive boolean. Strings, numbers, boxed booleans, objects, and other truthy values reject before package mutation.

Live `ShapeModel.isTextBox` returns `false` for attribute absence, accepts true tokens `1/true/on` and false tokens `0/false/off`, and returns `undefined` for malformed, qualified-lookalike, duplicate-attribute, or ambiguous-owner state. Its setter accepts only booleans: true canonicalizes to `txBox="1"`, false removes the unique direct attribute, and a canonical same-value assignment is an exact bytes/journal no-op. One alias or malformed token can be repaired; ambiguous structure rejects without changes.

Plain/rich text, `addPlaceholder()`, direct layout/master methods, declarative master text/placeholder objects, and all six presentation formats share the state. Layout-placeholder materialization retains the source direct boolean. Named-placeholder population validates the call-site value, then lets the layout source win without mutating the layout/master source. The state remains independent from preset/custom geometry, adjustments/`rectRadius`, transforms, text, fill, line, arrows, shadow, hyperlinks, and placeholder identity.

Legal PptxGenJS 4.0.1 boolean public inputs have the same final semantics. PptxGenJS runtime also treats arbitrary truthy values as true; native deliberately does not copy that behavior. Final full Vitest is 1337 passed / 1 skipped and the separate performance gate is 1/1 at 704ms. The actual 57-file tarball, installed Node/browser conditional export/declarations/CLI, and real Google Chrome report `textShapeIsTextBox: true`; Chrome has zero validation, console, page, or network errors.

The two-slide QA deck has 0 PowerPoint 2010 errors and only one expected portability warning from a fixture external hyperlink. Toggling one shape changes only `/ppt/slides/slide1.xml`, adding exactly ` txBox="1"` to the target `<p:cNvSpPr/>`; the other 21 parts remain byte-identical. Source and LibreOffice renders have zero overflow and passed slide-by-slide review. LibreOffice save removes every true `txBox` state from both native and independent PptxGenJS 4.0.1 files, so this is a documented client rewrite boundary rather than a native workaround or a claimed full round trip.

## Split rich-text runs into paragraphs

```ts
const rich = slide.addRichText([{
  align: 'center',
  runs: [
    { text: 'First paragraph', breakLine: true },
    { text: '', breakLine: true }, // preserves an empty middle paragraph
    { text: 'Third paragraph', softBreakBefore: true },
  ],
}]);

console.log(rich.richText.length); // 3; snapshots never expose breakLine markers
```

`RichTextRun.breakLine?: boolean` is transient input syntax for creation and whole-rich-text replacement. A true non-final run ends the current paragraph. Middle, empty, and consecutive flags preserve the corresponding empty paragraphs; a trailing flag is consumed without adding a trailing empty paragraph. Every split segment copies the source paragraph's alignment, RTL, margins, indent, bullet, level, spacing, and tab stops. `softBreakBefore` stays attached to its run even when that run becomes the first run of a split paragraph. Run-local URL or internal-slide hyperlinks remain attached and are reindexed against the canonical paragraph/run matrix.

The field covers slide/layout/master content, placeholder prompt/population, declarative masters, live editing, duplicate/move/rollback/reopen, and all six presentation formats. Getters return only explicit canonical `RichTextParagraph[]`; they never infer or expose a private marker. `breakLine` is not an outer `AddTextOptions` field, and CR/LF in run text is not an alias for it.

Legal boolean PptxGenJS 4.0.1 output reaches equivalent paragraph, property, and hyperlink semantics. Native strictly rejects truthy/falsy strings, numbers, null, objects, and boxed booleans. PptxGenJS suppresses a first-run soft break after splitting, while native retains the existing reversible `softBreakBefore` contract. Final release gates are 1350 passed / 1 skipped tests plus performance 1/1. The 57-file tarball SHA-256 is `d06b84c0c3b8ff8e610c87c55b0fe9b67de6b41e59b5ec7fad62b206fdbe2699`; installed Node/types/browser/CLI and real Chrome report `richTextBreakLine: true`, with zero Chrome validation, console, page, or network errors.

The four-slide source deck validates at PowerPoint 2010 with 0 errors / 0 warnings. One link-target edit changes only `slide1.xml` and its relationships; the other 24 parts remain byte-identical. Five-page source and LibreOffice visual decks have no overflow, clipping, or unexpected wrapping. LibreOffice retains visible paragraphs, empty lines, soft breaks, and internal links, but merges adjacent runs, omits empty tooltips, pushes master content into layouts, renames placeholders, and drops the master placeholder prompt; owner identity is therefore not claimed as a complete round trip.

Remaining advanced text/table, `tableToSlides`, output/runtime helpers, and peer/client audit work are still required before full PptxGenJS parity is claimed.

## Read the library runtime version

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

`PPTX_VERSION` is a browser-safe compile-time literal and `PptxVersion` is its literal type. Read-only `PptxDocument.version` remains the same through create/open/write/reopen, performs no package lookup, and does not mutate the presentation. It is unrelated to OOXML extended-properties `AppVersion` or the producer/version of an input file.

A PptxGenJS 4.0.1 instance correctly reports its own `'4.0.1'`, while this runtime reports its own `'0.1.0'`; parity means public availability, stability, and manifest synchronization, not equal cross-library strings. Manifest drift tests cover all three manifests, and CLI `--version` plus JSON doctor consume the same constant. Final release gates are 1354 passed / 1 skipped tests and performance 1/1 at 617ms. Both TypeScript checks, both bundles, and declaration generation pass. The actual 58-file tarball has SHA-256 `ce300d3c5da10a8fbdb9910b10497d02af496532b99329e18a314c9604e6f9a8`; installed Node/types/browser/CLI and real Google Chrome report `presentationVersion: true`, with zero Chrome validation, console, page, or network errors.

Full PptxGenJS parity is not yet claimed at this historical checkpoint. The `OutputType` runtime catalog is completed in a later section; six actual `write({ outputType })` return semantics, stream, compression, remaining runtime constants, advanced text/table, `tableToSlides`, and the final peer/client audit remain pending.

## Read the current presentation layout

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

Getter-only `PptxDocument.presLayout` projects `{ name, width, height }` from the single `p:sldSz` / `slideSize` source of truth; dimensions are EMU. Exact 10×7.5, 10×5.625, and 10×6.25 inch canvases map to `screen4x3`, `screen16x9`, and `screen16x10`; every other valid size, including `wide`, maps to `custom`. Each read returns a detached plain-object snapshot, causes no OPC mutation, immediately follows a `slideSize` edit, and remains stable through write/reopen.

PptxGenJS 4.0.1 exposes the same EMU values for its default, four built-ins, and custom layouts. Native does not expose undeclared `_sizeW` / `_sizeH` fields or a mutable internal alias. A `defineLayout()` custom name exists only in the PptxGenJS process and is not stored in PPTX, so native uses the recoverable canonical `custom` name instead of inventing a named-layout registry.

Final release gates are 1363 passed / 1 skipped tests and performance 1/1 at 1.01s. Both TypeScript checks, both bundles, and declaration generation pass. The actual 59-file tarball has SHA-256 `a07a11156840071f0945289c0a48fdd9741549d2003ca21006e6efab28104b3d`; installed Node/types/browser/CLI and real Google Chrome report `presentationLayouts: true`, with zero Chrome validation, console, page, or network errors.

Full PptxGenJS parity is not yet claimed at this historical checkpoint. The `OutputType` runtime catalog is completed in a later section; six actual `write({ outputType })` return semantics, stream, compression, remaining runtime constants, advanced text/table, `tableToSlides`, and the final peer/client audit remain pending.

## Enumerate horizontal text alignments

```ts
import { TEXT_ALIGNMENTS, type TextAlignment } from '@jiayunxie/pptx';

for (const alignment of TEXT_ALIGNMENTS) {
  const value: TextAlignment = alignment;
  console.log(value);
}
```

`TEXT_ALIGNMENTS` is a frozen readonly tuple in the stable order `left`, `center`, `right`, `justify`. `TextAlignment` is derived directly from that tuple, keeping runtime discovery and the TypeScript union synchronized. Every token remains valid for plain/rich text and table/table-cell horizontal alignment; the existing `left → l`, `center → ctr`, `right → r`, and `justify → just` OOXML mappings, rejection behavior, and write/reopen semantics are unchanged.

PptxGenJS 4.0.1 instance `AlignH` exposes the same four keys and values. Native matches the runtime values and enumeration order through the existing root-catalog design rather than adding `PptxDocument.AlignH`, an enum-shaped object, or a mutable alias. Final release gates are 1368 passed / 1 skipped tests plus performance 1/1 at 1.17s. Both TypeScript checks, both bundles, and declaration generation pass. The actual 59-file tarball has SHA-256 `46a88495acbffb2f81eb99f290bd15928974ddad86f507941c1ea5bfb65eaa90`; installed Node/types/browser/CLI and real Google Chrome report `horizontalAlignments: true`, with zero Chrome validation, console, page, or network errors.

Overall PptxGenJS parity at this checkpoint was approximately 96%; `AlignV` is completed in the next section.

## Enumerate vertical text alignments

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

`TEXT_VERTICAL_ALIGNMENTS` is a frozen readonly tuple in the stable order `top`, `middle`, `bottom`. `TextBoxVerticalAlignment` is derived directly from that tuple, keeping runtime discovery and the TypeScript union synchronized. Every token remains valid for text boxes, slide numbers, and table/table-cell vertical alignment; the existing `top → t`, `middle → ctr`, and `bottom → b` OOXML mappings, rejection behavior, defaults, and write/reopen semantics are unchanged.

PptxGenJS 4.0.1 instance `AlignV` exposes the same three keys and values. Native matches the runtime values and enumeration order through the existing root-catalog design rather than adding `PptxDocument.AlignV`, an enum-shaped object, or a mutable alias. Final clean gates are 72 passed / 1 skipped test files, 1373 passed / 1 skipped tests, and performance 1/1 at 939ms. Both TypeScript checks, both bundles, and declaration generation pass. The actual 59-file tarball has SHA-256 `aaaa5e0ceb053a472af49732784e0ea5babb00968734ec5093cd4f80afc34095`; installed Node/types/browser/CLI and real Google Chrome report `verticalAlignments: true`. Chrome values, text/table reopen, and frozen-catalog checks pass with zero validation, console, page, or network errors.

Overall PptxGenJS parity at this checkpoint is approximately 97%; the `OutputType` runtime catalog is completed in the next section.

## Enumerate presentation output types

```ts
import { OUTPUT_TYPES, type OutputType } from '@jiayunxie/pptx';

const outputTypes: readonly OutputType[] = OUTPUT_TYPES;
for (const outputType of outputTypes) {
  console.log(outputType);
}
```

`OUTPUT_TYPES` is a frozen readonly tuple in the stable order `arraybuffer`, `base64`, `binarystring`, `blob`, `nodebuffer`, `uint8array`. `OutputType` is derived directly from that tuple. The SDK output layer owns the catalog, and the aggregate root reuses the same object; reading or iterating it never accesses or mutates a presentation package.

PptxGenJS 4.0.1 instance `OutputType` exposes the same six keys and values. Native matches the runtime values and order without adding `PptxDocument.OutputType`, an enum-shaped object, or a mutable alias. `STREAM` is not a member of that public enum and remains part of the separate stream API. This historical checkpoint publishes only the catalog; all six selectable write return types are completed in the next section.

Final release gates are 73 passed / 1 skipped test files, 1378 passed / 1 skipped tests, and performance 1/1 at 1.10s. Both TypeScript checks, both bundles, and declaration generation pass. The actual 60-file tarball has SHA-256 `31a38643c8c851ae24a381a68cd225972b76dbf7b37758c16efd2fe27248df0d`; installed Node/types/browser/CLI and real Google Chrome report `outputTypes: true`. Chrome exact values, frozen catalog, and mutation-isolation checks pass; page, bundle, and blob requests return 200 with zero console, page, or network errors.

Overall PptxGenJS parity remains approximately 97%. Six-value `write({ outputType })` return semantics are completed in the next section.

## Select `write()` output types

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

`write({ outputType })` accepts all six public tokens. `arraybuffer` returns a standalone `ArrayBuffer`; `base64` returns raw base64 without a data-URI prefix; `binarystring` returns one byte per code unit; `blob` returns a Blob with `application/zip`; `nodebuffer` returns a Node `Buffer`; and `uint8array` returns a plain `Uint8Array`. Browser requests for `nodebuffer` reject with `nodebuffer is not supported by this platform`.

Omitted options, `{}`, and validation-only options retain the default `Uint8Array` result. Generic `WriteOptions<T>` and `WriteOutput<T>` preserve literal output-type inference. The public structural type for `nodebuffer` is `Uint8Array`, avoiding a `node:buffer` dependency in browser declarations, while the Node runtime value remains a `Buffer`. Every conversion starts from the same canonical ZIP bytes and leaves the package, diagnostics, and mutation journal unchanged.

Explicit `outputType: 'blob'` follows the PptxGenJS/ZIP contract and uses `application/zip`; the convenience `writeBlob()` keeps the presentation MIME `application/vnd.openxmlformats-officedocument.presentationml.presentation`. Existing `writeFile()` and `download()` contracts are unchanged.

Final clean gates are 74 passed / 1 skipped test files, 1383 passed / 1 skipped tests, and performance 1/1 at 966ms. Both TypeScript checks, Node/browser bundles, and declaration generation pass. The actual 61-file tarball has SHA-256 `26bbc7eb7c33eb194388576db2c2eaab33c80d0d99b19ed7a9b4a7375c3f9f37`; installed Node/types/browser/CLI and real Google Chrome report `writeOutputTypes: true`. All six Node outputs and all five portable browser outputs are byte-identical and reopen successfully, with zero Chrome validation, console, page, or network errors.

Overall PptxGenJS parity remains approximately 97%. Node readable output is completed in the next section.

## Write to a Node Readable stream

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

`stream(options?: WriteBaseOptions)` is Node-only and returns a real non-object-mode `Readable` with `pipe()`, async iteration, `data/end/error` events, pause/resume/read/destroy, and ordered chunks no larger than 64 KiB. Concatenated bytes are byte-identical to `write()` for the same state and reopen directly. Strict/permissive validation and compatibility diagnostics use the same private canonical-byte path.

The ZIP writer still produces the complete canonical `Uint8Array` in memory before exposing the Readable. This API provides backpressure-aware downstream delivery, not constant-memory ZIP generation or earlier time-to-first-byte. Browsers reject before validation or ZIP writing with `PptxDocument.stream() is only supported in Node.js`; use `write()`, `writeBlob()`, or `download()` there.

PptxGenJS 4.0.1 public `stream()` actually returns a Buffer rather than a Readable; `write({ outputType: 'nodebuffer' })` already matches that byte-result behavior. Native reserves `stream()` for a real Node stream and keeps `STREAM` out of `OUTPUT_TYPES`. Final clean gates are 75 passed / 1 skipped test files, 1390 passed / 1 skipped tests, and performance 1/1 at 682ms. The actual 61-file tarball SHA-256 is `37b1d6bec7b5a144d577c57b61c0777f2aad8515015e9cbee05abd55f8e067d2`; installed Node/types/browser/CLI and real Chrome report `nodeReadableStream: true`, with zero Chrome validation, console, page, or network errors.

Overall PptxGenJS parity remains approximately 97%. Compression policy is completed in the next section.

## Control presentation ZIP compression

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

`compression` accepts primitive booleans only. Omitted, `undefined`, or `false` uses ZIP STORE for created or modified presentations; `true` uses DEFLATE level 6. The policy is identical across all six `write({ outputType })` results, `stream()`, `writeFile()`, `writeBlob()`, and `download()`. Compression changes only the ZIP representation, not OOXML semantics, diagnostics, or the mutation journal. Strings, numbers, objects, and boxed booleans reject before OPC writing with `PptxDocument output compression must be a boolean`.

There is one lossless fast path: an opened, completely unchanged presentation returns its original bytes when compression is omitted or `undefined`, preserving source compression, entry metadata, and unrelated bytes. Explicit `false` or `true` always repacks as STORE or DEFLATE. Native matches PptxGenJS 4.0.1's legal boolean intent but deliberately does not copy its explicit-`outputType` compression drop or truthy non-boolean coercion.

Final clean gates are 1400 passed / 1 skipped tests and performance 1/1 at 749.5ms. Both TypeScript checks, Node/browser bundles, and declaration generation pass. The actual 61-file tarball SHA-256 is `4bbaa25b83a0d20dd3d2239708c628afec79bcff19c69faca6fe67b03e3bd990`; representative packed Node STORE/DEFLATE outputs are 149,598/9,347 bytes, and real Chrome outputs are 84,062/9,270 bytes. Installed Node/types/browser/CLI and Chrome report `compressionPolicy: true`; the Chrome download uses ZIP method 8, reopens successfully, and has zero console/page/network errors.

Overall PptxGenJS parity remains approximately 97%. The `SchemeColor` runtime helper is completed in the next section.

## Use PptxGenJS-compatible scheme colors

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

`SCHEME_COLORS` exposes the stable ten-entry order `text1→tx1`, `text2→tx2`, `background1→bg1`, `background2→bg2`, and `accent1..accent6`; `SchemeColor` is derived directly from those values. The model owns and freezes the catalog, while the SDK and aggregate root share that same object. Reading it never accesses or mutates a presentation package, and Node and browser builds expose identical state.

PptxGenJS 4.0.1's `SchemeColor` getter has the same keys, values, and order but returns a shared mutable object. Native uses a frozen root mapping and does not add `PptxDocument.SchemeColor`, an instance getter, or a second catalog. `SchemeColor` names only the ten PptxGenJS helper values; existing native color APIs continue to accept their wider validated DrawingML scheme-color subset, so this is not the exhaustive OOXML scheme-color type.

Final clean full Vitest is 77 passed / 1 skipped test files and 1404 passed / 1 skipped tests, with performance 1/1 at 736ms. Both TypeScript checks, Node/browser bundles, declaration generation, and installed Node/types/browser/CLI gates pass. The actual 62-file tarball SHA-256 is `5d7096b0347d605c105dff15bb357781c4dcaa1cb7c3eff69f89ea6baa70e742`; Node and real Google Chrome report `schemeColors: true`. Write/reopen, frozen/shared identity, and mutation isolation pass with zero Chrome validation, console, page, or network errors. Evidence is retained at `/tmp/pptx-scheme-color-artifacts.AOU1Qb`.

## Read and bulk-edit table-level vertical alignment

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

`TableModel.verticalAlignment` is a consensus projection over every physical cell's direct `tcPr@anchor`. It returns `top`, `middle`, or `bottom` only when one or more cells all have that same safe value; mixed, absent, empty, or unsafe state returns `undefined`. Assignment atomically overwrites every physical cell, including merge continuations, while `undefined` clears all direct anchors. Same-value assignment and an all-absent clear are exact no-ops. DrawingML stores no table creation default here, and this property neither synthesizes nor remembers one; inspect `rows[].cells[].verticalAlignment` for mixed detail. Unsafe writes fail with `ModelParseError` and zero package mutation while preserving text, borders, fill, margins, direction, fit, grid, rows, and transform.

PptxGenJS 4.0.1 exposes only creation-time table/cell `valign`; its final direct cell anchors are read exactly by this consensus getter. The native bulk editor is a lossless existing-deck extension over that same OOXML state. Focused verification is 4 files / 521 tests; final full verification is 78 passed / 1 skipped test files and 1411 passed / 1 skipped tests, with performance 1/1 at 885ms. The actual 62-file tarball SHA-256 is `6ce48d8bb73d59148754f14dc379b9cd11ba34d358dd8e7ebba7b72cf8208f1e`; installed Node/types/browser/CLI and real Chrome report `tableVerticalAlignment: true`, with zero Chrome validation, console, page, or network errors. Evidence is retained at `/tmp/pptx-table-vertical-alignment-artifacts.1kZjyy`.

## Read and bulk-edit table-level text direction

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
table.textDirection = 'horz';     // writes explicit vert="horz" to every physical cell
table.textDirection = undefined;  // clears every direct tcPr@vert
```

`TableModel.textDirection` is a strict consensus projection over every physical cell's direct `tcPr@vert`. It returns `horz`, `vert`, `vert270`, or `wordArtVert` only when every cell has the same valid direct token; absent, mixed, empty, or unsafe state returns `undefined`. Absence is never synthesized as `horz`. Assignment atomically overwrites every physical cell, including merge continuations: `horz` writes an explicit attribute, while only `undefined` clears it. Same-value assignment and an all-absent clear are exact no-ops. Inspect `rows[].cells[].textDirection` for mixed detail. Unsafe edits fail with `ModelParseError` and zero partial package mutation while preserving unrelated cell/table state.

PptxGenJS 4.0.1 collapses resolved horizontal creation to attribute absence, so importing that output yields `table.textDirection === undefined`, not `horz`; the three non-horizontal creation values import as their final direct state. Explicit native `horz` and existing-deck bulk editing are lossless extensions over the same OOXML direct state. Focused verification is 5 files / 529 tests; final full verification is 79 passed / 1 skipped test files and 1419 passed / 1 skipped tests, with performance 1/1 at 1118ms. The actual 62-file tarball SHA-256 is `5f427a8ff77cf64f6dda593ec02fdbe405c44d22481f0357bf05fa39b63ec92d`; installed Node/types/browser/CLI and real Chrome report `tableTextDirection: true`, with zero Chrome validation, console, page, or network errors. Evidence is retained at `/tmp/pptx-table-text-direction-artifacts.BksCOP`.

## Read and bulk-edit table-level horizontal alignment

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
table.horizontalAlignment = 'justify'; // overwrites every physical cell
table.horizontalAlignment = 'left';    // writes explicit pPr@algn="l"
table.horizontalAlignment = undefined; // clears every direct pPr@algn
```

`TableModel.horizontalAlignment` is a strict consensus projection over every physical cell's unique single-paragraph direct `pPr@algn`. It returns `left`, `center`, `right`, or `justify` only when every cell has the same safe direct value; absent, mixed, empty, multi-paragraph, or unsafe state returns `undefined`, and absence is never synthesized as `left`. Assignment atomically overwrites every physical cell, including merge continuations. `left` writes an explicit `algn="l"`, while only `undefined` clears the attribute. Same-value assignment and an all-absent clear are exact no-ops. Inspect `rows[].cells[].horizontalAlignment` for mixed detail. Unsafe edits fail with `ModelParseError` and zero partial package mutation while preserving unrelated cell/table state.

All four legal PptxGenJS 4.0.1 table-alignment creation values import from their final direct state; omitted alignment remains `undefined`, and table center plus a right cell override projects to mixed `undefined`. Native existing-deck bulk editing is a lossless extension over the same OOXML state. Focused verification is 6 files / 537 tests; final full verification is 80 passed / 1 skipped test files and 1427 passed / 1 skipped tests, with performance 1/1 at 1.34s. The actual 62-file tarball SHA-256 is `03b376861aeb799fa21a99dd105871b8943e29bd4fe51c875a508ff295b9f9c0`; installed Node/types/browser/CLI and real Chrome report `tableHorizontalAlignment: true`. CLI inspection finds exactly four direct `pPr@algn="r"` tokens and no `tcPr/bodyPr@algn`, while Chrome has zero validation, console, page, or network errors. Evidence is retained at `/tmp/pptx-table-horizontal-alignment-artifacts.oe2f5A`.

## Read and bulk-edit table-level margins

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
table.margins = 6;               // writes 6pt on all four sides
table.margins = { top: 2, left: 4 }; // whole-replaces and clears right/bottom
table.margins = undefined;       // clears direct marL/marR/marT/marB
```

`TableModel.margins` is a strict consensus projection over the unique direct `tcPr@marL/marR/marT/marB` state of every physical cell, including merge continuations. It returns a detached `TextBoxMargins` only when every cell has the same non-empty, safe complete or partial side set and values. Absent, mixed-key/value, empty, malformed, repeated, or ambiguous state returns `undefined`. It never synthesizes canonical defaults, table-style values, or creation input; inspect `rows[].cells[].margins` for mixed detail.

Assignment accepts a point scalar, TRBL tuple, partial named object, `{}`, or `undefined` and whole-replaces every physical cell in one transaction. Scalar/TRBL writes all four sides, partial input clears omitted sides, and `{}`/`undefined` clears all four. Numerically equal assignment and an all-absent clear are exact bytes/journal no-ops. Unsafe edits fail with `ModelParseError` and zero partial package mutation while preserving text, borders, fill, alignment, direction, fit, grid, rows, transform, and relationships.

PptxGenJS 4.0.1 exposes only creation-time margins, leaving final direct cell attributes. Omitted creation imports as its explicit canonical `{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }`; uniform table margins project directly, while one cell override produces mixed `undefined`. Its legacy runtime treats a first value below 1 as inches and a value at least 1 as points. Native always uses points, so PptxGenJS `[0.05, 0.1, 0.15, 0.2]` matches native `[3.6, 7.2, 10.8, 14.4]` by final EMU without copying the input-unit ambiguity.

Focused verification is 7 files / 547 tests; final full verification is 81 passed / 1 skipped test files and 1437 passed / 1 skipped tests, with performance 1/1 at 1.65s. The actual 62-file tarball SHA-256 is `428f47de86cebb89ae19a59b4b5500f3c67c116f63107253e2bf997b04008e37`; installed Node/types/browser/CLI and real Chrome report `tableMargins: true`. CLI inspection finds exactly four direct `tcPr` elements with `marL="50800" marR="25400" marT="12700" marB="38100"`; Chrome has zero validation, console, page, or network errors. Evidence is retained at `/tmp/pptx-table-margins-artifacts.gPmz7V`.

## Read and bulk-edit table-level fill

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
console.log(table.fill); // undefined: mixed direct cell state
console.log(table.rows[0].cells[1].fill); // { kind: 'none' }
table.fill = { kind: 'none' }; // overwrites every physical cell
table.fill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'D9EAF7' },
  transparency: 0, // writes explicit direct alpha=100000
};
table.fill = undefined; // clears every direct fill choice
```

`TableModel.fill` returns a detached `TableCellFill` only when every physical cell, including merge continuations, has the same safe supported direct `tcPr` fill choice. Absent, mixed, malformed, advanced, or ambiguous state returns `undefined`; table styles, effective defaults, and creation input are never resolved. Mixed detail remains available through `rows[].cells[].fill`. Assignment atomically whole-replaces every physical cell. `{ kind: 'none' }` writes direct `a:noFill`, solid supports strict sRGB/theme color and optional finite `0..100` transparency, and `undefined` clears the direct choice. Exact no-ops preserve bytes and the mutation journal; unsafe edits raise `ModelParseError` with zero partial package mutation.

Legal uniform PptxGenJS 4.0.1 table solid fill projects from its final direct state, omitted fill stays `undefined`, and a differing cell override yields mixed `undefined`. PptxGenJS collapses `type: 'none'` with omission and transparency zero with omitted alpha. Native keeps direct none, absence, and explicit-zero transparency distinct. Existing-deck consensus read and bulk editing are native lossless extensions.

Focused verification is 8 files / 557 tests; final full verification is 1447 passed / 1 skipped tests, with performance 1/1 at 1.17s. The actual 62-file tarball SHA-256 is `ae7f09233b2ff596c21ec0dab891d5069d810d64921bce9ba96cd08771c6cfdc`; installed Node/types/browser/CLI, `pptx-inspect`, and real Google Chrome report `tableFill: true`. The final four physical cells each contain exactly one direct `a:noFill`; PowerPoint 2010 validation and Chrome validation/console/page/network errors are all zero. Evidence is retained at `/private/tmp/pptx-table-fill-artifacts.5MqrXK`.

## Read and bulk-edit table-level borders

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
console.log(table.borders); // undefined: mixed direct cell state
console.log(table.rows[0].cells[1].borders); // direct none on all four sides
table.borders = {
  top: {
    kind: 'line',
    color: { kind: 'srgb', value: 'D9EAF7' },
    width: 2,
  },
  bottom: { kind: 'none' },
}; // partial whole replacement; right/left are cleared
table.borders = { kind: 'none' }; // broadcasts four sides to every physical cell
table.borders = undefined; // clears every direct L/R/T/B side
```

`TableModel.borders` returns a detached `TableCellBorders` only when every physical cell, including merge continuations, has the same non-empty safe supported direct complete or partial L/R/T/B vector. All-absent, mixed, malformed, advanced, repeated, or ambiguous state returns `undefined`; table styles, shared edges, effective borders, defaults, and creation input are never resolved. Mixed detail remains available through `rows[].cells[].borders`. Assignment accepts scalar, exact TRBL, partial named-object, `{}`, or `undefined` input and atomically whole-replaces every physical cell. Scalar writes four sides, TRBL/named input clears omitted sides, and empty/`undefined` clears all four. Exact no-ops preserve bytes and the mutation journal; unsafe edits raise `ModelParseError` with zero partial package mutation.

PptxGenJS 4.0.1 materializes omitted borders as uniform four-side direct none. Legal uniform scalar/TRBL output projects to one consensus, while a legal differing cell override produces mixed `undefined`. Native bulk editing is a lossless existing-deck extension over final direct cell state. Native keeps direct absence, direct none, zero-width line, omitted style, and explicit solid distinct; it does not copy PptxGenJS's empty-object default gray 1pt solid, omitted-type solid materialization, TRBL-zero-to-1pt behavior, or short-tuple padding.

Focused verification is 9 files / 567 tests; final full verification is 1457 passed / 1 skipped tests, with performance 1/1 at 1.20s. The actual 62-file tarball SHA-256 is `47d9666c1dac8454524a87f7ca1898af0442c6faa7816e39cec34ff42dbf0d48`; installed Node/types/browser/CLI, `pptx-inspect`, and real Google Chrome report `tableBorders: true` / `tableBordersInspect: true`. Each final physical cell has exactly one direct `lnL/lnR/lnT/lnB` no-fill set; PowerPoint 2010 validation is 0 errors / 0 warnings, and Chrome validation/console/page/network errors are zero. Design, plan, correction, core, and package-proof commits are `8d6dfae`, `d4bf9c9`, `faca4ba`, `9195e57`, and `f8a9d0d`; evidence is retained at `/tmp/pptx-table-borders-artifacts.vyi1yo`.

### Table-cell hyperlink creation, snapshots, and editing

```ts
import { PptxDocument, TableModel } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const source = document.addSlide();
const target = document.addSlide();
const table = source.addTable([[
  { text: 'Website', options: { hyperlink: { url: 'https://example.com' } } },
  { text: 'Details', options: { hyperlink: { slide: 2, tooltip: '' } } },
]], { name: 'Cell links' });

console.log(table.rows[0].cells[0].hyperlink); // { url: 'https://example.com' }
console.log(table.rows[0].cells[1].hyperlink); // { slide: 2, tooltip: '' }
document.moveSlide(document.slides.indexOf(target), 0);
console.log(table.rows[0].cells[1].hyperlink); // { slide: 1, tooltip: '' }
table.setCellText(0, 1, 'Open details'); // preserves the cell hyperlink
table.setCellHyperlink(0, 0, {
  url: 'https://example.com/docs',
  tooltip: 'Read the docs',
});
table.setCellHyperlink(0, 1, undefined); // clears the click, preserving run style

const reopened = await PptxDocument.open(await document.write());
const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
const reopenedTable = reopenedSource.shapes.find(
  ({ name }) => name === 'Cell links',
) as TableModel;
console.log(reopenedTable.rows[0].cells[0].hyperlink);
// { url: 'https://example.com/docs', tooltip: 'Read the docs' }
console.log(reopenedTable.rows[0].cells[1].hyperlink); // undefined
```

`AddTableCellOptions.hyperlink` and `TableModel.setCellHyperlink(rowIndex, columnIndex, value)` accept a URL or a one-based current-presentation slide target. Omitted tooltip and explicit `''` remain distinct; `undefined` clears the selected zero-based physical cell's direct run click. Readonly `TableCell.hyperlink` immediately returns a detached frozen direct-state snapshot. The link belongs to the plain cell's only direct run; `setCellText()` and supported cell-property edits preserve it. Internal relationships retain target-slide identity, so moving the target updates the reported ordinal. Existing duplicate, delete, rollback, six-format, and write/reopen relationship lifecycles apply.

Native-created linked cells own independent relationships even when multiple cells use the same URL. Imported shared IDs use clone-on-write for target changes. Equal values are exact part/relationship/journal no-ops; tooltip-only edits retain the ID, unique target changes update in place, and clear/replace garbage-collect only the last reference. Adding a click supplies direct single underline when absent; clearing removes only the click and preserves underline plus every other run property. There is no table-level hyperlink default. This scalar reader/editor accepts only one direct text body, paragraph, run, run-properties node, and text node; rich, multi-run, and multi-paragraph default/local links are read through `TableCell.richText` and whole-replaced through `setCellRichText()`.

Legal PptxGenJS 4.0.1 URL/slide output imports and edits, including its extra `invalidUrl=""`, `action=""`, and `history="1"` run-click attributes; PptxGenJS has no existing-deck table-cell hyperlink editor. It materializes omitted tooltip as empty and writes `_rId` back into caller hyperlink objects. Native preserves omitted versus empty state, never mutates input, and does not copy loose coercion or dangling-relationship behavior.

Final verification is 84 passed / 1 skipped full test files and 1476 passed / 1 skipped tests, with performance 1/1 at 1.63s. The actual 62-file tarball SHA-256 is `2d06b955b48a25fc6f1e06accf2bd059045a7b3db04a6f3640c5bdca987ea816`; installed Node/types/browser/CLI, `pptx-inspect`, and real Google Chrome report `tableCellHyperlinkEditing: true`. The final evidence deck has 22 parts, 23 relationships, and 3 slides; its first table slide has 6 cells, 2 clicks, 2 matching relationships, and 6 preserved underlines. PowerPoint 2010 validation has 0 errors and only 2 expected `OPC_EXTERNAL_RELATIONSHIP` warnings, while Chrome validation/console/page/network errors are zero. Design, plan, rollback, core, and package-proof commits are `4fe0c43`, `0e566bd`, `dca33ba`, `99a6d3b`, and `93b6b09`; evidence is retained at `/tmp/pptx-table-cell-hyperlink-editing-UphgYg`.

### Rich and multi-paragraph table-cell text

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
          text: 'Quarterly summary',
          style: {
            bold: true,
            fontSize: 18,
            color: { kind: 'scheme', value: 'accent1' },
          },
        }],
      },
      {
        runs: [
          { text: 'Default link' },
          {
            text: ' · OpenAI',
            style: { italic: true, hyperlink: { url: 'https://openai.com' } },
          },
          {
            text: 'soft break without the default link',
            softBreakBefore: true,
            style: { hyperlink: false },
          },
        ],
      },
    ],
    options: { hyperlink: { url: 'https://example.com/report' } },
  },
  'first\r\n\rthird',
]], { name: 'Rich cells' });

const cell = table.rows[0]!.cells[0]!;
console.log(cell.text);     // paragraph and soft-break boundaries project to \n
console.log(cell.richText); // detached readonly RichTextParagraph[]

table.setCellRichText(0, 0, [{
  runs: [
    { text: 'Replaced', style: { bold: true } },
    { text: ' · docs', style: { hyperlink: { url: 'https://example.com/docs' } } },
  ],
}]);
await document.writeFile('rich-table-cells.pptx');
```

`AddTableCell.text` accepts either a string or `readonly RichTextParagraph[]`; structured arrays must be wrapped in a `{ text, options? }` cell and cannot be supplied as bare cells. CRLF/CR strings normalize to LF before paragraph splitting, preserving consecutive and trailing empty paragraphs. Structured input directly represents multiple paragraphs; `softBreakBefore` is an in-paragraph soft break, while `breakLine` can split a paragraph during creation normalization and never appears in snapshots. `TableCell.text` projects both paragraph and soft-break boundaries to `\n`. `TableCell.richText` preserves paragraph, run, alignment, list/spacing/tab, and supported font, size, bold/italic, color, underline, strike, highlight, outline, glow, baseline, character-spacing, transparency, and hyperlink state.

Cell `options.hyperlink` is the default run link: non-empty runs without a local override inherit it, explicit run links own independent relationships, and `hyperlink: false` suppresses inheritance. Rich-cell links are read through `richText[].runs[].style.hyperlink` and whole-replaced through `setCellRichText()`. Plain `TableCell.hyperlink` / `setCellHyperlink()` remain the scalar API for an exact safe single-paragraph/single-run cell. Slide targets retain identity; replacement reuses equal-target IDs, clone-on-writes shared relationships, and garbage-collects only the final reference.

`setCellRichText(rowIndex, columnIndex, value)` uses zero-based physical row/cell coordinates and preserves `bodyPr`, `lstStyle`, `tcPr`, table geometry, neighboring cells, and live model identity. A structurally and semantically equal replacement is an exact bytes/relationships/journal no-op. Inputs and returned snapshots detach immediately. `setCellText()` is intentionally limited to an exact safe direct single paragraph and run, preserving its style and hyperlink. Use `setCellRichText()` for rich, multi-run, multi-paragraph, field, or break cells; unsafe scalar edits fail with `ModelParseError` before mutation.

Legal PptxGenJS 4.0.1 CR/LF, `breakLine`, multi-paragraph, paragraph alignment/bullet, run font/size/bold/italic/color/underline, cell-default link, and run-local link output imports and remains editable. Native uses canonical `RichTextParagraph[]`, never mutates caller values or writes duplicate `pPr`, preserves omitted versus explicit-empty tooltip state, and rejects coercible, dangling, or descriptor-unsafe input before mutation.

Final full verification is 85 passed / 1 skipped test files and 1487 passed / 1 skipped tests in 143.22s; the 1000-part performance test is 1648ms. TypeScript project references, Node/browser bundles, and declarations pass. The actual 62-file tarball SHA-256 is `7de2354ac691ad09b58e0e103fd07ff1428caa799b548d2a65d9d19a4e0fd79f`; installed Node, NodeNext types, browser, CLI, Inspector, and PptxGenJS 4.0.1 import/edit checks pass. The final edited package has 20 parts / 19 relationships, changes only the slide XML and its relationships part, and retains two clicks with no dangling ID. PowerPoint 2010 reports zero errors and only expected external-link warnings. Google Chrome 150.0.7871.188 reports creation/snapshot/edit/link true and zero validation/console/page/network errors. Implementation and review commits are `3eb6f37`, `6b40fc5`, `fd6fc43`, `d0aa76d`, and `0e3c36a`; evidence is retained at `/tmp/pptx-table-cell-rich-text-DA3x3Z`.

### Table and cell text-style defaults

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const slide = document.addSlide();
const table = slide.addTable([[
  'inherits table defaults',
  {
    text: [{
      spacing: { after: 10 },
      runs: [
        { text: 'inherits cell defaults' },
        { text: ' · local', style: { fontSize: 12, bold: false } },
        {
          text: ' · link',
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

table.setCellText(0, 0, 'preserves the materialized direct run style');
table.setCellRichText(0, 1, [{ runs: [{ text: 'replacement does not re-inherit' }] }]);
```

`AddTableOptions` and `AddTableCellOptions` both expose `fontFamily`, `fontSize`, `bold`, `color`, and `spacing`. Creation precedence is table → cell → explicit paragraph/run. Font family, size, bold, and color override independently; spacing overlays independently by `before`, `after`, and `line`. Explicit `bold: false`, paragraph `spacing: false`, and `line: false` block the corresponding inherited values. Font families are non-empty XML-safe strings, font size is strict 1–4000pt, bold is boolean-only, and color/spacing reuse the existing strict native values. Native names remain `fontFamily` and structured `spacing`; PptxGenJS aliases such as `fontFace`, `paraSpaceBefore`, `paraSpaceAfter`, `lineSpacing`, and `lineSpacingMultiple` are rejected.

Resolved values are materialized only as direct paragraph/run OOXML on each physical cell; no table/cell creation metadata is retained. `TableCell.richText` immediately exposes final direct state. `setCellText()` preserves the current safe plain-run style template, while `setCellRichText()` whole-replaces content without reapplying creation defaults. Empty-paragraph `endParaRPr` carries the resolved font family and size. A cell-default hyperlink retains outer color; a run-local hyperlink without explicit run color skips outer color, while explicit run color still wins.

Legal PptxGenJS 4.0.1 table/cell `fontFace`, font size, bold, color, cell paragraph spacing, rich overrides, empty paragraphs, and hyperlink final state import, edit, and reopen. Native additionally propagates table-level spacing and corrects PptxGenJS truthy fallback that overwrites cell/run `bold: false`, as well as its caller-option mutation. Table-cell merges, physical row/column CRUD, auto-page/repeated headers, and automatic content measurement/layout recomputation are completed below. At this historical checkpoint overall parity was approximately 99.3%; its remaining-work statement is retained as stage evidence.

Final full verification is 85 passed / 1 skipped test files and 1497 passed / 1 skipped tests in 167.50s; the 1000-part performance test is 1565ms. TypeScript, Node/browser bundles, and declarations pass. The actual 62-file tarball SHA-256 is `79ed789e6d4f218cc5c838af9e5965e96bd7e35f132d2a630a85ac5dd39ed222`; installed Node, NodeNext types, browser conditional export, CLI, and Inspector all report table text defaults true. The final evidence deck has 18 parts / 15 relationships and 1 slide / 1 table / 3 cells. PowerPoint 2010 reports zero errors and warnings. Google Chrome 150.0.7871.188 reports create/snapshot/plain edit/rich replacement/reopen true and zero validation/console/page/network errors. Implementation, review, and published-proof commits are `e8cd0c7`, `0fc1567`, `af4e419`, `d6f3fd9`, `6e5df9a`, `145148b`, and `2eb1a5f`; evidence is retained at `/tmp/pptx-table-text-defaults-proof.ViSdTX`.

### Create, inspect, and edit table-cell merges

```ts
const table = slide.addTable([
  [{ text: 'Summary', options: { colspan: 2, rowspan: 2 } }, 'Total'],
  ['42'],
]);

console.log(table.mergeRegions);
table.unmergeCell(1, 1); // any physical member resolves the whole region
table.mergeCells(0, 0, 2, 2);
```

Creation accepts logical rows. The sum of first-row `colspan` values defines the physical column count; later rows skip columns still occupied by an active `rowspan`, so a fully covered row may be written as `[]`. `colspan` and `rowspan` accept positive safe integers only, with omitted or `1` meaning no merge in that dimension. Expansion and validation finish before package mutation; holes, overlaps, out-of-bounds spans, non-rectangular layouts, and layouts above the safety limit are rejected.

Snapshots and editors use zero-based physical row/column coordinates. `TableModel.mergeRegions` returns `[]` for a recognized unmerged table, a row-major detached deep-frozen `TableMergeRegion[]` for recognized regions, and `undefined` for unsafe or ambiguous topology. Every physical cell in a recognized region exposes `TableCell.merge` with the same anchor and dimensions; `isAnchor` marks the top-left cell. `mergeCells()` must cover at least two cells and rejects any non-identical overlap, while `unmergeCell()` accepts either the anchor or any continuation.

Merge and unmerge operations change only `rowSpan`, `gridSpan`, `vMerge`, and `hMerge` on physical-cell start tags. Hidden continuation text, style, relationships, and unknown XML are preserved and become visible after unmerge; existing cell editors continue to address that state by physical coordinate. Repeating an identical merge or unmerging an unmerged cell is an exact bytes/relationships/journal no-op. Malformed or unrecognizable topology remains byte-preserved and rejects semantic editing.

Legal PptxGenJS 4.0.1 horizontal, vertical, rectangular, and offset span output imports, edits, and can be created natively with the same final merge semantics. Native deliberately rejects its invalid lopsided non-span rows, negative/fractional spans, and out-of-bounds rowspans before observable mutation.

Final focused verification is 5/5 test files and 594/594 tests in 28.11s. The full suite is 86 passed / 1 skipped test files and 1512 passed / 1 skipped tests in 73.26s; the independent 1000-part performance gate is 1/1 at 709ms. TypeScript project references, Node/browser bundles, and declarations pass. Two independently built 59-file dist manifests match exactly, and both 62-file actual tarballs are byte-identical with SHA-256 `0c85afa9bed6a04faa5d3dab6934a3974cea731091dc673ab2ff6e92cb83343d`. Installed Node, NodeNext types, the browser conditional export, CLI, and Inspector report `tableCellMerges: true` / `tableCellMergesInspect: true`.

Google Chrome 150.0.7871.188 reports create/read/frozen snapshot/unmerge/edit/remerge/reopen true with zero validation, console, page, or network errors. The browser evidence deck has 18 parts / 15 relationships, 1 slide / 1 table, a 2×3 physical-cell matrix, and one 2×2 merge region. All four anchor/continuation token forms are present, the slide relationship part contains only its valid layout owner, and PowerPoint 2010 validation is 0 errors / 0 warnings. Recognition, creation, snapshot/editor, SDK/adapter, documentation, and package-proof commits are `688f9f6`, `3d93f07`, `db01937`, `b2f6846`, `5832399`, `7073eae`, and `f174519`; evidence is retained at `/tmp/pptx-table-cell-merges-artifacts.B7ZhGQ`. Physical row/column CRUD, auto-page/repeated headers, and automatic content measurement/layout recomputation are completed in the following sections. The 99.3% figure here is a historical checkpoint.

`PRESET_SHAPE_TYPES` is the frozen discovery catalog for all 178 canonical preset geometries accepted by `SlideModel.addShape()`. `AddShapeOptions` accepts `name`, strict `adjustments`, strict `fill`, strict `line`, strict `arrows`, strict `shadow`, strict `hyperlink`, and native EMU/OOXML-angle transform fields; use `inches()` and `degrees()` for ergonomic conversion. Omitted geometry starts at x/y/width/height = 1 inch with zero rotation and no flips; omitted fill creates direct no-fill, and omitted line keeps the canonical empty line container. Inputs are strict, descriptor-safe, detached before mutation, and reject unknown fields. The catalog uses the valid OOXML `foldedCorner`; PptxGenJS 4.0.1's invalid `folderCorner` token and runtime-only `custGeom` value are not accepted as presets.

`ShapeModel.presetType` reads only one safe direct canonical preset geometry. Reassigning the same type is an exact no-op; changing the type replaces only the geometry and clears old adjustment guides while preserving transform, name, fill, line, arrows, effects, text, order, and model identity. Creation, duplicate isolation, rollback, write/reopen, Node/browser bundles, and PptxGenJS public output are covered.

`SlideModel.addCustomShape(geometry, options)` creates OOXML custom geometry, and live `ShapeModel.customGeometry` reads or whole-replaces it. `CustomGeometry` contains one or more `CustomGeometryPath` values; each path has direct positive-safe-integer width/height, optional `none | norm | lighten | lightenLess | darken | darkenLess` fill, stroke/extrusion flags, and a dense command list. The command union covers move, line, arc, quadratic Bézier, cubic Bézier, and close; repeated move commands create multiple subpaths, multiple paths and empty command lists are supported. Numeric point/radius values use direct EMU and numeric angles use direct `1/60000°`; call `inches()` and `degrees()` when conversion is desired. Point and arc fields may alternatively reference one guide or DrawingML built-in token. Inputs detach immediately, getters return detached deep-frozen snapshots, same-value assignment is an exact bytes/journal no-op, and `undefined` clear is intentionally not supported. Assigning `presetType` converts custom to preset; assigning `customGeometry` converts preset back to custom while preserving shape identity and unrelated style/relationship state.

```ts
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
}, { fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } } });
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
const liveEvaluation = formulaShape.evaluateCustomGeometry();
const pureEvaluation = evaluateCustomGeometry(formulaShape.customGeometry!, {
  width: formulaShape.transform.width,
  height: formulaShape.transform.height,
});
```

`CustomGeometryValue` is a safe integer or a non-empty, XML-safe, no-whitespace, non-decimal string token. `CustomGeometryGuide` maps optional `adjustments` / `guides` to `a:avLst` / `a:gdLst`; names are globally unique across both lists. `CustomGeometryFormula` accepts exact readonly tuples for every DrawingML operator: unary `val/abs/sqrt`, binary `at2/cos/max/min/sin/tan`, and ternary `*/`, `+-`, `+/`, `?:`, `cat2`, `mod`, `pin`, `sat2`. The strict codec validates operator, arity, lexical values, structure, and uniqueness. Evaluation processes all adjustments in source order, then all shape guides in source order; it distinguishes cycles, forward references, and unknown tokens. A completed custom guide shadows a same-name built-in for subsequent expressions, while the defining expression can still read that built-in.

`CustomGeometryXyHandle` and `CustomGeometryPolarHandle` form the ordered `CustomGeometryHandle` discriminated union exposed through optional `CustomGeometry.handles`. `kind: 'xy'` maps to `a:ahXY`; `kind: 'polar'` maps to `a:ahPolar`; both require `position`, which maps to one direct `a:pos`. XY `xGuide/yGuide/minX/maxX/minY/maxY` map to `gdRefX/gdRefY/minX/maxX/minY/maxY`, while polar `radiusGuide/angleGuide/minRadius/maxRadius/minAngle/maxAngle` map to `gdRefR/gdRefAng/minR/maxR/minAng/maxAng`. Position and XY/radius bounds accept a direct shape-coordinate safe integer or token; numeric angle bounds use direct `1/60000°`. Every optional attribute is independently present, mixed kinds retain exact source order, and omitted or empty handles normalize to no own snapshot property. Inputs detach immediately; getters recursively freeze detached handle positions and lists; same-value whole-geometry assignment is an exact bytes/journal no-op.

`CustomGeometryConnectionSite` entries form the ordered optional `CustomGeometry.connectionSites` list. The list maps to `a:cxnLst`; each entry maps to one `a:cxn`, required `angle` maps to `ang`, and required `position` maps to one direct `a:pos`. Numeric position coordinates are direct safe integers in custom-geometry coordinate space, while numeric angles are direct `1/60000°`; either may instead be a guide or DrawingML built-in token. The codec preserves direct tokens; the evaluator resolves them to numbers without normalizing angles or testing path attachment. Site order and duplicates are retained. Omitted and empty lists normalize to no own snapshot property.

`CustomGeometryTextRectangle` is exposed as optional `CustomGeometry.textRectangle`. Its required `left`, `top`, `right`, and `bottom` fields map directly to `a:rect@l/t/r/b`; each accepts a safe integer or guide/DrawingML built-in token. Omission and explicit `{ left: 'l', top: 't', right: 'r', bottom: 'b' }` both fold to the canonical default with no own `textRectangle` snapshot property. LibreOffice `textAreaLeft/Top/Right/Bottom` guide normalization is accepted by the same strict token branch.

The strict reader requires namespace-correct direct handle, connection-site, and text-rectangle children; exact attributes; exactly one valid position per handle or site; and at most one complete valid direct `a:rect`. Malformed state remains byte-preserved but makes `customGeometry` read as `undefined` and rejects replacement before mutation. Inputs detach immediately, getter snapshots recursively deep-freeze all nested state including the text rectangle, same-value assignment is an exact bytes/journal no-op, and a changed value whole-replaces the geometry.

`evaluateCustomGeometry(geometry, context)` is a pure API whose `CustomGeometryEvaluationContext` contains positive safe-integer `width` and `height`. `ShapeModel.evaluateCustomGeometry()` reads those values from the live shape transform and returns `undefined` for a preset or strict-unreadable custom shape. Both resolve all 37 DrawingML built-ins: `3cd4/3cd8/5cd8/7cd8/b/cd2/cd4/cd8/h/hc/hd2/hd3/hd4/hd5/hd6/hd8/l/ls/r/ss/ssd2/ssd4/ssd6/ssd8/ssd16/ssd32/t/vc/w/wd2/wd3/wd4/wd5/wd6/wd8/wd10/wd32`. They evaluate all 17 operators and replace tokens throughout guides, handles, connection sites, the text rectangle, and every path command. Finite fractional results are retained, negative zero becomes zero, and DrawingML `*/` or `+/` division by zero returns zero. An omitted rectangle materializes as `{ left: 0, top: 0, right: width, bottom: height }`.

The public evaluated surface is `EvaluatedCustomGeometry` plus `EvaluatedCustomGeometryGuide`, `EvaluatedCustomGeometryPoint`, `EvaluatedCustomGeometryTextRectangle`, `EvaluatedCustomGeometryCommand`, `EvaluatedCustomGeometryXyHandle`, `EvaluatedCustomGeometryPolarHandle`, `EvaluatedCustomGeometryHandle`, `EvaluatedCustomGeometryConnectionSite`, and `EvaluatedCustomGeometryPath`. `CustomGeometryEvaluationErrorCode` and `CustomGeometryEvaluationError` expose `unknown-token`, `forward-reference`, `cyclic-reference`, `invalid-domain`, and `non-finite-result`, with optional `guideName` and `token`. Geometry/context inputs detach before calculation, no package mutation occurs, and the complete result tree—including context, lists, commands, and points—is recursively frozen.

PptxGenJS 4.0.1 legal `ShapeType.custGeom` points import as the same final path snapshot. Its runtime treats numeric values below 100 and numeric strings as inches, keeps numbers at or above 100 direct, resolves percentages against the full slide, and ignores arc endpoint `x/y`; native deliberately accepts only explicit direct values and does not reproduce those heuristics or coercions. PptxGenJS exposes no public guide-formula, arbitrary adjustment-handle, connection-site, or text-rectangle input; it emits only an empty `a:cxnLst` and the canonical default `a:rect`. Evaluator parity is therefore limited to supported final numeric paths and the default rectangle; formulas, handles, sites, and arbitrary rectangles are native extensions. Evaluation intentionally does not scale path coordinates, derive arc endpoints or resolved bounds, drag handles, or snap/create connectors.

An actual packed-package gallery covers 4 slides and 22 evaluator targets. The original and LibreOffice round-trip files both reopen as 22/22 strictly evaluable and validate against PowerPoint 2010 with 0 errors and 0 warnings. LibreOffice rewrites all 22 direct expression trees; numeric paths plus effective text rectangles match for 21/22, with only the `sqrt` sample endpoint changed from `600000` to `0`. It rewrites all guide arrays and removes or rewrites one target's handle and connection-site metadata, so these direct arrays are documented as client normalization rather than equality guarantees.

`ShapeAdjustment` is the readonly `{ name: string; value: number }` direct guide value used by preset-only `AddShapeOptions.adjustments` and `ShapeModel.adjustments`. Values are safe integers written as the operand of `a:gd@fmla="val N"`; no shape-specific conversion or range inference occurs. Lists are ordered with unique XML-safe names, inputs detach immediately, and getters return detached deep-frozen snapshots. Assignment is a whole-list replacement, `[]` keeps an empty `a:avLst`, and the setter does not accept `undefined`; assigning the same ordered list is an exact bytes/journal no-op. Complex formulas, duplicate guides, wrong namespaces, or ambiguous preset geometry read as `undefined`, and replacement rejects before mutation. Changing `presetType` resets the list while same-type assignment preserves it. PptxGenJS 4.0.1 valid `rectRadius`, `angleRange`, and `arcThicknessRatio` output imports as the same final list, while native preserves explicit zero and rejects its string coercion, truthiness loss, shortcut precedence, and malformed/unsafe passthrough. The separate API above covers custom geometry paths, arbitrary guide formulas, handles, connection sites, text rectangles, and numeric evaluation.

`ShapeFill` supports `{ kind: 'none' }` or a solid six-digit sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. `AddShapeOptions.fill` and `AddTextOptions.fill` reuse this value. `ShapeModel.fill` returns a detached direct-state snapshot and supports same-value no-op, whole replacement, and clear: none writes direct `a:noFill`, while assigning `undefined` removes the direct fill choice. It does not calculate inherited/effective color. Existing gradient, picture, pattern, and group fills remain lossless during unrelated edits and can be explicitly replaced or cleared, but advanced fill creation remains outside this simple-fill API. PptxGenJS 4.0.1's explicit none and zero transparency omit direct fill/alpha state where native preserves explicit intent; effective rendering is equivalent, not byte-identical. PptxGenJS also turns an empty or missing-color fill into black and accepts deprecated `alpha`, while native rejects both forms before mutation and uses `transparency`.

`ShapeLine` supports `{ kind: 'none' }` or a strict solid line with six-digit sRGB/theme color, optional finite 0–100 transparency, optional 0–1584 point width, and optional `solid | dash | dashDot | lgDash | lgDashDot | lgDashDotDot | sysDash | sysDot` dash. Omitted width/dash materialize as 1pt/solid, while zero width is preserved. `ShapeModel.line` returns a detached direct-state snapshot; same-value assignment is an exact no-op, none writes direct line no-fill, and `undefined` clears only owned width/fill/dash while preserving the line container, arrowheads, joins, extensions, and unrelated attributes. Unique gradient/picture/pattern/group line fill or custom dash can be explicitly replaced or cleared, but their creation is outside this simple-line API. PptxGenJS 4.0.1 maps its `dashType` to native `dash`; its omitted/explicit-none line, empty/missing-color fallback to `333333`, zero-width fallback to 1pt, omitted direct alpha for zero transparency, and ignored deprecated `alpha`/`lineDash` differ from native strict direct-intent semantics.

`ShapeArrowType` is the closed `none | arrow | diamond | oval | stealth | triangle` union. `AddShapeOptions.arrows` and `ShapeModel.arrows` use a detached `ShapeArrows` snapshot with optional `begin` / `end`; assignment is a whole replacement, so an omitted side is cleared, explicit `none` remains a direct endpoint, and `undefined` or `{}` clears both endpoints. Arrow edits own only direct head/tail children: clearing arrows preserves line width/fill/dash, while clearing `ShapeModel.line` preserves arrows. Arrows-only creation writes no implicit line color, width, or dash. Existing legal endpoint `w` / `len` values (`sm | med | lg`) remain lossless during type edits but are not exposed as editable size state. PptxGenJS 4.0.1 instead materializes `333333`/1pt/solid for arrow-only input, ignores empty/nested deprecated aliases, maps top-level deprecated aliases, and can pass invalid runtime tokens through; native rejects aliases and invalid tokens before mutation.

`ShapeShadow` is the strict `kind: 'outer' | 'inner'` union used by `AddShapeOptions.shadow`, `AddTextOptions.shadow`, and `ShapeModel.shadow`. Both kinds accept sRGB/theme color, finite `0..1` opacity, `0..100` point blur, `0 <= angle < 360` degrees, and `0..200` point distance; only outer accepts `rotateWithShape`. Defaults are black, 0.75 opacity, 8pt blur, 270°, 4pt distance, and outer rotate false. Explicit zero survives normalization. Inputs are deeply detached before mutation, getter snapshots are detached and deep-frozen, assignment is a whole replacement, same-value assignment is an exact bytes/journal no-op, and `undefined` removes only the direct inner/outer child while retaining `effectLst` and legal sibling effects. PptxGenJS 4.0.1 omission and `type: 'none'` map to native `undefined`, and its legacy `offset` maps conceptually to native `distance`; native deliberately rejects its zero-value fallback, ignored rotate flag, invalid passthrough, and malformed inner closing tag. Generic/advanced effects, preset shadow, custom shadow transforms, and image/table/chart/media shadow APIs remain outside this focused surface.

`Hyperlink` is a mutually exclusive `{ url, tooltip? } | { slide, tooltip? }` union used by `AddShapeOptions.hyperlink`, `AddTextOptions.hyperlink`, `ShapeModel.hyperlink`, `RichTextRunStyle.hyperlink`, `AddTableCellOptions.hyperlink`, and `TableModel.setCellHyperlink()`; readonly `TableCell.hyperlink` exposes the supported plain-cell direct state. URLs must be non-empty XML-safe strings; slide targets are one-based positive safe integers that must exist when assigned. Inputs and frozen getter snapshots are detached. Assignment is a whole replacement, so an omitted tooltip removes the direct attribute, an explicit empty tooltip preserves `tooltip=""`, and `undefined` clears the owned click link. Same-value assignment is an exact no-op. Internal relationships preserve target-slide identity across insert, delete, and reorder; duplicate self-links retarget to the duplicate, target deletion cleans click/hover references, and shared relationships use reference-aware clone-on-write and garbage collection. Text outer creation provides the shape click and default run link; explicit run links own independent relationships, and `false` suppresses the default. Native-created linked plain table cells own independent relationships on their only direct run; imported sharing is preserved until a target edit requires clone-on-write. PptxGenJS 4.0.1 materializes omitted tooltip as empty and can console-ignore or coerce invalid runtime values into duplicate or dangling links; native rejects those values before mutation. External links intentionally produce the validator's portability warning. Text-shape simple-line, arrow, simple-shadow, outer-hyperlink, and per-run rich-text hyperlink creation/editing plus plain table-cell hyperlink creation/read/edit/clear are supported.

Hover editing, table graphic-frame/image/chart/media hyperlink creation, action navigation, arrow size, cap/compound/alignment/join editing, advanced line fill/custom dash creation, remaining text geometry shortcuts, and percentage positions remain pending. Plain single-run table-cell scalar links and rich/multi-paragraph cell default/local run links are supported.

`CreatePresentationOptions.title` and live `document.title` use the direct core-properties title. Omitted creation input writes no title, `''` writes an explicit empty title, and `undefined` clears only the direct field. Values are strict XML-safe strings; reads follow the package-root core-properties relationship instead of assuming a part URI or prefix, same-value/absent-clear operations are exact no-ops, missing metadata can be created, and unrelated subject/creator/revision/unknown content is preserved. Unsafe malformed or ambiguous ownership is rejected rather than guessed. PptxGenJS 4.0.1 defaults its own public `title` to `PptxGenJS Presentation`; native omitted creation intentionally remains `undefined`.

`CreatePresentationOptions.author` and live `document.author` use only the direct Dublin Core creator. Native omitted creation preserves the canonical `@jiayunxie/pptx`, `''` writes an explicit empty creator, and `undefined` clears only creator. Strict XML-safe values, relationship-based lookup, alternate part URI/prefix support, same-value/absent-clear exact no-ops, missing-part creation, and malformed/ambiguous rejection match the title lifecycle. Author edits preserve `cp:lastModifiedBy`, title, subject, revision, timestamps, unknown children, relationships, and unrelated parts. PptxGenJS 4.0.1 instead defaults author to `PptxGenJS` and mirrors it into creator and lastModifiedBy; native intentionally keeps lastModifiedBy independent.

`CreatePresentationOptions.lastModifiedBy` and live `document.lastModifiedBy` own only direct core-properties `cp:lastModifiedBy`. Native omitted and runtime-`undefined` creation retain canonical `@jiayunxie/pptx`, `''` writes an explicit empty property, and `undefined` clears only lastModifiedBy. Values are strict XML-safe strings; no system-user lookup, save-time refresh, revision increment, timestamp update, fallback, or coercion occurs. Relationship- and namespace-aware reads support alternate legal part URIs and prefixes without mutating or falling back to creator. Same-value and absent-clear operations are exact no-ops, missing metadata can be created with one canonical `cp` binding, and creator, title, subject, revision, timestamps, unknown children, relationships, and unrelated parts remain unchanged. Malformed or ambiguous ownership is rejected before mutation. PptxGenJS 4.0.1 exposes no independent lastModifiedBy property and mirrors public `author` into creator plus lastModifiedBy; native keeps the two fields independently editable.

`CreatePresentationOptions.createdAt` and live `document.createdAt` own only the unique direct Dublin Core Terms `created` element whose XSI type QName resolves to `{http://purl.org/dc/terms/}W3CDTF`. Values are `string | undefined` and must match `YYYY-MM-DDTHH:mm:ss[.fraction](Z|±HH:mm)` exactly, use a valid Gregorian date in years 0001–9999, and use an offset no greater than ±14:00. Invalid lexical, calendar, timezone, non-string, or coercible values reject before mutation. Omitted and runtime-`undefined` creation preserve canonical bytes and return `undefined`; assigning `undefined` removes only created. Reads locate the core-properties relationship and resolve element, XSI attribute, and QName namespaces rather than assuming a part URI or prefix. Missing metadata is created minimally, replacement repairs simple invalid text or missing/wrong unique type state, and same-value/absent-clear operations are exact no-ops while modified, creator, lastModifiedBy, revision, unknown content, relationships, and unrelated parts remain unchanged. Native `write()` never reads the clock or refreshes timestamps. PptxGenJS 4.0.1 exposes no created setter and emits a UTC seconds value on each public `write()`; its final output imports through the same property.

`CreatePresentationOptions.modifiedAt` and live `document.modifiedAt` own only the unique direct Dublin Core Terms `modified` element with the same expanded-name-correct `W3CDTF` type and strict lexical/calendar/timezone contract as created-at. Omitted and runtime-`undefined` creation preserve canonical bytes and return `undefined`; assigning `undefined` removes only modified. Reads, missing-part creation, valid repair, exact no-ops, alternate URI/prefix support, and unsafe ownership rejection use the same relationship- and QName-aware rules. Modified-at edits preserve created, creator, lastModifiedBy, revision, unknown content, relationships, and unrelated parts. Native `write()` never reads the clock, refreshes modified-at, or derives it from another field. PptxGenJS 4.0.1 exposes no modified setter and emits a UTC seconds value on every public `write()`; adapter imports that final typed state without reproducing its hidden clock side effect.

`CreatePresentationOptions.subject` and live `document.subject` own only direct Dublin Core `dc:subject`. Native omitted and runtime-`undefined` creation remain `undefined`, `''` writes an explicit empty subject, and `undefined` clears only subject. Strict XML-safe values, relationship-based lookup, alternate part URI/prefix support, same-value/absent-clear exact no-ops, missing-part creation, and malformed/ambiguous rejection match the title lifecycle. Subject edits preserve title, creator, `cp:lastModifiedBy`, revision, timestamps, unknown children, relationships, and unrelated parts. PptxGenJS 4.0.1 defaults subject to `PptxGenJS Presentation`; custom and empty output imports exactly, while native intentionally keeps omitted state absent.

`CreatePresentationOptions.revision` and live `document.revision` own only direct core-properties `cp:revision`. The lexical `string | undefined` value accepts one or more ASCII digits, preserves leading zeros, and rejects empty, signed, decimal, exponent, Unicode-digit, or non-string values without coercion. Native zero-input and runtime-`undefined` creation retain canonical `'1'`, matching PptxGenJS 4.0.1; assigning `undefined` clears only revision and never updates timestamps or lastModifiedBy. Relationship- and namespace-aware reads return `undefined` rather than inventing a default for missing, malformed, ambiguous, or lexical-invalid existing state. Valid replacement or clear can repair simple invalid text, missing metadata can be created with one canonical `cp` namespace binding, and same-value/absent-clear operations are exact no-ops while adjacent core properties, relationships, and unrelated parts remain unchanged. PptxGenJS runtime can emit values outside its documented whole-number contract; adapter preserves those bytes, but native does not expose or create them as supported revision state.

`CreatePresentationOptions.company` and live `document.company` use only direct `Company` in the extended-properties part. Native omitted creation remains `undefined`, `''` writes an explicit empty Company, and `undefined` clears only that field. Strict XML-safe values are escaped; reads follow the root extended-properties relationship and namespace URI instead of assuming an app.xml URI or default namespace. Alternate URI/prefix state, same-value/absent-clear exact no-ops, missing-part creation, schema-friendly insertion, and malformed/ambiguous rejection are supported while Application, AppVersion, PresentationFormat, statistics, vectors, link state, unknown children, relationships, and unrelated parts remain unchanged. PptxGenJS 4.0.1 defaults company to `PptxGenJS`; its XML-safe custom/empty output imports exactly, but its runtime does not escape company metacharacters, while native always emits valid XML.

`CreatePresentationOptions.theme` and live `document.theme` expose the presentation theme's Latin heading/body typefaces as `headFontFace` and `bodyFontFace`. Zero-input native creation retains `Aptos Display` / `Aptos`; an explicit empty or partial theme uses the PptxGenJS 4.0.1 fallbacks `Calibri Light` / `Calibri`. Snapshots are detached, and assigning `document.theme` is a whole replacement, while `ThemeModel.setFonts()` edits only the supplied major/minor Latin field. Inputs are strict descriptor-safe objects with non-whitespace XML-safe strings. Resolution follows only the unique internal theme relationship directly owned by the presentation; font scripts, East Asian and complex-script faces, panose data, colors, format schemes, extensions, and unrelated parts are preserved.

```ts
const themed = PptxDocument.create({
  theme: { headFontFace: 'Noto Sans Display' },
}); // Noto Sans Display / Calibri
themed.theme = { bodyFontFace: 'Noto Sans' }; // Calibri Light / Noto Sans
themed.masterLayoutTheme.presentationTheme?.setFonts({
  majorLatin: 'Aptos Display',
}); // Aptos Display / Noto Sans
```

Presentation sections support both native creation and existing-deck editing:

```ts
const intro = document.addSection({ title: 'Intro' });
document.addSlide({ sectionTitle: 'Intro' });
document.renameSection(intro.id, 'Overview');
document.assignSlideToSection(0, intro.id);
```

`document.sections` returns a detached snapshot: absence and a valid empty section list are `[]`, while structurally unsafe or ambiguous section XML is `undefined`. Each section's detached `slideIds` are presentation slide IDs. Section commands address stable IDs, preserve empty sections and loose slides, and reject unsafe state before mutation; deleting a section leaves its slides loose instead of deleting them. Once sections are active, an `addSlide()` without `sectionTitle` creates or continues the canonical `Default-N` section. Slide duplication copies membership, deletion removes dangling membership, and moving a slide preserves its section while reordering member IDs to match presentation order. Duplicate titles are valid, and `addSlide({ sectionTitle })` selects the first exact match.

Slide visibility is a live creation and existing-deck property:

```ts
const hiddenSlide = document.addSlide();
hiddenSlide.hidden = true;
hiddenSlide.hidden = false;
```

`SlideModel.hidden` reads direct slide-root state: absence, `show="1"`, `show="true"`, and `show="on"` mean visible `false`; `show="0"`, `show="false"`, and `show="off"` mean hidden `true`. Ambiguous ownership or an unknown token returns `undefined`. The setter accepts only boolean, writes canonical `show="0"` for true, and removes the direct unqualified attribute for false. Qualified lookalikes and unknown XML remain untouched. Duplicating a slide preserves visibility, moving it does not change visibility, and section membership is independent from visibility.

Speaker notes are a live plain-text slide property:

```ts
const slide = document.addSlide();
slide.addNotes('Opening context\nKey talking point');
slide.notes = 'Revised talking point';
slide.notes = '';
slide.notes = undefined;
```

`SlideModel.notes` is `string | undefined`. A lazy native slide has `undefined`; assigning `''` creates an explicit empty notes body, while assigning `undefined` removes only that slide's notes relationship and owned notes part. `addNotes(string)` returns the same slide, normalizes CRLF/CR to LF, preserves other whitespace, escapes XML metacharacters, and rejects non-strings or illegal XML controls before mutation. Creation reuses the valid shared notes master or safely creates a missing canonical master from an unambiguous theme; clearing one slide never removes the shared master or another slide's notes. Duplicating a noted slide creates an independent notes part retargeted to the duplicate while sharing the master; moving, deleting, writing, and reopening preserve the expected lifecycle. PptxGenJS 4.0.1 eagerly creates an empty notes part even when `addNotes()` is omitted, whereas native creation deliberately keeps the lazy `undefined` state. Rich notes, notes-page layout editing, comments, header/footer/date controls, and slide-number controls are not part of this plain-text API.

`AddTableOptions.align` and `AddTableCellOptions.align` reuse the strict `TextAlignment` values `left`, `center`, `right`, and `justify`, mapped to direct `a:pPr@algn` tokens `l`, `ctr`, `r`, and `just`. A table value supplies the default for every cell that omits cell alignment or supplies runtime `undefined`; a valid cell value wins, and explicit `RichTextParagraph.align` wins for that paragraph. Omitted or runtime-`undefined` table alignment preserves current bytes and never synthesizes effective left. Final ownership is each physical cell paragraph's direct `a:pPr@algn`, not `tcPr`, `bodyPr`, or retained table metadata, so clearing a cell later does not reapply the creation default. Rich/multi-paragraph alignment is read and whole-replaced through `TableCell.richText` / `setCellRichText()`. The scalar `TableCell.horizontalAlignment` getter and `setCellHorizontalAlignment()` editor deliberately require one exact direct paragraph; missing, malformed, ambiguous, or multi-paragraph state returns `undefined` or rejects unsafe mutation. The scalar editor uses physical zero-based row/cell coordinates, writes `l`/`ctr`/`r`/`just`, and clears only the unqualified direct token with `undefined`; same-value and absent-clear calls are exact no-ops. PptxGenJS 4.0.1 materializes supported table/cell/paragraph values and precedence into the same importable direct state; native existing-deck editing is a lossless extension because PptxGenJS has no existing-deck editor. PptxGenJS still silently drops an unknown table runtime value while native creation/editing rejects invalid values before mutation.

`AddTableOptions.textDirection` and `AddTableCellOptions.textDirection` accept the strict table-cell values `horz`, `vert`, `vert270`, and `wordArtVert`. A table value is materialized onto every physical cell whose cell value is omitted or runtime `undefined`; any explicit cell value wins. Explicit cell `horz` blocks a non-horizontal table value while still writing no direct token. Creation owns direct physical-cell `tcPr@vert`: omitted, runtime-`undefined`, and resolved horizontal values write no attribute, while the three non-horizontal values write their exact token and appear immediately in `TableCell.textDirection`. Omitted/runtime-undefined table direction preserves existing bytes, and explicit table `horz` produces the same direction bytes for uncovered cells. Values are normalized descriptor-safely and detached immediately. Only final cell state is retained, so clearing a cell later does not reapply the table value. `setCellTextDirection()` deliberately retains direct-editor semantics, so assigning `horz` writes `vert="horz"` and assigning `undefined` clears the attribute. `TableModel.textDirection` reads a uniform valid direct token across every physical cell and atomically broadcasts or clears the direct state; absent, mixed, empty, or unsafe state reads as `undefined`, and mixed detail remains available through `rows[].cells[].textDirection`. Supported table/cell creation and precedence match PptxGenJS 4.0.1 final state; its resolved `horz` creation imports as absent/`undefined`, while explicit native `horz` and existing-deck bulk editing are lossless extensions. Native rejects PptxGenJS's invalid runtime-token passthrough.

`AddTableOptions.margin` supplies a point-based table creation default using a scalar, exact `[top, right, bottom, left]` tuple, or partial named sides. Resolution is per side: canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. A cell scalar/TRBL overrides all four table sides; a partial cell object overrides only the supplied sides; omitted, runtime-undefined, or empty cell margin inherits the table value. Omitted, runtime-undefined, or empty table margin preserves the original canonical bytes. The resolved state is written only to physical-cell `tcPr@marL/marR/marT/marB`, before optional anchor, L/R/T/B borders, and fill. `TableCell.margins` exposes it immediately, and `setCellMargins()` remains a whole-replacement physical-cell editor; clearing a cell does not reapply the creation default because no creation-default metadata is retained. `TableModel.margins` now provides strict uniform complete/partial consensus plus atomic whole-table replacement or clear. PptxGenJS's legacy dual-unit interpretation is intentionally not copied; native inputs always use points.

`AddTableOptions.border` accepts the same strict scalar, exact `[top, right, bottom, left]` tuple, or partial named `TableCellBorderInput` as a cell. A non-empty table value is materialized only onto physical cells whose normalized cell border is absent; any non-empty cell scalar, TRBL, named value, or explicit none blocks the entire table value. Missing sides in the chosen value use canonical direct no-fill and never inherit across layers. Empty/all-undefined table input preserves omitted bytes, while empty/all-undefined cell input inherits a supplied table border. Inputs detach deeply and resolved `TableCell.borders` are available immediately. No creation-default metadata is retained, so `setCellBorders()` replaces or clears only final direct cell state and never re-inherits. `TableModel.borders` projects one uniform non-empty supported complete/partial direct vector across all physical cells and atomically whole-replaces or clears that state. Supported explicit table/cell scalar/TRBL none, solid/dash sRGB lines, scalar zero, and cell overrides match PptxGenJS 4.0.1 final state. Native empty values, omitted style, TRBL zero, strict four-item tuples, named sides, and theme colors retain the documented stricter semantics.

`AddTableOptions.fill` accepts the same strict `TableCellFill` as a cell: `{ kind: 'none' }` or a solid six-digit sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. It is a whole-value creation default materialized into physical cells whose cell fill is omitted or `undefined`; a cell solid or none value completely overrides it. Omitted/runtime-undefined table fill preserves the original bytes, while `{}` is invalid. Inputs detach deeply and resolved `TableCell.fill` values are available immediately. No creation-default metadata is retained, so `setCellFill()` replaces or clears only final direct cell state and never re-inherits the creation input. `TableModel.fill` now projects one uniform supported direct value across all physical cells and atomically whole-replaces or clears that state. Serialization remains margins, optional anchor, L/R/T/B borders, then fill. Supported table solid fills match PptxGenJS 4.0.1 final cell state; native direct none, absence, and explicit-zero alpha remain distinguishable where PptxGenJS collapses them.

Cell-level table text-fit creation uses only `AddTableCellOptions.fit` with strict `none`, `shrink`, and `resize` values. Omitted, runtime `undefined`, and `none` produce byte-identical self-closing `bodyPr` and an immediate direct `TableCell.textFit` snapshot of `undefined`; `shrink` and `resize` write direct `normAutofit` and `spAutoFit` and snapshot immediately. Creation never writes `noAutofit`, retains table metadata, measures content, calculates final font scale, or recomputes cell/table geometry. `setCellTextFit()` keeps direct-editor semantics: `none` and `undefined` clear without restoring creation input, and fit remains independent from `tcPr@vert` direction. Table-level fit creation/default/getter/editor remains unsupported. PptxGenJS 4.0.1 ignores table/cell runtime `fit`, `autoFit`, and `shrinkText`, so native shrink/resize creation is an explicit extension rather than a parity claim.

Creation is native and does not require PptxGenJS. New documents include the complete default presentation chain and support preset or custom EMU slide sizes, direct presentation RTL, editable canvas size, plain/rich text, strict shapes, media, charts, and tables. `slide.addTable()` accepts logical rows of strings or `{ text: string | readonly RichTextParagraph[], options? }` cells, including strict `colspan` / `rowspan`, rich/multi-paragraph text and run links, table/cell font family, font size, bold, color, paragraph spacing, align, border, fill, margin, text direction and vertical alignment defaults, cell fit, exact geometry, column widths, automatic/zero/minimum row heights, measurement weights, repeated headers, rich text-row fragmentation, placeholder auto-page, and `newAutoPagedSlides`. Creation defaults materialize only into physical-cell direct state and are not retained as metadata. Existing tables expose live plain/rich cell snapshots and indexed editors, strict grid/row-size editing, relationship-safe rich replacement, merge-region snapshots and lossless merge/unmerge, merge-aware physical row/column insertion and deletion, and uniform direct table-level consensus/bulk editing for vertical alignment, text direction, horizontal alignment, margins, borders, and fill. Explicit `horz`, direct no-fill, and explicit-zero alpha remain distinct from `undefined` clear; mixed detail stays in `rows[].cells[]`, exact no-ops preserve bytes/journal, and unsafe edits reject without partial mutation. Table-level fit, diagonal/advanced borders and fills, creation styles, logical content insertion, and `tableToSlides` remain pending. The dedicated sections above and below define exact creation precedence, numeric ranges, OOXML mappings, PptxGenJS boundaries, and the wider text/shape/media APIs.

### Insert and delete table rows and columns

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
table.setCellText(1, 1, 'Editable hidden continuation');
table.deleteRows(4);
table.deleteColumns(3);
```

`insertRows()`, `deleteRows()`, `insertColumns()`, and `deleteColumns()` use zero-based physical coordinates, matching `rows[].cells[]`, merge snapshots, and indexed cell editors. Insert may address the current count to append. Delete must address an existing item and may not remove the last row or column. `count` defaults to one and is a positive safe integer; the resulting table is capped at 1,000,000 physical cells.

Row heights are non-negative EMU values, where zero means automatic height. Column widths are positive EMU values. A scalar broadcasts across the inserted items; an array must be dense and exactly match `count`. An omitted middle size copies the direct size at the insertion index, while append copies the last size. Column CRUD always synchronizes transform width to the grid sum. Row CRUD synchronizes transform height only when every direct row height is positive; otherwise it preserves the existing height because the library does not measure content.

Insertion strictly inside a merge expands its span; insertion at the anchor coordinate occurs before the region. Deletion shrinks regions, promotes the top-left surviving physical cell when the old anchor is removed, and dissolves a 1×1 result. New cells are canonical editable empty plain cells and do not inherit adjacent content or style. Surviving source bytes, hidden content, style, relationships, and opaque XML are retained. A deleted cell relationship is collected only after its final reference disappears anywhere on the slide, within the same package transaction as the structural edit.

PptxGenJS 4.0.1 has creation-time table rows, `rowH` / `colW`, and auto-page helpers but no existing-deck row/column editor. Tables produced through its public API can be imported and edited with these native lossless methods. Auto-page, repeated headers, and automatic content measurement/layout recomputation are supported below; this structural-edit API still does not insert logical content, and `tableToSlides` remains pending.

Final focused verification is 5 files / 611 tests in 29.96s. The full suite is 87 passed / 1 skipped test files and 1535 passed / 1 skipped tests in 64.27s; the independent 1000-part performance gate is 1/1 with the core test at 1204ms, test file at 1207ms, and total at 2.52s. TypeScript project references, the root build, Node/browser bundles, and declarations pass. Two clean 59-file dist manifests match exactly with manifest SHA-256 `51d0c19da69fbd81682933d4a5418ff58ef2a805b4164d624f150d1674924e41`. Both actual tarballs contain 62 files, are 660,178 bytes and byte-identical, and have SHA-256 `17d43a887a9871fd4910bcf33415d985b4d8f1968b4020670a64166c148aeaa4`. Installed Node, NodeNext declarations, the browser conditional export, CLI, and Inspector report `tableStructureEditing: true`.

Google Chrome 150.0.7871.188 reports every create/row-insert/column-insert/new-cell-edit/row-delete/column-delete/dimensions/merge/survivor/relationships/reopen stage true with zero validation, console, page, or network errors. Both Node and browser evidence decks have 18 parts / 16 relationships, 1 slide / 1 table, and a 4×4 physical matrix. Column widths `[914400, 457200, 1828800, 2743200]` sum to the transform width 5943600; row heights `[457200, 228600, 914400, 1371600]` sum to the transform height 2971800. The final 3×3 merge contains 1 anchor, 2 top, 2 left, and 4 interior continuations; hidden inserted text and a styled survivor remain. Two clicks share `rId2`, exactly one external relationship survives, and there are no orphan hyperlinks. PowerPoint 2010 validation is 0 errors / 1 expected `OPC_EXTERNAL_RELATIONSHIP` portability warning. Implementation, contract, documentation, and package-proof commits are `ee68731`, `d70c2af`, `099f345`, `89f9b1b`, `2250826`, and `1ab602f`; evidence is retained at `/tmp/pptx-table-structure-editing-proof.S1rVAZ`. The CRUD workstream is 8/8 complete; explicit-row-height auto-page/repeated headers are completed below.

### Auto-page tables with repeated headers

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create();
document.addSlide();
const source = document.addSlide();
const following = document.addSlide();

source.addTable([
  [{ text: 'Region', options: { autoPageCharWeight: -0.25 } }, 'Revenue'],
  ['Unit', 'USD'],
  [{
    text: [{ runs: [
      { text: 'North '.repeat(700), style: { bold: true } },
      { text: '$120', softBreakBefore: true, style: { italic: true } },
    ] }],
    options: { colspan: 2, margin: 0 },
  }],
  ['South', '$95'],
  ['East', '$110'],
], {
  autoPage: true,
  autoPageCharWeight: 0,
  autoPageLineWeight: 0,
  autoPageRepeatHeader: true,
  autoPageHeaderRows: 2,
  autoPageSlideStartY: inches(0.75),
  slideMargin: [inches(0.5), inches(0.4), inches(0.5), inches(0.4)],
  y: inches(5),
  columnWidths: [inches(3), inches(2)],
  // Omitted rowHeights measure every row from its content.
});

console.log(source.newAutoPagedSlides.length); // number of continuation slides
console.log(document.slides.indexOf(following)); // after every generated continuation
```

`autoPage: true` supports three row-height forms. Omitting both `height` and `rowHeights` measures every row; zero measures that row; a positive value is that row's minimum. Supplying only `height` still distributes fixed positive row heights. Thus `rowHeights: [0, inches(0.5), 0]` is an automatic/minimum/automatic vector. An all-positive vector without a measurement weight retains the original fixed structural-pagination behavior and does not re-estimate content. Geometry, row heights, column widths, start positions, and slide margins are EMU inputs.

Measurement uses 12pt when no font size is declared. Each Unicode cluster's inline advance uses `2.3 + autoPageCharWeight` as its divisor, while natural line height uses `1.67 + autoPageLineWeight` as its modifier. Both weights are strict finite values in `[-1, 1]`; cell values override table values, and omitted and explicit zero are numerically equivalent. Invalid values are never clamped or coerced. Combining marks, variation selectors, skin tones, and ZWJ sequences remain one cluster; whitespace, ASCII punctuation, Latin/digits, and wide characters have deterministic widths. Rich-run font size/character spacing, soft breaks, cell margins, paragraph margins/indents/bullets/tabs, and before/after/exact/multiple spacing participate. A colspan measures against the summed spanned width, while rowspan content contributes a minimum constraint to its complete merge block.

The source page starts at the table's own `y`. Continuation pages start at `autoPageSlideStartY`, the current runtime named-layout top margin, or the canonical 0.5-inch margin, in that order, and stop before the bottom margin. Explicit `slideMargin` is either one non-negative EMU value or an exact `[top, right, bottom, left]` tuple and takes precedence over layout margins.

When `autoPageRepeatHeader` is true, `autoPageHeaderRows` defaults to one and may select multiple header rows. Rowspan blocks remain atomic, and a merge may not cross the header/body boundary. An oversized measured row without rowspan is fragmented only at complete measured line bands, preserving paragraph/run styles, soft breaks, spacing, URL/internal-slide links, and page-local relationships. A fixed minimum that causes overflow, a merge/rowspan block, or one line band larger than body capacity rejects explicitly. Every page contains an ordinary editable table with exact row-height sums, column widths, and valid merge semantics. Internal slide targets are resolved before inserting any continuation slide.

With a table placeholder selector, the source table uses the owner's X/Y/width, continuation pages reuse its X/width with the configured continuation Y, and the owner bottom is an additional lower bound. Each page height is its actual row-height sum and is never stretched to the placeholder height. Placeholder identity, layout, section membership, slide-number cache, and stable internal targets are synchronized across generated pages.

```ts
source.addTable([['Header'], ['Body']], {
  autoPage: true,
  placeholder: 'data_table',
  autoPageSlideStartY: inches(1.25),
  rowHeights: [0, inches(0.4)],
});
```

Continuation slides are newly created contiguously after the source, use the same layout, and retain the source section membership. An existing following slide is never reused or modified. Table creation, slide insertion, relationships, order, sections, and runtime state commit in one transaction.

`source.newAutoPagedSlides` is a frozen readonly runtime snapshot of the continuation slides created by the latest successful `addTable()` call. It excludes the source. A later successful ordinary or no-overflow table resets it to empty; a failed call preserves the last successful result; deleted continuations are filtered; duplicates and reopened documents start empty because the metadata is not serialized.

Native does not copy PptxGenJS 4.0.1 caller mutation, coercion/clamping, or existing-following-slide reuse. Current overall PptxGenJS parity is approximately 99.7%. The only remaining capability item is `tableToSlides`, followed by the final peer/client audit, so 100% or full parity is not yet claimed.

The final focused gate covers 7 files (6 passed / 1 skipped) with 39 passed / 629 skipped tests in 5.44s. The full gate covers 89 files (88 passed / 1 skipped) with 1579 passed / 1 skipped tests in 40.17s. The independent 1000-part performance gate is 1/1 with a 682ms core test and 1.89s Vitest duration. TypeScript typecheck, root build, and the `@jiayunxie/pptx` package build complete in 1.40s, 1.10s, and 7.00s. Two 59-file dist manifests are identical with manifest SHA-256 `f5fd12f308fc360fb49edd0c4d35e4e43aaf5e3bd7ad1ed3c3caebe9d4a25a8e`; two 62-file actual tarballs are byte-identical with SHA-256 `3db5ca1dcf61b81ac6072639e342d8b5e1bbca9e35c2025476cc9f4d781432d3`. Installed Node, NodeNext declarations, and the browser conditional export report `tableAutoPage: true`; the installed CLI/Inspector also reports `tableAutoPageInspect: true`.

Google Chrome 150.0.7871.188 reports create/edit/move/delete/relationship/reopen true with zero validation, console, page, or network errors. Both Node and browser evidence decks contain 26 parts / 32 relationships and 5 slides / 2 tables. Creation splits seven rows into 4/4/3 across three table pages with two repeated header rows; all three pages retain one layout, contiguous section membership, header/body merges, rich styles, and stable target-slide identity. After the final move/edit/delete lifecycle, both retained table slides have 4×3 physical cells, 457200 row heights, 914400 column widths, and a 2743200×1828800 transform. The source's four clicks own 3 external + 1 internal relationships; the continuation's five clicks own 3 external + 2 internal relationships, with exact page-local ID ownership. Node/browser deck SHA-256 values are `8665affbdbc2e96eecb7683770685d08cf6761ee4d12c4d72ee057cea9943cb7` and `a6a8b4458af12dd95635e6ea3443346910de07cbbd5a91ee68b5f86e8d82e862`; both validate under the PowerPoint 2010 profile at 0 errors / 6 expected `OPC_EXTERNAL_RELATIONSHIP` warnings. The commit chain is `0b63c07`, `a52708a`, `062e0b3`, `d30a4e6`, `3ce9cc7`, `d66f884`, `c377799`, and `0ab580f`; evidence is retained at `/tmp/pptx-table-auto-page-proof.IQxIFi`. This is the explicit-row-height structural checkpoint; its 99.3% and next-step wording describe that stage. The subsequent automatic measurement/layout commits are `4482555`, `7a262ae`, `95b98ce`, `78cb279`, `6633696`, `e64e232`, `d77f54b`, and `2f7595a`; the focused gate is 7 test files / 107 passed, with typecheck, build, and diff check passing. Current progress is approximately 99.7%.

## Edit an existing presentation

```ts
import { PptxDocument } from '@jiayunxie/pptx';

const document = await PptxDocument.open('input.pptx');
document.transaction((draft) => {
  draft.rtlMode = false;
  draft.slides[0].title.text = 'Updated';
  draft.duplicateSlide(0);
});
await document.writeFile('output.pptx');
```

Transactions are synchronous and roll back all package graph changes when the callback or structural validation fails.
Slide, shape, master, layout, and theme objects keep stable identity across repeated reads and supported edits.
Slide duplication isolates owned chart/notes/comment dependencies while preserving shared layouts, images, media, and opaque targets.
Editing a shared image payload or chart XML uses clone-on-write so only the selected shape changes; chart-owned dependencies are cloned with it.

## Optional codecs

Optional capabilities are exposed as namespaces from the same package.

```ts
import { PptxDocument, transitions, animations } from '@jiayunxie/pptx';

const document = await PptxDocument.open('input.pptx');
const transitionCodec = transitions.installTransitionPlugin(document);
const timingCodec = animations.installAnimationPlugin(document);
```

## CLI

```sh
npx @jiayunxie/pptx@next --json doctor
pptx-inspect --json package inspect deck.pptx
```

The CLI is offline by default. Write operations require an explicit output path and support dry-run validation.

## Requirements

- Node.js 20 or newer, or a modern browser
- ESM

## Browser

The same import path selects the browser bundle automatically through conditional exports.

```ts
const document = await PptxDocument.open(fileInput.files[0]);
document.slides[0].title.text = 'Updated';
await document.download('updated.pptx');
```

Project documentation and source: [github.com/Xiejiayun/pptx](https://github.com/Xiejiayun/pptx)
