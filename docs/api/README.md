# Public API 0.1

## Create, open, and save

```ts
import { degrees, inches, PRESET_SHAPE_TYPES, PptxDocument } from '@pptx/sdk';

const created = PptxDocument.create({
  author: 'Presentation Team',
  company: 'Acme & Partners',
  createdAt: '2024-02-29T12:34:56.123+05:30',
  format: 'pptx',
  lastModifiedBy: 'Presentation Team',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
  revision: '7',
  rtlMode: true,
  slideSize: '16:9',
  subject: 'Revenue & Forecast',
  title: 'Quarterly Review',
});
created.addSlide();
created.author = 'Updated Author';
created.author = '';
created.author = undefined;
created.company = 'Updated Company';
created.company = '';
created.company = undefined;
created.createdAt = '2026-07-30T00:00:00Z';
created.createdAt = undefined;
created.lastModifiedBy = 'Updated Editor';
created.lastModifiedBy = '';
created.lastModifiedBy = undefined;
created.modifiedAt = '2026-07-30T01:02:03Z';
created.modifiedAt = undefined;
created.subject = 'Updated Subject';
created.subject = '';
created.subject = undefined;
created.revision = '008';
created.revision = undefined;
created.title = 'Updated Review';
created.title = '';
created.title = undefined;
created.rtlMode = false;
created.rtlMode = undefined;
await created.writeFile('created.pptx');

const customSize = PptxDocument.create({
  slideSize: { width: inches(11.7), height: inches(8.3) },
});
customSize.slideSize = { width: inches(10), height: inches(7.5) };

const document = await PptxDocument.open('input.pptx', {
  limits: { maxPartBytes: 128 * 1024 * 1024 },
  signal: abortController.signal,
});

await document.writeFile('output.pptx', {
  compatibility: 'powerpoint-2010',
  mode: 'strict',
});
```

`create()` is synchronous and starts with zero slides plus a default master, blank layout, theme, notes master, and document properties. Built-in slide sizes are `4:3`, `16:9` (the default), `16:10`, and `wide`; a `{ width, height }` value accepts any OOXML-valid 1–56 inch dimensions in EMU. `document.slideSize` reads or changes the slide canvas without scaling existing shapes or changing the notes page. `CreatePresentationOptions.rtlMode` writes the direct presentation-level RTL flag; `document.rtlMode` reads or replaces that direct value, with false writing `rtl="0"` and undefined clearing it. This global flag is independent from `AddTextOptions.rtlMode` and `RichTextParagraph.rtl`, so it does not rewrite paragraph direction or alignment. All six presentation formats can be created without using PptxGenJS; macro-enabled formats start without a VBA project.

`CreatePresentationOptions.title` and `document.title` accept only strings without invalid XML control characters. Omitting `title` creates no direct title, `''` creates a direct empty `dc:title`, and assigning `undefined` removes only that direct field while preserving the core-properties part and its subject, creator, revision, timestamps, unknown children, and lexical formatting. Reads and edits locate the part through the package-root core-properties relationship rather than a fixed URI or namespace prefix; editing an existing package with no such relationship creates `/docProps/core.xml` or the next free URI. Reading never mutates, assigning the decoded current value or clearing an absent title is an exact bytes/journal no-op, and unsafe dangling, external, wrong-type, malformed, or ambiguous ownership is rejected rather than guessed.

`CreatePresentationOptions.author` and `document.author` use the direct Dublin Core `dc:creator` and accept the same strict XML-safe strings. Omitted native creation preserves the canonical `@jiayunxie/pptx` creator, `''` writes an explicit empty creator, and assigning `undefined` removes only the direct creator. The getter never falls back to `cp:lastModifiedBy`; author edits preserve lastModifiedBy, title, subject, revision, timestamps, custom/unknown children, relationships, and unrelated parts. Relationship-based lookup, missing-part creation, exact same-value/absent-clear no-ops, prefix reuse, and malformed/ambiguous rejection match the title lifecycle. PptxGenJS 4.0.1 instead defaults author to `PptxGenJS` and mirrors every author value into both creator and lastModifiedBy; native intentionally keeps lastModifiedBy independently owned.

`CreatePresentationOptions.lastModifiedBy` and `document.lastModifiedBy` own only direct core-properties `cp:lastModifiedBy`. Native omitted and runtime-`undefined` creation preserve the canonical `@jiayunxie/pptx`; `''` writes an explicit empty property, while assigning `undefined` removes only lastModifiedBy. Values are strict XML-safe strings with no trimming, coercion, user lookup, save-time refresh, revision increment, or timestamp update. Reads use the package-root core-properties relationship and namespace URI, accept alternate legal part URIs and prefixes, never fall back to creator, and never mutate. Missing metadata can be created with one canonical `cp` namespace binding; same-value and absent-clear operations are exact bytes/journal no-ops, while creator, title, subject, revision, timestamps, unknown children, relationships, and unrelated parts remain unchanged. Malformed or ambiguous ownership is rejected before mutation. PptxGenJS 4.0.1 has no independent public lastModifiedBy field: its public `author` writes creator and lastModifiedBy as a mirrored pair, while native keeps both properties independently editable.

`CreatePresentationOptions.createdAt` and `document.createdAt` own only the unique direct Dublin Core Terms `created` element with an XSI type QName resolving to `{http://purl.org/dc/terms/}W3CDTF`. Values are `string | undefined` and must match `YYYY-MM-DDTHH:mm:ss[.fraction](Z|±HH:mm)` exactly, use a valid Gregorian date in years 0001–9999, and use an offset no greater than ±14:00; empty, whitespace-padded, timezone-less, coercible, or calendar-invalid values are rejected before mutation. Omitted and runtime-`undefined` creation preserve canonical bytes and return `undefined`; assigning `undefined` removes only created. Reads follow the package-root core-properties relationship and resolve element, XSI attribute, and QName namespaces rather than assuming a part URI or lexical prefix. Missing metadata is created minimally, valid replacement repairs simple invalid text or missing/wrong unique type state, and identical valid assignment or absent clear is an exact bytes/journal no-op. Modified time, creator, lastModifiedBy, revision, other properties, unknown content, relationships, and unrelated parts remain unchanged. Native `write()` never reads the clock or refreshes timestamps; PptxGenJS 4.0.1 has no public created setter and emits a UTC seconds value during each public `write()`, whose final OOXML state imports through this same property.

`CreatePresentationOptions.modifiedAt` and `document.modifiedAt` own only the unique direct Dublin Core Terms `modified` element with the same expanded-name-correct `W3CDTF` type and strict lexical/calendar/timezone contract as created-at. Omitted and runtime-`undefined` creation preserve canonical bytes and return `undefined`; assigning `undefined` removes only modified. Reads, minimal missing-part creation, type repair, exact same-value/absent-clear no-ops, unsafe ownership rejection, and alternate URI/prefix support use the same relationship- and QName-aware rules. Modified-at edits preserve created, creator, lastModifiedBy, revision, other properties, unknown content, relationships, and unrelated parts. Native `write()` is deterministic and never refreshes modified-at or derives it from another field. PptxGenJS 4.0.1 has no public modified setter and emits a UTC seconds value during each public `write()`; adapter imports that final typed state without reproducing the hidden clock side effect.

`CreatePresentationOptions.subject` and `document.subject` own only direct Dublin Core `dc:subject`. Native omitted and runtime-`undefined` creation preserve the canonical core-properties bytes and return `undefined`; `''` writes an explicit empty subject, while assigning `undefined` removes only that direct field. Values use the same strict XML-safe string validation, relationship-based part discovery, alternate URI/prefix support, missing-part creation, same-value/absent-clear exact no-ops, and malformed/ambiguous rejection as title and author. Subject edits preserve title, creator, lastModifiedBy, revision, timestamps, unknown children, relationships, and unrelated parts. PptxGenJS 4.0.1 defaults subject to `PptxGenJS Presentation`; its custom and empty outputs import exactly, while native intentionally does not inject that brand default.

`CreatePresentationOptions.revision` and `document.revision` own only direct core-properties `cp:revision`. Values are lexical `string | undefined`; a string must contain one or more ASCII digits, so `'0'`, `'7'`, and `'007'` are valid and leading zeros are preserved exactly, while empty, signed, decimal, exponent, Unicode-digit, and non-string inputs are rejected without coercion. Native zero-input and runtime-`undefined` creation preserve canonical `'1'`, matching PptxGenJS 4.0.1. Assigning `undefined` removes only direct revision and never updates timestamps or lastModifiedBy. Reads use the root core-properties relationship and namespace URI, return `undefined` for missing, malformed, ambiguous, or lexical-invalid existing state, and never synthesize a default; valid replacement or clear can repair simple lexical-invalid state while preserving title, subject, creator, lastModifiedBy, timestamps, unknown children, relationships, and unrelated parts. Same-value and absent-clear operations are exact no-ops, and missing metadata can be created with a single canonical `cp` binding. PptxGenJS runtime does not enforce its documented whole-number constraint; adapter round-trips those invalid bytes, but the strict native snapshot does not expose them as supported revision values.

`CreatePresentationOptions.company` and `document.company` own only direct `Company` in the OOXML extended-properties part. Native omitted creation preserves the canonical app-properties bytes and returns `undefined`; `''` writes an explicit empty Company, and assigning `undefined` removes only that direct field. Values are strict XML-safe strings with correct metacharacter escaping. Reads and edits follow the package-root extended-properties relationship, accept any legal part URI plus default or prefixed namespace form, and create `/docProps/app.xml` or the next free URI when the relationship is absent. New Company is inserted before the first conventional following property without reordering existing children. Same-value and absent-clear calls are exact bytes/journal no-ops; Application, AppVersion, PresentationFormat, statistics, vectors, link state, unknown children, relationships, and unrelated parts remain unchanged. Dangling, external, wrong-type, malformed, duplicate-relationship, or duplicate-Company ownership is rejected before mutation. PptxGenJS 4.0.1 defaults company to `PptxGenJS` and its XML-safe custom/empty outputs import exactly, but it directly interpolates company without XML escaping; native intentionally escapes metacharacters instead of reproducing malformed output.

Inputs: `Uint8Array`, `ArrayBuffer`, `Blob`/`File`, Web `ReadableStream`, or async byte iterable. Node.js additionally accepts a file path or Node readable stream. `write()` returns `Uint8Array`; browsers can use `writeBlob()` or `download()`.

## Embedded raster images

```ts
import {
  calculateRasterImageSizing,
  degrees,
  inches,
  inspectRasterImage,
  type AddImageSourceOptions,
  type RasterImageSizing,
} from '@pptx/sdk';

declare const updatedPngBytes: Uint8Array;
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
const image = await created.addImage(0, 'chart.png', options);
const info = inspectRasterImage(updatedPngBytes);
console.log(info); // { contentType: 'image/png', width, height }
const crop = calculateRasterImageSizing(info, {
  type: 'crop',
  width: inches(3),
  height: inches(2),
  source: { x: 400, y: 225, width: 800, height: 450 },
});
console.log(crop); // target width/height plus four-edge sourceRectangle
image.setTransform({ x: inches(1.5) });
image.sourceRectangle = { left: 10, top: -5, right: 5, bottom: 0 };
image.sourceRectangle = undefined;
image.replaceData(updatedPngBytes, 'image/png');
```

`PptxDocument.addImage(slideIndex, source, options?)` returns `Promise<ImageModel>`. `RasterImageSource` accepts Node file paths, HTTP/HTTPS URLs, browser-relative URLs, strict base64 data URIs, `Uint8Array`, `ArrayBuffer`, `Blob`/`File`, Web `ReadableStream`, and async byte iterables. PNG/JPEG/GIF content type and raw pixel dimensions are detected exclusively from byte signatures and format structure. `AddImageSourceOptions.contentType` is an optional canonical MIME assertion, and `signal` aborts file, Fetch, Blob, or stream loading. A source assertion mismatch, unsupported/truncated bytes, failed load, invalid slide index, or invalid options rejects before package mutation. File extensions and transport metadata are intentionally ignored.

`inspectRasterImage(bytes)` exposes the same canonical `{ contentType, width, height }` inspection without mutation. Pure `calculateRasterImageSizing(info, sizing)` accepts a positive-safe-integer EMU target frame. `contain` preserves the complete source and may produce negative source-rectangle edges; `cover` fills the frame with symmetric crop; `crop` maps an explicit finite source-pixel region. Results and nested source rectangles are detached and frozen, with each edge quantized once to 0.001%. Omitted transforms remain x/y 0, width/height one inch, rotation 0, and both flips false; intrinsic pixels do not automatically determine slide size.

When `AddImageSourceOptions.sizing` is present, top-level `width` or `height` is a type and runtime error. High-level direct `sourceRectangle` is not accepted. The complete options/sizing tree is normalized and detached before source I/O, and the sizing calculation completes before package mutation; invalid sizing, source failure, MIME mismatch, extreme contain, or out-of-bounds crop leaves the package unchanged. Omitted names use the current slide image count as `Image N`; omitted alt text is `preencoded.png`, while direct empty name and alt text remain empty.

Each call allocates one unique `/ppt/media/imageN.{png|jpeg|gif}` part with the exact supplied content type and bytes, one internal image relationship from the slide, and canonical `p:pic` XML with rectangular geometry, `noChangeAspect`, and stretch/fillRect. Normalization happens before mutation, while part/relationship/XML writes run in one package transaction; any failure restores parts, content types, relationships, ZIP state, slide XML, shape IDs, object identity, and mutation journal. Unsupported or unknown options, empty/wrong-type bytes, noncanonical MIME values, unsafe transforms, invalid XML strings, and zero/nonpositive extents are rejected with no package change.

The synchronous lower-level `SlideModel.addImage(bytes, options)` requires a non-empty copied `Uint8Array` plus canonical `RasterImageContentType`, and may accept direct `sourceRectangle`. `ImageSourceRectangle.left/top/right/bottom` are percentages where `1` means 1%; every edge must be less than 100, horizontal and vertical pair sums must each leave positive source area, and negative values support contain. The direct value is quantized to 0.001% DrawingML integer percentages. The returned live `ImageModel.sourceRectangle` getter provides a detached frozen snapshot; assignment whole-replaces or repairs the single direct `a:srcRect`, same-value assignment is an exact no-op, and `undefined` clears it while preserving the image payload, relationship, transform, effects, and neighboring fill state.

The returned `ImageModel` also exposes its inherited name/transform lifecycle plus `sourcePartUri`, `externalUrl`, and `replaceData()`. An exclusive embedded target is replaced in place; a shared target or relationship is cloned and only that image is retargeted. `duplicateSlide()` therefore starts with shared image parts and replacement isolates the edited duplicate. Creation and lifecycle are covered across pptx/pptm/ppsx/ppsm/potx/potm, Node and browser bundles, declarations, CLI smoke, and write/reopen.

Valid PptxGenJS 4.0.1 public path/data PNG/JPEG/GIF output reaches the same loader state. Six public contain/cover/equal-ratio/crop cases also match native final transform and direct `srcRect` integer percentages exactly. Native stays strict instead of reproducing PptxGenJS path/data precedence, console-only invalid handling, truthy sizing fallback, ambiguous outer/nested dimensions, or noncanonical `image/jpg` output. The actual tarball passes Node/browser/declaration/CLI smoke, and two clean builds produce identical SHA-256 manifests for all 38 dist files.

The sizing gallery has four slides, 40 shapes, and 12 images. Source and LibreOffice round-trip packages reopen strictly, validate against PowerPoint 2010 with 0 errors and 0 warnings, report zero overflow, and were visually inspected page by page. LibreOffice retains 12/12 payload hashes, content types, names, non-empty alt text, order, and internal relationships; it deduplicates the 12 repeated payload targets to three and rewrites all 12 picture blocks. The maximum transform quantization is 360 EMU, the maximum source-rectangle quantization is 0.007%, and two flips normalize equivalently to an added 180-degree rotation. Source 180-DPI pages are 2400×1350. The round-trip page width is client-normalized, producing 2401×1350 directly at 180 DPI and a proportional 2400×1350 inspection raster.

## Embedded SVG images

```ts
import {
  calculateImageSizing,
  inches,
  inspectSvgImage,
} from '@pptx/sdk';

declare const svgBytes: Uint8Array;
declare const fallbackPngBytes: Uint8Array;
declare const replacementSvgBytes: Uint8Array;
declare const replacementPngBytes: Uint8Array;

const svgInfo = inspectSvgImage(svgBytes);
const placement = calculateImageSizing(svgInfo, {
  type: 'cover',
  width: inches(6),
  height: inches(4),
});
const svgImage = await created.addImage(0, svgBytes, {
  contentType: 'image/svg+xml',
  fallback: fallbackPngBytes,
  name: 'Architecture',
  altText: 'System architecture diagram',
  sizing: { type: 'contain', width: inches(6), height: inches(4) },
});
console.log(placement.sourceRectangle);
console.log(svgImage.isSvg, svgImage.fallbackPartUri, svgImage.svgPartUri);

const lowLevelSvg = created.slides[0]!.addSvgImage(svgBytes, fallbackPngBytes, {
  x: inches(6.5),
  width: inches(4),
  height: inches(3),
});
lowLevelSvg.replaceSvgData(replacementSvgBytes, replacementPngBytes);
```

`PptxDocument.addImage()` detects strict SVG XML from the same Node path, HTTP/HTTPS URL, browser-relative URL, canonical base64 data URI, `Uint8Array`, `ArrayBuffer`, Blob/File, Web Stream, and async-iterable source union used by raster images. `ImageContentType` adds canonical `image/svg+xml`; `AddImageSourceOptions.contentType` remains an optional assertion. `inspectImage()` dispatches raster signatures or SVG XML, while `inspectSvgImage()` returns frozen `{ contentType: 'image/svg+xml', width, height }` intrinsic information. Generic `calculateImageSizing()` supports contain, cover, and source-coordinate crop for both raster and SVG; raster-named inspector/sizing aliases remain available.

`AddImageSourceOptions.fallback` is valid only for SVG and must resolve to a signature-valid PNG. Fallback resolution order is explicit PNG, browser Canvas rasterization (maximum side 8,192 and maximum 16,777,216 pixels), then a detached built-in transparent PNG. Every high-level generated `.png` part contains a real PNG signature. Callers that require full appearance in fallback-only clients should always supply a high-fidelity PNG; the API does not execute SVG scripts, fetch external SVG resources, or guarantee arbitrary rasterization fidelity.

The synchronous `SlideModel.addSvgImage(svgBytes, fallbackPngBytes, options?)` requires and copies non-empty `Uint8Array` payloads, leaving SVG/PNG byte correctness to the low-level caller, then commits a canonical two-part picture in one package transaction. The high-level API performs strict SVG and PNG inspection before calling it. A safe pair exposes `ImageModel.isSvg === true`, `sourcePartUri === fallbackPartUri`, and a distinct `svgPartUri`. The reader resolves the Office SVG extension by namespace and relationship semantics rather than a fixed prefix and accepts LibreOffice's relationship-identified `image/svg` normalization. External SVG relationships and ambiguous/malformed pair state are not exposed as editable SVG state.

`ImageModel.replaceData()` rejects any Office SVG extension candidate. `replaceSvgData(svgBytes, fallbackPngBytes)` validates and copies both non-empty byte inputs before mutation and updates an exclusive canonical pair in place; shared or noncanonical targets are cloned and only the selected picture is retargeted. Slide duplication initially shares both targets. Pair replacement, rollback, six-format write/reopen, arbitrary prefix import, PptxGenJS import, and LibreOffice-normalized shared-target clone-on-write are covered. A missing/unsafe pair or either replacement failure leaves parts, relationships, content types, slide XML, model identity, ZIP state, and mutation journal unchanged.

PptxGenJS 4.0.1 conformance covers three public cases: data-contain, path-cover, and data-crop. Native matches the final picture order, transform, direct `srcRect`, extension URI/namespace, relationship roles, SVG payload, metadata, and canonical content types. The intentional divergence is PptxGenJS's path SVG fallback containing SVG bytes in a `.png` part; native always emits valid PNG fallback bytes.

The packed Node/browser/type/CLI smoke covers explicit/default fallbacks, sizing, duplicate sharing, paired clone-on-write, replacement/reopen, Canvas fallback, two internal relationships, and PowerPoint 2010 validation. Two clean builds have identical SHA-256 manifests for all 38 dist files. The five-slide gallery contains 13 shapes, eight SVG pictures, seven SVG parts, seven PNG fallbacks, and 16 image relationships. Source and LibreOffice round-trip packages strictly reopen and validate with 0 errors and 0 warnings. LibreOffice preserves shape order, names, alt text, SVG hashes, relationship roles, and 7+7 targets, normalizes `image/svg+xml` to `image/svg`, quantizes position/size by at most 360 EMU and `srcRect` by at most 0.003%, and may render the PNG fallback after save.

SVG DOM editing, external SVG relationships, image rounding/transparency, alt-text editing, image hyperlink/shadow and advanced placeholder styling, and public per-image deletion/media garbage collection remain outside this slice. Picture-placeholder population itself is supported by the named master/layout API below. Strict embedded media creation is documented in the Media section below.

## Presentation format

`document.format` is detected from the presentation part content type and is one of `pptx`, `pptm`, `ppsx`, `ppsm`, `potx`, or `potm`. `document.formatProfile` also reports whether the package is macro-enabled, a slideshow, or a template. Unknown presentation content types are rejected instead of being treated as `.pptx`.

## Semantic model

```ts
document.slides[0].title.text = 'Updated';
document.slides[0].shapes[0].setTransform({ x: inches(1.5) });
document.duplicateSlide(0);
document.moveSlide(1, 0);
document.deleteSlide(2);
```

Plain speaker notes can be created and edited through the live slide model:

```ts
const slide = document.addSlide();
slide.addNotes('Opening context\nKey talking point');
slide.notes = 'Revised talking point';
slide.notes = '';
slide.notes = undefined;
```

`SlideModel.notes` returns `string | undefined`: absence is lazy `undefined`, an explicit empty body is `''`, and `undefined` assignment clears only the selected slide's notes relationship and owned notes part. Both the property setter and chainable `addNotes(string)` accept only XML-safe strings, normalize CRLF/CR to LF, and preserve leading/trailing whitespace. Reads follow the unique internal slide→notesSlide relationship, validate its slide backlink and shared notes-master chain, and flatten the unique direct body placeholder to plain text without mutating the package. Same-value assignment and clearing an absent value are exact byte/journal no-ops. Creation repairs a safely missing body placeholder and can create one canonical notes master only from fully absent, unambiguous topology using the presentation theme or first ordered slide-master theme; partial or ambiguous ownership is rejected before mutation. Duplication clones and retargets the per-slide notes part while retaining the shared master, deletion garbage-collects only unreferenced per-slide notes, and move/sections/hidden state remain independent. PptxGenJS 4.0.1 public output eagerly materializes empty notes for an omitted call; native creation intentionally distinguishes that from lazy `undefined`. This API is plain-text only and does not edit rich notes, notes-page layout, comments, header/footer/date fields, or slide numbers.

Preset shapes can be created with ordered adjustments plus direct fill, line, arrow, shadow, and hyperlink state,
and those direct values can be read or replaced:

```ts
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
shape.presetType = 'hexagon';
shape.fill = { kind: 'none' };
shape.line = { kind: 'line', color: { kind: 'scheme', value: 'accent2' } };
shape.arrows = { begin: 'diamond' }; // clears the omitted end
shape.arrows = { begin: 'none', end: 'oval' };
shape.arrows = undefined; // clears both endpoints, preserves line style
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
console.log(shape.presetType); // 'hexagon'
console.log(PRESET_SHAPE_TYPES.length); // 178
```

`PRESET_SHAPE_TYPES` is a runtime-frozen tuple of the 178 canonical preset tokens; `PresetShapeType` is derived from it. `SlideModel.addShape(type, options?)` accepts `name`, strict `adjustments`, strict `fill`, strict `line`, strict `arrows`, strict `shadow`, strict `hyperlink`, and `Partial<Transform>` in native EMU and OOXML-angle units. Omitted x/y/width/height are each one inch, rotation is zero, both flips are false, the default name is `Shape ${id}`, omitted fill creates direct `a:noFill`, and omitted line keeps an empty direct `a:ln`. Options must be ordinary or null-prototype objects with supported own data properties; accessors, inherited/unknown/symbol keys, invalid XML strings, unsafe numbers, non-positive extents, invalid rotation, and non-boolean flips are rejected before package mutation.

`ShapeModel.presetType` returns a canonical token only for one namespace-correct direct `p:spPr/a:prstGeom`; missing, unknown, nested, malformed, repeated, or qualified-lookalike geometry returns `undefined` without mutation. Assigning the current token preserves the exact geometry bytes, including adjustments. Assigning another canonical token whole-replaces only that direct geometry with one empty `a:avLst`, intentionally clearing stale adjustments while preserving non-visual identity, transform, fill, line, arrows, effects, extensions, text, shape order, and the live model object. Creation and replacement are transactional and remain isolated through duplication, rollback, all six formats, write, and reopen. PptxGenJS 4.0.1 public `ShapeType/addShape/write` output is compared semantically for every legal token. Its `folderCorner` value is invalid OOXML and reads as `undefined`; native uses valid `foldedCorner`. `custGeom` is not a preset token and is exposed through `addCustomShape()` / `customGeometry` instead.

`CustomGeometryValue`, `CustomGeometryFormula`, `CustomGeometryGuide`, `CustomGeometryPoint`, `CustomGeometryXyHandle`, `CustomGeometryPolarHandle`, `CustomGeometryHandle`, `CustomGeometryConnectionSite`, `CustomGeometryTextRectangle`, `CustomGeometryCommand`, `CustomGeometryPathFill`, `CustomGeometryPath`, and `CustomGeometry` describe the public custom-geometry tree consumed by `SlideModel.addCustomShape(geometry, options)` and returned by `ShapeModel.customGeometry`. Path width/height are positive safe-integer coordinate-space extents. Point/radius values are direct EMU and arc angles are OOXML `1/60000°` when numeric; `inches()` and `degrees()` are explicit ergonomic converters. Point and arc fields are `CustomGeometryValue`, so they may instead reference one guide or DrawingML built-in token. The command union contains `moveTo`, `lineTo`, `arcTo`, `quadraticBezierTo`, `cubicBezierTo`, and `close`. Repeated moves form multiple subpaths; multiple paths, empty command lists, negative numeric point coordinates, zero numeric angles, and optional path `fill` / `stroke` / `extrusionOk` flags are supported. Arc commands deliberately have no endpoint because OOXML derives it from the current point, radii, and angles.

`CustomGeometry.adjustments` and `CustomGeometry.guides` map to `a:avLst` and `a:gdLst`. Each ordered `CustomGeometryGuide` has a globally unique name and a typed formula. `CustomGeometryFormula` supports the complete 17-operator grammar with exact tuple arity: unary `val`, `abs`, `sqrt`; binary `at2`, `cos`, `max`, `min`, `sin`, `tan`; ternary `*/`, `+-`, `+/`, `?:`, `cat2`, `mod`, `pin`, `sat2`. Every operand is a `CustomGeometryValue`. Its numeric branch must be a finite safe integer and normalizes negative zero to zero; its string branch must be non-empty, contain no XML whitespace or invalid XML 1.0 character, and not be a signed decimal integer string. Names use the same single-token lexical contract and must be unique across both lists. The strict codec validates the stored tree; the public evaluator separately resolves dependencies, built-ins, arithmetic domains, and numeric tree values.

`CustomGeometry.handles` is an ordered `CustomGeometryHandle[]` matching the `a:ahLst` choice order. `CustomGeometryXyHandle` has `kind: 'xy'` and maps to `a:ahXY`; its `xGuide`, `yGuide`, `minX`, `maxX`, `minY`, and `maxY` fields map to `gdRefX`, `gdRefY`, `minX`, `maxX`, `minY`, and `maxY`. `CustomGeometryPolarHandle` has `kind: 'polar'` and maps to `a:ahPolar`; `radiusGuide`, `angleGuide`, `minRadius`, `maxRadius`, `minAngle`, and `maxAngle` map to `gdRefR`, `gdRefAng`, `minR`, `maxR`, `minAng`, and `maxAng`. Both kinds require `position`, written as one direct `a:pos`. Position plus XY/radius bounds use direct shape-coordinate safe integers or tokens; numeric angle bounds use direct `1/60000°`. Every guide/bound property is independently optional: the codec does not require refs and bounds together, paired min/max values, resolved guide names, or ordered numeric bounds. Omitted or empty handles normalize to no own property, while mixed XY/polar order and exact optional-property presence are preserved.

`CustomGeometry.connectionSites` is an ordered `CustomGeometryConnectionSite[]` mapped to the direct `a:cxnLst`. Each entry maps to one direct `a:cxn`, requires `angle` mapped to its unqualified `ang` attribute, and requires `position` mapped to one direct `a:pos`. Numeric position coordinates are safe integers written directly in custom-geometry coordinate space; numeric angles are safe integers in direct `1/60000°`. Both fields may instead use the same guide/built-in token branch as `CustomGeometryValue`; the codec preserves those direct tokens and the evaluator resolves them to numbers. Source order and duplicate sites are retained. Neither layer clamps or normalizes angles or verifies attachment to a path. Omitted and empty lists normalize to no own `connectionSites` property.

`CustomGeometry.textRectangle?` is one `CustomGeometryTextRectangle`; required `left`, `top`, `right`, and `bottom` map to `a:rect@l/t/r/b`. Each edge is a `CustomGeometryValue`, so it accepts a finite safe integer or a valid guide/DrawingML built-in token. Omission and explicit `{ left: 'l', top: 't', right: 'r', bottom: 'b' }` both fold to the canonical default and produce no own `textRectangle` property. LibreOffice `textAreaLeft/Top/Right/Bottom` guide normalization is accepted without special coercion.

Creation and replacement strictly normalize dense ordinary arrays and descriptor-safe own data fields, detach the caller, and reject unknown, missing, non-safe, invalid lexical/formula/handle/connection-site/text-rectangle, non-positive numeric extent/radius, or invalid sequence state before mutation. Getter snapshots are detached and deeply frozen through guide lists, formula objects, operand tuples, handle/connection lists and positions, text rectangles, paths, and commands. Empty or omitted guide/handle/connection lists and canonical-default text rectangles normalize to absent optional properties. Assignment replaces the complete custom geometry; assigning the same semantic snapshot is an exact bytes/journal no-op, and `undefined` clear is intentionally not accepted. Setting `presetType` converts custom geometry to preset; setting `customGeometry` converts a preset shape back while preserving model identity, name, transform, styles, text, effects, relationships, and unrelated bytes. The strict reader accepts only namespace-correct direct XY/polar handles and connection sites plus at most one complete direct text rectangle with exact attributes. Malformed state remains byte-preserved, makes the snapshot `undefined`, and rejects replacement before package mutation.

`evaluateCustomGeometry(geometry, context)` evaluates a normalized `CustomGeometry` without reading or mutating a package. `CustomGeometryEvaluationContext` is exactly `{ readonly width: number; readonly height: number }`; both values must be positive safe integers. `ShapeModel.evaluateCustomGeometry()` is the live convenience method: it uses the current shape transform as context, returns `undefined` for preset geometry or an unreadable strict custom snapshot, and otherwise returns the same numeric tree. Resizing the shape therefore changes the next live result without changing the stored formulas.

All 37 DrawingML built-ins are supported: `3cd4`, `3cd8`, `5cd8`, `7cd8`, `b`, `cd2`, `cd4`, `cd8`, `h`, `hc`, `hd2`, `hd3`, `hd4`, `hd5`, `hd6`, `hd8`, `l`, `ls`, `r`, `ss`, `ssd2`, `ssd4`, `ssd6`, `ssd8`, `ssd16`, `ssd32`, `t`, `vc`, `w`, `wd2`, `wd3`, `wd4`, `wd5`, `wd6`, `wd8`, `wd10`, and `wd32`. Adjustment guides are evaluated first in source order, followed by shape guides in source order. A successfully evaluated custom guide shadows a same-name built-in for later values; while defining that guide, a same-name operand still resolves to the built-in. Cycles take precedence over forward-reference reporting, and unknown tokens are distinguished from both.

All 17 formula operators produce JavaScript numbers using DrawingML units. Finite fractional output is preserved; no integer rounding is applied. Negative zero is canonicalized to positive zero. `*/` and `+/` return zero when their denominator is zero. Negative `sqrt`, non-positive evaluated arc radii, and non-finite arithmetic are rejected. Every operand is resolved in source order before the operator is applied.

The evaluator resolves adjustment/shape guide values, handle positions and optional numeric bounds, connection-site positions and angles, text-rectangle edges, and every token-bearing path command field. It preserves handle guide-reference names and path width/height/flags. When `CustomGeometry.textRectangle` is omitted, evaluation always materializes `{ left: 0, top: 0, right: context.width, bottom: context.height }`; omitted optional arrays remain absent. It intentionally does not scale path coordinates from path extents into the shape context, derive arc endpoints, calculate geometric bounds, drag handles, or snap/create connectors.

The full public output type set is `EvaluatedCustomGeometryGuide`, `EvaluatedCustomGeometryPoint`, `EvaluatedCustomGeometryTextRectangle`, `EvaluatedCustomGeometryCommand`, `EvaluatedCustomGeometryXyHandle`, `EvaluatedCustomGeometryPolarHandle`, `EvaluatedCustomGeometryHandle`, `EvaluatedCustomGeometryConnectionSite`, `EvaluatedCustomGeometryPath`, and `EvaluatedCustomGeometry`. The public error surface is `CustomGeometryEvaluationErrorCode` plus `CustomGeometryEvaluationError(code, message, guideName?, token?)`; codes are `unknown-token`, `forward-reference`, `cyclic-reference`, `invalid-domain`, and `non-finite-result`. Standalone geometry/context normalization detaches caller state, evaluation performs no package mutation, and the complete output graph—including context, arrays, entries, commands, and points—is recursively frozen.

PptxGenJS 4.0.1 legal `ShapeType.custGeom` points import as matching final native snapshots: first ordinary point becomes `moveTo`, later ordinary points become `lineTo`, `moveTo: true` starts another subpath, and arc/quadratic/cubic/close map to their direct command branches. Its runtime converts numeric values below 100 and numeric strings as inches, keeps numeric values at or above 100 direct, resolves percentages against the full slide rather than the shape, ignores arc endpoint `x/y`, omits unknown curve kinds, and can emit malformed first-curve/zero-radius/negative-radius/unsafe output. The adapter imports supported final paths and byte-preserves malformed shapes with `customGeometry === undefined`; native does not copy those heuristics or coercions and requires explicit direct values. PptxGenJS 4.0.1 has no public guide-formula, arbitrary adjustment-handle, connection-site, or text-rectangle input and emits only an empty `a:cxnLst` plus the canonical default `a:rect`; evaluator parity is limited to those supported final numeric paths and that default rectangle. Formulas, handles, sites, arbitrary rectangles, and their evaluation are native extensions.

The release gallery was produced from the actual npm tarball and contains 4 slides, 54 shapes, and 22 evaluator targets. Strict reopen evaluates 22/22 targets; the source and LibreOffice round-trip files both pass PowerPoint 2010 validation with 0 errors and 0 warnings, render at 2400×1350 without overflow, and were visually reviewed slide by slide. LibreOffice rewrites direct expressions for 22/22 targets. Numeric path commands plus effective text rectangles match for 21/22; only `Operator sqrt` changes its path endpoint from `600000` to `0`. All 22 guide arrays are rewritten, and the handles and connection sites of the combined handle/site target are removed or rewritten, leaving 21/22 direct matches for each. Both saved files still evaluate strictly for 22/22 targets; these are recorded client rewrites, not hidden as native parity.

`ShapeAdjustment` is the readonly `{ readonly name: string; readonly value: number }` direct guide value used by preset-only `AddShapeOptions.adjustments` and `ShapeModel.adjustments`. A value must be a safe integer and is written directly as the operand of `a:gd@fmla="val N"`; the API performs no shape-specific conversion, clamping, or range inference. Lists must be dense, ordered, and uniquely named. Normalization is descriptor-safe and getter-free, immediately detaches from the caller, and getter snapshots are detached and deeply frozen. Assignment replaces the whole ordered list, `[]` keeps an empty `a:avLst`, and the setter does not accept `undefined`; assigning the same list is an exact bytes/journal no-op. Only one namespace-correct direct preset geometry with one direct adjustment list and simple `val` guides is supported. Complex formulas, duplicate or ambiguous guides, wrong namespaces, unsafe integers, and custom geometry read as `undefined`; replacement throws `ModelParseError` before package changes. A different `presetType` resets adjustments, while same-type assignment preserves their exact bytes. Valid PptxGenJS 4.0.1 `rectRadius`, `angleRange`, and `arcThicknessRatio` output imports as the same final integer list. Native deliberately retains explicit zero and rejects PptxGenJS string coercion, zero truthiness loss, shortcut precedence, ignored thickness-without-angles, and malformed/unsafe passthrough. The separate custom-geometry API above covers paths, guide formulas, handles, connection sites, text rectangles, and numeric evaluation; the preset-only adjustment API intentionally remains a direct-state list.

`ShapeFill` supports direct none or solid sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. `AddShapeOptions.fill` and `AddTextOptions.fill` use this same value. `ShapeModel.fill` returns a detached direct-state snapshot; same-value assignment is an exact no-op, `{ kind: 'none' }` writes direct no-fill, and `undefined` clears only the direct fill choice. Existing gradient, picture, pattern, and group fills survive unrelated edits and can be explicitly replaced or cleared, but their creation remains outside this simple-fill API.

`ShapeLine` supports direct none or a solid sRGB/theme line with optional finite 0–100 transparency, optional 0–1584 point width, and optional `solid | dash | dashDot | lgDash | lgDashDot | lgDashDotDot | sysDash | sysDot` dash. Omitted width/dash materialize as 1pt/solid; zero width remains direct zero. `ShapeModel.line` returns a detached direct-state snapshot; same-value assignment is an exact no-op, `{ kind: 'none' }` writes direct line no-fill, and `undefined` clears only owned width/fill/dash while preserving the line container, arrowheads, joins, extensions, and unrelated attributes. Existing advanced line fills and custom dashes survive unrelated edits and can be explicitly replaced or cleared, but their creation remains outside this simple-line API.

`ShapeArrowType` is `none | arrow | diamond | oval | stealth | triangle`. `ShapeArrows` has optional readonly `begin` / `end` fields and is used by `AddShapeOptions.arrows` and `ShapeModel.arrows`. Inputs and getter results are detached; the setter is a whole replacement, so a missing side clears that endpoint, explicit `none` stays distinguishable from absence, and `undefined` or an empty object clears both. Same-value assignment is an exact no-op. Arrow edits preserve line width/fill/dash, joins, extensions, advanced line state, and unrelated attributes; line clear/edit operations preserve arrows. Arrows-only creation writes a line container with endpoints but does not synthesize color, width, or dash. A unique safe existing endpoint may carry legal `w` / `len` values `sm | med | lg`; type replacement preserves them lexically, but size is not returned or editable. Malformed, duplicate, reversed, wrong-namespace, or unsupported endpoints read as `undefined` and reject arrow mutation without package changes.

`ShapeShadow` is the strict outer/inner direct-state union used by `AddShapeOptions.shadow`, `AddTextOptions.shadow`, and `ShapeModel.shadow`. Both branches accept optional `RichTextColor`, finite `0..1` opacity, `0..100` point blur, `0 <= angle < 360` degrees, and `0..200` point distance; only outer accepts `rotateWithShape`. Omitted fields normalize to black, 0.75, 8pt, 270°, 4pt, and outer rotate false, while every explicit zero remains zero. Inputs and nested colors are detached before mutation; getter snapshots are detached and deep-frozen. Assignment is a whole replacement, same-value assignment is an exact bytes/journal no-op, kind switches replace only the direct shadow child, and `undefined` removes only that child while retaining `effectLst` plus legal glow, preset-shadow, reflection, soft-edge, blur, or fill-overlay siblings. A malformed, ambiguous, wrong-namespace, `effectDag`, or unsafe schema-order state reads as `undefined` and rejects mutation without package changes. PptxGenJS 4.0.1 omitted shadow and `type: 'none'` map to native `undefined`; its `offset` is native `distance`. Native preserves explicit zero, honors outer rotation, emits legal inner XML, supports theme colors, and rejects invalid passthrough instead of copying PptxGenJS's falsy fallback, ignored rotate flag, malformed inner closing tag, or out-of-range output. Generic effect stacks, custom shadow transforms, preset-shadow editing, and shadow APIs for images, tables, charts, media, and other owners remain outside these focused shape/text APIs.

`Hyperlink` is the mutually exclusive `{ readonly url: string; readonly tooltip?: string } | { readonly slide: number; readonly tooltip?: string }` value used by `AddShapeOptions.hyperlink`, `AddTextOptions.hyperlink`, `ShapeModel.hyperlink`, and `RichTextRunStyle.hyperlink`. A URL must be a non-empty XML-safe string; a slide number must be a one-based positive safe integer resolving to a current presentation slide at assignment time. Inputs must be descriptor-safe ordinary or null-prototype objects with exactly one target and no unknown keys. Getter results are detached frozen direct-state snapshots. Tooltip absence remains property absence, while direct empty remains `tooltip: ''`; assignment is a whole replacement, omitted tooltip clears only that attribute, and `undefined` removes the supported whole-shape click element. Same-value assignment is an exact bytes/journal no-op. URL/slide switching reuses an unshared relationship or clones on write when its ID is referenced elsewhere, and clear or replacement garbage-collects only unreferenced relationships. Internal links retain target-part identity while slide insert/delete/reorder changes the reported one-based ordinal; duplicate self-links retarget to the duplicate, and deleting a target removes incoming DrawingML click/hover elements before deleting their relationships. `AddTextOptions.hyperlink` supplies the non-visual click and default run link. `RichTextRunStyle.hyperlink?: Hyperlink | false` overrides that default with an independent run relationship, inherits it when omitted, or suppresses it with `false`; explicit run underline always wins. `ShapeModel.hyperlink` deliberately remains whole-shape-only, while run links are read and edited through `ShapeModel.richText`. Unsupported hover editing, extra action/sound/history state, duplicate/malformed click ownership, or dangling/wrong-type relationships are never guessed. PptxGenJS 4.0.1 materializes omitted tooltip as direct empty and may console-ignore, coerce, duplicate, or dangle invalid runtime targets; its rich outer hyperlink emits broken `rIdundefined` references, while legal rich per-run links omit the shape click and allocate separate relationships. Native accepts its external-run `action=""` but rejects its defects before mutation. External hyperlinks produce the expected portability warning rather than a package error. Table/image/chart/media/group/graphic-frame hyperlink creation, hover links, action-only navigation, and relative/file safety policy remain outside this API.

Arrow size, cap/compound/alignment/join editing, generic/advanced effects, custom shadow transforms, non-shape/text shadow APIs, custom-geometry path scaling/arc endpoint and bounds calculation/handle dragging/connector snapping and creation, text `rectRadius` / `isTextBox` / `breakLine`, advanced line fill/custom dash creation, and percentage positions remain pending. Text-shape simple-line, arrow, simple-shadow, outer-hyperlink, and preset-geometry creation reuse the same codecs below.

Shape kinds include `text`, `shape`, `image`, `table`, `chart`, `graphic-frame`, and `group`. Images expose embedded part URIs and replacement; tables support basic native creation plus rows/cells, cell text, borders, fill, margins, horizontal/vertical alignment, text-direction, and text-fit editing; charts expose cached series and lossless chart XML editing.

### Text-shape direct fill

```ts
const slide = document.addSlide();
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
plain.fill = undefined;
```

`AddTextOptions.fill?: ShapeFill` is normalized before any package mutation. The supported union is exactly direct none or direct solid with a required strict six-digit sRGB/supported scheme color and optional finite `0..100` transparency rounded to `0.001%`. Omitted, runtime-`undefined`, and explicit none creation all produce canonical direct `a:noFill`; solid explicit zero transparency produces direct `a:alpha val="100000"`. Unknown, inherited, accessor, or symbol keys; non-ordinary objects; missing color; invalid lexical colors; and non-finite or out-of-range transparency reject without changing parts, relationships, XML, runtime caches, shape order, or the mutation journal.

Plain/rich text, `addPlaceholder()`, title/body placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share the same renderer and validation contract. The resulting `ShapeModel.fill` is immediately readable and editable. Inputs detach before mutation; snapshots are detached and frozen; duplicate, move, outer rollback, six-format write/reopen, stable identity, sibling isolation, and placeholder-source isolation retain the direct fill. Clearing the live model with `undefined` differs from creation omission: the former removes the direct choice, while the latter intentionally preserves the canonical text-box default no-fill.

PptxGenJS 4.0.1 writes direct no-fill for omitted text fill, omits the direct choice for `{ type: 'none' }`, and omits alpha for explicit zero transparency. Native deliberately preserves explicit none and zero direct intent. Supported solid/scheme/non-zero-alpha output reaches the same final semantics. Gradient/pattern/picture/group text-fill creation remains outside this simple creator. Text outer simple line, arrows, simple shadow, hyperlink, and preset geometry are supported below; `rectRadius`, `isTextBox`, and combined `breakLine` behavior remain pending.

### Text-shape direct simple line

```ts
const outlined = document.addSlide().addText('Outlined text box', {
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '2F5597' },
    transparency: 25,
    width: 2.5,
    dash: 'dashDot',
  },
});
const themed = document.layouts[0].addRichText([{
  runs: [{ text: 'Theme outline' }],
}], {
  line: { kind: 'line', color: { kind: 'scheme', value: 'accent2' } },
});

outlined.line = { kind: 'line', color: { kind: 'scheme', value: 'accent3' } };
outlined.line = { kind: 'none' };
outlined.line = undefined;
```

`AddTextOptions.line?: ShapeLine` is normalized before package mutation. The union is exactly direct none or a solid line with required strict six-digit sRGB/supported scheme color plus optional finite `0..100` transparency, finite `0..1584` point width, and one of eight preset dash tokens. Transparency is rounded to `0.001%`, width to one EMU, and omitted width/dash materialize as 1pt/solid. Omitted, runtime-`undefined`, and explicit none creation preserve canonical direct text-box no-fill; explicit zero transparency writes `a:alpha val="100000"` and zero width writes `a:ln@w="0"`. PptxGenJS-shaped aliases, missing color, invalid range/token, non-ordinary objects, unknown/inherited/accessor/symbol keys, and class instances reject before parts, relationships, XML, caches, shape order, or the mutation journal change.

Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share one renderer. Line stays after geometry and shape fill. The resulting `ShapeModel.line` immediately returns a detached normalized snapshot and keeps the existing read/whole-replace/clear semantics: same value is an exact no-op, none writes direct no-fill, and `undefined` removes owned width/fill/dash while preserving the line container and unrelated state. Duplicate, move, outer rollback, six-format write/reopen, stable identity, sibling isolation, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 emits empty `a:ln` for omitted/none/empty/missing-color text line, omits direct width/dash for their defaults, collapses width/transparency zero, honors nested deprecated `alpha`, and ignores nested `lineDash`. Native uses explicit reversible direct state and rejects permissive aliases/fallbacks. Supported sRGB/theme, non-zero transparency, positive width, and all eight dash tokens reach equivalent final semantics. Gradient/pattern/picture/group line fills, custom dash, cap/compound/alignment/join, and the remaining text geometry shortcuts stay outside this simple-line slice. Text arrows, simple shadow, hyperlink, and preset geometry are supported below.

Release evidence is 1268 passed / 1 skipped tests, performance 1/1 at 553ms, both TypeScript builds, both package builds, an actual 57-file tarball with `textShapeLines: true`, real-Chrome exact immediate/detached/reopen state with zero console/page/network errors, and installed CLI PowerPoint 2010 validation at 0 errors / 0 warnings.

### Text-shape direct arrows

```ts
const arrowed = document.addSlide().addText('Flow', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
});
const themed = document.layouts[0].addRichText([{
  runs: [{ text: 'Theme endpoints' }],
}], {
  arrows: { begin: 'none', end: 'stealth' },
});

arrowed.line = undefined;
arrowed.arrows = { begin: 'oval' };
arrowed.arrows = undefined;
```

`AddTextOptions.arrows?: ShapeArrows` is normalized before package mutation. Optional `begin` / `end` are limited to `none | arrow | diamond | oval | stealth | triangle`; omission, runtime `undefined`, or `{}` creates no direct endpoints, while explicit `none` is retained. PptxGenJS `beginArrowType` / `endArrowType`, deprecated `lineHead` / `lineTail`, empty or invalid tokens, unknown/inherited/accessor/symbol keys, class instances, arrays, and other non-ordinary objects reject without changing parts, relationships, XML, caches, shape order, or the mutation journal.

Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share the existing `ShapeArrows` codec and one text renderer. Direct child order is line fill/dash, `headEnd`, then `tailEnd`. Arrow-only native creation preserves canonical direct `a:noFill`. `ShapeModel.arrows` immediately returns a detached snapshot and uses whole replacement; explicit endpoint `none` survives, a missing side is removed, and `undefined` clears both endpoints. Line and endpoint ownership are independent in both directions. Duplicate, move, outer rollback, asynchronous declarative detachment, all six formats, write/reopen, stable identity, sibling isolation, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 emits an empty line for omitted text line paint and omits no-fill for arrow-only or `{ type: 'none' }` plus endpoints. It ignores empty endpoints and nested/top-level `lineHead` / `lineTail`, while invalid runtime tokens pass through as malformed direct endpoint state. Native keeps canonical no-fill and rejects aliases or invalid tokens before mutation. All six legal endpoint tokens, begin/end/both, explicit `none`, and combined solid line semantics are otherwise compatible; malformed imported endpoints remain losslessly preserved but are absent from the strict snapshot.

Release evidence is 1274 passed / 1 skipped tests, performance 1/1 at 581ms, both TypeScript and package builds, an actual 57-file tarball with `textShapeArrows: true`, real-browser exact immediate/detached/reopen state with zero errors, and installed CLI PowerPoint 2010 validation at 0 errors / 0 warnings. The following section documents text-shape simple shadow support.

### Text-shape direct simple shadow

```ts
const shadowed = document.addSlide().addText('Shadowed heading', {
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
const inner = document.layouts[0].addRichText([{
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

shadowed.shadow = { kind: 'inner', opacity: 0.25 };
shadowed.shadow = undefined;
```

`AddTextOptions.shadow?: ShapeShadow` uses the same strict direct-state normalizer as preset shapes. Outer and inner accept valid sRGB/theme color, finite `0..1` opacity, `0..100pt` blur, `0 <= angle < 360°`, and `0..200pt` distance; only outer accepts boolean `rotateWithShape`. Defaults are black, 0.75, 8pt, 270°, 4pt, and outer rotate false. Every explicit zero is retained. PptxGenJS `type` / `offset` aliases, coercible strings, invalid ranges, unknown/inherited/accessor/symbol keys, arrays, and class instances reject before package mutation.

Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share the same renderer. Direct child order is geometry, fill, line/endpoints, then one `effectLst`. Fill, line, arrows, and effect ownership are independent. The returned `ShapeModel.shadow` immediately exposes a detached deep-frozen snapshot; assignment is a whole replacement, same-value assignment is an exact bytes/journal no-op, and `undefined` clears only the direct inner/outer child while preserving a safe effect-list container plus sibling effects. Duplicate, move, outer rollback, asynchronous declarative detachment, all six formats, write/reopen, stable identity, sibling isolation, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 omits direct effects for omitted shadow and `{ type: 'none' }`; supported outer cases reach equivalent final semantics, and legacy `offset` corresponds to native `distance`. It falls back from runtime zero to defaults, ignores text `rotateWithShape: true`, warning-corrects or loosely converts some invalid type/color/number inputs, and emits a mismatched closing tag for inner shadow. Native deliberately retains zero, supports theme color and rotate true, emits legal inner XML, and rejects invalid input with zero mutation.

Release evidence is 1280 passed / 1 skipped tests, performance 1/1 at 607ms, both TypeScript builds, both package builds, and the declaration build. The actual 57-file tarball reports `textShapeShadows: true` from installed Node, declarations, browser export, and installed CLI checks. Real-browser exact immediate/detached/reopen state has zero console/page/network errors, and installed CLI PowerPoint 2010 validation is 0 errors / 0 warnings.

### Text-shape outer hyperlink

```ts
const plain = slide.addText('Open product', {
  hyperlink: { url: 'https://example.com/product', tooltip: 'View product' },
});
const rich = slide.addRichText([{
  runs: [
    { text: 'Go to ' },
    { text: 'details', style: { underline: false } },
  ],
}], {
  hyperlink: { slide: 2, tooltip: '' },
});

plain.hyperlink = { url: 'mailto:team@example.com' };
plain.hyperlink = undefined;
```

`AddTextOptions.hyperlink?: Hyperlink` is available on plain/rich text, `addPlaceholder()`, placeholder population, slide/layout/master wrappers, and declarative master text/placeholder objects. Normalization is descriptor-safe and getter-free, detaches immediately, and rejects invalid, coercible, or dangling targets before mutation. Creation writes one relationship ID to the direct non-visual `hlinkClick` and each non-empty run that does not override it; empty paragraphs/runs do not create references.

`RichTextRunStyle.hyperlink?: Hyperlink | false` is the direct run surface. An explicit value allocates an independent relationship even when another run has the same target; omission inherits `AddTextOptions.hyperlink` during creation, and `false` suppresses that default. A linked run without an explicit underline gets `u="sng"`; any explicit underline, including `false`, remains authoritative. Strict reads expose only one valid direct namespace-correct run click and preserve omitted versus empty tooltip. Whole-rich-text replacement provides exact no-op, index-stable ID reuse, unique-target in-place update, shared clone-on-write, clear/GC, validation-before-mutation, transaction rollback, and write/reopen behavior.

The returned live `ShapeModel.hyperlink` is the whole-shape click surface; replacement/clear uses reference-aware clone-on-write and garbage collection while preserving run-local links. URL/internal target identity and tooltip state survive slide/layout/master/placeholder/declarative owners, duplicate, move, target delete, rollback, all six formats, and write/reopen. Duplicate self-links retarget to the duplicate, and target deletion removes incoming direct shape/run clicks.

PptxGenJS 4.0.1 writes `tooltip=""` when omitted. Its plain outer output has shape/run links sharing one relationship; rich outer output instead writes broken `rIdundefined` references. Its legal rich per-run form writes only run links and allocates one relationship per run. Native accepts external-run `action=""` but not orphaned/dangling IDs, falsy-underline loss, loose coercion, or console-only failure.

Release evidence is 1303 passed / 1 skipped tests and performance 1/1 at 624ms. Model, SDK, root, and adapter suites are 199/199, 191/191, 13/13, and 80/80; both TypeScript builds, both tsup builds, and declaration build pass. The actual 57-file tarball plus installed Node/declarations/browser/CLI and real Chrome report `richTextRunHyperlinks: true`; Chrome has zero validation/console/page/network errors. The 24-part external deck validates with 0 errors and 8 expected portability warnings; the 20-part internal-only deck validates with 0 errors / 0 warnings. Package inspect, slide listing, and exact part reads pass.

### Text-shape preset geometry

```ts
const text = document.addSlide().addText('Shaped text', {
  shape: 'ellipse',
  line: { kind: 'line', color: { kind: 'scheme', value: 'accent1' } },
});

console.log(text.presetType); // 'ellipse'
text.presetType = 'hexagon';
```

`AddTextOptions.shape?: PresetShapeType` accepts exactly the 178 frozen canonical values in `PRESET_SHAPE_TYPES`; omitted and own data-property `undefined` both create `rect`. Plain/rich text, `addPlaceholder()`, named placeholder population, slide/layout/master methods, and declarative master text/placeholder objects share this contract. The resulting `ShapeModel.presetType` is live immediately. Same-token assignment is an exact bytes/journal no-op that retains existing adjustments; another token replaces only direct geometry with a canonical empty adjustment list.

The field is normalized as an own data property before package mutation. Empty/falsy values, `folderCorner`, `custGeom`, unknown strings, numbers, booleans, coercion, inherited values, and accessors are rejected rather than trimmed, converted, or guessed. Native exposes correct OOXML `foldedCorner`; PptxGenJS 4.0.1 instead exposes invalid `folderCorner` and omits `foldedCorner`. Native compares the 177 common valid tokens semantically but does not copy PptxGenJS's falsy fallback, unchecked passthrough, line-without-line-option exception, or runtime custom-geometry route. Use `addCustomShape()` / `ShapeModel.customGeometry` for custom geometry.

Geometry ownership is independent from transform, fill, line, arrows, shadow, whole-shape/run hyperlinks, text body, and placeholder identity. In particular, `shape: 'line'` selects preset line geometry while `AddTextOptions.line` controls outline style. Duplicate, move, rollback, all six presentation formats, write/reopen, stable identity, and layout/master placeholder-source isolation are covered.

Final verification is 1313 passed / 1 skipped tests and performance 1/1 at 578ms; both TypeScript checks, both tsup builds, and declaration build pass. The actual 57-file tarball (SHA-256 `1412706458c883b9e4dfa3d87e6577ab86df57b78999a2796b2c6e69647be0f9`) reports `textShapePresetGeometry: true` from installed Node/types/browser/CLI. Real Chrome reports the same with zero validation/console/page/network errors. The representative two-slide source validates at PowerPoint 2010 0/0; after LibreOffice save all 17 `(text, presetType)` pairs survive, with 0 errors and two placeholder-owner warnings. Both versions render without overflow and match in slide-by-slide visual review.

This slice does not complete advanced text or full PptxGenJS parity. `rectRadius`, `isTextBox`, `breakLine`, advanced text/table and `tableToSlides`, output/runtime helpers, and the peer-range full-suite audit remain pending.

```ts
const text = document.addSlide().addText('Quarterly results\nQ4 forecast', {
  name: 'Heading',
  x: inches(1),
  y: inches(0.75),
  width: inches(8),
  height: inches(1),
  align: 'center',
  fit: 'shrink',
  paragraphIndent: 12,
  paragraphMarginLeft: 18,
  paragraphMarginRight: 18,
  valign: 'middle',
  vert: 'vert270',
  wrap: true,
});
text.text = 'Updated results\nApproved';
text.textFit = 'resize';
text.verticalAlignment = 'bottom';
text.textDirection = 'wordArtVert';
text.textWrap = false;
```

`addText()` creates plain-text paragraphs with name and transform options. CRLF and CR normalize to LF; consecutive and trailing line breaks remain empty paragraphs. Setting `.text` replaces the visible text using the first paragraph as the style template. Use the structured API below when run styles must remain distinct.

```ts
const rich = document.addSlide().addRichText([
  {
    indent: -12,
    marginLeft: 12,
    marginRight: 12,
    runs: [{ text: 'Direct non-list paragraph margin' }],
  },
  {
    align: 'right',
    rtl: false,
    marginLeft: false,
    marginRight: 18,
    indent: false,
    bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
    level: 2,
    spacing: { before: 6, after: 8, line: { kind: 'multiple', factor: 1.5 } },
    tabStops: [{ position: 2.5, alignment: 'decimal' }],
    runs: [
      { text: 'Revenue\t', style: { bold: true, fontSize: 24 } },
      { text: '+18%', style: { italic: true, lang: 'fr-CA', color: { kind: 'srgb', value: '00A651' }, transparency: 25 } },
    ],
  },
], {
  lang: 'en-US',
  paragraphIndent: 12,
  paragraphMarginLeft: 24,
  paragraphMarginRight: 24,
  rtlMode: true,
});
rich.richText = [{ runs: [{ text: 'Approved', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 50 } }] }];
```

`richText` is an immutable paragraph/run value snapshot. It reads, creates, and replaces each paragraph's alignment, RTL mode, direct left/right margins, signed ordinary indent, Unicode bullet or automatic numbering, list level, paragraph spacing, tab stops, and run styles. `AddTextOptions` supplies paragraph creation defaults; each `RichTextParagraph` can override them, with `rtl: false`, `marginLeft: false`, `marginRight: false`, `indent: false`, `bullet: false`, `level: 0`, `spacing: false`, or `tabStops: false` suppressing the corresponding default. `paragraphMarginLeft` / `marginLeft` and `paragraphMarginRight` / `marginRight` use points from 0 through 4032 and map only direct `pPr@marL` / `pPr@marR`; direct zero remains distinguishable from false/absence. `paragraphIndent` / `indent` maps signed `-4032..4032` points to direct `pPr@indent`: positive values indent the first line and negative values create a hanging indent. New ordinary paragraphs retain canonical direct zero, while `false` or an omitted field in a later rich-text replacement clears the direct value. Ordinary numeric indent and an active bullet cannot share one paragraph because list indentation owns `indent`; `indent: false` can suppress an outer ordinary default for a list paragraph. Numeric left margin has the same conflict on `marL`; right margin remains independent and can coexist with bullets or numbering. Margin names and the indent sign map directly to physical OOXML values and do not swap under RTL. `AddTextOptions.rtlMode` applies to every created plain/rich paragraph; paragraph `rtl` overrides it, and omitting `rtl` in a later `shape.richText` replacement clears the direct `pPr@rtl` value. RTL does not change alignment, `bodyPr@rtlCol`, presentation direction, or run order. Bullets support a custom character and 0–4032pt per-level indent. Numbering supports the 16 PptxGenJS styles, a 1–32767 start value, and indent. List `level` is zero-based from 0–8; nested bullet margin is `indent × (level + 1)`. Spacing uses points for before/after/exact and a factor such as `1.5` for multiple line spacing. Tab stop positions use inches and support left, center, right, and decimal alignment; `[]` is an explicit empty list. Run styles include font family, point size, direct language, bold, italic, sRGB/theme color, and soft breaks. `AddTextOptions.lang` is the plain/rich creation default, while `RichTextRunStyle.lang` overrides a run; omitted creation language uses `en-US`. The getter exposes only a non-empty direct `rPr@lang`, not inherited or alternate language. Setting `richText` preserves text-body metadata and unrelated same-position paragraph properties but intentionally replaces old runs and supported paragraph values; hyperlinks and advanced typography are separate capabilities.

`RichTextRunStyle.transparency` controls only the run's main text fill. It accepts a finite percentage from 0 (opaque) through 100 (fully transparent), rounded to the nearest 0.001%. Explicit zero remains a direct alpha override; omission writes no alpha, and omission in a later `shape.richText` replacement clears the old override. Without an explicit color it applies to the default `tx1` text color. The getter requires one unambiguous direct solid-fill color and alpha transform; it does not infer inherited transparency or reuse alpha from glow, outline, highlight, underline, shape, or table-cell fills.

Text-box `margin` values use points and accept one number, a `[top, right, bottom, left]` tuple, or a named object. `shape.textMargins` reads only direct `bodyPr` overrides as a named object; assigning a scalar/tuple/object replaces the four supported direct sides, while `undefined` or `{}` clears them. This is separate from ordinary paragraph `marginLeft` / `marginRight` / `indent` and list-owned hanging indent. Unlike PptxGenJS 4.0.1's asymmetric-tuple runtime bug, native tuple creation follows the documented TRBL order.

Text-box `valign` accepts `top`, `middle`, or `bottom`; omission creates an explicit middle anchor. `shape.verticalAlignment` reads or replaces only the direct text-body anchor, and assigning `undefined` removes that direct override while preserving margins, autofit metadata, and other text-body content.

Text-box `wrap` accepts a boolean; omission and true create explicit automatic wrapping, while false keeps text on an unwrapped line. `shape.textWrap` reads or replaces only the direct text-body wrapping token, and assigning `undefined` removes that direct override without changing fit, margins, vertical alignment, or text content.

Text-box `vert` accepts `eaVert`, `horz`, `mongolianVert`, `vert`, `vert270`, `wordArtVert`, or `wordArtVertRtl`. Omission writes no direction, so it remains distinct from explicit `horz`. `shape.textDirection` reads or replaces only an exact valid direct `bodyPr@vert`; assigning `undefined` clears that override while preserving unknown direction tokens during unrelated edits. Table-cell direction is a separate capability.

```ts
import {
  emuToInches,
  inches,
  type AddTableCellOptions,
  type AddTableCellInput,
  type AddTableOptions,
} from '@pptx/sdk';

const tableOptions: AddTableOptions = {
  name: 'Revenue table',
  x: inches(1),
  y: inches(1.25),
  width: inches(8),
  height: inches(2.25),
  columnWidths: [inches(2.5), inches(3.5), inches(2)],
  rowHeights: [inches(0.5), inches(0.75), inches(1)],
  align: 'center',
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 1.5,
    style: 'solid',
  },
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent4' },
    transparency: 20,
  },
  margin: { top: 9, left: 18 },
  textDirection: 'vert270',
  valign: 'middle',
};
const headerOptions: AddTableCellOptions = {
  align: 'center',
  fit: 'shrink',
  textDirection: 'vert',
  valign: 'top',
  margin: { top: 4, left: 8 },
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 1,
    style: 'dash',
  },
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
};
const tableRows: readonly (readonly AddTableCellInput[])[] = [
  [
    { text: 'Region', options: headerOptions },
    'Revenue',
    {
      text: 'Growth',
      options: {
        border: {
          bottom: {
            kind: 'line',
            color: { kind: 'srgb', value: '70AD47' },
            width: 2,
          },
        },
        fill: { kind: 'none' },
      },
    },
  ],
  ['East', { text: '$1.2M' }, '12%'],
  [{ text: 'West' }, '$980K', { text: '' }],
];
const createdTable = document.addSlide().addTable(tableRows, tableOptions);
const currentWidthsInches = createdTable.columnWidths?.map(emuToInches);
const currentHeightsInches = createdTable.rowHeights?.map(emuToInches);
createdTable.setColumnWidths([
  inches(2),
  inches(3),
  inches(3),
]);
createdTable.setRowHeights([
  inches(0.75),
  inches(1.25),
  inches(0.5),
]);
createdTable.setRowHeights([0, inches(1), 0]);
createdTable.setCellText(2, 2, '8%');
```

`addTable()` accepts a non-empty, dense, rectangular matrix whose cells are single-paragraph strings or strict plain-text objects shaped as `{ text: string, options?: { align?: TextAlignment; border?: TableCellBorderInput; fill?: TableCellFill; fit?: TextBoxFit; margin?: TextBoxMarginInput; textDirection?: TableCellTextDirection; valign?: TextBoxVerticalAlignment } }`; both forms may be mixed. `AddTableOptions.margin` and cell `options.margin` are point-only and accept one number for all sides, an exact `[top, right, bottom, left]` tuple, or a partial named object. Resolution is per side: canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. A cell scalar/TRBL replaces all table sides, while a partial cell object replaces only supplied sides; omitted, runtime-undefined, or empty cell margin inherits the table layer. Omitted, runtime-undefined, or empty table margin preserves the existing canonical bytes. Each supplied value is rounded to the nearest EMU and must fit signed Int32, so zero and negative direct margins are supported. Cell-level creation `valign` accepts only `top`, `middle`, or `bottom` and emits the selected physical cell's direct `tcPr@anchor` as `t`, `ctr`, or `b`; without a table default, omission or runtime `undefined` writes no anchor and preserves the same bytes as an equivalent string cell. `AddTableOptions.valign` accepts the same exact values and materializes its value onto every cell without a cell-level value; an explicit cell value wins. An omitted or runtime-undefined table value preserves existing bytes. Table- and cell-level `textDirection` accept exactly `horz`, `vert`, `vert270`, or `wordArtVert`. `AddTableOptions.textDirection` materializes its value onto every physical cell whose cell option is omitted or runtime `undefined`; any explicit cell value wins. Explicit cell `horz` therefore blocks a non-horizontal table default while still creating no direct direction attribute. With no inherited non-horizontal value, omitted, runtime `undefined`, and explicit `horz` create no direct `tcPr@vert`, while the other three values write their exact token. Omitted or runtime-undefined table direction preserves existing bytes. `AddTableOptions.fill` and cell `options.fill` use the strict `TableCellFill` whole value: `{ kind: 'none' }` or `{ kind: 'solid', color, transparency? }`, with a six-digit sRGB or theme color and finite 0–100 transparency rounded to 0.001%. A cell solid or none value completely overrides the table value; omitted, runtime-undefined, or empty cell options inherit a supplied table fill. Omitted or runtime-undefined table fill preserves the original bytes, while an empty fill object is invalid. Table align, border, fill, margin, textDirection, and valign defaults are materialized only on physical cells and are not retained as table metadata, so later cell replace/clear operations work on final direct state and never reapply a creation default. The `tcPr` start tag serializes `marL/marR/marT/marB` first, then optional `anchor` and `vert`, followed by L/R/T/B borders and fill. `AddTableOptions.border` and cell `options.border` accept one side value for all four sides, an exact TRBL tuple, or a partial named object. A side is strict `{ kind: 'none' }` or `{ kind: 'line', color, width, style? }`: color is sRGB/theme, width is finite `0..1584` points quantized to the nearest EMU, and style is omitted, `solid`, or `dash`. A non-empty table value is copied only to cells with no effective cell border input; any non-empty cell scalar, TRBL, named value, or explicit none blocks the whole table value. A partial value defines only its own sides, and the remaining sides use the canonical direct no-fill baseline rather than inheriting across layers. Omitted/undefined/empty/all-undefined table border preserves existing bytes; omitted/undefined/empty/all-undefined cell border inherits a supplied table border. With no table border, those absent values and scalar none retain the canonical no-fill bytes. Omitted solid transparency writes no alpha and explicit zero writes `alpha=100000`. Cell, options, margin, border, side, fill, and nested color objects must be ordinary or null-prototype objects containing only supported own data properties. Accessors, inherited values, extra or symbol keys, exotic arrays, non-string text, CR/LF, and invalid XML characters are rejected without invoking a getter. Rows and all nested values are normalized into detached snapshots immediately. Geometry uses EMU: x/y default to 0.5 inch, width defaults to one inch per column, and height defaults to one inch with automatic rows (`a:tr@h=0`); use `inches()` for inch-based geometry. `columnWidths` and `rowHeights` each accept one positive EMU value repeated across the corresponding axis or a dense array whose length exactly matches the column/row count. Values are rounded to safe integers, input arrays are detached, and overall width/height are derived from the vector sum when omitted; when both an overall dimension and its vector are supplied, their rounded values must match exactly. Without a vector, explicit width/height retain their existing exact distribution behavior. Emitted explicit transform dimensions always equal the grid/row sum, while omitted `rowHeights` preserve automatic zero-height rows. New cells return a live `TableModel`; resolved direction, horizontal/vertical alignment, borders, fill, margins, and text fit are immediately visible as detached `TableCell.textDirection`, `TableCell.horizontalAlignment`, `TableCell.verticalAlignment`, `TableCell.borders`, `TableCell.fill`, `TableCell.margins`, and `TableCell.textFit` values, then editable through the physical-cell setters. Cell options other than align/border/fill/fit/margin/textDirection/valign—including merge, hyperlinks, and rich or multi-paragraph text—plus table-level fit creation/defaults, table-level direction/border/margin/valign/fill getters or editors, diagonal/advanced borders, advanced fills, table creation styles, auto-page, repeated headers, and content measurement remain unsupported and are rejected rather than ignored.

`AddTableOptions.align` and `AddTableCellOptions.align` reuse the exact `TextAlignment` values `left`, `center`, `right`, and `justify`, mapped to each created cell's single paragraph direct `a:pPr@algn` tokens `l`, `ctr`, `r`, and `just`, respectively. A table value is materialized onto every cell that omits cell alignment or supplies runtime `undefined`; a valid cell value wins. When the table value is omitted or runtime-`undefined`, current bytes are preserved and no effective left token is synthesized. Final ownership is the physical cell paragraph's direct `a:pPr@algn`, never `tcPr`, `bodyPr`, or retained table metadata, so later cell clearing does not reapply the creation default. The direct token survives supported cell text/property edits, slide duplication, write, and reopen, and is exposed through the strict single-paragraph snapshot/editor described below. Supported table and cell values, including cell precedence, match PptxGenJS 4.0.1 final state; PptxGenJS silently drops an unknown table runtime value, whereas native creation throws `TypeError` before mutation. Rich or multi-paragraph cell alignment remains unsupported.

`TableModel.columnWidths` reads the unique direct `tblGrid` as a detached exact-EMU snapshot. A malformed or ambiguous grid returns `undefined` instead of guessing from the transform or cells. `setColumnWidths()` accepts a positive scalar broadcast or a dense descriptor-safe exact-length array, rounds each item to a safe EMU integer, rejects unsafe sums, and atomically updates both `gridCol@w` and `ext@cx`. A valid grid/transform mismatch is repaired; a numeric no-op preserves the original slide bytes and mutation journal. Unsafe existing grid or transform XML raises `ModelParseError` without mutation. Inherited `setTransform({ width })` still changes only the transform, so use `setColumnWidths()` when changing table width distribution.

`TableModel.rowHeights` reads all direct `tr@h` values as a detached exact-EMU snapshot; zero is a valid automatic row height. A malformed or ambiguous direct row vector returns `undefined` without guessing from transform height or cell content. `setRowHeights()` accepts a non-negative scalar broadcast or a dense descriptor-safe exact-length array, rejects raw negative values, and rounds each item to a safe EMU integer. When every target is positive, their safe exact sum is written to `ext@cy` and a valid rows/transform mismatch is repaired. When any target is zero, row tokens are updated but the already-valid transform height is preserved because the rendered automatic height cannot be derived from the numeric row sum. Numeric no-ops preserve original tokens, slide bytes, and the mutation journal; unsafe existing rows or transform XML raise `ModelParseError` without mutation. Creation remains stricter: explicit `addTable({ rowHeights })` values must be positive, while omitting `rowHeights` creates automatic rows. Inherited `setTransform({ height })` still changes only the transform, so use `setRowHeights()` to edit row tokens. Rich/multi-paragraph cell text and alignment, cell options other than align/border/fill/fit/margin/textDirection/valign, table-level fit creation/defaults and table-level direction/border/margin/valign/fill getters or editors, diagonal/advanced borders, advanced fills, colspan/rowspan or merge editing, row insertion/deletion, table/cell creation styles, auto-page/repeated headers, hyperlinks, and content measurement are not yet supported by native creation.

```ts
import {
  TableModel,
  type TableCellBorders,
  type TableCellFill,
  type TableCellTextDirection,
  type TextBoxFit,
  type TextBoxMargins,
  type TextBoxVerticalAlignment,
  type TextAlignment,
} from '@pptx/sdk';

const table = document.slides[0].shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
const current: TableCellTextDirection | undefined =
  table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, 'vert270');
table?.setCellTextDirection(0, 1, 'wordArtVert');
table?.setCellTextDirection(0, 2, 'horz');
table?.setCellTextDirection(0, 2, undefined);
const currentHorizontalAlignment: TextAlignment | undefined =
  table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'left');
table?.setCellHorizontalAlignment(0, 1, 'center');
table?.setCellHorizontalAlignment(0, 2, 'right');
table?.setCellHorizontalAlignment(0, 3, 'justify');
table?.setCellHorizontalAlignment(0, 3, undefined);
const currentFit: TextBoxFit | undefined = table?.rows[0]?.cells[0]?.textFit;
table?.setCellTextFit(0, 0, 'shrink');
table?.setCellTextFit(0, 1, 'resize');
table?.setCellTextFit(0, 2, 'none');
table?.setCellTextFit(0, 2, undefined);
const currentAlignment: TextBoxVerticalAlignment | undefined =
  table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, 'top');
table?.setCellVerticalAlignment(0, 1, 'middle');
table?.setCellVerticalAlignment(0, 2, 'bottom');
table?.setCellVerticalAlignment(0, 2, undefined);
const currentMargins: TextBoxMargins | undefined = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, 7.2);
table?.setCellMargins(0, 1, [3.6, 7.2, 10.8, 14.4]);
table?.setCellMargins(0, 2, { top: 4, left: 8 });
table?.setCellMargins(0, 2, undefined);
const currentBorders: TableCellBorders | undefined = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: 'FF0000' },
  width: 2,
  style: 'solid',
});
table?.setCellBorders(0, 1, [
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1, style: 'dash' },
  { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 1.5 },
  { kind: 'none' },
  undefined,
]);
table?.setCellBorders(0, 2, { top: { kind: 'none' } });
table?.setCellBorders(0, 2, undefined);
const currentFill: TableCellFill | undefined = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'srgb', value: 'FF0000' },
});
table?.setCellFill(0, 1, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
});
table?.setCellFill(0, 2, { kind: 'none' });
table?.setCellFill(0, 2, undefined);
```

Table creation names the options `AddTableOptions.textDirection` and `AddTableCellOptions.textDirection`; both accept only `horz`, `vert`, `vert270`, or `wordArtVert`. The table value is materialized onto every physical cell whose cell value is omitted or runtime `undefined`, while any explicit cell value wins. Explicit cell `horz` therefore blocks a non-horizontal table default but still creates no direct attribute. Omitted, runtime-undefined, and resolved horizontal creation write no `tcPr@vert`; the three non-horizontal values write their exact token, never `bodyPr@vert`, and appear immediately in the live `TableCell.textDirection` snapshot. Omitted/runtime-undefined table direction preserves existing bytes, and explicit table `horz` produces the same direction bytes when cells do not override it. Inputs follow the same descriptor-safe, getter-free, detached normalization as other table options, and invalid or text-box-only tokens are rejected before mutation. Creation retains only final physical-cell direct state, not table metadata, so a later clear never re-inherits the original table value. `setCellTextDirection()` addresses physical zero-based row/cell positions and deliberately retains lossless editor semantics: explicit `horz` writes `vert="horz"`, while `undefined` clears the direct attribute. The getter requires one direct `tcPr`, one unqualified `vert`, and one exact public token; no table-level getter/editor or inheritance resolution is retained. PptxGenJS 4.0.1 produces the same supported table/cell final state and precedence, including horizontal collapse; its runtime invalid-token passthrough remains unsupported by the strict native API. This four-value table-cell API is separate from the seven-value text-box API.

Table-cell `horizontalAlignment` is an immutable `TextAlignment` snapshot for a strict cell containing exactly one direct text body, one direct paragraph, and one direct paragraph-properties element. It reads only the unqualified direct `pPr@algn` tokens `l`, `ctr`, `r`, and `just` as `left`, `center`, `right`, and `justify`; missing, malformed, ambiguous, zero-paragraph, or multi-paragraph cells return `undefined` without resolving effective alignment or a retained table default. `setCellHorizontalAlignment()` uses physical zero-based row/cell coordinates, writes the canonical direct token, and uses `undefined` to clear only that unqualified direct attribute. Assigning the current value or clearing an absent token is an exact no-op. Unsafe cell structure raises `ModelParseError` without mutation, while unrelated paragraph content, namespaced attributes, cell properties, and neighbor cells remain byte-preserved. PptxGenJS 4.0.1 table/cell alignment is materialized into the same direct cell-paragraph state and imports to the same snapshots; native existing-deck editing is a lossless extension because PptxGenJS has no existing-deck editor. Rich or multi-paragraph cell alignment remains unsupported.

Table-cell creation names the option `AddTableCellOptions.valign`; the immutable snapshot is `TableCell.verticalAlignment`. Both reuse exact `TextBoxVerticalAlignment` values `top`, `middle`, and `bottom`, stored only in the selected physical cell's direct `tcPr@anchor` as `t`, `ctr`, or `b`. Omitted or runtime-undefined cell values write no anchor unless `AddTableOptions.valign` supplies the same strict value as a creation default; a cell value overrides that default. The table value is materialized directly onto uncovered physical cells and is not retained as table metadata, so a later `setCellVerticalAlignment(..., undefined)` clears the anchor without inheritance. Omitting the table value preserves the equivalent default bytes. The returned live table exposes every resolved value immediately, and `setCellVerticalAlignment()` writes the canonical token or clears it with `undefined`. The strict snapshot getter requires exactly one direct `tcPr` and one unqualified supported anchor. It does not read or change the separate `bodyPr@anchor`, resolve inheritance, or support `just` / `dist`. PptxGenJS 4.0.1 produces the same final direct state for valid table- and cell-level values, including cell precedence and no anchor when both layers are omitted. Runtime invalid PptxGenJS tokens remain opaque and import as `undefined`, while the native API rejects them. A table-level getter/editor remains unsupported.

Table-cell `margins` reuses the point-based `TextBoxMargins` value shape but owns only direct `tcPr@marL/marR/marT/marB`; it never reads or changes text-box `bodyPr@*Ins`. `AddTableOptions.margin`, `AddTableCellOptions.margin`, and `setCellMargins()` accept a point scalar, `[top, right, bottom, left]` tuple, partial named object, `{}` or `undefined`, with descriptor-safe getter-free normalization and immediate detachment. Creation resolves canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. Cell scalar/TRBL replaces all four inherited sides; partial cell input replaces only supplied sides; omitted/undefined/empty cell input inherits the table layer. Omitted/undefined/empty table input preserves canonical bytes. The resolved values are materialized directly on physical cells and immediately exposed by `TableCell.margins`; no table metadata or getter/editor is retained, so clearing a cell does not reapply the creation default. The cell editor addresses zero-based physical row/cell positions and is a whole replacement, so its omitted named sides are cleared. Values quantize to signed-Int32 EMU and may be zero, negative, positive, or fractional points. The snapshot getter requires exactly one direct `tcPr`, reads each unique unqualified signed-Int32 integer independently, and returns only valid direct sides. Creation writes `marL/marR/marT/marB`, then optional `anchor`, followed by L/R/T/B borders and fill. PptxGenJS 4.0.1 writes the same narrow defaults and materializes table-level values into cells, but its legacy runtime treats the first value `<1` as inches and `>=1` as points. Thus native table or cell `7.2` equals PptxGenJS `0.1`, native `[3.6, 7.2, 10.8, 14.4]` equals PptxGenJS `[0.05, 0.1, 0.15, 0.2]`, while native `0.1` intentionally means 0.1pt. Adapter imports expose the resulting OOXML in points instead of guessing the original unit.

Table-cell `borders` is a detached partial snapshot of the selected physical cell's same-prefix direct `lnL/lnR/lnT/lnB`. Each supported side is `{ kind: 'none' }` or `{ kind: 'line', color, width, style? }`, where width is finite `0..1584` points, color is strict sRGB/theme, and style is omitted, `solid`, or `dash`. `AddTableOptions.border`, `AddTableCellOptions.border`, and `setCellBorders()` accept one scalar border, an exact `[top, right, bottom, left]` tuple, or a partial named object. During creation, a non-empty table value is materialized only onto cells whose normalized cell border is absent. Any non-empty cell scalar, TRBL, named value, or explicit none blocks the entire table value; missing sides in that cell value remain canonical none and do not inherit table sides. Empty/all-undefined cell values inherit a supplied table value, while empty/all-undefined table values preserve omitted bytes. Creation overlays the chosen value on four canonical direct no-fill sides; the editor is intentionally a whole replacement, so its missing sides are removed and `{}` / `undefined` clear all four direct sides without re-inheritance. Explicit none writes a direct zero-width `noFill` line. Public TRBL input is serialized in OOXML's required L/R/T/B order before cell fill; dash maps to `sysDash`, explicit solid maps to `solid`, omitted style writes no direct dash token, zero-width line remains a line, and widths are quantized to the nearest EMU. Border input at every layer is descriptor-safe, getter-free, null-prototype-compatible, and detached. The strict getter omits malformed or unsupported sides independently and never treats diagonals, cell fill, text outlines, advanced dash presets, joins, arrows, or shared-edge/effective style as supported state; unrelated edits preserve them. For supported table/cell scalar or TRBL none, sRGB line, solid/dash, scalar zero width, and whole-value cell overrides, native creation matches PptxGenJS 4.0.1 final direct state. Native `{}` means no effective input while PptxGenJS `{}` defaults to `666666`/1pt/solid and a cell `{}` blocks its table value; native omitted style writes no direct dash while PptxGenJS defaults it to direct solid. Native scalar and TRBL zero widths stay zero, whereas PptxGenJS table TRBL `pt: 0` changes to 1pt; PptxGenJS also accepts short runtime tuples and materializes missing sides as none, while native tuples must be dense and exactly four items. Native named sides and theme border colors are strict extensions. Adapter snapshots reflect final XML rather than guessing the original input layer. A table-level border getter/editor, diagonal borders, transparency, advanced line properties, shared-edge precedence, and layout recomputation remain unsupported.

Table-cell `fill` is a detached direct-state snapshot with explicit `{ kind: 'none' }` and `{ kind: 'solid', color, transparency? }` variants. Solid fill supports strict six-digit sRGB or the existing theme-color tokens; transparency is a finite `0..100` percentage quantized to `0.001%`. Omitted transparency writes no alpha, while explicit zero writes direct opaque alpha. During native creation, `AddTableOptions.fill` is a whole-value default materialized into every physical cell whose `AddTableCellOptions.fill` is omitted or `undefined`; a cell solid or none value completely overrides it. Table and cell fill/color inputs are descriptor-safe, getter-free, ordinary/null-prototype-only, and deeply detached immediately. Omitted or runtime-undefined table fill preserves the original bytes; `{}` is invalid. The returned live table immediately exposes detached resolved `TableCell.fill` values. `setCellFill()` then addresses only zero-based physical row/cell direct state: `none` writes `tcPr/a:noFill`, while `undefined` removes the direct fill choice and never re-inherits the creation default. The strict getter requires a unique direct `tcPr` and one unambiguous same-prefix noFill or solidFill; it never reads border/text descendant fills or resolves table styles. Unsupported gradient/pattern/picture/group fills remain preserved during unrelated edits and can be explicitly replaced or cleared; a table-level fill getter/editor is not supported. PptxGenJS 4.0.1 produces the same final direct cell state for supported table solid fills and cell solid overrides, but collapses omitted and `type: 'none'` to no direct fill and explicit zero transparency to no alpha. Native direct none and explicit zero therefore remain intentionally distinguishable. PptxGenJS may also emit invalid alpha for out-of-range runtime values; adapter imports preserve that XML but do not fabricate a valid snapshot.

Table-cell `textFit` reuses the immutable `TextBoxFit` values `none`, `shrink`, and `resize`, backed only by the selected physical cell's direct `txBody/bodyPr` fit choice. Native plain-text cell creation accepts `AddTableCellOptions.fit` at cell level only: omitted, runtime `undefined`, and `none` produce byte-identical self-closing `bodyPr` and an immediate direct snapshot of `undefined`; `shrink` and `resize` write direct `normAutofit` and `spAutoFit` and snapshot immediately. Creation never writes `noAutofit`, retains table metadata, measures text, calculates a final font scale, or recomputes cell/table geometry. The getter requires exactly one direct text body, one direct body properties element, and one unambiguous supported fit child; it reads an existing `noAutofit` as `none` and otherwise returns `undefined` for absent or malformed state. `setCellTextFit()` uses physical zero-based row/cell positions: `shrink` and `resize` write `normAutofit` and `spAutoFit`, while `none` and `undefined` both clear the direct choice without restoring creation input or creating `noAutofit`. Reassigning the current shrink/resize mode preserves any calculated `fontScale` and `lnSpcReduction`. Fit remains independent from text direction in `tcPr@vert`. Table-level fit creation/defaults are not supported. PptxGenJS 4.0.1 has no table-cell fit API and ignores runtime `fit`, `autoFit`, and `shrinkText` values supplied to table/cell options, so native shrink/resize creation is an explicit extension rather than a parity claim.

## Native charts

`CHART_TYPES` is the frozen public catalog of `area`, `bar`, `bar3D`, `bubble`, `doughnut`, `line`, `pie`, `radar`, and `scatter`. `SlideModel.addChart()` and `PptxDocument.addChart(slideIndex, ...)` expose the same overloads:

```ts
addChart(type: ChartType, series: readonly ChartSeriesInput[], frame?: AddChartOptions): Promise<ChartModel>;
addChart(groups: readonly ChartGroupInput[], frame?: AddChartOptions): Promise<ChartModel>;
```

`ChartSeriesInput` requires a non-empty `name` and finite non-empty `values`. Categorical groups require equal-length `categories`; scatter requires equal-length `xValues`; bubble additionally requires equal-length positive `sizes`. Pie and doughnut accept one series. Compatible bar/area/line groups may use `axis: 'primary' | 'secondary'`; secondary-only, axis-free, 3D, bubble, and incompatible XY/category combinations reject before part or ID allocation. `AddChartOptions` owns frame metadata and EMU/OOXML-angle geometry: `name`, `altText`, `x`, `y`, `width`, `height`, `rotation`, and flips.

Every native chart receives a chart part, an internal package relationship, and a deterministic embedded XLSX. One normalized workbook plan drives worksheet cells, chart formulas, and string/numeric caches. Workbook bytes are generated before the synchronous presentation transaction, so invalid input or workbook work leaves the package unchanged. All six presentation formats use the same creation and reopen path.

`ChartModel.definition` returns a detached, recursively frozen `ChartDefinition` with ordered groups, series, axes, and normalized options. Root options cover language/style, rounded corners, blank handling, title, legend, chart/plot areas, primary and secondary axes, data table, colors, and 3D view. Group options cover data labels, per-series fill/line/marker, and type-specific grouping, direction, gaps, overlap, first-slice angle, hole size, line/scatter/radar markers and smoothing, bubble scale/sign/size representation, and 3D depth. Set root or group options with `replaceDefinition()` after creation; `replaceSeries()` is the shorthand for replacing data on a one-group chart.

`replaceDefinition()` is a synchronized semantic operation: it updates supported chart-owned XML, A1 formulas, caches, and workbook cells together; option-only changes preserve workbook bytes; a proven equal recognized definition is an exact no-op. Existing unsupported extensions, style/color parts, and unowned chart XML remain preserved. Imported shared targets use relationship-aware clone-on-write and keep the selected model identity stable. `remove()` / `slide.deleteChart(shapeId)` removes the frame and relationship and collects only unreferenced owned chart/workbook/style/color parts. `setXml()` remains the explicit raw escape hatch and can intentionally create workbook/cache divergence.

`chart.diagnostics()` returns `CHART_RELATIONSHIP_INVALID`, `CHART_STRUCTURE_UNSUPPORTED`, `CHART_STRUCTURE_AMBIGUOUS`, `CHART_CACHE_INVALID`, `CHART_AXIS_INVALID`, `CHART_WORKBOOK_MISSING`, `CHART_WORKBOOK_CACHE_DIVERGENCE`, or `MODERN_CHART_EXTENSION`. Recognized charts require exact formula/cache/workbook agreement. A standard chart with valid caches but no `externalData` is `cache-only`: opaque client formula placeholders are ignored because there is no workbook to address, diagnostics emit the missing-workbook warning, and the first semantic replacement creates a canonical workbook. Office 2016 `cx:*` charts, external workbooks, and chart animations remain preservation-only.

PptxGenJS 4.0.1 conformance covers all nine public chart types, bar+line primary/secondary combinations, data/formulas/caches/XLSX relationships, and representative title/legend/axis/label/data-table/series/type-specific options. The actual package passes Node, real-Chrome, declaration, and installed-CLI smoke with `nativeCharts: true`; the 11-slide gallery has ten charts/workbooks, zero owned orphans, 0 PowerPoint-2010 errors/warnings, and zero 180-DPI overflow. LibreOffice retains chart types and cached data but removes workbooks on save; these reopen as cache-only warnings and can be canonicalized by semantic editing. Its `bar3D` renderer is title-only for both native and PptxGenJS controls. Built-in trendline/error-bar creation and modern inspection remain in the advanced-charts plugin.

Text-box `fit` accepts `none`, `shrink`, or `resize`. Omission and none write no autofit child for PptxGenJS and PowerPoint 2013 compatibility; shrink writes `normAutofit`, and resize writes `spAutoFit`. `shape.textFit` reads only a unique direct choice, including an existing explicit `noAutofit` as none. Assigning none or `undefined` clears the direct choice; reassigning the current shrink/resize mode preserves any PowerPoint-calculated scale attributes. Final shrink factors may be calculated by PowerPoint only after editing text or resizing the shape.

Slide and shape model objects have stable identity within a document: repeated collection reads and slide reordering return the same member instances. Their properties remain live and read the current OOXML rather than a cached snapshot. Master, layout, and theme models follow the same rule within `document.masterLayoutTheme`.

`duplicateSlide()` deep-clones owned dependencies such as charts and their embedded workbooks while retaining shared image, media, and layout targets. `deleteSlide()` garbage-collects only unreferenced owned dependency subgraphs; opaque and shared targets are preserved.

`ImageModel.replaceData()` and `ChartModel.setXml()` edit an exclusive target in place. If a target or relationship is shared, the edited shape is redirected to a private clone; chart clones include owned workbook/style/color dependencies. The operation is atomic and keeps the shape model object identity stable.

## Transactions

```ts
document.transaction((draft) => {
  draft.slides[0].title.text = 'Updated';
  draft.duplicateSlide(0);
});
```

Transactions are synchronous and nestable. A thrown error or package validation failure restores parts, content types, relationships, ZIP entries, and the mutation journal to the transaction savepoint. Complete asynchronous preparation before entering the callback.

## Slide backgrounds, gradients, and transparency

```ts
import type {
  SetSlideBackgroundImageOptions,
  SimpleFill,
  SlideBackground,
  SlideBackgroundImage,
} from '@pptx/sdk';

const slide = document.slides[0];
slide.background = {
  kind: 'linear-gradient',
  angle: 45,
  stops: [
    { offset: 0, color: '#2563EB' },
    { offset: 1, color: '#7C3AED', alpha: 0.65 },
  ],
};

slide.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
};
await document.setSlideBackgroundImage(0, './background.png');
slide.background = { kind: 'none' };
slide.background = undefined;
```

`SlideBackground` is `SimpleFill | GradientFill | SlideBackgroundImage`. `SimpleFill` is explicit `{ kind: 'none' }` or a solid sRGB/theme `RichTextColor` with optional `0..100` transparency. Gradients retain sRGB, scRGB, scheme, system, or preset sources and their ordered OOXML transforms. An image contains `kind: 'image'`, `image/png | image/jpeg | image/gif`, and detached `Uint8Array` bytes. Getter results are detached and frozen where applicable; image bytes are copied on each read.

Only the direct `p:sld/p:cSld/p:bg/p:bgPr` state is projected. `undefined` removes direct `p:bg` and restores layout/master inheritance; `{ kind: 'none' }` writes legal direct `a:noFill`. The reader returns `undefined` for unsupported or ambiguous `p:bgRef`, pattern/group fill, wrong namespaces, multiple choices, external/dangling image relationships, or malformed content without modifying the package. Explicit supported replacement canonicalizes an opaque direct background while preserving unrelated slide content. Same-value supported assignment and clearing an absent direct background are exact no-ops.

`setSlideBackgroundImage(slideIndex, source, options?)` accepts `RasterImageSource`. `SetSlideBackgroundImageOptions` contains only optional `contentType` assertion and `AbortSignal`. The resolver accepts Node path, HTTP/HTTPS and browser-relative URL, strict data URI, `Uint8Array`, `ArrayBuffer`, Blob/File, Web stream, and async byte iterable; signature validation and all async I/O finish before mutation. Duplicates initially share an internal image target, a different write clones on first mutation, and replacement/clear/slide deletion remove only relationships and media parts that have no remaining incoming reference.

PptxGenJS 4.0.1 solid, transparency, and PNG background output imports to the same supported state. PptxGenJS `{ type: 'none' }` writes no direct background, so it imports as inherited; `{ type: 'none', color }` emits an empty `p:bgPr`, which the strict reader treats as unsupported. Native explicit none intentionally writes legal `a:noFill` instead of reproducing either behavior. `SlideLayoutModel.background` and `SlideMasterModel.background` use the same supported direct owner-aware state. `p:bgRef` semantic editing, pattern/group fill, and image crop/tile/effects remain outside this API. The separate transient slide default text color is documented below.

Packed Node, real-Chrome, declaration, and installed-CLI smoke report `slideBackgrounds: true`. Two clean builds have an identical 48-file dist manifest (`e42633dfd50e9f8731e780f6b911f691845c530f5c1a0b9e5f356f93a1a0f423`). Full Vitest is 1156 passed with one performance test skipped by default; its separate performance gate is 1/1. The native gallery has 11 slides, 41 parts, 39 relationships, and three background media parts; both native and the seven-slide PptxGenJS control validate 0/0 and render without overflow. LibreOffice preserves slide order and every image payload hash but normalizes explicit no-fill to inheritance and gradient rotation/fill-rectangle metadata. Local PowerPoint automation returned `-9074`, so no PowerPoint round-trip pass is claimed.

## Slide numbers and presentation start number

```ts
import { inches, type SlideNumberOptions } from '@pptx/sdk';

const options: SlideNumberOptions = {
  x: inches(8.1),
  y: inches(5),
  width: inches(1.4),
  height: inches(0.35),
  align: 'justify',
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

document.firstSlideNumber = 5;
document.slides[0].slideNumber = options;
document.layouts[0].slideNumber = { x: inches(0.5), align: 'left' };
document.masters[0].slideNumber = { x: inches(4.5), align: 'center' };
document.slides[0].slideNumber = undefined;
```

The public value types are `SlideNumber`, `SlideNumberOptions`, `SlideNumberColor`, `SlideNumberMargins`, `SlideNumberMarginInput`, `SlideNumberTextStyle`, and `SlideNumberTextStyleOptions`. Geometry uses integer EMU; margins and font size use points; transparency uses `0..100` percent. The default normalized field is x/y 0, width 800000, height 300000, left aligned, LTR, and `en-US` non-bold/non-italic text. sRGB values are six-digit uppercase hex; scheme colors use supported DrawingML tokens. Inputs are closed, descriptor-safe, fully validated before package access, copied, and deeply frozen on read.

`SlideModel.slideNumber`, `LayoutModel.slideNumber`, and `MasterModel.slideNumber` project only one namespace-correct direct `p:ph type="sldNum"` field in the owner part. Layout and master values do not mutate child slides, and slide assignment does not create hidden fields in the unique master or default layout; the master setter additionally owns direct `p:hf@sldNum`. Equal semantic assignment, absent clear, and same start-number assignment are exact part/relationship/graph/journal no-ops. A unique recognized supported value is patched locally; one unambiguous opaque placeholder may be canonicalized by explicit assignment. Multiple placeholders, duplicate shape IDs, malformed fields, wrong namespaces, extra runs/paragraphs, or unsupported lexical values read as `undefined` and are rejected before unsafe mutation.

`CreatePresentationOptions.firstSlideNumber` and `PresentationModel.firstSlideNumber` accept only signed Int32 safe integers. `undefined` removes direct `p:presentation@firstSlideNum`, whose OOXML default is 1. Slide cached text is `(firstSlideNumber ?? 1) + zeroBasedIndex`; layout/master cached text is `‹#›`. Start changes and slide duplicate/move/delete synchronously update safely recognized direct slide caches inside the lifecycle transaction. The three warning diagnostics are `SLIDE_NUMBER_SHAPE_ID_COLLISION`, `SLIDE_NUMBER_MASTER_DISABLED`, and `SLIDE_NUMBER_CACHE_NONCANONICAL`.

PptxGenJS 4.0.1 public output imports its default/full style, sRGB/scheme color, left/center/right/justify alignment, top/middle/bottom vertical alignment, scalar/four-side margin, font, language, bold, italic, and transparency cases. Native corrects several observed defects instead of imitating them: fixed shape id 25 can collide, zero width/height fall back truthily, RTL and transparency are not faithfully emitted in some cases, layout/master use noncanonical caches, and the generated master placeholder is disabled. Native uses unique IDs, legal positive extents, canonical caches, and an enabled master flag.

The actual 54-file tarball contains 51 `dist` files and passes installed Node, real-Chrome, browser conditional-export, declaration, and CLI checks with `slideNumbers: true`. Two package builds produce an identical 51-file manifest (`3d77e6f56b8f299f2d580112fd0ebe77d0a98c38c07764259a3735064d5f9bea`). The focused suite is 448/448; full Vitest is 1194 passed with one performance test skipped by default, and the separate performance gate is 1/1. The 16-slide native gallery contains 48 parts and 45 relationships and validates 0/0; the 16-slide PptxGenJS control contains 82 parts and 95 relationships and validates with zero errors plus four expected warnings. All 32 pages render at 180 DPI without edge contact and were inspected individually.

LibreOffice 26.8 renders direct slide fields but displays its own 1..16 sequence instead of the native start 5. Its saved package retains slide order/titles, 15 direct slide owners, the cleared direct field, field types, alignment, and principal explicit styling. It removes `firstSlideNum` and rewrites cached text, default fonts/languages, and layout/master placeholder state, producing zero errors and 15 normalization warnings. Prefer direct slide fields when LibreOffice-visible portability matters instead of relying only on layout/master placeholders. Local PowerPoint 16.112 automation started the application but yielded no active presentation or saved/rendered output, so this environment does not establish a PowerPoint round-trip pass.

## Slide default text color

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

`SlideModel.color` has type `Readonly<RichTextColor> | undefined`. The setter accepts only the closed sRGB/theme union used by rich text: sRGB is normalized to uppercase six-digit hex without `#`, and scheme values must be supported DrawingML tokens. Inputs are descriptor-safe and detached; getters are frozen snapshots. Equal assignment is an exact package/journal no-op. `undefined` clears the transient state.

The value affects only subsequently created `addText()` / `addRichText()` runs. A run's explicit `style.color` takes precedence, while a run with transparency but no local color inherits the slide default. Changing or clearing `slide.color` does not scan or recolor existing shapes or tables. Tables, masters, layouts, and placeholders do not inherit this transient state; named master/layout/placeholder population is complete, while full theme text cascade and advanced-table color behavior remain separate advanced-text/table work.

The transient default is not serialized because OOXML has no legal direct slide-level default-text-color field. Creation materializes the chosen color into each run's standard `a:solidFill`. Reopened documents therefore have `slide.color === undefined`, while materialized run colors remain readable and visually unchanged. Duplicate copies the current state, move retains it, delete removes it, URI reuse cannot leak it, and all operations preserve sibling isolation and transaction rollback.

All six presentation formats, valid PptxGenJS 4.0.1 sRGB/theme/override output, Node, browser, declarations, and CLI are covered. The focused run is 10 passed / 409 skipped; full Vitest is 1205 passed / 1 skipped, performance is 1/1 at 998ms, and typecheck/build pass. The actual tarball contains 54 files and 51 `dist` files; its dist manifest SHA-256 is `467d87ffea6994355c357dbad3b1ea18afa8538b1bacb85b6de43de90ad16829` and tarball SHA-256 is `6812000a83247fdf2d63eddf81ec6ffb43c721d478e4cdcbbf4c4a3ce2b65ad1`.

The real-Chrome result exactly matches live inherited/override/transparency/duplicate state, reopened defaults are absent, and console/page/network errors are zero. Native and PptxGenJS galleries contain 11/9 slides, 38/52 parts, and 35/58 relationships; both validate with 0 errors / 0 warnings under the PowerPoint 2010 profile. All 20 pages were inspected at 180 DPI with zero overflow and 106px minimum margins. LibreOffice 26.8 retains slide and text order, custom sRGB, theme, override, and 40% transparency, while normalizing native `tx1` to equivalent `dk1`; the saved file remains 0/0. Local PowerPoint 16.112 returned `-9074` for both native and control inputs without loading or producing PPTX/PDF output, so no PowerPoint round-trip pass is claimed.

## Master, layout, placeholder, and theme

```ts
import { inches, PLACEHOLDER_TYPES, PptxDocument } from '@pptx/sdk';

const document = PptxDocument.create({ slideSize: 'wide' });
const layout = await document.defineSlideMaster({
  title: 'BRAND',
  background: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
  },
  margin: [inches(0.5), inches(0.5), inches(0.5), inches(0.5)],
  slideNumber: { x: inches(11), width: inches(1) },
  objects: [
    { kind: 'rect', options: { x: 0, y: 0, width: inches(13.333), height: inches(0.2) } },
    {
      kind: 'placeholder',
      text: 'Presentation title',
      options: {
        name: 'title_box',
        type: 'title',
        index: 101,
        x: inches(1),
        y: inches(1),
        width: inches(8),
        height: inches(1),
      },
    },
  ],
});

const slide = document.addSlide({ masterName: layout.name });
slide.addText('Quarterly results', { placeholder: 'title_box' });
console.log(PLACEHOLDER_TYPES); // title, body, pic, chart, tbl, media
```

`PptxDocument.defineSlideMaster(options)` is asynchronous because image, image-background, and chart definition objects may require source/workbook preparation. `DefineSlideMasterOptions.title` is a unique presentation-wide layout name. Optional `master` selects an attached same-document `SlideMasterModel`; omission uses the first safe attached master. Optional background accepts every supported direct background plus `{ kind: 'image-source', source, contentType?, signal? }`. `margin` is a non-negative EMU scalar or exact `[top, right, bottom, left]` tuple bounded by the current slide size. `slideNumber` uses the ordinary direct layout field options.

`SlideMasterObject` is the closed ordered union `rect | line | text | placeholder | image | chart`. Rect and line use `AddShapeOptions`; text accepts a string or rich-text paragraphs; image accepts every portable `ImageSource`; chart accepts normalized chart groups. Placeholder options require an XML-safe unique name and one of the six frozen `PLACEHOLDER_TYPES`, with an optional unique integer index from 0 through 4,294,967,294. All asynchronous work completes before one synchronous package transaction, so source, chart, relationship, part, XML, or definition failure has no observable partial mutation.

PptxGenJS calls the definition a master, but native `defineSlideMaster()` deliberately creates a named layout under a real parent master. Correspondingly, `addSlide({ masterName })` strictly resolves exactly one attached `SlideLayoutModel.name`. Unknown/duplicate names fail instead of falling back. Ordinary layout objects stay on the layout and are inherited. Layout placeholder prompts stay on the layout, while slide creation materializes empty owners with the same unique `{ type, index }` and geometry. A selector is either the unique layout placeholder name or a `PlaceholderIdentity`. Text/rich-text and shapes fill title/body owners, image/SVG fills `pic`, charts fill `chart`, tables fill `tbl`, and audio/video fills `media`; domain mismatch, ambiguous ownership, or a second fill fails before mutation.

`PptxDocument.masters` and `.layouts` return stable live semantic wrappers. `SlideMasterModel` exposes `partUri`, `layouts`, `theme`, `background`, `shapes`, `placeholders`, and `slideNumber`. `SlideLayoutModel` exposes `partUri`, `name`, `masterPartUri`, runtime `margin`, `background`, `shapes`, `placeholders`, and `slideNumber`. Both wrappers expose `addPlaceholder()`, `addText()`, `addRichText()`, `addShape()`, `addImage()`, `addSvgImage()`, and `addChart()`. Existing content and supported background/placeholder relationships are therefore editable without dropping the raw codec surface.

`replaceSlideMaster(layout, options)` prepares a complete definition and atomically replaces only layout-owned content, background, relationships, slide number, and transient margin. It retains the target part URI, `SlideLayoutModel` identity, master layout ID, and every incoming slide relationship; a changed `master` safely relinks both sides. An equivalent recognized definition is an exact no-op. `deleteSlideMaster(layout, replacement?)` rejects a used layout without a replacement. With an attached same-document replacement, it retargets incoming slides before deleting the layout and collecting only unreferenced owned dependencies. Deleted handles become stale, and later URI reuse does not revive them.

Layout `margin` is intentionally transient because OOXML and PptxGenJS do not serialize this `tableToSlides` hint. It is frozen during the current document session, updates on whole replacement, is removed on delete, and reopens as `undefined`. Background, ordinary content, placeholder identity/geometry, slide numbers, relationships, and payloads are persistent.

The theme API remains available alongside semantic wrappers:

```ts
const theme = document.themes[0];
theme.setColor('accent1', '#2563EB');

const chain = document.masterLayoutTheme.materializeInheritedStyle(
  document.slides[0].partUri,
  document.slides[0].shapes[0].id,
);
```

`masterLayoutTheme` continues to expose raw create/copy/delete/relink operations. Semantic wrapper collections synchronize with those raw operations and retain stable identity while their parts remain attached.

Focused master/layout/placeholder tests report 45 passed / 434 skipped; full Vitest reports 1256 passed / 1 skipped, performance is 1/1 at 578ms, and typecheck/build pass. The installed 57-file tarball has 54 `dist` files. Two clean builds have byte-identical sorted dist-hash manifests and tarballs, with SHA-256 `0a8e958ccde379ae071a7388dc4c29278ac5033a8641976324fcd5820339ad27` and `8362a3af38a4a7e8316a7e49e8cb3f4fb405753bd20cc935db609441819ca5e8`. Packed Node, TypeScript, CLI, and real Chrome all report `masterLayouts: true`; Chrome restores all six placeholder domains, selected layout targets, master/layout backgrounds, payload hashes, and chart definitions with zero validation, console, page, or network errors.

The two-slide native gallery has 32 parts, 29 relationships, two layouts, and one master and validates 0/0 under the PowerPoint 2010 profile. The two-slide PptxGenJS control has 36 parts and 34 relationships. Eight source/LibreOffice-round-trip pages were rendered at 2400×1350 and 180 DPI and inspected individually; full-bleed fixture backgrounds give an expected 0px minimum non-white margin. LibreOffice 26.8 preserves two slides, two layouts, and one master but rewrites placeholder identities and slide-number caches and removes audio plus embedded chart workbooks. Local PowerPoint 16.112 returned `-9074` for both native and control inputs and produced no PPTX/PDF, so neither result is reported as a full round-trip pass.

Full theme text cascade, percentage coordinates, advanced text/table/media/chart styles, and broad client certification remain pending. Advanced text now includes text-shape direct fill, simple line, arrows, simple shadow, outer hyperlink, per-run rich-text hyperlink, and preset geometry creation/read/edit.

`AddTextOptions.rectRadius` is next, followed by `isTextBox` / `breakLine`, the remaining advanced-text items, advanced table/`tableToSlides`, output/runtime helpers, and the peer-range full-suite audit.

## Media

```ts
import { inches } from '@pptx/sdk';

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
  contentType: 'audio/mpeg',
  play: 'auto',
  loop: true,
  hideWhenStopped: true,
  volume: 0.5,
});

declare const videoBytes: Uint8Array;
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
await audio.replaceSource('narration.wav', {
  contentType: 'audio/wav',
  fileName: 'narration.wav',
});
await audio.replacePoster('poster.gif', {
  contentType: 'image/gif',
});
await audio.replacePoster(); // reset to the built-in PNG
```

`MediaSource` is a Node path, strict base64 data URI, `Uint8Array`, `ArrayBuffer`, Blob/File, Web `ReadableStream<MediaByteChunk>`, or async iterable of byte chunks. A string beginning with HTTP/HTTPS is an external media relationship and is never fetched; other URL schemes are rejected. Posters use the same embedded-source union except that HTTP/HTTPS poster URLs are rejected. Omitted posters use a built-in one-pixel PNG.

`AddMediaOptions` contains optional `name`, `altText`, `contentType`, `fileName`, `poster`, `posterContentType`, EMU `x`/`y`/`width`/`height`, `play: 'click' | 'auto'`, `loop`, `hideWhenStopped`, `volume: 0..1`, and an async `transcode(bytes, contentType, kind)` hook. Audio supports `audio/mpeg` (`.mp3`), `audio/mp4` (`.m4a`), `audio/wav` (`.wav`), and `audio/ogg` (`.ogg`). Video supports `video/mp4` (`.mp4`/`.m4v`), `video/quicktime` (`.mov`), and `video/webm` (`.webm`). Posters support `image/png` (`.png`), `image/jpeg` (`.jpg`/`.jpeg`), and `image/gif` (`.gif`).

`addAudio()` and `addVideo()` return a runtime `MediaModel`. Within one document, the result is identical by reference to the corresponding member of `document.media(slideIndex)`, `slide.media`, and `slide.shapes`; collection reads and slide moves preserve that identity. Its live getters are `kind`, `shapeId`/`id`, `slidePartUri`, `name`, `altText`, `settings`, `mediaPartUri`, `externalUrl`, `posterPartUri`, and the inherited `transform`. Set `name`, `altText`, `settings`, or transform directly. A setting assignment synchronizes exact `px:playback` preference/ownership state with native PowerPoint timing. `settings = undefined` removes both the extension and only the library-owned native media graph. Returned settings and transform snapshots are detached and frozen.

`MediaModel.replaceSource(source, options?)` preserves kind and returns the same model. `ReplaceMediaSourceOptions` accepts only `contentType`, `fileName`, and `transcode`; all creation source forms are supported, including embedded↔external transitions. `replacePoster(source?, options?)` accepts only `contentType` and `fileName`; PNG/JPEG/GIF sources are supported, and omitted source resets the built-in PNG. `remove()` delegates to `slide.deleteMedia(shapeId)`. A removed or semantically changed picture invalidates the old live handle through the normal shape-resolution error path.

Descriptor resolution is explicit MIME assertion → data-URI MIME → known extension from `fileName`, path, URL path, or `File.name` → canonical domain default (`audio/mpeg`, `video/mp4`, or `image/png`). Explicit and declared MIME must match. A recognized extension must belong to the correct domain and match the resolved MIME; `.m4v` and `.jpeg` are preserved when explicitly inferred, otherwise the first canonical extension is used. Unknown extensions are ignored as evidence. `Blob.type` is not used. A transcode result must be an ordinary `{ bytes: Uint8Array, contentType, extension? }` object, and any explicit extension must be lowercase and exactly compatible with its output MIME.

Media data URIs require the exact `data:<supported-mime>;base64,<payload>` form. Payloads must use the standard alphabet, complete padding, and canonical padding bits; empty data, whitespace, percent encoding, URL-safe characters, extra commas, and malformed or noncanonical padding are rejected.

Request objects are descriptor-safe, getter-free, and reject unknown properties. Options, byte arrays, ArrayBuffers, transcode inputs/results, and final embedded payloads are detached from the caller. All asynchronous path/Blob/stream I/O, transcode work, poster loading, descriptor resolution, SHA-256 lookup, and XML-definition work completes before mutation. Part, content-type, relationship, and slide XML writes then run inside one synchronous package transaction; any create, replacement, or deletion failure leaves the package graph, ZIP state, model identity, and mutation journal unchanged.

Embedded payloads deduplicate only when SHA-256 and exact MIME both match. Every embedded media picture uses a canonical `a:audioFile` or `a:videoFile`, a standard kind relationship, a Microsoft media relationship, an internal poster image relationship, a media click action, and a rectangular `p:pic`. Duplicated slides initially share media and poster targets. Source/poster replacement updates an exclusive compatible target in place or uses content-aware deduplication/clone-on-write, allocating relationships when an rId is shared by multiple XML nodes and retargeting only the edited picture. Superseded relationships are removed only after their XML reference count reaches zero; media/poster parts are collected only after package-graph incoming reaches zero. Object and slide deletion follow the same reference-aware GC rule.

Playback is native without a plugin. `px:playback` retains exact preferences and the owned media/play/pause timing IDs, while the slide's `p:timing` contains the executable `p:audio`/`p:video`, `p:cMediaNode`, play command, and click pause command where applicable. `play: 'click'` emits an interactive click sequence; `play: 'auto'` emits a with-previous sequence. `loop` maps to indefinite repeat, `hideWhenStopped` maps to `showWhenStopped`, and `volume` maps to OOXML's 0–100000 volume.

Effective read order is a valid `px:playback` preference, then—only when no preference exists—a recognized unique/direct/complete native media graph, then empty settings. A native-only graph can therefore be read and adopted without rewriting its timing bytes when the first assignment is semantically identical. Recorded ownership is repaired when IDs are stale. The codec never claims ordinary animations, nested/non-direct media graphs, multiple matches, finite repeats, cross-slide media, trim/bookmarks, or otherwise unsupported branches.

Create, set, clear, materialize, duplicate, delete, and timing ID allocation are transactional. Timing IDs are unique across media and ordinary animations on the whole slide. Diagnostics are warnings with codes `MEDIA_TIMING_MISSING`, `MEDIA_TIMING_STALE`, `MEDIA_TIMING_UNSUPPORTED`, `MEDIA_TIMING_AMBIGUOUS`, `MEDIA_TIMING_DANGLING_TARGET`, and `MEDIA_TIMING_KIND_MISMATCH`. External media still emits portability diagnostics; under the PowerPoint 2010 profile, `audio/ogg` and `video/webm` emit expected codec warnings.

PptxGenJS 4.0.1 valid public embedded-media cases are semantically covered, including data/path, audio/video, cover, extension, object name, transform, and repeated-path deduplication. Import accepts its audio `a:videoFile`, `audio/mp3`, and duplicate-audio Microsoft-media relationship defects. Reads and non-source edits preserve those legacy primary roles; `replaceSource()` canonicalizes only the selected picture. Native creation always uses `a:audioFile`, canonical `audio/mpeg`, and the standard audio relationship.

The actual 45-file tarball passes Node/real-Chrome/declaration/installed-CLI smoke with `nativeMediaTiming: true`; two clean builds have identical SHA-256 manifests for all 42 dist files. All six presentation formats pass native timing create/edit/duplicate/delete/reopen. The nine-slide playable gallery contains 12 media objects across MP3/M4A/WAV/OGG and MP4/MOV/WebM, all three poster MIME types, ten deduplicated media/poster parts, and zero orphans. It strictly reopens, renders at 180 DPI without overflow, passes slide-by-slide visual inspection, and validates against PowerPoint 2010 with 0 errors plus only the expected OGG/WebM warnings. LibreOffice 26.8 preserves nine slides and text but removes all media, posters, media relationships, and timing on save; its output still strictly reopens and validates 0/0. Local PowerPoint 16.112 automation returned `-9074` for this gallery and both independent control files, so it is not reported as a successful round trip.

Trim/bookmarks, finite repeats, narration/cross-slide audio, captions/subtitles, online video, remote-fetch embedding, media crop/rounding/shadow/hyperlink and advanced placeholder styles, a built-in transcoding engine, and broad client certification remain pending. Native standard-chart creation, semantic editing, direct slide backgrounds, slide numbers, slide default text color, and master/layout/placeholder support are complete; the next PptxGenJS parity slice is advanced text.

## Diagnostics and errors

Errors: `PackageError`, `ParseError`, `ValidationError`, `OpaqueMutationError`, `PptxGenJSAdapterError`.

Every diagnostic has severity, code, message, and optional part URI, XML path, object id, compatibility profile, and suggestion. Strict mode blocks only error diagnostics.
