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

SVG is the next image item. Rounding/transparency, alt-text editing, image hyperlinks/shadows/placeholders, and public image deletion/media garbage collection remain pending.

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

`ShapeFill` supports `{ kind: 'none' }` or a solid six-digit sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. `ShapeModel.fill` returns a detached direct-state snapshot and supports same-value no-op, whole replacement, and clear: none writes direct `a:noFill`, while assigning `undefined` removes the direct fill choice. It does not calculate inherited/effective color. Existing gradient, picture, pattern, and group fills remain lossless during unrelated edits and can be explicitly replaced or cleared, but advanced fill creation remains outside this simple-fill API. PptxGenJS 4.0.1's explicit none and zero transparency omit direct fill/alpha state where native preserves explicit intent; effective rendering is equivalent, not byte-identical. PptxGenJS also turns an empty or missing-color fill into black and accepts deprecated `alpha`, while native rejects both forms before mutation and uses `transparency`.

`ShapeLine` supports `{ kind: 'none' }` or a strict solid line with six-digit sRGB/theme color, optional finite 0–100 transparency, optional 0–1584 point width, and optional `solid | dash | dashDot | lgDash | lgDashDot | lgDashDotDot | sysDash | sysDot` dash. Omitted width/dash materialize as 1pt/solid, while zero width is preserved. `ShapeModel.line` returns a detached direct-state snapshot; same-value assignment is an exact no-op, none writes direct line no-fill, and `undefined` clears only owned width/fill/dash while preserving the line container, arrowheads, joins, extensions, and unrelated attributes. Unique gradient/picture/pattern/group line fill or custom dash can be explicitly replaced or cleared, but their creation is outside this simple-line API. PptxGenJS 4.0.1 maps its `dashType` to native `dash`; its omitted/explicit-none line, empty/missing-color fallback to `333333`, zero-width fallback to 1pt, omitted direct alpha for zero transparency, and ignored deprecated `alpha`/`lineDash` differ from native strict direct-intent semantics.

`ShapeArrowType` is the closed `none | arrow | diamond | oval | stealth | triangle` union. `AddShapeOptions.arrows` and `ShapeModel.arrows` use a detached `ShapeArrows` snapshot with optional `begin` / `end`; assignment is a whole replacement, so an omitted side is cleared, explicit `none` remains a direct endpoint, and `undefined` or `{}` clears both endpoints. Arrow edits own only direct head/tail children: clearing arrows preserves line width/fill/dash, while clearing `ShapeModel.line` preserves arrows. Arrows-only creation writes no implicit line color, width, or dash. Existing legal endpoint `w` / `len` values (`sm | med | lg`) remain lossless during type edits but are not exposed as editable size state. PptxGenJS 4.0.1 instead materializes `333333`/1pt/solid for arrow-only input, ignores empty/nested deprecated aliases, maps top-level deprecated aliases, and can pass invalid runtime tokens through; native rejects aliases and invalid tokens before mutation.

`ShapeShadow` is the strict `kind: 'outer' | 'inner'` union used by `AddShapeOptions.shadow` and `ShapeModel.shadow`. Both kinds accept sRGB/theme color, finite `0..1` opacity, `0..100` point blur, `0 <= angle < 360` degrees, and `0..200` point distance; only outer accepts `rotateWithShape`. Defaults are black, 0.75 opacity, 8pt blur, 270°, 4pt distance, and outer rotate false. Explicit zero survives normalization. Inputs are deeply detached before mutation, getter snapshots are detached and deep-frozen, assignment is a whole replacement, same-value assignment is an exact bytes/journal no-op, and `undefined` removes only the direct inner/outer child while retaining `effectLst` and legal sibling effects. PptxGenJS 4.0.1 omission and `type: 'none'` map to native `undefined`, and its legacy `offset` maps conceptually to native `distance`; native deliberately rejects its zero-value fallback, ignored rotate flag, invalid passthrough, and malformed inner closing tag. Generic/advanced effects, preset shadow, custom shadow transforms, and non-shape shadow APIs remain outside this focused surface.

`Hyperlink` is a mutually exclusive `{ url, tooltip? } | { slide, tooltip? }` union used by `AddShapeOptions.hyperlink` and `ShapeModel.hyperlink`. URLs must be non-empty XML-safe strings; slide targets are one-based positive safe integers that must exist when assigned. Inputs and frozen getter snapshots are detached. Assignment is a whole replacement, so an omitted tooltip removes the direct attribute, an explicit empty tooltip preserves `tooltip=""`, and `undefined` clears the click link. Same-value assignment is an exact no-op. Internal relationships preserve target-slide identity across insert, delete, and reorder; duplicate self-links retarget to the duplicate, target deletion cleans click/hover references, and shared relationships use reference-aware clone-on-write and garbage collection. PptxGenJS 4.0.1 materializes an omitted tooltip as empty and can console-ignore or coerce invalid runtime values into duplicate or dangling links; native rejects those values before mutation. External links intentionally produce the validator's portability warning. Hover editing, text-run/table/image/chart/media hyperlink creation, action navigation, arrow size, cap/compound/alignment/join editing, advanced line fill/custom dash creation, shape-text creation options, and percentage positions remain pending.

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

`AddTableOptions.align` and `AddTableCellOptions.align` reuse the strict `TextAlignment` values `left`, `center`, `right`, and `justify`, mapped to each created cell paragraph's direct `a:pPr@algn` tokens `l`, `ctr`, `r`, and `just`, respectively. A table value is materialized onto every single-paragraph cell that omits cell alignment or supplies runtime `undefined`; a valid cell value wins. Omitted or runtime-`undefined` table alignment preserves current bytes and never synthesizes effective left. Final ownership is direct `a:pPr@algn`, not `tcPr`, `bodyPr`, or retained table metadata, so clearing a cell later does not reapply the creation default. `TableCell.horizontalAlignment` reads only one strict direct paragraph as the four public values; missing, malformed, ambiguous, or multi-paragraph state returns `undefined`. `TableModel.setCellHorizontalAlignment()` uses physical zero-based row/cell coordinates, writes `l`/`ctr`/`r`/`just`, and clears only the unqualified direct token with `undefined`. Same-value and absent-clear calls are exact no-ops; unsafe structure rejects without mutation, while unrelated paragraph/cell state remains preserved. PptxGenJS 4.0.1 materializes supported table/cell values and precedence into the same importable direct state; native existing-deck editing is a lossless extension because PptxGenJS has no existing-deck editor. PptxGenJS still silently drops an unknown table runtime value while native creation/editing rejects invalid values before mutation. Rich or multi-paragraph cell alignment remains pending.

`AddTableOptions.textDirection` and `AddTableCellOptions.textDirection` accept the strict table-cell values `horz`, `vert`, `vert270`, and `wordArtVert`. A table value is materialized onto every physical cell whose cell value is omitted or runtime `undefined`; any explicit cell value wins. Explicit cell `horz` blocks a non-horizontal table value while still writing no direct token. Creation owns direct physical-cell `tcPr@vert`: omitted, runtime-`undefined`, and resolved horizontal values write no attribute, while the three non-horizontal values write their exact token and appear immediately in `TableCell.textDirection`. Omitted/runtime-undefined table direction preserves existing bytes, and explicit table `horz` produces the same direction bytes for uncovered cells. Values are normalized descriptor-safely and detached immediately. Only final cell state is retained, so clearing a cell later does not reapply the table value. `setCellTextDirection()` deliberately retains direct-editor semantics, so assigning `horz` writes `vert="horz"` and assigning `undefined` clears the attribute. Supported table/cell creation and precedence match PptxGenJS 4.0.1 final state; native rejects its invalid runtime-token passthrough. A table-level direction getter/editor remains pending.

`AddTableOptions.margin` supplies a point-based table creation default using a scalar, exact `[top, right, bottom, left]` tuple, or partial named sides. Resolution is per side: canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. A cell scalar/TRBL overrides all four table sides; a partial cell object overrides only the supplied sides; omitted, runtime-undefined, or empty cell margin inherits the table value. Omitted, runtime-undefined, or empty table margin preserves the original canonical bytes. The resolved state is written only to physical-cell `tcPr@marL/marR/marT/marB`, before optional anchor, L/R/T/B borders, and fill. `TableCell.margins` exposes it immediately, and `setCellMargins()` remains a whole-replacement physical-cell editor; clearing a cell does not reapply the creation default because no table metadata is retained. A table-level margin getter/editor and PptxGenJS's legacy dual-unit interpretation remain unsupported; native inputs always use points.

`AddTableOptions.border` accepts the same strict scalar, exact `[top, right, bottom, left]` tuple, or partial named `TableCellBorderInput` as a cell. A non-empty table value is materialized only onto physical cells whose normalized cell border is absent; any non-empty cell scalar, TRBL, named value, or explicit none blocks the entire table value. Missing sides in the chosen value use canonical direct no-fill and never inherit across layers. Empty/all-undefined table input preserves omitted bytes, while empty/all-undefined cell input inherits a supplied table border. Inputs detach deeply and resolved `TableCell.borders` are available immediately. No table border metadata, getter, or editor is retained, so `setCellBorders()` replaces or clears only final direct cell state and never re-inherits. Supported explicit table/cell scalar/TRBL none, solid/dash sRGB lines, scalar zero, and cell overrides match PptxGenJS 4.0.1 final state. Native empty values, omitted style, TRBL zero, strict four-item tuples, named sides, and theme colors retain the documented stricter semantics.

`AddTableOptions.fill` accepts the same strict `TableCellFill` as a cell: `{ kind: 'none' }` or a solid six-digit sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. It is a whole-value creation default materialized into physical cells whose cell fill is omitted or `undefined`; a cell solid or none value completely overrides it. Omitted/runtime-undefined table fill preserves the original bytes, while `{}` is invalid. Inputs detach deeply and resolved `TableCell.fill` values are available immediately. No table fill metadata, getter, or editor is retained, so `setCellFill()` replaces or clears only final direct cell state and never re-inherits the creation default. Serialization remains margins, optional anchor, L/R/T/B borders, then fill. Supported table solid fills match PptxGenJS 4.0.1 final cell state; native direct none and explicit-zero alpha remain distinguishable where PptxGenJS omits them.

Cell-level table text-fit creation uses only `AddTableCellOptions.fit` with strict `none`, `shrink`, and `resize` values. Omitted, runtime `undefined`, and `none` produce byte-identical self-closing `bodyPr` and an immediate direct `TableCell.textFit` snapshot of `undefined`; `shrink` and `resize` write direct `normAutofit` and `spAutoFit` and snapshot immediately. Creation never writes `noAutofit`, retains table metadata, measures content, calculates final font scale, or recomputes cell/table geometry. `setCellTextFit()` keeps direct-editor semantics: `none` and `undefined` clear without restoring creation input, and fit remains independent from `tcPr@vert` direction. Table-level fit creation/default/getter/editor remains unsupported. PptxGenJS 4.0.1 ignores table/cell runtime `fit`, `autoFit`, and `shrinkText`, so native shrink/resize creation is an explicit extension rather than a parity claim.

Creation is native and does not require PptxGenJS. A new document starts with zero slides and a complete default master/layout/theme chain; its size can use a built-in preset or custom EMU dimensions. `CreatePresentationOptions.rtlMode` and editable `document.rtlMode` control only the direct presentation root; false writes an explicit LTR flag and undefined clears it. `document.slideSize` also reads or changes the canvas of an existing deck without silently scaling its content. Plain-text boxes preserve paragraphs and empty lines; point-based text-box margins accept a scalar, TRBL tuple, or named sides and remain editable through `shape.textMargins`. Text-box `valign` supports top, middle, and bottom through `shape.verticalAlignment`; boolean automatic wrapping remains editable through `shape.textWrap`; all seven OOXML `vert` directions remain editable through `shape.textDirection`; none/shrink/resize fit modes remain editable through `shape.textFit`. `slide.addTable()` creates strict non-empty rectangular matrices of single-paragraph strings and `{ text, options?: { align, border, fill, fit, margin, textDirection, valign } }` plain objects, which may be mixed, with optional EMU name/x/y/width/height/columnWidths/rowHeights, optional table `align`/`border`/`fill`/`margin`/`textDirection`/`valign`, and a live `TableModel`. Table- and cell-level `textDirection` accept the four strict table-cell tokens and write only final direct physical-cell `tcPr@vert`. A table value materializes onto cells with omitted/undefined cell direction; any cell value wins, including explicit `horz`, which blocks a non-horizontal table value while still writing no token. Resolved omission/undefined/`horz` collapse to no attribute, while `vert`/`vert270`/`wordArtVert` remain visible immediately and after reopen. Cell- and table-level creation `valign` accept only top/middle/bottom and write direct physical-cell `tcPr@anchor="t/ctr/b"`; a cell value overrides the table value. Omitted or runtime-undefined table `valign` preserves existing bytes, while omitted cells inherit a supplied table value. No table creation-default metadata is retained, so clearing any materialized cell property later leaves it unset. The returned table exposes each resolved direction and horizontal/vertical alignment immediately through `TableCell.textDirection`, `TableCell.horizontalAlignment`, and `TableCell.verticalAlignment` and their physical-cell setters. Creation margin is point-only and accepts scalar, exact TRBL, or named sides; supplied sides overlay top/bottom 3.6pt and left/right 7.2pt, while omitted/undefined/empty named input preserves the default bytes. Values round to signed-Int32 EMU and support zero, negative, positive, and fractional points. `marL/marR/marT/marB` attributes precede optional `anchor` and `vert`, followed by L/R/T/B border children and fill. Table- and cell-level creation border accepts scalar, exact TRBL, or named sides using strict none or sRGB/theme lines with `0..1584` point width and omitted/solid/dash style. A non-empty cell border blocks the whole table border; empty/all-undefined cell values inherit, and missing sides in the selected value use four canonical direct no-fill borders. Table- and cell-level creation fill supports direct none or solid sRGB/theme color with 0–100 transparency rounded to 0.001%; cell fill is a whole-value override of table fill. Omitted/undefined table border/fill preserves original bytes, omitted/undefined cell border/fill inherits a supplied table value, and an empty fill is invalid. Omitted or undefined margin and both `valign` layers, empty cell options, empty margin, and an empty border retain the appropriate default bytes. All nested object inputs are descriptor-safe, accept ordinary or null-prototype objects, invoke no getters, and detach immediately. x/y default to 0.5 inch, width to one inch per column, and total height to one inch with automatic zero-height rows. Creation `columnWidths` and `rowHeights` accept a positive scalar repeated over the corresponding axis or an exact-length array; omitted overall dimensions are derived from the explicit vector sum, while supplied dimensions must equal it. Existing tables expose detached strict `TableModel.columnWidths` snapshots and atomic `setColumnWidths()` editing that synchronizes direct grid widths and transform width. They also expose detached strict `TableModel.rowHeights` snapshots where zero means automatic, plus `setRowHeights()` editing that synchronizes transform height only when every target row is positive and otherwise preserves the valid transform height. Malformed grids or rows read as undefined; unsafe required transforms are rejected during mutation. Cell options other than align/border/fill/fit/margin/textDirection/valign, table-level fit creation/default/getter/editor, table-level direction/border/margin/fill/valign getters or editors, diagonal/advanced borders, advanced fills, rich/multi-paragraph text and alignment, hyperlinks, merges, row insertion/deletion, creation styles, auto-page, repeated headers, content measurement, and layout recomputation are still pending; unsupported object properties are rejected, and changing inherited transform dimensions does not resize the grid or rows. Created and existing tables expose strict four-value `TableCell.textDirection` snapshots and physical-cell editing through `TableModel.setCellTextDirection()` on direct `tcPr@vert`; strict four-value `TableCell.horizontalAlignment` snapshots and `TableModel.setCellHorizontalAlignment()` edit direct single-paragraph `pPr@algn`; top/middle/bottom `TableCell.verticalAlignment` snapshots and `TableModel.setCellVerticalAlignment()` edit direct `tcPr@anchor`; point-based partial `TableCell.margins` snapshots and whole-replacement `TableModel.setCellMargins()` edit direct `tcPr@marL/marR/marT/marB`; strict four-side `TableCell.borders` snapshots and `TableModel.setCellBorders()` edit direct none/solid/dash borders with point width and sRGB/theme color; strict solid/no-fill `TableCell.fill` snapshots and `TableModel.setCellFill()` edit direct cell fill with sRGB/theme color and percentage transparency. Table- and cell-level horizontal alignment creation and imported materialized state match PptxGenJS 4.0.1 for left/center/right/justify with cell precedence; native existing-deck editing is a lossless extension, while PptxGenJS silently drops an unknown table runtime value and the strict native API rejects it before mutation. Valid table/cell direction creation and precedence match PptxGenJS final direct state; its runtime invalid direction passthrough remains unsupported by the strict native API. Valid table/cell valign creation also matches PptxGenJS 4.0.1 final direct state; PptxGenJS runtime invalid valign passthrough remains unsupported by the strict native API. Margin creation matches PptxGenJS 4.0.1 final direct state when native point values are paired with its legacy `<1` inch / `>=1` point inputs—for example native `7.2` equals PptxGenJS `0.1`; native `0.1` intentionally remains 0.1pt. Border creation is PptxGenJS-compatible for explicit table/cell scalar/TRBL none, solid/dash sRGB lines, scalar zero width, and cell whole-value override; native empty border, omitted style, TRBL zero, exact tuple length, named sides, and theme colors retain stricter direct-state semantics. Supported table solid fills and cell solid overrides match PptxGenJS final direct state; native direct none and explicit zero remain distinguishable while PptxGenJS collapses them to absent direct state. Tables also expose strict `TableCell.textFit` snapshots and `TableModel.setCellTextFit()` editing for direct `bodyPr` none/shrink/resize choices; PptxGenJS 4.0.1 has no matching table fit API. Plain and rich paragraphs support horizontal alignment, point-based direct left/right margins, signed direct first-line/hanging indent, RTL defaults with explicit per-paragraph LTR/RTL overrides, Unicode bullets, all 16 PptxGenJS numbering styles, list levels 0–8, paragraph spacing, and left/center/right/decimal tab stops. `paragraphMarginLeft` / `paragraphMarginRight` / `paragraphIndent` supply creation defaults, and `RichTextParagraph.marginLeft` / `marginRight` / `indent` override or clear them. Positive indent moves the first line inward; negative indent creates a hanging paragraph. Numeric left margin or ordinary indent cannot share one paragraph with an active bullet because the list owns direct `marL` and `indent`; right margin remains independent. `indent: false` suppresses an outer ordinary default so a list can keep its own hanging indent. Direct sides and indent signs do not swap under RTL. Global RTL and paragraph RTL are independent. `addRichText()` and `shape.richText` add structured runs with fonts, sizes, languages, bold/italic, point-based character spacing, text/highlight colors, main-fill transparency from 0 (opaque) through 100 (fully transparent), superscript/subscript/custom baselines, soft breaks, solid sRGB/theme outlines, glow with point radii and opacity, all valid OOXML underline styles with independent colors, and single/double strike. Transparency is rounded to 0.001%, applies to the default `tx1` text color when color is omitted, and is independent from glow opacity and every other fill. `AddTextOptions.lang` and paragraph `rtlMode` supply creation defaults; `RichTextRunStyle.lang` and `RichTextParagraph.rtl` provide run/paragraph overrides.

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
