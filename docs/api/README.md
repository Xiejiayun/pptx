# Public API 0.1

## Create, open, and save

```ts
import { inches, PptxDocument } from '@pptx/sdk';

const created = PptxDocument.create({
  format: 'pptx',
  rtlMode: true,
  slideSize: '16:9',
});
created.addSlide();
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

Inputs: `Uint8Array`, `ArrayBuffer`, `Blob`/`File`, Web `ReadableStream`, or async byte iterable. Node.js additionally accepts a file path or Node readable stream. `write()` returns `Uint8Array`; browsers can use `writeBlob()` or `download()`.

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

Shape kinds include `text`, `shape`, `image`, `table`, `chart`, `graphic-frame`, and `group`. Images expose embedded part URIs and replacement; tables support basic native creation plus rows/cells, cell text, borders, fill, margins, text-direction, text-fit, and vertical-alignment editing; charts expose cached series and lossless chart XML editing.

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
import { emuToInches, inches, type AddTableOptions } from '@pptx/sdk';

const tableOptions: AddTableOptions = {
  name: 'Revenue table',
  x: inches(1),
  y: inches(1.25),
  width: inches(8),
  height: inches(2.25),
  columnWidths: [inches(2.5), inches(3.5), inches(2)],
  rowHeights: [inches(0.5), inches(0.75), inches(1)],
};
const createdTable = document.addSlide().addTable([
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', '12%'],
  ['West', '$980K', ''],
], tableOptions);
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

`addTable()` currently accepts a non-empty, dense, rectangular matrix of single-paragraph strings. Geometry uses EMU: x/y default to 0.5 inch, width defaults to one inch per column, and height defaults to one inch with automatic rows (`a:tr@h=0`); use `inches()` for inch-based inputs. `columnWidths` and `rowHeights` each accept one positive EMU value repeated across the corresponding axis or a dense array whose length exactly matches the column/row count. Values are rounded to safe integers, input arrays are detached, and overall width/height are derived from the vector sum when omitted; when both an overall dimension and its vector are supplied, their rounded values must match exactly. Without a vector, explicit width/height retain their existing exact distribution behavior. Emitted explicit transform dimensions always equal the grid/row sum, while omitted `rowHeights` preserve automatic zero-height rows. New cells materialize 7.2pt horizontal and 3.6pt vertical margins plus four direct no-fill borders, return a live `TableModel`, and can immediately use all existing cell and transform editors.

`TableModel.columnWidths` reads the unique direct `tblGrid` as a detached exact-EMU snapshot. A malformed or ambiguous grid returns `undefined` instead of guessing from the transform or cells. `setColumnWidths()` accepts a positive scalar broadcast or a dense descriptor-safe exact-length array, rounds each item to a safe EMU integer, rejects unsafe sums, and atomically updates both `gridCol@w` and `ext@cx`. A valid grid/transform mismatch is repaired; a numeric no-op preserves the original slide bytes and mutation journal. Unsafe existing grid or transform XML raises `ModelParseError` without mutation. Inherited `setTransform({ width })` still changes only the transform, so use `setColumnWidths()` when changing table width distribution.

`TableModel.rowHeights` reads all direct `tr@h` values as a detached exact-EMU snapshot; zero is a valid automatic row height. A malformed or ambiguous direct row vector returns `undefined` without guessing from transform height or cell content. `setRowHeights()` accepts a non-negative scalar broadcast or a dense descriptor-safe exact-length array, rejects raw negative values, and rounds each item to a safe EMU integer. When every target is positive, their safe exact sum is written to `ext@cy` and a valid rows/transform mismatch is repaired. When any target is zero, row tokens are updated but the already-valid transform height is preserved because the rendered automatic height cannot be derived from the numeric row sum. Numeric no-ops preserve original tokens, slide bytes, and the mutation journal; unsafe existing rows or transform XML raise `ModelParseError` without mutation. Creation remains stricter: explicit `addTable({ rowHeights })` values must be positive, while omitting `rowHeights` creates automatic rows. Inherited `setTransform({ height })` still changes only the transform, so use `setRowHeights()` to edit row tokens. Cell objects, rich/multi-paragraph text, colspan/rowspan or merge editing, row insertion/deletion, table/cell creation styles, auto-page/repeated headers, hyperlinks, and content measurement are not yet supported by native creation.

```ts
import {
  TableModel,
  type TableCellBorders,
  type TableCellFill,
  type TableCellTextDirection,
  type TextBoxFit,
  type TextBoxMargins,
  type TextBoxVerticalAlignment,
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

Table-cell `textDirection` is an immutable snapshot with the dedicated values `horz`, `vert`, `vert270`, and `wordArtVert`. `setCellTextDirection()` addresses physical zero-based row/cell positions and edits only the selected cell's direct `tcPr@vert`; explicit `horz` writes `vert="horz"`, while `undefined` clears the direct attribute. The getter requires one direct `tcPr`, one unqualified `vert`, and one exact public token; it does not resolve table defaults or inheritance. PptxGenJS 4.0.1 materializes a table-level direction onto individual cells, so imported cells expose those wire attributes; PptxGenJS collapses omitted and explicit `horz` to no attribute, and both therefore import as `undefined`. This four-value cell API is intentionally separate from the seven-value text-box API and from unsupported table-level direction defaults.

Table-cell `verticalAlignment` reuses `TextBoxVerticalAlignment` values `top`, `middle`, and `bottom`, but stores them only in the selected physical cell's direct `tcPr@anchor` as `t`, `ctr`, or `b`. `setCellVerticalAlignment()` writes the canonical token or clears it with `undefined`; the strict snapshot getter requires exactly one direct `tcPr` and one unqualified supported anchor. It does not read or change the separate `bodyPr@anchor`, resolve effective defaults, or support `just` / `dist`. PptxGenJS 4.0.1 leaves the direct anchor absent when table/cell `valign` is omitted, copies an explicit table-level value into uncovered cells, and lets a cell value override it; runtime invalid tokens remain opaque and import as `undefined`.

Table-cell `margins` reuses the point-based `TextBoxMargins` value shape but owns only the selected physical cell's direct `tcPr@marL/marR/marT/marB`; it never reads or changes text-box `bodyPr@*Ins`. `setCellMargins()` addresses zero-based physical row/cell positions and accepts a point scalar, `[top, right, bottom, left]` tuple, partial named object, `{}` or `undefined`; the named/object form is a whole replacement, so omitted sides are cleared. The snapshot getter requires exactly one direct `tcPr`, reads each unique unqualified signed-Int32 integer independently, and returns only valid direct sides. PptxGenJS 4.0.1 writes narrow direct defaults when margin is omitted, materializes a table-level margin into cells, and retains a legacy branch that treats a first value below 1 as inches but a first value of at least 1 as points. Adapter imports expose the resulting OOXML in points instead of guessing the original unit.

Table-cell `borders` is a detached partial snapshot of the selected physical cell's same-prefix direct `lnL/lnR/lnT/lnB`. Each supported side is `{ kind: 'none' }` or `{ kind: 'line', color, width, style? }`, where width is finite `0..1584` points, color is strict sRGB/theme, and style is omitted, `solid`, or `dash`. `setCellBorders()` accepts one scalar border for all sides, an exact `[top, right, bottom, left]` tuple, or a partial named whole replacement; omitted named/tuple sides are cleared, and `{}` / `undefined` clear all four direct sides. Explicit none writes a direct zero-width `noFill` line, while an omitted side removes that line, so table-style fallback remains distinct. Public TRBL input is serialized in OOXML's required L/R/T/B order; dash maps to `sysDash`, zero-width line remains a line, and widths are quantized to the nearest EMU. The strict getter omits malformed or unsupported sides independently and never treats diagonals, cell fill, text outlines, advanced dash presets, joins, arrows, or shared-edge/effective style as supported state; unrelated edits preserve them. PptxGenJS 4.0.1 materializes four direct noFill sides when border is omitted, copies table-level scalar/TRBL values into cells, defaults visible borders to `666666`/1pt/solid, preserves cell-level `pt: 0`, and changes table-tuple `pt: 0` to its 1pt default. Adapter snapshots reflect final XML rather than guessing the original input layer.

Table-cell `fill` is a detached direct-state snapshot with explicit `{ kind: 'none' }` and `{ kind: 'solid', color, transparency? }` variants. Solid fill supports strict six-digit sRGB or the existing theme-color tokens; transparency is a finite `0..100` percentage quantized to `0.001%`. Omitted transparency writes no alpha, while explicit zero writes direct opaque alpha. `setCellFill()` addresses zero-based physical row/cell positions: `none` writes direct `tcPr/a:noFill`, while `undefined` removes the direct fill choice so table-style fallback remains distinct. The strict getter requires a unique direct `tcPr` and one unambiguous same-prefix noFill or solidFill; it never reads border/text descendant fills or resolves table styles. Unsupported gradient/pattern/picture/group fills remain preserved during unrelated edits and can be explicitly replaced or cleared. PptxGenJS 4.0.1 materializes table-level fill into cells, collapses omitted and `type: 'none'` to no direct fill, and may emit invalid alpha for out-of-range runtime values; adapter imports preserve that XML but do not fabricate a valid snapshot.

Table-cell `textFit` reuses the immutable `TextBoxFit` values `none`, `shrink`, and `resize`, backed only by the selected physical cell's direct `txBody/bodyPr` fit choice. The getter requires exactly one direct text body, one direct body properties element, and one unambiguous supported fit child; it reads an existing `noAutofit` as `none` and otherwise returns `undefined` for absent or malformed state. `setCellTextFit()` uses physical zero-based row/cell positions: `shrink` and `resize` write `normAutofit` and `spAutoFit`, while `none` and `undefined` both clear the direct choice without creating `noAutofit`. Reassigning the current shrink/resize mode preserves any calculated `fontScale` and `lnSpcReduction`. The operation does not change `tcPr@vert`, compute final font scaling, resize a table, or add table-level defaults. PptxGenJS 4.0.1 has no table-cell fit API and ignores runtime `fit`, `autoFit`, and `shrinkText` values supplied to table/cell options.

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

## Gradients and transparency

```ts
document.slides[0].background = {
  kind: 'linear-gradient',
  angle: 45,
  stops: [
    { offset: 0, color: '#2563EB' },
    { offset: 1, color: '#7C3AED', alpha: 0.65 },
  ],
};
```

Colors can use sRGB, scRGB, scheme, system, or preset sources. OOXML transforms retain their original order.

## Master, layout, and theme

```ts
const theme = document.themes[0];
theme.setColor('accent1', '#2563EB');

const chain = document.masterLayoutTheme.materializeInheritedStyle(
  document.slides[0].partUri,
  document.slides[0].shapes[0].id,
);
```

`masterLayoutTheme` also exposes create/copy/delete/relink operations for masters, layouts, and themes.

## Media

```ts
await document.addAudio(0, audioBuffer, {
  contentType: 'audio/mpeg',
  play: 'click',
  volume: 0.8,
});

await document.addVideo(0, 'https://example.com/video.mp4');
```

Local media can come from paths, bytes, ArrayBuffers, or streams. External URLs are never fetched automatically and produce portability diagnostics.

## Diagnostics and errors

Errors: `PackageError`, `ParseError`, `ValidationError`, `OpaqueMutationError`, `PptxGenJSAdapterError`.

Every diagnostic has severity, code, message, and optional part URI, XML path, object id, compatibility profile, and suggestion. Strict mode blocks only error diagnostics.
