# @jiayunxie/pptx

Lossless bidirectional PPTX OOXML editing for Node.js and TypeScript.

## Install

This release is a technical preview published under the `next` tag.

```sh
npm install @jiayunxie/pptx@next
```

## Create a presentation

```ts
import { inches, PptxDocument } from '@jiayunxie/pptx';

const document = PptxDocument.create({
  rtlMode: true,
  slideSize: { width: inches(11.7), height: inches(8.3) },
});
const slide = document.addSlide();
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
    valign: 'top',
    margin: { top: 4, left: 8 },
    border: { kind: 'line', color: { kind: 'scheme', value: 'accent2' }, width: 1, style: 'dash' },
    fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
  } }, { text: 'Revenue', options: { fill: { kind: 'none' } } }],
  ['East', { text: '$1.2M' }],
  [{ text: 'West' }, '$980K'],
], { name: 'Revenue table', x: inches(1), y: inches(3), height: inches(2), columnWidths: [inches(2.5), inches(3.5)], rowHeights: [inches(0.5), inches(0.75), inches(0.75)], border: { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1.5, style: 'solid' }, fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent4' }, transparency: 20 }, margin: { top: 9, left: 18 }, valign: 'middle' });
table.setColumnWidths([inches(2), inches(4)]);
table.setRowHeights([inches(0.5), inches(1), inches(0.5)]);
table.setCellVerticalAlignment(0, 0, 'bottom');
table.setCellBorders(0, 0, { bottom: { kind: 'line', color: { kind: 'srgb', value: '70AD47' }, width: 2 } });
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } });
await document.writeFile('created.pptx');
```

`AddTableCellOptions.align` reuses the strict `TextAlignment` values `left`, `center`, `right`, and `justify`, mapped to the created cell paragraph's direct `a:pPr@algn` tokens `l`, `ctr`, `r`, and `just`, respectively. Omitted or runtime-`undefined` values preserve the prior bytes and never synthesize effective left; alignment is not written to `tcPr` or `bodyPr`. The direct token survives supported cell text/property edits, slide duplication, write, and reopen. PptxGenJS 4.0.1 produces the same final state for supported values but silently drops an unknown runtime value, while native creation rejects it with `TypeError`. Table-level `align`, an existing-cell alignment snapshot, and a horizontal-alignment editor remain pending.

`AddTableOptions.margin` supplies a point-based table creation default using a scalar, exact `[top, right, bottom, left]` tuple, or partial named sides. Resolution is per side: canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. A cell scalar/TRBL overrides all four table sides; a partial cell object overrides only the supplied sides; omitted, runtime-undefined, or empty cell margin inherits the table value. Omitted, runtime-undefined, or empty table margin preserves the original canonical bytes. The resolved state is written only to physical-cell `tcPr@marL/marR/marT/marB`, before optional anchor, L/R/T/B borders, and fill. `TableCell.margins` exposes it immediately, and `setCellMargins()` remains a whole-replacement physical-cell editor; clearing a cell does not reapply the creation default because no table metadata is retained. A table-level margin getter/editor and PptxGenJS's legacy dual-unit interpretation remain unsupported; native inputs always use points.

`AddTableOptions.border` accepts the same strict scalar, exact `[top, right, bottom, left]` tuple, or partial named `TableCellBorderInput` as a cell. A non-empty table value is materialized only onto physical cells whose normalized cell border is absent; any non-empty cell scalar, TRBL, named value, or explicit none blocks the entire table value. Missing sides in the chosen value use canonical direct no-fill and never inherit across layers. Empty/all-undefined table input preserves omitted bytes, while empty/all-undefined cell input inherits a supplied table border. Inputs detach deeply and resolved `TableCell.borders` are available immediately. No table border metadata, getter, or editor is retained, so `setCellBorders()` replaces or clears only final direct cell state and never re-inherits. Supported explicit table/cell scalar/TRBL none, solid/dash sRGB lines, scalar zero, and cell overrides match PptxGenJS 4.0.1 final state. Native empty values, omitted style, TRBL zero, strict four-item tuples, named sides, and theme colors retain the documented stricter semantics.

`AddTableOptions.fill` accepts the same strict `TableCellFill` as a cell: `{ kind: 'none' }` or a solid six-digit sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. It is a whole-value creation default materialized into physical cells whose cell fill is omitted or `undefined`; a cell solid or none value completely overrides it. Omitted/runtime-undefined table fill preserves the original bytes, while `{}` is invalid. Inputs detach deeply and resolved `TableCell.fill` values are available immediately. No table fill metadata, getter, or editor is retained, so `setCellFill()` replaces or clears only final direct cell state and never re-inherits the creation default. Serialization remains margins, optional anchor, L/R/T/B borders, then fill. Supported table solid fills match PptxGenJS 4.0.1 final cell state; native direct none and explicit-zero alpha remain distinguishable where PptxGenJS omits them.

Creation is native and does not require PptxGenJS. A new document starts with zero slides and a complete default master/layout/theme chain; its size can use a built-in preset or custom EMU dimensions. `CreatePresentationOptions.rtlMode` and editable `document.rtlMode` control only the direct presentation root; false writes an explicit LTR flag and undefined clears it. `document.slideSize` also reads or changes the canvas of an existing deck without silently scaling its content. Plain-text boxes preserve paragraphs and empty lines; point-based text-box margins accept a scalar, TRBL tuple, or named sides and remain editable through `shape.textMargins`. Text-box `valign` supports top, middle, and bottom through `shape.verticalAlignment`; boolean automatic wrapping remains editable through `shape.textWrap`; all seven OOXML `vert` directions remain editable through `shape.textDirection`; none/shrink/resize fit modes remain editable through `shape.textFit`. `slide.addTable()` creates strict non-empty rectangular matrices of single-paragraph strings and `{ text, options?: { align, border, fill, margin, valign } }` plain objects, which may be mixed, with optional EMU name/x/y/width/height/columnWidths/rowHeights, optional table `border`/`fill`/`valign`, and a live `TableModel`. Cell- and table-level creation `valign` accept only top/middle/bottom and write direct physical-cell `tcPr@anchor="t/ctr/b"`; a cell value overrides the table value. Omitted or runtime-undefined table `valign` preserves existing bytes, while omitted cells inherit a supplied table value. No table default metadata is retained, so clearing a materialized cell later leaves it unset. The returned table exposes each resolved value immediately through `TableCell.verticalAlignment` and `setCellVerticalAlignment()`. Creation margin is point-only and accepts scalar, exact TRBL, or named sides; supplied sides overlay top/bottom 3.6pt and left/right 7.2pt, while omitted/undefined/empty named input preserves the default bytes. Values round to signed-Int32 EMU and support zero, negative, positive, and fractional points. `marL/marR/marT/marB` attributes precede optional `anchor`, followed by L/R/T/B border children and fill. Table- and cell-level creation border accepts scalar, exact TRBL, or named sides using strict none or sRGB/theme lines with `0..1584` point width and omitted/solid/dash style. A non-empty cell border blocks the whole table border; empty/all-undefined cell values inherit, and missing sides in the selected value use four canonical direct no-fill borders. Table- and cell-level creation fill supports direct none or solid sRGB/theme color with 0–100 transparency rounded to 0.001%; cell fill is a whole-value override of table fill. Omitted/undefined table border/fill preserves original bytes, omitted/undefined cell border/fill inherits a supplied table value, and an empty fill is invalid. Omitted or undefined margin and both `valign` layers, empty cell options, empty margin, and an empty border retain the appropriate default bytes. All nested object inputs are descriptor-safe, accept ordinary or null-prototype objects, invoke no getters, and detach immediately. x/y default to 0.5 inch, width to one inch per column, and total height to one inch with automatic zero-height rows. Creation `columnWidths` and `rowHeights` accept a positive scalar repeated over the corresponding axis or an exact-length array; omitted overall dimensions are derived from the explicit vector sum, while supplied dimensions must equal it. Existing tables expose detached strict `TableModel.columnWidths` snapshots and atomic `setColumnWidths()` editing that synchronizes direct grid widths and transform width. They also expose detached strict `TableModel.rowHeights` snapshots where zero means automatic, plus `setRowHeights()` editing that synchronizes transform height only when every target row is positive and otherwise preserves the valid transform height. Malformed grids or rows read as undefined; unsafe required transforms are rejected during mutation. Cell options other than align/border/fill/margin/valign, direction/fit creation, table-level horizontal alignment, existing-cell horizontal-alignment snapshot/editing, table-level border/margin/fill/valign getters or editors, diagonal/advanced borders, advanced fills, rich/multi-paragraph text, hyperlinks, merges, row insertion/deletion, creation styles, auto-page, repeated headers, content measurement, and layout recomputation are still pending; unsupported object properties are rejected, and changing inherited transform dimensions does not resize the grid or rows. Created and existing tables expose strict four-value `TableCell.textDirection` snapshots and physical-cell editing through `TableModel.setCellTextDirection()` on direct `tcPr@vert`; top/middle/bottom `TableCell.verticalAlignment` snapshots and `TableModel.setCellVerticalAlignment()` edit direct `tcPr@anchor`; point-based partial `TableCell.margins` snapshots and whole-replacement `TableModel.setCellMargins()` edit direct `tcPr@marL/marR/marT/marB`; strict four-side `TableCell.borders` snapshots and `TableModel.setCellBorders()` edit direct none/solid/dash borders with point width and sRGB/theme color; strict solid/no-fill `TableCell.fill` snapshots and `TableModel.setCellFill()` edit direct cell fill with sRGB/theme color and percentage transparency. Cell-level horizontal alignment creation matches PptxGenJS 4.0.1 final direct state for left/center/right/justify; PptxGenJS silently drops an unknown runtime value while the strict native API rejects it. Valid table/cell valign creation also matches PptxGenJS 4.0.1 final direct state; PptxGenJS runtime invalid valign passthrough remains unsupported by the strict native API. Margin creation matches PptxGenJS 4.0.1 final direct state when native point values are paired with its legacy `<1` inch / `>=1` point inputs—for example native `7.2` equals PptxGenJS `0.1`; native `0.1` intentionally remains 0.1pt. Border creation is PptxGenJS-compatible for explicit table/cell scalar/TRBL none, solid/dash sRGB lines, scalar zero width, and cell whole-value override; native empty border, omitted style, TRBL zero, exact tuple length, named sides, and theme colors retain stricter direct-state semantics. Supported table solid fills and cell solid overrides match PptxGenJS final direct state; native direct none and explicit zero remain distinguishable while PptxGenJS collapses them to absent direct state. Tables also expose strict `TableCell.textFit` snapshots and `TableModel.setCellTextFit()` editing for direct `bodyPr` none/shrink/resize choices; PptxGenJS 4.0.1 has no matching table fit API. Plain and rich paragraphs support horizontal alignment, point-based direct left/right margins, signed direct first-line/hanging indent, RTL defaults with explicit per-paragraph LTR/RTL overrides, Unicode bullets, all 16 PptxGenJS numbering styles, list levels 0–8, paragraph spacing, and left/center/right/decimal tab stops. `paragraphMarginLeft` / `paragraphMarginRight` / `paragraphIndent` supply creation defaults, and `RichTextParagraph.marginLeft` / `marginRight` / `indent` override or clear them. Positive indent moves the first line inward; negative indent creates a hanging paragraph. Numeric left margin or ordinary indent cannot share one paragraph with an active bullet because the list owns direct `marL` and `indent`; right margin remains independent. `indent: false` suppresses an outer ordinary default so a list can keep its own hanging indent. Direct sides and indent signs do not swap under RTL. Global RTL and paragraph RTL are independent. `addRichText()` and `shape.richText` add structured runs with fonts, sizes, languages, bold/italic, point-based character spacing, text/highlight colors, main-fill transparency from 0 (opaque) through 100 (fully transparent), superscript/subscript/custom baselines, soft breaks, solid sRGB/theme outlines, glow with point radii and opacity, all valid OOXML underline styles with independent colors, and single/double strike. Transparency is rounded to 0.001%, applies to the default `tx1` text color when color is omitted, and is independent from glow opacity and every other fill. `AddTextOptions.lang` and paragraph `rtlMode` supply creation defaults; `RichTextRunStyle.lang` and `RichTextParagraph.rtl` provide run/paragraph overrides.

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
