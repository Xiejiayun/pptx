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
  ['Region', 'Revenue'],
  ['East', '$1.2M'],
  ['West', '$980K'],
], { name: 'Revenue table', x: inches(1), y: inches(3), height: inches(2), columnWidths: [inches(2.5), inches(3.5)], rowHeights: [inches(0.5), inches(0.75), inches(0.75)] });
table.setCellFill(0, 0, { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } });
await document.writeFile('created.pptx');
```

Creation is native and does not require PptxGenJS. A new document starts with zero slides and a complete default master/layout/theme chain; its size can use a built-in preset or custom EMU dimensions. `CreatePresentationOptions.rtlMode` and editable `document.rtlMode` control only the direct presentation root; false writes an explicit LTR flag and undefined clears it. `document.slideSize` also reads or changes the canvas of an existing deck without silently scaling its content. Plain-text boxes preserve paragraphs and empty lines; point-based text-box margins accept a scalar, TRBL tuple, or named sides and remain editable through `shape.textMargins`. Text-box `valign` supports top, middle, and bottom through `shape.verticalAlignment`; boolean automatic wrapping remains editable through `shape.textWrap`; all seven OOXML `vert` directions remain editable through `shape.textDirection`; none/shrink/resize fit modes remain editable through `shape.textFit`. `slide.addTable()` creates strict non-empty rectangular single-paragraph string matrices with optional EMU name/x/y/width/height/columnWidths/rowHeights and returns a live `TableModel`; x/y default to 0.5 inch, width to one inch per column, and total height to one inch with automatic zero-height rows. `columnWidths` and `rowHeights` accept a positive scalar repeated over the corresponding axis or an exact-length array; omitted overall dimensions are derived from the explicit vector sum, while supplied dimensions must equal it. Existing-table column-width/row-height reading and editing, cell objects, rich text, merges, creation styles, and auto-page are still pending; changing inherited transform dimensions does not resize the grid or rows. Created and existing tables expose strict four-value `TableCell.textDirection` snapshots and physical-cell editing through `TableModel.setCellTextDirection()` on direct `tcPr@vert`; top/middle/bottom `TableCell.verticalAlignment` snapshots and `TableModel.setCellVerticalAlignment()` edit direct `tcPr@anchor`; point-based partial `TableCell.margins` snapshots and `TableModel.setCellMargins()` edit direct `tcPr@marL/marR/marT/marB`; strict four-side `TableCell.borders` snapshots and `TableModel.setCellBorders()` edit direct none/solid/dash borders with point width and sRGB/theme color; strict solid/no-fill `TableCell.fill` snapshots and `TableModel.setCellFill()` edit direct cell fill with sRGB/theme color and percentage transparency. They also expose strict `TableCell.textFit` snapshots and `TableModel.setCellTextFit()` editing for direct `bodyPr` none/shrink/resize choices; PptxGenJS 4.0.1 has no matching table fit API. Plain and rich paragraphs support horizontal alignment, point-based direct left/right margins, signed direct first-line/hanging indent, RTL defaults with explicit per-paragraph LTR/RTL overrides, Unicode bullets, all 16 PptxGenJS numbering styles, list levels 0–8, paragraph spacing, and left/center/right/decimal tab stops. `paragraphMarginLeft` / `paragraphMarginRight` / `paragraphIndent` supply creation defaults, and `RichTextParagraph.marginLeft` / `marginRight` / `indent` override or clear them. Positive indent moves the first line inward; negative indent creates a hanging paragraph. Numeric left margin or ordinary indent cannot share one paragraph with an active bullet because the list owns direct `marL` and `indent`; right margin remains independent. `indent: false` suppresses an outer ordinary default so a list can keep its own hanging indent. Direct sides and indent signs do not swap under RTL. Global RTL and paragraph RTL are independent. `addRichText()` and `shape.richText` add structured runs with fonts, sizes, languages, bold/italic, point-based character spacing, text/highlight colors, main-fill transparency from 0 (opaque) through 100 (fully transparent), superscript/subscript/custom baselines, soft breaks, solid sRGB/theme outlines, glow with point radii and opacity, all valid OOXML underline styles with independent colors, and single/double strike. Transparency is rounded to 0.001%, applies to the default `tx1` color when color is omitted, and is independent from glow opacity and every other fill. `AddTextOptions.lang` and paragraph `rtlMode` supply creation defaults; `RichTextRunStyle.lang` and `RichTextParagraph.rtl` provide run/paragraph overrides.

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
