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
  author: 'Presentation Team',
  rtlMode: true,
  slideSize: { width: inches(11.7), height: inches(8.3) },
  title: 'Quarterly Review',
});
document.author = 'Updated Author';
document.title = 'Updated Review';
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

`CreatePresentationOptions.title` and live `document.title` use the direct core-properties title. Omitted creation input writes no title, `''` writes an explicit empty title, and `undefined` clears only the direct field. Values are strict XML-safe strings; reads follow the package-root core-properties relationship instead of assuming a part URI or prefix, same-value/absent-clear operations are exact no-ops, missing metadata can be created, and unrelated creator/revision/unknown content is preserved. Unsafe malformed or ambiguous ownership is rejected rather than guessed. PptxGenJS 4.0.1 defaults its own public `title` to `PptxGenJS Presentation`; native omitted creation intentionally remains `undefined`.

`CreatePresentationOptions.author` and live `document.author` use only the direct Dublin Core creator. Native omitted creation preserves the canonical `@jiayunxie/pptx`, `''` writes an explicit empty creator, and `undefined` clears only creator. Strict XML-safe values, relationship-based lookup, alternate part URI/prefix support, same-value/absent-clear exact no-ops, missing-part creation, and malformed/ambiguous rejection match the title lifecycle. Author edits preserve `cp:lastModifiedBy`, title, subject, revision, timestamps, unknown children, relationships, and unrelated parts. PptxGenJS 4.0.1 instead defaults author to `PptxGenJS` and mirrors it into creator and lastModifiedBy; native intentionally keeps lastModifiedBy independent.

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
