# @jiayunxie/pptx

Lossless bidirectional PPTX OOXML editing for Node.js and TypeScript.

## Install

This release is a technical preview published under the `next` tag.

```sh
npm install @jiayunxie/pptx@next
```

## Create a presentation

```ts
import { degrees, inches, PRESET_SHAPE_TYPES, PptxDocument } from '@jiayunxie/pptx';

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
});
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

`PRESET_SHAPE_TYPES` is the frozen discovery catalog for all 178 canonical preset geometries accepted by `SlideModel.addShape()`. `AddShapeOptions` accepts `name`, strict `fill`, strict `line`, and native EMU/OOXML-angle transform fields; use `inches()` and `degrees()` for ergonomic conversion. Omitted geometry starts at x/y/width/height = 1 inch with zero rotation and no flips; omitted fill creates direct no-fill, and omitted line keeps the canonical empty line container. Inputs are strict, descriptor-safe, detached before mutation, and reject unknown fields. The catalog uses the valid OOXML `foldedCorner`; PptxGenJS 4.0.1's invalid `folderCorner` token and runtime-only `custGeom` value are not accepted as presets.

`ShapeModel.presetType` reads only one safe direct canonical preset geometry. Reassigning the same type is an exact no-op; changing the type replaces only the geometry and clears old adjustment handles while preserving transform, name, fill, line, effects, text, order, and model identity. Creation, duplicate isolation, rollback, write/reopen, Node/browser bundles, and PptxGenJS public output are covered.

`ShapeFill` supports `{ kind: 'none' }` or a solid six-digit sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. `ShapeModel.fill` returns a detached direct-state snapshot and supports same-value no-op, whole replacement, and clear: none writes direct `a:noFill`, while assigning `undefined` removes the direct fill choice. It does not calculate inherited/effective color. Existing gradient, picture, pattern, and group fills remain lossless during unrelated edits and can be explicitly replaced or cleared, but advanced fill creation remains outside this simple-fill API. PptxGenJS 4.0.1's explicit none and zero transparency omit direct fill/alpha state where native preserves explicit intent; effective rendering is equivalent, not byte-identical. PptxGenJS also turns an empty or missing-color fill into black and accepts deprecated `alpha`, while native rejects both forms before mutation and uses `transparency`.

`ShapeLine` supports `{ kind: 'none' }` or a strict solid line with six-digit sRGB/theme color, optional finite 0–100 transparency, optional 0–1584 point width, and optional `solid | dash | dashDot | lgDash | lgDashDot | lgDashDotDot | sysDash | sysDot` dash. Omitted width/dash materialize as 1pt/solid, while zero width is preserved. `ShapeModel.line` returns a detached direct-state snapshot; same-value assignment is an exact no-op, none writes direct line no-fill, and `undefined` clears only owned width/fill/dash while preserving the line container, arrowheads, joins, extensions, and unrelated attributes. Unique gradient/picture/pattern/group line fill or custom dash can be explicitly replaced or cleared, but their creation is outside this simple-line API. PptxGenJS 4.0.1 maps its `dashType` to native `dash`; its omitted/explicit-none line, empty/missing-color fallback to `333333`, zero-width fallback to 1pt, omitted direct alpha for zero transparency, and ignored deprecated `alpha`/`lineDash` differ from native strict direct-intent semantics. Arrow type/size, cap/compound/alignment/join editing, advanced line fill/custom dash creation, shadow, hyperlink, adjustment editing, custom geometry, shape-text creation options, and percentage positions remain pending.

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
