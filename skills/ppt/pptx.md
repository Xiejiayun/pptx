# Native PPTX creation and editing

This reference explains how to create and edit presentations with `@jiayunxie/pptx`. It adapts the narrative, visual-design, and render-review discipline of Anthropic's presentation guidance to this repository's native semantic APIs.

## Contents

1. [Default source policy](#default-source-policy)
2. [ThemeSpec](#themespec)
3. [Typography](#typography)
4. [Spacing and layout families](#spacing-and-layout-families)
5. [Install and import](#install-and-import)
6. [Units, layouts, and formats](#units-layouts-and-formats)
7. [Create from zero](#create-from-zero)
8. [Open and edit](#open-and-edit)
9. [Text and rich text](#text-and-rich-text)
10. [Shapes and styling](#shapes-and-styling)
11. [Images, SVG, and backgrounds](#images-svg-and-backgrounds)
12. [Tables and pagination](#tables-and-pagination)
13. [Native charts](#native-charts)
14. [Audio and video](#audio-and-video)
15. [Masters, layouts, and placeholders](#masters-layouts-and-placeholders)
16. [Metadata, sections, notes, and slide numbers](#metadata-sections-notes-and-slide-numbers)
17. [Output and runtime environments](#output-and-runtime-environments)
18. [Inspection, validation, and diff](#inspection-validation-and-diff)
19. [Narrative and visual design](#narrative-and-visual-design)
20. [180-second execution and QA](#180-second-execution-and-qa)
21. [Deliberate boundaries](#deliberate-boundaries)
22. [Common mistakes](#common-mistakes)

## Default source policy

The default create path is offline and deterministic. Use sources in this order:

1. user-provided files;
2. assets already present in the repository or selected template;
3. native editable PowerPoint text, shapes, tables, charts, and diagrams;
4. simple local SVG or raster assets derived from already available inputs.

Do not browse, search, scrape, or download stock imagery or other assets during ordinary deck generation. Do not call image search or image generation by default. External acquisition is a separate, explicitly requested workflow and is outside the 180-second budget. When no photographic asset exists, use a designed native composition instead of a placeholder image, broken URL, or last-minute search.

## ThemeSpec

Before creating slides, derive one compact `ThemeSpec` from the subject. This is a generator-side design contract, not a fabricated package API.

```js
const themeSpec = {
  mode: 'sandwich', // light, dark, or sandwich
  primary: '123D2A',
  secondary: '3D6B45',
  accent: 'E0A83E',
  background: 'F3F0E4',
  surface: 'DDE8D3',
  text: '173126',
  mutedText: '607267',
  motif: 'layered canopy arcs',
  displayFont: 'Cambria',
  bodyFont: 'Arial',
  sizes: { cover: 56, title: 40, heading: 22, body: 16, caption: 11 },
  margin: 0.6,
  gap: 0.3,
};
```

Use the primary/background family for roughly 60–70% of the visual field, secondary for structure, accent sparingly for emphasis, surface for panels, text for primary copy, and mutedText for captions. Adapt hues to the query instead of defaulting to generic blue. Choose one content-specific motif and repeat it with restraint. When no treatment is implied, use sandwich mode: dark cover and conclusion with lighter content slides.

Use these starting palettes as adaptable color logic, never as fixed templates:

| Starting palette | primary | secondary | accent | background | surface | text | mutedText |
|---|---|---|---|---|---|---|---|
| Forest canopy | `123D2A` | `3D6B45` | `E0A83E` | `F3F0E4` | `DDE8D3` | `173126` | `607267` |
| Deep ocean | `103C5A` | `2E7185` | `F0A35B` | `F2F6F7` | `D8E8EA` | `17313D` | `5D717A` |
| Editorial clay | `713C2F` | `B2694F` | `E3B94F` | `F6EFE5` | `EAD8C8` | `342720` | `74645A` |
| Night sky | `171B35` | `414A78` | `E7C568` | `F2F0EA` | `DADDEA` | `20243A` | `696E84` |

Change hue, value, and contrast to fit the subject, audience, and supplied brand. Keep the semantic roles even when every hexadecimal value changes.

Map the contract to real public fields:

```js
const document = PptxDocument.create({ slideSize: 'wide', title: 'Amazon biodiversity' });
const slide = document.addSlide();
slide.background = { kind: 'solid', color: { kind: 'srgb', value: themeSpec.background } };
slide.addShape('roundRect', {
  x: inches(themeSpec.margin), y: inches(1.5), width: inches(5.8), height: inches(3.8),
  fill: { kind: 'solid', color: { kind: 'srgb', value: themeSpec.surface } },
  line: { kind: 'line', color: { kind: 'srgb', value: themeSpec.secondary }, width: 1 },
});
slide.addRichText([{ runs: [{
  text: 'A living climate engine',
  style: { fontFamily: themeSpec.displayFont, fontSize: themeSpec.sizes.title, bold: true,
    color: { kind: 'srgb', value: themeSpec.text } },
}] }], { x: inches(0.6), y: inches(0.6), width: inches(12.1), height: inches(0.7), margin: 0 });
```

## Typography

Use predictable Office and LibreOffice pairs: Cambria + Arial, Cambria + Calibri, Arial + Arial, or Calibri + Calibri. Do not default to Aptos. Honor a user-specified face and reserve about 10% extra text-box width for substitution variance.

- Cover title: 50–64 pt.
- Slide title: 36–44 pt.
- Section heading: 20–24 pt.
- Body: 14–18 pt; prefer 16 pt.
- Captions and labels: 10–12 pt.

Reduce copy or split the slide before shrinking below these ranges. Use left-aligned body copy, real bullet paragraphs instead of Unicode bullets, and zero text-box margin when exact edge alignment matters.

## Spacing and layout families

Use a 16:9 canvas by default, at least 0.5-inch outer margins, and one fixed internal gap of either 0.3 or 0.5 inch for the whole deck. Establish shared alignment anchors before placing shapes. Give each slide one dominant focal point and a meaningful visual element: an available image, native chart, table, diagram, icon-like shape, or large typographic statistic.

Choose among eight layout families:

1. **cinematic cover** — minimal title, short framing line, and one dominant full-field composition;
2. **section divider** — a numbered or named transition with the deck motif and generous negative space;
3. **statement plus hero visual** — one declarative takeaway paired with the strongest available visual;
4. **asymmetric two-column** — unequal text/visual regions such as 40/60 or 35/65;
5. **large statistic plus explanation** — one oversized number with context and a compact evidence cue;
6. **comparison** — two aligned alternatives with shared measures and clear contrast;
7. **process or timeline** — a directional sequence built from native shapes and connectors;
8. **native chart or table plus takeaway** — editable evidence occupying most of the slide with a concise implication.

Do not repeat a family on consecutive content slides unless continuity requires it. Do not default to title-and-bullets, repeated card grids, title underlines, decorative bars, or accent stripes. Prefer one main idea, no more than three supporting points, and one clear visual hierarchy.

## Install and import

Install the current technical-preview release:

```sh
npm install @jiayunxie/pptx@next
```

Import the public entry point. Do not import private `dist` modules.

```js
import { degrees, inches, PptxDocument } from '@jiayunxie/pptx';
```

Node.js 20 or later is required. Modern browsers use the conditional browser export.

## Units, layouts, and formats

Presentation coordinates and sizes are EMU. Convert inches with `inches()` and angles with `degrees()` instead of hard-coding conversion factors.

```js
const document = PptxDocument.create({ slideSize: 'wide' });
const slide = document.addSlide();
slide.addShape('rect', {
  x: inches(0.75), y: inches(0.75),
  width: inches(4), height: inches(2),
  rotation: degrees(2),
});
```

Use the preset slide sizes when they fit. A custom size uses `{ width, height }` in EMU. All six supported presentation formats preserve the package's content and relationship semantics; choose the desired extension/output contract rather than renaming files after writing.

Plan the canvas before placing content. A wide deck is 13.333 by 7.5 inches. Keep an explicit safe margin and use a small coordinate helper layer in generators so alignment stays consistent.

## Create from zero

Native creation does not require PptxGenJS or a template.

```js
const document = PptxDocument.create({
  author: 'Research Team',
  company: 'Example',
  subject: 'Biodiversity',
  title: 'Amazon Rainforest',
  slideSize: 'wide',
});

const cover = document.addSlide();
cover.background = {
  kind: 'solid',
  color: { kind: 'srgb', value: '0B2E22' },
};
cover.addRichText([{ runs: [{
  text: 'Amazon biodiversity',
  style: {
    fontFamily: 'Cambria', fontSize: 56, bold: true,
    color: { kind: 'srgb', value: 'FFFFFF' },
  },
}] }], {
  x: inches(0.8), y: inches(4.7),
  width: inches(8.5), height: inches(0.8),
});
cover.addNotes('[Sources]\nhttps://example.org/source');

await document.writeFile('amazon-biodiversity.pptx', { compression: true });
```

Create the story outline before coding. One slide should answer one question. Use native text, shapes, charts, and tables so recipients can edit the result.

## Open and edit

Open an existing deck, work through live semantic models, and write a new file.

```js
const document = await PptxDocument.open('input.pptx');
const slide = document.slides[0];
slide.title.text = 'Updated safely';
slide.addNotes('Updated presenter guidance');

await document.writeFile('output.pptx', {
  compatibility: 'powerpoint-current',
  mode: 'strict',
});
```

Preserve the input. Unknown package parts and unsupported XML are byte-preserved during unrelated semantic edits. Use getters and setters on the model rather than unzipping and rewriting XML. For a narrow edit, diff the source and result to confirm that only expected parts changed.

Slides can be added, duplicated, moved, and deleted. Reopen the output and verify slide order, titles, notes, and any edited object. Never assume a successful write proves the presentation is visually correct.

## Text and rich text

Use `addText()` for ordinary content and `addRichText()` for paragraph/run-level control.

```js
slide.addRichText([{ runs: [{
  text: 'One clear message',
  style: {
    fontFamily: 'Cambria', fontSize: 40, bold: true,
    color: { kind: 'srgb', value: '113D2C' },
  },
}] }], {
  x: inches(0.8), y: inches(0.6),
  width: inches(11.7), height: inches(0.7),
  fit: 'shrink', wrap: true,
});

slide.addRichText([
  { runs: [
    { text: '42%', style: { fontSize: 18, bold: true, color: { kind: 'srgb', value: 'E4A83D' } } },
    { text: ' of the region remains under mounting pressure.', style: { fontSize: 18 } },
  ] },
  { bullet: true, runs: [{ text: 'Protect connected habitat.', style: { fontSize: 18 } }] },
], {
  x: inches(0.8), y: inches(1.7),
  width: inches(5.2), height: inches(2.2),
  fit: 'shrink',
});
```

Text models support alignment, vertical alignment, direction, language, RTL, margins, indents, wrapping, autofit/shrink behavior, fills, lines, arrows, shadows, rounded corners, and hyperlinks. Links may target URLs, email, or another slide. Use real bullet paragraphs rather than Unicode bullet characters.

Keep copy concise. If text needs aggressive shrinking, rewrite or split the slide. Avoid putting ordinary labels inside screenshots.

## Shapes and styling

`addShape()` supports the audited preset-shape catalog and strict custom geometry. Apply explicit fill, line, shadow, rotation, and hyperlink styles.

```js
slide.addShape('roundRect', {
  x: inches(0.8), y: inches(2),
  width: inches(3.3), height: inches(1.4),
  fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEBD8' } },
  line: { kind: 'line', color: { kind: 'srgb', value: '5D7A52' }, width: 1.25 },
  shadow: { kind: 'outer', color: { kind: 'srgb', value: '000000' }, opacity: 0.18, blur: 4, angle: 45, distance: 2 },
});
```

Shape fills support no-fill and solid color. Slide, layout, and master backgrounds additionally support the documented direct gradient and image forms. Lines support colors, transparency, width, dashes, and arrowheads. Getters return detached snapshots; assign a new value to update a style. Use scheme colors when a theme should drive recoloring and sRGB values when exact brand color is required.

Prefer a few large shapes that support the composition. Avoid decorative clutter, excessive shadows, and repetitive card grids.

## Images, SVG, and backgrounds

Add raster imagery from a file, URL, data URI, bytes, browser `Blob`/`File`, streams, or async byte iterables. Signatures, not file extensions, determine PNG/JPEG/GIF content.

```js
await document.addImage(0, 'assets/canopy.png', {
  x: inches(0), y: inches(0),
  sizing: { type: 'cover', width: inches(13.333), height: inches(7.5) },
  altText: 'Aerial view of dense Amazon rainforest canopy',
});
```

Use `contain` when the full image must remain visible, `cover` for edge-to-edge frames, and crop sizing for a specific pixel region. Set accessible alt text. Images can be transformed, rounded, made transparent, shadowed, and replaced without breaking unrelated relationships.

SVG creation and loading are supported through the public SVG APIs. A live SVG image exposes semantic metadata, transform, style, hyperlink, and paired SVG/PNG payload replacement; arbitrary SVG DOM subtree editing is intentionally outside the API. Prefer SVG for diagrams and icons when practical. Do not use programmatic vector drawings to imitate photographic imagery; use sourced or generated raster visuals for that role.

Use `document.setSlideBackgroundImage(slideIndex, source)` for a full slide image background. `slide.background`, layout backgrounds, and master backgrounds support semantic direct background editing. Clearing with `undefined` restores inheritance.

## Tables and pagination

Create editable tables with strings or rich cell values.

```js
slide.addTable([
  ['Layer', 'Example', 'Function'],
  ['Canopy', 'Brazil nut tree', 'Carbon and habitat'],
  ['Forest floor', 'Fungi', 'Nutrient cycling'],
], {
  x: inches(0.8), y: inches(1.6),
  columnWidths: [inches(2), inches(3), inches(5)],
  rowHeights: [inches(0.5), inches(0.65), inches(0.65)],
  border: { kind: 'line', color: { kind: 'srgb', value: '86A882' }, width: 1 },
  fill: { kind: 'solid', color: { kind: 'srgb', value: 'F2F6EC' } },
});
```

Tables support row and column geometry, cell merges, rich text, links, typography, alignment, borders, fills, margins, text direction, vertical alignment, fit, repeated headers, and automatic pagination. Existing tables expose semantic cell, row, column, merge, and bulk-style editors.

Use automatic pagination when rows cannot fit legibly. Configure repeated header rows and inspect every generated page. In a browser, `document.tableToSlides(elementId, options)` converts an HTML table into ordinary editable slides. It can paginate into new slides and use layout placeholder geometry. Do not squeeze a long table onto one slide.

## Native charts

`slide.addChart()` creates native, editable PowerPoint charts. Use a single chart type or a combo-chart definition.

```js
const chart = await slide.addChart('bar', [
  { name: 'Area', categories: ['2000', '2010', '2020'], values: [100, 84, 71] },
], {
  x: inches(0.9), y: inches(1.7),
  width: inches(7.2), height: inches(4.7),
});
await chart.replaceDefinition({
  groups: [{
    type: 'bar',
    series: [{ name: 'Area', categories: ['2000', '2010', '2020'], values: [100, 84, 71] }],
    options: { dataLabels: { showValue: true, position: 'outsideEnd' } },
  }],
  options: { legend: { visible: true, position: 'bottom' } },
});
```

Charts are backed by native chart/workbook parts and can be reopened and semantically edited. Use the chart type that expresses the relationship: bars for comparison, lines for change over time, and pies only for a small part-to-whole relationship. Label units and sources. Do not render charts as screenshots unless the user explicitly needs a fixed visual.

## Audio and video

Use `document.addAudio()` and `document.addVideo()` for supported embedded media. Audio supports MP3, M4A, WAV, and OGG. Embedded video supports MP4/M4V, MOV, and WebM. Posters support PNG, JPEG, and GIF.

```js
await document.addVideo(0, 'assets/forest.mp4', {
  x: inches(7.8), y: inches(1.4),
  width: inches(4.7), height: inches(2.8),
  poster: 'assets/forest-poster.png',
});
```

HTTP(S) video URLs remain external and are not fetched. Provide an embedded local poster for predictable appearance. Confirm the user's target clients before relying on codecs or playback behavior.

## Masters, layouts, and placeholders

`document.masters` and `document.layouts` expose live semantic wrappers. They support direct backgrounds, slide numbers, shapes, placeholders, text, rich text, images, SVG, and charts. Use a master/layout when several slides share structure or branding.

Select the intended layout when adding slides. Prefer placeholders for title and body regions so the deck remains easy to edit. When replacing or deleting masters, use the semantic methods so slides and owned relationships are retargeted safely. Do not duplicate the same brand furniture manually on every slide unless it is deliberately slide-specific.

## Metadata, sections, notes, and slide numbers

Set document metadata on creation or through the document properties: author, company, title, subject, created/modified times, last editor, revision, RTL mode, and slide size.

```js
const section = document.addSection({ title: 'Species' });
slide.addNotes([
  '[Sources]',
  'https://example.org/report — statistic used on slide',
  'Local asset: assets/canopy.png — supplied by research team',
].join('\n'));

slide.slideNumber = {
  x: inches(12.2), y: inches(7.0),
  width: inches(0.5), height: inches(0.25),
  align: 'right', style: { fontSize: 10 },
};
```

Sections organize the presentation without flattening slide semantics. Notes are plain text and are independently owned by each slide. Use `[Sources]` as a distinct block so provenance is discoverable. Slide numbers can be defined on slides, layouts, or masters. Reopen the deck to confirm ownership and inheritance.

## Output and runtime environments

Write directly to a file in Node:

```js
await document.writeFile('deck.pptx', { compression: true });
```

`write({ outputType })` supports `arraybuffer`, `base64`, `binarystring`, `blob`, `nodebuffer`, and `uint8array`. `writeBlob()` returns a browser presentation Blob, `download()` triggers browser download, and `stream()` returns a backpressure-aware Node Readable after the package has been assembled. Use `compression: true` for DEFLATE or `false` for STORE; non-boolean values are rejected.

```js
const bytes = await document.write({ outputType: 'uint8array', compression: true });
const blob = await document.writeBlob({ compression: true });
```

An opened and completely unchanged presentation can preserve its original bytes when compression is omitted. An explicit compression value repacks the ZIP but does not change OOXML semantics.

## Inspection, validation, and diff

Use the installed CLI or repository build for stable JSON diagnostics.

```sh
pptx-inspect --json doctor
pptx-inspect --json slides list deck.pptx
pptx-inspect --json package inspect deck.pptx
pptx-inspect --json package validate deck.pptx --profile powerpoint-2010
pptx-inspect --json package diff before.pptx after.pptx
```

Inspection answers what is in the package; validation checks relationships, package structure, and compatibility diagnostics; diff confirms mutation scope. A zero-error package can still have visual defects, so always render after validation.

For narrow edits, inspect and validate both before and after. Preserve the original, use a new output file, and investigate any unexpected changed part before delivery.

## Narrative and visual design

Start from the audience and desired decision, not from a slide template. Write a short narrative arc: context, evidence, implication, response. Give each slide a declarative takeaway and ensure the sequence makes sense when only titles are read.

Choose a deliberate visual system:

- one display face and one body face, or a disciplined single-family hierarchy;
- a small semantic palette with accessible contrast;
- consistent margins, baseline rhythm, and alignment anchors;
- a controlled photography or illustration style;
- varied layouts that serve the content rather than repeating a card grid.

Use image composition intentionally. Request negative space where text must sit, crop around the focal subject, and apply a controlled overlay when text crosses photography. Do not reuse the same foreground image more than once. Prefer full-bleed visuals, diagrams, and native data graphics to dense paragraphs.

Minimum practical text sizes are 50 pt for a cover title, 36 pt for slide titles, 20 pt for section headings, 14 pt for body copy, and 10 pt for captions. These are floors, not targets. Reduce copy before reducing type.

## 180-second execution and QA

The warm path starts when the query is available and ends when the final PPTX passes content, structural, and visual checks.

| Stage | Budget |
|---|---:|
| Narrative outline | 20 seconds |
| Theme and layout assignment | 25 seconds |
| One native PPTX generator execution | 75 seconds |
| Reopen, validate, render, and inspect in parallel | 45 seconds |
| One repair and targeted recheck buffer | 15 seconds |
| Total | 180 seconds |

Run content QA by reopening the written file and verifying slide count, order, titles, notes, metadata, links, chart values, table cells, alt text, and `[Sources]` blocks. In parallel, run package inspection and compatibility validation. For edits, also diff against the untouched source.

Render every slide once, create a montage, and inspect it for story flow and repetitive composition. Inspect readable slide images for overlap, clipping, unexpected wrapping, low contrast, awkward crops, tiny text, misalignment, and excess empty space. Run the available automated overflow check, but do not treat it as a substitute for visual inspection.

If the first output has no defect, deliver it without inventing a correction. If a concrete defect exists, make at most one targeted repair within the fast path, rerender only affected slides, repeat affected semantic checks, and run final package validation. Report a missed time budget as a performance defect rather than skipping structural or visual evidence.

## Deliberate boundaries

The repository audits the PptxGenJS 4.0.1 public capability surface and provides additional lossless semantic editing. That is API-surface evidence, not a universal rendering certification.

Do not claim broad stable certification for PowerPoint, Keynote, Google Slides, or LibreOffice unless the current release has independent client evidence. Use `powerpoint-2010` or another documented validation profile as a structural compatibility check, not as proof of identical rendering in every client.

Provider-specific online-video metadata, remote-fetch embedding, trim/bookmarks, finite repeats, captions/subtitles, built-in transcoding, rich notes-page editing, comments, and some advanced native-client extensions remain deliberate boundaries. Preserve unknown content during unrelated edits and explain any limitation that affects the request.

## Common mistakes

- Overwriting the source instead of writing a separate output file.
- Editing ZIP/XML directly when a semantic API exists.
- Using raw numbers as inches rather than converting with `inches()`.
- Assuming package validation replaces rendering and visual review.
- Delivering a render that contains an observed defect without using the targeted repair buffer.
- Filling slides with prose or shrinking text below a readable size.
- Repeating one card layout or one foreground image across the deck.
- Stretching images instead of using `cover`, `contain`, or a deliberate crop.
- Rasterizing editable text, tables, or charts.
- Omitting alt text and `[Sources]` notes.
- Treating an external video URL as embedded media.
- Claiming universal client compatibility from an OOXML validator alone.
