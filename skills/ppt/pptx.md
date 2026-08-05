# Native PPTX creation and editing

This reference explains how to create and edit presentations with `@jiayunxie/pptx`. It adapts the narrative, visual-design, and render-review discipline of Anthropic's presentation guidance to this repository's native semantic APIs.

## Contents

1. [Install and import](#install-and-import)
2. [Units, layouts, and formats](#units-layouts-and-formats)
3. [Create from zero](#create-from-zero)
4. [Open and edit](#open-and-edit)
5. [Text and rich text](#text-and-rich-text)
6. [Shapes and styling](#shapes-and-styling)
7. [Images, SVG, and backgrounds](#images-svg-and-backgrounds)
8. [Tables and pagination](#tables-and-pagination)
9. [Native charts](#native-charts)
10. [Audio and video](#audio-and-video)
11. [Masters, layouts, and placeholders](#masters-layouts-and-placeholders)
12. [Metadata, sections, notes, and slide numbers](#metadata-sections-notes-and-slide-numbers)
13. [Output and runtime environments](#output-and-runtime-environments)
14. [Inspection, validation, and diff](#inspection-validation-and-diff)
15. [Narrative and visual design](#narrative-and-visual-design)
16. [Required QA loop](#required-qa-loop)
17. [Deliberate boundaries](#deliberate-boundaries)
18. [Common mistakes](#common-mistakes)

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
    fontFamily: 'Aptos Display', fontSize: 34, bold: true,
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
    fontFamily: 'Aptos Display', fontSize: 28, bold: true,
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
  'Generated image: canopy prompt, 2026-08-05',
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

Minimum practical text sizes are approximately 50 pt for a cover title, 35 pt for slide titles, 24 pt for subheads, and 16 pt for body copy. These are floors, not targets. Reduce copy before reducing type.

## Required QA loop

Use a two-part QA process.

Content QA:

1. Reopen the written file.
2. Verify slide count, order, titles, notes, metadata, links, chart values, table cells, alt text, and source blocks.
3. Run package inspection and compatibility validation.
4. For edits, run a package diff against the untouched source.

Visual QA:

1. Render every slide to images.
2. Create a montage to review story flow, rhythm, and repetition.
3. Inspect each slide at readable resolution for overlap, clipping, unexpected wrapping, low contrast, awkward crops, tiny text, misalignment, and excess empty space.
4. Run automated slide tests where available, but do not substitute them for visual inspection.
5. Make at least one concrete improvement found during the first review.
6. Regenerate, revalidate, rerender, and inspect the corrected deck.

The final deliverable is the corrected PPTX, not the first file that writes successfully.

## Deliberate boundaries

The repository audits the PptxGenJS 4.0.1 public capability surface and provides additional lossless semantic editing. That is API-surface evidence, not a universal rendering certification.

Do not claim broad stable certification for PowerPoint, Keynote, Google Slides, or LibreOffice unless the current release has independent client evidence. Use `powerpoint-2010` or another documented validation profile as a structural compatibility check, not as proof of identical rendering in every client.

Provider-specific online-video metadata, remote-fetch embedding, trim/bookmarks, finite repeats, captions/subtitles, built-in transcoding, rich notes-page editing, comments, and some advanced native-client extensions remain deliberate boundaries. Preserve unknown content during unrelated edits and explain any limitation that affects the request.

## Common mistakes

- Overwriting the source instead of writing a separate output file.
- Editing ZIP/XML directly when a semantic API exists.
- Using raw numbers as inches rather than converting with `inches()`.
- Assuming package validation replaces rendering and visual review.
- Delivering the first render without a correction pass.
- Filling slides with prose or shrinking text below a readable size.
- Repeating one card layout or one foreground image across the deck.
- Stretching images instead of using `cover`, `contain`, or a deliberate crop.
- Rasterizing editable text, tables, or charts.
- Omitting alt text and `[Sources]` notes.
- Treating an external video URL as embedded media.
- Claiming universal client compatibility from an OOXML validator alone.
