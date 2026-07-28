# Public API 0.1

## Create, open, and save

```ts
import { inches, PptxDocument } from '@pptx/sdk';

const created = PptxDocument.create({
  format: 'pptx',
  slideSize: '16:9',
});
created.addSlide();
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

`create()` is synchronous and starts with zero slides plus a default master, blank layout, theme, notes master, and document properties. Built-in slide sizes are `4:3`, `16:9` (the default), `16:10`, and `wide`; a `{ width, height }` value accepts any OOXML-valid 1–56 inch dimensions in EMU. `document.slideSize` reads or changes the slide canvas without scaling existing shapes or changing the notes page. All six presentation formats can be created without using PptxGenJS; macro-enabled formats start without a VBA project.

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

Shape kinds include `text`, `shape`, `image`, `table`, `chart`, `graphic-frame`, and `group`. Images expose embedded part URIs and replacement; tables expose rows/cells and cell text editing; charts expose cached series and lossless chart XML editing.

```ts
const text = document.addSlide().addText('Quarterly results\nQ4 forecast', {
  name: 'Heading',
  x: inches(1),
  y: inches(0.75),
  width: inches(8),
  height: inches(1),
  align: 'center',
  valign: 'middle',
  wrap: true,
});
text.text = 'Updated results\nApproved';
text.verticalAlignment = 'bottom';
text.textWrap = false;
```

`addText()` creates plain-text paragraphs with name and transform options. CRLF and CR normalize to LF; consecutive and trailing line breaks remain empty paragraphs. Setting `.text` replaces the visible text using the first paragraph as the style template. Use the structured API below when run styles must remain distinct.

```ts
const rich = document.addSlide().addRichText([
  {
    align: 'right',
    bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
    level: 2,
    spacing: { before: 6, after: 8, line: { kind: 'multiple', factor: 1.5 } },
    tabStops: [{ position: 2.5, alignment: 'decimal' }],
    runs: [
      { text: 'Revenue\t', style: { bold: true, fontSize: 24 } },
      { text: '+18%', style: { italic: true, color: { kind: 'srgb', value: '00A651' } } },
    ],
  },
]);
rich.richText = [{ runs: [{ text: 'Approved', style: { color: { kind: 'scheme', value: 'accent1' } } }] }];
```

`richText` is an immutable paragraph/run value snapshot. It reads, creates, and replaces each paragraph's alignment, Unicode bullet or automatic numbering, list level, paragraph spacing, tab stops, and run styles. `AddTextOptions` supplies paragraph creation defaults; each `RichTextParagraph` can override them, with `bullet: false`, `level: 0`, `spacing: false`, or `tabStops: false` suppressing the corresponding default. Bullets support a custom character and 0–4032pt per-level indent. Numbering supports the 16 PptxGenJS styles, a 1–32767 start value, and indent. List `level` is zero-based from 0–8; nested bullet margin is `indent × (level + 1)`. Spacing uses points for before/after/exact and a factor such as `1.5` for multiple line spacing. Tab stop positions use inches and support left, center, right, and decimal alignment; `[]` is an explicit empty list. Run styles include font family, point size, bold, italic, sRGB/theme color, and soft breaks. Setting `richText` preserves text-body metadata and unrelated same-position paragraph properties but intentionally replaces old runs and supported paragraph values; arbitrary paragraph indentation, hyperlinks, and advanced typography are separate capabilities.

Text-box `margin` values use points and accept one number, a `[top, right, bottom, left]` tuple, or a named object. `shape.textMargins` reads only direct `bodyPr` overrides as a named object; assigning a scalar/tuple/object replaces the four supported direct sides, while `undefined` or `{}` clears them. This is separate from paragraph `marL`, first-line indent, and bullet hanging indent. Unlike PptxGenJS 4.0.1's asymmetric-tuple runtime bug, native tuple creation follows the documented TRBL order.

Text-box `valign` accepts `top`, `middle`, or `bottom`; omission creates an explicit middle anchor. `shape.verticalAlignment` reads or replaces only the direct text-body anchor, and assigning `undefined` removes that direct override while preserving margins, autofit metadata, and other text-body content.

Text-box `wrap` accepts a boolean; omission and true create explicit automatic wrapping, while false keeps text on an unwrapped line. `shape.textWrap` reads or replaces only the direct text-body wrapping token, and assigning `undefined` removes that direct override without changing fit, margins, vertical alignment, or text content.

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
