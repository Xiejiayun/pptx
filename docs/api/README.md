# Public API 0.1

## Open and save

```ts
import { PptxDocument } from '@pptx/sdk';

const document = await PptxDocument.open('input.pptx', {
  limits: { maxPartBytes: 128 * 1024 * 1024 },
  signal: abortController.signal,
});

await document.writeFile('output.pptx', {
  compatibility: 'powerpoint-2010',
  mode: 'strict',
});
```

Inputs: file path, `Buffer`/`Uint8Array`, `ArrayBuffer`, or Node readable stream. `write()` returns `Uint8Array`.

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
