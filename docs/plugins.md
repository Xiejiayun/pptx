# Optional feature plugins

Plugins are independent `0.1.0` packages. Core never imports them; without installation their OOXML and related parts remain opaque and losslessly preserved.

```ts
import { PptxDocument } from '@pptx/sdk';
import { installTransitionPlugin } from '@pptx/plugin-transitions';
import { installAnimationPlugin } from '@pptx/plugin-animations';
import { installAdvancedChartPlugin } from '@pptx/plugin-advanced-charts';
import { installSmartArtPlugin } from '@pptx/plugin-smartart';

const document = await PptxDocument.open('input.pptx');
const transitions = installTransitionPlugin(document);
const animations = installAnimationPlugin(document);
const charts = installAdvancedChartPlugin(document);
const smartArt = installSmartArtPlugin(document);
```

## Transitions

Read/write common `<p:transition>` effects, speed, duration, click/automatic advance, and transition sound relationships. Morph/extensions are preserved; PowerPoint 2010 receives a diagnostic, and new Morph creation is blocked rather than emitting invalid XML.

## Animation and timing

Decode the timing tree, add appear/fade/wipe/fly or motion-path effects, model triggers/delay/duration/repeat/text ranges, validate shape targets, and retarget shape ids. Installation converts Media codec autoplay/loop/volume preferences into native `cMediaNode` timing nodes.

## Advanced charts

Inspect combination/modern chart types, axes, series, trendlines, error bars, and data labels. Cached values can be edited; when an embedded workbook exists, divergence produces a diagnostic until `replaceEmbeddedWorkbook()` is called. Image fallback is only created by an explicit method call.

## SmartArt

Resolve data/layout/quick-style/colors/fallback-drawing parts, read/replace text, and add/delete nodes plus parent connections. Editing data with a fallback drawing produces `SMARTART_RELAYOUT_REQUIRED`; style/color/drawing parts are not silently rebuilt or removed.

