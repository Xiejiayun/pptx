# Optional feature plugins

Plugins are internal `0.1.1` workspace packages bundled into the public aggregate package. Core never imports them; when unused, their OOXML and related parts remain opaque and losslessly preserved.

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

Decode the timing tree, add appear/fade/wipe/fly or motion-path effects, model triggers/delay/duration/repeat/text ranges, validate shape targets, and retarget shape ids. Core media creation and `MediaModel.settings` now write native timing directly without this plugin. Installation remains backward-compatible: it uses the shared media timing codec to materialize legacy preference-only files once, leaves healthy/native-only or unsafe imports unchanged, and shares the slide-wide timing ID allocator with ordinary animations.

## Advanced charts

Core now creates, reads, semantically edits, duplicates, and deletes all nine standard PptxGenJS chart types plus compatible primary/secondary combinations while keeping caches, A1 formulas, and embedded workbooks synchronized. For recognized standard charts, the plugin delegates series-value changes to `ChartModel.replaceDefinition()` so the core synchronization and clone-on-write rules remain authoritative.

The plugin adds inspection for modern charts and advanced series state such as trendlines, error bars, and data labels. Raw or advanced XML edits can intentionally diverge from an embedded workbook; diagnostics remain until a semantic core replacement or explicit `replaceEmbeddedWorkbook()` proves equality. Image fallback is created only by an explicit method call. Office 2016 `cx:*` creation/editing, external-workbook editing, and chart animations remain preservation-only.

## SmartArt

Resolve data/layout/quick-style/colors/fallback-drawing parts, read/replace text, and add/delete nodes plus parent connections. Editing data with a fallback drawing produces `SMARTART_RELAYOUT_REQUIRED`; style/color/drawing parts are not silently rebuilt or removed.
