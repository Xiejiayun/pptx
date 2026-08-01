# Chart Creation and Semantic Editing Design

## Goal

Complete the standard-chart phase of the PptxGenJS parity route: create charts from a zero-slide native presentation,
read and semantically edit existing chart data, keep chart caches and the embedded workbook synchronized, and manage
the complete chart dependency graph through duplicate, move, delete, write, and reopen.

The phase covers the nine PptxGenJS 4.0.1 public chart types (`area`, `bar`, `bar3D`, `bubble`, `doughnut`, `line`,
`pie`, `radar`, and `scatter`), compatible combination charts, their public data shapes, and the chart options needed
for title, legend, axes, gridlines, labels, data table, grouping, type-specific presentation, and chart/plot styling.

The public API does not need to reproduce PptxGenJS parameter objects byte for byte. It must provide equivalent
creation and editing capability through the native `Presentation -> Slide -> ChartModel` object model, and the
PptxGenJS adapter must prove valid public-output conformance.

This phase does not create Office 2016 modern `cx:*` charts, externally linked workbooks, chart animations, or
PowerPoint-only automatic layout artifacts. Those inputs remain losslessly preserved and diagnosed. Modern-chart
inspection stays in the advanced-charts plugin.

## Evidence and Current-State Audit

The repository already has useful chart foundations, but no native creation path:

- `ChartModel.chartPartUri` resolves a chart relationship from an existing graphic frame;
- `ChartModel.series` reads a shallow name/category/value projection from chart caches;
- `ChartModel.setXml()` validates XML and performs relationship-aware clone-on-write for a shared chart part;
- slide duplication deep-clones chart-owned dependencies, including embedded workbooks, styles, and colors;
- chart and workbook dependency garbage collection is already classified as owned;
- the advanced-charts plugin can inspect chart types and axes, patch numeric cache values, append trendlines,
  error bars, and data labels, replace a workbook payload, and emit a cache/workbook divergence diagnostic.

The missing behaviors are structural:

- no `SlideModel.addChart()` or `PptxDocument.addChart()` exists;
- no stable public chart-type or chart-definition type exists;
- the current reader does not distinguish categorical, scatter, bubble, or combination data;
- cache edits do not update the workbook and can knowingly create divergent state;
- semantic edits bypass chart clone-on-write and package-level atomicity;
- no chart-specific validation participates in `PptxDocument.write()`;
- no chart object deletion API removes the graphic frame and owned dependency subgraph;
- packed Node/browser/type smoke and real-client evidence cover only raw XML editing.

PptxGenJS 4.0.1 produces one chart part, one chart relationship part, and one embedded XLSX per public chart object.
Its chart XML always carries `c:externalData`, formulas, and caches. The nine public chart types map to
`areaChart`, `barChart`, `bar3DChart`, `bubbleChart`, `doughnutChart`, `lineChart`, `pieChart`, `radarChart`, and
`scatterChart`. Valid categorical charts use name/category/value caches; scatter uses x/y numeric caches; bubble uses
x/y/size numeric caches. Combination charts place multiple direct chart-type elements in one plot area and may use
primary and secondary axis pairs.

## Approaches Considered

### Delegate creation to PptxGenJS

Generate a temporary PptxGenJS presentation, copy its chart parts into the native package, and continue editing with
the existing model. This would accelerate initial output, but it would make the aggregate package depend on
PptxGenJS at runtime and would fail the repository rule that the native package itself must satisfy the parity matrix.

### Create cache-only charts

Write chart XML and formulas without an embedded workbook. PowerPoint and LibreOffice can display many such files
from caches, but Edit Data, future semantic mutation, and workbook/cache consistency would be unreliable. This is a
display-only shortcut and does not meet the chart phase completion gate.

### Native chart codec plus deterministic workbook generator -- selected

Add strict chart state, rendering, and editing modules to `@pptx/model`. Build a minimal XLSX with `OpcPackage` so
the implementation uses the existing browser-safe ZIP/OPC dependency and adds no runtime package. Generate workbook
bytes before the synchronous presentation transaction, then commit the chart part, workbook part, relationships, and
slide graphic frame atomically.

This approach preserves the current dependency lifecycle, supports both Node and browsers, gives semantic editing one
source of truth, and permits direct PptxGenJS output comparison without importing PptxGenJS into production code.

## Public API

### Type catalog

`@pptx/model` exports a frozen runtime catalog and matching literal type:

```ts
export const CHART_TYPES = [
  'area', 'bar', 'bar3D', 'bubble', 'doughnut', 'line', 'pie', 'radar', 'scatter',
] as const;

export type ChartType = typeof CHART_TYPES[number];
export type ChartCategory = string | number;
export type ChartCategories =
  | readonly ChartCategory[]
  | readonly (readonly string[])[];
```

Multi-level categories are represented outermost by level and innermost by point, matching the normalized
PptxGenJS `labels: string[][]` state. Every level must have the same non-zero point count.

### Series input and snapshot

One value shape is accepted for all chart types; the selected type determines the legal fields:

```ts
export interface ChartSeriesInput {
  readonly name: string;
  readonly categories?: ChartCategories;
  readonly values: readonly number[];
  readonly xValues?: readonly number[];
  readonly sizes?: readonly number[];
}

export interface ChartSeries extends ChartSeriesInput {
  readonly categories?: ChartCategories;
  readonly values: readonly number[];
  readonly xValues?: readonly number[];
  readonly sizes?: readonly number[];
}
```

- area, bar, bar3D, doughnut, line, pie, and radar require `categories` and `values`, and reject `xValues`/`sizes`;
- scatter requires equal-length `xValues` and `values`, and rejects categories/sizes;
- bubble requires equal-length `xValues`, `values`, and positive finite `sizes`, and rejects categories;
- names are non-empty XML-safe strings; every numeric value is finite; sparse arrays and accessors are rejected;
- pie and doughnut accept exactly one series, matching their single-ring public behavior for this phase;
- snapshots are deeply detached and frozen.

Unlike PptxGenJS's special first-series X-vector convention for scatter and bubble, each native series carries its own
X values. The adapter conformance layer maps the shared PptxGenJS X vector to each projected native series. This is a
strict superset and avoids coupling unrelated series lengths.

### Definition and combination groups

```ts
export interface ChartGroupInput {
  readonly type: ChartType;
  readonly series: readonly ChartSeriesInput[];
  readonly axis?: 'primary' | 'secondary';
  readonly options?: ChartGroupOptions;
}

export interface ChartDefinitionInput {
  readonly groups: readonly ChartGroupInput[];
  readonly options?: ChartOptions;
}

export interface ChartDefinition {
  readonly groups: readonly Readonly<ChartGroupInput>[];
  readonly options: Readonly<ChartOptions>;
}
```

A normal chart has one group. A combination chart has two or more groups. Pie, doughnut, bar3D, and bubble groups do
not combine in this phase because their plot-area and axis grammars are not safely composable with the PptxGenJS
combination surface. Axis-based area/bar/line/radar groups may share a primary axis pair; one or more groups may use a
secondary pair. Scatter may combine only with scatter groups because it requires value/value axes.

The creation overloads retain an easy PptxGenJS migration path:

```ts
await slide.addChart('bar', series, options);
await slide.addChart([
  { type: 'bar', series: columns },
  { type: 'line', series: trend, axis: 'secondary' },
], options);

await document.addChart(0, 'line', series, options);
```

Both methods return `Promise<ChartModel>` because the embedded XLSX is generated before the package transaction.

### Live model

`ChartModel` retains `chartPartUri`, `xml`, and guarded `setXml()` for lossless escape-hatch editing, then adds:

```ts
class ChartModel extends BaseShapeModel {
  get workbookPartUri(): string | undefined;
  get definition(): Readonly<ChartDefinition> | undefined;
  get series(): readonly Readonly<ChartSeries>[];
  get diagnostics(): readonly ChartDiagnostic[];
  async replaceDefinition(value: ChartDefinitionInput): Promise<this>;
  async replaceSeries(value: readonly ChartSeriesInput[]): Promise<this>;
  remove(): void;
}
```

`definition` is `undefined` when the chart is modern, malformed, ambiguous, or outside the supported grammar. The
legacy `series` projection remains available and flattens recognized groups in plot order; it returns an empty frozen
array for unsupported state. `replaceSeries()` is allowed only for a recognized single-group chart and retains its
type/options. `replaceDefinition()` is whole replacement for supported semantic chart state. Raw `setXml()` remains
explicitly capable of making workbook caches diverge and therefore produces a diagnostic until a semantic replacement
or explicit advanced-plugin workbook replacement restores consistency.

## Strict Chart Recognition

The reader claims semantic ownership only when all required elements are direct and namespace-correct:

- exactly one direct `c:chartSpace/c:chart/c:plotArea` chain;
- one or more supported direct chart-type children in plot order;
- exactly one direct `c:ser` sequence per group, with unique non-negative `c:idx` and `c:order` values;
- exactly one supported series-name cache and the correct category/value or x/y/size cache set for the group type;
- one direct `c:ptCount` per cache whose value equals the number of unique direct points;
- point indexes are canonical non-negative integers, unique, in range, and returned in numeric order;
- numeric values are strict finite decimal lexical values; category strings preserve exact text;
- formulas are optional for cache-only imported charts, but when present they must be ordinary `Sheet1` A1 ranges;
- chart axis IDs are positive unsigned integers, unique by axis object, and every group reference resolves;
- `c:externalData`, when present, resolves through exactly one internal package relationship to one XLSX part.

Unknown style children, extension lists, rich title/legend text, colors, and client-specific ordering outside owned data
spans are preserved. Unsupported or ambiguous structure is never coerced into a partial definition. Read operations
do not change bytes, relationships, the package journal, or model identity.

The reader exposes an internal state with `recognized`, `cache-only`, `workbook-divergent`, `modern`, `unsupported`,
and `ambiguous` statuses. Diagnostics and mutators consume that state instead of repeating broad descendant searches.

## Deterministic Embedded Workbook

The workbook generator uses `OpcPackage.create()` and emits a minimal valid XLSX with:

- root office-document relationship;
- `/xl/workbook.xml`, its worksheet relationship, and `/xl/worksheets/sheet1.xml`;
- `/xl/styles.xml` with General/string-compatible defaults;
- inline strings instead of a shared-string table, eliminating index drift during edits;
- no volatile timestamps, calculation chain, external links, macros, tables, or unused theme parts;
- deterministic row/column order and stable ZIP entry order.

Categorical data uses category level columns followed by one value column per series. Scatter uses an X and Y column
pair per series. Bubble uses X, Y, and size columns per series. Combination groups share one worksheet and allocate
columns in group/series order. Chart formulas are rendered from the same column plan that renders the workbook, so
cache and worksheet ranges cannot be computed independently.

The generator returns bytes without touching the presentation. Creation and semantic editing await those bytes first,
then enter one `OpcPackage.transaction()` to update the presentation. If workbook generation fails, the presentation
is unchanged.

On semantic edit of a recognized imported chart, the native implementation replaces or creates the embedded workbook
with the canonical deterministic workbook. It does not patch arbitrary external formulas. That normalization is
limited to the chart-owned workbook and owned cache/formula spans; unrelated chart styling and extension bytes stay
unchanged.

## Canonical Chart and Graphic-Frame Encoding

Creation adds, in one transaction:

1. an allocated `/ppt/charts/chartN.xml` part with the standard chart content type;
2. an allocated `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx` part;
3. a chart-to-workbook package relationship and `c:externalData` reference;
4. a slide-to-chart relationship;
5. a `p:graphicFrame` before a direct slide `p:extLst`, if present.

The frame contains a unique positive shape ID, escaped name/alt text, `a:xfrm`, and a chart graphic-data URI. Defaults
are x/y 1 inch, width 6 inches, height 4 inches, zero rotation, and no flips. Public coordinates are EMU, consistent
with the rest of the native model.

Every chart part contains canonical chart-space namespaces, date system, chart, plot area, direct group elements,
caches/formulas, required axes, plot visibility, blank handling, chart/plot shape properties, and external data.
Axis IDs are allocated deterministically within the chart part and are independent of slide shape IDs.

Creation never copies PptxGenJS random UUIDs, timestamps, application metadata, malformed defaults, or redundant
style bytes. Conformance compares public semantics, relationships, workbook cell values, formulas, transforms, and
supported options rather than package byte identity.

## Supported Chart Options

`ChartOptions` is descriptor-safe, detached before workbook generation, and divided by ownership:

- frame: `name`, `altText`, x/y/width/height, rotation, horizontal/vertical flip;
- title: text, visibility, overlay, font face/size/bold/italic/color, rotation, and manual position;
- legend: visibility, position, overlay, font face/size/bold/italic/color;
- chart and plot areas: strict none/solid fill plus strict none/solid line;
- series colors, opacity, border, and supported line/marker settings;
- data labels: value/category/series/percent/bubble flags, position, number format, leader lines, and font;
- category/value/secondary axes: visibility, position, title, bounds, major/minor unit, logarithmic base, number format,
  orientation, label position/rotation/font, line, ticks, and major/minor gridlines;
- data table: visibility, horizontal/vertical borders, outline, legend keys, font size, and number format;
- common: display blanks, language, style number, right-angle/rotation/perspective for bar3D;
- type-specific: area/bar grouping, bar direction/gap/overlap, doughnut hole size, pie/doughnut first-slice angle,
  line smoothing/marker, and radar style.

Each option has one canonical default documented in its public type tests. Values are validated before workbook work.
Omitted options do not copy PptxGenJS truthy fallbacks; explicit zero and false remain representable where OOXML
allows them. Unsupported combinations reject before mutation.

Chart option editing is initially exposed through `replaceDefinition()`. Narrow live accessors may be added later only
when they preserve the same state reader, validators, and transaction semantics.

## Semantic Mutation and No-Op Rules

`replaceDefinition()` performs these stages:

1. detach and normalize the complete input without reading accessors;
2. require a recognized standard chart and plan compatible plot/axis ownership;
3. compare normalized public state to the current definition; return exact no-op when both caches and workbook are
   already synchronized;
4. generate the canonical workbook bytes outside the presentation transaction;
5. enter one transaction, re-resolve the live shape, clone the owned chart subgraph when shared, patch owned chart
   data/type/options, replace the workbook, and retarget the graphic frame when cloning;
6. validate the resulting chart graph before returning the same `ChartModel` identity.

Errors in normalization, workbook generation, chart parsing, clone allocation, relationship creation, XML insertion,
part writing, or validation leave parts, content types, relationships, slide/chart XML, ZIP state, journal, and object
identity unchanged.

Changing a single-group chart among compatible category/value types preserves reusable axes and styles where legal.
Switching between axis and axis-free types replaces only the owned plot-area chart/axis spans. Scatter/bubble or combo
conversions use whole supported plot-area replacement because their data and axis grammars differ.

## Duplicate, Move, Delete, and Garbage Collection

Existing slide duplication remains the source of truth: it deep-clones the chart part and every owned internal
dependency, including the workbook, while preserving shared/external dependencies. Semantic edits after duplication
therefore isolate by chart part; imported shared charts still use mutation-time clone-on-write.

Moving a slide changes only presentation order and leaves chart relationships untouched. Stable `ChartModel` handles
continue resolving by original slide part and shape ID.

`ChartModel.remove()` removes the target graphic frame and its slide relationship, then garbage-collects the chart
root only when no incoming reference remains. Owned workbook/style/color dependencies are collected recursively;
shared images, themes, and external targets remain. Unsupported chart XML may still be deleted because deletion owns
the entire graphic-frame relationship root and does not infer internal chart semantics.

Removing the last chart does not renumber remaining shape, relationship, chart, workbook, series, or axis IDs.

## Advanced-Charts Plugin Integration

The plugin remains the home of trendline, error-bar, modern-chart, and explicit image-fallback APIs. It must reuse the
new strict chart state helpers for standard series resolution and workbook consistency instead of maintaining a
separate descendant-wide cache reader.

`setSeriesValues()` becomes a compatibility wrapper over the core semantic replacement path and therefore updates
both cache and workbook. Existing divergence diagnostics remain for raw XML edits or direct workbook replacement that
cannot prove equality. Adding a trendline, error bars, or data labels changes only the standard chart part and does not
alter data/workbook equality.

Modern charts remain inspectable and diagnostic-only. The plugin never converts them to standard charts implicitly.

## Diagnostics and Validation

The chart phase adds precise diagnostics:

- `CHART_RELATIONSHIP_INVALID`: missing, external, repeated, wrong-type, or dangling slide/chart/workbook relation;
- `CHART_STRUCTURE_UNSUPPORTED`: standard chart exists but its direct grammar cannot be safely projected;
- `CHART_STRUCTURE_AMBIGUOUS`: repeated roots, plot areas, chart groups, series ownership, or cache containers;
- `CHART_CACHE_INVALID`: point count, index, lexical value, or vector length is invalid;
- `CHART_AXIS_INVALID`: group axis references are missing, duplicated, dangling, or incompatible;
- `CHART_WORKBOOK_MISSING`: a semantically recognized chart has no embedded workbook;
- `CHART_WORKBOOK_CACHE_DIVERGENCE`: workbook cells and chart caches/formulas disagree;
- `MODERN_CHART_EXTENSION`: existing information diagnostic retained by the plugin.

Chart errors participate in `PptxDocument.write()` for every slide. Strict writes reject broken relationships,
invalid caches, dangling axes, and unsafe ambiguity. A cache-only imported chart is writable with a warning when its
display state is valid; semantic mutation materializes a synchronized workbook. Unknown chart options and extension
content are preserved and do not become errors merely because the native API does not expose them.

## Conformance and Verification

Completion evidence must cover:

- unit tests for descriptor-safe normalization, strict direct-state parsing, workbook cells/formulas, canonical XML,
  option defaults, malformed input, no-op bytes, and rollback at every mutation boundary;
- all nine public types plus valid combination groups, categorical/multi-level/numeric categories, scatter X/Y,
  bubble X/Y/size, explicit zero values, Unicode/XML escaping, and special number formats;
- create/edit/type-change/group-change/duplicate/move/delete/write/reopen in `pptx`, `pptm`, `potx`, `potm`, `ppsx`,
  and `ppsm`;
- PptxGenJS 4.0.1 public-output imports for each type and representative options, with native semantic projection and
  subsequent synchronized edits;
- advanced-plugin compatibility and exact modern-chart preservation;
- actual packed Node, real-Chrome, declaration, and installed-CLI smoke with a top-level `nativeCharts: true` gate;
- a real-data gallery created only through the installed tarball, with package inspection, strict reopen,
  PowerPoint-profile validation, 180-DPI rendering, overflow checks, and slide-by-slide visual review;
- PowerPoint and LibreOffice open/save/reopen evidence that records workbook/cache, series, option, rendering, and
  dependency normalization without claiming unsupported clients as successful.

Every implementation task ends with focused tests, diff review, one commit, push, and local/remote divergence `0 0`.
Temporary workbooks, fixtures, packed consumers, galleries, renders, and client round trips stay under
`/tmp/pptx-native-charts-*`.

## Implementation Sequence

1. strict chart types, definition normalization, and standard chart state reader;
2. deterministic workbook column planner and XLSX generator;
3. categorical and axis-free chart creation;
4. scatter and bubble chart creation;
5. combination groups and primary/secondary axes;
6. atomic semantic replacement, type conversion, and workbook/cache synchronization;
7. chart option creation and semantic replacement;
8. deletion lifecycle, diagnostics, validation, and advanced-plugin integration;
9. PptxGenJS conformance, packed/browser/client evidence, and public documentation.

The first implementation item deliberately ends at strict read-only semantic state. It creates no parts and changes
no existing chart bytes, making the public type and ownership grammar independently reviewable before workbook and
mutation code depend on it.
