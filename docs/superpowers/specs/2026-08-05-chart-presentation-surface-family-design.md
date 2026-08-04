# PptxGenJS Chart Presentation Surface Family Design

Date: 2026-08-05

Status: confirmed under the user's autonomous-decision authorization

## Scope and Batch Boundary

Close exactly 91 unverified PptxGenJS 4.0.1 atoms in one chart-presentation
capability family. The batch contains 88 already implemented presentation
atoms plus three importer gaps that share the same hierarchical data-label
reader:

- `interface:IChartOpts@property:showLabel`;
- `interface:IChartOpts@property:showPercent`;
- `union:interface:IChartOpts@property:dataLabelPosition#bestFit`.

The other presentation atoms cover chart geometry and identity, bar grouping,
gaps and overlap, colors, group data labels, data tables, blank handling, pie
and doughnut geometry, legends, line and marker presentation, radar styles,
titles, and 3-D view settings. Point fills, point lines/effects, scatter custom
label text/fields, line caps, `bubble3D`, and multi-level category data remain
separate capability families because they need new public point or residual
models rather than evidence-only closure.

The exact source of truth is `/tmp/chart-presentation-91-matrix-map.json`: 91
unique atoms, 26 `supported`, 65 `deliberate-difference`, and no defect
exclusions.

## Hierarchical Data-Label Readback

Keep `ChartCommonGroupOptions.dataLabels` as the public group-level semantic
surface. Imported PptxGenJS pie and doughnut files often put effective label
state under `c:ser/c:dLbls`, with a default series layer and a complete set of
per-point `c:dLbl` overrides. The reader resolves those layers without adding
a lossy per-point API in this batch.

For every chart group:

1. Read direct group `c:dLbls` as the fallback.
2. Inspect every direct series `c:dLbls`.
3. A series layer with no point labels is promotable when its label shape and
   extension payload are absent or empty and its supported scalar state can be
   represented by `ChartDataLabelOptions`.
4. A series layer with point labels is promotable only when point indices are
   unique, dense, and cover the complete series value count; every point is
   free of custom `c:tx`, fields, extensions, and non-empty label shape state;
   and all effective point options are semantically identical.
5. Point booleans explicitly set to false override true series defaults.
   `dLblPos`, `showCatName`, and `showPercent` therefore preserve the effective
   PptxGenJS pie semantics, including `bestFit` and false values.
6. Promote the series result only when every series has one promotable layer
   and all promoted results are equal. Otherwise keep the direct group result
   and leave the unsupported narrow layer unread.

Custom scatter labels, `customXY` fields, non-empty label backgrounds,
point-specific state, unknown extensions, and partial point coverage are never
promoted. Their XML remains owned by the imported package, not by the group
model.

## Safe Editing and Losslessness

Unrelated chart edits must not rewrite direct group or series data-label XML.
The editor compares the current and requested group `dataLabels` values for
each structurally unchanged group:

- when label options are unchanged, `dLbls` is excluded from the owned patch
  set, preserving imported point labels byte-for-byte;
- when label options change, the direct group layer is synchronized with the
  canonical renderer;
- only series layers accepted by the same safe-promotion predicate are
  removed, so the newly edited group layer becomes effective;
- unsafe custom or point-style series layers remain untouched.

Clearing a promoted label object removes both its canonical group layer and
the safely promotable source series layers. A failed or unsafe replacement
still runs inside the existing package transaction and must leave package
bytes, relationships, cached model state, and the mutation journal unchanged.

## Compatibility Classification

Twenty-six atoms are supported with the same effective semantics. Sixty-five
are deliberate API-shape differences: Native uses strict nested chart,
group, series, marker, data-label, legend, data-table, title, and 3-D objects
instead of permissive PptxGenJS top-level aliases or abbreviated tokens.

`showPercent` is supported after hierarchical readback. `showLabel` is a
deliberate difference because Native names the semantic
`showCategoryName`. `bestFit` is a deliberate difference because Native
promotes the uniform effective point value into group
`ChartDataLabelOptions.position` and emits canonical group XML on explicit
edit.

## Evidence Pipeline

Agent A owns the immutable 91-ID matrix and importer review. Agent B owns the
PptxGenJS runtime/control mapping. Agent C owns actual packed-package,
persistent Chrome, OOXML, and PowerPoint 2010 evidence. Evidence agents write
only `/tmp`; the main agent owns every repository edit, review, commit, and
push.

One aggregate per gate covers the family:

- focused model/editor/adapter/SDK tests once;
- one actual npm tarball build/install smoke;
- one persistent real-Chrome session;
- one OOXML/package and PowerPoint 2010 validation batch;
- one TypeScript gate and one full suite;
- two deterministic audit generations.

## Acceptance Criteria

- Exactly 91 atoms leave `unverified`: 26 become `supported` and 65 become
  `deliberate-difference`.
- Totals become 742 supported, 496 deliberate differences, 94 deprecated
  aliases, 370 defect exclusions, and 72 unverified: 1,702 of 1,774 atoms
  classified (95.94%).
- Actual PptxGenJS pie/doughnut `showLabel`, `showPercent`, and `bestFit`
  output reads back with its effective point semantics.
- Explicit label edit/clear canonicalizes only safe promoted layers; unrelated
  edits preserve custom scatter text/fields, point styles, and extensions.
- Focused tests, TypeScript, deterministic matrix generation, actual npm
  tarball, persistent Chrome, OOXML, PowerPoint 2010, and the full suite pass.
- Review produces one capability-family commit and push with no unrelated
  tracked files.
