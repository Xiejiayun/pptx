# PptxGenJS Chart Axes Surface Family Design

Date: 2026-08-05

Status: confirmed under the user's autonomous-decision authorization

## Context and Batch Boundary

The declaration audit has 155 unverified atoms in the complete chart-axis
family. The boundary includes the top-level and nested category/value axis
owners, series-axis properties, all declared crosses and display-unit union
members, the multi-axis containers, the four format/grid aliases whose IDs do
not contain `Axis`, and eleven PptxGenJS 4.0.1 runtime defects.

This is one capability-family batch. It does not split already implemented
axis foundations from date, display-unit, frequency, multi-level, secondary,
or series-axis behavior. Chart data labels, series presentation, plot/chart
areas, 3-D view, grouping, markers, and chart type catalogs remain separate
families.

## Native Model

Replace the permissive shared axis surface with three public option types:

- `ChartCategoryAxisOptions` owns category/date scaling, crossing,
  label-frequency, multi-level-label, and time-unit semantics;
- `ChartValueAxisOptions` owns numeric scaling, logarithmic base, crossing,
  built-in display units, and display-unit-label visibility;
- `ChartSeriesAxisOptions` owns editable 3-D series-axis visibility, labels,
  line/grid, units, title, orientation, and frequency.

Common font, title, position, orientation, tick, line, gridline, and number
format fields remain structurally shared but are normalized through
axis-specific allowlists. This prevents category time units from appearing on
value/series axes, logarithmic bases from appearing on category/series axes,
and display units from appearing outside value axes.

For scatter and bubble charts, the existing `categoryAxis` slot continues to
represent the horizontal `c:valAx` for PptxGenJS compatibility, but it accepts
only the common/category-declared controls in this batch. A dedicated
horizontal numeric-axis surface is outside the 155-atom PptxGenJS declaration
boundary and is not inferred by weakening the strict category interface.

Category axes use `kind: 'category' | 'date'`. Supplying a time unit infers a
date axis; explicitly requesting a category axis together with time units is
rejected before mutation. Date units are the closed `days | months | years`
domain. Label frequencies are positive integers. Crossings accept finite
numbers or `autoZero`. Value display units use the nine declared built-in
tokens, and a display-unit label is rejected unless a display unit is present.

`ChartOptions` exposes primary and secondary category/value objects plus the
single series axis used by a 3-D bar axis set. Series-axis objects are serialized only for 3-D bar axis sets; they
remain inert nowhere and invalid combinations fail normalization or semantic
rendering before the package changes.

## OOXML and Round Trip

The renderer emits `c:catAx`, `c:dateAx`, `c:valAx`, and `c:serAx` with strict
schema ordering. Numeric crossings use `c:crossesAt`; the canonical default
uses `c:crosses val="autoZero"`. Date axes emit base/major/minor time units,
category and series frequencies emit `c:tickLblSkip`, category multi-level
intent emits `c:noMultiLvlLbl`, and value display units emit `c:dispUnits`
with an optional `c:dispUnitsLbl`.

The state reader resolves the series member of each three-axis set and reads
axis-kind-specific options. Default values remain canonical omissions after
reopen, while non-default and numeric-zero values remain observable. Create,
read, edit, clear, duplicate, transaction rollback, write, and reopen all use
the same detached frozen definition model.

## Compatibility Classification

Seventy-nine atoms are supported directly. Sixty-five are deliberate API
shape differences where native uses named primary/secondary axis objects,
structured colors/lines/titles, strict numeric frequencies, closed time-unit
tokens, or canonical format fields instead of PptxGenJS's flat permissive
properties and aliases.

Eleven atoms are defect-excluded from the PptxGenJS 4.0.1 parity denominator:

- `axisPos` and its four tokens are inert and produce the baseline axis XML;
- the three series time-unit properties reject all three values because the
  runtime validates the property name instead of the property value;
- series label positions `high`, `nextTo`, and `none` all serialize as `low`.

The valid series label-position property and `low` token remain supported.
Numeric crossing zero is supported natively even though one PptxGenJS writer
path produces malformed `crossesAt="autoZero"`; the declared numeric union is
not excluded as a whole.

## Evidence and Pipeline

Agent A owns the exact 155-atom matrix slice and status generation. Agent B
owns one PptxGenJS 4.0.1 runtime/control probe. Agent C owns native, actual npm
tarball, persistent real-Chrome, OOXML, and PowerPoint 2010 evidence. Evidence
agents write only `/tmp` fragments; the main agent is the only shared-file
writer and owns implementation, review, artifact regeneration, commit, push,
and progress reporting.

Evidence is aggregated once for the family:

- focused model, adapter, and SDK tests run once;
- all six presentation formats share one native create/edit/duplicate/reopen
  aggregate;
- the actual npm tarball is built and installed once;
- a single persistent Chrome state covers the complete advanced-axis slice;
- OOXML/package validation and the full test suite run once at the batch gate.

## Acceptance Criteria

- Exactly 155 atoms leave `unverified`: 79 become `supported`, 65 become
  `deliberate-difference`, and 11 become `defect-excluded`.
- Totals become 716 supported, 431 deliberate differences, 94 deprecated
  aliases, 370 defect exclusions, and 163 unverified: 1,611 of 1,774 atoms
  classified (90.81%).
- Native supports strict category/date, value, and editable series axes through
  create, read, edit, clear, duplicate, rollback, write, and reopen.
- Manifest scope/count assertions, focused tests, deterministic artifact
  generation, TypeScript, actual npm tarball, real Chrome, OOXML, PowerPoint
  2010, and the batch full suite pass.
- Review produces one capability-family commit and push with no unrelated
  tracked files.
