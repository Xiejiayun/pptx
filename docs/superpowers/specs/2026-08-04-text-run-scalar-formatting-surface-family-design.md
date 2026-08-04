# PptxGenJS Text Run Scalar Formatting Surface Family Design

## Context

The PptxGenJS 4.0.1 declaration audit has 54 unverified scalar text-formatting atoms. They are the Cartesian product of six inherited owners and nine properties:

- owners: `PlaceholderProps`, `SlideNumberProps`, `TableCellProps`, `TableProps`, `TableToSlidesProps`, and `TextPropsOptions`;
- properties: `bold`, `breakLine`, `color`, `fontFace`, `fontSize`, `highlight`, `italic`, `lang`, and `softBreakBefore`.

The native library already has rich-text runs, table text defaults, slide-number styles, editable placeholders, strict semantic colors, line-break normalization, package clients, and browser coverage. The remaining work is to lock the exact PptxGenJS owner matrix, add one compact all-format native aggregate, expose a family-level package/browser result, and regenerate the evidence-backed matrix.

## Batch Boundary

This batch contains exactly the 54 atoms above. It excludes adjacent paragraph, geometry, and effect fields such as `align`, `margin`, `valign`, `transparency`, `underline`, `tabStops`, `textDirection`, `bullet`, `strike`, `glow`, `outline`, `baseline`, `fit`, `wrap`, and `rtlMode`. The already closed `IChartOpts` inherited text fields are not reopened.

This boundary is intentionally larger than a single-property batch but smaller than all remaining `TextBaseProps`: the nine fields share the same run-property writer, owner matrix, native models, and client lifecycle without mixing paragraph or shape-body semantics.

## Locked PptxGenJS Runtime Matrix

`S` means supported, `D` means deliberate difference, and `X` means inert upstream declaration excluded as a defect.

| Owner | bold | breakLine | color | fontFace | fontSize | highlight | italic | lang | softBreakBefore |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| PlaceholderProps | S | X | D | D | S | D | S | S | X |
| SlideNumberProps | S | X | D | D | S | X | X | X | X |
| TableCellProps | D | S | D | D | S | D | S | S | S |
| TableProps | S | X | D | D | S | X | X | X | X |
| TableToSlidesProps | X | X | X | X | X | X | X | X | X |
| TextPropsOptions | S | S | D | D | S | D | S | S | S |

The direct runtime control must prove:

- `bold`, `color`, `fontFace`, and `fontSize` are active for every owner except `TableToSlidesProps`;
- `highlight`, `italic`, and `lang` are active only for text, placeholder, and table-cell writers;
- `breakLine` is active only for text and table-cell rich runs and materializes paragraph boundaries;
- `softBreakBefore` is active only before a non-first text or table-cell run and materializes `a:br`;
- slide numbers emit `bold` on `a:rPr` and color/font/size in `a:defRPr`, while ignoring the other five fields;
- table defaults propagate only bold/color/font/size to cells;
- all nine top-level `TableToSlidesProps` fields are byte-for-byte inert, while a computed-CSS positive control proves the importer can produce styled cell runs;
- PptxGenJS truthy fallback prevents a legal cell `bold: false` from clearing inherited table `bold: true`;
- standalone or inherited color/highlight strings, font names, and valid point sizes/languages serialize to exact run properties;
- invalid or declaration-external inputs and caller mutation are recorded as runtime differences but do not broaden native contracts.

## Native Contract

Native maps the effective capabilities as follows:

- `color` and `highlight` use strict `RichTextColor` values instead of permissive PptxGenJS strings;
- `fontFace` maps to the consistently named native `fontFamily` field;
- `fontSize`, `bold`, `italic`, and `lang` remain scalar run-style fields;
- `breakLine` and `softBreakBefore` use `RichTextRun` and canonical paragraph/`a:br` OOXML;
- table defaults use `AddTableOptions` and `AddTableCellOptions`, with explicit rich-text runs for per-run styles;
- slide numbers use `SlideNumberOptions.style` for the four upstream-effective fields;
- placeholders accept rich text through `addPlaceholder`, including create, edit, duplicate, and reopen lifecycles.

Native preserves legal false values, rejects invalid values before mutation, does not mutate caller objects, and supports semantic sRGB/scheme colors. It does not add inert flat aliases merely because they appear through PptxGenJS declaration inheritance.

## Classification

The exact family totals are:

- `supported`: 19
- `deliberate-difference`: 14
- `defect-excluded`: 21

The fourteen deliberate differences are:

- five active `color` owners because native uses `RichTextColor`;
- five active `fontFace` owners because native uses `fontFamily` and rejects invalid XML strings;
- three active `highlight` owners because native uses `RichTextColor` and does not copy PptxGenJS's conditional-emission defect;
- `TableCellProps.bold` because native preserves legal `false` while PptxGenJS truthy fallback can overwrite it with the table default.

The twenty-one defect exclusions are exactly the inert cells in the runtime matrix. All other active cells are supported.

After this batch, expected global totals are 571 supported, 192 deliberate differences, 83 deprecated aliases, 333 defect exclusions, and 595 unverified atoms: 1,179 of 1,774 closed (66.46%).

## Evidence Architecture

- Adapter: one aggregate test named `locks scalar text formatting behavior across every declared owner`, reusing existing focused controls for line breaks, slide numbers, table defaults, and invalid truthiness.
- SDK: one aggregate test named `creates and reopens scalar text formatting owners in all six formats`, reusing existing deep edit, rollback, duplicate, and owner tests.
- npm package: one `textRunScalarFamilyState` that combines existing slide-number, table-default, and rich-line-break gates with a small missing run-style fixture.
- Browser: the same stable family state in one persistent real-Chrome run, with console, page, and network error gates.
- OOXML: exact `a:rPr`, `a:defRPr`, `a:highlight`, `a:latin/a:ea/a:cs`, paragraph, and `a:br` checks plus `pptx-inspect` validation and part-isolation diff.
- Matrix: entries generated from frozen owner/property classification tables, with exact family/global count assertions.

Only the main agent edits repository files. Evidence agents remain read-only, and no two agents modify the same file.

## Acceptance Criteria

- Exactly 54 matrix atoms change from `unverified`, with totals 19 supported, 14 deliberate differences, and 21 defect exclusions.
- The aggregate PptxGenJS control proves every owner/property cell and its positive control.
- Native creation and reopen pass in all six formats with exact OOXML and zero error/warning diagnostics.
- The packed npm tarball and persistent Chrome report `textRunScalarFamily=true`; all family-state fields are true and Chrome error counts are zero.
- `pptx-inspect` validates created and edited files and shows only the intended slide part changed.
- Focused tests, TypeScript, audit tests, syntax, full Vitest, and generated-artifact checks pass.
- Review finds no unrelated staged path; one capability-family commit is pushed and local/remote divergence is `0 0`.
