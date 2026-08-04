# PptxGenJS Text Direction Surface Family Design

## Context

The PptxGenJS 4.0.1 declaration audit contains 38 unverified text-direction atoms. They form one capability family:

- six remaining owners inherit `textDirection`, each contributing the property plus four union tokens;
- `TextPropsOptions.vert` contributes the property plus seven union tokens.

Five `IChartOpts.textDirection` atoms are already closed by the inert-chart family and are outside this batch. Native already has strict table-cell and text-box direction models, creation, editing, rollback, duplicate, reopen, package, and partial browser evidence. The missing work is direct owner control plus compact six-format, packed-package, real-browser, and matrix evidence.

## Goals

- Close exactly the 38 remaining text-direction atoms in one capability-family batch.
- Lock the PptxGenJS runtime behavior for every remaining declared owner and legal token.
- Reuse existing deep lifecycle tests instead of duplicating them in every layer.
- Add compact aggregate evidence for six OOXML formats, the packed npm package, persistent Chrome, and package isolation.
- Regenerate the compatibility matrix from evidence only.

## Non-goals

- Do not reopen the five completed `IChartOpts.textDirection` atoms.
- Do not add ignored `textDirection` aliases to native text boxes, placeholders, slide numbers, or `tableToSlides`.
- Do not broaden native table directions to text-box-only or invalid tokens.
- Do not split this family into per-owner or per-token commits.

## Surface Inventory

The 38 target atoms are:

- `TableCellProps.textDirection`: property plus `horz | vert | vert270 | wordArtVert`.
- `TableProps.textDirection`: property plus the same four tokens.
- `PlaceholderProps.textDirection`, `SlideNumberProps.textDirection`, `TableToSlidesProps.textDirection`, and `TextPropsOptions.textDirection`: property plus the same four tokens for each owner.
- `TextPropsOptions.vert`: property plus `eaVert | horz | mongolianVert | vert | vert270 | wordArtVert | wordArtVertRtl`.

## Locked Upstream Behavior

One aggregate PptxGenJS control must prove:

- `TableCellProps.textDirection` and `TableProps.textDirection` are active.
- `vert`, `vert270`, and `wordArtVert` write direct `a:tcPr@vert` values.
- `horz` has canonical horizontal semantics and is omitted from table-cell OOXML.
- Table-level direction materializes to cells, while an explicit cell value overrides the table default.
- `TextPropsOptions.vert` writes all seven declared tokens exactly to `a:bodyPr@vert`.
- `PlaceholderProps.textDirection`, `SlideNumberProps.textDirection`, `TableToSlidesProps.textDirection`, and `TextPropsOptions.textDirection` accept all four declared tokens but do not affect owner-attributable output.
- An outer `vert` value wins when `TextPropsOptions.vert` and the ignored `textDirection` declaration are both present.
- Run-local `vert` is ignored because direction is a text-body property.
- PptxGenJS passes declaration-external truthy tokens through; native rejects them before mutation. This runtime defect is recorded in notes but creates no extra declaration atom.

## Native Contract

Native keeps two strict models:

```ts
type TableCellTextDirection = 'horz' | 'vert' | 'vert270' | 'wordArtVert';

type TextBoxTextDirection =
  | 'eaVert'
  | 'horz'
  | 'mongolianVert'
  | 'vert'
  | 'vert270'
  | 'wordArtVert'
  | 'wordArtVertRtl';
```

Table directions serialize on `a:tcPr@vert`; text-box directions serialize on `a:bodyPr@vert`. Creation may omit horizontal table direction because it is the canonical default, while live editing may retain an explicit horizontal token. Both paths preserve the same final semantics. Invalid values fail atomically.

## Classification

The exact expected batch totals are:

- `supported`: 18
- `deliberate-difference`: 0
- `defect-excluded`: 20

Supported:

- five `TableCellProps.textDirection` atoms;
- five `TableProps.textDirection` atoms;
- eight `TextPropsOptions.vert` atoms.

Defect-excluded:

- five atoms for each ignored `textDirection` owner: `PlaceholderProps`, `SlideNumberProps`, `TableToSlidesProps`, and `TextPropsOptions`.

## Evidence Architecture

- Adapter control: one aggregate owner/token fixture, plus reuse of existing table and text-direction comparison tests.
- Native SDK: reuse existing deep edit/rollback tests and add one compact six-format aggregate for all active tokens, exact OOXML, write, reopen, and diagnostics.
- Packed package: one `textDirectionFamilyState` that reuses the existing table state and adds all seven text-box tokens.
- Browser: the same stable family state in persistent Chrome, retaining console, page, and network error gates.
- OOXML: exact `bodyPr` versus `tcPr` attributes plus `pptx-inspect` validate/diff/part-read isolation.
- Matrix: generated entries from frozen owner/token lists with exact family and global count assertions.

Only the main agent edits repository files. Evidence agents remain read-only.

## Acceptance Criteria

- Exactly 38 matrix atoms change from `unverified`, and every changed ID is in this family.
- Batch totals equal 18 supported and 20 defect-excluded.
- Focused adapter and SDK family tests pass.
- TypeScript, surface-audit tests, npm tarball smoke, persistent Chrome, `pptx-inspect`, and the full repository suite pass.
- Review finds no unrelated staged path.
- One family commit is pushed and local/remote divergence is `0 0`.
