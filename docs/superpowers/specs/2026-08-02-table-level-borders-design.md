# Table-Level Borders Design

## Status

Validated under the standing direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. Product and implementation choices are delegated, so this design proceeds without a separate decision stop. It follows the completed table-level vertical-alignment, text-direction, horizontal-alignment, margins, and fill consensus/bulk editors.

## Goal

Add a lossless native `TableModel.borders` read/write surface that projects one uniform supported direct physical-cell border vector and atomically applies or clears direct `a:tcPr/a:lnL`, `a:lnR`, `a:lnT`, and `a:lnB` state across every physical cell in an existing table.

## Current context

The library already supports:

- strict descriptor-safe `TableCellBorderInput` scalar, exact `[top, right, bottom, left]` tuple, partial named-object, empty-object, and `undefined` input;
- direct none or strict sRGB/theme line values with finite `0..1584` point width and omitted/solid/dash style;
- canonical point-to-EMU quantization, with zero-width line distinct from no-fill and omitted style distinct from explicit solid;
- table-level border creation materialized into physical cells, with whole-value cell overrides and canonical direct no-fill for unspecified creation sides;
- detached partial `TableCell.borders` direct-state snapshots;
- `TableModel.setCellBorders()` whole-replacement editing and clearing;
- exact single-cell no-ops, unique advanced-side replacement, ambiguous-state protection, duplicate isolation, rollback, write/reopen, and unrelated-state preservation;
- PptxGenJS 4.0.1 table/cell border creation conformance for supported final direct states.

DrawingML has no table-wide stored border property. PptxGenJS's table-level `border` option is a creation value materialized into four direct line elements on every physical `a:tcPr`; only that final cell state survives write/reopen. The remaining gap is therefore a strict uniform projection and bulk editor over recoverable direct cell sides.

## Considered approaches

### 1. Partial-vector consensus property plus atomic all-cell setter — selected

Expose live `TableModel.borders` accessors. The getter returns a detached `TableCellBorders` object only when every physical cell has the same non-empty, structurally safe set of direct supported sides and values. Assignment accepts the existing `TableCellBorderInput` forms, normalizes once before package access, and whole-replaces all four owned sides on every physical cell in one transaction.

This reuses the public value model, reflects only serialized OOXML, preserves partial direct state, and gives existing-deck editing the same semantics already available for one cell.

### 2. Require all four direct sides in the getter

Requiring a complete TRBL vector would simplify comparison and matches normal table creation output. Existing decks and the current cell editor can legally contain only some direct sides. Rejecting uniform partial state would lose recoverable information and make table projection inconsistent with `TableCell.borders`, so this approach is rejected.

### 3. Resolve effective or shared-edge borders

Synthesizing missing sides from table styles, creation input, neighboring cells, merged-cell topology, or shared-edge precedence could return a visual border for more decks. Those sources are not one recoverable direct table property, and the result would require layout/style evaluation outside this slice. This approach is rejected.

No generic higher-order table-consensus abstraction is added. The four-side safety and equality rules are property-specific, and the existing focused helper is the smallest reliable change.

## Public API

`TableModel` gains one live property with separate getter and setter types, matching the established `margins` pattern:

```ts
class TableModel {
  get borders(): TableCellBorders | undefined;
  set borders(value: TableCellBorderInput | undefined);
}
```

Usage:

```ts
const table = slide.addTable([
  ['North', '$1.2M'],
  ['South', '$980K'],
], {
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 1,
    style: 'solid',
  },
});

table.borders; // one uniform four-side direct snapshot
table.setCellBorders(0, 1, { kind: 'none' });
table.borders; // undefined: physical cells are mixed
table.borders = {
  top: {
    kind: 'line',
    color: { kind: 'srgb', value: '4472C4' },
    width: 2,
    style: 'dash',
  },
  bottom: { kind: 'none' },
}; // clears left/right on every physical cell
table.borders = { kind: 'none' }; // writes direct none on all four sides
table.borders = undefined; // clears all four direct sides
```

No singular alias, retained table-default object, mixed sentinel, effective-value API, or PptxGenJS-shaped compatibility facade is added.

## Getter semantics

The getter resolves the exact direct table path and enumerates all physical `a:tc` elements, including merge continuations.

- One or more cells with the same non-empty safe direct side set and values returns one detached `TableCellBorders` snapshot.
- Uniform complete and uniform partial vectors are both valid.
- Each side preserves direct `{ kind: 'none' }` versus a direct line.
- Line equality includes color kind/value, quantized width, and style presence/value. Omitted style is not equal to explicit `solid`.
- Numerically equivalent width tokens compare after strict EMU parsing, and sRGB values compare after uppercase canonicalization.
- All four sides absent returns `undefined`.
- Present/absent mixtures, differing side sets, differing values, or four direct none versus four absent return `undefined`.
- Empty tables, ambiguous direct table paths, missing/repeated direct `tcPr`, repeated owned side elements, malformed lines, unsupported direct line state, and other unsafe state return `undefined` without throwing during a read.
- Wrong-prefix children, diagonals, descendants, border-like text lines, and elements outside the selected direct `tcPr` do not participate.

The existing `TableCell.borders` reader remains unchanged and continues its per-side tolerant snapshot behavior. The table-level getter uses a stricter private cell-state read so malformed or unsupported owned sides cannot be mistaken for valid absence. Callers needing mixed or per-cell detail inspect `rows[].cells[].borders`.

## Setter semantics

Assignment calls `normalizeTableCellBorders(value, 'Table borders')` before package access. Accepted inputs are:

- one direct border scalar for all four sides;
- one dense ordinary four-value `[top, right, bottom, left]` tuple whose items may be `undefined`;
- one ordinary/null-prototype partial named object;
- `{}`, an all-`undefined` tuple, or `undefined` to clear all four direct sides.

The existing normalizer continues to reject null, class instances, accessors, inherited or symbol properties, extra keys, sparse or wrong-length tuples, missing/invalid kinds, colors, widths or styles, non-finite widths, and values outside `0..1584` before mutation. Normalized nested colors and side objects are detached immediately.

For a valid assignment:

1. resolve one live table snapshot inside one OPC transaction;
2. require one complete non-empty direct physical-cell set;
3. call the existing `replaceTableCellBorders()` for every physical cell on the same in-memory slide XML document;
4. serialize and commit the slide at most once, and only if at least one cell changed.

Scalar input writes all four sides. Tuple and named input use the existing editor's whole-replacement semantics: supplied sides are written and omitted sides are removed from every cell. Empty input and `undefined` clear all four sides. No clear restores a creation default, table style, or neighboring edge.

Assigning the same semantic uniform vector is an exact bytes and mutation-journal no-op. Clearing an all-absent table is also an exact no-op. Unique malformed or advanced side elements may be normalized by legal replacement or removed by clear, matching the cell editor. Missing/repeated direct `tcPr` or repeated same-prefix owned side elements make the edit unsafe and throw `ModelParseError`. Package bytes are committed only after every physical cell succeeds, so an unsafe later cell cannot leave a partial package mutation.

## OOXML ownership and preservation

The editor owns only exact same-prefix direct `lnL`, `lnR`, `lnT`, and `lnB` children under each physical cell's unique direct `tcPr`. The public getter recognizes only strict direct none and strict solid-color lines with the supported optional dash state. Replacement or clear may intentionally replace/remove one unique unsupported advanced side, matching `setCellBorders()` whole-replacement semantics.

The property never reads or writes `lnTlToBr`, `lnBlToTr`, text outlines, table styles, effective theme resolution, adjacent shared edges, merge topology, extensions, or descendants outside the direct cell-properties owner. Bulk editing preserves text, runs, paragraph properties, horizontal/vertical alignment, margins, text direction, fit, fill, row/grid geometry, transforms, merge-continuation identity, diagonal lines, extensions, relationships, and unrelated shapes or package parts. New sides retain the existing L/R/T/B schema order before direct cell fill and `extLst`.

The property is live against current OOXML and preserves `TableModel` identity. Slide duplication copies current direct cell state; later source and duplicate edits remain isolated. Move, write, reopen, and all six presentation formats preserve the projection. User transactions roll back the entire bulk edit and mutation journal. Validation failure or unsafe-state edits preserve bytes, relationships, diagnostics, model collections, and journal state.

## PptxGenJS conformance boundary

PptxGenJS 4.0.1 has table/cell creation options but no live table model or existing-deck editor. Conformance covers legal supported final direct output:

- omitted border imports as the explicit uniform four-side no-fill state that PptxGenJS materializes;
- a uniform legal table scalar or TRBL border imports as one native `TableCellBorders` consensus;
- a legal differing cell override produces mixed direct state and projects `undefined` while per-cell snapshots remain available;
- native assignment changes only final direct cell L/R/T/B elements and survives write/reopen.

Existing strict differences remain documented: PptxGenJS empty objects materialize default gray lines; omitted line type materializes explicit solid; TRBL zero-width items may become 1pt; short runtime tuples are accepted and padded. Native empty input means clear/no layer, omitted style stays omitted, zero width stays zero, tuples are exact and dense, and named sides/theme colors are supported. Native bulk editing is an intentional existing-deck extension.

## Test design

### Internal and model tests

- Add focused internal tests for uniform complete and partial vectors, none, sRGB/theme line, omitted/solid/dash style, zero/fractional/maximum width, lexical width and color equivalence, all absent, present/absent and none/line mixtures, differing side sets/values, malformed and advanced sides, repeated sides, missing/repeated `tcPr`, empty rows/table, ambiguous table paths, merge continuations, diagonals, and qualified/descendant lookalikes.
- Apply scalar, exact TRBL, partial named object, all-undefined tuple, empty object, and `undefined` across multi-row/multi-column tables.
- Require partial whole-replacement clearing, exact same-value and all-absent-clear no-ops, one committed slide mutation for a real change, late-cell atomicity, unique advanced-side replacement/clear, and preservation of unrelated XML.
- Require invalid-input early rejection, transaction rollback, stable model identity, duplicate isolation, move, all six formats, write, and reopen.
- Require public typing to read `TableCellBorders | undefined`, accept `TableCellBorderInput | undefined`, and reject unsupported values.
- Re-run completed table-level fill, margins, horizontal alignment, text direction, and vertical alignment tests.

### PptxGenJS evidence

- Import real PptxGenJS 4.0.1 omitted borders and require explicit uniform four-side none consensus.
- Import legal uniform scalar and TRBL table borders and require equivalent native direct snapshots.
- Import one legal cell override and require mixed `undefined` plus the exact per-cell matrix.
- Normalize an imported mixed table with the native setter, write/reopen, and require every physical direct side vector to match.
- Preserve the documented empty/default-style/TRBL-zero runtime differences.

### Package and browser evidence

- Require generated declarations to contain both accessors inside the packed `TableModel` block.
- Extend actual-tarball Node, TypeScript, browser-condition, installed CLI package inspection, and real-Chrome smoke with stable `tableBorders: true` and detailed state.
- Cover uniform read, detached read, no-op isolation, mixed projection, scalar/partial/none overwrite, clear, reopen, and invalid failure isolation.
- Inspect packed PPTX slide XML and require the expected direct L/R/T/B state per physical cell without diagonal, text-line, or fill false positives.
- Use `pptx-inspect` broad package inspection and PowerPoint 2010 validation before exact slide-part reads; keep validation, console, page, and network error counts at zero.

## Delivery boundaries

The work is delivered as separately reviewed and pushed commits:

1. this design;
2. the implementation plan;
3. core helper/property implementation, source tests, and PptxGenJS conformance;
4. actual-package declarations, Node/browser/CLI, and real-Chrome evidence;
5. public documentation, compatibility matrix, changelog, and progress closeout.

Each commit stages only its declared files. `.pnpm-store/`, generated workspace tarballs, and retained `/tmp` evidence are never committed.

## Non-goals

- Persisting or reconstructing a table creation default.
- Filling absent sides from canonical defaults, table styles, neighboring cells, merge topology, or shared-edge precedence.
- Returning a mixed sentinel or per-cell matrix from the table-level getter.
- Adding diagonal, transparency, gradient/pattern line fill, advanced dash, compound/cap/join/alignment/arrow/effect state to the public border model.
- Changing `TableCell.borders`, `setCellBorders()`, table creation precedence/output, zero-width, omitted-style, or explicit-none semantics.
- Adding table-level fit, merge editing, row/column insertion/deletion, auto-page, repeated headers, or `tableToSlides`.
- Performing content measurement or layout/geometry recomputation.
- Generalizing table properties behind a higher-order consensus framework.
- Copying PptxGenJS invalid-value coercion or default materialization into native APIs.

## Acceptance criteria

The item is complete when `TableModel.borders` safely projects one uniform non-empty supported direct complete or partial physical-cell border vector, atomically whole-replaces or clears it across every physical cell, preserves absence and explicit direct intent, preserves unrelated bytes and lifecycle behavior, rejects invalid or unsafe edits without package mutation, imports legal PptxGenJS table output within the documented direct-state boundary, keeps completed table-level properties unchanged, passes focused/full/performance/type/bundle/declaration gates, the actual tarball and real Chrome report `tableBorders: true` with zero browser errors, documentation moves the table-level border getter/editor to supported, and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
