# Table-Level Fill Design

## Status

Validated under the standing direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. The user delegated product and implementation choices, so this design proceeds without a separate decision gate. It follows the completed table-level vertical-alignment, text-direction, horizontal-alignment, and margins consensus/bulk editors.

## Goal

Add a lossless native `TableModel.fill` read/write surface that projects one uniform supported direct physical-cell fill and atomically applies or clears the direct `a:tcPr` fill choice across every physical cell in an existing table.

## Current context

The library already supports:

- strict descriptor-safe `TableCellFill` input for direct none or solid sRGB/theme fill with optional `0..100` transparency;
- canonical transparency quantization to `0.001%`, with omitted alpha distinct from explicit zero transparency;
- table-level fill creation materialized into physical cells, with whole-value cell overrides;
- `TableCell.fill` detached direct-state snapshots;
- `TableModel.setCellFill()` whole-replacement editing and clearing;
- exact single-cell no-ops, malformed single-choice replacement, ambiguous-state protection, duplicate isolation, rollback, write/reopen, and unrelated-state preservation;
- PptxGenJS 4.0.1 table/cell fill creation conformance for supported final direct states.

DrawingML has no table-wide stored fill property. PptxGenJS's table-level `fill` option is a creation value materialized into direct fill choices on physical `a:tcPr` elements; only that final cell state survives write/reopen. The remaining gap is therefore a uniform projection and bulk editor over recoverable supported direct cell fills.

## Considered approaches

### 1. Strict direct-fill consensus property plus atomic all-cell setter — selected

Expose live `TableModel.fill` accessors. The getter returns a detached `TableCellFill` only when every physical cell has the same structurally valid direct none or solid fill. Assignment normalizes once before package access and whole-replaces every physical cell's direct fill choice in one transaction.

This reuses the established public value model and cell codec, reflects only serialized OOXML, preserves explicit none and explicit zero transparency, and adds the smallest coherent existing-deck editing surface.

### 2. Return a mixed sentinel or per-cell matrix

A `mixed` value or table-level matrix would distinguish absence, unsupported state, and heterogeneity. It would add a second result model even though per-cell detail already exists in `rows[].cells[].fill`, and it would diverge from the completed table-level consensus properties. It is rejected.

### 3. Resolve effective table-style fill or retain creation metadata

Computing an effective fill from table styles, defaults, or original creation input could produce a value when direct cells are absent. Those sources are not a single direct table property, and creation metadata does not survive reopen. Synthesizing them would make imported and created decks behave differently and make clear operations unexpectedly re-inherit. It is rejected.

## Public API

`TableModel` gains one live property:

```ts
class TableModel {
  get fill(): TableCellFill | undefined;
  set fill(value: TableCellFill | undefined);
}
```

Usage:

```ts
const table = slide.addTable([
  ['North', '$1.2M'],
  ['South', '$980K'],
], {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
});

table.fill; // uniform direct solid fill
table.setCellFill(0, 1, { kind: 'none' });
table.fill; // undefined: physical cells are mixed
table.fill = { kind: 'none' }; // writes direct noFill to every physical cell
table.fill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'D9EAF7' },
  transparency: 0,
};
table.fill = undefined; // removes every supported direct fill choice
```

No `setFill()` alias, mixed sentinel, effective-value API, retained table-default object, or PptxGenJS-shaped compatibility facade is added.

## Getter semantics

The getter resolves the exact direct table path and enumerates every physical `a:tc`, including merge continuations.

- One or more cells with the same safe direct `{ kind: 'none' }` returns a detached none snapshot.
- One or more cells with the same safe direct solid color kind/value and the same optional transparency returns a detached solid snapshot.
- sRGB values compare after canonical uppercase parsing; alpha compares after strict numeric parsing into the public transparency value.
- Omitted alpha and explicit zero transparency remain different direct states.
- All cells with no supported direct fill return `undefined`.
- Present/absent mixtures, none/solid mixtures, differing colors, or differing transparency return `undefined`.
- A single malformed or unsupported direct fill in any cell makes the table projection `undefined`; no approximation of gradient, pattern, picture, or group fill is returned.
- Empty tables, ambiguous direct table paths, missing/repeated direct `tcPr`, or multiple direct same-prefix fill choices return `undefined` without throwing during a read.
- Qualified/wrong-prefix children and descendant fill lookalikes outside the selected direct `tcPr` do not participate.

The getter composes the existing strict `readTableCellFill()` with semantic `simpleFillsEqual()` comparison. `TableCell.fill` remains unchanged. Callers needing mixed or per-cell detail inspect `rows[].cells[].fill`.

## Setter semantics

Assignment calls `normalizeTableCellFill(value, 'Table fill')` before package access. Accepted inputs are:

- `{ kind: 'none' }`;
- `{ kind: 'solid', color }`;
- `{ kind: 'solid', color, transparency }`;
- `undefined` to clear the direct fill choice.

The existing normalizer continues to reject null, arrays, class instances, accessors, inherited or symbol properties, extra keys, missing/invalid kinds or colors, non-finite transparency, and values outside `0..100` before mutation. Normalized nested color state is detached immediately.

For a valid assignment:

1. resolve one live table snapshot inside one OPC transaction;
2. require one complete non-empty direct physical-cell set;
3. call the existing `replaceTableCellFill()` for every physical cell on the same in-memory slide XML document;
4. serialize and commit the slide at most once, and only if at least one cell changed.

`none` writes direct same-prefix `noFill`. Solid input writes one direct same-prefix `solidFill`; omitted transparency emits no alpha, while explicit zero emits `alpha=100000`. `undefined` removes the owned direct fill choice from every cell. No clear operation restores a creation default or table-style result.

Assigning the same semantic direct value to every cell is an exact bytes and mutation-journal no-op. Clearing an all-absent table is also an exact no-op. A single malformed supported fill choice or one advanced fill choice can be normalized by legal replacement or removed by clear, matching the existing cell editor. Missing/repeated direct `tcPr` or multiple direct same-prefix fill choices make the edit unsafe and throw `ModelParseError`. Package bytes are committed only after every physical cell succeeds, so an unsafe later cell cannot leave a partial package mutation.

## OOXML ownership and preservation

The editor owns only the exact same-prefix direct fill choice among `noFill`, `solidFill`, `gradFill`, `blipFill`, `pattFill`, and `grpFill` under each physical cell's unique direct `tcPr`. The public getter recognizes only strict direct none and solid values; replacement or clear may intentionally replace/remove one unique unsupported advanced choice, matching `setCellFill()` whole-replacement semantics.

The property never reads or writes border fills, text fills, table styles, effective theme resolution, extensions, or descendants outside the direct cell-properties owner. Bulk editing preserves text, runs, paragraph properties, horizontal/vertical alignment, margins, text direction, fit, borders, row/grid geometry, transforms, merge-continuation identity, extensions, relationships, and unrelated shapes or package parts. New fill XML remains before direct `extLst` through the existing cell replacer.

The property is live against current OOXML and preserves `TableModel` identity. Slide duplication copies current direct cell state; later source and duplicate edits remain isolated. Move, write, reopen, and all six presentation formats preserve the projection. User transactions roll back the entire bulk edit and mutation journal. Validation failure or unsafe-state edits preserve bytes, relationships, diagnostics, model collections, and journal state.

## PptxGenJS conformance boundary

PptxGenJS 4.0.1 has table/cell creation options but no live table model or existing-deck editor. Conformance covers legal supported final direct output:

- omitted table/cell fill imports as absent table consensus `undefined`;
- a uniform legal table solid fill imports as one native `TableCellFill` consensus;
- a legal cell solid override produces mixed direct state and projects `undefined` while per-cell snapshots remain available;
- native assignment changes only final direct cell fill choices and survives write/reopen.

PptxGenJS collapses its none input to absence and omits alpha for explicit zero transparency. Native preserves direct `{ kind: 'none' }` and explicit zero as distinct serialized intent, so those cases are documented strict extensions rather than claimed byte parity. Runtime invalid/coercible fills remain outside the strict native API.

## Test design

### Internal and model tests

- Add focused internal tests for uniform none, sRGB/theme solid, omitted/zero/fractional/full transparency, canonical color parsing, all absent, present/absent and none/solid mixtures, differing colors/alpha, malformed and advanced choices, repeated choices, missing/repeated `tcPr`, empty rows/table, ambiguous table paths, merge continuations, and qualified/descendant lookalikes.
- Apply none, multiple solid variants, and clear across multi-row/multi-column tables.
- Require exact same-value and all-absent-clear no-ops, one committed slide mutation for a real change, late-cell atomicity, advanced-choice replacement/clear, and preservation of unrelated XML.
- Require invalid-input early rejection, transaction rollback, stable model identity, duplicate isolation, move, all six formats, write, and reopen.
- Require public typing to read and assign `TableCellFill | undefined` and reject unsupported values.
- Re-run completed table-level margins, horizontal alignment, text direction, and vertical alignment tests.

### PptxGenJS evidence

- Import real PptxGenJS 4.0.1 omitted fill and require `undefined` consensus.
- Import a uniform legal table solid fill and require the equivalent native direct snapshot.
- Import one legal cell solid override and require mixed `undefined` plus the exact per-cell matrix.
- Normalize an imported mixed table with the native setter, write/reopen, and require every physical direct fill to match.
- Preserve the documented differences for PptxGenJS none collapse and explicit-zero alpha collapse.

### Package and browser evidence

- Require generated declarations to contain the getter and setter inside the packed `TableModel` block.
- Extend actual-tarball Node, TypeScript, browser-condition, installed CLI package inspection, and real-Chrome smoke with stable `tableFill: true` and detailed state.
- Cover uniform read, detached read, no-op isolation, mixed projection, none/solid overwrite, clear, reopen, and invalid failure isolation.
- Inspect packed PPTX slide XML and require one expected direct fill choice per physical cell without border/text-fill false positives.
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
- Resolving effective table-style, theme-inherited, neighboring-cell, or shared-edge fill.
- Returning a mixed sentinel or per-cell matrix from the table-level getter.
- Adding gradient, pattern, picture, or group fill to the public `TableCellFill` value model.
- Changing `TableCell.fill`, `setCellFill()`, table creation precedence/output, or explicit-none/zero-alpha semantics.
- Adding table-level border or fit, merge editing, row/column insertion/deletion, auto-page, repeated headers, or `tableToSlides`.
- Performing content measurement or layout/geometry recomputation.
- Generalizing table properties behind a higher-order consensus framework.
- Copying PptxGenJS invalid-value coercion or collapsed direct intent into native APIs.

## Acceptance criteria

The item is complete when `TableModel.fill` safely projects one uniform supported direct physical-cell fill, atomically whole-replaces or clears it across every physical cell, preserves absence as `undefined`, preserves explicit none and explicit zero transparency, preserves unrelated bytes and lifecycle behavior, rejects invalid or unsafe edits without package mutation, imports legal PptxGenJS table output within the documented direct-state boundary, keeps completed table-level properties unchanged, passes focused/full/performance/type/bundle/declaration gates, the actual tarball and real Chrome report `tableFill: true` with zero browser errors, documentation moves the table-level fill getter/editor to supported, and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
