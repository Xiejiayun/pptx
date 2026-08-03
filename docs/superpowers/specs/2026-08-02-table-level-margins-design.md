# Table-Level Margins Design

## Status

Validated under the standing direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. The user delegated product and implementation choices, so this design proceeds without a separate decision gate. It follows the completed table-level vertical-alignment, text-direction, and horizontal-alignment consensus/bulk editors.

## Goal

Add a lossless native `TableModel.margins` read/write surface that projects one uniform direct physical-cell margin vector and atomically applies or clears direct `a:tcPr@marL/marR/marT/marB` state across every physical cell in an existing table.

## Current context

The library already supports:

- strict point-based scalar, four-value `[top, right, bottom, left]`, partial named-object, empty-object, and `undefined` margin input;
- descriptor-safe, getter-free normalization that detaches caller input and quantizes point values to signed Int32 EMU;
- canonical table creation defaults of 3.6pt top/bottom and 7.2pt left/right;
- table creation values materialized into physical cells with cell-level side precedence;
- `TableCell.margins` direct partial snapshots;
- `TableModel.setCellMargins()` whole-replacement editing and clearing;
- exact single-cell no-ops, malformed-state protection, duplicate isolation, rollback, all-six-format write/reopen, and unrelated-state preservation;
- PptxGenJS 4.0.1 table/cell creation conformance, including its legacy less-than-one inches versus at-least-one points input rule.

DrawingML has no table-wide stored margin property. PptxGenJS's table-level `margin` option is a creation value materialized into direct attributes on every physical `a:tcPr`; only that final cell state survives write/reopen. The remaining gap is therefore a uniform projection and bulk editor over recoverable direct cell attributes.

## Considered approaches

### 1. Partial-vector consensus property plus atomic all-cell setter — selected

Expose live `TableModel.margins` accessors. The getter returns a detached `TextBoxMargins` object only when every physical cell has the same non-empty, structurally safe set of direct margin sides and numeric values. Assignment accepts the existing `TextBoxMarginInput` forms, normalizes once before package access, and whole-replaces every physical cell's four direct margin attributes in one transaction.

This reuses the public margin value model, reflects only serialized OOXML, preserves partial direct state, and gives existing-deck editing the same lossless semantics already available for one cell.

### 2. Require all four sides in the getter

Requiring a complete TRBL vector would simplify the result, and newly created tables normally materialize all four sides. Existing decks can legally contain only some direct margin attributes, and the current cell API intentionally exposes those partial values. Rejecting uniform partial state would lose recoverable information and make the table projection inconsistent with cell snapshots. It is rejected.

### 3. Effective defaults or retained creation metadata

Filling absent sides from canonical defaults, a table style, or the original creation input could always produce a four-side value. Those sources are not a single recoverable direct table property, and creation metadata does not survive reopen. Synthesizing them would make imported and created decks behave differently and blur direct versus effective state. It is rejected.

## Public API

`TableModel` gains one live property with separate getter and setter types, matching the established `ShapeModel.textMargins` pattern:

```ts
class TableModel {
  get margins(): TextBoxMargins | undefined;
  set margins(value: TextBoxMarginInput | undefined);
}
```

Usage:

```ts
const table = slide.addTable([
  ['North', '$1.2M'],
  ['South', '$980K'],
], { margin: [3.6, 7.2, 10.8, 14.4] });

table.margins; // { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 }
table.setCellMargins(0, 1, { top: 2 });
table.margins; // undefined: physical cells are mixed
table.margins = 6; // all four sides on every physical cell become 6pt
table.margins = { top: 2, left: 4 }; // clears right/bottom on every cell
table.margins = undefined; // clears all four direct attributes
```

No alias type, table-default object, `mixed` sentinel, effective-value API, or compatibility facade is added.

## Getter semantics

The getter resolves the exact direct table path and enumerates all physical `a:tc` elements, including merge continuations.

- One or more cells with the same non-empty safe direct side set and values returns one detached `TextBoxMargins` snapshot.
- A uniform complete vector and a uniform partial vector are both valid.
- Values compare after signed Int32 numeric parsing, so lexically different tokens such as `12700` and `0012700` represent the same 1pt value.
- All four sides absent returns `undefined`.
- Present/absent mixtures, differing side sets, or differing numeric values return `undefined`.
- Empty tables, ambiguous direct table paths, missing/repeated direct `tcPr`, repeated owned attributes, malformed/out-of-range owned tokens, and other unsafe state return `undefined` without throwing during a read.
- Qualified attributes and descendant lookalikes outside the selected direct `tcPr` do not participate.

The existing `TableCell.margins` reader remains unchanged and continues its per-side tolerant snapshot behavior. The table-level getter uses a stricter private cell-state read so one malformed owned side cannot be mistaken for a valid uniform partial vector. Callers needing mixed or per-cell detail inspect `rows[].cells[].margins`.

## Setter semantics

Assignment calls `normalizeTextBoxMargins(value, 'Table margins')` before package access. Accepted inputs are:

- one finite number for all four sides;
- one dense ordinary four-value `[top, right, bottom, left]` tuple;
- one ordinary/null-prototype partial named object;
- `{}` or `undefined` to clear all four sides.

Invalid, accessor-backed, inherited, sparse, subclassed, symbol-bearing, non-finite, or signed-Int32-overflow input rejects with the existing stable `TypeError` or `RangeError` before mutation.

For a valid assignment:

1. resolve one live table snapshot inside one OPC transaction;
2. require one complete non-empty direct physical-cell set;
3. call the existing `replaceTableCellMargins()` for every physical cell on the same in-memory slide XML document;
4. serialize and commit the slide at most once, and only if at least one cell changed.

Scalar and tuple input write all four sides. Partial named input performs the existing whole replacement: supplied sides are written and omitted sides are removed from every cell. `{}` and `undefined` both clear all four direct attributes. No operation restores a creation default after clearing.

Assigning a numerically equivalent uniform value is an exact bytes and mutation-journal no-op. Clearing an all-absent table is also an exact no-op. Mixed values and malformed single attributes may be normalized by legal replacement or clearing. Missing/repeated direct `tcPr` or repeated owned attributes make the edit unsafe and throw `ModelParseError`. Because package bytes are committed only after every physical cell succeeds, an unsafe later cell cannot leave a partial package mutation.

## OOXML and lifecycle behavior

The editor owns only unqualified direct `marL`, `marR`, `marT`, and `marB` attributes on each physical cell's unique direct `a:tcPr`. It never reads or writes text-body `lIns/rIns/tIns/bIns`, table styles, table metadata, extensions, or effective layout values.

Bulk editing preserves text, runs, paragraph properties, horizontal and vertical alignment, text direction, fit, borders, fill, row/grid geometry, transforms, merge-continuation identity, extensions, relationships, and unrelated shapes or package parts. Attribute insertion order continues to follow the existing cell margin replacer.

The property is live against current OOXML and preserves `TableModel` identity. Slide duplication copies current direct cell state; later source and duplicate edits remain isolated. Move, write, reopen, and all six presentation formats preserve the projection. User transactions roll back the whole bulk edit and mutation journal. Validation failure or unsafe-state edits preserve bytes, relationships, diagnostics, model collections, and journal state.

## Conformance boundary

PptxGenJS 4.0.1 has creation options but no live table model or existing-deck editor. Conformance covers its legal final direct output:

- omitted margins import as the explicit canonical cell vector that PptxGenJS materializes;
- a uniform table-level scalar or TRBL value imports as one native point-based table consensus;
- one legal cell override produces mixed direct state and projects `undefined`;
- native assignment changes only final direct cell margin attributes and survives write/reopen.

PptxGenJS's legacy runtime interprets the first margin value below 1 as inches and values at least 1 as points. Native continues to use points exclusively. Tests compare equivalent final EMU, for example PptxGenJS `0.1` with native `7.2`, rather than copying input-unit ambiguity. Native bulk editing is an intentional existing-deck extension.

## Test design

### Internal and model tests

- Add focused internal tests for uniform complete and partial vectors, numeric lexical equivalence, all absent, present/absent mixed, differing keys/values, malformed/out-of-range/repeated attributes, missing/repeated `tcPr`, empty rows/table, ambiguous table paths, merge continuations, and qualified/descendant lookalikes.
- Apply scalar, TRBL, partial named object, empty object, and `undefined` across multi-row/multi-column tables.
- Require partial whole-replacement clearing, exact same-value and all-absent-clear no-ops, one committed slide mutation for a real change, late-cell atomicity, and preservation of unrelated XML.
- Require invalid-input early rejection, transaction rollback, stable model identity, duplicate isolation, move, all six formats, write, and reopen.
- Require public typing to read `TextBoxMargins | undefined`, accept `TextBoxMarginInput | undefined`, and reject unsupported values.
- Re-run completed table-level horizontal alignment, text direction, and vertical alignment tests.

### PptxGenJS evidence

- Import real PptxGenJS 4.0.1 omitted margins and require the explicit canonical direct consensus.
- Import a uniform table margin using its legacy inch input and require the equivalent native point vector.
- Import one legal cell override and require `undefined` plus the exact per-cell matrix.
- Normalize an imported mixed table with the native setter, write/reopen, and require every physical direct margin vector to match.

### Package and browser evidence

- Require generated declarations to contain the getter and setter inside the packed `TableModel` block.
- Extend actual-tarball Node, TypeScript, browser-condition, installed CLI package inspection, and real-Chrome smoke with stable `tableMargins: true` and detailed state.
- Cover uniform read, read isolation, mixed projection, scalar overwrite, partial whole replacement, clear, no-op, invalid failure isolation, and write/reopen.
- Inspect packed PPTX slide XML and require the expected four direct margin attributes per physical cell without text-body inset false positives.
- Keep validation, console, page, and network error counts at zero.

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
- Filling absent sides from canonical defaults, table styles, or neighboring cells.
- Returning a `mixed` sentinel or per-cell matrix from the table-level getter.
- Changing `TableCell.margins`, `setCellMargins()`, table creation precedence/defaults/output, or point-only native input semantics.
- Adding table-level border, fill, fit, merge editing, row/column insertion/deletion, auto-page, repeated headers, or `tableToSlides`.
- Performing content measurement or layout/geometry recomputation.
- Generalizing table properties behind a higher-order consensus framework.
- Copying PptxGenJS invalid-value coercion or legacy input-unit ambiguity into native APIs.

## Acceptance criteria

The item is complete when `TableModel.margins` safely projects one uniform non-empty direct complete or partial physical-cell margin vector, atomically whole-replaces or clears it across every physical cell, preserves absence as `undefined`, preserves unrelated bytes and lifecycle behavior, rejects invalid or unsafe edits without package mutation, imports legal PptxGenJS table output within the documented unit/direct-state boundary, keeps completed table-level alignment and direction properties unchanged, passes focused/full/performance/type/bundle/declaration gates, the actual tarball and real Chrome report `tableMargins: true` with zero browser errors, documentation moves the table-level margins getter/editor to supported, and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
