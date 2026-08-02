# Table-Level Vertical Alignment Design

## Status

Validated under the standing direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. The user delegated product and implementation choices, so the selected design proceeds without a separate decision gate. This is the first advanced-table item after completing all six PptxGenJS 4.0.1 presentation runtime catalogs.

## Goal

Add a lossless native `TableModel.verticalAlignment` read/write surface that projects a uniform table-wide value from physical cell `a:tcPr@anchor` attributes and atomically applies or clears that direct state across every physical cell in an existing table.

## Current context

The library already supports:

- table-level and cell-level `valign: 'top' | 'middle' | 'bottom'` during `addTable()` creation;
- cell precedence over the table creation option;
- canonical `top -> t`, `middle -> ctr`, and `bottom -> b` DrawingML mapping;
- `TableCell.verticalAlignment` direct snapshots;
- `TableModel.setCellVerticalAlignment()` direct single-cell editing, clearing, no-op behavior, malformed-state protection, duplication isolation, rollback, and write/reopen;
- PptxGenJS 4.0.1 public-output conformance for table/cell creation values.

DrawingML has no table-wide vertical-alignment property. PptxGenJS's table-level `valign` option is a creation default that is materialized into each physical `<a:tc>` as direct `a:tcPr@anchor`. The creation default itself is not stored and cannot be recovered after reopening a PPTX. The remaining native gap is therefore a semantic projection and bulk editor over final direct cell state, not a new package field.

## Considered approaches

### 1. Consensus property plus atomic all-cell setter — selected

Expose a live `TableModel.verticalAlignment` property. Its getter returns a value only when the table contains at least one physical cell and every physical cell has the same safe direct alignment. It returns `undefined` for absent, mixed, empty, or unsafe state. Assignment applies one value to all physical cells; assigning `undefined` clears every direct anchor.

This matches the existing model-property style, reflects only recoverable OOXML, supports editing imported decks, and keeps cell-level overrides available through `setCellVerticalAlignment()`.

### 2. Array or `mixed` sentinel

Returning a matrix or adding a `'mixed'` token would distinguish mixed, absent, and malformed state. The existing `rows[].cells[].verticalAlignment` matrix already provides that detail, while a sentinel would widen the public alignment type with a value that is neither legal input nor PptxGenJS-compatible. It is rejected.

### 3. Synthetic table default metadata

Retaining the original creation option would allow a table getter to remember intent even after cell overrides. That state is not serialized by DrawingML, cannot be recovered after reopen, and would diverge between created and imported decks. Adding a custom extension solely for this convenience would alter files and exceed parity scope. It is rejected.

## Public API

`TableModel` gains one live property:

```ts
class TableModel {
  get verticalAlignment(): TextBoxVerticalAlignment | undefined;
  set verticalAlignment(value: TextBoxVerticalAlignment | undefined);
}
```

Usage:

```ts
const table = slide.addTable([
  ['North', '$1.2M'],
  ['South', '$980K'],
], { valign: 'middle' });

table.verticalAlignment; // 'middle'
table.setCellVerticalAlignment(0, 1, 'top');
table.verticalAlignment; // undefined: physical cells are mixed
table.verticalAlignment = 'bottom'; // all physical cells become bottom aligned
table.verticalAlignment = undefined; // clears all direct anchors
```

The property uses `TextBoxVerticalAlignment`, the same public union already shared by text boxes, slide numbers, table creation, and table cells. No alias type, table-default object, `mixed` token, or compatibility facade is added.

## Getter semantics

The getter enumerates all physical `<a:tc>` elements under the table in row and column order.

- One or more cells with one identical, supported direct `anchor` value returns `top`, `middle`, or `bottom`.
- All anchors absent returns `undefined`.
- A mixture of values, or a mixture of present and absent values, returns `undefined`.
- An empty table, repeated direct `tcPr`, repeated `anchor`, or unsupported anchor token returns `undefined` rather than throwing during a read.
- Merged-cell continuation elements are physical cells and participate in the consensus. The item does not add merge-aware visible-cell semantics.

Callers that need to distinguish absent, mixed, and unsafe state continue to inspect `table.rows[].cells[].verticalAlignment` and diagnostics/package XML. The table-level property intentionally answers only whether one safe uniform direct value exists.

## Setter semantics

Assignment validates the value before package access. Only `top`, `middle`, `bottom`, and `undefined` are accepted; coercible, boxed, whitespace-padded, OOXML-token, object, array, and symbol inputs throw `TypeError` without mutation.

For a valid assignment:

1. resolve one live table snapshot inside an OPC transaction;
2. require at least one physical cell;
3. patch every cell's unique direct `tcPr@anchor` on the same in-memory XML document;
4. commit the slide XML once only if at least one cell changed.

Assigning an already uniform value is an exact byte and mutation-journal no-op. Assigning `undefined` to a table whose cells already omit anchors is also an exact no-op. Mixed state is normalized by overwriting every cell. Clearing removes only direct `anchor` attributes and preserves margins, text direction, borders, fill, text bodies, row/grid geometry, transforms, extensions, relationships, and unrelated shapes/parts.

An empty physical-cell set, missing/repeated direct `tcPr`, or repeated direct `anchor` makes the edit unsafe and throws `ModelParseError`. A single unsupported anchor token may be repaired by assigning a legal value or cleared with `undefined`, matching existing single-cell replacement behavior. Because no slide bytes are committed until all cells are processed, a later malformed cell cannot leave a partial bulk edit.

## OOXML and lifecycle behavior

The editor owns only `anchor` on each physical cell's unique direct `a:tcPr`. It preserves attribute order rules already implemented by `replaceTableCellVerticalAlignment()` and serializes the slide once per successful bulk edit.

The property is live against current OOXML and keeps existing `TableModel` identity. Slide duplication copies the then-current direct cell state; later source or duplicate edits remain isolated. Move, write, reopen, and all six presentation formats preserve the projection. User transactions roll back the entire bulk edit, including the mutation journal. A failed validation or unsafe-state edit leaves bytes, relationships, diagnostics, model collections, and journal unchanged.

## Conformance boundary

PptxGenJS 4.0.1 has no live table model or post-creation table alignment editor. Conformance therefore covers its legal final output:

- a uniform public table-level `valign` output imports with a matching native table property;
- a cell override imports as mixed and projects `undefined`;
- native assignment changes only final direct cell anchors and remains valid after write/reopen.

Native editing is an intentional extension required for full PPT editing. Invalid PptxGenJS runtime values that pass through or disappear are not copied; the native setter remains strict.

## Test design

### Internal and model tests

- Read uniform top/middle/bottom, all-absent, present/absent mixed, mixed legal values, unsupported tokens, repeated anchors, repeated/missing `tcPr`, and empty tables.
- Apply top/middle/bottom and clear across multi-row/multi-column tables, including imported mixed state.
- Require exact no-op bytes/journal, one committed slide mutation for a real change, and preservation of every unrelated cell/table property.
- Require invalid-input early rejection, malformed-late-cell atomicity, transaction rollback, stable model identity, duplicate isolation, move, write, and reopen.
- Require the public property type to accept the three tokens plus `undefined` and reject invalid values.

### PptxGenJS evidence

- Import real PptxGenJS 4.0.1 uniform table-level `valign` output and require the native consensus value.
- Import table-level middle with one top/bottom cell override and require `undefined` plus the exact cell matrix.
- Normalize the imported mixed table through the native setter, write/reopen, and require all physical anchors to match.

### Package and browser evidence

- Require generated declarations to contain the getter and setter.
- Extend actual-tarball Node, TypeScript, browser-condition, installed CLI package inspection, and real-Chrome smoke with a stable `tableVerticalAlignment: true` state.
- In Node and Chrome, cover uniform read, mixed projection, bulk overwrite, clear, write/reopen, mutation isolation, and zero validation/browser errors.
- Inspect the packed PPTX slide part and require the expected direct `anchor="t/ctr/b"` state without unrelated package changes.

## Delivery boundaries

The work is delivered as separately reviewed and pushed commits:

1. this design;
2. the implementation plan;
3. core helper/property implementation, source tests, and PptxGenJS conformance;
4. actual-package declarations, Node/browser/CLI, and real-Chrome evidence;
5. public documentation, compatibility matrix, changelog, and progress closeout.

Each commit stages only its declared files. `.pnpm-store/`, generated workspace tarballs, and retained `/tmp` evidence are never committed.

## Non-goals

- Adding a serializable or transient table creation-default field.
- Returning a `mixed` sentinel or a matrix from the table-level getter.
- Changing `TableCell.verticalAlignment` or `setCellVerticalAlignment()`.
- Adding table-level horizontal alignment, text direction, margins, borders, fill, or fit in the same item.
- Adding merge creation/editing, row/column insertion/deletion, auto-page, repeated headers, or `tableToSlides`.
- Copying PptxGenJS invalid-value coercion or warning-only behavior.
- Changing OOXML for tables that are not assigned through the new property.

## Acceptance criteria

The item is complete when `TableModel.verticalAlignment` safely projects one uniform direct cell value, atomically applies or clears it across all physical cells, preserves unrelated bytes and lifecycle behavior, rejects invalid/unsafe edits without mutation, imports legal PptxGenJS table output, passes focused/full/performance/type/bundle/declaration gates, the actual tarball and real Chrome report `tableVerticalAlignment: true` with zero browser errors, documentation moves the table-level vertical-alignment getter/editor to supported, and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
