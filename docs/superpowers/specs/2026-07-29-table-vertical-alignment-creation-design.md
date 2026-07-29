# Table Vertical Alignment Creation Design

## Goal

Add strict table-level `valign` creation to native `slide.addTable()` so one top, middle, or bottom default is materialized onto every cell that does not provide its own cell-level `valign`. Match PptxGenJS 4.0.1 final direct OOXML while preserving the native API's validation, lossless omission, immediate model access, and transactional behavior.

## Public Contract

`AddTableOptions` gains:

```ts
readonly valign?: TextBoxVerticalAlignment;
```

The only accepted values are `top`, `middle`, and `bottom`. Table-level omission and runtime `undefined` do not change existing output. A cell-level `AddTableCellOptions.valign` value overrides the table value. A string cell, `{ text }`, `{ text, options: {} }`, or `{ text, options: { valign: undefined } }` inherits a defined table value.

The table default is a creation input, not persistent table metadata. Creation materializes the resolved value into each physical cell's direct `a:tcPr@anchor` as `t`, `ctr`, or `b`. After creation, `TableCell.verticalAlignment` exposes that direct state and `TableModel.setCellVerticalAlignment()` edits or clears it. Clearing a cell does not reapply the original table default because the source layer is intentionally not retained.

## OOXML Ownership and Ordering

There is no new table-level OOXML property. The resolved value belongs only to each created cell's direct `tcPr@anchor`; it is never written to `txBody/bodyPr@anchor` or stored in an extension.

Cell property serialization remains:

1. `marL`, `marR`, `marT`, and `marB` attributes;
2. optional resolved `anchor` attribute;
3. `lnL`, `lnR`, `lnT`, and `lnB` children;
4. optional cell fill child.

If neither table nor cell provides a value, no anchor is emitted and existing default cell bytes remain unchanged.

## Normalization and Precedence

`normalizeTableDefinition()` continues to normalize every cell into a detached, descriptor-safe value. It then reads `AddTableOptions` through the existing own-data-property validator, adds `valign` to the supported key set, and normalizes a defined value with `normalizeTextBoxVerticalAlignment()`.

When a table value exists, normalization creates resolved rows in which only cells with no normalized `verticalAlignment` receive the table value. Explicit cell `top`, `middle`, or `bottom` remains unchanged. Cell runtime `undefined` is intentionally equivalent to omission and therefore inherits. No additional renderer branch or long-lived table-default field is needed; the existing cell renderer receives the resolved direct value.

Inputs remain detached immediately. Accessors, inherited properties, symbol keys, extra keys, aliases, coercions, case variants, whitespace variants, `just`, `dist`, `mid`, `center`, and `distributed` are rejected. PptxGenJS runtime alias or invalid-token passthrough is not copied into the strict native API.

## Error and Transaction Semantics

Invalid table `valign` throws `TypeError` before any slide bytes, package relationships, ZIP state, model identity, or mutation journal entry changes. Accessor-backed table options are rejected without invoking the getter. Existing invalid-cell behavior and validation order remain unchanged.

`SlideModel.addTable()` keeps its existing package transaction. No new editor transaction or model cache is introduced.

## Compatibility

For public values, native table-level creation matches PptxGenJS 4.0.1 final state:

- a table value is copied to uncovered cells;
- a valid cell value wins;
- omitted table and cell values leave the direct anchor absent;
- every materialized value uses `tcPr@anchor`.

Native differences remain deliberate: runtime-invalid PptxGenJS values can pass through as opaque anchor tokens, while native creation rejects them; PptxGenJS accepts undocumented runtime aliases in some paths, while native creation accepts only the public three-value union.

## Verification

Tests must prove:

- strict table top/middle/bottom normalization and exact `t/ctr/b` output;
- propagation to string and object cells, including cell runtime `undefined`;
- valid cell override precedence;
- byte equality when table `valign` is omitted or runtime `undefined`;
- attribute/child ordering and absence of `bodyPr@anchor`;
- immediate snapshots, per-cell editing/clearing, duplicate, write, reopen, rollback, and invalid-input isolation;
- PptxGenJS 4.0.1 final-state equality for supported values and documented invalid-token differences;
- public declarations plus Node, browser, CLI, and packed-tarball smoke;
- PowerPoint 2010 validation, package-diff isolation, LibreOffice rendering, Poppler rasterization, and overflow checks.

## Scope Boundary

This slice adds table-level `valign` only during native table creation. It does not add a `TableModel.verticalAlignment` property, a table-level alignment editor, horizontal alignment creation, table-level text direction, table-level fit, rich or multi-paragraph cells, merges, hyperlinks, auto-page, repeated headers, style inheritance, effective-default computation, row insertion/deletion, or layout recomputation.
