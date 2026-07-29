# Table Margin Creation Design

## Goal

Add strict table-level `margin` creation to native `slide.addTable()` so one point-based scalar, TRBL tuple, or partial named default is materialized onto every physical cell. Cell-level `AddTableCellOptions.margin` remains the higher-precedence layer. Match PptxGenJS 4.0.1 final direct OOXML for its public scalar/TRBL values while preserving the native API's stable point units, strict validation, immediate model access, and transactional behavior.

## Public Contract

`AddTableOptions` gains:

```ts
readonly margin?: TextBoxMarginInput;
```

The value model remains:

```ts
type TextBoxMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | TextBoxMargins;
```

All native values are points. A scalar applies to all four sides. A tuple uses exact top/right/bottom/left order. A named object may provide any subset of `top`, `right`, `bottom`, and `left`.

Creation resolves each side independently in this order:

1. canonical cell defaults: top/bottom `3.6pt`, left/right `7.2pt`;
2. supplied table margin sides;
3. supplied cell margin sides.

Therefore a cell scalar or tuple replaces all four table sides, while a partial named cell value overrides only its supplied sides and inherits remaining table/default sides. A string cell, `{ text }`, empty cell options, omitted/runtime-undefined cell margin, `{}`, or an all-undefined named cell margin inherits the table layer. Omitted/runtime-undefined/empty table margin preserves current bytes.

The table value is only a creation input. Resolved values are materialized on physical cells, and no table metadata or source-layer state is retained. `TableCell.margins` immediately exposes the four direct values. `TableModel.setCellMargins()` keeps its existing whole-replacement editor semantics; clearing a created cell does not reapply the original table margin.

## Approach Selection

Three approaches were considered:

1. Overlay normalized sides during table-definition normalization. This is selected because it reuses the descriptor-safe point normalizer and existing one-pass renderer, preserves creation/editor semantic separation, and stores only the final cell state.
2. Create the table and call `setCellMargins()` for each cell. This would repeatedly parse and serialize the slide, produce unnecessary mutations, and incorrectly bind creation defaults to the editor's whole-replacement behavior.
3. Retain a table-level default in `TableModel` or custom OOXML metadata. PowerPoint has no matching table margin property in this output path, PptxGenJS materializes the value, and retained state would create unsupported re-inheritance after clearing.

The native API does not copy PptxGenJS's legacy unit switch because a number whose unit changes at `1` cannot form a stable native value contract.

## Normalization and Precedence

`normalizeTableDefinition()` first normalizes every cell into a detached value, then reads table options through the existing own-data-property validator. `margin` is added to the allowed table keys and a defined value is passed to `normalizeTextBoxMargins()` with table-specific error context.

The shared normalizer continues to enforce:

- finite numbers quantized to the nearest EMU and bounded by signed Int32;
- ordinary dense exact-length tuples containing only own data indices and `length`;
- ordinary or null-prototype named objects containing only own data sides;
- rejection of accessors, inherited values, class/exotic objects, sparse arrays, extra keys, and symbol keys without invoking getters;
- immediate detachment from caller-owned arrays and objects.

When the table layer contains at least one normalized side, each normalized cell receives a detached merged margin object. Cell normalized sides overwrite matching table sides. Empty table values may take the same merge path but must render byte-identically to omission.

Invalid table margin input throws before geometry rendering, package mutation, relationship changes, ZIP state changes, or mutation journal entries. Existing cell validation order and package transaction behavior remain unchanged.

## OOXML Ownership and Ordering

No new table-level OOXML property is written. Every resolved side is emitted only as the physical cell's direct `a:tcPr@marL`, `marR`, `marT`, or `marB` integer attribute.

Cell property serialization remains:

1. `marL`, `marR`, `marT`, and `marB` attributes;
2. optional `anchor` attribute from cell/table `valign`;
3. `lnL`, `lnR`, `lnT`, and `lnB` border children;
4. optional cell fill child.

Examples after quantization:

- table `margin: 0` gives four zero attributes on uncovered cells;
- table `margin: [3.6, 7.2, 10.8, 14.4]` gives T/R/B/L `45,720/91,440/137,160/182,880` EMU;
- table `margin: { top: 9, left: 18 }` keeps canonical right/bottom and changes top/left;
- with that table value, cell `margin: { bottom: 12 }` keeps table top/left, canonical right, and changes bottom;
- cell `margin: 0` overrides all four table sides.

Margins never write `bodyPr` insets, resize cells, change row/column dimensions, or trigger content measurement.

## PptxGenJS 4.0.1 Compatibility

PptxGenJS 4.0.1 declares both table and cell `margin` as `number | [number, number, number, number]`. At generation time it copies a table value to cells without a cell value and serializes only direct `tcPr` margins. A cell value wins. With both layers omitted, it uses the same narrow direct defaults as native creation.

For final direct state, native table values can match PptxGenJS values by expressing the same dimensions in points:

- native `0` matches PptxGenJS `0`;
- native `1` matches PptxGenJS `1`;
- native `7.2` matches PptxGenJS `0.1`;
- native `[3.6, 7.2, 10.8, 14.4]` matches PptxGenJS `[0.05, 0.1, 0.15, 0.2]`;
- native `-7.2` matches PptxGenJS `-0.1`.

Intentional differences remain:

- native always uses points, while PptxGenJS treats a margin whose first value is below `1` as inches and a first value at least `1` as points;
- native adds descriptor-safe partial named table and cell values with per-side layering;
- native rejects coercion, non-finite values, unsafe quantized values, and malformed object shapes instead of emitting ambiguous or invalid XML.

The adapter continues to expose final direct EMU values as points and does not infer the original input layer or unit.

## Verification

Tests must prove:

- scalar, TRBL, partial named, empty, runtime-undefined, zero, negative, fractional, and signed-Int32-bound table normalization;
- canonical → table → cell side precedence for strings and every supported object-cell form;
- table omission/runtime undefined/empty byte equality and caller detachment;
- exact marL/marR/marT/marB token order, coexistence with `anchor`, L/R/T/B borders, and fill;
- immediate snapshots, whole-replacement edit/clear behavior, duplicate isolation, rollback, write, and reopen;
- invalid table inputs are getter-free and leave slide bytes, identity, and mutation journal unchanged;
- PptxGenJS 4.0.1 final-state equality for paired scalar/TRBL values and documented dual-unit differences;
- public declarations plus Node, browser, CLI, and packed-tarball smoke;
- PowerPoint 2010 validation, package-diff isolation, LibreOffice rendering, Poppler rasterization, and overflow checks.

## Scope Boundary

This slice adds table-level `margin` only during native table creation. It does not add `TableModel.margins`, a table-level margin getter/editor, PptxGenJS dual-unit interpretation, horizontal alignment, text direction/fit creation, table-level border/fill, rich or multi-paragraph cells, merge/span, hyperlinks, table styles, auto-page, repeated headers, row insertion/deletion, content measurement, effective-style inheritance, or layout recomputation.
