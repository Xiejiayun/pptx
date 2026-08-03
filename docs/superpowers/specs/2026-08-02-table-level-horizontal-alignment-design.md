# Table-Level Horizontal Alignment Design

## Status

Validated under the standing direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. Product and implementation choices are delegated, so the selected design proceeds without a separate decision gate. This is the advanced-table item immediately after table-level direct text direction.

## Goal

Add a lossless native `TableModel.horizontalAlignment` read/write surface that projects one uniform direct single-paragraph table-cell horizontal alignment and atomically applies or clears that direct state across every physical cell in an existing table.

## Current context

The library already supports:

- strict `left`, `center`, `right`, and `justify` values through `AddTableOptions.align` and per-cell creation options;
- cell-value precedence over the table creation option;
- PptxGenJS-compatible direct `a:pPr@algn` output for supported values;
- direct `TableCell.horizontalAlignment` snapshots for cells with exactly one direct text body, paragraph, paragraph-properties element, and legal alignment token;
- `TableModel.setCellHorizontalAlignment()` editing, clearing, exact no-ops, malformed-state protection, duplication isolation, rollback, and write/reopen;
- a shared exact direct-table physical-cell resolver used by table-level vertical alignment and text direction;
- PptxGenJS 4.0.1 final-output conformance for table and cell creation values.

DrawingML has no table-wide horizontal-alignment property. PptxGenJS's table-level `align` option is materialized into each uncovered physical cell's paragraph as direct `a:pPr@algn`. The original creation option is not recoverable after import or reopen. The remaining gap is therefore a direct-state consensus projection and bulk editor, not persisted table metadata or effective-style calculation.

## Considered approaches

### 1. Shared physical-cell resolver plus alignment-specific consensus — selected

Extend `table-cell-horizontal-alignment.internal.ts` with table-level read and replace helpers that consume `readDirectTablePhysicalCells()`. Keep paragraph structure, token parsing, and replacement in the alignment module, and expose one live property on `TableModel`.

This follows the two completed table-level properties, reuses their structural safety boundary, and keeps paragraph-specific behavior explicit.

### 2. Duplicate the table-path parser

Copying direct `graphicFrame -> graphic -> graphicData -> tbl -> tr -> tc` traversal into the alignment module would be locally small but would create another source of truth for malformed tables and merge continuations. It is rejected.

### 3. Generic consensus and bulk-property framework

A higher-order helper could parameterize value readers and replacers for vertical alignment, text direction, and horizontal alignment. Their owned OOXML structures differ (`tcPr` attributes versus nested paragraph properties), and three simple consumers do not justify obscuring error ownership or mutation behavior. It is rejected as premature.

## Public API

`TableModel` gains one live property:

```ts
class TableModel {
  get horizontalAlignment(): TextAlignment | undefined;
  set horizontalAlignment(value: TextAlignment | undefined);
}
```

Usage:

```ts
const table = slide.addTable([
  ['North', '$1.2M'],
  ['South', '$980K'],
], { align: 'center' });

table.horizontalAlignment; // 'center'
table.setCellHorizontalAlignment(0, 1, 'right');
table.horizontalAlignment; // undefined: physical cells are mixed
table.horizontalAlignment = 'justify'; // all physical cells become justified
table.horizontalAlignment = 'left'; // writes explicit direct algn="l"
table.horizontalAlignment = undefined; // clears every direct algn attribute
```

The property uses the existing `TextAlignment` union. No alias type, table-default object, `mixed` token, array getter, compatibility facade, or additional creation option is added.

## Getter semantics

The getter resolves all physical `<a:tc>` elements under the exact direct table path in row-major order, including merge continuations.

- One or more cells with the same supported direct single-paragraph token returns `left`, `center`, `right`, or `justify`.
- All direct alignment attributes absent returns `undefined`; absence is not synthesized as `left`.
- A mixture of values, or a mixture of present and absent values, returns `undefined`.
- Zero or multiple direct text bodies, zero or multiple direct paragraphs, zero or multiple direct `pPr` elements, repeated `algn`, unsupported tokens, empty rows/tables, and ambiguous table paths return `undefined` without throwing during a read.
- Qualified or descendant lookalikes outside the exact owned direct path do not participate.
- Rich or multi-paragraph cells are outside the table-level consensus even if their paragraphs happen to share a value.

Callers needing mixed or per-cell detail continue to inspect `table.rows[].cells[].horizontalAlignment`. The table property answers only whether one safe uniform direct value exists.

## Setter semantics

Assignment validates before package access. Only `left`, `center`, `right`, `justify`, and `undefined` are accepted; coercible, boxed, whitespace-padded, OOXML wire tokens, unsupported DrawingML values, objects, arrays, and symbols throw `TypeError` without mutation.

For a valid assignment:

1. resolve one live table snapshot inside one OPC transaction;
2. require one complete non-empty direct physical-cell set;
3. require each cell to contain exactly one direct text body and exactly one direct paragraph;
4. patch every paragraph's unique direct `pPr@algn` in the same in-memory slide XML document;
5. commit the slide XML once only when at least one cell changed.

Missing `pPr` is safe: assigning a value inserts one before the paragraph's first direct element using the paragraph prefix. Assigning `undefined` to a paragraph without direct alignment is an exact no-op and does not create `pPr`. A repeated direct `pPr` or `algn`, multiple paragraphs, or malformed paragraph template makes the edit unsafe and throws `ModelParseError`.

Assigning an already uniform explicit value is an exact byte and mutation-journal no-op. Assigning `undefined` when every cell already omits alignment is also an exact no-op. Mixed or unsupported direct state may be normalized by overwriting every cell with a legal value; an unsupported direct `algn` may also be cleared, matching the single-cell editor. Assigning `left` deliberately writes explicit `algn="l"`; clearing is the only operation that removes the owned attribute.

The bulk edit preserves cell text, runs, paragraph children and unrelated attributes, vertical alignment, text direction, margins, borders, fill, fit, row/grid geometry, transforms, extensions, relationships, and unrelated shapes or parts. Because slide bytes are committed only after every cell succeeds, an unsafe later cell cannot leave a partial package mutation.

## OOXML and lifecycle behavior

The editor owns only the unqualified direct `algn` attribute on each physical cell's unique direct single paragraph `a:pPr`. It never writes alignment to `a:tcPr`, `a:bodyPr`, table properties, a style part, or an extension. Existing `replaceTableCellHorizontalAlignment()` remains the single token and lossless-patch implementation.

The property is live against current OOXML and preserves existing `TableModel` identity. Slide duplication copies current direct paragraph state; later source and duplicate edits are isolated. Move, write, reopen, and all six presentation formats preserve the projection. User transactions roll back the whole bulk edit and mutation journal. Validation failures and unsafe-state edits preserve package bytes, relationships, diagnostics, model collections, and journal state.

## Conformance boundary

PptxGenJS 4.0.1 exposes creation options but no live table model or post-creation table editor. Conformance covers final supported OOXML output:

- uniform table-level `left`, `center`, `right`, or `justify` output imports with the matching native consensus;
- omitted alignment imports as `undefined`, not synthesized `left`;
- a legal cell override produces mixed direct state and projects `undefined`;
- native assignment changes only final direct paragraph alignment and remains valid after write/reopen.

Native bulk editing is an intentional existing-deck extension. PptxGenJS silently drops some invalid table runtime values, while the native creation and editing APIs remain strict. Rich and multi-paragraph cell alignment remains outside this item.

## Test design

### Internal and model tests

- Read uniform explicit left/center/right/justify; all absent; present/absent mixed; mixed legal values; unsupported tokens; repeated attributes; repeated/missing `pPr`; missing/multiple text bodies; zero/multiple paragraphs; empty rows/table; ambiguous direct table paths; merge continuations; and opaque descendant lookalikes.
- Apply all four values and clear across multi-row/multi-column tables, including cells with absent `pPr`, imported mixed/unsupported state, and merge continuations.
- Require explicit `left` insertion, exact same-value and all-absent-clear no-ops, one committed slide mutation for a real change, and preservation of unrelated paragraph/cell/table state.
- Require invalid-input early rejection, malformed-late-cell atomicity, transaction rollback, stable model identity, duplicate isolation, move, write, and reopen.
- Require public typing to accept the four tokens plus `undefined` and reject unsupported values.
- Re-run completed table-level vertical-alignment and text-direction tests to prove the shared physical-cell boundary remains unchanged.

### PptxGenJS evidence

- Import real PptxGenJS 4.0.1 uniform left/center/right/justify table-level output and require matching native consensus.
- Import omitted table alignment and require `undefined` without effective-left synthesis.
- Import one legal cell override and require `undefined` plus the exact per-cell matrix.
- Normalize an imported mixed table through the native setter, write/reopen, and require every physical direct token to match.

### Package and browser evidence

- Require generated declarations to contain the getter and setter.
- Extend actual-tarball Node, TypeScript, browser-condition, installed CLI package inspection, and real-Chrome smoke with stable `tableHorizontalAlignment: true` state.
- Cover uniform read, read isolation, mixed projection, all four bulk assignments, explicit-left behavior, clear, no-op, invalid failure isolation, duplicate isolation, and write/reopen.
- Inspect packed PPTX slide XML and require expected direct `pPr@algn` state without `tcPr@algn` or `bodyPr@algn` false positives.
- Keep validation, console, page, and network error counts at zero.

## Delivery boundaries

The work is delivered as separately reviewed and pushed commits:

1. this design;
2. the implementation plan;
3. core property implementation, source tests, and PptxGenJS conformance;
4. actual-package declarations, Node/browser/CLI, and real-Chrome evidence;
5. public documentation, compatibility matrix, changelog, and progress closeout.

Each commit stages only its declared files. `.pnpm-store/`, generated workspace tarballs, and retained `/tmp` evidence are never committed.

## Non-goals

- Persisting or reconstructing a table creation default.
- Treating absent `algn` as direct `left` in the getter.
- Returning a `mixed` sentinel or matrix from the table-level getter.
- Changing `TableCell.horizontalAlignment`, `setCellHorizontalAlignment()`, or table creation precedence/output.
- Supporting rich or multi-paragraph cell consensus or editing.
- Supporting additional DrawingML alignment tokens outside the current four-value union.
- Adding table-level fit, margins, borders, fill, merge editing, row/column insertion/deletion, auto-page, repeated headers, or `tableToSlides` in this item.
- Generalizing table properties behind a higher-order consensus framework.
- Copying PptxGenJS invalid-value coercion, dropping, or warning-only behavior.

## Acceptance criteria

The item is complete when `TableModel.horizontalAlignment` safely projects one uniform explicit direct single-paragraph cell value, atomically applies or clears it across every physical cell, writes explicit `left` while preserving absence as `undefined`, preserves unrelated bytes and lifecycle behavior, rejects invalid or unsafe edits without package mutation, imports legal PptxGenJS table output within the documented direct-state boundary, keeps completed table-level vertical alignment and text direction unchanged, passes focused/full/performance/type/bundle/declaration gates, the actual tarball and real Chrome report `tableHorizontalAlignment: true` with zero browser errors, documentation moves the table-level horizontal-alignment getter/editor to supported, and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
