# Table-Level Text Direction Design

## Status

Validated under the standing direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. Product and implementation choices are delegated, so the selected design proceeds without a separate decision gate. This is the advanced-table item immediately after table-level direct vertical alignment.

## Goal

Add a lossless native `TableModel.textDirection` read/write surface that projects one uniform direct table-cell text direction and atomically applies or clears that direct state across every physical cell in an existing table.

## Current context

The library already supports:

- strict `horz`, `vert`, `vert270`, and `wordArtVert` values through `AddTableOptions.textDirection` and per-cell creation options;
- cell-value precedence over the table creation option;
- PptxGenJS-compatible creation output where `horz` and omitted state both omit `a:tcPr@vert`, while the three non-horizontal values write direct tokens;
- direct `TableCell.textDirection` snapshots;
- `TableModel.setCellTextDirection()` editing, clearing, exact no-ops, malformed-state protection, duplication isolation, rollback, and write/reopen;
- PptxGenJS 4.0.1 final-output conformance for table and cell creation values.

DrawingML has no table-wide text-direction property. PptxGenJS's table-level option is materialized into physical cell `a:tcPr@vert` attributes, except that horizontal is represented by absence. The original creation option is therefore not recoverable from an imported or reopened deck. The native gap is a direct-state consensus projection and bulk editor, not persisted table metadata.

## Considered approaches

### 1. Shared physical-cell resolver plus direction-specific consensus — selected

Extract the exact direct `graphicFrame -> graphic -> graphicData -> tbl -> tr -> tc` resolver already proven by table-level vertical alignment into a focused internal helper. Both table-level properties use the same physical-cell set, while their value parsing and patching remain in separate property-specific modules.

This removes structural duplication, guarantees matching malformed/merge behavior, and creates a narrow reusable boundary for later table-level direct-state editors without introducing a generic value abstraction.

### 2. Duplicate the table-path parser in the text-direction module

Copying the existing resolver is the smallest local diff, but it creates two sources of truth for table structure and makes later table-level properties likely to repeat the same code. Small future changes could produce inconsistent safety decisions. It is rejected.

### 3. Generic table consensus and bulk-property framework

A higher-order reader/replacer could parameterize cell parsing and patching for every direct property. At two consumers this obscures simple control flow, complicates error ownership, and forces unlike OOXML structures into one abstraction. It is rejected as premature.

## Public API

`TableModel` gains one live property:

```ts
class TableModel {
  get textDirection(): TableCellTextDirection | undefined;
  set textDirection(value: TableCellTextDirection | undefined);
}
```

Usage:

```ts
const table = slide.addTable([
  ['North', '$1.2M'],
  ['South', '$980K'],
], { textDirection: 'vert270' });

table.textDirection; // 'vert270'
table.setCellTextDirection(0, 1, 'vert');
table.textDirection; // undefined: physical cells are mixed
table.textDirection = 'wordArtVert'; // all physical cells become stacked
table.textDirection = 'horz'; // writes explicit direct horz on every cell
table.textDirection = undefined; // clears every direct vert attribute
```

The property uses the existing `TableCellTextDirection` union. No alias type, table-default object, `mixed` token, compatibility facade, or additional creation option is added.

## Getter semantics

The getter resolves all physical `<a:tc>` elements under the table in direct row and column order.

- One or more cells with the same supported direct `vert` token returns `horz`, `vert`, `vert270`, or `wordArtVert`.
- All direct attributes absent returns `undefined`; absence is not synthesized as `horz`.
- A mixture of values, or a mixture of present and absent values, returns `undefined`.
- Empty tables, incomplete direct rows, ambiguous table paths, repeated/missing direct `tcPr`, repeated `vert`, and unsupported tokens return `undefined` without throwing during a read.
- Merge continuations are physical cells and participate in the consensus.
- Descendant lookalikes outside the exact direct table path do not participate.

The distinction between absent and explicit `horz` is intentional. PowerPoint treats absence as horizontal, and PptxGenJS omits explicit horizontal creation values, but the native getter reports only recoverable direct state. Callers that need the cell matrix continue to inspect `table.rows[].cells[].textDirection`.

## Setter semantics

Assignment validates before package access. Only `horz`, `vert`, `vert270`, `wordArtVert`, and `undefined` are accepted; coercible, boxed, whitespace-padded, unsupported OOXML tokens, objects, arrays, and symbols throw `TypeError` without mutation.

For a valid assignment:

1. resolve one live table snapshot inside one OPC transaction;
2. require one complete non-empty direct physical-cell set;
3. patch every cell's unique direct `tcPr@vert` in the same in-memory XML document;
4. commit the slide XML once only when at least one cell changed.

Assigning an already uniform explicit value is an exact byte and mutation-journal no-op. Assigning `undefined` when all cells already omit `vert` is also an exact no-op. Mixed state is normalized by overwriting every cell. Assigning `horz` deliberately writes `vert="horz"` to each cell, matching the existing single-cell editor and preserving an explicit editable direct state; clearing is the only operation that removes the attribute.

Clearing removes only direct `vert` attributes and preserves margins, alignment, borders, fill, vertical alignment, fit, text bodies, row/grid geometry, transforms, extensions, relationships, and unrelated shapes or parts. An empty or ambiguous physical-cell set, missing/repeated direct `tcPr`, or repeated direct `vert` makes the edit unsafe and throws `ModelParseError`. One unsupported direct token may be repaired by assigning a supported value or cleared with `undefined`, matching single-cell replacement behavior. Slide bytes are committed only after every cell succeeds, so a malformed later cell cannot cause a partial package mutation.

## Shared structural helper

Create a focused internal module that exports one function equivalent to:

```ts
function readDirectTablePhysicalCells(
  frame: XmlElement,
): readonly XmlElement[] | undefined;
```

The helper owns only exact structural resolution:

- root local name must be `graphicFrame`;
- exactly one direct `graphic`, `graphicData`, and `tbl` must exist;
- at least one direct `tr` must exist;
- every direct row must contain at least one direct `tc`;
- returned order is physical row-major order and includes merge continuations.

It does not read values, inspect `tcPr`, mutate XML, throw model errors, or infer visible cells. Vertical alignment moves to this helper without changing its API or behavior. Text direction consumes it for the new getter and setter. Existing column-width and row-height structural logic remains untouched because those editors also own grid/transform invariants beyond a physical-cell set.

## OOXML and lifecycle behavior

The editor owns only unqualified direct `vert` on each physical cell's unique direct `a:tcPr`. Attribute insertion, replacement, and whitespace-preserving deletion reuse `replaceTableCellTextDirection()`. The slide is serialized at most once per successful bulk edit.

The property is live against current OOXML and preserves existing `TableModel` identity. Slide duplication copies the current direct cell state; later source and duplicate edits are isolated. Move, write, reopen, and all six presentation formats preserve the direct projection. User transactions roll back the entire bulk edit and mutation journal. Validation failures and unsafe-state edits preserve package bytes, relationships, diagnostics, model collections, and journal state.

## Conformance boundary

PptxGenJS 4.0.1 exposes creation options but no live table model or post-creation table text-direction editor. Conformance therefore covers final supported OOXML output:

- uniform non-horizontal table-level output imports with the matching native property;
- explicit PptxGenJS `horz` collapses to attribute absence and therefore imports as `undefined` at table level;
- a cell override produces mixed direct state and projects `undefined`;
- native assignment changes only final direct cell attributes and remains valid after write/reopen.

Native bulk editing and explicit direct `horz` are intentional extensions required for complete PPT editing. Invalid PptxGenJS runtime values such as `eaVert` remain preserved in OOXML but are not exposed as valid native state; assigning a legal native value may repair them.

## Test design

### Internal and model tests

- Read uniform explicit `horz`, `vert`, `vert270`, and `wordArtVert`; all absent; present/absent mixed; mixed legal values; unsupported tokens; repeated attributes; repeated/missing `tcPr`; empty rows/table; ambiguous direct table paths; and opaque descendant lookalikes.
- Apply all four values and clear across multi-row/multi-column tables including merge continuations and imported mixed/unsupported state.
- Require explicit `horz` insertion, exact same-value and all-absent-clear no-ops, one committed slide mutation for a real change, and preservation of unrelated cell/table properties.
- Require invalid-input early rejection, malformed-late-cell atomicity, transaction rollback, stable model identity, duplicate isolation, move, write, and reopen.
- Require public typing to accept the four tokens plus `undefined` and reject unsupported values.
- Re-run existing table-level vertical-alignment tests after extracting the shared structural resolver.

### PptxGenJS evidence

- Import real PptxGenJS 4.0.1 uniform `vert`, `vert270`, and `wordArtVert` table-level output and require matching native consensus.
- Import PptxGenJS horizontal/omitted output and require `undefined`, documenting horizontal collapse.
- Import a non-horizontal table default with horizontal, alternate-direction, and unsupported cell overrides and require `undefined` plus the exact cell matrix.
- Normalize an imported mixed table through the native setter, write/reopen, and require every physical direct token to match.

### Package and browser evidence

- Require generated declarations to contain the getter and setter.
- Extend actual-tarball Node, TypeScript, browser-condition, installed CLI package inspection, and real-Chrome smoke with a stable `tableTextDirection: true` state.
- Cover uniform read, read isolation, mixed projection, all four bulk assignments, explicit-horizontal behavior, clear, no-op, invalid failure isolation, duplicate isolation, and write/reopen.
- Inspect packed PPTX slide XML and require the expected direct `tcPr@vert` state without `bodyPr@vert` false positives.
- Keep validation, console, page, and network error counts at zero.

## Delivery boundaries

The work is delivered as separately reviewed and pushed commits:

1. this design;
2. the implementation plan;
3. shared structural helper, core property implementation, source tests, and PptxGenJS conformance;
4. actual-package declarations, Node/browser/CLI, and real-Chrome evidence;
5. public documentation, compatibility matrix, changelog, and progress closeout.

Each commit stages only its declared files. `.pnpm-store/`, generated workspace tarballs, and retained `/tmp` evidence are never committed.

## Non-goals

- Persisting or reconstructing a table creation default.
- Treating absent `vert` as direct `horz` in the getter.
- Returning a `mixed` sentinel or matrix from the table-level getter.
- Changing `TableCell.textDirection`, `setCellTextDirection()`, or creation precedence/output.
- Supporting additional DrawingML direction tokens outside the PptxGenJS-compatible table union.
- Adding table-level fit, horizontal alignment, margins, borders, fill, merge editing, row/column insertion/deletion, auto-page, repeated headers, or `tableToSlides` in this item.
- Generalizing all table properties behind a higher-order consensus framework.
- Copying PptxGenJS invalid-value coercion or warning-only behavior.

## Acceptance criteria

The item is complete when `TableModel.textDirection` safely projects one uniform explicit direct cell value, atomically applies or clears it across every physical cell, writes explicit `horz` while preserving absence as `undefined`, preserves unrelated bytes and lifecycle behavior, rejects invalid or unsafe edits without package mutation, imports legal PptxGenJS table output within the documented horizontal-collapse boundary, keeps table-level vertical alignment unchanged after the shared-helper extraction, passes focused/full/performance/type/bundle/declaration gates, the actual tarball and real Chrome report `tableTextDirection: true` with zero browser errors, documentation moves the table-level text-direction getter/editor to supported, and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
