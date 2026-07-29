# Table Horizontal Alignment Creation Design

## Goal

Add strict table-level `align` creation to native `slide.addTable()` so one left, center, right, or justify default is materialized onto every single-paragraph cell that does not provide its own cell-level `align`. Match PptxGenJS 4.0.1 final direct OOXML while preserving strict native validation, lossless omission, detached normalization, and transactional behavior.

This is one independently reviewable parity slice. Cell-level horizontal-alignment creation already exists. Existing-deck horizontal-alignment snapshots and editing remain separate later items because a pre-existing cell can contain multiple paragraphs with different direct states.

## Public Contract

`AddTableOptions` gains:

```ts
readonly align?: TextAlignment;
```

The only accepted values are `left`, `center`, `right`, and `justify`. Table-level omission and runtime `undefined` do not change existing output. A valid cell-level `AddTableCellOptions.align` value overrides the table value. A string cell, `{ text }`, `{ text, options: {} }`, or `{ text, options: { align: undefined } }` inherits a defined table value.

The table default is a creation input, not persistent table metadata. Creation materializes the resolved value into each physical cell's only paragraph as direct `a:p/a:pPr@algn`. The current `TableModel.setCellText()` replaces only text, so it preserves the materialized paragraph property. A future cell horizontal-alignment editor will operate on the direct paragraph state and will not reapply the original table default after clearing it.

## Alternatives

1. Resolve the table value onto normalized cells before rendering. This is the selected approach. It follows table-level `border`, `fill`, `margin`, and `valign` propagation, keeps precedence observable in a detached definition, and reuses the existing paragraph renderer without another token map.
2. Retain table-level alignment metadata and compute effective alignment dynamically. This would introduce a style-inheritance layer that neither the current model nor PptxGenJS's emitted final state has, so it is not selected.
3. Pass a table default directly into `renderTableCell()`. This would couple normalization precedence to rendering and make strict pre-mutation validation harder to verify, so it is not selected.

## OOXML Ownership and Ordering

There is no table-level OOXML alignment property. The resolved value belongs only to the created cell's single paragraph:

| API value | `a:pPr@algn` |
| --- | --- |
| `left` | `l` |
| `center` | `ctr` |
| `right` | `r` |
| `justify` | `just` |

`renderTableCell()` continues to build the paragraph through `normalizeRichText()` and `renderRichTextParagraphs(..., { defaultAlign })`. It must not copy the alignment token map or manually construct a second paragraph-property path. Alignment is never written to `a:tcPr`, `a:bodyPr`, table properties, or an extension.

The existing canonical paragraph ordering remains unchanged. Explicit table `left` is materialized as `algn="l"`; it is not optimized away as an effective default. If neither table nor cell provides alignment, direct `algn` remains absent and current table bytes remain unchanged.

## Normalization and Precedence

`normalizeTableDefinition()` continues to normalize every input cell into an owned, descriptor-safe value before reading table options. `align` is added to the table option allowlist. A defined table value is normalized with the existing `normalizeTextAlignment()`.

When the table value exists, normalization creates resolved rows in which only cells whose normalized `alignment` is absent receive the table value. Explicit cell `left`, `center`, `right`, or `justify` remains unchanged. Cell runtime `undefined` is equivalent to omission and therefore inherits. The resulting resolved alignment is stored only on each normalized cell; `NormalizedTableDefinition` gains no table-default field.

Alignment propagation remains independent of border, fill, margin, and vertical-alignment propagation. Their existing cell override and merge rules do not change. No input object, array, color, border, fill, or margin value is retained by reference.

## Validation and Transaction Semantics

Table options keep the existing ordinary/null-prototype, own-data-property contract. Accessors, inherited properties, class instances, arrays, extra string keys, and symbol keys are rejected without invoking getters.

Omitted `align` and `align: undefined` are accepted. Defined values reject null, booleans, numbers, empty strings, case or whitespace variants, OOXML tokens `l`, `ctr`, `r`, and `just`, unsupported DrawingML states such as `dist`, `thaiDist`, and `justLow`, arrays, objects, and symbols. There is no coercion, trimming, case folding, alias, or fallback.

Invalid table alignment throws `TypeError` before slide bytes, package relationships, ZIP state, live model identity, or the mutation journal changes. Existing matrix and cell validation order remains unchanged. `SlideModel.addTable()` keeps its package transaction; no new editor transaction or cache is introduced.

## PptxGenJS 4.0.1 Compatibility

Direct runtime probes of PptxGenJS 4.0.1 confirm:

- valid table `left`, `center`, `right`, and `justify` are copied to uncovered cells as `l`, `ctr`, `r`, and `just`;
- cell omission and runtime `undefined` inherit a valid table value;
- a valid cell value overrides the table value;
- omitted or runtime-undefined table alignment leaves direct `algn` absent;
- invalid table values including `dist`, `thaiDist`, `justLow`, `bogus`, `ctr`, null, false, and zero are silently ignored for uncovered cells while a valid cell override still renders.

Native creation matches the supported final states and precedence. The invalid-runtime behavior deliberately differs: native rejects invalid input rather than silently producing unaligned cells. The PptxGenJS adapter continues to consume the already-materialized paragraph state and requires no production change for this slice.

## Verification

Tests must prove:

1. Internal normalization and exact output for table left/center/right/justify, string/object cells, empty options, runtime-undefined cell alignment, and all valid cell overrides.
2. Omitted and runtime-undefined table alignment are byte-identical to current output; no table, cell-property, or body-property alignment is emitted.
3. Strict invalid values and accessor-backed options fail getter-free before package mutation.
4. Public model lifecycle preserves the resolved direct tokens through text and other cell-property edits, duplicate isolation, outer rollback, write, reopen, and immediate model use.
5. SDK coverage proves public typing, invalid-input isolation, target-slide-only mutation, and interoperability with border, fill, margin, vertical alignment, unequal columns, and unequal rows.
6. PptxGenJS 4.0.1 comparison covers supported final-state equality, valid cell precedence, and the documented invalid fallback difference.
7. Packed Node, browser, declaration, and CLI smoke proves the public option survives the distributed package surface.
8. Changelog, API README, package README, and compatibility baseline distinguish supported table-level creation from still-unsupported existing-deck horizontal-alignment read/edit.
9. TypeScript project references, focused and full tests, performance gate, actual tarball, and staged diff review pass.
10. Real native source/edited/reopened and PptxGenJS baseline decks pass the PowerPoint 2010 validation profile, package-diff isolation, LibreOffice/Poppler visual inspection, and overflow checks without repair, clipping, overlap, or unexpected wrapping.

## Scope Boundary

This slice adds table-level `align` only during native creation for the existing single-paragraph cell input. It does not add `TableModel.horizontalAlignment`, `TableCell.horizontalAlignment`, a table- or cell-level horizontal-alignment editor, effective inheritance, rich or multi-paragraph cell creation, `dist` / `thaiDist` / `justLow`, text-direction or fit creation, hyperlinks, merge/span, auto-page, repeated headers, style computation, row insertion/deletion, content measurement, or layout recomputation. It does not alter other shape paragraph APIs or the PptxGenJS adapter's production behavior.
