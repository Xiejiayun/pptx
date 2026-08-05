# Core Content and Primitive Inputs Surface Family Design

## Scope

Close exactly 14 PptxGenJS 4.0.1 declaration atoms as one evidence-backed
capability family:

- `TableCell.text`, `TableCell.options`, and the string/recursive-cell text
  union members;
- `TextProps.text`, `TextProps.options`, and the string/rich-array
  `Slide.addText()` union members;
- hex/theme `Color` members;
- percentage/numeric `Coord` members;
- scalar/four-value `Margin` members.

The expected classification is five supported and nine deliberate
differences. No deprecated alias, defect exclusion, or unsupported entry is
expected unless the locked runtime control disproves the current analysis.

## Reuse boundary

The native implementation already covers the legal capabilities:

- `SlideModel.addText()` and `addRichText()` expose plain and explicit
  paragraph/run content;
- table creation and live table cells expose plain, structured, and rich
  content with typed options;
- `RichTextColor` separates strict sRGB and scheme state;
- `SlideCoordinate` accepts percentage strings and explicit EMU values;
- `TextBoxMarginInput` accepts scalar and documented top/right/bottom/left
  tuple state.

No new text, table, color, coordinate, or margin codec is added. PptxGenJS
flat rich arrays, recursive table-cell arrays, permissive color strings,
implicit-inch numbers, and asymmetric margin ordering remain explicit
projection differences instead of being copied into the native API.

## Runtime control and classification

One aggregate PptxGenJS 4.0.1 control locks the exact 14-ID inventory and the
5/9 split. A single deck exercises plain/rich text, bare/structured/rich
table cells, hex/theme color, numeric/percentage coordinates, and scalar/
asymmetric margins. Direct OOXML locks the PptxGenJS four-value margin order;
the corresponding native fixture proves the documented native order.

Supported candidates are `TextProps.text`, plain table-cell text, plain
`addText()` string input, percentage coordinates, and scalar margins. The
remaining nine entries retain the same legal final capability through a
stricter or more explicit native representation.

## Shared lifecycle and artifact proof

One shared two-slide lifecycle probe covers all 14 atoms. It creates plain and
rich text plus a mixed-content table, then verifies exact no-op, invalid-input
isolation, rollback, live edits, duplicate or source isolation, write/reopen,
owner-bound OOXML, relationship stability, and zero error/warning diagnostics.

The same probe runs against the actual npm tarball and its browser conditional
export in persistent Chrome. Both retained decks receive tarball-CLI inspect,
PowerPoint 2010 validation, slide/part reads, and decompressed package-part
comparison. Existing package/browser anchors are reused where they already
prove the state; the probe adds only the missing aggregate lifecycle boundary.

## Verification and matrix result

The family receives one focused gate, one full Vitest/TypeScript gate, one
runtime-control gate, deterministic audit generation, and one artifact batch.
Expected matrix result: 1,734/1,774 classified, 40 unverified, 97.75%, with
unsupported, stale, and diagnostics remaining zero.
