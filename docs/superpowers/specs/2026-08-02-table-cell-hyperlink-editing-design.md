# Table-Cell Hyperlink Editing and Clearing Design

## Goal

Add lossless editing and clearing for supported plain single-run table-cell hyperlinks without expanding into rich or multi-paragraph table text. The public operation is:

```ts
table.setCellHyperlink(rowIndex, columnIndex, value);
```

`value` is `Hyperlink | undefined`; `undefined` clears the direct cell-run click. `TableCell.hyperlink` remains a readonly detached snapshot, and no table-level hyperlink default or bulk editor is added.

## Alternatives

1. Add `TableModel.setCellHyperlink()`. This is selected because it matches `setCellText()`, `setCellFill()`, and the other physical-cell editors while keeping snapshot objects immutable.
2. Make `TableCell.hyperlink` assignable. Rejected because `TableCell` is a detached value, not a stable live model, and mutation would create misleading ownership and stale-index behavior.
3. Re-render the whole table from semantic rows. Rejected because it would discard unsupported table XML, relationship IDs, imported run properties, and exact no-op behavior.

## Supported Ownership Boundary

Rows and cells use zero-based physical indexes. Editing first resolves one exact direct table structure and the selected physical cell. The cell must have exactly one namespace-correct direct `txBody`, one direct paragraph, one direct run, one direct `rPr`, and one direct text node under the same structural rules as `TableCell.hyperlink` reading.

An unlinked supported cell is editable. A supported direct `hlinkClick` is editable only when its relationship, target mode/type, action, and current slide target decode unambiguously. A present but malformed, dangling, duplicate, wrong-namespace, wrong-type, or unsupported-action click makes the cell unsafe and raises `ModelParseError` before mutation. Rich, multi-run, multi-paragraph, field, break, or ambiguous structures remain unsupported.

## Public Semantics

`setCellHyperlink()` performs strict `Hyperlink` normalization before entering the OPC transaction. URL and slide remain mutually exclusive; tooltips preserve property absence versus explicit empty; internal slide targets must resolve before any relationship or XML mutation.

The method provides these states:

- Equal current and requested values are exact part/relationship/journal no-ops.
- Changing only tooltip omission/value patches the direct click locally and keeps its relationship ID.
- Changing URL to URL, URL to slide, slide to URL, or slide target updates an unshared relationship in place.
- When the current relationship ID is referenced elsewhere, target changes allocate a new relationship and update only the selected cell.
- Adding to an unlinked supported cell creates one relationship and one direct run click.
- Clearing removes only the selected direct click and garbage-collects its relationship only when no references remain.

Every native-created cell keeps independent relationship ownership. Imported shared IDs are handled with clone-on-write rather than normalized globally.

## XML Editing

Extend the focused table-cell hyperlink internal module with strict editable-state inspection and local replacement. Reuse the existing hyperlink normalizer, decoder, renderer, equality rules, relationship reference counter, and `patchExistingHyperlink()` behavior through a new text-run replacement helper in the hyperlink codec.

Insertion places the direct `a:hlinkClick` before direct mouse-over/extension anchors when present, otherwise at the end of `rPr`. It declares the relationship namespace locally when no effective declaration exists. If the run has no direct underline attribute, insertion adds `u="sng"`; an existing underline value is preserved. Target/tooltip edits preserve unrelated `rPr` attributes and children plus unowned click attributes/children, while canonicalizing the owned relationship ID, tooltip, and internal-slide action fields.

Clear removes only `hlinkClick`. It deliberately preserves the direct underline and every other run property because existing decks do not expose whether underline was hyperlink-generated or independently authored. Target-slide deletion retains the existing cleanup behavior and likewise does not infer style ownership.

## Transaction and Relationship Lifecycle

`SlideModel` owns the relationship mutation because it has the presentation, slide part, and OPC transaction. `TableModel.setCellHyperlink()` delegates with the table shape ID and physical indexes.

The operation validates input and the internal target first, then executes cell resolution, current-state decoding, relationship update/addition, local XML patching, slide write, and orphan cleanup in one synchronous transaction. Any parser, allocator, relationship, XML, or slide-write failure restores parts, relationships, ZIP entry dates, and the journal.

Existing duplicate, slide move, target deletion, six-format, write/reopen, and live table identity behavior continues to use part-target identity. Duplicate self-links retarget to the duplicate through the existing lifecycle codec.

## PptxGenJS Boundary

PptxGenJS 4.0.1 has table-cell hyperlink creation but no existing-deck editing API. Native imports its legal URL/internal-slide cell clicks, including omitted-tooltip materialization and extra run-click attributes, then edits the supported final direct state.

Native does not copy caller `_rId` mutation or loose runtime coercion. Tooltip edits may canonicalize the owned empty external `action` field while preserving unrelated compatible attributes such as history, target frame, highlight, and sound/extension content. This is an existing-deck lossless extension, not a claim that PptxGenJS exposes equivalent editing calls.

## Verification

Focused tests cover:

- add, URL/slide switch, tooltip absent/empty/value changes, and clear;
- exact no-op and absent clear;
- unique relationship ID reuse and shared clone-on-write;
- equal-target independent native cells;
- clear/replace reference-aware garbage collection;
- direct underline insertion and style preservation on clear;
- PptxGenJS extra-attribute import/editing;
- unsafe structural/relationship state rejection;
- invalid input/target and injected relationship/XML failures with zero observable mutation;
- setCellText/property-edit coexistence, duplicate, reorder, target deletion, self-link, all six formats, and write/reopen;
- model, SDK, root package, TypeScript, actual tarball, browser conditional export, CLI inspection, `pptx-inspect`, and real Chrome proof.

Final gates remain focused Vitest, TypeScript checks, Node/browser builds and declarations, full Vitest with two workers, independent performance, actual package inspection, PowerPoint 2010 validation, and zero Chrome validation/console/page/network errors.

## Out of Scope

- Table-level hyperlink defaults or bulk editing.
- Rich, multi-run, nested, or multi-paragraph cell hyperlinks.
- Hover, action-only, custom-show, macro, sound, history/highlight, or hyperlink-color editing.
- Automatic removal of underline or other run formatting on clear.
- Relative/file URL safety policy changes.
- Table graphic-frame, image, chart, media, group, or other owner hyperlinks.
