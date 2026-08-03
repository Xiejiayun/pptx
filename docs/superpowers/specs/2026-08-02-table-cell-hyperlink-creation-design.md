# Table-Cell Hyperlink Creation Design

## Status

Validated under the standing direction to complete PptxGenJS parity one independently reviewed, committed, and pushed item at a time. Product and implementation choices are delegated, so this design proceeds without a separate decision stop. It follows the completed table creation properties and the existing shape/text-run hyperlink ownership model.

## Goal

Add strict native hyperlink creation for one plain single-paragraph table cell, plus a readonly direct-state snapshot, by extending `AddTableCellOptions.hyperlink` and `TableCell.hyperlink` with the existing public `Hyperlink` value.

## Current context

The library already supports:

- strict descriptor-safe `Hyperlink` values with mutually exclusive external `{ url, tooltip? }` and internal `{ slide, tooltip? }` branches;
- non-empty XML-safe URL validation, positive safe one-based slide numbers, and exact optional-tooltip presence;
- external hyperlink and internal slide relationships on slide, layout, and master parts;
- direct text-run `a:rPr/a:hlinkClick` rendering and reading, including automatic single underline when no direct underline is supplied;
- package-atomic relationship creation, target-part identity across slide reorder, relationship-aware slide duplication, incoming-link cleanup on slide deletion, and write/reopen preservation;
- plain single-paragraph table cells with one direct DrawingML run and descriptor-safe creation options;
- table-cell text, alignments, direction, fit, margins, borders, fill, geometry, duplication, and all six presentation formats;
- PptxGenJS 4.0.1 table-cell `hyperlink` creation for URL and slide targets.

The remaining focused gap is that native table-cell options reject `hyperlink`, table creation does not allocate cell-run relationships, and `TableCell` does not expose a supported cell-run hyperlink snapshot.

## Considered approaches

### 1. One relationship and one direct run hyperlink per linked cell — selected

Normalize each cell hyperlink with the existing codec, validate all internal targets before relationship mutation, allocate one relationship per linked physical cell inside the table-creation transaction, and render the click on that cell's single run. Expose the same direct link through `TableCell.hyperlink` when the cell has exactly one supported paragraph/run/property path.

This matches PptxGenJS's legal final cell output, preserves independent relationship ownership, and reuses the established native hyperlink rules without introducing another public value model.

### 2. Reuse one relationship for cells with the same target

Relationship reuse would reduce relationship count. It would also couple later clone-on-write, replacement, and garbage collection across unrelated cells, while PptxGenJS and the current rich-text run creator allocate independent run relationships. This approach is rejected.

### 3. Put the hyperlink on the table graphic frame

A non-visual click on `p:cNvPr` would make the whole table clickable rather than the selected cell text. It would not import as a cell-run hyperlink and would conflict with future graphic-frame hyperlink work. This approach is rejected.

No generic relationship planner is added. The existing table normalizer, relationship input type, and rich-text hyperlink renderer already provide the smallest safe path.

## Public API

The existing public types gain one optional property each:

```ts
interface AddTableCellOptions {
  readonly hyperlink?: Hyperlink;
}

interface TableCell {
  readonly hyperlink?: Hyperlink;
}
```

Usage:

```ts
const document = PptxDocument.create();
const source = document.addSlide();
document.addSlide();

const table = source.addTable([[
  {
    text: 'Open documentation',
    options: {
      hyperlink: {
        url: 'https://example.com/docs',
        tooltip: 'Open the documentation',
      },
    },
  },
  {
    text: 'Go to details',
    options: { hyperlink: { slide: 2 } },
  },
]]);

table.rows[0]!.cells[0]!.hyperlink;
// { url: 'https://example.com/docs', tooltip: 'Open the documentation' }
table.rows[0]!.cells[1]!.hyperlink;
// { slide: 2 }
```

There is no table-level hyperlink default, `TableModel.hyperlink` alias, target string shorthand, mixed sentinel, or cell hyperlink setter in this item. Cell hyperlink editing remains a separate independently reviewed slice.

## Normalization and validation

`AddTableCellOptions` adds only the exact `hyperlink` key. Omitted and runtime-`undefined` values create no link. A present value is normalized with `normalizeHyperlink(value, "Table cell R,C hyperlink")` while the complete table definition is still being detached.

The accepted input remains exactly one ordinary or null-prototype object with one target:

- `{ url: string, tooltip?: string }`, where URL is non-empty and XML-safe;
- `{ slide: number, tooltip?: string }`, where slide is a positive safe one-based integer;
- an absent tooltip remains absent, while `tooltip: ''` remains an explicit empty attribute.

Class instances, arrays, accessors, inherited/symbol/unknown properties, both target branches, neither target branch, empty or non-string URLs, non-string tooltips, invalid XML controls, non-numeric/fractional/unsafe/non-positive slide numbers, and out-of-range current-presentation slide targets are rejected. All cells and all internal targets are validated before the first relationship is added. Input objects are detached and never receive PptxGenJS-style `_rId` mutation.

## Relationship preparation and rendering

The normalized table definition retains an optional frozen normalized hyperlink on each cell. `SlideModel.addTable()` performs two phases inside its existing OPC transaction:

1. walk the resolved row/cell matrix and prepare an equally shaped matrix of optional `RelationshipInput` values, resolving internal one-based slide numbers to current target part URIs without package mutation;
2. after the full matrix is valid, add one relationship for each linked cell and produce an equally shaped optional relationship-ID matrix.

External cells use the standard hyperlink relationship type with `TargetMode="External"`. Internal cells use the standard slide relationship type with a relative internal target. Equal targets intentionally receive independent relationships.

`renderTableGraphicFrame()` receives the relationship-ID matrix. Each linked cell calls the existing rich-text paragraph renderer with its normalized hyperlink and matching ID. The output remains one paragraph and one run, with the click under `a:rPr`; the renderer supplies `u="sng"` because the current plain cell API has no explicit run underline. A relationship namespace is declared on the graphic-frame fragment only when at least one cell is linked.

Hyperlink and relationship ID presence must match for every cell, and matrix dimensions must match the normalized row/cell dimensions. Any mismatch is an internal `TypeError`. If relationship creation, rendering, placeholder replacement, shape resolution, or package validation later fails, the enclosing transaction rolls back all slide XML, relationships, diagnostics, identities, and journal changes.

## Read snapshot

`TableCell.hyperlink` reads only a supported direct plain-cell path:

- one direct DrawingML `txBody`;
- one direct DrawingML paragraph;
- one direct DrawingML run;
- one direct DrawingML run-properties element;
- one safe supported direct `hlinkClick` decoded by the existing text-run hyperlink reader.

A legal external link returns its relationship URL. A legal internal link resolves the target part URI to the current unique one-based presentation slide number. The returned `Hyperlink` is detached and frozen. Missing clicks, dangling/wrong-type/ambiguous relationships, unsupported actions, malformed properties, multiple paragraphs/runs, fields, or other unsupported cell text structures return `undefined` rather than guessing.

The snapshot represents only the direct single-run link. It does not synthesize links from a table style, table graphic frame, paragraph end properties, neighboring cells, or retained creation metadata. Existing `TableCell.text` and per-cell property snapshots remain unchanged.

## Lifecycle and preservation

Cell hyperlinks are ordinary DrawingML run hyperlinks after creation, so existing package lifecycle rules apply:

- external URLs remain external relationships and produce the existing portability warning, not a package error;
- internal links retain target-part identity when slides are inserted, moved, or reordered, while the readonly snapshot reports the new one-based ordinal;
- slide duplication clones the source slide relationship set, preserves external targets, and retargets duplicate self-links to the duplicate through the existing duplication rules;
- deleting an internal target removes incoming DrawingML click elements and their now-unreferenced relationships before deleting the target part;
- `setCellText()` changes only the first direct text node and preserves the run click and relationship;
- border, fill, margin, alignment, direction, fit, geometry, table-level bulk edits, and unrelated shape edits preserve the cell click and relationship;
- write/reopen and all six presentation formats preserve supported external and internal snapshots.

Deletion cleanup deliberately preserves unrelated run formatting, including the direct underline that was emitted with the original linked run. Hover links, sounds, action-only navigation, and custom hyperlink history state are not created or edited.

## PptxGenJS conformance boundary

PptxGenJS 4.0.1 legal plain table-cell URL output writes a run click, a hyperlink relationship, and `u="sng"`; legal slide output writes a slide relationship plus `action="ppaction://hlinksldjump"`. Native creation produces the same final target, action, tooltip value, and underlined cell-run behavior and imports those legal outputs through `TableCell.hyperlink`.

Documented strict differences remain:

- PptxGenJS materializes an omitted tooltip as a direct empty attribute; native preserves tooltip absence and preserves explicit empty separately;
- PptxGenJS URL clicks include empty/default history, frame, sound, and invalid-URL attributes; native emits only the relationship ID plus requested tooltip/action;
- PptxGenJS mutates caller hyperlink objects with `_rId`, may prefer one truthy target when both are present, and may log, ignore, coerce, or leave defective state for invalid runtime values; native detaches input and rejects invalid values before relationship mutation;
- PptxGenJS supports nested/rich table-cell text forms; native rich/multi-paragraph table cells and per-run cell links remain outside this plain-cell item.

Native readonly import and package lifecycle behavior are lossless extensions because PptxGenJS has no existing-deck editor.

## Test design

### Internal and model tests

- Normalize valid URL, slide, absent tooltip, empty tooltip, runtime-`undefined`, null-prototype, and detached input values.
- Reject unknown keys, accessors, inherited/symbol keys, class instances, both/neither targets, invalid strings, invalid slide numbers, and out-of-range internal targets before any relationship is added.
- Render exact one-run URL and slide clicks with matching independent relationships, conditional relationship namespace, XML escaping, automatic underline, empty text, and unchanged unlinked-cell bytes.
- Require matrix dimension/presence mismatch failures in the internal renderer.
- Read legal native and PptxGenJS external/internal clicks, absent versus empty tooltip, current slide ordinal, and frozen detached results.
- Return `undefined` for missing, dangling, wrong-type, ambiguous, malformed, multi-paragraph/run, unsupported-action, field, and descendant-lookalike state.
- Require full transaction rollback after a late render/shape failure, stable table identity, input detachment, exact unrelated-state preservation, `setCellText()` preservation, source/duplicate isolation, slide move/reorder, duplicate self-link retargeting, target deletion cleanup, write/reopen, and all six formats.
- Require public typing to accept `Hyperlink` on `AddTableCellOptions`, expose `Hyperlink | undefined` on `TableCell`, and reject aliases or invalid branches.
- Re-run completed table creation/property and shape/rich-text hyperlink suites.

### PptxGenJS evidence

- Generate real PptxGenJS 4.0.1 plain table cells with URL and internal slide links and import their exact final relationships/clicks.
- Require equivalent native final target, action, tooltip, underline, and readonly snapshot after write/reopen.
- Record PptxGenJS's explicit-empty omitted tooltip, extra URL click attributes, and `_rId` caller mutation without copying them into native behavior.
- Keep nested/rich cell runs outside the parity claim.

### Package and browser evidence

- Require generated declarations to expose `AddTableCellOptions.hyperlink?: Hyperlink` and `TableCell.hyperlink?: Hyperlink` from the actual tarball.
- Extend actual-tarball Node, TypeScript, browser-condition, installed CLI, `pptx-inspect`, and real-Google-Chrome smoke with stable `tableCellHyperlinks: true` state.
- Cover URL/slide creation, absent/empty tooltip, input detachment, independent relationships, readonly snapshots, text-edit preservation, reorder, reopen, and invalid failure isolation.
- Inspect packed slide XML and relationships, require exact cell-run clicks and targets, and validate with the PowerPoint 2010 profile before narrow part reads.
- Keep validation, console, page, and network error counts at zero.

## Delivery boundaries

The work is delivered as separately reviewed and pushed commits:

1. this design;
2. the implementation plan;
3. core creation/snapshot implementation, source tests, and PptxGenJS conformance;
4. actual-package declarations, Node/browser/CLI, and real-Chrome evidence;
5. public documentation, compatibility matrix, changelog, and progress closeout.

Each commit stages only its declared files. `.pnpm-store/`, generated workspace tarballs, and retained `/tmp` evidence are never committed.

## Non-goals

- Table-level hyperlink defaults or table graphic-frame hyperlinks.
- `TableModel.setCellHyperlink()` or live hyperlink editing/clearing.
- Multiple paragraphs/runs, nested PptxGenJS table-cell arrays, rich cell text, or per-run cell link creation.
- Hover links, sounds, action-only navigation, relative/file safety policy, history/highlight flags, or hyperlink color customization.
- Removing synthesized underline when an incoming internal target is deleted.
- Cell merge, row/column CRUD, auto-page, repeated headers, content measurement, layout recomputation, or `tableToSlides`.
- Reusing relationships between cells or copying PptxGenJS `_rId` mutation and invalid-value coercion.

## Acceptance criteria

The item is complete when `AddTableCellOptions.hyperlink` creates strict independent URL and slide relationships on the cell's one direct text run, `TableCell.hyperlink` reads the supported direct state after creation/import/reopen, invalid input and failures leave zero package mutation, existing cell/table properties and lifecycle behavior remain unchanged, legal PptxGenJS 4.0.1 plain-cell final output imports within the documented boundary, focused/full/performance/type/bundle/declaration gates pass, the actual tarball and real Chrome report `tableCellHyperlinks: true` with zero browser errors, documentation moves plain table-cell hyperlink creation to supported while leaving editing/rich-cell links pending, and every sub-item is reviewed, committed, pushed, and verified at remote divergence `0 0`.
