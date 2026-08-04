# PptxGenJS tableToSlides Surface Family Design

## Context

The PptxGenJS 4.0.1 declaration audit has 42 unverified atoms reachable from
`PptxGenJS.tableToSlides()` and `TableToSlidesProps`. The native SDK already
implements browser DOM/CSS capture, deterministic column resolution, pagination,
master/layout inheritance, repeated headers, four page additions, transactional
rollback, editable output, all six formats, packed-package use, and real Chrome
reopen. This batch closes the declaration atoms with one direct upstream control
and reuses those existing lifecycle artifacts.

The standing project direction approves capability-family batching, read-only
evidence agents, main-agent-only repository edits, and one reviewed commit and
push per family.

## Batch Boundary

This batch contains exactly:

- `class:PptxGenJS#tableToSlides`;
- 27 `interface:TableToSlidesProps@property:*` atoms;
- eight nested atoms under `addImage`, `addShape`, `addTable`, and `addText`;
- six union atoms belonging to `border`, `colW`, and `rowH`.

It excludes the already closed inherited scalar-text, fill, underline, tab-stop,
and text-direction atoms. It also excludes ordinary `Slide.addTable()` paging,
which has separate evidence and declaration owners.

## Locked Classification

### Supported: 4

- `autoPageCharWeight`
- `autoPageLineWeight`
- `autoPageRepeatHeader`
- `masterSlideName`

These names and legal final-state effects map directly to native
`TableToSlidesOptions`.

### Deliberate differences: 20

- `class:PptxGenJS#tableToSlides`;
- top-level `addImage`, `addShape`, `addTable`, `addText`;
- the eight nested addition fields;
- `autoPageSlideStartY`, `h`, `slideMargin`, `w`, `x`, `y`;
- `verbose`.

Native returns `Promise<readonly SlideModel[]>`, freezes the result, snapshots
inputs, and rolls back atomically instead of synchronously returning `undefined`
while mutating caller objects. Addition records use `source`, `type`, native row
and rich-text types, strict geometry, and explicit EMU or `inches()` units.
Native maps `w`/`h` to exact `width`/`height` semantics and preserves legal zero
coordinates that PptxGenJS truthy fallbacks discard. PptxGenJS `verbose` writes
process-global console traces without changing the deck; native deliberately
keeps document creation free of global logging.

### Deprecated aliases: 2

- `addHeaderToEach` → `autoPageRepeatHeader`
- `newSlideStartY` → `autoPageSlideStartY`

The aggregate control must prove both aliases remain active in PptxGenJS 4.0.1
and that each canonical target is closed in the same family.

### Defect exclusions: 16

- properties `align`, `autoPage`, `autoPageHeaderRows`, `border`, `colW`,
  `margin`, `objectName`, `rowH`, `transparency`, and `valign`;
- both declared union branches for each of `border`, `colW`, and `rowH`.

PptxGenJS `tableToSlides()` always paginates even when `autoPage` is false,
ignores `autoPageHeaderRows` and repeats all `<thead>` rows, overwrites caller
`colW` from DOM widths, never forwards `rowH`, and derives alignment, borders,
and vertical alignment from computed CSS instead of inherited flat fields.
`margin` participates only in an inconsistent pagination measurement fallback
and does not produce the declared cell-margin state. Native exposes strict
corrections for several of these cases, but native extensions do not turn an
upstream inert or miswired declaration into a supported PptxGenJS capability.

## Evidence Architecture

- Adapter: add one aggregate test named `locks every declared tableToSlides option and nested addition against PptxGenJS 4.0.1`.
- Native: reuse `packages/sdk/src/table-to-slides.test.ts` for normalization,
  pagination, geometry, layouts, additions, rollback, six formats, and reopen.
- npm: reuse `tableToSlidesState` and its CLI package inspection in
  `scripts/smoke-npm-package.mjs`.
- Browser: reuse the persistent-Chrome `tableToSlidesState` in
  `scripts/playwright-browser-smoke.js`.
- OOXML: reuse exact grid, CSS style, header, addition, relationship, layout,
  section, editing, and PowerPoint 2010 validation assertions from the packed
  artifact; run `pptx-inspect` once at the family boundary.
- Matrix: generate all 42 entries from frozen property, nested-field, and union
  tables and assert both family and global counts.

Only the main agent edits repository files. Evidence agents remain read-only.

## Acceptance Criteria

- Exactly 42 atoms move from `unverified`: 4 supported, 20
  deliberate-difference, 2 deprecated-alias, and 16 defect-excluded.
- Global totals become 575 supported, 212 deliberate-difference, 85
  deprecated-alias, 349 defect-excluded, and 553 unverified: 1,221 of 1,774
  atoms closed (68.83%).
- The aggregate upstream control locks every target property, nested field,
  union branch, deprecated alias, output-relevant effect, and defect boundary.
- Existing native focused tests, packed tarball, persistent Chrome, OOXML
  inspection, TypeScript, audit tests, and full Vitest all pass.
- Review finds no unrelated staged path; one capability-family commit is pushed
  and local/remote divergence is `0 0`.
