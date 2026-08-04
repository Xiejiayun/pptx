# PptxGenJS addTable Core Surface Family Design

## Context

After the `tableToSlides` family, the PptxGenJS 4.0.1 declaration audit has 45
unverified atoms that form the ordinary editable-table core. The native SDK
already creates, imports, edits, duplicates, rolls back, serializes, validates,
and reopens tables across Node, the packed npm artifact, and real Chrome. This
batch closes the declaration inventory without adding a second table model or
duplicating the existing lifecycle probes.

The standing workflow is capability-family batching with three read-only
evidence agents and main-agent-only repository edits. Expensive artifact gates
run once at the family boundary, followed by one review, commit, and push.

## Batch Boundary

This batch contains exactly 45 atoms:

- `method:Slide#addTable`;
- the 20 still-unverified direct properties owned by `TableProps`;
- the ten still-unverified direct properties owned by `TableCellProps`;
- the three direct properties owned by `BorderProps`;
- three `BorderProps.type` union atoms;
- two `TableCellProps.border` union atoms;
- two `TableProps.border` union atoms;
- two `TableProps.colW` union atoms;
- two `TableProps.rowH` union atoms.

It excludes inherited text, fill, underline, tab-stop, bullet, and text-direction
atoms that earlier families already closed. It also excludes
`PptxGenJS.tableToSlides`, whose 42-atom family is already committed.

## Locked Classification

### Supported: 13

- `TableCellProps.align`
- `TableCellProps.colspan`
- `TableCellProps.hyperlink`
- `TableCellProps.rowspan`
- `TableCellProps.transparency`
- `TableCellProps.valign`
- `TableProps.align`
- `TableProps.autoPage`
- `TableProps.autoPageCharWeight`
- `TableProps.autoPageHeaderRows`
- `TableProps.autoPageLineWeight`
- `TableProps.autoPageRepeatHeader`
- `TableProps.valign`

These declarations have the same effective legal behavior and public names in
the native SDK. Native validation, detachment, editing, and rollback strengthen
the contract without changing the legal capability.

### Deliberate differences: 28

- `method:Slide#addTable`;
- `BorderProps.color`, `BorderProps.pt`, and `BorderProps.type`;
- the `dash`, `none`, and `solid` members of `BorderProps.type`;
- `TableCellProps.border` and `TableCellProps.margin`;
- both direct union branches of `TableCellProps.border`;
- `TableProps.autoPageSlideStartY`, `border`, `colW`, `h`, `margin`,
  `objectName`, `rowH`, `verbose`, `w`, `x`, and `y`;
- both direct union branches of `TableProps.border`, `TableProps.colW`, and
  `TableProps.rowH`.

PptxGenJS returns the owning `Slide`, mutates caller rows/options, uses implicit
inch coordinates, mixes inch and point margin rules, and represents borders as
`type/color/pt`. Native returns the editable `TableModel`, snapshots input,
rolls back atomically, requires explicit EMU or `inches()` geometry, names
`w/h/colW/rowH/objectName` as `width/height/columnWidths/rowHeights/name`, and
uses strict `kind/color/width/style` border state. The three border union tokens
remain deliberate differences because their native paths are `kind` or `style`,
matching the existing line and fill family classification rule.

Cell-level `transparency` is supported through native cell rich text because it
preserves the same legal color/alpha final state, matching the existing
TableCell scalar-text classification rule. `verbose` is a process-global console
side effect with no deck-state effect; native intentionally keeps table creation
free of global logging.

### Deprecated alias: 1

- `TableProps.newSlideStartY` → `TableProps.autoPageSlideStartY`

The control must prove that the old field still sets continuation-table Y in
PptxGenJS 4.0.1 and that the canonical target is closed in the same family.

### Defect exclusions: 3

- `TableCellProps.autoPageCharWeight`
- `TableCellProps.autoPageLineWeight`
- `TableProps.transparency`

PptxGenJS overwrites every caller cell `autoPageCharWeight` with the table-level
value or `null` before measuring content. It never reads cell-local
`autoPageLineWeight`; only the table-level value enters line-height math. It also
does not propagate table-level `transparency` to cells or rich-text runs. Native
implements strict intended cell weight overrides and rich-text transparency,
but those corrections do not turn the broken upstream declarations into
delivered PptxGenJS capabilities.

## Evidence Architecture

- Adapter: reuse existing basic creation, auto-page, weight, hyperlink, merge,
  alignment, vertical alignment, margin, and border controls. Add one aggregate
  control named `locks the remaining addTable core declarations against
  PptxGenJS 4.0.1` for the exact 45-ID inventory, method return/mutation,
  `objectName`, `verbose`, alias behavior, table transparency, and both inert
  cell weight declarations.
- Native: reuse `packages/sdk/src/index.test.ts` table creation, auto-page,
  hyperlink, merge, margin, border, alignment, all-format, edit, duplicate,
  rollback, and reopen tests.
- npm: reuse the existing table creation, auto-page, content measurement,
  merge, hyperlink, border, margin, horizontal alignment, and vertical
  alignment states in `scripts/smoke-npm-package.mjs`.
- Browser: reuse the matching persistent-Chrome states in
  `scripts/playwright-browser-smoke.js` and require zero validation, console,
  page, and network errors.
- OOXML: retain one packed table artifact and inspect table grid, row height,
  merge, hyperlink relationship, borders, margins, anchors, and editable reopen
  state with `pptx-inspect` at the family boundary.
- Matrix: generate all 45 entries from frozen owner/property/union tables and
  assert exact family and global counts.

Only the main agent edits repository files. Evidence agents remain read-only.

## Acceptance Criteria

- Exactly 45 atoms move from `unverified`: 13 supported, 28
  deliberate-difference, one deprecated-alias, and three defect-excluded.
- Global totals become 588 supported, 240 deliberate-difference, 86
  deprecated-alias, 352 defect-excluded, and 508 unverified: 1,266 of 1,774
  atoms closed (71.36%).
- The aggregate upstream control locks the three defect boundaries and the
  method, logging, naming, alias, and caller-mutation differences.
- Existing focused table tests, one packed tarball, one persistent Chrome run,
  OOXML inspection, TypeScript, audit tests, and full Vitest all pass.
- Review finds no unrelated staged path; one capability-family commit is pushed
  and local/remote divergence is `0 0`.
