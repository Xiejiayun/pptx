# PptxGenJS Text Paragraph/Layout Surface Family Design

Date: 2026-08-05

Status: confirmed under the user's autonomous-decision authorization

## Context and Batch Boundary

The declaration audit has ten unverified `TextPropsOptions` atoms for the
paragraph and text-box layout surface: `align`, `rtlMode`, `lineSpacing`,
`lineSpacingMultiple`, `paraSpaceBefore`, `paraSpaceAfter`, `margin`, the
deprecated `inset` spelling, `valign`, and `wrap`. Native already implements
their create, read, edit, clear, duplicate, transaction, write, and reopen
intent through structured rich-text paragraphs and text-box properties. This
batch binds that existing vertical slice to the compatibility matrix and adds
the missing aggregate control and real-Chrome gates.

The batch is owner-exact. It excludes `PlaceholderProps.margin`, the standalone
`Margin` union atoms, table owners, text fit, text direction, bullets, tab
stops, run effects, shape geometry, transforms, and hyperlinks. Those remain
separate capability families.

## Classification

Four atoms are direct support:

- `interface:TextPropsOptions@property:align` maps to paragraph `align` with the
  same left, center, right, and justify tokens;
- `interface:TextPropsOptions@property:rtlMode` maps to paragraph `rtl`;
- `interface:TextPropsOptions@property:valign` maps to
  `ShapeModel.verticalAlignment` with the same top, middle, and bottom tokens;
- `interface:TextPropsOptions@property:wrap` maps to `ShapeModel.textWrap` with
  the same boolean intent.

Five atoms are deliberate API-shape differences:

- `lineSpacing` and `lineSpacingMultiple` map to the mutually exclusive
  `ParagraphSpacing.line` exact/multiple union;
- `paraSpaceBefore` and `paraSpaceAfter` map to `ParagraphSpacing.before` and
  `ParagraphSpacing.after`;
- `margin` maps to `ShapeModel.textMargins`, where native keeps the documented
  top/right/bottom/left tuple order and strict validation instead of copying
  PptxGenJS 4.0.1's asymmetric tuple reorder.

The resulting OOXML intent is preserved for legal values: `spcPts`, `spcPct`,
`spcBef`, `spcAft`, and the four `bodyPr` inset attributes. Native rejects
malformed values before mutation and exposes reversible editing through the
structured properties.

`interface:TextPropsOptions@property:inset` is a deprecated alias whose canonical
target is `interface:TextPropsOptions@property:margin`. PptxGenJS 4.0.1 interprets
`inset` as inches, applies it uniformly to all four sides, and lets explicit
`margin` win when both are present. Native exposes only the canonical margin
surface in points.

## Evidence and Pipeline

Agent A owns the exact matrix slice and totals, Agent B owns the locked
PptxGenJS 4.0.1 runtime/control comparison, and Agent C owns native,
actual-tarball, browser, and OOXML evidence. All three lines are read-only;
the main agent is the only writer and owns review, artifact regeneration,
commit, push, and progress reporting.

Evidence is aggregated once per family:

- the existing PptxGenJS control already covers align, RTL, exact/multiple
  spacing, paragraph spacing, margin, vertical alignment, wrapping, OOXML,
  and reopen; add only the missing `inset` precedence and owner probes;
- native SDK coverage proves strict lifecycle behavior and exact OOXML for
  paragraph alignment, RTL, spacing, margins, vertical alignment, and wrap;
- the existing actual npm smoke proves installed-package types and runtime
  lifecycle for all canonical mappings;
- one real-Chrome state proves immediate values, exact OOXML, live edits,
  invalid rollback, duplicate isolation, write/reopen, and diagnostics.

## Acceptance Criteria

- Exactly ten atoms leave `unverified`: four become `supported`, five become
  `deliberate-difference`, and one becomes `deprecated-alias`; no other atom
  changes status.
- Totals become 628 supported, 359 deliberate differences, 94 deprecated
  aliases, 359 defect exclusions, and 334 unverified: 1,440 of 1,774 atoms
  classified (81.17%).
- Manifest scope/count assertions, focused adapter/SDK tests, deterministic
  artifact generation, TypeScript, actual npm tarball, real Chrome, OOXML,
  and the batch full suite pass.
- Review produces one capability-family commit and push with no unrelated
  tracked files.
