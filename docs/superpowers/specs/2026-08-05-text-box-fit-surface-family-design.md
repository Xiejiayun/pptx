# PptxGenJS Text Box Fit Surface Family Design

Date: 2026-08-05

Status: confirmed under the user's autonomous-decision authorization

## Context and Batch Boundary

The declaration audit has six unverified `TextPropsOptions` atoms for text-box
fit: the `fit` property, its `none`, `resize`, and `shrink` union members, and
the deprecated `autoFit` and `shrinkText` booleans. Native already implements
strict create, read, edit, clear, duplicate, write, and reopen behavior through
`AddTextOptions.fit`, `ShapeModel.textFit`, and `TextBoxFit`. This batch binds
that existing vertical slice to the compatibility matrix and adds the missing
real-Chrome aggregate gate.

The batch excludes table-cell fit, text overflow calculation, inherited
placeholder fit, wrapping, margins, alignment, direction, paragraph spacing,
and run formatting. Those remain separate capability families.

## Classification

The modern family is direct support:

- `interface:TextPropsOptions@property:fit`;
- `union:interface:TextPropsOptions@property:fit#none`;
- `union:interface:TextPropsOptions@property:fit#resize`;
- `union:interface:TextPropsOptions@property:fit#shrink`.

PptxGenJS and native use the same legal input tokens and final OOXML intent:
omitted and `none` write no autofit child, `shrink` writes
`a:normAutofit`, and `resize` writes `a:spAutoFit`. Native additionally rejects
invalid values before mutation and exposes reversible live editing; those are
strictness improvements, not a reduction or renaming of the legal capability.

The two legacy atoms are deprecated aliases whose canonical target is
`interface:TextPropsOptions@property:fit`:

- `autoFit: true` means `fit: 'resize'`;
- `shrinkText: true` means `fit: 'shrink'`.

Native intentionally exposes only the canonical union and does not copy the
deprecated booleans.

## Evidence and Pipeline

Agent A owns the exact matrix slice and total calculation, Agent B owns the
locked PptxGenJS 4.0.1 runtime/control comparison, and Agent C owns native,
actual-tarball, browser, and OOXML evidence. All three lines are read-only;
the main agent is the only writer and owns review, artifact regeneration,
commit, push, and progress reporting.

Evidence is aggregated once per family:

- PptxGenJS control covers omitted, all three modern tokens, invalid input,
  both legacy aliases, and inert run-level values;
- native SDK coverage proves create, edit, duplicate, rollback, exact OOXML,
  write, reopen, and invalid-input isolation;
- actual npm coverage proves installed-package types and runtime lifecycle;
- one real-Chrome state proves omitted/none/shrink/resize, live edits, invalid
  rollback, duplicate isolation, `writeBlob()`/reopen, MIME, and diagnostics.

## Acceptance Criteria

- Exactly six atoms leave `unverified`: four become `supported` and two become
  `deprecated-alias`; no other atom changes status.
- Totals become 624 supported, 354 deliberate differences, 93 deprecated
  aliases, 359 defect exclusions, and 344 unverified: 1,430 of 1,774 atoms
  classified (80.61%).
- Manifest scope/count assertions, focused adapter/SDK tests, deterministic
  artifact generation, TypeScript, the real-Chrome batch gate, and exact
  migration review pass.
- Review produces one capability-family commit and push with no unrelated
  tracked files.
