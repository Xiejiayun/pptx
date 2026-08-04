# PptxGenJS Rich-Text Effects Surface Family Design

Date: 2026-08-05

Status: confirmed under the user's autonomous-decision authorization

## Context and Batch Boundary

The declaration audit has sixteen unverified atoms for run-level rich-text
effects: outline size and color, glow opacity, size, and color, baseline,
superscript, subscript, character spacing, strike, and text transparency.
Native already implements their create, read, edit, clear, duplicate,
transaction, write, and reopen intent through `RichTextRunStyle`. This batch
binds that existing vertical slice to the compatibility matrix and adds the
missing aggregate six-format and real-Chrome gates.

The batch is owner-exact. It excludes shape line/fill effects, underline,
highlight, hyperlinks, paragraph layout, text geometry, table formatting, and
image effects. Those remain separate capability families.

## Classification

Nine atoms are direct support:

- `inline:interface:TextPropsOptions@property:outline@property:outline.size`;
- `interface:TextGlowProps@property:opacity` and `size`;
- `interface:TextPropsOptions@property:glow`, `outline`, `strike`, and
  `transparency`;
- `union:interface:TextPropsOptions@property:strike#sngStrike` and
  `union:interface:TextPropsOptions@property:strike#dblStrike`.

For positive outline sizes, both libraries write the same DrawingML width.
Native also preserves explicit zero; PptxGenJS 4.0.1 replaces zero with its
0.75-point falsy fallback. This is treated as a strict extension of the same
declared size capability rather than a separate API shape.

Seven atoms are deliberate API-shape differences:

- outline and glow colors use the strict structured `RichTextColor` union
  instead of PptxGenJS's permissive color scalar;
- `baseline`, `superscript`, and `subscript` map to the single
  `RichTextRunStyle.baseline` field, which uses direct OOXML percentage units
  and canonical named values instead of PptxGenJS's custom-baseline scaling
  and truthy-zero behavior;
- `charSpacing` maps to the explicit point-based `characterSpacing` field;
- boolean strike is accepted but differs deliberately because native writes
  `strike="noStrike"` for explicit false while PptxGenJS omits false.

Native preserves legal resulting DrawingML intent, rejects malformed values
before mutation, preserves explicit zero values, and exposes reversible live
editing through detached structured run styles.

## Evidence and Pipeline

Agent A owns the exact matrix slice and totals, Agent B owns the locked
PptxGenJS 4.0.1 runtime/control comparison, and Agent C owns native,
actual-tarball, browser, and OOXML evidence. All evidence lines are read-only;
the main agent is the only shared-file writer and owns review, artifact
regeneration, commit, push, and progress reporting.

Evidence is aggregated once per family:

- the existing PptxGenJS aggregate covers outline, glow, baseline,
  superscript, subscript, character spacing, and strike, while the dedicated
  transparency control covers its alpha behavior;
- existing native SDK tests prove strict lifecycle behavior and exact OOXML
  for each effect;
- one new SDK aggregate proves the combined effect state across all six
  presentation formats;
- the existing actual npm smoke proves installed-package types and runtime
  create/edit behavior for all mappings;
- one real-Chrome state proves immediate values, exact OOXML, live edits,
  invalid rollback, duplicate isolation, write/reopen, MIME, and diagnostics.

## Acceptance Criteria

- Exactly sixteen atoms leave `unverified`: nine become `supported` and seven
  become `deliberate-difference`; no other atom changes status.
- Totals become 637 supported, 366 deliberate differences, 94 deprecated
  aliases, 359 defect exclusions, and 318 unverified: 1,456 of 1,774 atoms
  classified (82.07%).
- Manifest scope/count assertions, focused adapter/SDK tests, deterministic
  artifact generation, TypeScript, actual npm tarball, real Chrome, OOXML,
  and the batch full suite pass.
- Review produces one capability-family commit and push with no unrelated
  tracked files.
