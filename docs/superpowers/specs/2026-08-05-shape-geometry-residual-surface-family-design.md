# Shape Geometry Residual 3 Surface Family Design

## Scope

Close exactly three remaining `ShapeProps` atoms: inert `align`, plus the
`angleRange` and `arcThicknessRatio` geometry shortcuts.

## Classification

- `align`: `defect-excluded`; PptxGenJS accepts it on `addShape()` but emits no
  text body or alignment state.
- `angleRange`: `deliberate-difference`; native exposes the same final
  `adj1`/`adj2` guides through ordered strict `ShapeAdjustment[]` state.
- `arcThicknessRatio`: `deliberate-difference`; native exposes the same final
  `adj3` guide without the permissive shortcut alias.

No product implementation is required. The family adds one aggregate runtime
control and reuses existing native adjustment lifecycle, npm, Chrome, OOXML,
six-format, and PowerPoint 2010 evidence.

## Verification

The aggregate locks exact classification, inert align owner/relationship state,
PptxGenJS shortcut output, and native direct-guide equivalence. Audit generation
must be byte-stable at 1,751 classified / 23 unverified with zero diagnostics.
The family is reviewed, committed, and pushed as one unit.
