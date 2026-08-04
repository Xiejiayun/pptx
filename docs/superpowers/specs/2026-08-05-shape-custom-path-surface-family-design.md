# PptxGenJS Shape Custom Path Surface Family Design

Date: 2026-08-05

Status: confirmed under the user's autonomous-decision authorization

## Context

The declaration audit has fifteen unverified atoms under `ShapeProps.points`.
Together they describe PptxGenJS custom paths: ordinary and explicit move
points, lines, quadratic and cubic Beziers, arcs, and path closure. Native
already owns a strict custom-geometry command tree and supports creation,
reading, whole replacement, preset/custom conversion, six-format reopening,
packed consumers, and browser bundles. This family binds that implementation
to the declaration matrix and adds the missing real-Chrome aggregate gate.

## Batch Boundary

The family contains exactly:

- `ShapeProps.points`;
- inline `points.close`, `curve`, `moveTo`, `x`, and `y`;
- inline `points.curve.type`, `hR`, `wR`, `stAng`, `swAng`, `x1`, `y1`,
  `x2`, and `y2`.

It excludes the shared `Coord` union, shape transform fields, `ShapeType` /
`custGeom`, preset adjustment shortcuts, guide formulas, handles, connection
sites, text rectangles, evaluator-only state, path fill/stroke flags, and
TextProps custom-path behavior.

## Approach Decision

Three approaches were considered:

1. Add `ShapeProps.points` as a permissive native alias. This would duplicate
   the existing geometry tree and copy unsafe unit guessing and coercion.
2. Mark the fifteen atoms supported because legal final OOXML imports into the
   same native commands. This would hide the intentionally different public
   API, units, validation, and malformed-output boundary.
3. Reuse native `CustomGeometryPath.commands` and classify all fifteen atoms
   as deliberate differences. This preserves strict explicit units while
   proving every legal PptxGenJS command has a native equivalent.

The third approach is selected.

## Locked Classification

All fifteen atoms are deliberate differences:

- the first ordinary point maps to `moveTo`; later ordinary points map to
  `lineTo`, while `moveTo: true` starts another subpath;
- `quadratic` maps `x1/y1` and `x/y` to one `quadraticBezierTo` command;
- `cubic` maps `x1/y1`, `x2/y2`, and `x/y` to one `cubicBezierTo` command;
- `arc` maps `wR/hR/stAng/swAng` to `arcTo`; PptxGenJS ignores the declared
  endpoint `x/y` on this branch;
- `close: true` maps to `close`, and an empty points list remains a legal
  empty path.

PptxGenJS 4.0.1 guesses units from value shape and magnitude, resolves
percentages against the full slide, coerces some malformed curves, and can
write invalid first arcs, radii, or unsafe coordinates. Native instead accepts
an explicit discriminated command union using direct DrawingML values (with
`inches()` and `degrees()` helpers), rejects invalid trees before mutation, and
byte-preserves unsafe imported geometry without exposing an editable snapshot.

## Evidence Architecture

- Runtime/control: reuse the aggregate all-command import and unit/malformed
  boundary controls against PptxGenJS 4.0.1.
- Native: map all atoms to `CustomGeometry`, `CustomGeometryPath`,
  `CustomGeometryCommand`, `SlideModel.addCustomShape`, and
  `ShapeModel.customGeometry`.
- Package: reuse the actual-tarball all-command and multi-path lifecycle.
- Browser: add one real-Chrome aggregate with all six command kinds,
  deep-frozen state, writeBlob/reopen, exact XML tags, and zero diagnostics.
- OOXML: reuse SDK create/edit/reopen coverage and inspect the generated Chrome
  deck once for package validity and exact command ordering.
- Matrix: generate exactly fifteen frozen entries and assert scope, evidence,
  exclusions, status totals, and exact migration.

Evidence agents remain read-only. Only the main agent edits shared files,
reviews, commits, and pushes.

## Acceptance Criteria

- Exactly fifteen atoms move from `unverified` to
  `deliberate-difference`; no other atom changes status.
- Global totals become 620 supported, 354 deliberate-difference, 91
  deprecated-alias, 359 defect-excluded, and 350 unverified: 1,424 of 1,774
  atoms classified (80.27%).
- Focused custom-geometry tests run once; audit tests, TypeScript,
  deterministic generation, exact migration review, and `git diff --check`
  pass.
- One actual npm tarball run, one persistent Chrome session, one OOXML
  inspection batch, and one full Vitest run pass at this family boundary.
- Review produces one capability-family commit and push with no unrelated
  tracked or staged files.
