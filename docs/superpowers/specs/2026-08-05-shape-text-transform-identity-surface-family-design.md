# Shape and Text Transform Identity Surface Family Design

## Scope

Close exactly 13 PptxGenJS 4.0.1 declaration atoms as one capability family:

- `ShapeProps`: `flipH`, `flipV`, `objectName`, `rectRadius`, `rotate`, and
  `shapeName`;
- `TextPropsOptions`: `flipH`, `flipV`, `isTextBox`, `objectName`,
  `rectRadius`, `rotate`, and `shape`.

The family covers zero-input creation, imported PptxGenJS output, live shape
and text identity edits, duplicate isolation, rollback, package consumers,
the browser bundle, and PowerPoint 2010 validation.

## Reuse boundary

Existing native APIs already cover the transform and geometry atoms:

- `BaseShapeModel.transform` and `setTransform()` expose strict rotation and
  horizontal/vertical flips;
- `ShapeModel.presetType`, adjustments, `rectRadius`, and `isTextBox` expose
  the supported shape/text geometry state;
- `AddShapeOptions.name` and `AddTextOptions.name` create XML-safe identity;
- the shared `BaseShapeModel.name` setter edits the unique presentation
  `cNvPr@name` without changing geometry, text, relationships, or other XML.

No second identity codec or transform implementation is added. The shared
name normalizer becomes owner-neutral and reports `Shape name` for invalid
ordinary shape, text, and image names.

## Safety and lifecycle

The permanent lifecycle gate uses one preset shape and one text shape. It
requires same-value exact no-op, invalid-input isolation, outer transaction
rollback, escaped and explicit-empty names, duplicate-slide isolation, and
write/reopen equivalence. Shape transform/adjustments and text transform,
content, preset geometry, `rectRadius`, `isTextBox`, and relationships must
remain unchanged.

Malformed, duplicated, missing, or wrong-namespace `nvSpPr/cNvPr` ownership
must reject the name edit without changing slide XML, package bytes, cached
identity, or the mutation journal.

## Upstream control and classification

One aggregate PptxGenJS 4.0.1 test locks the exact 13-ID inventory and reuses
the existing runtime fixtures for transforms, preset geometry, radius, and
text-box behavior.

- supported: `TextPropsOptions.isTextBox`, `TextPropsOptions.shape`;
- deliberate-difference: both owners' `flipH`, `flipV`, `objectName`,
  `rectRadius`, and `rotate` atoms;
- defect-excluded: `ShapeProps.shapeName`, which is declared but ignored by
  PptxGenJS at runtime.

## Verification

The family receives one focused gate and one batch artifact gate:

1. permanent model name lifecycle and malformed-owner rejection;
2. exact PptxGenJS 4.0.1 aggregate control and 2/10/1 matrix assertion;
3. reused transform, preset, radius, and text-box source tests;
4. actual npm tarball and persistent Chrome name edit/reopen probe;
5. PowerPoint 2010 validation and direct OOXML assertions;
6. deterministic matrix generation, zero stale/unsupported/diagnostics;
7. TypeScript and the full Vitest suite once for the family.

Expected matrix result: 1,720/1,774 classified, 54 unverified, 96.96%.
