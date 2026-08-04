# PptxGenJS Image Source, Sizing, and Transform Surface Family Design

Date: 2026-08-04

Status: confirmed under the user's autonomous-decision authorization

## Context

The declaration audit has fourteen unverified `ImageProps` atoms for image
source selection, intrinsic sizing, cropping, rotation, and flips. Native
already implements raster and SVG source loading, strict sizing calculation,
direct `a:srcRect` state, transform editing, six-format reopen, installed npm
package use, and persistent-browser use. This family closes the evidence gap;
it does not add a second image implementation.

## Batch Boundary

The family contains exactly:

- `ImageProps.data`, `path`, `sizing`, `flipH`, `flipV`, and `rotate`;
- inline `ImageProps.sizing.type`, `w`, `h`, `x`, and `y`;
- the `contain`, `cover`, and `crop` tokens of `ImageProps.sizing.type`.

It excludes already classified direct `x/y/w/h` and `placeholder`; metadata
`altText/objectName`; interaction `hyperlink`; appearance
`rounding/transparency/shadow`; and inherited `DataOrPathProps.data/path`, which
remain cross-owner atoms until media and text owners have shared evidence.

## Approach Decision

Three approaches were considered:

1. Copy the permissive PptxGenJS object shape and implicit-inch semantics into
   native. This would introduce ambiguous source and coordinate state into the
   strict editable API.
2. Mark all fourteen atoms supported because representative legal output
   matches. This would hide public naming, unit, source-precedence, truthy
   fallback, and invalid-input differences.
3. Reuse the existing vertical slice and classify all fourteen atoms as
   deliberate differences because the same sizing token can produce different
   legal output when outer geometry and intrinsic dimensions disagree. This is
   the selected approach.

## Locked Classification

All fourteen atoms are deliberate differences:

- `data` and `path`: PptxGenJS exposes two optional fields with permissive
  precedence and MIME fallbacks; native accepts one typed `ImageSource`,
  detects content from bytes, resolves it before mutation, and rejects
  ambiguous or unsafe state atomically.
- `sizing`: PptxGenJS combines outer geometry and a permissive inline object;
  native uses a discriminated `ImageSizing` option whose target size owns the
  final frame and whose crop source uses intrinsic pixels.
- inline `w/h`: native names these `width/height`, requires explicit EMU or
  `inches()` numeric units, and validates positive safe extents.
- inline `x/y`: PptxGenJS reuses layout coordinates for crop input; native
  places crop coordinates in a required `ImageCropRegion.source` pixel region.
- inline `type` and the `contain/cover/crop` tokens: native applies the same
  three operations to intrinsic image dimensions. PptxGenJS applies its sizing
  arithmetic to outer `w/h`, so a legal source whose intrinsic aspect ratio
  differs from the outer ratio produces a different `a:srcRect`; its missing
  extents and unknown-type failure timing also differ.
- `flipH/flipV/rotate`: native names these
  `flipHorizontal/flipVertical/rotation`, requires strict booleans and
  `OoxmlAngle`, and preserves zero instead of applying truthy fallbacks.

## Evidence Architecture

- Runtime/control: reuse the PptxGenJS 4.0.1 controls for embedded raster
  output, path/data source loading, contain/cover/crop final state, SVG source
  and OOXML roles, sizing fallbacks, and strict native divergence.
- Native: reuse `ImageSource`, `AddImageSourceOptions`, `resolveImageSource`,
  `ImageSizing`, `normalizeImageSizing`, `calculateImageSizing`,
  `ImageModel.sourceRectangle`, and strict transform normalization.
- Package/client: extend the actual installed-package `packedSvgDeck` state
  with `packedSvgPath` and all three sizing modes, and extend the same
  persistent-Chrome `svgDocument` flow with Blob, data URI, and browser-URL
  sources plus one `imageSourceSizingTransformState` assertion.
- OOXML: reuse raster/SVG internal image relationships, exact direct
  `a:srcRect`, transform attributes, PowerPoint 2010 validation, and all-six-
  format reopen tests.
- Matrix: generate exactly fourteen entries from frozen ID groups and assert
  the exact migration and global totals.

Evidence agents remain read-only. Only the main agent edits the manifest,
tests, generated matrix, design, and plan.

## Acceptance Criteria

- Exactly fourteen atoms move from `unverified` to `deliberate-difference`.
- Global totals become 620 supported, 325 deliberate-difference, 91
  deprecated-alias, 359 defect-excluded, and 379 unverified: 1,395 of 1,774
  atoms closed (78.64%).
- Focused image tests run once; audit tests, TypeScript, deterministic
  generation, exact migration review, and `git diff --check` pass.
- One actual npm tarball build, one persistent Chrome session, one OOXML
  inspection batch, and one full Vitest run pass at this image-family boundary.
- Review produces one capability-family commit and push with no unrelated
  tracked or staged files.
