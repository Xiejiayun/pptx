# PptxGenJS Shape and Text Shadow Surface Family Design

Date: 2026-08-04

Status: confirmed under the user's autonomous-decision authorization

## Context

The declaration audit has 12 unverified atoms for shared `ShadowProps` and two
in-scope owner properties. Native already implements strict simple outer/inner shadow
creation, reading, editing, clearing, duplication, rollback, six-format
reopen, packed consumers, browser consumers, and legal OOXML validation for
preset shapes and text shapes. This family closes only the declaration
matrix; it does not add another shadow implementation.

## Batch Boundary

The family contains exactly:

- `ShadowProps.angle`, `blur`, `color`, `offset`, `opacity`,
  `rotateWithShape`, and `type`;
- the `inner`, `none`, and `outer` members of `ShadowProps.type`;
- `ShapeProps.shadow` and `TextPropsOptions.shadow`.

It excludes `ImageProps.shadow`, advanced DrawingML effects, preset shadows,
effect DAGs, and shadow fields owned by tables, charts, or media.

## Approach Decision

Three approaches were considered:

1. Copy the loose PptxGenJS `ShadowProps` API into native. This would duplicate
   the existing `ShapeShadow` contract and copy permissive coercion and invalid
   output behavior.
2. Mark every atom supported because native has a broader legal shadow
   lifecycle. This would hide observable differences in explicit zero,
   defaults, names, colors, `none`, and two owner-specific upstream defects.
3. Reuse the existing vertical slices and close the declaration atoms
   conservatively with owner-aware runtime controls. This is the selected
   approach.

## Locked Classification

### Deliberate differences: 12

- `ShadowProps.angle`, `blur`, `color`, `offset`, `opacity`, and `type`;
- `ShadowProps.rotateWithShape`;
- `ShadowProps.type` members `inner`, `none`, and `outer`;
- `ShapeProps.shadow` and `TextPropsOptions.shadow`.

PptxGenJS uses `type`, `offset`, hex strings, implicit defaults, truthy
fallbacks, and permissive/coercible inputs. Native uses `kind`, `distance`,
strict sRGB/theme colors, validation-before-mutation, detached frozen state,
and preserves explicit zero. Legal outer shadows reach the same final simple
shadow semantics, while native represents `none` as absence and exposes a
safe editable lifecycle.

The shared declaration must be classified across all reachable owners.
PptxGenJS shape/text writers ignore `rotateWithShape: true` and produce a
mismatched closing tag for `inner`, but chart can serialize the rotate flag
and image/chart can serialize a legal inner shadow. These are owner-specific
runtime differences, not globally unusable declaration atoms. The two generic
atoms therefore remain deliberate differences; the `ShapeProps.shadow` and
`TextPropsOptions.shadow` notes record their local failures.

## Evidence Architecture

- Runtime/control: reuse the existing adapter tests `compares shape shadow
  public output and strict native divergences` and `compares text shape shadow
  public output and strict native divergences`, and add one aggregate
  cross-owner control for the exact 12-atom inventory, image/chart inner
  output, and chart rotate behavior.
- Native: reuse `ShapeShadow`, `ShapeModel.shadow`, the simple-shadow codec,
  and the shape/text SDK lifecycle tests.
- Package/client: reuse the current actual-tarball `shapeShadows` and
  `textShapeShadows` states, the browser-conditional `browserShadowChecks`,
  and the persistent Chrome `textShapeShadowState`.
- OOXML: reuse all-format SDK reopen tests and exact shadow-part/PowerPoint
  2010 validation evidence already exercised by the installed package smoke.
- Matrix: generate the exact family from frozen ID/status groups and assert
  exact migration and global counts.

Evidence agents remain read-only. Only the main agent edits the manifest,
generated matrix, tests, design, and plan.

## Acceptance Criteria

- Exactly 12 atoms move from `unverified`, all as deliberate differences.
- Global totals become 620 supported, 303 deliberate-difference, 91
  deprecated-alias, 359 defect-excluded, and 401 unverified: 1,373 of 1,774
  atoms closed (77.40%).
- The aggregate and both existing runtime controls, focused shadow/native
  tests, audit tests, deterministic matrix generation, exact migration
  review, and `git diff --check` pass.
- No new runtime implementation, package build, or browser session is created
  for this matrix-only family. Expensive package, Chrome, and full-suite gates
  remain shared at the multi-family batch boundary.
- Review produces one capability-family commit and push with no unrelated
  tracked or staged files.
