# PptxGenJS Placeholder Core Surface Family Design

Date: 2026-08-04

Status: confirmed under the user's autonomous-decision authorization

## Context

The declaration audit has eight unverified placeholder-core atoms. Native
already implements strict placeholder definition, materialization, population,
editing, duplication, rollback, six-format reopen, installed-package use,
browser use, and legal OOXML validation. This family closes the declaration
matrix and strengthens aggregate evidence; it does not add another placeholder
implementation.

## Batch Boundary

The family contains exactly:

- `PlaceholderProps.name`, `type`, `x`, `y`, `w`, and `h`;
- `ImageProps.placeholder`;
- `TextPropsOptions.placeholder`.

It excludes `PlaceholderProps.align`, `margin`, `transparency`, and `valign`,
all placeholder-style subproperties, chart/table/media selector declarations,
and the already classified `PLACEHOLDER_TYPE` and `PLACEHOLDER_TYPES` union
catalogs.

## Approach Decision

Three approaches were considered:

1. Copy PptxGenJS names, implicit-inch coordinates, string-only selectors,
   owner fallbacks, and mutation behavior into native. This would weaken the
   existing strict editable placeholder lifecycle.
2. Mark the eight atoms supported because their common legal cases work. This
   would hide observable type, identity, geometry, invalid-input, and delayed
   population differences.
3. Reuse the existing native vertical slices, add one exact aggregate runtime
   control, and classify every atom conservatively as a deliberate difference.
   This is the selected approach.

## Locked Classification

All eight atoms are deliberate differences.

- `name`: PptxGenJS separates its lookup name from the selection-pane object
  name and internally remaps the name on its cloned definition; native requires
  one detached, non-empty, unique owner name and persists it.
- `type`: PptxGenJS accepts invalid values by fallback and emits empty `pic`
  and `tbl` definitions as `body`; native preserves all six strict placeholder
  domains and exposes frozen `PlaceholderIdentity` state.
- `x` and `y`: PptxGenJS accepts loose `Coord` values and implicit-inch
  numbers; native accepts explicit EMU, `inches()`, or strict percentages and
  validates before mutation.
- `w` and `h`: native names the fields `width` and `height`, uses explicit
  units, and rejects non-positive extents; PptxGenJS exposes `w` and `h` and
  permits zero extent.
- `ImageProps.placeholder`: PptxGenJS accepts only a string selector, may
  degrade `pic` identity to `body`, and does not inherit the full owner extent;
  native accepts name or type/index identity, replaces the owner in place,
  preserves the full transform, and supports raster and SVG sources.
- `TextPropsOptions.placeholder`: PptxGenJS accepts a string selector and can
  create duplicate identities after delayed population; native accepts name or
  identity selectors, replaces the materialized owner in place, and rejects
  missing, ambiguous, wrong-domain, or already-filled owners atomically.

## Evidence Architecture

- Runtime/control: reuse `matches public slide master objects, topology, and
  empty placeholder geometry`, `compares public placeholder population
  payloads with strict native owners`, and `locks public slide master fallbacks
  while native definitions reject atomically`; add one family aggregate for the
  exact eight-atom inventory, percentage geometry, conflicting population
  geometry, selector failures, and the image type/extent defect.
- Native: reuse `AddPlaceholderOptions`, `PlaceholderIdentity`,
  `PlaceholderSelector`, `SlideModel.addPlaceholder`, text population, and
  raster/SVG image population.
- Package/client: reuse the installed-package `masterLayoutChecks` and the
  persistent-Chrome `masterLayoutState`, extending their shared state with one
  placeholder-transform equality check rather than opening new runs.
- OOXML: reuse empty-placeholder six-format reopen, population lifecycle,
  exact `p:ph` ownership, validator, and PowerPoint 2010 package checks.
- Matrix: generate exactly eight entries from one frozen ID group and assert
  exact migration and global counts.

Evidence agents remain read-only. Only the main agent edits the manifest,
tests, generated matrix, design, and plan.

## Acceptance Criteria

- Exactly eight atoms move from `unverified` to `deliberate-difference`.
- Global totals become 620 supported, 311 deliberate-difference, 91
  deprecated-alias, 359 defect-excluded, and 393 unverified: 1,381 of 1,774
  atoms closed (77.85%).
- The aggregate and existing runtime controls, focused native tests, audit
  tests, TypeScript, deterministic generation, exact migration review, and
  `git diff --check` pass.
- npm pack, persistent Chrome, OOXML package inspection, and full Vitest remain
  shared gates at the image-family batch boundary; no redundant session or
  package build is started for this family.
- Review produces one capability-family commit and push with no unrelated
  tracked or staged files.
