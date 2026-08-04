# PptxGenJS Slide and Section Surface Family Design

## Context

The PptxGenJS 4.0.1 declaration audit still contains 26 unverified atoms for
slide-level methods and state, the loose `PresSlide` continuation projection,
slide selection by master or section, and section options. The native SDK
already implements the legal slide, notes, hidden-state, background, default
color, slide-number, auto-page, master, and section lifecycles. This family
closes those atoms by reusing the existing implementations and adding one
aggregate upstream runtime control for the remaining ownership, return-value,
mutation, fallback, and alias boundaries.

This is the second family in the three-family lifecycle artifact batch.
Focused tests, review, commit, and push remain family-local. npm pack,
persistent Chrome, OOXML package inspection, and full Vitest run once at the
batch boundary after the master/background/slide-number family is closed.

## Batch Boundary

This family contains exactly 26 atoms:

- the five `Slide` methods `addImage`, `addMedia`, `addNotes`, `addShape`, and
  `addText`;
- the six `Slide` properties `background`, `bkgd`, `color`, `hidden`,
  `newAutoPagedSlides`, and `slideNumber`;
- the seven callable `PresSlide` properties `addChart`, `addImage`, `addMedia`,
  `addNotes`, `addShape`, `addTable`, and `addText`;
- the four stateful `PresSlide` properties `background`, `color`, `hidden`, and
  `slideNumber`;
- `AddSlideProps.masterName` and `AddSlideProps.sectionTitle`;
- `SectionProps.order` and `SectionProps.title`.

It excludes the nested members of background, color, slide-number, media,
image, shape, text, table, and chart option objects. Those atoms remain in
their owning capability families.

## Locked Classification

### Supported: 10

- `PresSlide.addChart`, `addImage`, `addNotes`, `addShape`, `addTable`, and
  `addText`;
- `PresSlide.hidden`;
- `Slide.addNotes` and `Slide.hidden`;
- `SectionProps.title`.

The six supported `PresSlide` atoms declare only a callable `Function`
property, not a signature or return contract. Native `SlideModel` exposes the
corresponding callable operations, while the separately inventoried `Slide`
method atoms own their return-value differences. `PresSlide` is reachable
through `newAutoPagedSlides`, whose runtime entries are ordinary slides;
PptxGenJS's internal master placeholder also being typed as `PresSlide` does
not invalidate that legal continuation path. Legal hidden booleans, notes,
and section titles serialize and reopen equivalently.

### Deliberate differences: 15

- `AddSlideProps.masterName` and `AddSlideProps.sectionTitle`;
- `PresSlide.addMedia`, `background`, `color`, and `slideNumber`;
- `SectionProps.order`;
- `Slide.addImage`, `addMedia`, `addShape`, and `addText`;
- `Slide.background`, `color`, `newAutoPagedSlides`, and `slideNumber`.

PptxGenJS slide creation silently falls back for an unknown master and warns
but still creates a loose slide for an unknown section. Native rejects either
unknown reference before mutation. PptxGenJS also treats `order: 0` as omitted,
while native preserves zero as the first insertion index.

PptxGenJS returns the owning slide from all five slide methods. Native returns
the same slide only from `addNotes`; it returns typed live image, media, shape,
and text models for the other four operations, with media split into explicit
audio, video, and online-video APIs. Native therefore has no generic
`PresSlide.addMedia` callable.

The background, default-color, slide-number, and auto-page properties preserve
legal final-state capability but deliberately use strict detached state,
transactional validation, canonical OOXML, and immutable continuation lists.
PptxGenJS retains or mutates caller objects, accepts permissive fallbacks,
reuses existing following slides during auto-page, and exposes several known
defaulting and fixed-ID defects.

### Deprecated alias: 1

- `Slide.bkgd`, canonicalized to `Slide.background`.

The alias is a working legacy setter for `background.color`. It does not
replace an already selected canonical background and is not exposed by the
native API.

### Defect exclusions: 0

Every atom in this family has at least one valid PptxGenJS runtime path. Narrow
invalid-reference, order-zero, and internal-master defects are controlled as
deliberate strict corrections instead of excluding an otherwise reachable
atom.

## Evidence Architecture

- Adapter: add one aggregate control titled `locks the slide and section
  declarations against PptxGenJS 4.0.1`. Reuse existing chart-return, master
  fallback, section, hidden, notes, background, default-color, slide-number,
  hyperlink mutation, and auto-page controls.
- Native: reuse `SlideModel` operations and properties, `PptxDocument.addSlide`,
  `PresentationModel.addSection`, and the all-format section, notes, hidden,
  background, slide-number, and auto-page tests.
- npm/browser: reuse the packed and persistent-browser master, section, notes,
  background, slide-number, and auto-page lifecycle states.
- OOXML: reuse the existing PowerPoint 2010 validation and exact package-part
  evidence. Inspect the retained lifecycle artifact once at the three-family
  batch boundary.
- Matrix: generate the exact 26 entries from frozen ID/status groups and assert
  the exact family and global counts.

Only the main agent edits repository files. Evidence agents remain read-only.

## Acceptance Criteria

- Exactly 26 atoms move from `unverified`: 10 supported, 15
  deliberate-difference, and one deprecated-alias.
- Global totals become 616 supported, 269 deliberate-difference, 87
  deprecated-alias, 355 defect-excluded, and 447 unverified: 1,327 of 1,774
  atoms closed (74.80%).
- The aggregate control locks the exact inventory, five slide method returns,
  six valid `PresSlide` callables, the missing generic media callable, the
  internal master declaration mismatch, caller-object mutation, alias
  precedence, auto-page reuse, master/section fallbacks, and order-zero
  behavior.
- Focused adapter/audit tests, TypeScript, deterministic matrix generation,
  exact migration review, and `git diff --check` pass before one
  capability-family commit and push.
- The worktree contains no staged package/browser/PPTX artifacts and local and
  remote main remain at divergence `0 0`.
