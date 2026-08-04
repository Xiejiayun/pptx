# PptxGenJS Master, Background, and Slide-Number Surface Family Design

## Context

The PptxGenJS 4.0.1 declaration audit still contains 34 unverified atoms for
slide-master definitions, direct backgrounds, and slide-number geometry and
formatting. Native already implements strict editable master layouts,
backgrounds, and slide-number owners. This family closes those atoms by
reusing the existing native, packed-package, browser, and OOXML evidence and
adding one aggregate upstream runtime control for declaration ownership,
aliases, invalid branches, mutation, defaults, and strict API differences.

This is the third and final family in the lifecycle artifact batch. Focused
tests, review, commit, and push remain family-local. npm pack, persistent
Chrome, OOXML package inspection, and full Vitest run once after this family is
committed.

## Batch Boundary

This family contains exactly 34 atoms:

- all 16 still-unverified `SlideMasterProps` atoms: six direct properties,
  eight inline object members, and the two `bkgd` union branches;
- all ten still-unverified `BackgroundProps` atoms: eight properties and the
  two declared `type` tokens;
- all eight still-unverified `SlideNumberProps` atoms: alignment, geometry,
  margin, transparency, and vertical alignment.

It excludes nested chart, image, shape, text, placeholder-option, fill, and
text-style members. Those atoms remain in their owning capability families.

## Locked Classification

### Supported: 4

- `SlideMasterProps.objects.placeholder.text`;
- `SlideMasterProps.title`;
- `SlideNumberProps.margin`;
- `SlideNumberProps.valign`.

Legal placeholder prompt text, master titles, scalar or four-sided margins,
and the three vertical alignment tokens serialize and reopen with the same
public meaning. Their native names and legal value domains do not require an
API-shape correction, including an explicit zero margin.

### Deliberate differences: 22

- the six valid inline master-object branches other than placeholder text
  and the defective declared chart branch: image, line, placeholder,
  placeholder options, rect, and text;
- `SlideMasterProps.background`, `margin`, `objects`, and `slideNumber`;
- `BackgroundProps.color`, `data`, `path`, `transparency`, and `type`, plus the
  `none` and `solid` tokens;
- `SlideNumberProps.align`, `h`, `w`, `x`, and `y`.

Native covers the same legal master-object, background, and slide-number
capabilities through strict discriminated objects, explicit media sources,
semantic width and height names, and explicit EMU units. PptxGenJS uses loose
unions, implicit inches or percentages, caller-owned state, truthy defaults,
and permissive fallbacks.

The aggregate control locks the material differences. PptxGenJS does not emit
a legal no-fill background for `type: 'none'`; it mutates direct background
inputs; zero slide-number width and height fall back to fixed defaults;
`align: 'justify'` becomes left. Native preserves explicit no-fill, snapshots
strict input, rejects zero dimensions before mutation, and retains justify.

### Deprecated aliases: 4

- `SlideMasterProps.bkgd` and its working string branch, canonicalized to
  `SlideMasterProps.background`;
- `BackgroundProps.alpha`, canonicalized to `transparency`;
- `BackgroundProps.fill`, canonicalized to `color`.

The master `bkgd` string, direct background `alpha`, and direct background
`fill` still affect output. Native exposes only the canonical strict forms.

### Defect exclusions: 4

- `SlideMasterProps.objects.chart`;
- the `SlideMasterProps.bkgd` `BackgroundProps` union branch;
- `BackgroundProps.src`;
- `SlideNumberProps.transparency`.

The declared master chart branch is typed as `IChartOpts`, but PptxGenJS
runtime requires an undeclared internal `{ type, data, opts }` record and
throws for the declared shape. Every legal `BackgroundProps` object supplied
through master `bkgd` is inert, direct `src` is inert, and top-level
slide-number transparency is ignored. Native corrections for canonical master
charts, backgrounds, image sources, and nested slide-number text style do not
turn these defective declaration atoms into supported atoms.

## Evidence Architecture

- Adapter: add one aggregate control titled `locks master, background, and
  slide-number declarations against PptxGenJS 4.0.1`. Reuse the existing
  master-object, master-background, slide-background, and slide-number
  controls.
- Native: reuse `PptxDocument.defineSlideMaster`, `SlideLayoutModel`,
  `SlideMasterObject`, `SlideModel.background`, `setSlideBackgroundImage`, and
  `SlideNumberOptions`, including the all-format lifecycle tests.
- npm/browser: reuse the packed and persistent-browser master-layout,
  background, and slide-number states. No second build or browser session is
  required for this family.
- OOXML: reuse the PowerPoint 2010 validation and exact package-part probes.
  Inspect the retained lifecycle artifacts once at the batch boundary.
- Matrix: generate the exact 34 entries from frozen ID/status groups and assert
  the exact family and global counts.

Only the main agent edits repository files. Evidence agents remain read-only.

## Acceptance Criteria

- Exactly 34 atoms move from `unverified`: four supported, 22
  deliberate-difference, four deprecated-alias, and four defect-excluded.
- Global totals become 620 supported, 291 deliberate-difference, 91
  deprecated-alias, 359 defect-excluded, and 413 unverified: 1,361 of 1,774
  atoms closed (76.72%).
- The aggregate control locks the exact inventory, declared chart failure,
  master alias branches, direct background aliases and defects, caller
  identity and mutation, valid master objects, slide-number defaults,
  alignment, margin, transparency, and strict native corrections.
- Focused adapter/audit tests, TypeScript, deterministic matrix generation,
  exact migration review, and `git diff --check` pass before one
  capability-family commit and push.
- The worktree contains no staged package/browser/PPTX artifacts and local and
  remote main remain at divergence `0 0`.
