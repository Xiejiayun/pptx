# PptxGenJS Underline Surface Family Design

## Context

The PptxGenJS 4.0.1 declaration audit still contains 120 unverified underline atoms. They are one capability family, not 120 independent implementation tasks: six declared owners each expose the `underline` property, the inline `color` and `style` fields, and seventeen declared style tokens.

The repository already has a strict native `RichTextRunStyle.underline` model, deterministic DrawingML serialization, reopen/edit support, packed-package type coverage, and browser rich-text coverage. The missing work is an evidence gap: the upstream owner matrix and token defects have not been locked in one control, and the native/package/browser evidence is not yet connected to all 120 declaration atoms.

## Goals

- Close exactly the 120 remaining underline atoms in one reviewed capability-family batch.
- Prove PptxGenJS runtime behavior for every owner and declared token.
- Preserve the native strict underline model rather than accepting upstream invalid DrawingML.
- Prove create, edit, clear, duplicate, rollback, reopen, package, browser, and OOXML behavior with aggregate fixtures.
- Update the generated compatibility matrix from evidence only.

## Non-goals

- Do not add PptxGenJS's misspelled `dotDashHeave` token to the native API.
- Do not implement inert `SlideNumberProps.underline` or `TableToSlidesProps.underline` declarations that PptxGenJS itself ignores.
- Do not split the family into per-owner, per-token, or per-atom commits.
- Do not change unrelated rich-text, table, or compatibility behavior.

## Surface Inventory

The family owns six declaration owners:

- Active in PptxGenJS output: `PlaceholderProps`, `TableCellProps`, `TableProps`, `TextPropsOptions`.
- Inert in PptxGenJS output: `SlideNumberProps`, `TableToSlidesProps`.

Each owner contributes exactly twenty atoms:

- `interface:<owner>@property:underline`;
- inline `underline.color` and `underline.style`;
- seventeen `underline.style` union tokens.

The seventeen declared tokens are:

```text
dash dashHeavy dashLong dashLongHeavy dbl dotDash dotDashHeave
dotDotDash dotDotDashHeavy dotted dottedHeavy heavy none sng
wavy wavyDbl wavyHeavy
```

## Locked Upstream Behavior

One aggregate PptxGenJS control must prove:

- All four active owners serialize all seventeen declared tokens into `a:rPr@u` in declaration order.
- A representative `{ style: 'dbl', color: 'FF0000' }` serializes `u="dbl"` and an exact `a:uFill/a:solidFill/a:srgbClr` chain.
- A color-only object emits underline fill without an explicit underline style and therefore relies on inheritance.
- An empty object emits neither `u` nor `uFill`.
- `none` emits valid `u="none"`.
- `dotDashHeave` is emitted literally even though it is not a DrawingML underline token.
- The undeclared runtime spelling `dotDashHeavy` emits valid DrawingML and remains a native extension, not a declaration atom.
- Both inert owners accept the values without throwing but emit no owner-attributable `u` or `uFill`; a positive control in the same fixture proves the owner content was created.

## Native Contract

Native keeps the strict public model:

```ts
type RichTextUnderlineStyle =
  | 'words' | 'sng' | 'dbl' | 'heavy'
  | 'dotted' | 'dottedHeavy'
  | 'dash' | 'dashHeavy' | 'dashLong' | 'dashLongHeavy'
  | 'dotDash' | 'dotDashHeavy'
  | 'dotDotDash' | 'dotDotDashHeavy'
  | 'wavy' | 'wavyHeavy' | 'wavyDbl';

interface RichTextUnderline {
  readonly style?: RichTextUnderlineStyle;
  readonly color?: RichTextColor;
}
```

`underline: false` is the native semantic equivalent of PptxGenJS `style: 'none'`. Native normalizes a color-only underline to a single underline with that color, rejects empty objects and unsupported tokens before mutation, and preserves strict sRGB or scheme colors. Creation and editing may use explicit rich-text runs for table/placeholder owners; this is an intentional API-shape difference, not a capability loss.

## Classification

The exact expected family totals are:

- `supported`: 68
- `deliberate-difference`: 8
- `defect-excluded`: 44

For each of the four active owners:

- the top-level property is `supported`;
- inline `underline.style` is `supported`;
- inline `underline.color` is `deliberate-difference` because native uses strict semantic colors and normalizes color-only input;
- the fifteen declared, valid, same-named style tokens are `supported`;
- `none` is `deliberate-difference` because native expresses the same semantic state with `false`;
- `dotDashHeave` is `defect-excluded` because PptxGenJS writes invalid DrawingML.

For each inert owner, all twenty atoms are `defect-excluded`.

## Evidence Architecture

- Adapter control: one aggregate six-owner and seventeen-token PptxGenJS fixture.
- Native SDK: one six-format aggregate lifecycle covering common styles, strict colors, `false`, snapshot isolation, edit/clear, duplicate, rollback, reopen, exact OOXML, and diagnostics.
- Packed package: one aggregate `underlineFamilyState`, executed once from the npm tarball.
- Browser: the same stable state executed once in persistent Chromium with empty error logs.
- OOXML: exact underline attributes/fill plus `pptx-inspect` package validation and mutation-isolation diff.
- Matrix: generated entries from frozen owner/token lists with exact count assertions.

Only the main agent edits repository files. Evidence agents remain read-only, and temporary harnesses or package artifacts never enter Git.

## Acceptance Criteria

- Exactly 120 matrix atoms change from `unverified` and every changed ID belongs to this family.
- Family totals equal 68 supported, 8 deliberate-difference, and 44 defect-excluded.
- Focused adapter and SDK tests pass once for the family.
- TypeScript, surface-audit tests, npm tarball smoke, persistent Chromium, `pptx-inspect`, and the full repository suite pass.
- Review finds no unrelated staged path.
- One family commit is pushed and local/remote divergence is `0 0`.
