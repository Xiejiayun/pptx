# Image Identity and Visual Effects Surface Family Design

## Scope

Close these five PptxGenJS 4.0.1 declaration atoms as one capability family:

- `interface:ImageProps@property:altText`
- `interface:ImageProps@property:objectName`
- `interface:ImageProps@property:rounding`
- `interface:ImageProps@property:shadow`
- `interface:ImageProps@property:transparency`

The family must work for zero-input raster and SVG creation, imported images,
live edits, duplicated slides, transaction rollback, packed npm consumers, and
the browser bundle.

## Public semantics

- `ImageModel.name` is editable and maps to `p:cNvPr@name`. This is the strict
  native counterpart of PptxGenJS `objectName`.
- `ImageModel.altText` is editable and maps to `p:cNvPr@descr`. Setting
  `undefined` removes the attribute.
- `ImageModel.rounding` maps `false` to `a:prstGeom@prst="rect"` and `true` to
  `ellipse`.
- `ImageModel.transparency` is a percentage in `[0, 100]`. It maps to
  `a:blip/a:alphaModFix@amt = (100 - transparency) * 1000`; zero removes the
  element.
- `ImageModel.shadow` reuses strict native `ShapeShadow`. Native preserves a
  legal `rotateWithShape: true` value instead of reproducing PptxGenJS 4.0.1's
  false serialization bug.
- `AddImageOptions`, `AddSvgImageOptions`, and SDK `AddImageSourceOptions`
  accept the same appearance fields so the family is available when creating
  a deck from zero.

## OOXML ownership and safety

Each editor owns only its direct semantic fragment:

- identity owns one unqualified `name` or `descr` attribute on the single
  presentation `cNvPr` descendant;
- rounding owns the unqualified `prst` attribute on one direct DrawingML
  `prstGeom`, and only accepts imported `rect` or `ellipse` state;
- transparency owns one direct DrawingML `alphaModFix` under the image blip;
- shadow owns one strict inner/outer shadow under `spPr/effectLst`.

Duplicate attributes, ambiguous geometry, malformed effect lists, unknown
blip-effect ordering, or conflicting namespaces are not rewritten. An edit
throws before package mutation. Unrelated elements, extensions, attributes,
relationships, and media bytes remain unchanged.

## Creation and normalization

Creation normalizes appearance before opening the package transaction:

- `rounding` defaults to `false`;
- `transparency` defaults to `0` and is quantized to OOXML thousandths of a
  percent;
- `shadow` is detached and deeply frozen by the existing strict shadow
  normalizer;
- invalid option containers, accessors, XML strings, booleans, percentages,
  and shadow values are rejected without reading unsafe getters.

Raster and SVG creation share the same normalized appearance. SVG blips place
`alphaModFix` before their existing `extLst`.

## Verification

One family gate covers:

1. creation normalization and canonical raster/SVG XML;
2. imported read/edit/no-op behavior and malformed-state rejection;
3. live model create, edit, duplicate isolation, rollback, relationship
   stability, and write/reopen equivalence;
4. PptxGenJS runtime/control comparison for the five atoms;
5. SDK declarations and source preparation;
6. npm tarball and persistent Chrome lifecycle probes;
7. PowerPoint 2010 validation and direct OOXML assertions;
8. surface audit counts, stale-entry detection, and zero diagnostics.

## Matrix classification

- supported: `altText`, `rounding`, `transparency`;
- deliberate-difference: `objectName` (native `name`) and `shadow` (strict
  `ShapeShadow`, preserving legal rotate-with-shape state).
