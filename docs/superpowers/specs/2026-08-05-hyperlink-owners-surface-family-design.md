# Hyperlink Owners Surface Family Design

## Scope

Close exactly six PptxGenJS 4.0.1 declaration atoms as one capability family:
the `url`, `slide`, and `tooltip` fields of `HyperlinkProps`, plus hyperlink
ownership on shapes, images, and text options. Runtime evidence fixes the final
classification at five supported atoms and one deliberate difference.

## Native ownership model

Shapes, text, and pictures use the existing strict hyperlink codec and
relationship lifecycle. `ImageModel` gains a hyperlink property that delegates
to the same slide-owned reader/editor used by shapes. Raster and SVG creation
options carry one normalized, detached hyperlink value into their existing
outer package transaction; no image-specific codec or relationship model is
introduced.

Internal slide targets are validated before media parts or relationships are
allocated. The created picture receives its link before the existing outer
transaction completes, so any failure rolls back media, XML, and relationships
as one unit. High-level SDK options detach the nested hyperlink synchronously
before awaiting file, URL, or binary image sources.

## Compatibility boundary

Native state requires one non-empty URL or one positive, in-range, one-based
slide target, with an optional string tooltip. Inputs must be ordinary own-data
objects and are never mutated. Omitted and explicitly empty tooltips remain
distinct. Existing same-value no-op, relationship-ID reuse, clone-on-write,
garbage collection, slide reorder/delete behavior, duplicate isolation, and
transaction rollback apply unchanged to image owners.

`TextPropsOptions.hyperlink` is a deliberate difference. PptxGenJS 4.0.1 can
emit dangling `rIdundefined` clicks for a legal rich-text outer hyperlink.
Native preserves the legal whole-shape/run semantic with valid relationship
ownership and does not reproduce that upstream defect.

## Evidence and verification

One aggregate PptxGenJS runtime control fixes the exact six-ID inventory and
5/1 split. Focused model and SDK tests cover raster/SVG create, edit, clear,
duplicate, move, rollback, write, and reopen across all six presentation
formats, including invalid-input zero mutation and preservation of image
appearance/media state.

One shared lifecycle probe runs through the workspace build, the produced npm
tarball, and the persistent browser session. Retained decks receive exact
OOXML/relationship checks and PowerPoint 2010 validation. The family runs one
focused gate and one batch-level TypeScript/full-suite/package/browser/audit
gate. The expected matrix result is 1,740/1,774 classified (98.08%), with 34
unverified atoms and zero unsupported, stale, or diagnostic entries.
