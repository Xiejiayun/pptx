# Preset Shape Adjustments Design

## 1. Objective

Add native create, read, replace, and clear support for preset-shape adjustment
guides stored in `p:sp/p:spPr/a:prstGeom/a:avLst/a:gd`. The public model must
cover every simple preset guide rather than only PptxGenJS 4.0.1's
`rectRadius`, `angleRange`, and `arcThicknessRatio` shortcuts, while matching
the final OOXML those shortcuts produce.

This item does not add `a:custGeom`, arbitrary guide formulas, adjustment
handles, connection sites, custom paths, or geometry evaluation. Those remain
the next custom-geometry item.

## 2. Evidence and compatibility target

PptxGenJS 4.0.1 public `addShape()` output maps all three shortcut options to
simple preset adjustment guides:

- `rectRadius` emits `name="adj"` and a `val` formula calculated from the
  smaller rendered dimension;
- `angleRange` emits `adj1` and `adj2` in OOXML angle units;
- `arcThicknessRatio`, when `angleRange` is present and the ratio is truthy,
  emits `adj3` scaled by 50000.

The public-output probe also establishes the behaviors that native conformance
tests must record:

- omitted and numeric-zero `rectRadius` both emit an empty `a:avLst`;
- numeric-zero `angleRange` values are retained;
- numeric-zero `arcThicknessRatio` is dropped;
- negative, over-range, string-coerced, and shape-inappropriate runtime values
  are serialized without validation;
- `rectRadius` wins over `angleRange`, and thickness without `angleRange` is
  ignored.

Native code matches every valid final guide list. It does not reproduce
truthiness loss, string coercion, invalid ranges, or silent option precedence.
Callers express the desired final guide integers directly.

## 3. Alternatives considered

### A. Copy the three PptxGenJS shortcut fields

This is familiar to PptxGenJS users but is incomplete for the much larger set
of preset geometries, couples `rectRadius` to shape dimensions, and still does
not provide an existing-deck editor. Rejected as too narrow.

### B. Expose raw guide XML

This could represent every formula but would bypass descriptor safety, XML
escaping, immutable snapshots, semantic no-op detection, and lossless mutation
isolation. Rejected because it weakens the model contract.

### C. Expose typed simple adjustment guides

Use one ordered list of `{ name, value }` entries for creation and editing.
This directly represents all PptxGenJS valid output, covers other simple preset
adjustments, and leaves complex formulas for a later geometry model. Selected.

## 4. Public API

```ts
export interface ShapeAdjustment {
  readonly name: string;
  readonly value: number;
}

export interface AddShapeOptions extends Partial<Transform> {
  readonly adjustments?: readonly ShapeAdjustment[];
}

export class ShapeModel {
  get adjustments(): readonly ShapeAdjustment[] | undefined;
  set adjustments(value: readonly ShapeAdjustment[]);
}
```

`AddShapeOptions.adjustments` omission and runtime `undefined` preserve the
existing canonical empty adjustment-list bytes. A supplied empty list produces
the same canonical empty list. `ShapeModel.adjustments` returns `undefined`
only when the owned geometry state cannot be interpreted safely; it returns a
frozen empty array for a supported empty list.

The setter accepts a list only. `[]` is the explicit clear operation and
removes all direct guides while retaining the adjustment-list container.
`undefined` is rejected at runtime instead of combining clear with unsupported
getter state.

## 5. Value contract

An adjustment list must be a dense ordinary JavaScript array. Sparse arrays,
array subclasses, proxies that violate ordinary descriptor access, symbol
properties, accessors, and non-index own properties are rejected. Each entry
must be an ordinary or null-prototype object with exactly two own data
properties, `name` and `value`.

Names must be non-empty strings with no XML-invalid characters. Duplicate names
are rejected because guide resolution by name would otherwise be ambiguous.
Names are XML-attribute escaped during rendering and their caller-provided
order is retained.

Values must be finite safe integers. No percentage, angle, point, inch, or
shape-specific conversion occurs in this API. This keeps one exact unit: the
integer operand of `fmla="val N"`. It also supports PptxGenJS angle output,
ratio output, negative values, and every safe integer without guessing a
preset's valid semantic range.

Normalization reads property descriptors rather than invoking getters, copies
all data before package mutation, freezes every copied entry, and freezes the
copied list. Getter snapshots are newly detached and deeply frozen.

## 6. Supported OOXML state

The reader starts from one namespace-correct direct `p:spPr/a:prstGeom` with a
canonical preset type. The geometry must contain exactly one direct
namespace-correct `a:avLst`. The list may be empty or contain only direct
same-namespace `a:gd` children.

Each supported guide has exactly one unqualified `name` attribute and one
unqualified `fmla` attribute, no other non-namespace attributes, and no child
elements. The formula must consist of `val`, one or more XML whitespace
characters, and one signed decimal integer operand with no leading or trailing
tokens. Optional `+` and `-` signs are accepted, and the parsed value must be a
safe integer. Namespace declarations do not affect the value contract.

The reader returns `undefined` without mutation for missing or repeated direct
geometry/list elements, wrong namespaces, non-guide list children, duplicate
names, missing/repeated/qualified owned attributes, unsupported attributes,
child elements, non-`val` formulas, multiple operands, or unsafe integers.

Complex formulas such as `*/`, `+-`, `pin`, `min`, and guide-name references
are deliberately unsupported in this item. They stay byte-preserved during
all unrelated edits.

## 7. Creation

`normalizePresetShape()` adds `adjustments` to the descriptor-safe option-key
set and normalizes it before any transaction mutation, relationship creation,
or shape-ID allocation. The normalized state stores a frozen list.

`renderPresetShapeXml()` renders the list inside the existing direct
`a:avLst`. Omission and an explicit empty list retain the previously published
`<a:avLst/>` bytes. Non-empty values use deterministic compact XML:

```xml
<a:prstGeom prst="blockArc"><a:avLst><a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 0"/><a:gd name="adj3" fmla="val 25000"/></a:avLst></a:prstGeom>
```

Adjustment creation is independent from fill, line, arrows, hyperlink,
shadow, transform, name, IDs, relationships, and shape order.

## 8. Existing-deck editing

A focused internal module owns adjustment normalization, reading, rendering,
semantic equality, and replacement. `SlideModel` resolves the live shape by ID
and performs replacement inside the existing OPC transaction mechanism.

For a supported state, assigning the same ordered names and integer values is
an exact no-op: slide bytes, package mutation journal, live object identity,
and relationships remain unchanged.

Assigning a different list replaces only the direct `a:avLst` element span,
using its existing in-scope DrawingML prefix. The replacement preserves the
`a:prstGeom` element, its attributes, namespace context, preset type, and all
unrelated geometry-local bytes. Clearing writes an empty same-prefix
`a:avLst`. No other shape property or package part changes.

If geometry or adjustment state is unsupported, the getter returns
`undefined`; setter normalization errors throw `TypeError` or `RangeError`,
and unsafe source structure throws `ModelParseError` with the slide part URI.
All failures occur with zero package mutation.

Changing `ShapeModel.presetType` to a different type keeps its existing
whole-geometry replacement behavior and therefore resets adjustments. Setting
the same type remains an exact no-op and preserves the current list.

## 9. Isolation and lifecycle

The capability must preserve:

- caller detachment and frozen snapshots;
- stable `ShapeModel` identity after changes;
- source/duplicate isolation;
- outer-transaction rollback of bytes, snapshots, IDs, and journal;
- all six presentation formats (`pptx`, `pptm`, `potx`, `potm`, `ppsx`,
  `ppsm`);
- write/reopen values and order;
- unrelated XML, relationships, content types, and opaque parts.

## 10. Verification

### Internal codec and source-span editing

- normalize empty, one-guide, multi-guide, negative, zero, min/max safe integer,
  escaped name, null-prototype entry, detachment, and deep freeze;
- reject sparse/exotic arrays, symbols, accessors, unknown/missing fields,
  empty/invalid names, duplicates, non-number/fractional/non-finite/unsafe values;
- read canonical and alternate-prefix lists, XML whitespace, optional signs,
  and namespace declarations;
- reject every malformed, ambiguous, complex-formula, and wrong-namespace
  state without mutation;
- verify ordered semantic equality, exact no-op, same-prefix replacement, clear,
  and preservation of geometry-local siblings and all external bytes.

### Public creation and editing

- create/read/edit/clear through `AddShapeOptions`, `ShapeModel`, and
  `SlideModel`;
- coexist with transform, fill, line, arrows, hyperlink, shadow, text, and
  preset-type editing;
- prove validation before shape ID allocation and relationships;
- prove duplicate isolation, rollback, stable identity, all-format lifecycle,
  and write/reopen.

### PptxGenJS and release surface

- generate public PptxGenJS 4.0.1 `rectRadius`, `angleRange`, and
  `arcThicknessRatio` cases and compare imported final guide snapshots with
  native shapes;
- record truthiness, coercion, precedence, and invalid-runtime divergences;
- cover Node, browser, declarations, and CLI through an actual packed tarball;
- run the complete suite, typecheck, performance test, build, package smoke,
  validator, LibreOffice render, overflow check, and artifact-tool import on a
  representative adjustment gallery;
- update README, API docs, changelog, compatibility baseline, and remaining-gap
  text.

## 11. Completion gate

Preset-shape adjustments are marked supported only when:

1. native creation, read, ordered replacement, clear, and reopen work for
   simple `val` guides;
2. malformed and complex source state is preserved and cannot be accidentally
   edited;
3. identity, no-op, detachment, rollback, duplicate isolation, six formats,
   and unrelated-byte preservation have direct tests;
4. valid PptxGenJS shortcut output imports to the same snapshots and documented
   runtime divergences have explicit evidence;
5. packed public surfaces, full checks, validation, and visual QA pass;
6. custom geometry and arbitrary formulas remain explicitly listed as the next
   unsupported scope.
