# Public API 0.1

## Create, open, and save

```ts
import { degrees, inches, PRESET_SHAPE_TYPES, PptxDocument } from '@pptx/sdk';

const created = PptxDocument.create({
  author: 'Presentation Team',
  company: 'Acme & Partners',
  createdAt: '2024-02-29T12:34:56.123+05:30',
  format: 'pptx',
  lastModifiedBy: 'Presentation Team',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
  revision: '7',
  rtlMode: true,
  slideSize: '16:9',
  subject: 'Revenue & Forecast',
  title: 'Quarterly Review',
});
created.addSlide();
created.author = 'Updated Author';
created.author = '';
created.author = undefined;
created.company = 'Updated Company';
created.company = '';
created.company = undefined;
created.createdAt = '2026-07-30T00:00:00Z';
created.createdAt = undefined;
created.lastModifiedBy = 'Updated Editor';
created.lastModifiedBy = '';
created.lastModifiedBy = undefined;
created.modifiedAt = '2026-07-30T01:02:03Z';
created.modifiedAt = undefined;
created.subject = 'Updated Subject';
created.subject = '';
created.subject = undefined;
created.revision = '008';
created.revision = undefined;
created.title = 'Updated Review';
created.title = '';
created.title = undefined;
created.rtlMode = false;
created.rtlMode = undefined;
await created.writeFile('created.pptx');

const customSize = PptxDocument.create({
  slideSize: { width: inches(11.7), height: inches(8.3) },
});
customSize.slideSize = { width: inches(10), height: inches(7.5) };

const document = await PptxDocument.open('input.pptx', {
  limits: { maxPartBytes: 128 * 1024 * 1024 },
  signal: abortController.signal,
});

await document.writeFile('output.pptx', {
  compatibility: 'powerpoint-2010',
  mode: 'strict',
});
```

`create()` is synchronous and starts with zero slides plus a default master, blank layout, theme, notes master, and document properties. Built-in slide sizes are `4:3`, `16:9` (the default), `16:10`, and `wide`; a `{ width, height }` value accepts any OOXML-valid 1–56 inch dimensions in EMU. `document.slideSize` reads or changes the slide canvas without scaling existing shapes or changing the notes page. `CreatePresentationOptions.rtlMode` writes the direct presentation-level RTL flag; `document.rtlMode` reads or replaces that direct value, with false writing `rtl="0"` and undefined clearing it. This global flag is independent from `AddTextOptions.rtlMode` and `RichTextParagraph.rtl`, so it does not rewrite paragraph direction or alignment. All six presentation formats can be created without using PptxGenJS; macro-enabled formats start without a VBA project.

`CreatePresentationOptions.title` and `document.title` accept only strings without invalid XML control characters. Omitting `title` creates no direct title, `''` creates a direct empty `dc:title`, and assigning `undefined` removes only that direct field while preserving the core-properties part and its subject, creator, revision, timestamps, unknown children, and lexical formatting. Reads and edits locate the part through the package-root core-properties relationship rather than a fixed URI or namespace prefix; editing an existing package with no such relationship creates `/docProps/core.xml` or the next free URI. Reading never mutates, assigning the decoded current value or clearing an absent title is an exact bytes/journal no-op, and unsafe dangling, external, wrong-type, malformed, or ambiguous ownership is rejected rather than guessed.

`CreatePresentationOptions.author` and `document.author` use the direct Dublin Core `dc:creator` and accept the same strict XML-safe strings. Omitted native creation preserves the canonical `@jiayunxie/pptx` creator, `''` writes an explicit empty creator, and assigning `undefined` removes only the direct creator. The getter never falls back to `cp:lastModifiedBy`; author edits preserve lastModifiedBy, title, subject, revision, timestamps, custom/unknown children, relationships, and unrelated parts. Relationship-based lookup, missing-part creation, exact same-value/absent-clear no-ops, prefix reuse, and malformed/ambiguous rejection match the title lifecycle. PptxGenJS 4.0.1 instead defaults author to `PptxGenJS` and mirrors every author value into both creator and lastModifiedBy; native intentionally keeps lastModifiedBy independently owned.

`CreatePresentationOptions.lastModifiedBy` and `document.lastModifiedBy` own only direct core-properties `cp:lastModifiedBy`. Native omitted and runtime-`undefined` creation preserve the canonical `@jiayunxie/pptx`; `''` writes an explicit empty property, while assigning `undefined` removes only lastModifiedBy. Values are strict XML-safe strings with no trimming, coercion, user lookup, save-time refresh, revision increment, or timestamp update. Reads use the package-root core-properties relationship and namespace URI, accept alternate legal part URIs and prefixes, never fall back to creator, and never mutate. Missing metadata can be created with one canonical `cp` namespace binding; same-value and absent-clear operations are exact bytes/journal no-ops, while creator, title, subject, revision, timestamps, unknown children, relationships, and unrelated parts remain unchanged. Malformed or ambiguous ownership is rejected before mutation. PptxGenJS 4.0.1 has no independent public lastModifiedBy field: its public `author` writes creator and lastModifiedBy as a mirrored pair, while native keeps both properties independently editable.

`CreatePresentationOptions.createdAt` and `document.createdAt` own only the unique direct Dublin Core Terms `created` element with an XSI type QName resolving to `{http://purl.org/dc/terms/}W3CDTF`. Values are `string | undefined` and must match `YYYY-MM-DDTHH:mm:ss[.fraction](Z|±HH:mm)` exactly, use a valid Gregorian date in years 0001–9999, and use an offset no greater than ±14:00; empty, whitespace-padded, timezone-less, coercible, or calendar-invalid values are rejected before mutation. Omitted and runtime-`undefined` creation preserve canonical bytes and return `undefined`; assigning `undefined` removes only created. Reads follow the package-root core-properties relationship and resolve element, XSI attribute, and QName namespaces rather than assuming a part URI or lexical prefix. Missing metadata is created minimally, valid replacement repairs simple invalid text or missing/wrong unique type state, and identical valid assignment or absent clear is an exact bytes/journal no-op. Modified time, creator, lastModifiedBy, revision, other properties, unknown content, relationships, and unrelated parts remain unchanged. Native `write()` never reads the clock or refreshes timestamps; PptxGenJS 4.0.1 has no public created setter and emits a UTC seconds value during each public `write()`, whose final OOXML state imports through this same property.

`CreatePresentationOptions.modifiedAt` and `document.modifiedAt` own only the unique direct Dublin Core Terms `modified` element with the same expanded-name-correct `W3CDTF` type and strict lexical/calendar/timezone contract as created-at. Omitted and runtime-`undefined` creation preserve canonical bytes and return `undefined`; assigning `undefined` removes only modified. Reads, minimal missing-part creation, type repair, exact same-value/absent-clear no-ops, unsafe ownership rejection, and alternate URI/prefix support use the same relationship- and QName-aware rules. Modified-at edits preserve created, creator, lastModifiedBy, revision, other properties, unknown content, relationships, and unrelated parts. Native `write()` is deterministic and never refreshes modified-at or derives it from another field. PptxGenJS 4.0.1 has no public modified setter and emits a UTC seconds value during each public `write()`; adapter imports that final typed state without reproducing the hidden clock side effect.

`CreatePresentationOptions.subject` and `document.subject` own only direct Dublin Core `dc:subject`. Native omitted and runtime-`undefined` creation preserve the canonical core-properties bytes and return `undefined`; `''` writes an explicit empty subject, while assigning `undefined` removes only that direct field. Values use the same strict XML-safe string validation, relationship-based part discovery, alternate URI/prefix support, missing-part creation, same-value/absent-clear exact no-ops, and malformed/ambiguous rejection as title and author. Subject edits preserve title, creator, lastModifiedBy, revision, timestamps, unknown children, relationships, and unrelated parts. PptxGenJS 4.0.1 defaults subject to `PptxGenJS Presentation`; its custom and empty outputs import exactly, while native intentionally does not inject that brand default.

`CreatePresentationOptions.revision` and `document.revision` own only direct core-properties `cp:revision`. Values are lexical `string | undefined`; a string must contain one or more ASCII digits, so `'0'`, `'7'`, and `'007'` are valid and leading zeros are preserved exactly, while empty, signed, decimal, exponent, Unicode-digit, and non-string inputs are rejected without coercion. Native zero-input and runtime-`undefined` creation preserve canonical `'1'`, matching PptxGenJS 4.0.1. Assigning `undefined` removes only direct revision and never updates timestamps or lastModifiedBy. Reads use the root core-properties relationship and namespace URI, return `undefined` for missing, malformed, ambiguous, or lexical-invalid existing state, and never synthesize a default; valid replacement or clear can repair simple lexical-invalid state while preserving title, subject, creator, lastModifiedBy, timestamps, unknown children, relationships, and unrelated parts. Same-value and absent-clear operations are exact no-ops, and missing metadata can be created with a single canonical `cp` binding. PptxGenJS runtime does not enforce its documented whole-number constraint; adapter round-trips those invalid bytes, but the strict native snapshot does not expose them as supported revision values.

`CreatePresentationOptions.company` and `document.company` own only direct `Company` in the OOXML extended-properties part. Native omitted creation preserves the canonical app-properties bytes and returns `undefined`; `''` writes an explicit empty Company, and assigning `undefined` removes only that direct field. Values are strict XML-safe strings with correct metacharacter escaping. Reads and edits follow the package-root extended-properties relationship, accept any legal part URI plus default or prefixed namespace form, and create `/docProps/app.xml` or the next free URI when the relationship is absent. New Company is inserted before the first conventional following property without reordering existing children. Same-value and absent-clear calls are exact bytes/journal no-ops; Application, AppVersion, PresentationFormat, statistics, vectors, link state, unknown children, relationships, and unrelated parts remain unchanged. Dangling, external, wrong-type, malformed, duplicate-relationship, or duplicate-Company ownership is rejected before mutation. PptxGenJS 4.0.1 defaults company to `PptxGenJS` and its XML-safe custom/empty outputs import exactly, but it directly interpolates company without XML escaping; native intentionally escapes metacharacters instead of reproducing malformed output.

Inputs: `Uint8Array`, `ArrayBuffer`, `Blob`/`File`, Web `ReadableStream`, or async byte iterable. Node.js additionally accepts a file path or Node readable stream. `write()` returns `Uint8Array`; browsers can use `writeBlob()` or `download()`.

## Presentation format

`document.format` is detected from the presentation part content type and is one of `pptx`, `pptm`, `ppsx`, `ppsm`, `potx`, or `potm`. `document.formatProfile` also reports whether the package is macro-enabled, a slideshow, or a template. Unknown presentation content types are rejected instead of being treated as `.pptx`.

## Semantic model

```ts
document.slides[0].title.text = 'Updated';
document.slides[0].shapes[0].setTransform({ x: inches(1.5) });
document.duplicateSlide(0);
document.moveSlide(1, 0);
document.deleteSlide(2);
```

Plain speaker notes can be created and edited through the live slide model:

```ts
const slide = document.addSlide();
slide.addNotes('Opening context\nKey talking point');
slide.notes = 'Revised talking point';
slide.notes = '';
slide.notes = undefined;
```

`SlideModel.notes` returns `string | undefined`: absence is lazy `undefined`, an explicit empty body is `''`, and `undefined` assignment clears only the selected slide's notes relationship and owned notes part. Both the property setter and chainable `addNotes(string)` accept only XML-safe strings, normalize CRLF/CR to LF, and preserve leading/trailing whitespace. Reads follow the unique internal slide→notesSlide relationship, validate its slide backlink and shared notes-master chain, and flatten the unique direct body placeholder to plain text without mutating the package. Same-value assignment and clearing an absent value are exact byte/journal no-ops. Creation repairs a safely missing body placeholder and can create one canonical notes master only from fully absent, unambiguous topology using the presentation theme or first ordered slide-master theme; partial or ambiguous ownership is rejected before mutation. Duplication clones and retargets the per-slide notes part while retaining the shared master, deletion garbage-collects only unreferenced per-slide notes, and move/sections/hidden state remain independent. PptxGenJS 4.0.1 public output eagerly materializes empty notes for an omitted call; native creation intentionally distinguishes that from lazy `undefined`. This API is plain-text only and does not edit rich notes, notes-page layout, comments, header/footer/date fields, or slide numbers.

Preset shapes can be created with ordered adjustments plus direct fill, line, arrow, shadow, and hyperlink state,
and those direct values can be read or replaced:

```ts
const slide = document.addSlide();
const shape = slide.addShape('roundRect', {
  x: inches(1),
  y: inches(1),
  width: inches(3),
  height: inches(2),
  rotation: degrees(15),
  name: 'Feature card',
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 20,
  },
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '1F4E78' },
    transparency: 10,
    width: 2.5,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
  shadow: {
    kind: 'outer',
    color: { kind: 'srgb', value: '000000' },
    opacity: 0.35,
    blur: 6,
    angle: 45,
    distance: 4,
  },
  hyperlink: {
    url: 'https://example.com/docs',
    tooltip: 'Open documentation',
  },
});
shape.presetType = 'hexagon';
shape.fill = { kind: 'none' };
shape.line = { kind: 'line', color: { kind: 'scheme', value: 'accent2' } };
shape.arrows = { begin: 'diamond' }; // clears the omitted end
shape.arrows = { begin: 'none', end: 'oval' };
shape.arrows = undefined; // clears both endpoints, preserves line style
shape.shadow = {
  kind: 'inner',
  color: { kind: 'scheme', value: 'accent2' },
  opacity: 0.5,
  blur: 3,
  angle: 270,
  distance: 2,
};
shape.shadow = undefined;
document.addSlide(); // creates slide 2 as an internal-link target
shape.hyperlink = { slide: 2, tooltip: 'Go to details' };
shape.hyperlink = { url: 'mailto:team@example.com', tooltip: '' };
shape.hyperlink = undefined;
const arc = slide.addShape('blockArc', {
  adjustments: [
    { name: 'adj1', value: 16_200_000 },
    { name: 'adj2', value: 0 },
    { name: 'adj3', value: 25_000 },
  ],
});
arc.adjustments = [{ name: 'adj1', value: 10_800_000 }];
arc.adjustments = [];
console.log(shape.presetType); // 'hexagon'
console.log(PRESET_SHAPE_TYPES.length); // 178
```

`PRESET_SHAPE_TYPES` is a runtime-frozen tuple of the 178 canonical preset tokens; `PresetShapeType` is derived from it. `SlideModel.addShape(type, options?)` accepts `name`, strict `adjustments`, strict `fill`, strict `line`, strict `arrows`, strict `shadow`, strict `hyperlink`, and `Partial<Transform>` in native EMU and OOXML-angle units. Omitted x/y/width/height are each one inch, rotation is zero, both flips are false, the default name is `Shape ${id}`, omitted fill creates direct `a:noFill`, and omitted line keeps an empty direct `a:ln`. Options must be ordinary or null-prototype objects with supported own data properties; accessors, inherited/unknown/symbol keys, invalid XML strings, unsafe numbers, non-positive extents, invalid rotation, and non-boolean flips are rejected before package mutation.

`ShapeModel.presetType` returns a canonical token only for one namespace-correct direct `p:spPr/a:prstGeom`; missing, unknown, nested, malformed, repeated, or qualified-lookalike geometry returns `undefined` without mutation. Assigning the current token preserves the exact geometry bytes, including adjustments. Assigning another canonical token whole-replaces only that direct geometry with one empty `a:avLst`, intentionally clearing stale adjustments while preserving non-visual identity, transform, fill, line, arrows, effects, extensions, text, shape order, and the live model object. Creation and replacement are transactional and remain isolated through duplication, rollback, all six formats, write, and reopen. PptxGenJS 4.0.1 public `ShapeType/addShape/write` output is compared semantically for every legal token. Its `folderCorner` value is invalid OOXML and reads as `undefined`; native uses valid `foldedCorner`. `custGeom` is not a preset token and is exposed through `addCustomShape()` / `customGeometry` instead.

`CustomGeometryPoint`, `CustomGeometryCommand`, `CustomGeometryPathFill`, `CustomGeometryPath`, and `CustomGeometry` describe the public custom-path tree consumed by `SlideModel.addCustomShape(geometry, options)` and returned by `ShapeModel.customGeometry`. All numeric values are strict direct safe integers: points, path extents, and arc radii are EMU, while arc angles are OOXML `1/60000°`; `inches()` and `degrees()` are explicit ergonomic converters. The command union contains `moveTo`, `lineTo`, `arcTo`, `quadraticBezierTo`, `cubicBezierTo`, and `close`. Repeated moves form multiple subpaths; multiple paths, empty command lists, negative point coordinates, zero angles, and optional path `fill` / `stroke` / `extrusionOk` flags are supported. Arc commands deliberately have no endpoint because OOXML derives it from the current point, radii, and angles.

Creation and replacement strictly normalize dense ordinary arrays and descriptor-safe own data fields, detach the caller, and reject unknown, missing, non-safe, non-positive extent/radius, or invalid sequence state before mutation. Getter snapshots are detached and deeply frozen. Assignment replaces the complete custom geometry; assigning the same semantic snapshot is an exact bytes/journal no-op, and `undefined` clear is intentionally not accepted. Setting `presetType` converts custom geometry to preset; setting `customGeometry` converts a preset shape back while preserving model identity, name, transform, styles, text, effects, relationships, and unrelated bytes. Existing non-path custom geometry with arbitrary guide formulas, adjustment handles, connection sites, or a custom text rectangle remains byte-preserved, but reads `undefined` and cannot be replaced through this focused API. Geometry evaluation remains pending.

PptxGenJS 4.0.1 legal `ShapeType.custGeom` points import as matching final native snapshots: first ordinary point becomes `moveTo`, later ordinary points become `lineTo`, `moveTo: true` starts another subpath, and arc/quadratic/cubic/close map to their direct command branches. Its runtime converts numeric values below 100 and numeric strings as inches, keeps numeric values at or above 100 direct, resolves percentages against the full slide rather than the shape, ignores arc endpoint `x/y`, omits unknown curve kinds, and can emit malformed first-curve/zero-radius/negative-radius/unsafe output. The adapter imports supported final paths and byte-preserves malformed shapes with `customGeometry === undefined`; native does not copy those heuristics or coercions and requires explicit direct safe-integer values.

`ShapeAdjustment` is the readonly `{ readonly name: string; readonly value: number }` direct guide value used by `AddShapeOptions.adjustments` and `ShapeModel.adjustments`. A value must be a safe integer and is written directly as the operand of `a:gd@fmla="val N"`; the API performs no shape-specific conversion, clamping, or range inference. Lists must be dense, ordered, and uniquely named. Normalization is descriptor-safe and getter-free, immediately detaches from the caller, and getter snapshots are detached and deeply frozen. Assignment replaces the whole ordered list, `[]` keeps an empty `a:avLst`, and the setter does not accept `undefined`; assigning the same list is an exact bytes/journal no-op. Only one namespace-correct direct preset geometry with one direct adjustment list and simple `val` guides is supported. Complex formulas, duplicate or ambiguous guides, wrong namespaces, unsafe integers, and custom geometry read as `undefined`; replacement throws `ModelParseError` before package changes. A different `presetType` resets adjustments, while same-type assignment preserves their exact bytes. Valid PptxGenJS 4.0.1 `rectRadius`, `angleRange`, and `arcThicknessRatio` output imports as the same final integer list. Native deliberately retains explicit zero and rejects PptxGenJS string coercion, zero truthiness loss, shortcut precedence, ignored thickness-without-angles, and malformed/unsafe passthrough. Custom geometry paths use the separate API above; arbitrary guide formulas, handles, connection sites, custom text rectangles, and geometry evaluation remain outside this adjustment API.

`ShapeFill` supports direct none or solid sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. `ShapeModel.fill` returns a detached direct-state snapshot; same-value assignment is an exact no-op, `{ kind: 'none' }` writes direct no-fill, and `undefined` clears only the direct fill choice. Existing gradient, picture, pattern, and group fills survive unrelated edits and can be explicitly replaced or cleared, but their creation remains outside this simple-fill API.

`ShapeLine` supports direct none or a solid sRGB/theme line with optional finite 0–100 transparency, optional 0–1584 point width, and optional `solid | dash | dashDot | lgDash | lgDashDot | lgDashDotDot | sysDash | sysDot` dash. Omitted width/dash materialize as 1pt/solid; zero width remains direct zero. `ShapeModel.line` returns a detached direct-state snapshot; same-value assignment is an exact no-op, `{ kind: 'none' }` writes direct line no-fill, and `undefined` clears only owned width/fill/dash while preserving the line container, arrowheads, joins, extensions, and unrelated attributes. Existing advanced line fills and custom dashes survive unrelated edits and can be explicitly replaced or cleared, but their creation remains outside this simple-line API.

`ShapeArrowType` is `none | arrow | diamond | oval | stealth | triangle`. `ShapeArrows` has optional readonly `begin` / `end` fields and is used by `AddShapeOptions.arrows` and `ShapeModel.arrows`. Inputs and getter results are detached; the setter is a whole replacement, so a missing side clears that endpoint, explicit `none` stays distinguishable from absence, and `undefined` or an empty object clears both. Same-value assignment is an exact no-op. Arrow edits preserve line width/fill/dash, joins, extensions, advanced line state, and unrelated attributes; line clear/edit operations preserve arrows. Arrows-only creation writes a line container with endpoints but does not synthesize color, width, or dash. A unique safe existing endpoint may carry legal `w` / `len` values `sm | med | lg`; type replacement preserves them lexically, but size is not returned or editable. Malformed, duplicate, reversed, wrong-namespace, or unsupported endpoints read as `undefined` and reject arrow mutation without package changes.

`ShapeShadow` is the strict outer/inner direct-state union used by `AddShapeOptions.shadow` and `ShapeModel.shadow`. Both branches accept optional `RichTextColor`, finite `0..1` opacity, `0..100` point blur, `0 <= angle < 360` degrees, and `0..200` point distance; only outer accepts `rotateWithShape`. Omitted fields normalize to black, 0.75, 8pt, 270°, 4pt, and outer rotate false, while every explicit zero remains zero. Inputs and nested colors are detached before mutation; getter snapshots are detached and deep-frozen. Assignment is a whole replacement, same-value assignment is an exact bytes/journal no-op, kind switches replace only the direct shadow child, and `undefined` removes only that child while retaining `effectLst` plus legal glow, preset-shadow, reflection, soft-edge, blur, or fill-overlay siblings. A malformed, ambiguous, wrong-namespace, `effectDag`, or unsafe schema-order state reads as `undefined` and rejects mutation without package changes. PptxGenJS 4.0.1 omitted shadow and `type: 'none'` map to native `undefined`; its `offset` is native `distance`. Native preserves explicit zero, honors outer rotation, emits legal inner XML, supports theme colors, and rejects invalid passthrough instead of copying PptxGenJS's falsy fallback, ignored rotate flag, malformed inner closing tag, or out-of-range output. Generic effect stacks, custom shadow transforms, preset-shadow editing, and shadow APIs for images, text creation, tables, charts, media, and other owners remain outside this focused shape API.

`Hyperlink` is the mutually exclusive `{ readonly url: string; readonly tooltip?: string } | { readonly slide: number; readonly tooltip?: string }` value used by `AddShapeOptions.hyperlink` and `ShapeModel.hyperlink`. A URL must be a non-empty XML-safe string; a slide number must be a one-based positive safe integer resolving to a current presentation slide at assignment time. Inputs must be descriptor-safe ordinary or null-prototype objects with exactly one target and no unknown keys. Getter results are detached frozen direct-state snapshots. Tooltip absence remains property absence, while direct empty remains `tooltip: ''`; assignment is a whole replacement, omitted tooltip clears only that attribute, and `undefined` removes the supported click element. Same-value assignment is an exact bytes/journal no-op. URL/slide switching reuses an unshared relationship or clones on write when its ID is referenced elsewhere, and clear or replacement garbage-collects only unreferenced relationships. Internal links retain target-part identity while slide insert/delete/reorder changes the reported one-based ordinal; duplicate self-links retarget to the duplicate, and deleting a target removes incoming DrawingML click/hover elements before deleting their relationships. Unsupported hover editing, extra action/sound/history state, duplicate/malformed click ownership, or dangling/wrong-type relationships are never guessed: reads return `undefined`, and writes reject without package changes. PptxGenJS 4.0.1 materializes omitted tooltip as direct empty and may console-ignore, coerce, duplicate, or dangle invalid runtime targets; native supports the valid final semantics but rejects those defects before mutation. External hyperlinks produce the expected portability warning rather than a package error. Text-run, table, image, chart, media, group, and graphic-frame hyperlink creation, hover links, action-only navigation, and relative/file safety policy remain outside this shape-level API.

Arrow size, cap/compound/alignment/join editing, generic/advanced effects, custom shadow transforms, non-shape shadow APIs, arbitrary custom-geometry formulas/handles/connections/text rectangles/evaluation, shape-text creation options, advanced line fill/custom dash creation, and percentage positions remain pending.

Shape kinds include `text`, `shape`, `image`, `table`, `chart`, `graphic-frame`, and `group`. Images expose embedded part URIs and replacement; tables support basic native creation plus rows/cells, cell text, borders, fill, margins, horizontal/vertical alignment, text-direction, and text-fit editing; charts expose cached series and lossless chart XML editing.

```ts
const text = document.addSlide().addText('Quarterly results\nQ4 forecast', {
  name: 'Heading',
  x: inches(1),
  y: inches(0.75),
  width: inches(8),
  height: inches(1),
  align: 'center',
  fit: 'shrink',
  paragraphIndent: 12,
  paragraphMarginLeft: 18,
  paragraphMarginRight: 18,
  valign: 'middle',
  vert: 'vert270',
  wrap: true,
});
text.text = 'Updated results\nApproved';
text.textFit = 'resize';
text.verticalAlignment = 'bottom';
text.textDirection = 'wordArtVert';
text.textWrap = false;
```

`addText()` creates plain-text paragraphs with name and transform options. CRLF and CR normalize to LF; consecutive and trailing line breaks remain empty paragraphs. Setting `.text` replaces the visible text using the first paragraph as the style template. Use the structured API below when run styles must remain distinct.

```ts
const rich = document.addSlide().addRichText([
  {
    indent: -12,
    marginLeft: 12,
    marginRight: 12,
    runs: [{ text: 'Direct non-list paragraph margin' }],
  },
  {
    align: 'right',
    rtl: false,
    marginLeft: false,
    marginRight: 18,
    indent: false,
    bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
    level: 2,
    spacing: { before: 6, after: 8, line: { kind: 'multiple', factor: 1.5 } },
    tabStops: [{ position: 2.5, alignment: 'decimal' }],
    runs: [
      { text: 'Revenue\t', style: { bold: true, fontSize: 24 } },
      { text: '+18%', style: { italic: true, lang: 'fr-CA', color: { kind: 'srgb', value: '00A651' }, transparency: 25 } },
    ],
  },
], {
  lang: 'en-US',
  paragraphIndent: 12,
  paragraphMarginLeft: 24,
  paragraphMarginRight: 24,
  rtlMode: true,
});
rich.richText = [{ runs: [{ text: 'Approved', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 50 } }] }];
```

`richText` is an immutable paragraph/run value snapshot. It reads, creates, and replaces each paragraph's alignment, RTL mode, direct left/right margins, signed ordinary indent, Unicode bullet or automatic numbering, list level, paragraph spacing, tab stops, and run styles. `AddTextOptions` supplies paragraph creation defaults; each `RichTextParagraph` can override them, with `rtl: false`, `marginLeft: false`, `marginRight: false`, `indent: false`, `bullet: false`, `level: 0`, `spacing: false`, or `tabStops: false` suppressing the corresponding default. `paragraphMarginLeft` / `marginLeft` and `paragraphMarginRight` / `marginRight` use points from 0 through 4032 and map only direct `pPr@marL` / `pPr@marR`; direct zero remains distinguishable from false/absence. `paragraphIndent` / `indent` maps signed `-4032..4032` points to direct `pPr@indent`: positive values indent the first line and negative values create a hanging indent. New ordinary paragraphs retain canonical direct zero, while `false` or an omitted field in a later rich-text replacement clears the direct value. Ordinary numeric indent and an active bullet cannot share one paragraph because list indentation owns `indent`; `indent: false` can suppress an outer ordinary default for a list paragraph. Numeric left margin has the same conflict on `marL`; right margin remains independent and can coexist with bullets or numbering. Margin names and the indent sign map directly to physical OOXML values and do not swap under RTL. `AddTextOptions.rtlMode` applies to every created plain/rich paragraph; paragraph `rtl` overrides it, and omitting `rtl` in a later `shape.richText` replacement clears the direct `pPr@rtl` value. RTL does not change alignment, `bodyPr@rtlCol`, presentation direction, or run order. Bullets support a custom character and 0–4032pt per-level indent. Numbering supports the 16 PptxGenJS styles, a 1–32767 start value, and indent. List `level` is zero-based from 0–8; nested bullet margin is `indent × (level + 1)`. Spacing uses points for before/after/exact and a factor such as `1.5` for multiple line spacing. Tab stop positions use inches and support left, center, right, and decimal alignment; `[]` is an explicit empty list. Run styles include font family, point size, direct language, bold, italic, sRGB/theme color, and soft breaks. `AddTextOptions.lang` is the plain/rich creation default, while `RichTextRunStyle.lang` overrides a run; omitted creation language uses `en-US`. The getter exposes only a non-empty direct `rPr@lang`, not inherited or alternate language. Setting `richText` preserves text-body metadata and unrelated same-position paragraph properties but intentionally replaces old runs and supported paragraph values; hyperlinks and advanced typography are separate capabilities.

`RichTextRunStyle.transparency` controls only the run's main text fill. It accepts a finite percentage from 0 (opaque) through 100 (fully transparent), rounded to the nearest 0.001%. Explicit zero remains a direct alpha override; omission writes no alpha, and omission in a later `shape.richText` replacement clears the old override. Without an explicit color it applies to the default `tx1` text color. The getter requires one unambiguous direct solid-fill color and alpha transform; it does not infer inherited transparency or reuse alpha from glow, outline, highlight, underline, shape, or table-cell fills.

Text-box `margin` values use points and accept one number, a `[top, right, bottom, left]` tuple, or a named object. `shape.textMargins` reads only direct `bodyPr` overrides as a named object; assigning a scalar/tuple/object replaces the four supported direct sides, while `undefined` or `{}` clears them. This is separate from ordinary paragraph `marginLeft` / `marginRight` / `indent` and list-owned hanging indent. Unlike PptxGenJS 4.0.1's asymmetric-tuple runtime bug, native tuple creation follows the documented TRBL order.

Text-box `valign` accepts `top`, `middle`, or `bottom`; omission creates an explicit middle anchor. `shape.verticalAlignment` reads or replaces only the direct text-body anchor, and assigning `undefined` removes that direct override while preserving margins, autofit metadata, and other text-body content.

Text-box `wrap` accepts a boolean; omission and true create explicit automatic wrapping, while false keeps text on an unwrapped line. `shape.textWrap` reads or replaces only the direct text-body wrapping token, and assigning `undefined` removes that direct override without changing fit, margins, vertical alignment, or text content.

Text-box `vert` accepts `eaVert`, `horz`, `mongolianVert`, `vert`, `vert270`, `wordArtVert`, or `wordArtVertRtl`. Omission writes no direction, so it remains distinct from explicit `horz`. `shape.textDirection` reads or replaces only an exact valid direct `bodyPr@vert`; assigning `undefined` clears that override while preserving unknown direction tokens during unrelated edits. Table-cell direction is a separate capability.

```ts
import {
  emuToInches,
  inches,
  type AddTableCellOptions,
  type AddTableCellInput,
  type AddTableOptions,
} from '@pptx/sdk';

const tableOptions: AddTableOptions = {
  name: 'Revenue table',
  x: inches(1),
  y: inches(1.25),
  width: inches(8),
  height: inches(2.25),
  columnWidths: [inches(2.5), inches(3.5), inches(2)],
  rowHeights: [inches(0.5), inches(0.75), inches(1)],
  align: 'center',
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 1.5,
    style: 'solid',
  },
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent4' },
    transparency: 20,
  },
  margin: { top: 9, left: 18 },
  textDirection: 'vert270',
  valign: 'middle',
};
const headerOptions: AddTableCellOptions = {
  align: 'center',
  fit: 'shrink',
  textDirection: 'vert',
  valign: 'top',
  margin: { top: 4, left: 8 },
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 1,
    style: 'dash',
  },
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
};
const tableRows: readonly (readonly AddTableCellInput[])[] = [
  [
    { text: 'Region', options: headerOptions },
    'Revenue',
    {
      text: 'Growth',
      options: {
        border: {
          bottom: {
            kind: 'line',
            color: { kind: 'srgb', value: '70AD47' },
            width: 2,
          },
        },
        fill: { kind: 'none' },
      },
    },
  ],
  ['East', { text: '$1.2M' }, '12%'],
  [{ text: 'West' }, '$980K', { text: '' }],
];
const createdTable = document.addSlide().addTable(tableRows, tableOptions);
const currentWidthsInches = createdTable.columnWidths?.map(emuToInches);
const currentHeightsInches = createdTable.rowHeights?.map(emuToInches);
createdTable.setColumnWidths([
  inches(2),
  inches(3),
  inches(3),
]);
createdTable.setRowHeights([
  inches(0.75),
  inches(1.25),
  inches(0.5),
]);
createdTable.setRowHeights([0, inches(1), 0]);
createdTable.setCellText(2, 2, '8%');
```

`addTable()` accepts a non-empty, dense, rectangular matrix whose cells are single-paragraph strings or strict plain-text objects shaped as `{ text: string, options?: { align?: TextAlignment; border?: TableCellBorderInput; fill?: TableCellFill; fit?: TextBoxFit; margin?: TextBoxMarginInput; textDirection?: TableCellTextDirection; valign?: TextBoxVerticalAlignment } }`; both forms may be mixed. `AddTableOptions.margin` and cell `options.margin` are point-only and accept one number for all sides, an exact `[top, right, bottom, left]` tuple, or a partial named object. Resolution is per side: canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. A cell scalar/TRBL replaces all table sides, while a partial cell object replaces only supplied sides; omitted, runtime-undefined, or empty cell margin inherits the table layer. Omitted, runtime-undefined, or empty table margin preserves the existing canonical bytes. Each supplied value is rounded to the nearest EMU and must fit signed Int32, so zero and negative direct margins are supported. Cell-level creation `valign` accepts only `top`, `middle`, or `bottom` and emits the selected physical cell's direct `tcPr@anchor` as `t`, `ctr`, or `b`; without a table default, omission or runtime `undefined` writes no anchor and preserves the same bytes as an equivalent string cell. `AddTableOptions.valign` accepts the same exact values and materializes its value onto every cell without a cell-level value; an explicit cell value wins. An omitted or runtime-undefined table value preserves existing bytes. Table- and cell-level `textDirection` accept exactly `horz`, `vert`, `vert270`, or `wordArtVert`. `AddTableOptions.textDirection` materializes its value onto every physical cell whose cell option is omitted or runtime `undefined`; any explicit cell value wins. Explicit cell `horz` therefore blocks a non-horizontal table default while still creating no direct direction attribute. With no inherited non-horizontal value, omitted, runtime `undefined`, and explicit `horz` create no direct `tcPr@vert`, while the other three values write their exact token. Omitted or runtime-undefined table direction preserves existing bytes. `AddTableOptions.fill` and cell `options.fill` use the strict `TableCellFill` whole value: `{ kind: 'none' }` or `{ kind: 'solid', color, transparency? }`, with a six-digit sRGB or theme color and finite 0–100 transparency rounded to 0.001%. A cell solid or none value completely overrides the table value; omitted, runtime-undefined, or empty cell options inherit a supplied table fill. Omitted or runtime-undefined table fill preserves the original bytes, while an empty fill object is invalid. Table align, border, fill, margin, textDirection, and valign defaults are materialized only on physical cells and are not retained as table metadata, so later cell replace/clear operations work on final direct state and never reapply a creation default. The `tcPr` start tag serializes `marL/marR/marT/marB` first, then optional `anchor` and `vert`, followed by L/R/T/B borders and fill. `AddTableOptions.border` and cell `options.border` accept one side value for all four sides, an exact TRBL tuple, or a partial named object. A side is strict `{ kind: 'none' }` or `{ kind: 'line', color, width, style? }`: color is sRGB/theme, width is finite `0..1584` points quantized to the nearest EMU, and style is omitted, `solid`, or `dash`. A non-empty table value is copied only to cells with no effective cell border input; any non-empty cell scalar, TRBL, named value, or explicit none blocks the whole table value. A partial value defines only its own sides, and the remaining sides use the canonical direct no-fill baseline rather than inheriting across layers. Omitted/undefined/empty/all-undefined table border preserves existing bytes; omitted/undefined/empty/all-undefined cell border inherits a supplied table border. With no table border, those absent values and scalar none retain the canonical no-fill bytes. Omitted solid transparency writes no alpha and explicit zero writes `alpha=100000`. Cell, options, margin, border, side, fill, and nested color objects must be ordinary or null-prototype objects containing only supported own data properties. Accessors, inherited values, extra or symbol keys, exotic arrays, non-string text, CR/LF, and invalid XML characters are rejected without invoking a getter. Rows and all nested values are normalized into detached snapshots immediately. Geometry uses EMU: x/y default to 0.5 inch, width defaults to one inch per column, and height defaults to one inch with automatic rows (`a:tr@h=0`); use `inches()` for inch-based geometry. `columnWidths` and `rowHeights` each accept one positive EMU value repeated across the corresponding axis or a dense array whose length exactly matches the column/row count. Values are rounded to safe integers, input arrays are detached, and overall width/height are derived from the vector sum when omitted; when both an overall dimension and its vector are supplied, their rounded values must match exactly. Without a vector, explicit width/height retain their existing exact distribution behavior. Emitted explicit transform dimensions always equal the grid/row sum, while omitted `rowHeights` preserve automatic zero-height rows. New cells return a live `TableModel`; resolved direction, horizontal/vertical alignment, borders, fill, margins, and text fit are immediately visible as detached `TableCell.textDirection`, `TableCell.horizontalAlignment`, `TableCell.verticalAlignment`, `TableCell.borders`, `TableCell.fill`, `TableCell.margins`, and `TableCell.textFit` values, then editable through the physical-cell setters. Cell options other than align/border/fill/fit/margin/textDirection/valign—including merge, hyperlinks, and rich or multi-paragraph text—plus table-level fit creation/defaults, table-level direction/border/margin/valign/fill getters or editors, diagonal/advanced borders, advanced fills, table creation styles, auto-page, repeated headers, and content measurement remain unsupported and are rejected rather than ignored.

`AddTableOptions.align` and `AddTableCellOptions.align` reuse the exact `TextAlignment` values `left`, `center`, `right`, and `justify`, mapped to each created cell's single paragraph direct `a:pPr@algn` tokens `l`, `ctr`, `r`, and `just`, respectively. A table value is materialized onto every cell that omits cell alignment or supplies runtime `undefined`; a valid cell value wins. When the table value is omitted or runtime-`undefined`, current bytes are preserved and no effective left token is synthesized. Final ownership is the physical cell paragraph's direct `a:pPr@algn`, never `tcPr`, `bodyPr`, or retained table metadata, so later cell clearing does not reapply the creation default. The direct token survives supported cell text/property edits, slide duplication, write, and reopen, and is exposed through the strict single-paragraph snapshot/editor described below. Supported table and cell values, including cell precedence, match PptxGenJS 4.0.1 final state; PptxGenJS silently drops an unknown table runtime value, whereas native creation throws `TypeError` before mutation. Rich or multi-paragraph cell alignment remains unsupported.

`TableModel.columnWidths` reads the unique direct `tblGrid` as a detached exact-EMU snapshot. A malformed or ambiguous grid returns `undefined` instead of guessing from the transform or cells. `setColumnWidths()` accepts a positive scalar broadcast or a dense descriptor-safe exact-length array, rounds each item to a safe EMU integer, rejects unsafe sums, and atomically updates both `gridCol@w` and `ext@cx`. A valid grid/transform mismatch is repaired; a numeric no-op preserves the original slide bytes and mutation journal. Unsafe existing grid or transform XML raises `ModelParseError` without mutation. Inherited `setTransform({ width })` still changes only the transform, so use `setColumnWidths()` when changing table width distribution.

`TableModel.rowHeights` reads all direct `tr@h` values as a detached exact-EMU snapshot; zero is a valid automatic row height. A malformed or ambiguous direct row vector returns `undefined` without guessing from transform height or cell content. `setRowHeights()` accepts a non-negative scalar broadcast or a dense descriptor-safe exact-length array, rejects raw negative values, and rounds each item to a safe EMU integer. When every target is positive, their safe exact sum is written to `ext@cy` and a valid rows/transform mismatch is repaired. When any target is zero, row tokens are updated but the already-valid transform height is preserved because the rendered automatic height cannot be derived from the numeric row sum. Numeric no-ops preserve original tokens, slide bytes, and the mutation journal; unsafe existing rows or transform XML raise `ModelParseError` without mutation. Creation remains stricter: explicit `addTable({ rowHeights })` values must be positive, while omitting `rowHeights` creates automatic rows. Inherited `setTransform({ height })` still changes only the transform, so use `setRowHeights()` to edit row tokens. Rich/multi-paragraph cell text and alignment, cell options other than align/border/fill/fit/margin/textDirection/valign, table-level fit creation/defaults and table-level direction/border/margin/valign/fill getters or editors, diagonal/advanced borders, advanced fills, colspan/rowspan or merge editing, row insertion/deletion, table/cell creation styles, auto-page/repeated headers, hyperlinks, and content measurement are not yet supported by native creation.

```ts
import {
  TableModel,
  type TableCellBorders,
  type TableCellFill,
  type TableCellTextDirection,
  type TextBoxFit,
  type TextBoxMargins,
  type TextBoxVerticalAlignment,
  type TextAlignment,
} from '@pptx/sdk';

const table = document.slides[0].shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
const current: TableCellTextDirection | undefined =
  table?.rows[0]?.cells[0]?.textDirection;
table?.setCellTextDirection(0, 0, 'vert270');
table?.setCellTextDirection(0, 1, 'wordArtVert');
table?.setCellTextDirection(0, 2, 'horz');
table?.setCellTextDirection(0, 2, undefined);
const currentHorizontalAlignment: TextAlignment | undefined =
  table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'left');
table?.setCellHorizontalAlignment(0, 1, 'center');
table?.setCellHorizontalAlignment(0, 2, 'right');
table?.setCellHorizontalAlignment(0, 3, 'justify');
table?.setCellHorizontalAlignment(0, 3, undefined);
const currentFit: TextBoxFit | undefined = table?.rows[0]?.cells[0]?.textFit;
table?.setCellTextFit(0, 0, 'shrink');
table?.setCellTextFit(0, 1, 'resize');
table?.setCellTextFit(0, 2, 'none');
table?.setCellTextFit(0, 2, undefined);
const currentAlignment: TextBoxVerticalAlignment | undefined =
  table?.rows[0]?.cells[0]?.verticalAlignment;
table?.setCellVerticalAlignment(0, 0, 'top');
table?.setCellVerticalAlignment(0, 1, 'middle');
table?.setCellVerticalAlignment(0, 2, 'bottom');
table?.setCellVerticalAlignment(0, 2, undefined);
const currentMargins: TextBoxMargins | undefined = table?.rows[0]?.cells[0]?.margins;
table?.setCellMargins(0, 0, 7.2);
table?.setCellMargins(0, 1, [3.6, 7.2, 10.8, 14.4]);
table?.setCellMargins(0, 2, { top: 4, left: 8 });
table?.setCellMargins(0, 2, undefined);
const currentBorders: TableCellBorders | undefined = table?.rows[0]?.cells[0]?.borders;
table?.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: 'FF0000' },
  width: 2,
  style: 'solid',
});
table?.setCellBorders(0, 1, [
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 1, style: 'dash' },
  { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 1.5 },
  { kind: 'none' },
  undefined,
]);
table?.setCellBorders(0, 2, { top: { kind: 'none' } });
table?.setCellBorders(0, 2, undefined);
const currentFill: TableCellFill | undefined = table?.rows[0]?.cells[0]?.fill;
table?.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'srgb', value: 'FF0000' },
});
table?.setCellFill(0, 1, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
});
table?.setCellFill(0, 2, { kind: 'none' });
table?.setCellFill(0, 2, undefined);
```

Table creation names the options `AddTableOptions.textDirection` and `AddTableCellOptions.textDirection`; both accept only `horz`, `vert`, `vert270`, or `wordArtVert`. The table value is materialized onto every physical cell whose cell value is omitted or runtime `undefined`, while any explicit cell value wins. Explicit cell `horz` therefore blocks a non-horizontal table default but still creates no direct attribute. Omitted, runtime-undefined, and resolved horizontal creation write no `tcPr@vert`; the three non-horizontal values write their exact token, never `bodyPr@vert`, and appear immediately in the live `TableCell.textDirection` snapshot. Omitted/runtime-undefined table direction preserves existing bytes, and explicit table `horz` produces the same direction bytes when cells do not override it. Inputs follow the same descriptor-safe, getter-free, detached normalization as other table options, and invalid or text-box-only tokens are rejected before mutation. Creation retains only final physical-cell direct state, not table metadata, so a later clear never re-inherits the original table value. `setCellTextDirection()` addresses physical zero-based row/cell positions and deliberately retains lossless editor semantics: explicit `horz` writes `vert="horz"`, while `undefined` clears the direct attribute. The getter requires one direct `tcPr`, one unqualified `vert`, and one exact public token; no table-level getter/editor or inheritance resolution is retained. PptxGenJS 4.0.1 produces the same supported table/cell final state and precedence, including horizontal collapse; its runtime invalid-token passthrough remains unsupported by the strict native API. This four-value table-cell API is separate from the seven-value text-box API.

Table-cell `horizontalAlignment` is an immutable `TextAlignment` snapshot for a strict cell containing exactly one direct text body, one direct paragraph, and one direct paragraph-properties element. It reads only the unqualified direct `pPr@algn` tokens `l`, `ctr`, `r`, and `just` as `left`, `center`, `right`, and `justify`; missing, malformed, ambiguous, zero-paragraph, or multi-paragraph cells return `undefined` without resolving effective alignment or a retained table default. `setCellHorizontalAlignment()` uses physical zero-based row/cell coordinates, writes the canonical direct token, and uses `undefined` to clear only that unqualified direct attribute. Assigning the current value or clearing an absent token is an exact no-op. Unsafe cell structure raises `ModelParseError` without mutation, while unrelated paragraph content, namespaced attributes, cell properties, and neighbor cells remain byte-preserved. PptxGenJS 4.0.1 table/cell alignment is materialized into the same direct cell-paragraph state and imports to the same snapshots; native existing-deck editing is a lossless extension because PptxGenJS has no existing-deck editor. Rich or multi-paragraph cell alignment remains unsupported.

Table-cell creation names the option `AddTableCellOptions.valign`; the immutable snapshot is `TableCell.verticalAlignment`. Both reuse exact `TextBoxVerticalAlignment` values `top`, `middle`, and `bottom`, stored only in the selected physical cell's direct `tcPr@anchor` as `t`, `ctr`, or `b`. Omitted or runtime-undefined cell values write no anchor unless `AddTableOptions.valign` supplies the same strict value as a creation default; a cell value overrides that default. The table value is materialized directly onto uncovered physical cells and is not retained as table metadata, so a later `setCellVerticalAlignment(..., undefined)` clears the anchor without inheritance. Omitting the table value preserves the equivalent default bytes. The returned live table exposes every resolved value immediately, and `setCellVerticalAlignment()` writes the canonical token or clears it with `undefined`. The strict snapshot getter requires exactly one direct `tcPr` and one unqualified supported anchor. It does not read or change the separate `bodyPr@anchor`, resolve inheritance, or support `just` / `dist`. PptxGenJS 4.0.1 produces the same final direct state for valid table- and cell-level values, including cell precedence and no anchor when both layers are omitted. Runtime invalid PptxGenJS tokens remain opaque and import as `undefined`, while the native API rejects them. A table-level getter/editor remains unsupported.

Table-cell `margins` reuses the point-based `TextBoxMargins` value shape but owns only direct `tcPr@marL/marR/marT/marB`; it never reads or changes text-box `bodyPr@*Ins`. `AddTableOptions.margin`, `AddTableCellOptions.margin`, and `setCellMargins()` accept a point scalar, `[top, right, bottom, left]` tuple, partial named object, `{}` or `undefined`, with descriptor-safe getter-free normalization and immediate detachment. Creation resolves canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. Cell scalar/TRBL replaces all four inherited sides; partial cell input replaces only supplied sides; omitted/undefined/empty cell input inherits the table layer. Omitted/undefined/empty table input preserves canonical bytes. The resolved values are materialized directly on physical cells and immediately exposed by `TableCell.margins`; no table metadata or getter/editor is retained, so clearing a cell does not reapply the creation default. The cell editor addresses zero-based physical row/cell positions and is a whole replacement, so its omitted named sides are cleared. Values quantize to signed-Int32 EMU and may be zero, negative, positive, or fractional points. The snapshot getter requires exactly one direct `tcPr`, reads each unique unqualified signed-Int32 integer independently, and returns only valid direct sides. Creation writes `marL/marR/marT/marB`, then optional `anchor`, followed by L/R/T/B borders and fill. PptxGenJS 4.0.1 writes the same narrow defaults and materializes table-level values into cells, but its legacy runtime treats the first value `<1` as inches and `>=1` as points. Thus native table or cell `7.2` equals PptxGenJS `0.1`, native `[3.6, 7.2, 10.8, 14.4]` equals PptxGenJS `[0.05, 0.1, 0.15, 0.2]`, while native `0.1` intentionally means 0.1pt. Adapter imports expose the resulting OOXML in points instead of guessing the original unit.

Table-cell `borders` is a detached partial snapshot of the selected physical cell's same-prefix direct `lnL/lnR/lnT/lnB`. Each supported side is `{ kind: 'none' }` or `{ kind: 'line', color, width, style? }`, where width is finite `0..1584` points, color is strict sRGB/theme, and style is omitted, `solid`, or `dash`. `AddTableOptions.border`, `AddTableCellOptions.border`, and `setCellBorders()` accept one scalar border, an exact `[top, right, bottom, left]` tuple, or a partial named object. During creation, a non-empty table value is materialized only onto cells whose normalized cell border is absent. Any non-empty cell scalar, TRBL, named value, or explicit none blocks the entire table value; missing sides in that cell value remain canonical none and do not inherit table sides. Empty/all-undefined cell values inherit a supplied table value, while empty/all-undefined table values preserve omitted bytes. Creation overlays the chosen value on four canonical direct no-fill sides; the editor is intentionally a whole replacement, so its missing sides are removed and `{}` / `undefined` clear all four direct sides without re-inheritance. Explicit none writes a direct zero-width `noFill` line. Public TRBL input is serialized in OOXML's required L/R/T/B order before cell fill; dash maps to `sysDash`, explicit solid maps to `solid`, omitted style writes no direct dash token, zero-width line remains a line, and widths are quantized to the nearest EMU. Border input at every layer is descriptor-safe, getter-free, null-prototype-compatible, and detached. The strict getter omits malformed or unsupported sides independently and never treats diagonals, cell fill, text outlines, advanced dash presets, joins, arrows, or shared-edge/effective style as supported state; unrelated edits preserve them. For supported table/cell scalar or TRBL none, sRGB line, solid/dash, scalar zero width, and whole-value cell overrides, native creation matches PptxGenJS 4.0.1 final direct state. Native `{}` means no effective input while PptxGenJS `{}` defaults to `666666`/1pt/solid and a cell `{}` blocks its table value; native omitted style writes no direct dash while PptxGenJS defaults it to direct solid. Native scalar and TRBL zero widths stay zero, whereas PptxGenJS table TRBL `pt: 0` changes to 1pt; PptxGenJS also accepts short runtime tuples and materializes missing sides as none, while native tuples must be dense and exactly four items. Native named sides and theme border colors are strict extensions. Adapter snapshots reflect final XML rather than guessing the original input layer. A table-level border getter/editor, diagonal borders, transparency, advanced line properties, shared-edge precedence, and layout recomputation remain unsupported.

Table-cell `fill` is a detached direct-state snapshot with explicit `{ kind: 'none' }` and `{ kind: 'solid', color, transparency? }` variants. Solid fill supports strict six-digit sRGB or the existing theme-color tokens; transparency is a finite `0..100` percentage quantized to `0.001%`. Omitted transparency writes no alpha, while explicit zero writes direct opaque alpha. During native creation, `AddTableOptions.fill` is a whole-value default materialized into every physical cell whose `AddTableCellOptions.fill` is omitted or `undefined`; a cell solid or none value completely overrides it. Table and cell fill/color inputs are descriptor-safe, getter-free, ordinary/null-prototype-only, and deeply detached immediately. Omitted or runtime-undefined table fill preserves the original bytes; `{}` is invalid. The returned live table immediately exposes detached resolved `TableCell.fill` values. `setCellFill()` then addresses only zero-based physical row/cell direct state: `none` writes `tcPr/a:noFill`, while `undefined` removes the direct fill choice and never re-inherits the creation default. The strict getter requires a unique direct `tcPr` and one unambiguous same-prefix noFill or solidFill; it never reads border/text descendant fills or resolves table styles. Unsupported gradient/pattern/picture/group fills remain preserved during unrelated edits and can be explicitly replaced or cleared; a table-level fill getter/editor is not supported. PptxGenJS 4.0.1 produces the same final direct cell state for supported table solid fills and cell solid overrides, but collapses omitted and `type: 'none'` to no direct fill and explicit zero transparency to no alpha. Native direct none and explicit zero therefore remain intentionally distinguishable. PptxGenJS may also emit invalid alpha for out-of-range runtime values; adapter imports preserve that XML but do not fabricate a valid snapshot.

Table-cell `textFit` reuses the immutable `TextBoxFit` values `none`, `shrink`, and `resize`, backed only by the selected physical cell's direct `txBody/bodyPr` fit choice. Native plain-text cell creation accepts `AddTableCellOptions.fit` at cell level only: omitted, runtime `undefined`, and `none` produce byte-identical self-closing `bodyPr` and an immediate direct snapshot of `undefined`; `shrink` and `resize` write direct `normAutofit` and `spAutoFit` and snapshot immediately. Creation never writes `noAutofit`, retains table metadata, measures text, calculates a final font scale, or recomputes cell/table geometry. The getter requires exactly one direct text body, one direct body properties element, and one unambiguous supported fit child; it reads an existing `noAutofit` as `none` and otherwise returns `undefined` for absent or malformed state. `setCellTextFit()` uses physical zero-based row/cell positions: `shrink` and `resize` write `normAutofit` and `spAutoFit`, while `none` and `undefined` both clear the direct choice without restoring creation input or creating `noAutofit`. Reassigning the current shrink/resize mode preserves any calculated `fontScale` and `lnSpcReduction`. Fit remains independent from text direction in `tcPr@vert`. Table-level fit creation/defaults are not supported. PptxGenJS 4.0.1 has no table-cell fit API and ignores runtime `fit`, `autoFit`, and `shrinkText` values supplied to table/cell options, so native shrink/resize creation is an explicit extension rather than a parity claim.

Text-box `fit` accepts `none`, `shrink`, or `resize`. Omission and none write no autofit child for PptxGenJS and PowerPoint 2013 compatibility; shrink writes `normAutofit`, and resize writes `spAutoFit`. `shape.textFit` reads only a unique direct choice, including an existing explicit `noAutofit` as none. Assigning none or `undefined` clears the direct choice; reassigning the current shrink/resize mode preserves any PowerPoint-calculated scale attributes. Final shrink factors may be calculated by PowerPoint only after editing text or resizing the shape.

Slide and shape model objects have stable identity within a document: repeated collection reads and slide reordering return the same member instances. Their properties remain live and read the current OOXML rather than a cached snapshot. Master, layout, and theme models follow the same rule within `document.masterLayoutTheme`.

`duplicateSlide()` deep-clones owned dependencies such as charts and their embedded workbooks while retaining shared image, media, and layout targets. `deleteSlide()` garbage-collects only unreferenced owned dependency subgraphs; opaque and shared targets are preserved.

`ImageModel.replaceData()` and `ChartModel.setXml()` edit an exclusive target in place. If a target or relationship is shared, the edited shape is redirected to a private clone; chart clones include owned workbook/style/color dependencies. The operation is atomic and keeps the shape model object identity stable.

## Transactions

```ts
document.transaction((draft) => {
  draft.slides[0].title.text = 'Updated';
  draft.duplicateSlide(0);
});
```

Transactions are synchronous and nestable. A thrown error or package validation failure restores parts, content types, relationships, ZIP entries, and the mutation journal to the transaction savepoint. Complete asynchronous preparation before entering the callback.

## Gradients and transparency

```ts
document.slides[0].background = {
  kind: 'linear-gradient',
  angle: 45,
  stops: [
    { offset: 0, color: '#2563EB' },
    { offset: 1, color: '#7C3AED', alpha: 0.65 },
  ],
};
```

Colors can use sRGB, scRGB, scheme, system, or preset sources. OOXML transforms retain their original order.

## Master, layout, and theme

```ts
const theme = document.themes[0];
theme.setColor('accent1', '#2563EB');

const chain = document.masterLayoutTheme.materializeInheritedStyle(
  document.slides[0].partUri,
  document.slides[0].shapes[0].id,
);
```

`masterLayoutTheme` also exposes create/copy/delete/relink operations for masters, layouts, and themes.

## Media

```ts
await document.addAudio(0, audioBuffer, {
  contentType: 'audio/mpeg',
  play: 'click',
  volume: 0.8,
});

await document.addVideo(0, 'https://example.com/video.mp4');
```

Local media can come from paths, bytes, ArrayBuffers, or streams. External URLs are never fetched automatically and produce portability diagnostics.

## Diagnostics and errors

Errors: `PackageError`, `ParseError`, `ValidationError`, `OpaqueMutationError`, `PptxGenJSAdapterError`.

Every diagnostic has severity, code, message, and optional part URI, XML path, object id, compatibility profile, and suggestion. Strict mode blocks only error diagnostics.
