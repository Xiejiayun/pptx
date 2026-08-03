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

### Runtime version

```ts
import { PPTX_VERSION, PptxDocument, type PptxVersion } from '@pptx/sdk';

const current: PptxVersion = PPTX_VERSION; // '0.1.0'
const document = PptxDocument.create();
document.version satisfies PptxVersion;
```

`PPTX_VERSION` is the browser-safe compile-time package version, `PptxVersion` is its literal type, and `PptxDocument.version` is getter-only. Created and opened documents retain that value before and after write/reopen without reading a manifest at runtime or changing OPC state. This value identifies the current library runtime; it is not OOXML `AppVersion`, file-producer metadata, or a PowerPoint compatibility version.

PptxGenJS 4.0.1 reports `'4.0.1'` and native 0.1.0 reports `'0.1.0'`; those values should differ because each instance identifies its own library. Repository tests synchronize the constant with the root, SDK, and aggregate manifests, while CLI `--version` and JSON doctor reuse the same constant. The actual 58-file package (SHA-256 `ce300d3c5da10a8fbdb9910b10497d02af496532b99329e18a314c9604e6f9a8`) passes installed Node, declaration, browser-conditional, CLI, and real-Google-Chrome create/writeBlob/reopen coverage with `presentationVersion: true` and zero Chrome validation/console/page/network errors. Final gates are 1354 passed / 1 skipped tests, performance 1/1 at 617ms, both TypeScript checks, both bundles, and declaration generation.

This historical checkpoint does not complete PptxGenJS parity. The `OutputType` runtime catalog is completed in a later section; six actual `write({ outputType })` return semantics, stream, compression, remaining runtime constants, advanced text/table, `tableToSlides`, and the final peer/client audit remain pending.

### Presentation layout projection

```ts
import {
  inches,
  PptxDocument,
  type PresentationLayout,
  type PresentationLayoutName,
} from '@pptx/sdk';

const document = PptxDocument.create({ slideSize: '4:3' });
const layout: PresentationLayout = document.presLayout;
const name: PresentationLayoutName = layout.name; // 'screen4x3'

document.slideSize = { width: inches(11.7), height: inches(8.3) };
document.presLayout; // { name: 'custom', width: 10698480, height: 7589520 }
```

`PptxDocument.presLayout` is getter-only and derives a new detached `{ name, width, height }` snapshot from `slideSize` on every read. Dimensions use EMU. Exact standard dimensions map to `screen4x3`, `screen16x9`, or `screen16x10`; all other legal dimensions, including wide, map to `custom`. Reads do not mutate the package, invalid `p:sldSz` reuses the strict `slideSize` error, failed edits roll back, and successful edits survive write/reopen.

PptxGenJS 4.0.1 public runtime uses the same EMU values. Native intentionally omits the undeclared `_sizeW` / `_sizeH` fields and mutable internal alias. A PptxGenJS custom registry name is not serialized into PPTX, so native reports canonical `custom` after open/reopen. Installed Node, declarations, browser conditional export, CLI package inspection, and real Chrome report `presentationLayouts: true`; the final 59-file tarball SHA-256 is `a07a11156840071f0945289c0a48fdd9741549d2003ca21006e6efab28104b3d`. Final gates are 1363 passed / 1 skipped tests, performance 1/1 at 1.01s, both TypeScript checks, both bundles, declaration generation, and zero Chrome validation/console/page/network errors.

This historical checkpoint does not complete PptxGenJS parity. The `OutputType` runtime catalog is completed in a later section; six actual `write({ outputType })` return semantics, stream, compression, remaining runtime constants, advanced text/table, `tableToSlides`, and the final peer/client audit remain pending.

### Horizontal text alignment catalog

```ts
import { TEXT_ALIGNMENTS, type TextAlignment } from '@pptx/sdk';

const alignments: readonly TextAlignment[] = TEXT_ALIGNMENTS;
```

`TEXT_ALIGNMENTS` is exactly the frozen readonly tuple `['left', 'center', 'right', 'justify']`, and `TextAlignment` is derived as `(typeof TEXT_ALIGNMENTS)[number]`. SDK and aggregate-root exports reuse the model value rather than creating facade copies. These tokens continue to drive plain/rich text plus table/table-cell alignment through the unchanged `l`, `ctr`, `r`, and `just` OOXML mappings; catalog reads never mutate an OPC package.

PptxGenJS 4.0.1 public `AlignH` keys and values match the tuple in the same order. Native deliberately exposes the immutable catalog instead of an instance getter or mutable enum-shaped alias. Installed Node, declarations, browser conditional export, CLI package inspection, and real Chrome report `horizontalAlignments: true`; create and reopen preserve all four values. The final 59-file tarball SHA-256 is `46a88495acbffb2f81eb99f290bd15928974ddad86f507941c1ea5bfb65eaa90`. Final gates are 1368 passed / 1 skipped tests, performance 1/1 at 1.17s, both TypeScript checks, both bundles, declaration generation, and zero Chrome validation/console/page/network errors.

Overall PptxGenJS parity at this checkpoint was approximately 96%; `AlignV` is completed in the next section.

### Vertical text alignment catalog

```ts
import {
  PptxDocument,
  TEXT_VERTICAL_ALIGNMENTS,
  type TextBoxVerticalAlignment,
} from '@pptx/sdk';

const document = PptxDocument.create();
const slide = document.addSlide();
const alignment: TextBoxVerticalAlignment = TEXT_VERTICAL_ALIGNMENTS[1];

slide.addText('Centered vertically', { valign: alignment });
slide.slideNumber = { valign: alignment };
slide.addTable([[{ text: alignment, options: { valign: alignment } }]]);
```

`TEXT_VERTICAL_ALIGNMENTS` is exactly the frozen readonly tuple `['top', 'middle', 'bottom']`, and `TextBoxVerticalAlignment` is derived as `(typeof TEXT_VERTICAL_ALIGNMENTS)[number]`. SDK and aggregate-root exports reuse the model value rather than creating facade copies. The tokens continue to drive text-box, slide-number, and table/table-cell alignment through the unchanged `t`, `ctr`, and `b` OOXML mappings; catalog reads never mutate an OPC package.

PptxGenJS 4.0.1 public `AlignV` keys and values match the tuple in the same order. Native deliberately exposes the immutable catalog instead of an instance getter or mutable enum-shaped alias. Installed Node, declarations, browser conditional export, CLI package inspection, and real Chrome report `verticalAlignments: true`; text and table create/reopen preserve all three values. The final 59-file tarball SHA-256 is `aaaa5e0ceb053a472af49732784e0ea5babb00968734ec5093cd4f80afc34095`. Final clean gates are 72 passed / 1 skipped test files, 1373 passed / 1 skipped tests, performance 1/1 at 939ms, both TypeScript checks, both bundles, declaration generation, and zero Chrome validation/console/page/network errors.

Overall PptxGenJS parity at this checkpoint is approximately 97%; the `OutputType` runtime catalog is completed in the next section.

### Presentation output type catalog

```ts
import { OUTPUT_TYPES, type OutputType } from '@pptx/sdk';

const outputTypes: readonly OutputType[] = OUTPUT_TYPES;
```

`OUTPUT_TYPES` is exactly the frozen readonly tuple `['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']`, and `OutputType` is derived as `(typeof OUTPUT_TYPES)[number]`. The SDK output layer owns the value, while the aggregate root reuses the same object. Catalog discovery is environment-independent and produces no OPC mutation.

PptxGenJS 4.0.1 public `OutputType` keys and values match the tuple in the same order. Native deliberately exposes the immutable catalog instead of an instance getter or mutable enum-shaped alias. `STREAM` is excluded because it belongs to the separate stream API rather than the public instance enum. This historical catalog checkpoint does not change write behavior; selectable return types are completed below.

Installed Node, declarations, browser conditional export, CLI package inspection, and real Chrome report `outputTypes: true`. The final 60-file tarball SHA-256 is `31a38643c8c851ae24a381a68cd225972b76dbf7b37758c16efd2fe27248df0d`. Final gates are 73 passed / 1 skipped test files, 1378 passed / 1 skipped tests, performance 1/1 at 1.10s, both TypeScript checks, both bundles, declaration generation, Chrome HTTP 200 responses, and zero Chrome console/page/network errors.

Overall PptxGenJS parity remains approximately 97%. Six-value `write({ outputType })` return semantics are completed below.

### Presentation write output types

```ts
import {
  PptxDocument,
  type OutputType,
  type WriteBaseOptions,
  type WriteOptions,
  type WriteOutput,
} from '@pptx/sdk';

const document = PptxDocument.create();
const defaultBytes = await document.write(); // Uint8Array
const arrayBuffer = await document.write({ outputType: 'arraybuffer' });
const base64 = await document.write({ outputType: 'base64' });
const binaryString = await document.write({ outputType: 'binarystring' });
const zipBlob = await document.write({ outputType: 'blob' });
const nodeBuffer = await document.write({ outputType: 'nodebuffer' });
const bytes = await document.write({ outputType: 'uint8array' });

async function encode<T extends OutputType>(
  options: WriteOptions<T>,
): Promise<WriteOutput<T>> {
  return document.write(options);
}
```

`WriteBaseOptions` contains validation options shared by every output method. `WriteOptions<TOutputType = 'uint8array'>` adds `outputType?: TOutputType`, and `write<T>()` returns `Promise<WriteOutput<T>>`. Omitted options, `{}`, and `WriteBaseOptions` preserve the default `Promise<Uint8Array>` contract. Literal `arraybuffer` maps to `ArrayBuffer`; `base64` and `binarystring` map to `string`; `blob` maps to `Blob`; `nodebuffer` and `uint8array` map structurally to `Uint8Array`. At Node runtime, `nodebuffer` is a `Buffer`; the structural declaration deliberately avoids importing `node:buffer` into browser consumers.

Raw base64 has no data-URI prefix. Binary strings use one byte per code unit. ArrayBuffers are standalone, not views over a larger backing store. Explicit Blob output has MIME `application/zip`, while `writeBlob(options?: WriteBaseOptions)` retains the presentation MIME. Browser `nodebuffer` requests reject exactly with `nodebuffer is not supported by this platform`. All supported conversions use the same canonical ZIP bytes and do not mutate diagnostics or package state.

Installed Node, declarations, browser conditional export, CLI package inspection, and real Chrome report `writeOutputTypes: true`. The final 61-file tarball SHA-256 is `26bbc7eb7c33eb194388576db2c2eaab33c80d0d99b19ed7a9b4a7375c3f9f37`. Final gates are 74 passed / 1 skipped test files, 1383 passed / 1 skipped tests, performance 1/1 at 966ms, both TypeScript checks, both bundles, declaration generation, byte-identical/reopen checks for every available output, and zero Chrome validation/console/page/network errors.

Overall PptxGenJS parity remains approximately 97%. Node readable output is completed below.

### Node readable output

```ts
import { createWriteStream } from 'node:fs';
import {
  PptxDocument,
  type PptxNodeReadableStream,
  type WriteBaseOptions,
} from '@pptx/sdk';

const document = PptxDocument.create();
document.addSlide().addText('Stream output');

const options: WriteBaseOptions = { mode: 'strict' };
const readable: PptxNodeReadableStream = await document.stream(options);
for await (const chunk of readable) console.log(chunk.byteLength);

(await document.stream()).pipe(createWriteStream('output.pptx'));
```

`PptxNodeReadableStream` is a browser-safe structural declaration over the core binary Node Readable surface: async iteration, `pipe`, data/end/close/error listeners, pause/resume/isPaused/read/destroy, and readable state. Runtime values are actual `node:stream` `Readable` instances with `readableObjectMode === false`; declarations do not import `node:stream`, `node:buffer`, `NodeJS`, or Buffer types. Each internal chunk is at most 65,536 bytes.

`stream()` checks the Node runtime before validation or ZIP generation, then calls the same `#writeBytes(options)` path as other output methods and dynamically loads `node:stream`. Concatenated output is byte-identical to `write()` and reopens. The stream captures document state when its Promise resolves; later mutations affect only later writes. Consumer destroy and delivery do not mutate package state or diagnostics.

The complete canonical ZIP bytes are still generated before the Readable is returned, so peak ZIP-generation memory and time-to-first-byte do not improve. Browser calls reject exactly with `PptxDocument.stream() is only supported in Node.js`. PptxGenJS 4.0.1's method of the same name returns Buffer; native matches that separate result through `write({ outputType: 'nodebuffer' })` and provides real stream semantics here.

Installed Node, declarations, browser conditional export, CLI package inspection, and real Chrome report `nodeReadableStream: true`. The final 61-file tarball SHA-256 is `37b1d6bec7b5a144d577c57b61c0777f2aad8515015e9cbee05abd55f8e067d2`. Final gates are 75 passed / 1 skipped test files, 1390 passed / 1 skipped tests, performance 1/1 at 682ms, both TypeScript checks, Node/browser bundles, multi-chunk async/pipe byte equality, successful reopen, exact Chrome rejection/isolation, and zero Chrome validation/console/page/network errors.

Overall PptxGenJS parity remains approximately 97%. Compression policy is completed below.

### Presentation ZIP compression policy

```ts
const deflated = await document.write({
  outputType: 'uint8array',
  compression: true,
});
const stored = await document.write({ compression: false });
const readable = await document.stream({ compression: true }); // Node.js
await document.writeFile('output.pptx', { compression: true }); // Node.js
const browserBlob = await document.writeBlob({ compression: false });
await document.download('output.pptx', { compression: true }); // browser
```

`WriteBaseOptions.compression?: boolean` is shared by `write()`, `stream()`, `writeFile()`, `writeBlob()`, and `download()`. For created or modified documents, omitted/`undefined` and `false` select ZIP STORE, while `true` selects DEFLATE level 6. Every `WriteOptions<T>` output representation uses the same package policy. Non-primitive booleans reject before validation replacement or OPC generation with `PptxDocument output compression must be a boolean`, leaving diagnostics and package mutations unchanged.

Opened documents have a single preservation exception: if the package remains unchanged and compression is omitted or `undefined`, the original bytes are returned exactly. Explicit `false` or `true` bypasses that fast path and regenerates STORE or DEFLATE output. This keeps lossless editing as the default while making caller-selected compression deterministic.

PptxGenJS 4.0.1's legal primitive-boolean intent is supported. Native deliberately keeps compression orthogonal to explicit `outputType` and rejects truthy non-booleans instead of copying upstream's explicit-output omission and coercion behavior. Installed Node/types/browser/CLI and real Chrome report `compressionPolicy: true`. The final 61-file tarball SHA-256 is `4bbaa25b83a0d20dd3d2239708c628afec79bcff19c69faca6fe67b03e3bd990`; packed Node STORE/DEFLATE sizes are 149,598/9,347 bytes, Chrome sizes are 84,062/9,270 bytes, download uses method 8 and reopens, and Chrome console/page/network errors are zero. Final gates are 1400 passed / 1 skipped tests, performance 1/1 at 749.5ms, both TypeScript checks, both bundles, and declaration generation.

Overall PptxGenJS parity remains approximately 97%. The `SchemeColor` runtime helper is completed below.

### Presentation scheme-color catalog

```ts
import {
  PptxDocument,
  SCHEME_COLORS,
  type SchemeColor,
} from '@jiayunxie/pptx';

const text: SchemeColor = SCHEME_COLORS.text1;
const background: SchemeColor = SCHEME_COLORS.background1;
const accent: SchemeColor = SCHEME_COLORS.accent1;
const document = PptxDocument.create();

document.addSlide().addRichText([{
  runs: [
    { text: 'Theme text', style: { color: { kind: 'scheme', value: text } } },
    { text: ' accent', style: { color: { kind: 'scheme', value: accent } } },
  ],
}], { fill: { kind: 'solid', color: { kind: 'scheme', value: background } } });
```

`SCHEME_COLORS` is a frozen `Readonly` mapping with exact order and values `text1→tx1`, `text2→tx2`, `background1→bg1`, `background2→bg2`, and `accent1..accent6`. `SchemeColor` is the union of those ten values. The model, SDK, and root export one object identity in Node and browsers; catalog access is independent of document/package state.

This matches PptxGenJS 4.0.1's public `SchemeColor` keys, values, order, and legal output without copying its prototype getter or shared mutable enum object. `SchemeColor` is intentionally the PptxGenJS helper union, not the exhaustive native theme-color type: native color APIs retain their wider validated DrawingML token subset.

Final gates are 77 passed / 1 skipped test files, 1404 passed / 1 skipped tests, performance 1/1 at 736ms, both TypeScript checks, both bundles, declaration generation, and installed Node/types/browser/CLI checks. The actual 62-file tarball SHA-256 is `5d7096b0347d605c105dff15bb357781c4dcaa1cb7c3eff69f89ea6baa70e742`; Node and real Chrome report `schemeColors: true`, with successful write/reopen, frozen/shared identity, mutation isolation, and zero Chrome validation/console/page/network errors.

## Table-level vertical alignment

```ts
import { PptxDocument, type TextBoxVerticalAlignment } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { valign: 'middle' });

const uniform: TextBoxVerticalAlignment | undefined = table.verticalAlignment; // middle
table.setCellVerticalAlignment(0, 1, 'top');
const mixed = table.verticalAlignment; // undefined
table.verticalAlignment = 'bottom';
table.verticalAlignment = undefined;
```

`TableModel.verticalAlignment` reads consensus from every direct physical-cell `tcPr@anchor`: it returns one safe shared value, otherwise `undefined` for mixed, absent, empty, or unsafe state. Setting `top`, `middle`, or `bottom` atomically replaces every physical cell, including merge continuations; setting `undefined` clears every direct anchor. Exact no-ops do not change slide bytes or the mutation journal. This is not retained table metadata or an inherited default, so callers that need mixed detail inspect `rows[].cells[].verticalAlignment`. Unsafe edits raise `ModelParseError` without partial package mutation and preserve all unrelated cell/table state.

PptxGenJS 4.0.1 provides creation options but no existing-deck table object model; supported creation output imports into the same final direct anchors, and the native bulk editor is a lossless extension. Final verification is 4 focused files / 521 focused tests, 78 passed / 1 skipped full test files, 1411 passed / 1 skipped full tests, and performance 1/1 at 885ms. The actual 62-file tarball SHA-256 is `6ce48d8bb73d59148754f14dc379b9cd11ba34d358dd8e7ebba7b72cf8208f1e`; installed Node/types/browser/CLI and real Chrome report `tableVerticalAlignment: true`, with zero Chrome validation/console/page/network errors. Evidence is retained at `/tmp/pptx-table-vertical-alignment-artifacts.1kZjyy`.

## Table-level text direction

```ts
import { PptxDocument, type TableCellTextDirection } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { textDirection: 'vert' });

const uniform: TableCellTextDirection | undefined = table.textDirection; // vert
table.setCellTextDirection(0, 1, 'vert270');
const mixed = table.textDirection; // undefined
table.textDirection = 'horz';
table.textDirection = undefined;
```

`TableModel.textDirection` reads strict consensus from every direct physical-cell `tcPr@vert`. It returns `horz`, `vert`, `vert270`, or `wordArtVert` only when every cell has the same valid direct token; absent, mixed, empty, or unsafe state returns `undefined`, and absence is never synthesized as horizontal. Assigning a legal value atomically replaces every physical cell, including merge continuations; assigning `horz` writes an explicit token, while assigning `undefined` clears every direct token. Exact no-ops preserve slide bytes and the mutation journal. Callers needing mixed detail inspect `rows[].cells[].textDirection`. Unsafe edits raise `ModelParseError` without partial package mutation and preserve unrelated cell/table state.

PptxGenJS 4.0.1 collapses resolved horizontal table creation to attribute absence, so importing that output yields `undefined` at table level rather than `horz`; its three non-horizontal values import exactly. Explicit native `horz` and existing-deck bulk editing are intentional lossless extensions over the same OOXML direct state. Final verification is 5 focused files / 529 focused tests, 79 passed / 1 skipped full test files, 1419 passed / 1 skipped full tests, and performance 1/1 at 1118ms. The actual 62-file tarball SHA-256 is `5f427a8ff77cf64f6dda593ec02fdbe405c44d22481f0357bf05fa39b63ec92d`; installed Node/types/browser/CLI and real Chrome report `tableTextDirection: true`, with zero Chrome validation, console, page, or network errors. Evidence is retained at `/tmp/pptx-table-text-direction-artifacts.BksCOP`.

## Table-level horizontal alignment

```ts
import { PptxDocument, type TextAlignment } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { align: 'center' });

const uniform: TextAlignment | undefined = table.horizontalAlignment; // center
table.setCellHorizontalAlignment(0, 1, 'right');
const mixed = table.horizontalAlignment; // undefined
table.horizontalAlignment = 'justify';
table.horizontalAlignment = 'left'; // explicit direct left
table.horizontalAlignment = undefined;
```

`TableModel.horizontalAlignment` reads strict consensus from every physical cell's unique direct single-paragraph `pPr@algn`. It returns `left`, `center`, `right`, or `justify` only when every cell has that same safe direct value. Absent, mixed, empty, multi-paragraph, or unsafe state returns `undefined`; absent state is never synthesized as `left`. Assigning a legal value atomically replaces every physical cell, including merge continuations. `left` writes an explicit direct token, while `undefined` clears every direct alignment. Exact no-ops preserve slide bytes and the mutation journal. Callers needing mixed detail inspect `rows[].cells[].horizontalAlignment`. Unsafe edits raise `ModelParseError` without partial package mutation and preserve unrelated cell/table state.

PptxGenJS 4.0.1's four legal table-alignment creation values import exactly from their final direct state. Omitted alignment remains `undefined`; a table value plus a differing cell override projects to mixed `undefined`. Native existing-deck bulk editing is an intentional lossless extension. Final verification is 6 focused files / 537 focused tests, 80 passed / 1 skipped full test files, 1427 passed / 1 skipped full tests, and performance 1/1 at 1.34s. The actual 62-file tarball SHA-256 is `03b376861aeb799fa21a99dd105871b8943e29bd4fe51c875a508ff295b9f9c0`; installed Node/types/browser/CLI and real Chrome report `tableHorizontalAlignment: true`. CLI inspection finds exactly four direct final `pPr@algn="r"` tokens and no `tcPr/bodyPr@algn`; Chrome validation, console, page, and network errors are zero. Evidence is retained at `/tmp/pptx-table-horizontal-alignment-artifacts.oe2f5A`.

## Table-level margins

```ts
import { PptxDocument, type TextBoxMargins } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], { margin: [3.6, 7.2, 10.8, 14.4] });

const uniform: TextBoxMargins | undefined = table.margins;
table.setCellMargins(0, 1, { top: 9 });
console.log(table.margins);       // undefined: mixed direct cell state
table.margins = 6;               // writes 6pt on all four sides
table.margins = { top: 2, left: 4 }; // whole-replaces and clears right/bottom
table.margins = undefined;       // clears direct marL/marR/marT/marB
```

`TableModel.margins` reads strict consensus from every physical cell's unique direct `tcPr@marL/marR/marT/marB`. It returns a detached complete or partial `TextBoxMargins` only when every cell has the same non-empty safe side set and values. Absent, mixed-key/value, empty, malformed, repeated, or ambiguous state returns `undefined`; canonical defaults, table styles, and creation input are never synthesized. Callers needing mixed detail inspect `rows[].cells[].margins`.

Assignment accepts point scalar, TRBL, partial named-object, `{}`, and `undefined` input and atomically whole-replaces every physical cell, including merge continuations. Scalar/TRBL writes four sides, partial input clears omitted sides, and `{}`/`undefined` clears all sides. Exact no-ops preserve slide bytes and the mutation journal. Unsafe edits raise `ModelParseError` without partial package mutation and preserve unrelated cell/table state.

PptxGenJS 4.0.1 leaves only final direct cell margin attributes. Omitted output imports as its explicit canonical `{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }`; uniform values project directly and one cell override yields mixed `undefined`. Its legacy runtime treats a first value below 1 as inches and a value at least 1 as points, while native always uses points. Thus PptxGenJS `[0.05, 0.1, 0.15, 0.2]` and native `[3.6, 7.2, 10.8, 14.4]` are compared by final EMU rather than by ambiguous input units.

Final verification is 7 focused files / 547 focused tests, 81 passed / 1 skipped full test files, 1437 passed / 1 skipped full tests, and performance 1/1 at 1.65s. The actual 62-file tarball SHA-256 is `428f47de86cebb89ae19a59b4b5500f3c67c116f63107253e2bf997b04008e37`; installed Node/types/browser/CLI and real Chrome report `tableMargins: true`. CLI inspection finds four direct cells with exact `marL="50800" marR="25400" marT="12700" marB="38100"`; Chrome validation, console, page, and network errors are zero. Evidence is retained at `/tmp/pptx-table-margins-artifacts.gPmz7V`.

## Table-level fill

```ts
import { PptxDocument, type TableCellFill } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
});

const uniform: TableCellFill | undefined = table.fill;
table.setCellFill(0, 1, { kind: 'none' });
const mixed = table.fill; // undefined
const cellDetail = table.rows[0].cells[1].fill; // { kind: 'none' }
table.fill = { kind: 'none' };
table.fill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'D9EAF7' },
  transparency: 0,
};
table.fill = undefined;
```

`TableModel.fill` returns a detached `TableCellFill` only when every physical cell, including merge continuations, has the same safe supported direct `tcPr` fill choice. Absent, mixed, malformed, advanced, or ambiguous state returns `undefined`; it never resolves table styles, effective defaults, or creation input. Mixed detail remains in `rows[].cells[].fill`. Assignment atomically whole-replaces every physical cell. Explicit none writes direct `a:noFill`; solid accepts strict sRGB/theme color with optional finite `0..100` transparency; `undefined` removes the direct choice. Omitted alpha and explicit zero transparency remain distinct. Same-value assignment and an all-absent clear are exact no-ops, while unsafe edits raise `ModelParseError` with zero partial package mutation.

PptxGenJS 4.0.1 legal uniform table solid fill imports from final direct state, omission stays `undefined`, and a differing cell override projects to mixed `undefined`. Its `type: 'none'` collapses to absence and transparency zero collapses to omitted alpha; native preserves direct none, absence, and explicit-zero transparency as distinct direct intent. Existing-deck consensus and bulk editing are native lossless extensions.

Final verification is 8 focused files / 557 focused tests, 1447 passed / 1 skipped full tests, and performance 1/1 at 1.17s. The actual 62-file tarball SHA-256 is `ae7f09233b2ff596c21ec0dab891d5069d810d64921bce9ba96cd08771c6cfdc`; installed Node/types/browser/CLI, `pptx-inspect`, and real Google Chrome report `tableFill: true`. The final four physical cells each contain exactly one direct `a:noFill`; PowerPoint 2010 validation and Chrome validation/console/page/network errors are zero. Evidence is retained at `/private/tmp/pptx-table-fill-artifacts.5MqrXK`.

## Table-level borders

```ts
import { PptxDocument, type TableCellBorders } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const table = document.addSlide().addTable([
  ['North', 'South'],
  ['East', 'West'],
], {
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 1.5,
    style: 'dash',
  },
});

const uniform: TableCellBorders | undefined = table.borders;
table.setCellBorders(0, 1, { kind: 'none' });
const mixed = table.borders; // undefined
const cellDetail = table.rows[0].cells[1].borders; // direct none on four sides
table.borders = {
  top: {
    kind: 'line',
    color: { kind: 'srgb', value: 'D9EAF7' },
    width: 2,
  },
  bottom: { kind: 'none' },
}; // partial whole replacement; right/left are cleared
table.borders = { kind: 'none' };
table.borders = undefined;
```

`TableModel.borders` returns a detached `TableCellBorders` only when every physical cell, including merge continuations, has the same non-empty safe supported direct complete or partial L/R/T/B vector. All-absent, mixed, malformed, advanced, repeated, or ambiguous state returns `undefined`; it never resolves table styles, shared edges, effective borders, defaults, or creation input. Mixed detail remains in `rows[].cells[].borders`. Assignment accepts scalar, exact TRBL, partial named-object, `{}`, or `undefined` input and atomically whole-replaces every physical cell. Scalar writes four sides, TRBL/named input clears omitted sides, and empty/`undefined` clears all four. Exact no-ops preserve bytes and the mutation journal; unsafe edits raise `ModelParseError` with zero partial package mutation.

PptxGenJS 4.0.1 materializes omitted borders as uniform four-side direct none. Legal uniform scalar/TRBL output projects to one consensus, while a legal differing cell override produces mixed `undefined`. Native bulk editing is a lossless existing-deck extension over final direct cell state. Native keeps direct absence, direct none, zero-width line, omitted style, and explicit solid distinct; it does not copy PptxGenJS's empty-object default gray 1pt solid, omitted-type solid materialization, TRBL-zero-to-1pt behavior, or short-tuple padding.

Final verification is 9 focused files / 567 focused tests, 1457 passed / 1 skipped full tests, and performance 1/1 at 1.20s. The actual 62-file tarball SHA-256 is `47d9666c1dac8454524a87f7ca1898af0442c6faa7816e39cec34ff42dbf0d48`; installed Node/types/browser/CLI, `pptx-inspect`, and real Google Chrome report `tableBorders: true` / `tableBordersInspect: true`. Each final physical cell has exactly one direct `lnL/lnR/lnT/lnB` no-fill set; PowerPoint 2010 validation is 0 errors / 0 warnings, and Chrome validation/console/page/network errors are zero. Design, plan, correction, core, and package-proof commits are `8d6dfae`, `d4bf9c9`, `faca4ba`, `9195e57`, and `f8a9d0d`; evidence is retained at `/tmp/pptx-table-borders-artifacts.vyi1yo`.

## Table-cell hyperlink creation, snapshots, and editing

```ts
import { PptxDocument, TableModel } from '@jiayunxie/pptx';

const document = PptxDocument.create();
const source = document.addSlide();
const target = document.addSlide();
const table = source.addTable([[
  { text: 'Website', options: { hyperlink: { url: 'https://example.com' } } },
  { text: 'Details', options: { hyperlink: { slide: 2, tooltip: '' } } },
]], { name: 'Cell links' });

console.log(table.rows[0].cells[0].hyperlink); // { url: 'https://example.com' }
console.log(table.rows[0].cells[1].hyperlink); // { slide: 2, tooltip: '' }
document.moveSlide(document.slides.indexOf(target), 0);
console.log(table.rows[0].cells[1].hyperlink); // { slide: 1, tooltip: '' }
table.setCellText(0, 1, 'Open details'); // preserves the cell hyperlink
table.setCellHyperlink(0, 0, {
  url: 'https://example.com/docs',
  tooltip: 'Read the docs',
});
table.setCellHyperlink(0, 1, undefined); // clears click, preserving run style

const reopened = await PptxDocument.open(await document.write());
const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
const reopenedTable = reopenedSource.shapes.find(
  ({ name }) => name === 'Cell links',
) as TableModel;
console.log(reopenedTable.rows[0].cells[0].hyperlink);
// { url: 'https://example.com/docs', tooltip: 'Read the docs' }
console.log(reopenedTable.rows[0].cells[1].hyperlink); // undefined
```

`AddTableCellOptions.hyperlink?: Hyperlink` and `TableModel.setCellHyperlink(rowIndex, columnIndex, value)` accept a URL or a one-based current-presentation slide target. Omitted tooltip and explicit `''` remain distinct; `undefined` clears the selected zero-based physical cell's direct run click. Readonly `TableCell.hyperlink?: Hyperlink` immediately returns a detached frozen snapshot of the supported direct state. The link belongs to the plain cell's only direct run, survives `setCellText()` and supported property edits, and follows target-slide identity while reorder changes the reported ordinal. Duplicate, delete, rollback, all six formats, and write/reopen use the existing relationship lifecycle.

Native-created linked cells own independent relationships, including cells with equal URLs. Imported shared IDs use clone-on-write for target changes. Equal values are exact part/relationship/journal no-ops; tooltip-only edits retain the ID, unique target changes update in place, and clear/replace garbage-collect only the last reference. Adding a click supplies direct single underline when absent; clearing removes only the click and preserves underline plus every other run property. No `AddTableOptions.hyperlink` default is provided. The scalar reader/editor requires exactly one namespace-correct direct text body, paragraph, run, run-properties node, and text node; rich, multi-run, and multi-paragraph default/local links are read through `TableCell.richText` and whole-replaced through `setCellRichText()`.

Legal PptxGenJS 4.0.1 URL/slide output imports and edits, including extra `invalidUrl=""`, `action=""`, and `history="1"` run-click attributes; PptxGenJS has no existing-deck table-cell hyperlink editor. It materializes omitted tooltip as empty and writes `_rId` back into caller hyperlink objects. Native preserves omitted versus empty state, never mutates input, and does not copy loose coercion or dangling-relationship behavior.

Final verification is 84 passed / 1 skipped full test files and 1476 passed / 1 skipped tests, with performance 1/1 at 1.63s. The actual 62-file tarball SHA-256 is `2d06b955b48a25fc6f1e06accf2bd059045a7b3db04a6f3640c5bdca987ea816`; installed Node/types/browser/CLI, `pptx-inspect`, and real Google Chrome report `tableCellHyperlinkEditing: true`. The final evidence deck has 22 parts, 23 relationships, and 3 slides; its first table slide has 6 cells, 2 clicks, 2 matching relationships, and 6 preserved underlines. PowerPoint 2010 validation has 0 errors and only 2 expected `OPC_EXTERNAL_RELATIONSHIP` warnings, while Chrome validation/console/page/network errors are zero. Design, plan, rollback, core, and package-proof commits are `4fe0c43`, `0e566bd`, `dca33ba`, `99a6d3b`, and `93b6b09`; evidence is retained at `/tmp/pptx-table-cell-hyperlink-editing-UphgYg`.

## Rich and multi-paragraph table-cell text

```ts
import { PptxDocument, type RichTextParagraph } from '@pptx/sdk';

const document = PptxDocument.create();
const slide = document.addSlide();
const table = slide.addTable([[
  {
    text: [
      {
        align: 'center',
        runs: [{ text: 'Title', style: { bold: true, fontSize: 18 } }],
      },
      {
        runs: [
          { text: 'Default link' },
          { text: 'OpenAI', style: { hyperlink: { url: 'https://openai.com' } } },
          { text: 'soft break', softBreakBefore: true, style: { hyperlink: false } },
        ],
      },
    ],
    options: { hyperlink: { url: 'https://example.com' } },
  },
  'first\r\n\rthird',
]]);

const cell = table.rows[0]!.cells[0]!;
const plain: string = cell.text;
const rich: readonly RichTextParagraph[] = cell.richText;
table.setCellRichText(0, 0, [{ runs: [{ text: 'Replaced' }] }]);
```

`AddTableCell.text` accepts string or `readonly RichTextParagraph[]`; a rich array must be wrapped in a cell object. CRLF/CR strings normalize to LF and split into direct paragraphs while preserving empty and trailing lines. Structured input preserves multiple paragraphs and all supported paragraph/run styles. `softBreakBefore` remains an in-paragraph break, `breakLine` is normalized into a new paragraph, and `TableCell.text` projects either boundary to `\n`. `TableCell.richText` is a detached readonly snapshot.

Cell `options.hyperlink` defaults runs that do not provide a local override. Explicit run links own independent relationships and `hyperlink: false` suppresses the default. Rich links are read and whole-replaced through the rich snapshot; the scalar `TableCell.hyperlink` / `setCellHyperlink()` API remains limited to an exact plain single-run cell. `setCellRichText()` uses zero-based physical row/cell indexes, preserves body/cell/table structure and neighbors, and performs an exact no-op for an equal value. Relationship reuse, clone-on-write, final-reference garbage collection, slide-target identity, duplicate/delete, rollback, all six formats, and write/reopen are covered.

`setCellText()` is safe only for one exact direct paragraph and run, whose style/link template it preserves. Rich, multi-run, multi-paragraph, field, or break cells require `setCellRichText()` and reject scalar replacement before mutation. Inputs and snapshots detach from caller state. Legal PptxGenJS 4.0.1 CR/LF, `breakLine`, paragraph style, run style, cell-default link, and run-local link output imports and remains editable without copying caller mutation, duplicate `pPr`, tooltip collapse, loose coercion, or dangling-link defects.

Final gates are 85 passed / 1 skipped test files and 1487 passed / 1 skipped tests in 143.22s, plus the 1000-part performance test at 1648ms. TypeScript, Node/browser bundles, declarations, an actual 62-file tarball (SHA-256 `7de2354ac691ad09b58e0e103fd07ff1428caa799b548d2a65d9d19a4e0fd79f`), installed Node/NodeNext/browser/CLI/Inspector, and PptxGenJS import/edit checks pass. PowerPoint 2010 reports zero errors and only expected external-link warnings. Chrome 150.0.7871.188 reports all four rich-cell capability flags true and zero validation/console/page/network errors. Evidence is retained at `/tmp/pptx-table-cell-rich-text-DA3x3Z`.

## Table and cell text-style defaults

```ts
import { PptxDocument } from '@pptx/sdk';

const document = PptxDocument.create();
const slide = document.addSlide();
const table = slide.addTable([[
  'inherits table defaults',
  {
    text: [{
      spacing: { after: 10 },
      runs: [
        { text: 'inherits cell defaults' },
        { text: ' · local', style: { fontSize: 12, bold: false } },
        {
          text: ' · link',
          style: { hyperlink: { url: 'https://example.com' } },
        },
      ],
    }],
    options: {
      fontFamily: 'Courier New',
      bold: false,
      spacing: { before: 3 },
    },
  },
]], {
  fontFamily: 'Aptos',
  fontSize: 18.25,
  bold: true,
  color: { kind: 'scheme', value: 'accent1' },
  spacing: {
    before: 6,
    after: 8,
    line: { kind: 'multiple', factor: 1.5 },
  },
});
```

Both `AddTableOptions` and `AddTableCellOptions` expose `fontFamily?: string`, `fontSize?: number`, `bold?: boolean`, `color?: RichTextColor`, and `spacing?: ParagraphSpacing`. Resolution order is table → cell → explicit paragraph/run. Scalar run fields override independently; spacing overlays by `before`, `after`, and `line`. Explicit run/cell `bold: false`, paragraph `spacing: false`, and paragraph `line: false` block inherited values. Native rejects the PptxGenJS aliases `fontFace`, `paraSpaceBefore`, `paraSpaceAfter`, `lineSpacing`, and `lineSpacingMultiple`.

Only resolved direct OOXML is stored. `TableCell.richText` immediately reads it, `setCellText()` preserves a safe plain run's materialized style, and `setCellRichText()` does not reapply creation defaults. Font family and size are also written to empty-paragraph `endParaRPr`. Cell-default hyperlinks retain outer color; local run hyperlinks without explicit color skip the outer color. PptxGenJS 4.0.1 legal font/size/bold/color, cell spacing, rich override, empty paragraph, and hyperlink output imports and remains editable. Native additionally propagates table-level spacing, preserves explicit false values, and never mutates caller objects. Table-cell merges, physical row/column CRUD, auto-page/repeated headers, and automatic content measurement/layout recomputation are completed below. The 99.3% figure attached to this earlier specialty is retained as a historical checkpoint.

Final gates are 85 passed / 1 skipped test files and 1497 passed / 1 skipped tests in 167.50s, plus the 1000-part performance test at 1565ms. TypeScript, Node/browser bundles, declarations, and an actual 62-file tarball (SHA-256 `79ed789e6d4f218cc5c838af9e5965e96bd7e35f132d2a630a85ac5dd39ed222`) pass installed Node, NodeNext types, browser conditional-export, CLI, and Inspector probes. The retained 18-part / 15-relationship deck has one slide, one table, and three cells; PowerPoint 2010 reports zero errors and warnings. Chrome 150.0.7871.188 reports every table-text-default capability true and zero validation/console/page/network errors. Evidence is retained at `/tmp/pptx-table-text-defaults-proof.ViSdTX`.

## Table-cell colspan, rowspan, snapshots, and editing

```ts
const table = slide.addTable([
  [{ text: 'Summary', options: { colspan: 2, rowspan: 2 } }, 'Total'],
  ['42'],
]);

console.log(table.mergeRegions);
table.unmergeCell(1, 1); // any physical member resolves the whole region
table.mergeCells(0, 0, 2, 2);
```

Creation accepts logical rows. The sum of first-row `colspan` values defines the physical column count. Later rows place each logical cell at the leftmost free physical column after skipping active `rowspan` coverage; a fully covered later row may therefore be `[]`. Both span fields accept positive safe integers only, with omitted or `1` meaning an unmerged dimension. The complete logical-to-physical layout, text/style defaults, and hyperlink targets are validated before observable package mutation. Holes, overlaps, non-rectangular coverage, bottom/right overflow, or more than 1,000,000 expanded physical cells are rejected.

`TableMergeRegion` contains zero-based physical `rowIndex`, `columnIndex`, `rowspan`, and `colspan`. `TableModel.mergeRegions` returns `[]` for a recognized unmerged table, a row-major detached deeply frozen list for recognized merges, and `undefined` when direct table structure or merge topology is not safely recognizable. A recognized region gives every member `TableCell.merge` the same anchor and span values; `isAnchor` is true only on the top-left physical cell. Existing scalar/rich/style editors retain their physical-coordinate contract and may explicitly edit preserved continuation state.

`mergeCells(rowIndex, columnIndex, rowspan, colspan)` requires at least two cells, accepts only an in-bounds rectangle, and rejects any existing-region intersection unless it is the exact same region. `unmergeCell(rowIndex, columnIndex)` resolves either an anchor or continuation to the whole region. An exact repeated merge and an unmerged-cell unmerge preserve bytes, relationships, ZIP state, model identity, and the mutation journal.

Both editors modify only the direct physical-cell `rowSpan`, `gridSpan`, `vMerge`, and `hMerge` attributes. Continuation text, style, relationships, opaque children, and unknown attributes remain byte-preserved and become visible after unmerge. A malformed or ambiguous topology leaves `mergeRegions` undefined and rejects semantic merge/unmerge without changing the package.

Legal PptxGenJS 4.0.1 horizontal, vertical, rectangular, and offset spans reach the same final semantic state through native creation and remain readable/editable after import. Native rejects lopsided non-span rows, negative/fractional spans, and out-of-bounds rowspans that PptxGenJS can serialize into malformed or inconsistent tables.

Final focused verification is 5/5 test files and 594/594 tests in 28.11s. The full suite is 86 passed / 1 skipped test files and 1512 passed / 1 skipped tests in 73.26s; the independent 1000-part performance gate is 1/1 at 709ms. TypeScript project references, Node/browser bundles, and declarations pass. Two 59-file dist manifests match exactly, and two 62-file actual tarballs are byte-identical with SHA-256 `0c85afa9bed6a04faa5d3dab6934a3974cea731091dc673ab2ff6e92cb83343d`. Installed Node, NodeNext types, browser conditional export, CLI, and Inspector report `tableCellMerges: true` / `tableCellMergesInspect: true`.

Google Chrome 150.0.7871.188 reports create/read/frozen snapshot/unmerge/edit/remerge/reopen true with zero validation/console/page/network errors. The retained browser deck has 18 parts / 15 relationships, 1 slide / 1 table, a 2×3 physical matrix, and one 2×2 merge region. It contains all four anchor/continuation token forms, has only the valid slide-layout relationship, and validates under the PowerPoint 2010 profile at 0 errors / 0 warnings. Recognition, creation, snapshot/editor, SDK/adapter, documentation, and package-proof commits are `688f9f6`, `3d93f07`, `db01937`, `b2f6846`, `5832399`, `7073eae`, and `f174519`; evidence is retained at `/tmp/pptx-table-cell-merges-artifacts.B7ZhGQ`. Physical row/column CRUD, auto-page/repeated headers, and automatic content measurement/layout recomputation are completed in the following sections. The 99.3% figure here is a historical checkpoint.

## Table row and column insertion and deletion

```ts
const table = slide.addTable([
  ['A0', 'A1', 'A2'],
  ['B0', 'B1', 'B2'],
  ['C0', 'C1', 'C2'],
], {
  columnWidths: [inches(1), inches(2), inches(3)],
  rowHeights: [inches(0.5), inches(1), inches(1.5)],
});

table.mergeCells(0, 0, 2, 2);
table.insertRows(1, { count: 2, rowHeights: [inches(0.25), inches(0.5)] });
table.insertColumns(1, { columnWidths: inches(0.75) });
table.setCellText(1, 1, 'Editable hidden continuation');
table.deleteRows(4);
table.deleteColumns(3);
```

The four structural methods use zero-based physical coordinates. Insert accepts an index from zero through the current row or column count, including append. Delete must start at an existing item, and its range may not remove every row or every column. Omitted `count` is one; supplied counts are positive safe integers. The post-edit physical-cell count may not exceed 1,000,000.

`InsertTableRowsOptions.rowHeights` accepts a non-negative safe-EMU scalar or a dense readonly array whose length equals `count`; zero is an automatic row. `InsertTableColumnsOptions.columnWidths` accepts positive safe EMU values with the same scalar/exact-array rule. When omitted, an insertion before an existing item copies that item's direct size, and append copies the final item. Column insertion/deletion recomputes transform width from the exact grid sum. Row insertion/deletion recomputes transform height only when all resulting direct heights are positive; a table containing any automatic row retains its existing height.

An insertion at a merge anchor coordinate occurs before the region, while one strictly inside its bounds expands that region. Deletion contracts overlapping regions, promotes the top-left surviving physical cell when the original anchor is removed, and clears merge tokens when only one cell remains. Newly inserted cells use a canonical empty `txBody` that is immediately compatible with the existing text, rich-text, hyperlink, alignment, margin, border, fill, direction, and fit editors. Structural insertion deliberately does not accept content or infer adjacent style.

The implementation splices direct grid, row, and cell source spans. Surviving content, style, hidden continuation state, relationship IDs, unknown attributes/children, and lexical XML remain intact except for owned numeric or merge tokens whose semantics changed. Deleted relationships are collected only after their last slide-wide XML reference disappears. XML replacement and relationship collection share one OPC transaction, including outer rollback and failure recovery.

PptxGenJS 4.0.1 exposes table construction, sizes, and auto-page helpers but no existing-deck structural editor. Its legal public plain/rich/linked/merged/sized output can be imported and edited with these native methods. Malformed imported topology remains preservation-only. Auto-page, repeated headers, and automatic content measurement/layout recomputation are documented below. Logical content insertion and retained creation defaults remain outside this structural-edit API; browser DOM import is covered by the separate `tableToSlides()` API.

Final focused verification is 5 files / 611 tests in 29.96s. The full suite is 87 passed / 1 skipped test files and 1535 passed / 1 skipped tests in 64.27s; the independent 1000-part performance gate is 1/1 with the core test at 1204ms, test file at 1207ms, and total at 2.52s. TypeScript project references, the root build, Node/browser bundles, and declarations pass. Two clean 59-file dist manifests match exactly with manifest SHA-256 `51d0c19da69fbd81682933d4a5418ff58ef2a805b4164d624f150d1674924e41`; two 62-file, 660,178-byte actual tarballs are byte-identical with SHA-256 `17d43a887a9871fd4910bcf33415d985b4d8f1968b4020670a64166c148aeaa4`. Installed Node, NodeNext declarations, browser conditional export, CLI, and Inspector report `tableStructureEditing: true`.

Google Chrome 150.0.7871.188 reports every lifecycle stage true with zero validation/console/page/network errors. Node and browser evidence decks both contain 18 parts / 16 relationships, 1 slide / 1 table, and a 4×4 physical matrix. Their column widths `[914400, 457200, 1828800, 2743200]` sum to transform width 5943600; row heights `[457200, 228600, 914400, 1371600]` sum to transform height 2971800. The final 3×3 merge has 1 anchor, 2 top, 2 left, and 4 interior continuations; hidden inserted text and a styled survivor remain. Two clicks share `rId2`, one external relationship survives, and no hyperlink is orphaned. PowerPoint 2010 validation is 0 errors / 1 expected `OPC_EXTERNAL_RELATIONSHIP` portability warning. Commits are `ee68731`, `d70c2af`, `099f345`, `89f9b1b`, `2250826`, and `1ab602f`; evidence is retained at `/tmp/pptx-table-structure-editing-proof.S1rVAZ`. The CRUD workstream is 8/8 complete; explicit-row-height auto-page/repeated headers are completed below.

## Table auto-page and repeated headers

```ts
import { inches, PptxDocument, type SlideModel } from '@jiayunxie/pptx';

const document = PptxDocument.create();
document.addSlide();
const source = document.addSlide();
const following = document.addSlide();

source.addTable([
  [{ text: 'Region', options: { autoPageCharWeight: -0.25 } }, 'Revenue'],
  ['Unit', 'USD'],
  [{
    text: [{ runs: [
      { text: 'North '.repeat(700), style: { bold: true } },
      { text: '$120', softBreakBefore: true, style: { italic: true } },
    ] }],
    options: { colspan: 2, margin: 0 },
  }],
  ['South', '$95'],
  ['East', '$110'],
], {
  autoPage: true,
  autoPageCharWeight: 0,
  autoPageLineWeight: 0,
  autoPageRepeatHeader: true,
  autoPageHeaderRows: 2,
  autoPageSlideStartY: inches(0.75),
  slideMargin: [inches(0.5), inches(0.4), inches(0.5), inches(0.4)],
  y: inches(5),
  columnWidths: [inches(3), inches(2)],
  // Omitted rowHeights measure every row from its content.
});

const generated: readonly SlideModel[] = source.newAutoPagedSlides;
console.log(generated.length); // number of continuation slides
console.log(document.slides.indexOf(following)); // after every generated continuation
```

`AddTableOptions` exposes `autoPage`, `autoPageCharWeight`, `autoPageLineWeight`, `autoPageRepeatHeader`, `autoPageHeaderRows`, `autoPageSlideStartY`, and `slideMargin`; `AddTableCellOptions` exposes the two weight overrides. Geometry is direct EMU. `slideMargin` is a non-negative scalar broadcast or an exact dense `[top, right, bottom, left]` tuple. Explicit margins take precedence over the source slide's current runtime named-layout margin; after reopen that transient layout margin is absent, so the planner uses canonical 0.5-inch margins until it is reapplied.

Omitting both `height` and `rowHeights` measures every physical row. A zero entry measures that row, while a positive entry is its minimum; supplying only `height` still distributes fixed positive row heights. `rowHeights: [0, inches(0.5), 0]` is a mixed automatic/minimum vector. An all-positive vector without any table/cell measurement weight retains fixed structural pagination. A measured table's serialized rows are ordinary positive-height rows whose transform height is the exact per-page sum. Supplying `height` together with an automatic zero row is invalid.

The default measurement font size is 12pt. Cluster advance is based on `fontSize × units / (2.3 + autoPageCharWeight)` plus run character spacing, and natural line height uses `fontSize × (1.67 + autoPageLineWeight)`; the implementation rounds the final values to safe EMU. Both weights are strict finite values in `[-1, 1]`, cell overrides win, and omitted and explicit zero are numerically equivalent. Combining marks, variation selectors, emoji skin tones, and ZWJ sequences remain one cluster. Whitespace, ASCII punctuation, Latin/digits, and wide clusters use deterministic unit widths. Run font size/character spacing, soft breaks, cell margins, paragraph margins/indents/bullets/tabs, and before/after/exact/multiple spacing contribute. Colspan uses the exact sum of spanned column widths; rowspan content adds a minimum constraint distributed across the full merge block.

The source capacity begins at the table `y`; continuation capacity begins at `autoPageSlideStartY`, layout top margin, or canonical top margin, and ends before the chosen bottom margin. `autoPageRepeatHeader` defaults to false. When enabled, `autoPageHeaderRows` defaults to one and must select one or more physical rows within the table.

Pagination never splits a rowspan block. A merge that crosses the header/body boundary is invalid. An oversized measured row without rowspan is fragmented only at complete measured line bands, preserving rich paragraph/run styles, soft breaks, paragraph edge spacing, hyperlinks, and page-local relationship ownership. A fixed minimum that causes overflow, a merge block, or one line band larger than continuation body capacity raises a strict error. The planner emits ordinary page-local table definitions with exact row-height sums and no serialized auto-page metadata. Repeated links receive relationships owned by each page; internal slide targets are resolved by stable part identity before slide insertion.

For `placeholder`, the selected table owner supplies source X/Y/width and an owner-bottom limit. Continuation pages retain owner X/width, use `autoPageSlideStartY` for Y, and share the same bottom edge. Each page height is the exact sum of its measured rows and is never stretched to owner height. Placeholder identity, layout, section membership, slide-number cache, and stable internal-link targets remain synchronized.

```ts
source.addTable([['Header'], ['Body']], {
  autoPage: true,
  placeholder: 'data_table',
  autoPageSlideStartY: inches(1.25),
  rowHeights: [0, inches(0.4)],
});
```

Generated slides are new canonical blank slides inserted contiguously after the source. They use the same direct layout relationship, inherit exact source section membership, materialize layout placeholders/slide numbers like ordinary slides, and never reuse a pre-existing following slide. The source table, generated slide parts, relationships, presentation order, sections, model caches, and runtime result share one outer transaction and roll back together.

`SlideModel.newAutoPagedSlides` is getter-only and returns a frozen readonly snapshot of the continuation slides produced by the latest successful `addTable()` call. It excludes the source. A successful ordinary table or no-overflow auto-page resets the result to empty; a failed call preserves the previous successful result. Deleted generated slides are filtered, duplicate sources start empty, and write/reopen clears every result because this runtime metadata is not serialized.

Native rejects deprecated `addHeaderToEach` / `newSlideStartY`, coercible controls, malformed margins, and unsupported fragmentation boundaries, and does not copy PptxGenJS 4.0.1 caller mutation, weight clamping, or following-slide reuse. This paragraph records the historical 99.7% measurement checkpoint; `tableToSlides` is now supported below. Public capability coverage is 100%, while final full-parity certification still waits for the peer/client audit.

The table content measurement/layout recomputation workstream is 11/11 complete. The focused gate covers 8 files (7 passed / 1 skipped) with 107 passed / 644 skipped tests in 8.59s. The full gate covers 90 files (89 passed / 1 skipped) with 1654 passed / 1 skipped tests in 40.16s. The independent 1000-part performance gate is 1/1 with a 604ms core test and 1.68s Vitest duration. TypeScript typecheck, root build, and the `@jiayunxie/pptx` package build complete in 2.15s, 1.66s, and 6.51s. Two 59-file dist manifests are identical with manifest SHA-256 `c2e9acc7f14aebb32e425fc0abcc8f62677e5268ee883f941ccda74d74c1d5ab`; two 62-file actual tarballs are byte-identical with SHA-256 `1ce84a18208daa0b045ef45dfa4b79f4728daa5a3111c960caf823a1b334b4ac`. Installed Node and browser conditional-export `tableContentMeasurement` states are all true, while the installed CLI/Inspector reports `tableContentMeasurementInspect: true`. Google Chrome 150.0.7871.188 has 0/0/0 console/page/network errors.

Both Node and browser evidence decks contain 42 parts / 57 relationships and 12 slides / 11 tables, including nine paginated tables. Fixed rows are `[274320, 274320, 274320]`; Node automatic rows are `[549738, 183246]`; browser automatic rows are `[732984, 183246]`; and the minimum row is `320040`. All nine section pages, repeated headers, page-local transform/row-height sums, layouts, and placeholder identities pass. Eighteen clicks own eighteen exact page-local relationships with zero orphan links. Node/browser deck SHA-256 values are `b37d5204aab7cdb98d42e6b0e0eda7af782b660f15901440189fc434fefa214a` and `d734893e309809d5fa4cce1751db302fec9bf1b7c307c4b48fa8f705cfe8dfab`; both validate under the PowerPoint 2010 profile at 0 errors / 15 expected external-link warnings. The commit chain is `4482555`, `7a262ae`, `95b98ce`, `78cb279`, `6633696`, `e64e232`, `d77f54b`, `2f7595a`, `e765011`, and `6c63e0b`; complete evidence is retained at `/tmp/pptx-table-content-measurement-artifacts.zxJbXX`. The workstream was 100% complete and overall parity was approximately 99.7% at that historical checkpoint.

## Embedded raster images

```ts
import {
  calculateRasterImageSizing,
  degrees,
  inches,
  inspectRasterImage,
  type AddImageSourceOptions,
  type RasterImageSizing,
} from '@pptx/sdk';

declare const updatedPngBytes: Uint8Array;
const sizing: RasterImageSizing = {
  type: 'cover',
  width: inches(5),
  height: inches(3),
};
const options: AddImageSourceOptions = {
  name: 'Quarterly chart',
  altText: 'Revenue by quarter',
  x: inches(1),
  y: inches(1.5),
  sizing,
  rotation: degrees(10),
};
const image = await created.addImage(0, 'chart.png', options);
const info = inspectRasterImage(updatedPngBytes);
console.log(info); // { contentType: 'image/png', width, height }
const crop = calculateRasterImageSizing(info, {
  type: 'crop',
  width: inches(3),
  height: inches(2),
  source: { x: 400, y: 225, width: 800, height: 450 },
});
console.log(crop); // target width/height plus four-edge sourceRectangle
image.setTransform({ x: inches(1.5) });
image.sourceRectangle = { left: 10, top: -5, right: 5, bottom: 0 };
image.sourceRectangle = undefined;
image.replaceData(updatedPngBytes, 'image/png');
```

`PptxDocument.addImage(slideIndex, source, options?)` returns `Promise<ImageModel>`. `RasterImageSource` accepts Node file paths, HTTP/HTTPS URLs, browser-relative URLs, strict base64 data URIs, `Uint8Array`, `ArrayBuffer`, `Blob`/`File`, Web `ReadableStream`, and async byte iterables. PNG/JPEG/GIF content type and raw pixel dimensions are detected exclusively from byte signatures and format structure. `AddImageSourceOptions.contentType` is an optional canonical MIME assertion, and `signal` aborts file, Fetch, Blob, or stream loading. A source assertion mismatch, unsupported/truncated bytes, failed load, invalid slide index, or invalid options rejects before package mutation. File extensions and transport metadata are intentionally ignored.

`inspectRasterImage(bytes)` exposes the same canonical `{ contentType, width, height }` inspection without mutation. Pure `calculateRasterImageSizing(info, sizing)` accepts a positive-safe-integer EMU target frame. `contain` preserves the complete source and may produce negative source-rectangle edges; `cover` fills the frame with symmetric crop; `crop` maps an explicit finite source-pixel region. Results and nested source rectangles are detached and frozen, with each edge quantized once to 0.001%. Omitted transforms remain x/y 0, width/height one inch, rotation 0, and both flips false; intrinsic pixels do not automatically determine slide size.

When `AddImageSourceOptions.sizing` is present, top-level `width` or `height` is a type and runtime error. High-level direct `sourceRectangle` is not accepted. The complete options/sizing tree is normalized and detached before source I/O, and the sizing calculation completes before package mutation; invalid sizing, source failure, MIME mismatch, extreme contain, or out-of-bounds crop leaves the package unchanged. Omitted names use the current slide image count as `Image N`; omitted alt text is `preencoded.png`, while direct empty name and alt text remain empty.

Each call allocates one unique `/ppt/media/imageN.{png|jpeg|gif}` part with the exact supplied content type and bytes, one internal image relationship from the slide, and canonical `p:pic` XML with rectangular geometry, `noChangeAspect`, and stretch/fillRect. Normalization happens before mutation, while part/relationship/XML writes run in one package transaction; any failure restores parts, content types, relationships, ZIP state, slide XML, shape IDs, object identity, and mutation journal. Unsupported or unknown options, empty/wrong-type bytes, noncanonical MIME values, unsafe transforms, invalid XML strings, and zero/nonpositive extents are rejected with no package change.

The synchronous lower-level `SlideModel.addImage(bytes, options)` requires a non-empty copied `Uint8Array` plus canonical `RasterImageContentType`, and may accept direct `sourceRectangle`. `ImageSourceRectangle.left/top/right/bottom` are percentages where `1` means 1%; every edge must be less than 100, horizontal and vertical pair sums must each leave positive source area, and negative values support contain. The direct value is quantized to 0.001% DrawingML integer percentages. The returned live `ImageModel.sourceRectangle` getter provides a detached frozen snapshot; assignment whole-replaces or repairs the single direct `a:srcRect`, same-value assignment is an exact no-op, and `undefined` clears it while preserving the image payload, relationship, transform, effects, and neighboring fill state.

The returned `ImageModel` also exposes its inherited name/transform lifecycle plus `sourcePartUri`, `externalUrl`, and `replaceData()`. An exclusive embedded target is replaced in place; a shared target or relationship is cloned and only that image is retargeted. `duplicateSlide()` therefore starts with shared image parts and replacement isolates the edited duplicate. Creation and lifecycle are covered across pptx/pptm/ppsx/ppsm/potx/potm, Node and browser bundles, declarations, CLI smoke, and write/reopen.

Valid PptxGenJS 4.0.1 public path/data PNG/JPEG/GIF output reaches the same loader state. Six public contain/cover/equal-ratio/crop cases also match native final transform and direct `srcRect` integer percentages exactly. Native stays strict instead of reproducing PptxGenJS path/data precedence, console-only invalid handling, truthy sizing fallback, ambiguous outer/nested dimensions, or noncanonical `image/jpg` output. The actual tarball passes Node/browser/declaration/CLI smoke, and two clean builds produce identical SHA-256 manifests for all 38 dist files.

The sizing gallery has four slides, 40 shapes, and 12 images. Source and LibreOffice round-trip packages reopen strictly, validate against PowerPoint 2010 with 0 errors and 0 warnings, report zero overflow, and were visually inspected page by page. LibreOffice retains 12/12 payload hashes, content types, names, non-empty alt text, order, and internal relationships; it deduplicates the 12 repeated payload targets to three and rewrites all 12 picture blocks. The maximum transform quantization is 360 EMU, the maximum source-rectangle quantization is 0.007%, and two flips normalize equivalently to an added 180-degree rotation. Source 180-DPI pages are 2400×1350. The round-trip page width is client-normalized, producing 2401×1350 directly at 180 DPI and a proportional 2400×1350 inspection raster.

## Embedded SVG images

```ts
import {
  calculateImageSizing,
  inches,
  inspectSvgImage,
} from '@pptx/sdk';

declare const svgBytes: Uint8Array;
declare const fallbackPngBytes: Uint8Array;
declare const replacementSvgBytes: Uint8Array;
declare const replacementPngBytes: Uint8Array;

const svgInfo = inspectSvgImage(svgBytes);
const placement = calculateImageSizing(svgInfo, {
  type: 'cover',
  width: inches(6),
  height: inches(4),
});
const svgImage = await created.addImage(0, svgBytes, {
  contentType: 'image/svg+xml',
  fallback: fallbackPngBytes,
  name: 'Architecture',
  altText: 'System architecture diagram',
  sizing: { type: 'contain', width: inches(6), height: inches(4) },
});
console.log(placement.sourceRectangle);
console.log(svgImage.isSvg, svgImage.fallbackPartUri, svgImage.svgPartUri);

const lowLevelSvg = created.slides[0]!.addSvgImage(svgBytes, fallbackPngBytes, {
  x: inches(6.5),
  width: inches(4),
  height: inches(3),
});
lowLevelSvg.replaceSvgData(replacementSvgBytes, replacementPngBytes);
```

`PptxDocument.addImage()` detects strict SVG XML from the same Node path, HTTP/HTTPS URL, browser-relative URL, canonical base64 data URI, `Uint8Array`, `ArrayBuffer`, Blob/File, Web Stream, and async-iterable source union used by raster images. `ImageContentType` adds canonical `image/svg+xml`; `AddImageSourceOptions.contentType` remains an optional assertion. `inspectImage()` dispatches raster signatures or SVG XML, while `inspectSvgImage()` returns frozen `{ contentType: 'image/svg+xml', width, height }` intrinsic information. Generic `calculateImageSizing()` supports contain, cover, and source-coordinate crop for both raster and SVG; raster-named inspector/sizing aliases remain available.

`AddImageSourceOptions.fallback` is valid only for SVG and must resolve to a signature-valid PNG. Fallback resolution order is explicit PNG, browser Canvas rasterization (maximum side 8,192 and maximum 16,777,216 pixels), then a detached built-in transparent PNG. Every high-level generated `.png` part contains a real PNG signature. Callers that require full appearance in fallback-only clients should always supply a high-fidelity PNG; the API does not execute SVG scripts, fetch external SVG resources, or guarantee arbitrary rasterization fidelity.

The synchronous `SlideModel.addSvgImage(svgBytes, fallbackPngBytes, options?)` requires and copies non-empty `Uint8Array` payloads, leaving SVG/PNG byte correctness to the low-level caller, then commits a canonical two-part picture in one package transaction. The high-level API performs strict SVG and PNG inspection before calling it. A safe pair exposes `ImageModel.isSvg === true`, `sourcePartUri === fallbackPartUri`, and a distinct `svgPartUri`. The reader resolves the Office SVG extension by namespace and relationship semantics rather than a fixed prefix and accepts LibreOffice's relationship-identified `image/svg` normalization. External SVG relationships and ambiguous/malformed pair state are not exposed as editable SVG state.

`ImageModel.replaceData()` rejects any Office SVG extension candidate. `replaceSvgData(svgBytes, fallbackPngBytes)` validates and copies both non-empty byte inputs before mutation and updates an exclusive canonical pair in place; shared or noncanonical targets are cloned and only the selected picture is retargeted. Slide duplication initially shares both targets. Pair replacement, rollback, six-format write/reopen, arbitrary prefix import, PptxGenJS import, and LibreOffice-normalized shared-target clone-on-write are covered. A missing/unsafe pair or either replacement failure leaves parts, relationships, content types, slide XML, model identity, ZIP state, and mutation journal unchanged.

PptxGenJS 4.0.1 conformance covers three public cases: data-contain, path-cover, and data-crop. Native matches the final picture order, transform, direct `srcRect`, extension URI/namespace, relationship roles, SVG payload, metadata, and canonical content types. The intentional divergence is PptxGenJS's path SVG fallback containing SVG bytes in a `.png` part; native always emits valid PNG fallback bytes.

The packed Node/browser/type/CLI smoke covers explicit/default fallbacks, sizing, duplicate sharing, paired clone-on-write, replacement/reopen, Canvas fallback, two internal relationships, and PowerPoint 2010 validation. Two clean builds have identical SHA-256 manifests for all 38 dist files. The five-slide gallery contains 13 shapes, eight SVG pictures, seven SVG parts, seven PNG fallbacks, and 16 image relationships. Source and LibreOffice round-trip packages strictly reopen and validate with 0 errors and 0 warnings. LibreOffice preserves shape order, names, alt text, SVG hashes, relationship roles, and 7+7 targets, normalizes `image/svg+xml` to `image/svg`, quantizes position/size by at most 360 EMU and `srcRect` by at most 0.003%, and may render the PNG fallback after save.

SVG DOM editing, external SVG relationships, image rounding/transparency, alt-text editing, image hyperlink/shadow and advanced placeholder styling, and public per-image deletion/media garbage collection remain outside this slice. Picture-placeholder population itself is supported by the named master/layout API below. Strict embedded media creation is documented in the Media section below.

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

`CustomGeometryValue`, `CustomGeometryFormula`, `CustomGeometryGuide`, `CustomGeometryPoint`, `CustomGeometryXyHandle`, `CustomGeometryPolarHandle`, `CustomGeometryHandle`, `CustomGeometryConnectionSite`, `CustomGeometryTextRectangle`, `CustomGeometryCommand`, `CustomGeometryPathFill`, `CustomGeometryPath`, and `CustomGeometry` describe the public custom-geometry tree consumed by `SlideModel.addCustomShape(geometry, options)` and returned by `ShapeModel.customGeometry`. Path width/height are positive safe-integer coordinate-space extents. Point/radius values are direct EMU and arc angles are OOXML `1/60000°` when numeric; `inches()` and `degrees()` are explicit ergonomic converters. Point and arc fields are `CustomGeometryValue`, so they may instead reference one guide or DrawingML built-in token. The command union contains `moveTo`, `lineTo`, `arcTo`, `quadraticBezierTo`, `cubicBezierTo`, and `close`. Repeated moves form multiple subpaths; multiple paths, empty command lists, negative numeric point coordinates, zero numeric angles, and optional path `fill` / `stroke` / `extrusionOk` flags are supported. Arc commands deliberately have no endpoint because OOXML derives it from the current point, radii, and angles.

`CustomGeometry.adjustments` and `CustomGeometry.guides` map to `a:avLst` and `a:gdLst`. Each ordered `CustomGeometryGuide` has a globally unique name and a typed formula. `CustomGeometryFormula` supports the complete 17-operator grammar with exact tuple arity: unary `val`, `abs`, `sqrt`; binary `at2`, `cos`, `max`, `min`, `sin`, `tan`; ternary `*/`, `+-`, `+/`, `?:`, `cat2`, `mod`, `pin`, `sat2`. Every operand is a `CustomGeometryValue`. Its numeric branch must be a finite safe integer and normalizes negative zero to zero; its string branch must be non-empty, contain no XML whitespace or invalid XML 1.0 character, and not be a signed decimal integer string. Names use the same single-token lexical contract and must be unique across both lists. The strict codec validates the stored tree; the public evaluator separately resolves dependencies, built-ins, arithmetic domains, and numeric tree values.

`CustomGeometry.handles` is an ordered `CustomGeometryHandle[]` matching the `a:ahLst` choice order. `CustomGeometryXyHandle` has `kind: 'xy'` and maps to `a:ahXY`; its `xGuide`, `yGuide`, `minX`, `maxX`, `minY`, and `maxY` fields map to `gdRefX`, `gdRefY`, `minX`, `maxX`, `minY`, and `maxY`. `CustomGeometryPolarHandle` has `kind: 'polar'` and maps to `a:ahPolar`; `radiusGuide`, `angleGuide`, `minRadius`, `maxRadius`, `minAngle`, and `maxAngle` map to `gdRefR`, `gdRefAng`, `minR`, `maxR`, `minAng`, and `maxAng`. Both kinds require `position`, written as one direct `a:pos`. Position plus XY/radius bounds use direct shape-coordinate safe integers or tokens; numeric angle bounds use direct `1/60000°`. Every guide/bound property is independently optional: the codec does not require refs and bounds together, paired min/max values, resolved guide names, or ordered numeric bounds. Omitted or empty handles normalize to no own property, while mixed XY/polar order and exact optional-property presence are preserved.

`CustomGeometry.connectionSites` is an ordered `CustomGeometryConnectionSite[]` mapped to the direct `a:cxnLst`. Each entry maps to one direct `a:cxn`, requires `angle` mapped to its unqualified `ang` attribute, and requires `position` mapped to one direct `a:pos`. Numeric position coordinates are safe integers written directly in custom-geometry coordinate space; numeric angles are safe integers in direct `1/60000°`. Both fields may instead use the same guide/built-in token branch as `CustomGeometryValue`; the codec preserves those direct tokens and the evaluator resolves them to numbers. Source order and duplicate sites are retained. Neither layer clamps or normalizes angles or verifies attachment to a path. Omitted and empty lists normalize to no own `connectionSites` property.

`CustomGeometry.textRectangle?` is one `CustomGeometryTextRectangle`; required `left`, `top`, `right`, and `bottom` map to `a:rect@l/t/r/b`. Each edge is a `CustomGeometryValue`, so it accepts a finite safe integer or a valid guide/DrawingML built-in token. Omission and explicit `{ left: 'l', top: 't', right: 'r', bottom: 'b' }` both fold to the canonical default and produce no own `textRectangle` property. LibreOffice `textAreaLeft/Top/Right/Bottom` guide normalization is accepted without special coercion.

Creation and replacement strictly normalize dense ordinary arrays and descriptor-safe own data fields, detach the caller, and reject unknown, missing, non-safe, invalid lexical/formula/handle/connection-site/text-rectangle, non-positive numeric extent/radius, or invalid sequence state before mutation. Getter snapshots are detached and deeply frozen through guide lists, formula objects, operand tuples, handle/connection lists and positions, text rectangles, paths, and commands. Empty or omitted guide/handle/connection lists and canonical-default text rectangles normalize to absent optional properties. Assignment replaces the complete custom geometry; assigning the same semantic snapshot is an exact bytes/journal no-op, and `undefined` clear is intentionally not accepted. Setting `presetType` converts custom geometry to preset; setting `customGeometry` converts a preset shape back while preserving model identity, name, transform, styles, text, effects, relationships, and unrelated bytes. The strict reader accepts only namespace-correct direct XY/polar handles and connection sites plus at most one complete direct text rectangle with exact attributes. Malformed state remains byte-preserved, makes the snapshot `undefined`, and rejects replacement before package mutation.

`evaluateCustomGeometry(geometry, context)` evaluates a normalized `CustomGeometry` without reading or mutating a package. `CustomGeometryEvaluationContext` is exactly `{ readonly width: number; readonly height: number }`; both values must be positive safe integers. `ShapeModel.evaluateCustomGeometry()` is the live convenience method: it uses the current shape transform as context, returns `undefined` for preset geometry or an unreadable strict custom snapshot, and otherwise returns the same numeric tree. Resizing the shape therefore changes the next live result without changing the stored formulas.

All 37 DrawingML built-ins are supported: `3cd4`, `3cd8`, `5cd8`, `7cd8`, `b`, `cd2`, `cd4`, `cd8`, `h`, `hc`, `hd2`, `hd3`, `hd4`, `hd5`, `hd6`, `hd8`, `l`, `ls`, `r`, `ss`, `ssd2`, `ssd4`, `ssd6`, `ssd8`, `ssd16`, `ssd32`, `t`, `vc`, `w`, `wd2`, `wd3`, `wd4`, `wd5`, `wd6`, `wd8`, `wd10`, and `wd32`. Adjustment guides are evaluated first in source order, followed by shape guides in source order. A successfully evaluated custom guide shadows a same-name built-in for later values; while defining that guide, a same-name operand still resolves to the built-in. Cycles take precedence over forward-reference reporting, and unknown tokens are distinguished from both.

All 17 formula operators produce JavaScript numbers using DrawingML units. Finite fractional output is preserved; no integer rounding is applied. Negative zero is canonicalized to positive zero. `*/` and `+/` return zero when their denominator is zero. Negative `sqrt`, non-positive evaluated arc radii, and non-finite arithmetic are rejected. Every operand is resolved in source order before the operator is applied.

The evaluator resolves adjustment/shape guide values, handle positions and optional numeric bounds, connection-site positions and angles, text-rectangle edges, and every token-bearing path command field. It preserves handle guide-reference names and path width/height/flags. When `CustomGeometry.textRectangle` is omitted, evaluation always materializes `{ left: 0, top: 0, right: context.width, bottom: context.height }`; omitted optional arrays remain absent. It intentionally does not scale path coordinates from path extents into the shape context, derive arc endpoints, calculate geometric bounds, drag handles, or snap/create connectors.

The full public output type set is `EvaluatedCustomGeometryGuide`, `EvaluatedCustomGeometryPoint`, `EvaluatedCustomGeometryTextRectangle`, `EvaluatedCustomGeometryCommand`, `EvaluatedCustomGeometryXyHandle`, `EvaluatedCustomGeometryPolarHandle`, `EvaluatedCustomGeometryHandle`, `EvaluatedCustomGeometryConnectionSite`, `EvaluatedCustomGeometryPath`, and `EvaluatedCustomGeometry`. The public error surface is `CustomGeometryEvaluationErrorCode` plus `CustomGeometryEvaluationError(code, message, guideName?, token?)`; codes are `unknown-token`, `forward-reference`, `cyclic-reference`, `invalid-domain`, and `non-finite-result`. Standalone geometry/context normalization detaches caller state, evaluation performs no package mutation, and the complete output graph—including context, arrays, entries, commands, and points—is recursively frozen.

PptxGenJS 4.0.1 legal `ShapeType.custGeom` points import as matching final native snapshots: first ordinary point becomes `moveTo`, later ordinary points become `lineTo`, `moveTo: true` starts another subpath, and arc/quadratic/cubic/close map to their direct command branches. Its runtime converts numeric values below 100 and numeric strings as inches, keeps numeric values at or above 100 direct, resolves percentages against the full slide rather than the shape, ignores arc endpoint `x/y`, omits unknown curve kinds, and can emit malformed first-curve/zero-radius/negative-radius/unsafe output. The adapter imports supported final paths and byte-preserves malformed shapes with `customGeometry === undefined`; native does not copy those heuristics or coercions and requires explicit direct values. PptxGenJS 4.0.1 has no public guide-formula, arbitrary adjustment-handle, connection-site, or text-rectangle input and emits only an empty `a:cxnLst` plus the canonical default `a:rect`; evaluator parity is limited to those supported final numeric paths and that default rectangle. Formulas, handles, sites, arbitrary rectangles, and their evaluation are native extensions.

The release gallery was produced from the actual npm tarball and contains 4 slides, 54 shapes, and 22 evaluator targets. Strict reopen evaluates 22/22 targets; the source and LibreOffice round-trip files both pass PowerPoint 2010 validation with 0 errors and 0 warnings, render at 2400×1350 without overflow, and were visually reviewed slide by slide. LibreOffice rewrites direct expressions for 22/22 targets. Numeric path commands plus effective text rectangles match for 21/22; only `Operator sqrt` changes its path endpoint from `600000` to `0`. All 22 guide arrays are rewritten, and the handles and connection sites of the combined handle/site target are removed or rewritten, leaving 21/22 direct matches for each. Both saved files still evaluate strictly for 22/22 targets; these are recorded client rewrites, not hidden as native parity.

`ShapeAdjustment` is the readonly `{ readonly name: string; readonly value: number }` direct guide value used by preset-only `AddShapeOptions.adjustments` and `ShapeModel.adjustments`. A value must be a safe integer and is written directly as the operand of `a:gd@fmla="val N"`; the API performs no shape-specific conversion, clamping, or range inference. Lists must be dense, ordered, and uniquely named. Normalization is descriptor-safe and getter-free, immediately detaches from the caller, and getter snapshots are detached and deeply frozen. Assignment replaces the whole ordered list, `[]` keeps an empty `a:avLst`, and the setter does not accept `undefined`; assigning the same list is an exact bytes/journal no-op. Only one namespace-correct direct preset geometry with one direct adjustment list and simple `val` guides is supported. Complex formulas, duplicate or ambiguous guides, wrong namespaces, unsafe integers, and custom geometry read as `undefined`; replacement throws `ModelParseError` before package changes. A different `presetType` resets adjustments, while same-type assignment preserves their exact bytes. Valid PptxGenJS 4.0.1 `rectRadius`, `angleRange`, and `arcThicknessRatio` output imports as the same final integer list. Native deliberately retains explicit zero and rejects PptxGenJS string coercion, zero truthiness loss, shortcut precedence, ignored thickness-without-angles, and malformed/unsafe passthrough. The separate custom-geometry API above covers paths, guide formulas, handles, connection sites, text rectangles, and numeric evaluation; the preset-only adjustment API intentionally remains a direct-state list.

`ShapeFill` supports direct none or solid sRGB/theme color with optional finite 0–100 transparency rounded to 0.001%. `AddShapeOptions.fill` and `AddTextOptions.fill` use this same value. `ShapeModel.fill` returns a detached direct-state snapshot; same-value assignment is an exact no-op, `{ kind: 'none' }` writes direct no-fill, and `undefined` clears only the direct fill choice. Existing gradient, picture, pattern, and group fills survive unrelated edits and can be explicitly replaced or cleared, but their creation remains outside this simple-fill API.

`ShapeLine` supports direct none or a solid sRGB/theme line with optional finite 0–100 transparency, optional 0–1584 point width, and optional `solid | dash | dashDot | lgDash | lgDashDot | lgDashDotDot | sysDash | sysDot` dash. Omitted width/dash materialize as 1pt/solid; zero width remains direct zero. `ShapeModel.line` returns a detached direct-state snapshot; same-value assignment is an exact no-op, `{ kind: 'none' }` writes direct line no-fill, and `undefined` clears only owned width/fill/dash while preserving the line container, arrowheads, joins, extensions, and unrelated attributes. Existing advanced line fills and custom dashes survive unrelated edits and can be explicitly replaced or cleared, but their creation remains outside this simple-line API.

`ShapeArrowType` is `none | arrow | diamond | oval | stealth | triangle`. `ShapeArrows` has optional readonly `begin` / `end` fields and is used by `AddShapeOptions.arrows` and `ShapeModel.arrows`. Inputs and getter results are detached; the setter is a whole replacement, so a missing side clears that endpoint, explicit `none` stays distinguishable from absence, and `undefined` or an empty object clears both. Same-value assignment is an exact no-op. Arrow edits preserve line width/fill/dash, joins, extensions, advanced line state, and unrelated attributes; line clear/edit operations preserve arrows. Arrows-only creation writes a line container with endpoints but does not synthesize color, width, or dash. A unique safe existing endpoint may carry legal `w` / `len` values `sm | med | lg`; type replacement preserves them lexically, but size is not returned or editable. Malformed, duplicate, reversed, wrong-namespace, or unsupported endpoints read as `undefined` and reject arrow mutation without package changes.

`ShapeShadow` is the strict outer/inner direct-state union used by `AddShapeOptions.shadow`, `AddTextOptions.shadow`, and `ShapeModel.shadow`. Both branches accept optional `RichTextColor`, finite `0..1` opacity, `0..100` point blur, `0 <= angle < 360` degrees, and `0..200` point distance; only outer accepts `rotateWithShape`. Omitted fields normalize to black, 0.75, 8pt, 270°, 4pt, and outer rotate false, while every explicit zero remains zero. Inputs and nested colors are detached before mutation; getter snapshots are detached and deep-frozen. Assignment is a whole replacement, same-value assignment is an exact bytes/journal no-op, kind switches replace only the direct shadow child, and `undefined` removes only that child while retaining `effectLst` plus legal glow, preset-shadow, reflection, soft-edge, blur, or fill-overlay siblings. A malformed, ambiguous, wrong-namespace, `effectDag`, or unsafe schema-order state reads as `undefined` and rejects mutation without package changes. PptxGenJS 4.0.1 omitted shadow and `type: 'none'` map to native `undefined`; its `offset` is native `distance`. Native preserves explicit zero, honors outer rotation, emits legal inner XML, supports theme colors, and rejects invalid passthrough instead of copying PptxGenJS's falsy fallback, ignored rotate flag, malformed inner closing tag, or out-of-range output. Generic effect stacks, custom shadow transforms, preset-shadow editing, and shadow APIs for images, tables, charts, media, and other owners remain outside these focused shape/text APIs.

`Hyperlink` is the mutually exclusive `{ readonly url: string; readonly tooltip?: string } | { readonly slide: number; readonly tooltip?: string }` value used by `AddShapeOptions.hyperlink`, `AddTextOptions.hyperlink`, `ShapeModel.hyperlink`, `RichTextRunStyle.hyperlink`, and `AddTableCellOptions.hyperlink`; readonly `TableCell.hyperlink` exposes supported plain-cell direct state. A URL must be a non-empty XML-safe string; a slide number must be a one-based positive safe integer resolving to a current presentation slide at assignment time. Inputs must be descriptor-safe ordinary or null-prototype objects with exactly one target and no unknown keys. Getter results are detached frozen direct-state snapshots. Tooltip absence remains property absence, while direct empty remains `tooltip: ''`; shape assignment is a whole replacement, omitted tooltip clears only that attribute, and `undefined` removes the supported whole-shape click element. Same-value assignment is an exact bytes/journal no-op. URL/slide switching reuses an unshared relationship or clones on write when its ID is referenced elsewhere, and clear or replacement garbage-collects only unreferenced relationships. Internal links retain target-part identity while slide insert/delete/reorder changes the reported one-based ordinal; duplicate self-links retarget to the duplicate, and deleting a target removes incoming DrawingML click/hover elements before deleting their relationships. `AddTextOptions.hyperlink` supplies the non-visual click and default run link. `RichTextRunStyle.hyperlink?: Hyperlink | false` overrides that default with an independent run relationship, inherits it when omitted, or suppresses it with `false`; explicit run underline always wins. Each linked plain table cell owns an independent relationship on its only direct run. `ShapeModel.hyperlink` deliberately remains whole-shape-only, while shape run links are read and edited through `ShapeModel.richText`. Rich table-cell links are detached snapshots under `TableCell.richText` and are whole-replaced through `setCellRichText()` with relationship reuse/COW/GC. Unsupported hover editing, extra action/sound/history ownership, duplicate/malformed click ownership, or dangling/wrong-type relationships are never guessed. PptxGenJS 4.0.1 materializes omitted tooltip as direct empty and may console-ignore, coerce, duplicate, or dangle invalid runtime targets; its rich outer hyperlink emits broken `rIdundefined` references, while legal rich per-run links omit the shape click and allocate separate relationships. Native accepts its compatible extra external-run attributes but rejects its defects before mutation. External hyperlinks produce the expected portability warning rather than a package error. Table graphic-frame/image/chart/media/group hyperlink creation, hover links, action-only navigation, and relative/file safety policy remain outside this API.

Arrow size, cap/compound/alignment/join editing, generic/advanced effects, custom shadow transforms, non-shape/text shadow APIs, custom-geometry path scaling/arc endpoint and bounds calculation/handle dragging/connector snapping and creation, advanced line fill/custom dash creation, and percentage positions remain pending. Text-shape simple-line, arrow, simple-shadow, outer-hyperlink, preset-geometry, rounded-rectangle-radius, direct `isTextBox`, and rich-text `breakLine` creation/editing reuse the same codecs below.

Shape kinds include `text`, `shape`, `image`, `table`, `chart`, `graphic-frame`, and `group`. Images expose embedded part URIs and replacement; tables support basic native creation plus rows/cells, cell text, borders, fill, margins, horizontal/vertical alignment, text direction, cell text-fit editing, and plain single-run hyperlink creation/read snapshots, with table-level consensus/bulk editing for borders, fill, margins, horizontal/vertical alignment, and text direction; charts expose cached series and lossless chart XML editing.

### Text-shape direct fill

```ts
const slide = document.addSlide();
const plain = slide.addText('Solid text box', {
  fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: 'D9EAF7' },
    transparency: 25,
  },
});
const rich = slide.addRichText([{
  runs: [{ text: 'Theme-filled rich text' }],
}], {
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent2' },
    transparency: 0,
  },
});
const placeholder = slide.addPlaceholder('Filled placeholder', {
  name: 'filled_title',
  type: 'title',
  fill: { kind: 'none' },
});

plain.fill = { kind: 'solid', color: { kind: 'scheme', value: 'accent3' } };
plain.fill = { kind: 'none' };
plain.fill = undefined;
```

`AddTextOptions.fill?: ShapeFill` is normalized before any package mutation. The supported union is exactly direct none or direct solid with a required strict six-digit sRGB/supported scheme color and optional finite `0..100` transparency rounded to `0.001%`. Omitted, runtime-`undefined`, and explicit none creation all produce canonical direct `a:noFill`; solid explicit zero transparency produces direct `a:alpha val="100000"`. Unknown, inherited, accessor, or symbol keys; non-ordinary objects; missing color; invalid lexical colors; and non-finite or out-of-range transparency reject without changing parts, relationships, XML, runtime caches, shape order, or the mutation journal.

Plain/rich text, `addPlaceholder()`, title/body placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share the same renderer and validation contract. The resulting `ShapeModel.fill` is immediately readable and editable. Inputs detach before mutation; snapshots are detached and frozen; duplicate, move, outer rollback, six-format write/reopen, stable identity, sibling isolation, and placeholder-source isolation retain the direct fill. Clearing the live model with `undefined` differs from creation omission: the former removes the direct choice, while the latter intentionally preserves the canonical text-box default no-fill.

PptxGenJS 4.0.1 writes direct no-fill for omitted text fill, omits the direct choice for `{ type: 'none' }`, and omits alpha for explicit zero transparency. Native deliberately preserves explicit none and zero direct intent. Supported solid/scheme/non-zero-alpha output reaches the same final semantics. Gradient/pattern/picture/group text-fill creation remains outside this simple creator. Text outer simple line, arrows, simple shadow, hyperlink, preset geometry, rounded-rectangle radius, direct `isTextBox` state, and rich-text `breakLine` are supported below.

### Text-shape direct simple line

```ts
const outlined = document.addSlide().addText('Outlined text box', {
  line: {
    kind: 'line',
    color: { kind: 'srgb', value: '2F5597' },
    transparency: 25,
    width: 2.5,
    dash: 'dashDot',
  },
});
const themed = document.layouts[0].addRichText([{
  runs: [{ text: 'Theme outline' }],
}], {
  line: { kind: 'line', color: { kind: 'scheme', value: 'accent2' } },
});

outlined.line = { kind: 'line', color: { kind: 'scheme', value: 'accent3' } };
outlined.line = { kind: 'none' };
outlined.line = undefined;
```

`AddTextOptions.line?: ShapeLine` is normalized before package mutation. The union is exactly direct none or a solid line with required strict six-digit sRGB/supported scheme color plus optional finite `0..100` transparency, finite `0..1584` point width, and one of eight preset dash tokens. Transparency is rounded to `0.001%`, width to one EMU, and omitted width/dash materialize as 1pt/solid. Omitted, runtime-`undefined`, and explicit none creation preserve canonical direct text-box no-fill; explicit zero transparency writes `a:alpha val="100000"` and zero width writes `a:ln@w="0"`. PptxGenJS-shaped aliases, missing color, invalid range/token, non-ordinary objects, unknown/inherited/accessor/symbol keys, and class instances reject before parts, relationships, XML, caches, shape order, or the mutation journal change.

Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share one renderer. Line stays after geometry and shape fill. The resulting `ShapeModel.line` immediately returns a detached normalized snapshot and keeps the existing read/whole-replace/clear semantics: same value is an exact no-op, none writes direct no-fill, and `undefined` removes owned width/fill/dash while preserving the line container and unrelated state. Duplicate, move, outer rollback, six-format write/reopen, stable identity, sibling isolation, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 emits empty `a:ln` for omitted/none/empty/missing-color text line, omits direct width/dash for their defaults, collapses width/transparency zero, honors nested deprecated `alpha`, and ignores nested `lineDash`. Native uses explicit reversible direct state and rejects permissive aliases/fallbacks. Supported sRGB/theme, non-zero transparency, positive width, and all eight dash tokens reach equivalent final semantics. Gradient/pattern/picture/group line fills, custom dash, cap/compound/alignment/join, and the remaining text geometry shortcuts stay outside this simple-line slice. Text arrows, simple shadow, hyperlink, and preset geometry are supported below.

Release evidence is 1268 passed / 1 skipped tests, performance 1/1 at 553ms, both TypeScript builds, both package builds, an actual 57-file tarball with `textShapeLines: true`, real-Chrome exact immediate/detached/reopen state with zero console/page/network errors, and installed CLI PowerPoint 2010 validation at 0 errors / 0 warnings.

### Text-shape direct arrows

```ts
const arrowed = document.addSlide().addText('Flow', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
});
const themed = document.layouts[0].addRichText([{
  runs: [{ text: 'Theme endpoints' }],
}], {
  arrows: { begin: 'none', end: 'stealth' },
});

arrowed.line = undefined;
arrowed.arrows = { begin: 'oval' };
arrowed.arrows = undefined;
```

`AddTextOptions.arrows?: ShapeArrows` is normalized before package mutation. Optional `begin` / `end` are limited to `none | arrow | diamond | oval | stealth | triangle`; omission, runtime `undefined`, or `{}` creates no direct endpoints, while explicit `none` is retained. PptxGenJS `beginArrowType` / `endArrowType`, deprecated `lineHead` / `lineTail`, empty or invalid tokens, unknown/inherited/accessor/symbol keys, class instances, arrays, and other non-ordinary objects reject without changing parts, relationships, XML, caches, shape order, or the mutation journal.

Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share the existing `ShapeArrows` codec and one text renderer. Direct child order is line fill/dash, `headEnd`, then `tailEnd`. Arrow-only native creation preserves canonical direct `a:noFill`. `ShapeModel.arrows` immediately returns a detached snapshot and uses whole replacement; explicit endpoint `none` survives, a missing side is removed, and `undefined` clears both endpoints. Line and endpoint ownership are independent in both directions. Duplicate, move, outer rollback, asynchronous declarative detachment, all six formats, write/reopen, stable identity, sibling isolation, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 emits an empty line for omitted text line paint and omits no-fill for arrow-only or `{ type: 'none' }` plus endpoints. It ignores empty endpoints and nested/top-level `lineHead` / `lineTail`, while invalid runtime tokens pass through as malformed direct endpoint state. Native keeps canonical no-fill and rejects aliases or invalid tokens before mutation. All six legal endpoint tokens, begin/end/both, explicit `none`, and combined solid line semantics are otherwise compatible; malformed imported endpoints remain losslessly preserved but are absent from the strict snapshot.

Release evidence is 1274 passed / 1 skipped tests, performance 1/1 at 581ms, both TypeScript and package builds, an actual 57-file tarball with `textShapeArrows: true`, real-browser exact immediate/detached/reopen state with zero errors, and installed CLI PowerPoint 2010 validation at 0 errors / 0 warnings. The following section documents text-shape simple shadow support.

### Text-shape direct simple shadow

```ts
const shadowed = document.addSlide().addText('Shadowed heading', {
  line: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent2' },
    width: 2,
    dash: 'dashDot',
  },
  arrows: { begin: 'triangle', end: 'arrow' },
  shadow: {
    kind: 'outer',
    color: { kind: 'scheme', value: 'accent4' },
    opacity: 0.4,
    blur: 2,
    angle: 45,
    distance: 3,
    rotateWithShape: true,
  },
});
const inner = document.layouts[0].addRichText([{
  runs: [{ text: 'Inner shadow' }],
}], {
  shadow: {
    kind: 'inner',
    color: { kind: 'srgb', value: '667788' },
    opacity: 0,
    blur: 0,
    angle: 0,
    distance: 0,
  },
});

shadowed.shadow = { kind: 'inner', opacity: 0.25 };
shadowed.shadow = undefined;
```

`AddTextOptions.shadow?: ShapeShadow` uses the same strict direct-state normalizer as preset shapes. Outer and inner accept valid sRGB/theme color, finite `0..1` opacity, `0..100pt` blur, `0 <= angle < 360°`, and `0..200pt` distance; only outer accepts boolean `rotateWithShape`. Defaults are black, 0.75, 8pt, 270°, 4pt, and outer rotate false. Every explicit zero is retained. PptxGenJS `type` / `offset` aliases, coercible strings, invalid ranges, unknown/inherited/accessor/symbol keys, arrays, and class instances reject before package mutation.

Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrapper methods, and declarative `defineSlideMaster()` text/placeholder objects share the same renderer. Direct child order is geometry, fill, line/endpoints, then one `effectLst`. Fill, line, arrows, and effect ownership are independent. The returned `ShapeModel.shadow` immediately exposes a detached deep-frozen snapshot; assignment is a whole replacement, same-value assignment is an exact bytes/journal no-op, and `undefined` clears only the direct inner/outer child while preserving a safe effect-list container plus sibling effects. Duplicate, move, outer rollback, asynchronous declarative detachment, all six formats, write/reopen, stable identity, sibling isolation, and placeholder-source isolation are covered.

PptxGenJS 4.0.1 omits direct effects for omitted shadow and `{ type: 'none' }`; supported outer cases reach equivalent final semantics, and legacy `offset` corresponds to native `distance`. It falls back from runtime zero to defaults, ignores text `rotateWithShape: true`, warning-corrects or loosely converts some invalid type/color/number inputs, and emits a mismatched closing tag for inner shadow. Native deliberately retains zero, supports theme color and rotate true, emits legal inner XML, and rejects invalid input with zero mutation.

Release evidence is 1280 passed / 1 skipped tests, performance 1/1 at 607ms, both TypeScript builds, both package builds, and the declaration build. The actual 57-file tarball reports `textShapeShadows: true` from installed Node, declarations, browser export, and installed CLI checks. Real-browser exact immediate/detached/reopen state has zero console/page/network errors, and installed CLI PowerPoint 2010 validation is 0 errors / 0 warnings.

### Text-shape outer hyperlink

```ts
const plain = slide.addText('Open product', {
  hyperlink: { url: 'https://example.com/product', tooltip: 'View product' },
});
const rich = slide.addRichText([{
  runs: [
    { text: 'Go to ' },
    { text: 'details', style: { underline: false } },
  ],
}], {
  hyperlink: { slide: 2, tooltip: '' },
});

plain.hyperlink = { url: 'mailto:team@example.com' };
plain.hyperlink = undefined;
```

`AddTextOptions.hyperlink?: Hyperlink` is available on plain/rich text, `addPlaceholder()`, placeholder population, slide/layout/master wrappers, and declarative master text/placeholder objects. Normalization is descriptor-safe and getter-free, detaches immediately, and rejects invalid, coercible, or dangling targets before mutation. Creation writes one relationship ID to the direct non-visual `hlinkClick` and each non-empty run that does not override it; empty paragraphs/runs do not create references.

`RichTextRunStyle.hyperlink?: Hyperlink | false` is the direct run surface. An explicit value allocates an independent relationship even when another run has the same target; omission inherits `AddTextOptions.hyperlink` during creation, and `false` suppresses that default. A linked run without an explicit underline gets `u="sng"`; any explicit underline, including `false`, remains authoritative. Strict reads expose only one valid direct namespace-correct run click and preserve omitted versus empty tooltip. Whole-rich-text replacement provides exact no-op, index-stable ID reuse, unique-target in-place update, shared clone-on-write, clear/GC, validation-before-mutation, transaction rollback, and write/reopen behavior.

The returned live `ShapeModel.hyperlink` is the whole-shape click surface; replacement/clear uses reference-aware clone-on-write and garbage collection while preserving run-local links. URL/internal target identity and tooltip state survive slide/layout/master/placeholder/declarative owners, duplicate, move, target delete, rollback, all six formats, and write/reopen. Duplicate self-links retarget to the duplicate, and target deletion removes incoming direct shape/run clicks.

PptxGenJS 4.0.1 writes `tooltip=""` when omitted. Its plain outer output has shape/run links sharing one relationship; rich outer output instead writes broken `rIdundefined` references. Its legal rich per-run form writes only run links and allocates one relationship per run. Native accepts external-run `action=""` but not orphaned/dangling IDs, falsy-underline loss, loose coercion, or console-only failure.

Release evidence is 1303 passed / 1 skipped tests and performance 1/1 at 624ms. Model, SDK, root, and adapter suites are 199/199, 191/191, 13/13, and 80/80; both TypeScript builds, both tsup builds, and declaration build pass. The actual 57-file tarball plus installed Node/declarations/browser/CLI and real Chrome report `richTextRunHyperlinks: true`; Chrome has zero validation/console/page/network errors. The 24-part external deck validates with 0 errors and 8 expected portability warnings; the 20-part internal-only deck validates with 0 errors / 0 warnings. Package inspect, slide listing, and exact part reads pass.

### Text-shape preset geometry

```ts
const text = document.addSlide().addText('Shaped text', {
  shape: 'ellipse',
  line: { kind: 'line', color: { kind: 'scheme', value: 'accent1' } },
});

console.log(text.presetType); // 'ellipse'
text.presetType = 'hexagon';
```

`AddTextOptions.shape?: PresetShapeType` accepts exactly the 178 frozen canonical values in `PRESET_SHAPE_TYPES`; omitted and own data-property `undefined` both create `rect`. Plain/rich text, `addPlaceholder()`, named placeholder population, slide/layout/master methods, and declarative master text/placeholder objects share this contract. The resulting `ShapeModel.presetType` is live immediately. Same-token assignment is an exact bytes/journal no-op that retains existing adjustments; another token replaces only direct geometry with a canonical empty adjustment list.

The field is normalized as an own data property before package mutation. Empty/falsy values, `folderCorner`, `custGeom`, unknown strings, numbers, booleans, coercion, inherited values, and accessors are rejected rather than trimmed, converted, or guessed. Native exposes correct OOXML `foldedCorner`; PptxGenJS 4.0.1 instead exposes invalid `folderCorner` and omits `foldedCorner`. Native compares the 177 common valid tokens semantically but does not copy PptxGenJS's falsy fallback, unchecked passthrough, line-without-line-option exception, or runtime custom-geometry route. Use `addCustomShape()` / `ShapeModel.customGeometry` for custom geometry.

Geometry ownership is independent from transform, fill, line, arrows, shadow, whole-shape/run hyperlinks, text body, and placeholder identity. In particular, `shape: 'line'` selects preset line geometry while `AddTextOptions.line` controls outline style. Duplicate, move, rollback, all six presentation formats, write/reopen, stable identity, and layout/master placeholder-source isolation are covered.

Final verification is 1313 passed / 1 skipped tests and performance 1/1 at 578ms; both TypeScript checks, both tsup builds, and declaration build pass. The actual 57-file tarball (SHA-256 `1412706458c883b9e4dfa3d87e6577ab86df57b78999a2796b2c6e69647be0f9`) reports `textShapePresetGeometry: true` from installed Node/types/browser/CLI. Real Chrome reports the same with zero validation/console/page/network errors. The representative two-slide source validates at PowerPoint 2010 0/0; after LibreOffice save all 17 `(text, presetType)` pairs survive, with 0 errors and two placeholder-owner warnings. Both versions render without overflow and match in slide-by-slide visual review.

This is the historical checkpoint at completion of text preset geometry and did not complete advanced text or full PptxGenJS parity at that point. Direct `isTextBox` state and rich-text `breakLine` are supported below; advanced text/table, `tableToSlides`, output/runtime helpers, and the peer-range full-suite audit were then still pending.

### Text-shape rounded rectangle radius

```ts
const rounded = document.addSlide().addText('Rounded text', {
  shape: 'roundRect',
  rectRadius: inches(0.5),
  width: inches(4),
  height: inches(2),
});

rounded.adjustments = [{ name: 'adj', value: 12500 }];
```

`AddTextOptions.rectRadius?: Emu` is an own-data-property creation shortcut that is valid only when the resolved shape is `roundRect`. A supplied value must be a finite number in inclusive range `0..914400`, rounds to the nearest EMU with negative zero canonicalized to zero, and must remain a safe integer after rounding. Omitted, absent, and own `undefined` values leave the preset geometry's canonical empty `a:avLst`; explicit zero creates one direct `{ name: 'adj', value: 0 }` guide.

Creation derives the guide as `Math.round(rectRadius * 100000 / Math.min(finalWidth, finalHeight))`. Plain/rich text, `addPlaceholder()`, named-placeholder population, slide/layout/master methods, and declarative master text/placeholder objects share the contract. Placeholder population uses the selected owner's final inherited transform. The derived guide is ordinary `ShapeModel.adjustments` state: same-value assignment is an exact bytes/journal no-op, whole replacement and `[]` clear are supported, and later transform resizing does not recalculate it. Geometry adjustment ownership remains independent from fill, line, arrows, shadow, hyperlinks, text-body state, transform mutation, and placeholder identity.

Valid positive PptxGenJS 4.0.1 cases produce the same final guide. Native deliberately preserves explicit zero and rejects string coercion, wrong-shape use, negative or over-one-inch values, NaN, infinities, accessors, symbols, inherited keys, and unsafe object shapes before package mutation. It therefore does not reproduce PptxGenJS's zero/NaN truthiness loss, unchecked negative/over-range formulas, or `val Infinity` output.

Final verification is 1320 passed / 1 skipped tests; model, SDK, root, and adapter suites are 203/203, 195/195, 15/15, and 85/85. Both TypeScript checks, both bundles, and declaration generation pass. The actual 57-file tarball plus installed Node/types/browser/CLI and real Google Chrome report `textShapeRectRadius: true`; Chrome validation/console/page/network errors are zero. A three-slide source validates under PowerPoint 2010 at 0/0, mutation isolation changes only `slide1.xml`, overflow is zero, and every rendered page was reviewed. LibreOffice retains all explicit guide values and only normalizes the omitted default to `16667`.

`AddTextOptions.isTextBox?: boolean` controls only direct `p:sp/p:nvSpPr/p:cNvSpPr@txBox`. Omitted, own-data `undefined`, and `false` create attribute absence; `true` creates canonical unqualified `txBox="1"`. The field is read as an own data property, inherited state is ignored, accessors are not invoked, and any defined value must be a primitive boolean. Invalid runtime values reject before parts, relationships, XML, shape order, caches, or the mutation journal change.

`ShapeModel.isTextBox` is `boolean | undefined` on read and boolean-only on assignment. The getter maps absence to `false`, `1/true/on` to `true`, and `0/false/off` to `false`; malformed tokens, qualified lookalikes, repeated attributes, or missing/repeated direct owner structures return `undefined`. Assignment writes canonical `txBox="1"` for true or removes the unique direct attribute for false. A canonical same-value assignment is an exact bytes/journal no-op; a single alias or malformed token can be canonicalized, while ambiguous structure throws `ModelParseError` before mutation.

The contract covers plain/rich text, `addPlaceholder()`, layout/master direct methods, declarative master text/placeholder objects, duplicate/rollback/write/reopen, and all six presentation formats. Layout-placeholder materialization preserves the source direct state. Placeholder population validates the call-site field but the selected layout source wins, matching legal PptxGenJS 4.0.1 output, and source parts remain unchanged. `isTextBox` is independent from preset/custom geometry, adjustments/`rectRadius`, transforms, text bodies, styles, hyperlinks, and placeholder identity. PptxGenJS runtime truthiness outside its boolean type is intentionally rejected.

Final verification is 1337 passed / 1 skipped tests and performance 1/1 at 704ms. The actual 57-file tarball has SHA-256 `2c6afd9bdb1f4c076ff0d0eb8bc8e8711793ae46dc320b2978d08f5e3a44b41a`; installed Node/browser conditional export/declarations/CLI and real Google Chrome report `textShapeIsTextBox: true`, with zero Chrome validation/console/page/network errors. The two-slide source has 0 PowerPoint 2010 errors and one expected external-link warning; one toggle changes only `slide1.xml`, and the other 21 parts remain byte-identical. Source and LibreOffice renders have zero overflow. LibreOffice save removes true `txBox` state from both native and PptxGenJS files, so source OOXML remains the compatibility reference rather than a claimed client round trip.

### Rich-text run paragraph splitting

`RichTextRun.breakLine?: boolean` is a transient creation and whole-replacement field. A primitive `true` on a non-final run ends the current paragraph after that run. Middle, empty, and consecutive flagged runs preserve the corresponding canonical empty paragraphs; a final flag is consumed without creating a trailing empty paragraph. Omitted, `undefined`, and `false` do not split. Every segment receives a detached copy of the source paragraph's alignment, RTL, margins, indent, bullet, level, spacing, and tab stops.

`softBreakBefore` stays on its run even when splitting makes that run the first run of a paragraph. `RichTextRunStyle.hyperlink` URL/internal-slide targets also stay on their runs, and relationship allocation uses the final canonical paragraph/run indexes. The contract covers slide/layout/master content, placeholder prompt/population, declarative master objects, live `ShapeModel.richText` editing, duplicate/move/rollback/reopen, and all six formats. Getter snapshots contain only explicit `RichTextParagraph[]` and never expose or infer `breakLine` markers.

The field belongs only to `RichTextRun`; outer `AddTextOptions.breakLine` is rejected by types and runtime validation. CR/LF inside run text is not a field alias. Values other than primitive booleans reject before package mutation, including strings, numbers, null, objects, and boxed booleans. Legal boolean PptxGenJS 4.0.1 output has equivalent paragraph/property/hyperlink semantics. Its first-run soft-break suppression is intentionally not copied because native preserves explicit reversible `softBreakBefore` state.

Final verification is 1350 passed / 1 skipped tests plus performance 1/1. Both TypeScript checks, Node/browser bundles, and declaration generation pass. The actual 57-file tarball SHA-256 is `d06b84c0c3b8ff8e610c87c55b0fe9b67de6b41e59b5ec7fad62b206fdbe2699`; installed Node/types/browser/CLI and real Chrome report `richTextBreakLine: true`, with zero Chrome validation/console/page/network errors. The four-slide source validates at PowerPoint 2010 0/0, and an internal-link edit changes only `slide1.xml` plus its relationship part while the other 24 parts stay byte-identical. Source and LibreOffice visual decks have zero overflow. LibreOffice preserves visible paragraphs, empty lines, soft breaks, and internal links but merges adjacent runs, omits empty tooltips, pushes master content into layouts, renames placeholders, and drops the master placeholder prompt; owner identity is not a complete round-trip guarantee.

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

`addTable()` accepts non-empty logical rows of strings or strict `{ text: string | readonly RichTextParagraph[], options? }` objects whose supported cell options include strict `colspan` / `rowspan`, font family, font size, bold, color, paragraph spacing, align, border, fill, fit, hyperlink, margin, text direction, vertical alignment, and table/cell auto-page weights. String CR/LF normalization, empty paragraphs, structured paragraph/run styles, soft breaks, and cell-default/local run links are supported. Table defaults for font family, font size, bold, color, spacing, align, border, fill, margin, text direction, and vertical alignment are materialized into physical-cell direct state; cell values win and no creation metadata is retained. There is no table-level hyperlink default. Inputs are descriptor-safe, getter-free, detached, and serialized in stable `tcPr` order: margins, optional anchor/direction, L/R/T/B borders, then fill. A plain cell hyperlink attaches to its one direct run; rich cell defaults apply only to runs without a local override. Geometry uses EMU, with strict scalar or exact-length column/row vectors and synchronized transform dimensions. Automatic/zero/minimum rows support content measurement, repeated headers, rich row fragmentation, placeholder auto-page, continuation Y/margins, and `newAutoPagedSlides`; all-positive rows without weights preserve fixed structural pagination. Existing tables expose paragraph-aware `text`, detached `richText`, indexed safe plain/rich editors, readonly scalar hyperlink snapshots, detached merge snapshots, lossless merge/unmerge, merge-aware physical row/column CRUD, and table-level uniform consensus/all-physical-cell bulk editing through `verticalAlignment`, `textDirection`, `horizontalAlignment`, `margins`, `borders`, and `fill`; mixed detail remains under `rows[].cells[]`, and explicit direction `horz`, direct no-fill, zero-width line, omitted border style, explicit solid, and explicit-zero alpha stay distinct from `undefined` clear. Remaining native-extension work includes logical existing-table content insertion, table-level fit creation/defaults, diagonal/advanced borders and fills, and table creation styles; HTML import is supported by `tableToSlides()`. The detailed sections define exact creation precedence, numeric ranges, OOXML mappings, no-op behavior, and PptxGenJS boundaries for each property.

`AddTableOptions.align` and `AddTableCellOptions.align` reuse the exact `TextAlignment` values `left`, `center`, `right`, and `justify`, mapped to direct `a:pPr@algn` tokens `l`, `ctr`, `r`, and `just`. A table value supplies the default for every cell that omits cell alignment or supplies runtime `undefined`; a valid cell value wins, and an explicit `RichTextParagraph.align` wins for that paragraph. When the table value is omitted or runtime-`undefined`, current bytes are preserved and no effective left token is synthesized. Final ownership is each physical cell paragraph's direct `a:pPr@algn`, never `tcPr`, `bodyPr`, or retained table metadata, so later clearing does not reapply a creation default. Rich/multi-paragraph alignment is read and whole-replaced through `TableCell.richText` / `setCellRichText()`. The scalar `horizontalAlignment` snapshot/editor below deliberately requires one exact direct paragraph. Supported table, cell, and paragraph values and precedence match PptxGenJS 4.0.1 final state; PptxGenJS silently drops an unknown table runtime value, whereas native creation throws `TypeError` before mutation.

`TableModel.columnWidths` reads the unique direct `tblGrid` as a detached exact-EMU snapshot. A malformed or ambiguous grid returns `undefined` instead of guessing from the transform or cells. `setColumnWidths()` accepts a positive scalar broadcast or a dense descriptor-safe exact-length array, rounds each item to a safe EMU integer, rejects unsafe sums, and atomically updates both `gridCol@w` and `ext@cx`. A valid grid/transform mismatch is repaired; a numeric no-op preserves the original slide bytes and mutation journal. Unsafe existing grid or transform XML raises `ModelParseError` without mutation. Inherited `setTransform({ width })` still changes only the transform, so use `setColumnWidths()` when changing table width distribution.

`TableModel.rowHeights` reads all direct `tr@h` values as a detached exact-EMU snapshot; zero is a valid automatic row height. A malformed or ambiguous direct row vector returns `undefined` without guessing from transform height or cell content. `setRowHeights()` accepts a non-negative scalar broadcast or a dense descriptor-safe exact-length array, rejects raw negative values, and rounds each item to a safe EMU integer. When every target is positive, their safe exact sum is written to `ext@cy` and a valid rows/transform mismatch is repaired. When any target is zero, row tokens are updated but the already-valid transform height is preserved because this existing-table direct editor does not invoke creation-time measurement. Numeric no-ops preserve original tokens, slide bytes, and the mutation journal; unsafe existing rows or transform XML raise `ModelParseError` without mutation. During `addTable({ autoPage: true })`, omitting both `height` and `rowHeights`, or using zero row entries, enables measurement; positive row values are minima and all-positive fixed mode remains available. Inherited `setTransform({ height })` still changes only the transform, so use `setRowHeights()` to edit row tokens. Rich/multi-paragraph cell text, table/cell text-style defaults, paragraph/run overrides, default/local links, plain scalar hyperlink editing, colspan/rowspan creation/read/edit, physical row/column insertion/deletion, measured/fixed pagination, repeated headers, rich row fragmentation, and placeholder auto-page are supported in the sections above. Remaining table gaps include cell options outside the documented set, table-level fit creation/defaults, diagonal/advanced borders and fills, logical content insertion, table/cell creation styles, and `tableToSlides`.

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
const currentTableAlignment: TextBoxVerticalAlignment | undefined =
  table?.verticalAlignment;
if (table) table.verticalAlignment = 'middle';
if (table) table.verticalAlignment = undefined;
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

Table creation names the options `AddTableOptions.textDirection` and `AddTableCellOptions.textDirection`; both accept only `horz`, `vert`, `vert270`, or `wordArtVert`. The table value is materialized onto every physical cell whose cell value is omitted or runtime `undefined`, while any explicit cell value wins. Explicit cell `horz` therefore blocks a non-horizontal table default but still creates no direct attribute. Omitted, runtime-undefined, and resolved horizontal creation write no `tcPr@vert`; the three non-horizontal values write their exact token, never `bodyPr@vert`, and appear immediately in the live `TableCell.textDirection` snapshot. Omitted/runtime-undefined table direction preserves existing bytes, and explicit table `horz` produces the same direction bytes when cells do not override it. Inputs follow the same descriptor-safe, getter-free, detached normalization as other table options, and invalid or text-box-only tokens are rejected before mutation. Creation retains only final physical-cell direct state, not table metadata, so a later clear never re-inherits the original table value. `setCellTextDirection()` addresses physical zero-based row/cell positions and deliberately retains lossless editor semantics: explicit `horz` writes `vert="horz"`, while `undefined` clears the direct attribute. The cell getter requires one direct `tcPr`, one unqualified `vert`, and one exact public token. `TableModel.textDirection` requires the same safe direct value on every physical cell; it returns `undefined` for absent, mixed, empty, or unsafe state, without resolving inheritance or synthesizing absent state as `horz`. Assignment atomically broadcasts one legal token or clears all direct tokens with `undefined`, while `rows[].cells[].textDirection` retains mixed detail. PptxGenJS 4.0.1 produces the same supported table/cell final state and precedence, including horizontal collapse; imported horizontal creation therefore reads `undefined`, whereas explicit native `horz` and existing-deck bulk editing are lossless extensions. Its runtime invalid-token passthrough remains unsupported by the strict native API. This four-value table-cell API is separate from the seven-value text-box API.

Table-cell `horizontalAlignment` is an immutable `TextAlignment` snapshot for a strict cell containing exactly one direct text body, one direct paragraph, and one direct paragraph-properties element. It reads only the unqualified direct `pPr@algn` tokens `l`, `ctr`, `r`, and `just` as `left`, `center`, `right`, and `justify`; missing, malformed, ambiguous, zero-paragraph, or multi-paragraph cells return `undefined` without resolving effective alignment or a retained table default. `setCellHorizontalAlignment()` uses physical zero-based row/cell coordinates, writes the canonical direct token, and uses `undefined` to clear only that unqualified direct attribute. Assigning the current value or clearing an absent token is an exact no-op. Unsafe cell structure raises `ModelParseError` without mutation, while unrelated paragraph content, namespaced attributes, cell properties, and neighbor cells remain byte-preserved. PptxGenJS 4.0.1 table/cell alignment is materialized into the same direct cell-paragraph state and imports to the same snapshots; native existing-deck editing is a lossless extension because PptxGenJS has no existing-deck editor. Rich or multi-paragraph alignment is supported separately through `RichTextParagraph.align`, `TableCell.richText`, and `setCellRichText()`.

Table-cell creation names the option `AddTableCellOptions.valign`; the immutable snapshot is `TableCell.verticalAlignment`. Both reuse exact `TextBoxVerticalAlignment` values `top`, `middle`, and `bottom`, stored only in the selected physical cell's direct `tcPr@anchor` as `t`, `ctr`, or `b`. Omitted or runtime-undefined cell values write no anchor unless `AddTableOptions.valign` supplies the same strict value as a creation default; a cell value overrides that default. The table value is materialized directly onto uncovered physical cells and is not retained as table metadata, so a later `setCellVerticalAlignment(..., undefined)` clears the anchor without inheritance. Omitting the table value preserves the equivalent default bytes. The returned live table exposes every resolved value immediately, and `setCellVerticalAlignment()` writes the canonical token or clears it with `undefined`. The strict snapshot getter requires exactly one direct `tcPr` and one unqualified supported anchor. It does not read or change the separate `bodyPr@anchor`, resolve inheritance, or support `just` / `dist`. `TableModel.verticalAlignment` scans the exact direct `graphic → graphicData → tbl → tr → tc` path and returns a value only when one or more physical cells, including merge continuations, all expose the same safe direct anchor. Mixed, absent, empty, malformed, or ambiguous state returns `undefined`; assigning a legal value atomically replaces every physical cell, assigning `undefined` clears every direct anchor, and unsafe edits reject without partial mutation. PptxGenJS 4.0.1 produces the same final direct state for valid table- and cell-level creation values, including cell precedence and no anchor when both layers are omitted. Runtime invalid PptxGenJS tokens remain opaque and import as `undefined`, while the native API rejects them. Existing-deck consensus/bulk editing is an intentional native lossless extension because PptxGenJS has no table object editor.

Table-cell `margins` reuses the point-based `TextBoxMargins` value shape but owns only direct `tcPr@marL/marR/marT/marB`; it never reads or changes text-box `bodyPr@*Ins`. `AddTableOptions.margin`, `AddTableCellOptions.margin`, `setCellMargins()`, and `TableModel.margins` assignment accept a point scalar, `[top, right, bottom, left]` tuple, partial named object, `{}` or `undefined`, with descriptor-safe getter-free normalization and immediate detachment. Creation resolves canonical top/bottom 3.6pt and left/right 7.2pt, then table sides, then cell sides. Cell scalar/TRBL replaces all four inherited sides; partial cell input replaces only supplied sides; omitted/undefined/empty cell input inherits the table layer. Omitted/undefined/empty table input preserves canonical bytes. The resolved values are materialized directly on physical cells and immediately exposed by `TableCell.margins`; no creation-default metadata is retained, so clearing a cell does not reapply the creation input. The cell editor addresses zero-based physical row/cell positions and is a whole replacement, so its omitted named sides are cleared. `TableModel.margins` returns a detached complete or partial snapshot only when every physical cell has the same non-empty safe direct side set and values; assignment atomically whole-replaces or clears all physical cells, including merge continuations. Values quantize to signed-Int32 EMU and may be zero, negative, positive, or fractional points. The snapshot getters require exactly one direct `tcPr`, reads each unique unqualified signed-Int32 integer independently, and returns only valid direct sides. Creation writes `marL/marR/marT/marB`, then optional `anchor`, followed by L/R/T/B borders and fill. PptxGenJS 4.0.1 writes the same narrow defaults and materializes table-level values into cells, but its legacy runtime treats the first value `<1` as inches and `>=1` as points. Thus native table or cell `7.2` equals PptxGenJS `0.1`, native `[3.6, 7.2, 10.8, 14.4]` equals PptxGenJS `[0.05, 0.1, 0.15, 0.2]`, while native `0.1` intentionally means 0.1pt. Adapter imports expose the resulting OOXML in points instead of guessing the original unit.

Table-cell `borders` is a detached partial snapshot of the selected physical cell's same-prefix direct `lnL/lnR/lnT/lnB`. Each supported side is `{ kind: 'none' }` or `{ kind: 'line', color, width, style? }`, where width is finite `0..1584` points, color is strict sRGB/theme, and style is omitted, `solid`, or `dash`. `AddTableOptions.border`, `AddTableCellOptions.border`, `setCellBorders()`, and `TableModel.borders` assignment accept one scalar border, an exact `[top, right, bottom, left]` tuple, or a partial named object. During creation, a non-empty table value is materialized only onto cells whose normalized cell border is absent. Any non-empty cell scalar, TRBL, named value, or explicit none blocks the entire table value; missing sides in that cell value remain canonical none and do not inherit table sides. Empty/all-undefined cell values inherit a supplied table value, while empty/all-undefined table values preserve omitted bytes. Creation overlays the chosen value on four canonical direct no-fill sides; both cell and table editors are whole replacements, so missing sides are removed and `{}` / `undefined` clear all four direct sides without re-inheritance. Explicit none writes a direct zero-width `noFill` line. Public TRBL input is serialized in OOXML's required L/R/T/B order before cell fill; dash maps to `sysDash`, explicit solid maps to `solid`, omitted style writes no direct dash token, zero-width line remains a line, and widths are quantized to the nearest EMU. Border input at every layer is descriptor-safe, getter-free, null-prototype-compatible, and detached. The cell getter omits malformed or unsupported sides independently and never treats diagonals, cell fill, text outlines, advanced dash presets, joins, arrows, or shared-edge/effective style as supported state; unrelated edits preserve them. `TableModel.borders` uses a stricter all-physical-cell consensus: every cell must expose the same non-empty safe complete or partial vector, otherwise absent, mixed, malformed, advanced, repeated, or ambiguous state returns `undefined`. It never resolves table styles, adjacent shared edges, effective borders, defaults, or creation metadata. For supported table/cell scalar or TRBL none, sRGB line, solid/dash, scalar zero width, and whole-value cell overrides, native creation matches PptxGenJS 4.0.1 final direct state. PptxGenJS omitted border imports as the uniform four-side none it materializes; legal uniform values project directly and a differing cell override projects to mixed `undefined`. Native `{}` means clear/no effective layer while PptxGenJS `{}` defaults to `666666`/1pt/solid and a cell `{}` blocks its table value; native omitted style writes no direct dash while PptxGenJS defaults it to direct solid. Native scalar and TRBL zero widths stay zero, whereas PptxGenJS table TRBL `pt: 0` changes to 1pt; PptxGenJS also accepts short runtime tuples and materializes missing sides as none, while native tuples must be dense and exactly four items. Native named sides, theme border colors, and existing-deck bulk editing are strict lossless extensions. Adapter snapshots reflect final XML rather than guessing the original input layer. Diagonal borders, transparency, advanced line properties, shared-edge precedence, and layout recomputation remain unsupported.

Table-cell `fill` is a detached direct-state snapshot with explicit `{ kind: 'none' }` and `{ kind: 'solid', color, transparency? }` variants. Solid fill supports strict six-digit sRGB or the existing theme-color tokens; transparency is a finite `0..100` percentage quantized to `0.001%`. Omitted transparency writes no alpha, while explicit zero writes direct opaque alpha. During native creation, `AddTableOptions.fill` is a whole-value default materialized into every physical cell whose `AddTableCellOptions.fill` is omitted or `undefined`; a cell solid or none value completely overrides it. Table and cell fill/color inputs are descriptor-safe, getter-free, ordinary/null-prototype-only, and deeply detached immediately. Omitted or runtime-undefined table fill preserves the original bytes; `{}` is invalid. The returned live table immediately exposes detached resolved `TableCell.fill` values. `setCellFill()` then addresses only zero-based physical row/cell direct state: `none` writes `tcPr/a:noFill`, while `undefined` removes the direct fill choice and never re-inherits the creation default. The strict cell getter requires a unique direct `tcPr` and one unambiguous same-prefix noFill or solidFill; it never reads border/text descendant fills or resolves table styles. Unsupported gradient/pattern/picture/group fills remain preserved during unrelated edits and can be explicitly replaced or cleared. `TableModel.fill` returns the same detached value only when every physical cell has one identical supported direct fill; absent, mixed, malformed, advanced, or unsafe state is `undefined`. Assignment atomically whole-replaces or clears all physical cells, including merge continuations, with exact no-op and late-cell failure isolation. PptxGenJS 4.0.1 produces the same final direct cell state for supported table solid fills and cell solid overrides, but collapses omitted and `type: 'none'` to no direct fill and explicit zero transparency to no alpha. Native direct none, absence, and explicit zero therefore remain intentionally distinguishable. PptxGenJS may also emit invalid alpha for out-of-range runtime values; adapter imports preserve that XML but do not fabricate a valid snapshot.

Table-cell `textFit` reuses the immutable `TextBoxFit` values `none`, `shrink`, and `resize`, backed only by the selected physical cell's direct `txBody/bodyPr` fit choice. Native plain-text cell creation accepts `AddTableCellOptions.fit` at cell level only: omitted, runtime `undefined`, and `none` produce byte-identical self-closing `bodyPr` and an immediate direct snapshot of `undefined`; `shrink` and `resize` write direct `normAutofit` and `spAutoFit` and snapshot immediately. Creation never writes `noAutofit`, retains table metadata, measures text, calculates a final font scale, or recomputes cell/table geometry. The getter requires exactly one direct text body, one direct body properties element, and one unambiguous supported fit child; it reads an existing `noAutofit` as `none` and otherwise returns `undefined` for absent or malformed state. `setCellTextFit()` uses physical zero-based row/cell positions: `shrink` and `resize` write `normAutofit` and `spAutoFit`, while `none` and `undefined` both clear the direct choice without restoring creation input or creating `noAutofit`. Reassigning the current shrink/resize mode preserves any calculated `fontScale` and `lnSpcReduction`. Fit remains independent from text direction in `tcPr@vert`. Table-level fit creation/defaults are not supported. PptxGenJS 4.0.1 has no table-cell fit API and ignores runtime `fit`, `autoFit`, and `shrinkText` values supplied to table/cell options, so native shrink/resize creation is an explicit extension rather than a parity claim.

## Native charts

`CHART_TYPES` is the frozen public catalog of `area`, `bar`, `bar3D`, `bubble`, `doughnut`, `line`, `pie`, `radar`, and `scatter`. `SlideModel.addChart()` and `PptxDocument.addChart(slideIndex, ...)` expose the same overloads:

```ts
addChart(type: ChartType, series: readonly ChartSeriesInput[], frame?: AddChartOptions): Promise<ChartModel>;
addChart(groups: readonly ChartGroupInput[], frame?: AddChartOptions): Promise<ChartModel>;
```

`ChartSeriesInput` requires a non-empty `name` and finite non-empty `values`. Categorical groups require equal-length `categories`; scatter requires equal-length `xValues`; bubble additionally requires equal-length positive `sizes`. Pie and doughnut accept one series. Compatible bar/area/line groups may use `axis: 'primary' | 'secondary'`; secondary-only, axis-free, 3D, bubble, and incompatible XY/category combinations reject before part or ID allocation. `AddChartOptions` owns frame metadata and EMU/OOXML-angle geometry: `name`, `altText`, `x`, `y`, `width`, `height`, `rotation`, and flips.

Every native chart receives a chart part, an internal package relationship, and a deterministic embedded XLSX. One normalized workbook plan drives worksheet cells, chart formulas, and string/numeric caches. Workbook bytes are generated before the synchronous presentation transaction, so invalid input or workbook work leaves the package unchanged. All six presentation formats use the same creation and reopen path.

`ChartModel.definition` returns a detached, recursively frozen `ChartDefinition` with ordered groups, series, axes, and normalized options. Root options cover language/style, rounded corners, blank handling, title, legend, chart/plot areas, primary and secondary axes, data table, colors, and 3D view. Group options cover data labels, per-series fill/line/marker, and type-specific grouping, direction, gaps, overlap, first-slice angle, hole size, line/scatter/radar markers and smoothing, bubble scale/sign/size representation, and 3D depth. Set root or group options with `replaceDefinition()` after creation; `replaceSeries()` is the shorthand for replacing data on a one-group chart.

`replaceDefinition()` is a synchronized semantic operation: it updates supported chart-owned XML, A1 formulas, caches, and workbook cells together; option-only changes preserve workbook bytes; a proven equal recognized definition is an exact no-op. Existing unsupported extensions, style/color parts, and unowned chart XML remain preserved. Imported shared targets use relationship-aware clone-on-write and keep the selected model identity stable. `remove()` / `slide.deleteChart(shapeId)` removes the frame and relationship and collects only unreferenced owned chart/workbook/style/color parts. `setXml()` remains the explicit raw escape hatch and can intentionally create workbook/cache divergence.

`chart.diagnostics()` returns `CHART_RELATIONSHIP_INVALID`, `CHART_STRUCTURE_UNSUPPORTED`, `CHART_STRUCTURE_AMBIGUOUS`, `CHART_CACHE_INVALID`, `CHART_AXIS_INVALID`, `CHART_WORKBOOK_MISSING`, `CHART_WORKBOOK_CACHE_DIVERGENCE`, or `MODERN_CHART_EXTENSION`. Recognized charts require exact formula/cache/workbook agreement. A standard chart with valid caches but no `externalData` is `cache-only`: opaque client formula placeholders are ignored because there is no workbook to address, diagnostics emit the missing-workbook warning, and the first semantic replacement creates a canonical workbook. Office 2016 `cx:*` charts, external workbooks, and chart animations remain preservation-only.

PptxGenJS 4.0.1 conformance covers all nine public chart types, bar+line primary/secondary combinations, data/formulas/caches/XLSX relationships, and representative title/legend/axis/label/data-table/series/type-specific options. The actual package passes Node, real-Chrome, declaration, and installed-CLI smoke with `nativeCharts: true`; the 11-slide gallery has ten charts/workbooks, zero owned orphans, 0 PowerPoint-2010 errors/warnings, and zero 180-DPI overflow. LibreOffice retains chart types and cached data but removes workbooks on save; these reopen as cache-only warnings and can be canonicalized by semantic editing. Its `bar3D` renderer is title-only for both native and PptxGenJS controls. Built-in trendline/error-bar creation and modern inspection remain in the advanced-charts plugin.

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

## Slide backgrounds, gradients, and transparency

```ts
import type {
  SetSlideBackgroundImageOptions,
  SimpleFill,
  SlideBackground,
  SlideBackgroundImage,
} from '@pptx/sdk';

const slide = document.slides[0];
slide.background = {
  kind: 'linear-gradient',
  angle: 45,
  stops: [
    { offset: 0, color: '#2563EB' },
    { offset: 1, color: '#7C3AED', alpha: 0.65 },
  ],
};

slide.background = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 20,
};
await document.setSlideBackgroundImage(0, './background.png');
slide.background = { kind: 'none' };
slide.background = undefined;
```

`SlideBackground` is `SimpleFill | GradientFill | SlideBackgroundImage`. `SimpleFill` is explicit `{ kind: 'none' }` or a solid sRGB/theme `RichTextColor` with optional `0..100` transparency. Gradients retain sRGB, scRGB, scheme, system, or preset sources and their ordered OOXML transforms. An image contains `kind: 'image'`, `image/png | image/jpeg | image/gif`, and detached `Uint8Array` bytes. Getter results are detached and frozen where applicable; image bytes are copied on each read.

Only the direct `p:sld/p:cSld/p:bg/p:bgPr` state is projected. `undefined` removes direct `p:bg` and restores layout/master inheritance; `{ kind: 'none' }` writes legal direct `a:noFill`. The reader returns `undefined` for unsupported or ambiguous `p:bgRef`, pattern/group fill, wrong namespaces, multiple choices, external/dangling image relationships, or malformed content without modifying the package. Explicit supported replacement canonicalizes an opaque direct background while preserving unrelated slide content. Same-value supported assignment and clearing an absent direct background are exact no-ops.

`setSlideBackgroundImage(slideIndex, source, options?)` accepts `RasterImageSource`. `SetSlideBackgroundImageOptions` contains only optional `contentType` assertion and `AbortSignal`. The resolver accepts Node path, HTTP/HTTPS and browser-relative URL, strict data URI, `Uint8Array`, `ArrayBuffer`, Blob/File, Web stream, and async byte iterable; signature validation and all async I/O finish before mutation. Duplicates initially share an internal image target, a different write clones on first mutation, and replacement/clear/slide deletion remove only relationships and media parts that have no remaining incoming reference.

PptxGenJS 4.0.1 solid, transparency, and PNG background output imports to the same supported state. PptxGenJS `{ type: 'none' }` writes no direct background, so it imports as inherited; `{ type: 'none', color }` emits an empty `p:bgPr`, which the strict reader treats as unsupported. Native explicit none intentionally writes legal `a:noFill` instead of reproducing either behavior. `SlideLayoutModel.background` and `SlideMasterModel.background` use the same supported direct owner-aware state. `p:bgRef` semantic editing, pattern/group fill, and image crop/tile/effects remain outside this API. The separate transient slide default text color is documented below.

Packed Node, real-Chrome, declaration, and installed-CLI smoke report `slideBackgrounds: true`. Two clean builds have an identical 48-file dist manifest (`e42633dfd50e9f8731e780f6b911f691845c530f5c1a0b9e5f356f93a1a0f423`). Full Vitest is 1156 passed with one performance test skipped by default; its separate performance gate is 1/1. The native gallery has 11 slides, 41 parts, 39 relationships, and three background media parts; both native and the seven-slide PptxGenJS control validate 0/0 and render without overflow. LibreOffice preserves slide order and every image payload hash but normalizes explicit no-fill to inheritance and gradient rotation/fill-rectangle metadata. Local PowerPoint automation returned `-9074`, so no PowerPoint round-trip pass is claimed.

## Slide numbers and presentation start number

```ts
import { inches, type SlideNumberOptions } from '@pptx/sdk';

const options: SlideNumberOptions = {
  x: inches(8.1),
  y: inches(5),
  width: inches(1.4),
  height: inches(0.35),
  align: 'justify',
  rtl: true,
  valign: 'middle',
  margin: [1, 2, 3, 4],
  style: {
    fontFamily: 'Aptos',
    fontSize: 18,
    lang: 'zh-CN',
    bold: true,
    italic: true,
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
};

document.firstSlideNumber = 5;
document.slides[0].slideNumber = options;
document.layouts[0].slideNumber = { x: inches(0.5), align: 'left' };
document.masters[0].slideNumber = { x: inches(4.5), align: 'center' };
document.slides[0].slideNumber = undefined;
```

The public value types are `SlideNumber`, `SlideNumberOptions`, `SlideNumberColor`, `SlideNumberMargins`, `SlideNumberMarginInput`, `SlideNumberTextStyle`, and `SlideNumberTextStyleOptions`. Geometry uses integer EMU; margins and font size use points; transparency uses `0..100` percent. The default normalized field is x/y 0, width 800000, height 300000, left aligned, LTR, and `en-US` non-bold/non-italic text. sRGB values are six-digit uppercase hex; scheme colors use supported DrawingML tokens. Inputs are closed, descriptor-safe, fully validated before package access, copied, and deeply frozen on read.

`SlideModel.slideNumber`, `LayoutModel.slideNumber`, and `MasterModel.slideNumber` project only one namespace-correct direct `p:ph type="sldNum"` field in the owner part. Layout and master values do not mutate child slides, and slide assignment does not create hidden fields in the unique master or default layout; the master setter additionally owns direct `p:hf@sldNum`. Equal semantic assignment, absent clear, and same start-number assignment are exact part/relationship/graph/journal no-ops. A unique recognized supported value is patched locally; one unambiguous opaque placeholder may be canonicalized by explicit assignment. Multiple placeholders, duplicate shape IDs, malformed fields, wrong namespaces, extra runs/paragraphs, or unsupported lexical values read as `undefined` and are rejected before unsafe mutation.

`CreatePresentationOptions.firstSlideNumber` and `PresentationModel.firstSlideNumber` accept only signed Int32 safe integers. `undefined` removes direct `p:presentation@firstSlideNum`, whose OOXML default is 1. Slide cached text is `(firstSlideNumber ?? 1) + zeroBasedIndex`; layout/master cached text is `‹#›`. Start changes and slide duplicate/move/delete synchronously update safely recognized direct slide caches inside the lifecycle transaction. The three warning diagnostics are `SLIDE_NUMBER_SHAPE_ID_COLLISION`, `SLIDE_NUMBER_MASTER_DISABLED`, and `SLIDE_NUMBER_CACHE_NONCANONICAL`.

PptxGenJS 4.0.1 public output imports its default/full style, sRGB/scheme color, left/center/right/justify alignment, top/middle/bottom vertical alignment, scalar/four-side margin, font, language, bold, italic, and transparency cases. Native corrects several observed defects instead of imitating them: fixed shape id 25 can collide, zero width/height fall back truthily, RTL and transparency are not faithfully emitted in some cases, layout/master use noncanonical caches, and the generated master placeholder is disabled. Native uses unique IDs, legal positive extents, canonical caches, and an enabled master flag.

The actual 54-file tarball contains 51 `dist` files and passes installed Node, real-Chrome, browser conditional-export, declaration, and CLI checks with `slideNumbers: true`. Two package builds produce an identical 51-file manifest (`3d77e6f56b8f299f2d580112fd0ebe77d0a98c38c07764259a3735064d5f9bea`). The focused suite is 448/448; full Vitest is 1194 passed with one performance test skipped by default, and the separate performance gate is 1/1. The 16-slide native gallery contains 48 parts and 45 relationships and validates 0/0; the 16-slide PptxGenJS control contains 82 parts and 95 relationships and validates with zero errors plus four expected warnings. All 32 pages render at 180 DPI without edge contact and were inspected individually.

LibreOffice 26.8 renders direct slide fields but displays its own 1..16 sequence instead of the native start 5. Its saved package retains slide order/titles, 15 direct slide owners, the cleared direct field, field types, alignment, and principal explicit styling. It removes `firstSlideNum` and rewrites cached text, default fonts/languages, and layout/master placeholder state, producing zero errors and 15 normalization warnings. Prefer direct slide fields when LibreOffice-visible portability matters instead of relying only on layout/master placeholders. Local PowerPoint 16.112 automation started the application but yielded no active presentation or saved/rendered output, so this environment does not establish a PowerPoint round-trip pass.

## Slide default text color

```ts
slide.color = { kind: 'scheme', value: 'accent1' };
slide.addText('Uses accent1');
slide.addRichText([{
  runs: [
    { text: 'Inherited' },
    { text: ' override', style: { color: { kind: 'srgb', value: '00AA00' } } },
  ],
}]);
slide.color = undefined;
```

`SlideModel.color` has type `Readonly<RichTextColor> | undefined`. The setter accepts only the closed sRGB/theme union used by rich text: sRGB is normalized to uppercase six-digit hex without `#`, and scheme values must be supported DrawingML tokens. Inputs are descriptor-safe and detached; getters are frozen snapshots. Equal assignment is an exact package/journal no-op. `undefined` clears the transient state.

The value affects only subsequently created `addText()` / `addRichText()` runs. A run's explicit `style.color` takes precedence, while a run with transparency but no local color inherits the slide default. Changing or clearing `slide.color` does not scan or recolor existing shapes or tables. Tables, masters, layouts, and placeholders do not inherit this transient state; named master/layout/placeholder population is complete, while full theme text cascade and advanced-table color behavior remain separate advanced-text/table work.

The transient default is not serialized because OOXML has no legal direct slide-level default-text-color field. Creation materializes the chosen color into each run's standard `a:solidFill`. Reopened documents therefore have `slide.color === undefined`, while materialized run colors remain readable and visually unchanged. Duplicate copies the current state, move retains it, delete removes it, URI reuse cannot leak it, and all operations preserve sibling isolation and transaction rollback.

All six presentation formats, valid PptxGenJS 4.0.1 sRGB/theme/override output, Node, browser, declarations, and CLI are covered. The focused run is 10 passed / 409 skipped; full Vitest is 1205 passed / 1 skipped, performance is 1/1 at 998ms, and typecheck/build pass. The actual tarball contains 54 files and 51 `dist` files; its dist manifest SHA-256 is `467d87ffea6994355c357dbad3b1ea18afa8538b1bacb85b6de43de90ad16829` and tarball SHA-256 is `6812000a83247fdf2d63eddf81ec6ffb43c721d478e4cdcbbf4c4a3ce2b65ad1`.

The real-Chrome result exactly matches live inherited/override/transparency/duplicate state, reopened defaults are absent, and console/page/network errors are zero. Native and PptxGenJS galleries contain 11/9 slides, 38/52 parts, and 35/58 relationships; both validate with 0 errors / 0 warnings under the PowerPoint 2010 profile. All 20 pages were inspected at 180 DPI with zero overflow and 106px minimum margins. LibreOffice 26.8 retains slide and text order, custom sRGB, theme, override, and 40% transparency, while normalizing native `tx1` to equivalent `dk1`; the saved file remains 0/0. Local PowerPoint 16.112 returned `-9074` for both native and control inputs without loading or producing PPTX/PDF output, so no PowerPoint round-trip pass is claimed.

## Master, layout, placeholder, and theme

```ts
import { inches, PLACEHOLDER_TYPES, PptxDocument } from '@pptx/sdk';

const document = PptxDocument.create({ slideSize: 'wide' });
const layout = await document.defineSlideMaster({
  title: 'BRAND',
  background: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
  },
  margin: [inches(0.5), inches(0.5), inches(0.5), inches(0.5)],
  slideNumber: { x: inches(11), width: inches(1) },
  objects: [
    { kind: 'rect', options: { x: 0, y: 0, width: inches(13.333), height: inches(0.2) } },
    {
      kind: 'placeholder',
      text: 'Presentation title',
      options: {
        name: 'title_box',
        type: 'title',
        index: 101,
        x: inches(1),
        y: inches(1),
        width: inches(8),
        height: inches(1),
      },
    },
  ],
});

const slide = document.addSlide({ masterName: layout.name });
slide.addText('Quarterly results', { placeholder: 'title_box' });
console.log(PLACEHOLDER_TYPES); // title, body, pic, chart, tbl, media
```

`PptxDocument.defineSlideMaster(options)` is asynchronous because image, image-background, and chart definition objects may require source/workbook preparation. `DefineSlideMasterOptions.title` is a unique presentation-wide layout name. Optional `master` selects an attached same-document `SlideMasterModel`; omission uses the first safe attached master. Optional background accepts every supported direct background plus `{ kind: 'image-source', source, contentType?, signal? }`. `margin` is a non-negative EMU scalar or exact `[top, right, bottom, left]` tuple bounded by the current slide size. `slideNumber` uses the ordinary direct layout field options.

`SlideMasterObject` is the closed ordered union `rect | line | text | placeholder | image | chart`. Rect and line use `AddShapeOptions`; text accepts a string or rich-text paragraphs; image accepts every portable `ImageSource`; chart accepts normalized chart groups. Placeholder options require an XML-safe unique name and one of the six frozen `PLACEHOLDER_TYPES`, with an optional unique integer index from 0 through 4,294,967,294. All asynchronous work completes before one synchronous package transaction, so source, chart, relationship, part, XML, or definition failure has no observable partial mutation.

PptxGenJS calls the definition a master, but native `defineSlideMaster()` deliberately creates a named layout under a real parent master. Correspondingly, `addSlide({ masterName })` strictly resolves exactly one attached `SlideLayoutModel.name`. Unknown/duplicate names fail instead of falling back. Ordinary layout objects stay on the layout and are inherited. Layout placeholder prompts stay on the layout, while slide creation materializes empty owners with the same unique `{ type, index }` and geometry. A selector is either the unique layout placeholder name or a `PlaceholderIdentity`. Text/rich-text and shapes fill title/body owners, image/SVG fills `pic`, charts fill `chart`, tables fill `tbl`, and audio/video fills `media`; domain mismatch, ambiguous ownership, or a second fill fails before mutation.

`PptxDocument.masters` and `.layouts` return stable live semantic wrappers. `SlideMasterModel` exposes `partUri`, `layouts`, `theme`, `background`, `shapes`, `placeholders`, and `slideNumber`. `SlideLayoutModel` exposes `partUri`, `name`, `masterPartUri`, runtime `margin`, `background`, `shapes`, `placeholders`, and `slideNumber`. Both wrappers expose `addPlaceholder()`, `addText()`, `addRichText()`, `addShape()`, `addImage()`, `addSvgImage()`, and `addChart()`. Existing content and supported background/placeholder relationships are therefore editable without dropping the raw codec surface.

`replaceSlideMaster(layout, options)` prepares a complete definition and atomically replaces only layout-owned content, background, relationships, slide number, and transient margin. It retains the target part URI, `SlideLayoutModel` identity, master layout ID, and every incoming slide relationship; a changed `master` safely relinks both sides. An equivalent recognized definition is an exact no-op. `deleteSlideMaster(layout, replacement?)` rejects a used layout without a replacement. With an attached same-document replacement, it retargets incoming slides before deleting the layout and collecting only unreferenced owned dependencies. Deleted handles become stale, and later URI reuse does not revive them.

Layout `margin` is intentionally transient because OOXML and PptxGenJS do not serialize this table auto-page / `tableToSlides` hint. It is frozen during the current document session, supplies default margins for auto-page and DOM table slides using that named layout, updates on whole replacement, is removed on delete, and reopens as `undefined`. Background, ordinary content, placeholder identity/geometry, slide numbers, relationships, and payloads are persistent.

The theme API remains available alongside semantic wrappers:

```ts
const theme = document.themes[0];
theme.setColor('accent1', '#2563EB');

const chain = document.masterLayoutTheme.materializeInheritedStyle(
  document.slides[0].partUri,
  document.slides[0].shapes[0].id,
);
```

`masterLayoutTheme` continues to expose raw create/copy/delete/relink operations. Semantic wrapper collections synchronize with those raw operations and retain stable identity while their parts remain attached.

Focused master/layout/placeholder tests report 45 passed / 434 skipped; full Vitest reports 1256 passed / 1 skipped, performance is 1/1 at 578ms, and typecheck/build pass. The installed 57-file tarball has 54 `dist` files. Two clean builds have byte-identical sorted dist-hash manifests and tarballs, with SHA-256 `0a8e958ccde379ae071a7388dc4c29278ac5033a8641976324fcd5820339ad27` and `8362a3af38a4a7e8316a7e49e8cb3f4fb405753bd20cc935db609441819ca5e8`. Packed Node, TypeScript, CLI, and real Chrome all report `masterLayouts: true`; Chrome restores all six placeholder domains, selected layout targets, master/layout backgrounds, payload hashes, and chart definitions with zero validation, console, page, or network errors.

The two-slide native gallery has 32 parts, 29 relationships, two layouts, and one master and validates 0/0 under the PowerPoint 2010 profile. The two-slide PptxGenJS control has 36 parts and 34 relationships. Eight source/LibreOffice-round-trip pages were rendered at 2400×1350 and 180 DPI and inspected individually; full-bleed fixture backgrounds give an expected 0px minimum non-white margin. LibreOffice 26.8 preserves two slides, two layouts, and one master but rewrites placeholder identities and slide-number caches and removes audio plus embedded chart workbooks. Local PowerPoint 16.112 returned `-9074` for both native and control inputs and produced no PPTX/PDF, so neither result is reported as a full round-trip pass.

Full theme text cascade, percentage coordinates, advanced text/table/media/chart styles, and broad client certification remain pending. Advanced text now includes text-shape direct fill, simple line, arrows, simple shadow, outer hyperlink, per-run rich-text hyperlink, preset geometry, rounded-rectangle-radius, direct `isTextBox` creation/read/edit, and rich-text `breakLine` paragraph splitting.

Remaining advanced-text items are followed by advanced table/`tableToSlides`, output/runtime helpers, and the peer-range full-suite audit. Full PptxGenJS parity is not yet claimed.

## Media

```ts
import { inches } from '@pptx/sdk';

const poster =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAACXBIWXMAAAABAAAAAQBPJcTWAAAAEElEQVR4nGP8ywACLGCSAQANEQED1LYyQAAAAABJRU5ErkJggg==';

const audio = await document.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
  name: 'Opening narration',
  altText: 'Opening narration audio',
  poster,
  x: inches(1),
  y: inches(1),
  width: inches(3),
  height: inches(2),
  contentType: 'audio/mpeg',
  play: 'auto',
  loop: true,
  hideWhenStopped: true,
  volume: 0.5,
});

declare const videoBytes: Uint8Array;
await document.addVideo(0, videoBytes, {
  contentType: 'video/mp4',
  fileName: 'overview.mp4',
  poster,
});

audio.name = 'Opening narration edited';
audio.altText = undefined;
audio.settings = { play: 'click', loop: false, volume: 0.75 };
audio.setTransform({ x: inches(2), y: inches(1.5) });
await audio.replaceSource('https://example.com/narration.wav');
await audio.replaceSource('narration.wav', {
  contentType: 'audio/wav',
  fileName: 'narration.wav',
});
await audio.replacePoster('poster.gif', {
  contentType: 'image/gif',
});
await audio.replacePoster(); // reset to the built-in PNG
```

`MediaSource` is a Node path, strict base64 data URI, `Uint8Array`, `ArrayBuffer`, Blob/File, Web `ReadableStream<MediaByteChunk>`, or async iterable of byte chunks. A string beginning with HTTP/HTTPS is an external media relationship and is never fetched; other URL schemes are rejected. Posters use the same embedded-source union except that HTTP/HTTPS poster URLs are rejected. Omitted posters use a built-in one-pixel PNG.

`AddMediaOptions` contains optional `name`, `altText`, `contentType`, `fileName`, `poster`, `posterContentType`, EMU `x`/`y`/`width`/`height`, `play: 'click' | 'auto'`, `loop`, `hideWhenStopped`, `volume: 0..1`, and an async `transcode(bytes, contentType, kind)` hook. Audio supports `audio/mpeg` (`.mp3`), `audio/mp4` (`.m4a`), `audio/wav` (`.wav`), and `audio/ogg` (`.ogg`). Video supports `video/mp4` (`.mp4`/`.m4v`), `video/quicktime` (`.mov`), and `video/webm` (`.webm`). Posters support `image/png` (`.png`), `image/jpeg` (`.jpg`/`.jpeg`), and `image/gif` (`.gif`).

`addAudio()` and `addVideo()` return a runtime `MediaModel`. Within one document, the result is identical by reference to the corresponding member of `document.media(slideIndex)`, `slide.media`, and `slide.shapes`; collection reads and slide moves preserve that identity. Its live getters are `kind`, `shapeId`/`id`, `slidePartUri`, `name`, `altText`, `settings`, `mediaPartUri`, `externalUrl`, `posterPartUri`, and the inherited `transform`. Set `name`, `altText`, `settings`, or transform directly. A setting assignment synchronizes exact `px:playback` preference/ownership state with native PowerPoint timing. `settings = undefined` removes both the extension and only the library-owned native media graph. Returned settings and transform snapshots are detached and frozen.

`MediaModel.replaceSource(source, options?)` preserves kind and returns the same model. `ReplaceMediaSourceOptions` accepts only `contentType`, `fileName`, and `transcode`; all creation source forms are supported, including embedded↔external transitions. `replacePoster(source?, options?)` accepts only `contentType` and `fileName`; PNG/JPEG/GIF sources are supported, and omitted source resets the built-in PNG. `remove()` delegates to `slide.deleteMedia(shapeId)`. A removed or semantically changed picture invalidates the old live handle through the normal shape-resolution error path.

Descriptor resolution is explicit MIME assertion → data-URI MIME → known extension from `fileName`, path, URL path, or `File.name` → canonical domain default (`audio/mpeg`, `video/mp4`, or `image/png`). Explicit and declared MIME must match. A recognized extension must belong to the correct domain and match the resolved MIME; `.m4v` and `.jpeg` are preserved when explicitly inferred, otherwise the first canonical extension is used. Unknown extensions are ignored as evidence. `Blob.type` is not used. A transcode result must be an ordinary `{ bytes: Uint8Array, contentType, extension? }` object, and any explicit extension must be lowercase and exactly compatible with its output MIME.

Media data URIs require the exact `data:<supported-mime>;base64,<payload>` form. Payloads must use the standard alphabet, complete padding, and canonical padding bits; empty data, whitespace, percent encoding, URL-safe characters, extra commas, and malformed or noncanonical padding are rejected.

Request objects are descriptor-safe, getter-free, and reject unknown properties. Options, byte arrays, ArrayBuffers, transcode inputs/results, and final embedded payloads are detached from the caller. All asynchronous path/Blob/stream I/O, transcode work, poster loading, descriptor resolution, SHA-256 lookup, and XML-definition work completes before mutation. Part, content-type, relationship, and slide XML writes then run inside one synchronous package transaction; any create, replacement, or deletion failure leaves the package graph, ZIP state, model identity, and mutation journal unchanged.

Embedded payloads deduplicate only when SHA-256 and exact MIME both match. Every embedded media picture uses a canonical `a:audioFile` or `a:videoFile`, a standard kind relationship, a Microsoft media relationship, an internal poster image relationship, a media click action, and a rectangular `p:pic`. Duplicated slides initially share media and poster targets. Source/poster replacement updates an exclusive compatible target in place or uses content-aware deduplication/clone-on-write, allocating relationships when an rId is shared by multiple XML nodes and retargeting only the edited picture. Superseded relationships are removed only after their XML reference count reaches zero; media/poster parts are collected only after package-graph incoming reaches zero. Object and slide deletion follow the same reference-aware GC rule.

Playback is native without a plugin. `px:playback` retains exact preferences and the owned media/play/pause timing IDs, while the slide's `p:timing` contains the executable `p:audio`/`p:video`, `p:cMediaNode`, play command, and click pause command where applicable. `play: 'click'` emits an interactive click sequence; `play: 'auto'` emits a with-previous sequence. `loop` maps to indefinite repeat, `hideWhenStopped` maps to `showWhenStopped`, and `volume` maps to OOXML's 0–100000 volume.

Effective read order is a valid `px:playback` preference, then—only when no preference exists—a recognized unique/direct/complete native media graph, then empty settings. A native-only graph can therefore be read and adopted without rewriting its timing bytes when the first assignment is semantically identical. Recorded ownership is repaired when IDs are stale. The codec never claims ordinary animations, nested/non-direct media graphs, multiple matches, finite repeats, cross-slide media, trim/bookmarks, or otherwise unsupported branches.

Create, set, clear, materialize, duplicate, delete, and timing ID allocation are transactional. Timing IDs are unique across media and ordinary animations on the whole slide. Diagnostics are warnings with codes `MEDIA_TIMING_MISSING`, `MEDIA_TIMING_STALE`, `MEDIA_TIMING_UNSUPPORTED`, `MEDIA_TIMING_AMBIGUOUS`, `MEDIA_TIMING_DANGLING_TARGET`, and `MEDIA_TIMING_KIND_MISMATCH`. External media still emits portability diagnostics; under the PowerPoint 2010 profile, `audio/ogg` and `video/webm` emit expected codec warnings.

PptxGenJS 4.0.1 valid public embedded-media cases are semantically covered, including data/path, audio/video, cover, extension, object name, transform, and repeated-path deduplication. Import accepts its audio `a:videoFile`, `audio/mp3`, and duplicate-audio Microsoft-media relationship defects. Reads and non-source edits preserve those legacy primary roles; `replaceSource()` canonicalizes only the selected picture. Native creation always uses `a:audioFile`, canonical `audio/mpeg`, and the standard audio relationship.

The actual 45-file tarball passes Node/real-Chrome/declaration/installed-CLI smoke with `nativeMediaTiming: true`; two clean builds have identical SHA-256 manifests for all 42 dist files. All six presentation formats pass native timing create/edit/duplicate/delete/reopen. The nine-slide playable gallery contains 12 media objects across MP3/M4A/WAV/OGG and MP4/MOV/WebM, all three poster MIME types, ten deduplicated media/poster parts, and zero orphans. It strictly reopens, renders at 180 DPI without overflow, passes slide-by-slide visual inspection, and validates against PowerPoint 2010 with 0 errors plus only the expected OGG/WebM warnings. LibreOffice 26.8 preserves nine slides and text but removes all media, posters, media relationships, and timing on save; its output still strictly reopens and validates 0/0. Local PowerPoint 16.112 automation returned `-9074` for this gallery and both independent control files, so it is not reported as a successful round trip.

Trim/bookmarks, finite repeats, narration/cross-slide audio, captions/subtitles, online video, remote-fetch embedding, media crop/rounding/shadow/hyperlink and advanced placeholder styles, a built-in transcoding engine, and broad client certification remain pending. Native standard-chart creation, semantic editing, direct slide backgrounds, slide numbers, slide default text color, and master/layout/placeholder support are complete; the next PptxGenJS parity slice is advanced text.

## HTML table to editable slides

`PptxDocument.tableToSlides(elementId, options?)` requires a browser document, resolves the target only through `getElementById()`, requires a `<table>`, and returns `Promise<readonly SlideModel[]>`. The returned frozen array contains only HTML-generated pages; every page is an ordinary editable slide.

```ts
const pages = await document.tableToSlides('report-table', {
  masterSlideName: 'REPORT',
  autoPage: true,
  autoPageRepeatHeader: true,
  x: inches(0.5),
  y: inches(0.75),
  width: inches(12),
  slideMargin: inches(0.5),
  addImage: { source: logoBytes, options: { width: inches(1), height: inches(0.4) } },
  addShape: { type: 'rect', options: { width: inches(0.2), height: inches(0.2) } },
  addTable: { rows: [['Generated', 'Yes']] },
  addText: { text: [{ runs: [{ text: 'Confidential', style: { bold: true } }] }] },
});
```

Options expose strict `name`, `masterSlideName`, `autoPage`, measurement weights, repeat-header controls, continuation Y, margin, EMU x/y/width/height/column widths, and one image/shape/table/text template. `autoPage` defaults to true. Explicit false produces exactly one automatic-row table and rejects pagination-only options. Additional objects are committed on each HTML page in image → shape → table → text order; nested pages created by `addTable` are excluded from the return snapshot.

Snapshot order is `thead`, every `tbody`, then `tfoot`. Cell `innerText` and computed CSS are detached once. Supported CSS is sRGB foreground/background, font family/size/weight, alignment, vertical alignment, padding, and four borders. Transparent background becomes white; visible non-solid borders become dash. Explicit native `columnWidths` wins; otherwise visible pixel proportions plus header `data-pptx-width` and `data-pptx-min-width` constraints resolve an exact positive EMU vector. The two data attributes use inches at the HTML boundary.

Images resolve before mutation and only once. Each page owns its relationship; exact content-type-and-byte payloads share media parts, including SVG/fallback pairs, while replacement remains clone-on-write. DOM/CSS/layout/column/image failures occur before package writes. Slide creation, pagination, relationships, and all additions share one outer transaction and roll back together.

The API does not retain DOM/CSS objects, parse word-level HTML, or map alpha into table fill/borders. A hidden table without sufficient fixed-width constraints rejects. Native deliberately avoids PptxGenJS 4.0.1 selector interpolation, caller mutation, ignored `autoPage: false`, truthy coercion, fixed-width-as-minimum behavior, and silent invalid numbers. Public capability coverage is 100%; full-parity certification remains gated on the final actual-package/browser/client/peer audit.

## Diagnostics and errors

Errors: `PackageError`, `ParseError`, `ValidationError`, `OpaqueMutationError`, `PptxGenJSAdapterError`.

Every diagnostic has severity, code, message, and optional part URI, XML path, object id, compatibility profile, and suggestion. Strict mode blocks only error diagnostics.
