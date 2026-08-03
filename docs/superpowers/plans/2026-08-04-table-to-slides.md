# HTML Table to Slides Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a strict browser-DOM `PptxDocument.tableToSlides()` that converts styled HTML tables into editable, auto-paged PowerPoint slides with named layouts and per-page image/shape/table/text additions.

**Architecture:** Snapshot DOM rows and computed cell styles into immutable native table inputs, solve exact EMU columns from visible pixel proportions plus HTML width constraints, then reuse `SlideModel.addTable()` for content measurement, fragmentation, repeated headers, layouts, sections, relationships, and rollback. Resolve optional image sources before one outer package transaction; add all other per-page objects through existing native creation APIs.

**Tech Stack:** TypeScript 5 strict mode, browser DOM/CSSOM, immutable native table inputs, safe-integer EMU math, Vitest, lossless OOXML/OPC transactions, PptxGenJS 4.0.1 source and Chrome conformance, pnpm/tsup, actual npm tarballs, installed Node/NodeNext/browser/CLI/Inspector, Google Chrome, LibreOffice rendering, and PowerPoint 2010 validation.

## Global Constraints

- Public native geometry is EMU. Computed CSS pixels map to points only for font, padding, and border width; `data-pptx-width` and `data-pptx-min-width` remain PptxGenJS-compatible inches at the HTML boundary.
- `tableToSlides()` accepts a non-empty string ID, requires a browser document and a `<table>`, returns `Promise<readonly SlideModel[]>`, and never interpolates the ID into a CSS selector.
- `autoPage` defaults to true. Explicit false creates exactly one ordinary automatic-row table and rejects auto-page-only controls.
- DOM row order is `thead`, every `tbody`, then `tfoot`; all cell values come from one detached `innerText` and computed-style snapshot.
- Options and additions are descriptor-safe ordinary/null-prototype data records. No caller object, nested template, DOM node, or CSS object is mutated.
- Explicit `columnWidths` wins. Otherwise column pixels plus fixed/minimum HTML constraints produce positive safe EMU widths whose exact sum equals table width.
- Reuse existing strict table measurement, fragmentation, repeated headers, placeholder/layout state, same-layout insertion, sections, slide-number caches, relationships, and write/reopen behavior.
- Resolve DOM/CSS/layout/column/image preflight before package mutation; commit first slide, HTML table pages, and additional objects in one outer OPC transaction.
- Additional objects are added to each HTML-generated page in image → shape → table → text order. Nested additional-table pages are not included in the HTML result snapshot.
- Preserve PptxGenJS legal capabilities, but do not copy caller mutation, ignored false auto-page, truthy/coercible inputs, selector bugs, mutable line queues, silent NaN, or fixed/min width defects.
- Each task ends with focused review, commit, push, and `HEAD...origin/main` divergence `0 0`; never stage `.pnpm-store/` or retained `/tmp` evidence.

---

### Task 1: Strict Request Normalization and Detached DOM Rows

**Files:**
- Create: `packages/sdk/src/table-to-slides.ts`
- Create: `packages/sdk/src/table-to-slides.test.ts`

**Interfaces:**
- Consumes: unknown `elementId`, unknown options, global DOM-like document, table sections/cells, and cell `innerText` / span properties.
- Produces:

```ts
export interface TableToSlidesAddImage {
  readonly source: ImageSource;
  readonly options?: AddImageSourceOptions;
}

export interface TableToSlidesAddShape {
  readonly type: PresetShapeType;
  readonly options?: AddShapeOptions;
}

export interface TableToSlidesAddTable {
  readonly rows: readonly (readonly AddTableCellInput[])[];
  readonly options?: AddTableOptions;
}

export interface TableToSlidesAddText {
  readonly text: string | readonly RichTextParagraph[];
  readonly options?: AddTextOptions;
}

export interface TableToSlidesOptions {
  readonly name?: string;
  readonly masterSlideName?: string;
  readonly autoPage?: boolean;
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
  readonly autoPageRepeatHeader?: boolean;
  readonly autoPageHeaderRows?: number;
  readonly autoPageSlideStartY?: number;
  readonly slideMargin?: TableAutoPageMarginInput;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly addImage?: TableToSlidesAddImage;
  readonly addShape?: TableToSlidesAddShape;
  readonly addTable?: TableToSlidesAddTable;
  readonly addText?: TableToSlidesAddText;
}

export interface HtmlTableSnapshot {
  readonly rows: readonly (readonly HtmlTableCellSnapshot[])[];
  readonly headRowCount: number;
  readonly widthSourceRowIndex: number;
}
```

- [ ] **Step 1: Write failing normalization and DOM-order tests**

Create a DOM-like fixture builder in `table-to-slides.test.ts` whose cells expose `localName`, `innerText`, `offsetWidth`, `colSpan`, `rowSpan`, `getAttribute()`, and an owner-document style resolver. Add this exact core case:

```ts
it('snapshots thead, multiple tbody sections, and tfoot without retaining DOM', () => {
  const dom = tableFixture({
    head: [[cell('Head', { colSpan: 2, width: 200 })]],
    bodies: [
      [[cell('A\r\nB', { width: 80 }), cell('C', { width: 120 })]],
      [[cell('D', { rowSpan: 2, width: 80 }), cell('E', { width: 120 })]],
    ],
    foot: [[cell('Foot A', { width: 80 }), cell('Foot B', { width: 120 })]],
  });
  const snapshot = snapshotHtmlTable(dom.table, dom.getComputedStyle);
  expect(snapshot.headRowCount).toBe(1);
  expect(snapshot.rows.map((row) => row.map(({ text }) => text))).toEqual([
    ['Head'], ['A\nB', 'C'], ['D', 'E'], ['Foot A', 'Foot B'],
  ]);
  expect(snapshot.rows[0]![0]).toMatchObject({ colspan: 2, offsetWidth: 200 });
  expect(snapshot.rows[2]![0]).toMatchObject({ rowspan: 2 });
  dom.cells[0]!.innerText = 'mutated';
  expect(snapshot.rows[0]![0]!.text).toBe('Head');
  expect(Object.isFrozen(snapshot.rows)).toBe(true);
});
```

Add missing DOM/defaultView/getComputedStyle, empty/string/class target, non-table target, empty table, empty row, special-character ID, multiple body order, getter/unknown/symbol/inherited options, sparse tuples, nested add records, invalid booleans/weights/coordinates, caller mutation, and accessor-call-count-zero cases.

- [ ] **Step 2: Run the focused test and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts \
  -t "snapshots|normalizes|rejects"
```

Expected: FAIL because the module and public request types do not exist.

- [ ] **Step 3: Implement descriptor-safe request and row snapshot normalization**

Use a fixed allowed-key set and own data descriptors:

```ts
const OPTION_KEYS = new Set([
  'name', 'masterSlideName', 'autoPage', 'autoPageCharWeight',
  'autoPageLineWeight', 'autoPageRepeatHeader', 'autoPageHeaderRows',
  'autoPageSlideStartY', 'slideMargin', 'x', 'y', 'width', 'height',
  'columnWidths', 'addImage', 'addShape', 'addTable', 'addText',
]);

export function normalizeTableToSlidesRequest(
  elementId: unknown,
  options: unknown,
): Readonly<NormalizedTableToSlidesRequest> {
  const id = normalizeNonEmptyString(elementId, 'HTML table element ID');
  const input = readDataObject(options, OPTION_KEYS, 'HTML table slide options');
  const autoPage = normalizeOptionalBoolean(input.autoPage, 'HTML table autoPage') ?? true;
  return Object.freeze({ id, autoPage, ...normalizeKnownFields(input, autoPage) });
}
```

Resolve the target only through `document.getElementById(id)`. Structural-check `localName === 'table'`, then snapshot `tHead.rows`, every `tBodies`, and `tFoot.rows`. Read each platform value once; normalize CRLF/CR; store span only when greater than one; freeze cells, rows, and outer snapshot. Keep CSS mapping empty until Task 2.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/sdk/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/sdk/src/table-to-slides.ts packages/sdk/src/table-to-slides.test.ts
git commit -m "feat: normalize HTML table slide requests"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review zero mutation, own-data reads, platform-read count, row order, span omission, snapshot freezing, and error context. Expected divergence: `0 0`.

### Task 2: Computed CSS to Native Cell Styles

**Files:**
- Modify: `packages/sdk/src/table-to-slides.ts`
- Modify: `packages/sdk/src/table-to-slides.test.ts`

**Interfaces:**
- Consumes: one computed-style object per HTML cell.
- Produces: `HtmlTableCellSnapshot.options: Readonly<AddTableCellOptions>` with native sRGB/font/margin/border/alignment state.

- [ ] **Step 1: Write failing exact CSS mapping tests**

Add a cell fixture with:

```ts
style: {
  color: 'rgb(1, 2, 3)',
  'background-color': 'rgb(240, 241, 242)',
  'font-family': '"Noto Sans", Arial, sans-serif',
  'font-size': '18.5px',
  'font-weight': '600',
  'text-align': 'right',
  'vertical-align': 'bottom',
  direction: 'ltr',
  'padding-top': '7.5px',
  'padding-right': '11px',
  'padding-bottom': '3.25px',
  'padding-left': '5px',
  'border-top-style': 'solid',
  'border-top-width': '2px',
  'border-top-color': 'rgba(10, 20, 30, 0.5)',
  'border-right-style': 'dashed',
  'border-right-width': '1.5px',
  'border-right-color': 'rgb(40, 50, 60)',
  'border-bottom-style': 'none',
  'border-bottom-width': '0px',
  'border-bottom-color': 'rgb(0, 0, 0)',
  'border-left-style': 'dotted',
  'border-left-width': '3px',
  'border-left-color': 'rgb(70, 80, 90)',
}
```

Require:

```ts
expect(snapshot.rows[0]![0]!.options).toEqual({
  align: 'right',
  bold: true,
  color: { kind: 'srgb', value: '010203' },
  fill: { kind: 'solid', color: { kind: 'srgb', value: 'F0F1F2' } },
  fontFamily: 'Noto Sans',
  fontSize: 18.5,
  margin: [7.5, 11, 3.25, 5],
  valign: 'bottom',
  border: {
    top: { kind: 'line', color: { kind: 'srgb', value: '0A141E' }, width: 2, style: 'solid' },
    right: { kind: 'line', color: { kind: 'srgb', value: '28323C' }, width: 1.5, style: 'dash' },
    bottom: { kind: 'none' },
    left: { kind: 'line', color: { kind: 'srgb', value: '46505A' }, width: 3, style: 'dash' },
  },
});
```

Add transparent background → white, normal/400 → false, bold/bolder/500, justify, RTL start/end, top/middle, decimal RGB channels, rgba alpha ignored, all-empty optional CSS, nested visible text, unsupported color space, invalid channel/size/padding/border width, and one-getComputedStyle-call-per-cell cases.

- [ ] **Step 2: Run the CSS tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts \
  -t "computed CSS|transparent|border|alignment"
```

Expected: FAIL because snapshots contain no native style options.

- [ ] **Step 3: Implement deterministic CSSOM parsing**

Read each property through one cached accessor map. Implement `parseCssRgb()`, `parseCssLengthPx()`, `normalizeFontFamily()`, `normalizeCssAlignment()`, and four-side border mapping. Use round-to-byte for decimal RGB channels, preserve decimal points for font/padding/border before downstream OOXML quantization, map transparent background to white, and map visible non-solid border styles to `dash`.

Do not retain CSSStyleDeclaration or DOM references. Omit properties only when CSSOM returns an intentional empty/generic value; throw a cell-indexed error for malformed non-empty platform values.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/sdk/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/sdk/src/table-to-slides.ts packages/sdk/src/table-to-slides.test.ts
git commit -m "feat: map HTML table cell styles"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review PptxGenJS common-state equality, transparent handling, sRGB uppercase, start/end direction, no retained CSS objects, and exact point values.

### Task 3: Exact Column Width Solver and HTML Constraints

**Files:**
- Modify: `packages/sdk/src/table-to-slides.ts`
- Modify: `packages/sdk/src/table-to-slides.test.ts`

**Interfaces:**
- Consumes: snapshot cell pixel widths/spans, first header attributes, target table width, and optional explicit native widths.
- Produces:

```ts
export interface ResolvedHtmlTableColumns {
  readonly widths: readonly number[];
  readonly width: number;
}

export function resolveHtmlTableColumns(
  snapshot: Readonly<HtmlTableSnapshot>,
  targetWidth: number,
  explicit?: number | readonly number[],
): Readonly<ResolvedHtmlTableColumns>;
```

- [ ] **Step 1: Write failing width and constraint tests**

Cover a 900px 25%/75% row against `inches(10.8)` and require `[inches(2.7), inches(8.1)]`. Add odd EMU remainders and require `sum(widths) === width` exactly. Add colspan 2 with 300px plus 100px third column and require the span split evenly before proportioning.

Add first-header constraints:

```ts
expect(resolveHtmlTableColumns(snapshotWithConstraints([
  { pixels: 100, minInches: 0.8 },
  { pixels: 200, fixedInches: 4.5 },
  { pixels: 100 },
]), inches(8))).toEqual({
  widths: [inches(1.75), inches(4.5), inches(1.75)],
  width: inches(8),
});
```

Add fixed+minimum overflow expansion, fixed wins over min, colspan constraint distribution, explicit scalar/vector override, explicit width mismatch, all-zero hidden table, all-fixed hidden table, empty/negative/NaN/infinite attributes, too many physical columns, sub-EMU allocation, and safe-sum overflow.

- [ ] **Step 2: Run width tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts \
  -t "column width|data-pptx|hidden table"
```

Expected: FAIL because the width solver does not exist.

- [ ] **Step 3: Implement largest-remainder and water-filling resolution**

Expand a chosen first non-empty row to physical pixel weights. Allocate target EMU by floor plus descending fractional remainder with stable column-index ties. Parse header constraints as positive finite inches and convert with `inches()`.

Apply fixed columns first. For flexible columns, start at minimum widths and distribute remaining EMU proportionally with another largest-remainder pass. If fixed+minimum+one-EMU flexible floors exceed target, set actual width to that required sum. Freeze output and assert every width is a positive safe integer and exact safe sum.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/sdk/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/sdk/src/table-to-slides.ts packages/sdk/src/table-to-slides.test.ts
git commit -m "feat: resolve HTML table column widths"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review fixed/min semantics, exact sums, deterministic ties, hidden-source error, and mismatch delegation to strict table creation.

### Task 4: Public DOM Table Pagination and Lifecycle

**Files:**
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/sdk/src/table-to-slides.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/table-to-slides.test.ts`

**Interfaces:**
- Consumes: normalized request, snapshot, slide size, selected runtime layout margin, and resolved columns.
- Produces: `PptxDocument.tableToSlides(elementId, options): Promise<readonly SlideModel[]>`.

- [ ] **Step 1: Write failing public lifecycle tests**

Install a fake global document inside each test and restore its exact property descriptor in `finally`. Build a 40-row table with two header rows, rich-enough text, fixed viewport widths, and 12pt CSS. Require:

```ts
const slides = await deck.tableToSlides('sales:2026 table', {
  autoPageRepeatHeader: true,
  autoPageCharWeight: 0,
  autoPageLineWeight: 0,
  x: inches(0.75),
  y: inches(1),
  autoPageSlideStartY: inches(0.5),
  slideMargin: [inches(0.5), inches(0.4), inches(0.6), inches(0.4)],
});
expect(slides.length).toBeGreaterThan(1);
expect(Object.isFrozen(slides)).toBe(true);
expect(slides).toEqual(deck.slides);
expect(slides.every((slide) => slide.shapes.some(({ kind }) => kind === 'table'))).toBe(true);
expect(slides.slice(1).every((slide) =>
  slide.shapes.find(({ kind }) => kind === 'table')!.rows[0]!.cells[0]!.text === 'Header A'
)).toBe(true);
```

Add default true, explicit false exactly one slide, all-thead default header count, explicit header count, body fallback header, false+header control rejection, named layout margin/background/section/slide-number state, width/height bottom edge, append after existing slides, repeated invocation, frozen returned identities, move/delete/write/reopen, and input/DOM detachment.

Inject failures after first slide/table/generated slide and assert package part/relationship hashes, mutation journal, sections, existing model identities, and retry bytes match a clean run.

- [ ] **Step 2: Run public tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts \
  packages/sdk/src/index.test.ts \
  -t "tableToSlides|HTML table"
```

Expected: FAIL because `PptxDocument.tableToSlides` is absent.

- [ ] **Step 3: Implement layout preparation and one-transaction commit**

Add an internal preparer returning frozen rows/table options/master name. Resolve layout by exact name before mutation. Compute margins with explicit → layout → canonical order, default x/y/width, and transform `height` into an equivalent stricter bottom margin.

Commit with this structure:

```ts
async tableToSlides(
  elementId: string,
  options: TableToSlidesOptions = {},
): Promise<readonly SlideModel[]> {
  const prepared = prepareTableToSlides(this, elementId, options);
  const createdPartUris: string[] = [];
  try {
    const slides = this.opcPackage.transaction(() => {
      const first = this.addSlide(prepared.slideOptions);
      createdPartUris.push(first.partUri);
      first.addTable(prepared.rows, prepared.tableOptions);
      const pages = Object.freeze([first, ...first.newAutoPagedSlides]);
      createdPartUris.push(...first.newAutoPagedSlides.map(({ partUri }) => partUri));
      return pages;
    });
    return slides;
  } catch (error) {
    for (const partUri of createdPartUris) this.discardDetachedSlideModel(partUri);
    throw error;
  }
}
```

When autoPage is false, omit every auto-page-only key. When true, set omitted header count from thead or one body row. Preserve `name`, exact columns, first/continuation coordinates, weights, and effective margins.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts \
  packages/sdk/src/index.test.ts \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-content-measurement.internal.test.ts
pnpm --config.verify-deps-before-run=false exec tsc -b --pretty false
git diff --check
git add packages/sdk/src/index.ts packages/sdk/src/index.test.ts \
  packages/sdk/src/table-to-slides.ts packages/sdk/src/table-to-slides.test.ts
git commit -m "feat: generate slides from HTML tables"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review nested transaction behavior, detached model cleanup, section/layout ownership, false auto-page, header defaults, bottom edge, returned identity, and zero partial mutation.

### Task 5: Per-Page Shape, Table, and Text Additions

**Files:**
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/sdk/src/table-to-slides.ts`
- Modify: `packages/sdk/src/table-to-slides.test.ts`

**Interfaces:**
- Consumes: normalized `addShape`, `addTable`, and `addText` templates.
- Produces: one page-local object of each requested kind on every HTML-generated page, in shape → table → text order after the HTML table.

- [ ] **Step 1: Write failing additional-object tests**

Create a multi-page HTML table and request:

```ts
const pages = await deck.tableToSlides('data', {
  addShape: { type: 'roundRect', options: { x: 10, y: 20, width: 30, height: 40 } },
  addTable: { rows: [['K', 'V']], options: { x: 50, y: 60, rowHeights: 100 } },
  addText: { text: richParagraphs, options: { x: 70, y: 80, width: 90, height: 100 } },
});
expect(pages.every((page) => page.shapes.map(({ kind }) => kind).join(',') ===
  'table,shape,table,text')).toBe(true);
```

Require identical semantic templates but distinct page-local shape IDs/objects, rich-text links with page-local relationships, caller template detachment, placeholder placement on named layouts, and one nested auto-paged additional table whose continuation is not included in `pages`.

Inject invalid shape type, invalid table topology, text hyperlink target, placeholder, and generated-page failures; require whole operation rollback and byte-identical retry.

- [ ] **Step 2: Run focused tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts \
  packages/sdk/src/index.test.ts \
  -t "additional shape|additional table|additional text"
```

Expected: FAIL because additions are normalized but not committed.

- [ ] **Step 3: Commit additions after all HTML pages exist**

Inside the same outer transaction and after `[first, ...newAutoPagedSlides]` is complete:

```ts
for (const page of htmlPages) {
  if (prepared.addShape) page.addShape(prepared.addShape.type, prepared.addShape.options);
  if (prepared.addTable) page.addTable(prepared.addTable.rows, prepared.addTable.options);
  if (prepared.addText) {
    if (typeof prepared.addText.text === 'string') {
      page.addText(prepared.addText.text, prepared.addText.options);
    } else {
      page.addRichText(prepared.addText.text, prepared.addText.options);
    }
  }
}
```

Normalize outer descriptors without materializing or modifying nested native options. Use original HTML page snapshot for iteration even if an additional table inserts its own continuation slides.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts packages/sdk/src/index.test.ts \
  packages/model/src/model.test.ts
pnpm --config.verify-deps-before-run=false exec tsc -b --pretty false
git diff --check
git add packages/sdk/src/index.ts packages/sdk/src/index.test.ts \
  packages/sdk/src/table-to-slides.ts packages/sdk/src/table-to-slides.test.ts
git commit -m "feat: add objects to HTML table slides"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review object order, page-local IDs/links, stable HTML page snapshot, nested pagination boundary, caller detachment, and rollback.

### Task 6: Resolve-Once Image Additions and Atomic Media Lifecycle

**Files:**
- Create: `packages/sdk/src/prepared-image-source.ts`
- Create: `packages/sdk/src/prepared-image-source.test.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/sdk/src/table-to-slides.ts`

**Interfaces:**
- Consumes: one `ImageSource`, `AddImageSourceOptions`, optional abort, and each target slide.
- Produces:

```ts
export interface PreparedImageSource {
  readonly resolved: Readonly<ResolvedImageSource>;
  readonly options: Readonly<NormalizedAddImageSourceOptions>;
  readonly fallbackPngBytes?: Uint8Array;
}

export function prepareImageSource(
  source: ImageSource,
  options?: AddImageSourceOptions,
): Promise<Readonly<PreparedImageSource>>;

export function commitPreparedImage(
  slide: SlideModel,
  prepared: Readonly<PreparedImageSource>,
): ImageModel;
```

- [ ] **Step 1: Write failing shared-preparation and table-page tests**

Move existing `PptxDocument.addImage()` PNG/JPEG/GIF/SVG/fallback/sizing assertions into shared helper coverage without reducing cases. Add a fetch-counting source and multi-page table:

```ts
const pages = await deck.tableToSlides('data', {
  addImage: {
    source: 'https://images.example/logo.png',
    options: { x: 10, y: 20, width: 30, height: 40 },
  },
});
expect(fetchCount).toBe(1);
expect(pages.every((page) => page.shapes.filter(({ kind }) => kind === 'image').length === 1))
  .toBe(true);
expect(uniqueMediaPayloadTargets(deck)).toHaveLength(1);
```

Add SVG+fallback, Blob/data URI/bytes, asserted MIME, per-page placeholder sizing, page-local relationships, exact content-type-and-bytes payload dedup, invalid source/signature/fallback, abort before/during load, and failure after media commit. Add model coverage proving repeated raster and SVG payloads share parts while replacement remains clone-on-write. Require load failures occur before mutation and commit failures roll back all slides/media/relationships.

- [ ] **Step 2: Run image tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/prepared-image-source.test.ts \
  packages/sdk/src/index.test.ts \
  -t "prepared image|HTML table image"
```

Expected: FAIL because shared preparation and per-page image commit are absent.

- [ ] **Step 3: Extract existing image logic and integrate it**

Move option normalization, source resolution, content-type assertion, SVG fallback resolution, and placeholder-aware sizing from `PptxDocument.addImage()` into the helper without behavior changes. In `SlideModel.addImage()` and `SlideModel.addSvgImage()`, reuse an existing `/ppt/media` part only when both the exact content type and bytes match; otherwise allocate a new part. Preserve the existing relationship-per-picture model and replacement clone-on-write behavior. Keep `PptxDocument.addImage()` as:

```ts
async addImage(slideIndex: number, source: ImageSource, options: AddImageSourceOptions = {}) {
  const slide = this.slides[slideIndex];
  if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
  return commitPreparedImage(slide, await prepareImageSource(source, options));
}
```

`tableToSlides()` awaits optional preparation before its transaction, then calls `commitPreparedImage()` first on each HTML page. Clone byte arrays at public boundaries, reuse immutable prepared state, and use the new exact image-part dedup plus existing relationship-aware clone-on-write.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/prepared-image-source.test.ts \
  packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/svg-image-fallback.test.ts \
  packages/sdk/src/index.test.ts \
  packages/sdk/src/table-to-slides.test.ts \
  packages/model/src/model.test.ts
pnpm --config.verify-deps-before-run=false exec tsc -b --pretty false
git diff --check
git add packages/sdk/src/prepared-image-source.ts \
  packages/sdk/src/prepared-image-source.test.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts packages/sdk/src/table-to-slides.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: add images to HTML table slides"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review unchanged standalone addImage placement/source behavior, resolve-once count, byte detachment, SVG fallback, placeholder sizing, exact payload dedup, replacement clone-on-write, abort timing, and whole-operation isolation.

### Task 7: Public Surface, Six Formats, and PptxGenJS Contracts

**Files:**
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/build-npm-package-types.mjs`

**Interfaces:**
- Consumes: complete root/SDK exports and PptxGenJS 4.0.1 legal DOM behavior.
- Produces: closed declaration/runtime/format/conformance contract for `tableToSlides` and all public option/addition types.

- [ ] **Step 1: Write failing root/type/format/conformance tests**

In SDK/root tests, import every public type and require `typeof PptxDocument.prototype.tableToSlides === 'function'`. Build one fake-DOM deck in each `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, `.potm` format; write/reopen and require format preservation, editable tables, style state, page count, layout, and zero validation errors.

Add PptxGenJS 4.0.1 fixtures for common RGB/font/padding/border/alignment/spans/widths/repeated headers/master/additions. Compare semantic rows/styles/columns/layout/object counts, not private mutable line arrays. Add explicit difference assertions for void vs Promise, option mutation, ignored `autoPage: false`, fixed-width-as-min defect, and strict invalid inputs.

In the installed declaration consumer add:

```ts
const htmlPages: Promise<readonly SlideModel[]> = document.tableToSlides('table', {
  autoPage: true,
  autoPageRepeatHeader: true,
  addImage: { source: new Uint8Array(), options: { width: 1, height: 1 } },
  addShape: { type: 'rect' },
  addTable: { rows: [['A']] },
  addText: { text: [{ runs: [{ text: 'Title', style: { bold: true } }] }] },
});
```

Use non-empty signature-valid image bytes in runtime fixtures; the declaration snippet only needs typechecking.

- [ ] **Step 2: Run public tests and verify gaps**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "tableToSlides|HTML table"
node scripts/build-npm-package-types.mjs
```

Expected before closure: at least root capability, six-format, adapter, or packed declaration assertions fail.

- [ ] **Step 3: Close exports and aggregate contracts**

Export public types from SDK/root, add exact capability assertions, keep DOM-only runtime diagnostics clear, and update declaration builder imports. Do not add `tableToSlides` to model or adapter runtime classes.

- [ ] **Step 4: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
pnpm --config.verify-deps-before-run=false exec tsc -b --pretty false
pnpm --config.verify-deps-before-run=false run build
node scripts/build-npm-package-types.mjs
git diff --check
git add packages/sdk/src/index.ts packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/build-npm-package-types.mjs
git commit -m "test: verify HTML table slide contracts"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review root/SDK identity, public declaration closure, six formats, common PptxGenJS semantics, and documented differences.

### Task 8: Capability Documentation and Examples

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: verified public API and focused contract results.
- Produces: consistent user/API/compatibility/progress/changelog coverage without release-proof numbers guessed in advance.

- [ ] **Step 1: Add executable native examples**

Document browser usage with `await document.tableToSlides('report-table', { ... })`, EMU geometry, default true/explicit false, thead/body/tfoot, CSS mapping, data width attributes, repeated headers, master margin, and image/shape/table/rich-text additions. Show that returned pages are editable ordinary `SlideModel`s.

- [ ] **Step 2: Record exact boundaries**

Move DOM table importer from unsupported to supported in the compatibility matrix. Record async Promise/EMU, strict DOM/CSS parser, hidden table rule, cell-only styles, alpha/border-style limits, no word-level HTML, and deliberate 4.0.1 mutation/false-autoPage/column defects. State public capability coverage reaches 100% but final full-parity certification waits for peer/client audit.

- [ ] **Step 3: Search stale status and validate Markdown**

```bash
rg -n "tableToSlides|DOM table importer|99\.7|unsupported|remaining capability" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --check
```

Keep historical checkpoints labeled as historical; remove only current claims that the capability is missing.

- [ ] **Step 4: Review, commit, and push**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: document HTML table slides"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review naming, units, async examples, compatibility state, no premature release counts, and no premature full-parity claim.

### Task 9: Actual Package, Browser, and Client Proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual packed package, browser conditional export, real DOM/CSS, installed CLI/Inspector, and final evidence decks.
- Produces: stable `tableToSlides: true`, `tableToSlidesInspect: true`, Chrome state, deterministic package hashes, and retained client evidence.

- [ ] **Step 1: Extend installed package smoke**

Add a DOM-like Node consumer fixture for row/style/width/pagination/additions/reopen/failure isolation and expose:

```json
{
  "tableToSlides": true,
  "tableToSlidesState": {
    "created": true,
    "styles": true,
    "widths": true,
    "headers": true,
    "layout": true,
    "additions": true,
    "relationships": true,
    "edited": true,
    "reopened": true,
    "validationErrors": 0
  },
  "tableToSlidesInspect": true
}
```

Write a representative Node deck containing proportional/fixed/min columns, transparent/RGB styles, colspan/rowspan, multiple thead/tbody/tfoot rows, fragmented content, named layout/margins/section, internal/external links, and four additional objects.

- [ ] **Step 2: Extend real Chrome smoke**

Inside `scripts/playwright-browser-smoke.js`, create a visible styled HTML table in the page, call the packed browser API, inspect immediate pages/styles/widths/headers/layout/additions, edit/move/delete one result, writeBlob/reopen, download a dedicated evidence deck, and add exact expected JSON. Require all table states true and global console/page/network errors zero.

- [ ] **Step 3: Run syntax/focused/build/full/performance gates**

```bash
node --check scripts/playwright-browser-smoke.js
node --check scripts/smoke-npm-package.mjs
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/table-to-slides.test.ts packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm --config.verify-deps-before-run=false exec tsc -b --pretty false
pnpm --config.verify-deps-before-run=false run build
pnpm --config.verify-deps-before-run=false --filter @jiayunxie/pptx run build
pnpm --config.verify-deps-before-run=false exec vitest run --exclude packages/performance/src/performance.test.ts
pnpm --config.verify-deps-before-run=false exec vitest run packages/performance/src/performance.test.ts
```

- [ ] **Step 4: Build two clean artifacts and run consumers**

Create `/tmp/pptx-table-to-slides-artifacts.XXXXXX`. Run two clean package builds; record sorted dist file SHA-256 manifests; pack twice and require byte-identical actual tarballs. Run installed Node, NodeNext declarations, browser conditional export, CLI, Inspector, six-format smoke, and the full package smoke script.

- [ ] **Step 5: Run real browser and client validation**

Serve the extracted actual tarball over loopback, execute the checked-in callback in installed Google Chrome, retain full/compact JSON and dedicated PPTX. Require no console/page/network errors. Run `pptx-inspect` package/slide/part checks and PowerPoint 2010 validation; accept only intentional external-link warnings. Render source at 180 DPI, run overflow checks, inspect every page, and perform LibreOffice open/save/PDF degradation comparison without overstating CSS fidelity.

- [ ] **Step 6: Review, commit, and push**

```bash
git diff --check
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git commit -m "test: verify packed HTML table slides"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count HEAD...origin/main
```

Review actual tarball isolation, manifest/tarball determinism, state booleans, exact deck structure, relationship ownership, Chrome errors, client diagnostics, visual overflow, and retained evidence path.

### Task 10: Final TableToSlides Proof and Specialty Closure

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: exact Task 9 gate output, hashes, Chrome version/state, client diagnostics, deck metrics, commit chain, and evidence path.
- Produces: one consistent final proof record and transition to the final peer/client audit.

- [ ] **Step 1: Record only observed final evidence**

In all six documents add the same focused/full/performance counts and timings, typecheck/root/package build times, dist/tarball file counts and hashes, installed consumer states, Chrome version/errors, source/browser deck hashes, part/relationship/slide/table/addition counts, rows/headers/styles/width/layout/section/link metrics, PowerPoint/LibreOffice/render results, commit chain, and `/tmp` proof directory.

- [ ] **Step 2: Close the capability matrix without closing the total goal**

Set `tableToSlides` specialty to its final completed task count and public capability coverage to 100%. Remove current unsupported entries. Keep full PptxGenJS parity certification explicitly pending until the next final peer/client audit proves the entire declaration matrix and consumer corpus.

- [ ] **Step 3: Fresh review and consistency checks**

```bash
rg -n "tableToSlides|DOM table importer|100%|full parity|peer/client audit|actual tarball|Chrome|PowerPoint 2010" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --check
```

Cross-check every number/hash against retained files and command output. Confirm Markdown fences balance and historical percentages remain labeled as checkpoints.

- [ ] **Step 4: Commit, push, fetch, and confirm clean sync**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: record HTML table slide proof"
git -c http.version=HTTP/1.1 push origin main
git -c http.version=HTTP/1.1 fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected divergence: `0 0`; only existing `.pnpm-store/` remains untracked. Report completed item, remaining final audit, and progress, then immediately begin the peer/client full-surface audit under a separate audited plan.

## Plan Self-Review

- **Spec coverage:** Tasks 1–3 own strict DOM/CSS/column input; Tasks 4–6 own pagination, lifecycle, additions, image async, and rollback; Task 7 closes public/format/conformance contracts; Tasks 8–10 close docs and actual consumer/client proof. Every design requirement has an implementation or evidence owner.
- **Type consistency:** Public names match the design: `TableToSlidesOptions`, four `TableToSlidesAdd*` records, `PptxDocument.tableToSlides()`, and `Promise<readonly SlideModel[]>`. Existing native types are reused without aliases that change units.
- **Scope:** This plan completes one subsystem—editable HTML table import to slides. The final full-surface audit remains a separate next plan because it verifies all public subsystems rather than implementing this importer.
- **Execution:** The user delegated all choices and requested uninterrupted work. Execute inline in this task, one numbered task at a time, with review/commit/push/progress sync after every task.
