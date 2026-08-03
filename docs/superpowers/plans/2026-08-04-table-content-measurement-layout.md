# Table Content Measurement and Layout Recalculation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add deterministic automatic table-row measurement, strict table/cell auto-page weights, rich-text row fragmentation, placeholder auto-page, and exact per-page table layout recomputation without regressing fixed-height pagination.

**Architecture:** Normalize measurement intent into the existing detached table definition, resolve one immutable page-layout region, and use a new pure content-measurement module to materialize positive EMU rows before the existing page partitioner runs. Rich row fragments remain ordinary normalized physical rows, while the existing slide insertion, relationship, section, identity, and outer transaction paths commit each page; placeholder auto-page changes only effective placement and commit-time scaling policy.

**Tech Stack:** TypeScript 5, Vitest, immutable normalized rich-text values, lossless OOXML source spans, safe-integer EMU arithmetic, OPC transactions, pnpm/tsup, PptxGenJS 4.0.1 conformance fixtures, installed-package Node/NodeNext/browser/CLI smoke, Google Chrome, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- `autoPageCharWeight` and `autoPageLineWeight` are finite inclusive `[-1, 1]`; table values are defaults, cell values override, and explicit zero remains an own normalized property.
- Existing all-positive/no-weight `rowHeights` stay exact. Automatic/zero rows or any supplied weight activate measurement; positive values in measured mode are lower bounds and never shrink.
- A zero-containing `rowHeights` vector is legal only with `autoPage: true` and without `height`; ordinary table creation retains its current positive explicit-height contract.
- Measurement is pure and independent of installed fonts, Canvas, locale segmentation, package/XML state, relationships, time, and randomness.
- Default measurement uses 12pt, `2.3 + charWeight`, `1.67 + lineWeight`, canonical table-cell margins, and one rounded safe-integer EMU conversion per value.
- Do not split repeated headers, rowspan blocks, surrogate pairs, combining clusters, variation sequences, emoji modifiers, or ZWJ sequences.
- A content-driven oversized body row may split only at measured line-band boundaries; fragments retain cell formatting and semantic hyperlinks while page-local relationship IDs are regenerated.
- Placeholder auto-page uses placeholder X/Y/width and bottom boundary, does not stretch page rows to placeholder height, and keeps non-auto placeholder creation byte behavior unchanged.
- Resolve placeholder/layout state, measurement, fragmentation, page partition, slide insertion, and link targets before the first package mutation.
- Keep source/generated table writes, placeholders, relationships, presentation order, sections, slide numbers, caches, and `newAutoPagedSlides` failure-isolated under the existing outer transaction.
- Each task ends with focused review, commit, push, and `origin/main...HEAD` divergence `0 0`; do not stage `.pnpm-store/` or retained proof artifacts.

---

### Task 1: Strict Measurement Intent and Weight Normalization

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`
- Modify: `packages/model/src/table-auto-page.internal.ts`

**Interfaces:**
- Consumes: descriptor-safe table/cell option records, physical row count, normalized `rowHeights`, and `autoPage` state.
- Produces:

```ts
export interface NormalizedTableAutoPageRequest {
  readonly repeatHeader: boolean;
  readonly headerRows: number;
  readonly slideStartY?: number;
  readonly slideMargin?: readonly [number, number, number, number];
  readonly charWeight?: number;
  readonly lineWeight?: number;
  readonly measureContent: boolean;
}

export interface NormalizedTableCell {
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
}
```

- `measureContent` is true when rows are automatic/zero or any table/cell weight is present; optional weight properties preserve explicit zero.

- [ ] **Step 1: Write failing table/cell normalization tests**

Add to `table-create.internal.test.ts`:

```ts
it('normalizes automatic rows and strict table/cell measurement weights', () => {
  const definition = normalizeTableDefinition([
    [{ text: 'A', options: { autoPageCharWeight: -0.25, autoPageLineWeight: 0 } }],
    ['B'],
  ], {
    autoPage: true,
    autoPageCharWeight: 0,
    autoPageLineWeight: 0.5,
    rowHeights: [0, 400_000],
  });
  expect(definition.autoPage).toMatchObject({
    charWeight: 0,
    lineWeight: 0.5,
    measureContent: true,
  });
  expect(definition.rows[0]![0]).toMatchObject({
    autoPageCharWeight: -0.25,
    autoPageLineWeight: 0,
  });
  expect(definition.rowHeights).toEqual([0, 400_000]);
});
```

Add cases for omitted rows/height, all-positive fixed mode, weight-triggered minimum mode, scalar zero, mixed zero, zero plus `height`, weight without `autoPage`, cell weight without `autoPage`, `NaN`, infinities, `-1`, `1`, values just outside both bounds, numeric strings, accessors, inherited/symbol fields, class instances, and caller mutation after normalization. Require accessor call count zero and frozen auto-page metadata.

- [ ] **Step 2: Run the focused tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-create.internal.test.ts \
  -t "measurement weights|automatic rows"
```

Expected: FAIL because weight keys are unsupported and auto-page still rejects zero rows.

- [ ] **Step 3: Implement early auto-page state and strict weights**

Add `autoPageCharWeight` / `autoPageLineWeight` to table and cell allowed-key sets. Validate `autoPage` before row-vector normalization so zero is accepted only for the enabled path. Use one helper:

```ts
function normalizeAutoPageWeight(value: unknown, context: string): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < -1 || value > 1) {
    throw new RangeError(`${context} must be between -1 and 1`);
  }
  return Object.is(value, -0) ? 0 : value;
}
```

Remove only the automatic-row rejection from `normalizeTableAutoPageRequest()`. Keep placeholder rejection until Task 7. Reject zero row plus `height`, compute `measureContent` from zero/automatic rows and own weight presence, and keep the all-positive/no-weight definition identical to the current fixed path.

- [ ] **Step 4: Run focused/type gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-auto-page.internal.ts
git commit -m "feat: normalize table measurement intent"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review descriptor safety, explicit-zero ownership, fixed-mode equality, zero/height cross-field behavior, and absence of rendered XML changes. Expected divergence: `0 0`.

### Task 2: Shared Auto-Page Layout Region

**Files:**
- Modify: `packages/model/src/table-auto-page.internal.ts`
- Modify: `packages/model/src/table-auto-page.internal.test.ts`
- Modify: `packages/model/src/slide.ts`

**Interfaces:**
- Consumes: a normalized auto-page definition, slide size, optional runtime layout margins, and optional stricter bottom edge.
- Produces:

```ts
export interface TableAutoPageLayoutRegion {
  readonly firstY: number;
  readonly continuationY: number;
  readonly bottomEdge: number;
  readonly firstCapacity: number;
  readonly continuationCapacity: number;
}

export function resolveTableAutoPageLayout(
  definition: Readonly<NormalizedTableDefinition>,
  slideSize: Readonly<SlideSize>,
  layoutMargins?: Readonly<TableAutoPageMargins>,
  bottomEdgeOverride?: number,
): Readonly<TableAutoPageLayoutRegion>;

export function planTableAutoPages(
  definition: Readonly<NormalizedTableDefinition>,
  region: Readonly<TableAutoPageLayoutRegion>,
): readonly Readonly<NormalizedTableDefinition>[];
```

- [ ] **Step 1: Write failing region tests**

Add exact assertions:

```ts
it('resolves one frozen layout region for measurement and partition', () => {
  const source = definition([['A']], {
    y: 20,
    rowHeights: [10],
    autoPageSlideStartY: 5,
    slideMargin: [3, 4, 7, 6],
  });
  expect(resolveTableAutoPageLayout(source, SLIDE)).toEqual({
    firstY: 20,
    continuationY: 5,
    bottomEdge: 93,
    firstCapacity: 73,
    continuationCapacity: 88,
  });
});
```

Cover canonical/layout/explicit margin precedence, exact bottom override, override above canonical bottom, override at/before first or continuation Y, horizontal margin exhaustion, unsafe dimensions, safe-integer overflow, and frozen output. Re-run every existing partition fixture through the resolved region and require unchanged page rows/Y/heights.

- [ ] **Step 2: Run the region tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-auto-page.internal.test.ts
```

Expected: FAIL because the region resolver and new planner signature do not exist.

- [ ] **Step 3: Extract geometry without changing fixed pagination**

Move slide/margin/bottom/start/capacity validation out of `planTableAutoPages()` into `resolveTableAutoPageLayout()`. Normalize `bottomEdgeOverride` as a positive safe integer and use `Math.min(canonicalBottomEdge, bottomEdgeOverride)`. The planner must consume only the frozen region plus definition rows/heights and continue stripping `autoPage` from page definitions.

Update `SlideModel.addTable()` to resolve the region once before calling the planner. Do not invoke content measurement yet; automatic inputs may still fail safely in the planner without package mutation.

- [ ] **Step 4: Run regression gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/model.test.ts \
  -t "auto-page|auto-paged|newAutoPagedSlides"
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/table-auto-page.internal.ts \
  packages/model/src/table-auto-page.internal.test.ts packages/model/src/slide.ts
git commit -m "refactor: share table auto-page layout"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review unchanged fixed page partitions, one geometry authority, override validation, no package access, and exact EMU arithmetic. Expected divergence: `0 0`.

### Task 3: Deterministic Rich-Text Line Measurement

**Files:**
- Modify: `packages/model/src/rich-text.internal.ts`
- Create: `packages/model/src/table-content-measurement.internal.ts`
- Create: `packages/model/src/table-content-measurement.internal.test.ts`

**Interfaces:**
- Consumes: one normalized anchor cell, exact colspan width, and resolved table weights.
- Produces:

```ts
export interface MeasuredTableRunSlice {
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly text: string;
  readonly startsAtRunStart: boolean;
  readonly endsAtRunEnd: boolean;
  readonly retainsSoftBreak: boolean;
}

export interface MeasuredTableLine {
  readonly paragraphIndex: number;
  readonly slices: readonly Readonly<MeasuredTableRunSlice>[];
  readonly height: number;
  readonly startsParagraph: boolean;
  readonly endsParagraph: boolean;
}

export interface MeasuredTableCellContent {
  readonly paragraphs: readonly Readonly<NormalizedRichTextParagraph>[];
  readonly lines: readonly Readonly<MeasuredTableLine>[];
  readonly topMargin: number;
  readonly rightMargin: number;
  readonly bottomMargin: number;
  readonly leftMargin: number;
}

export function measureTableCellContent(
  cell: Readonly<NormalizedTableCell>,
  width: number,
  tableCharWeight?: number,
  tableLineWeight?: number,
): Readonly<MeasuredTableCellContent>;
```

- Export `NormalizedRichTextRun` and `NormalizedRichTextParagraph` as internal readonly types; do not re-export them from `packages/model/src/index.ts`.

- [ ] **Step 1: Write failing cluster/wrap tests**

Create `table-content-measurement.internal.test.ts` with zero cell margins and 12pt text. Set `advance = Math.round(12 * 12_700 / 2.3)` and `width = 5 * advance`; assert `Alpha` fits exactly, the next whitespace wraps, a long word splits without losing text, and increasing/decreasing char weight moves the exact boundary in opposite directions.

Add table-driven cases for ASCII punctuation, tabs with/without declared stops, Latin/Greek/Cyrillic, CJK, surrogate emoji, combining marks, variation selectors, emoji modifiers, ZWJ families, whitespace retention, empty runs, empty paragraphs, paragraph boundaries, run `softBreakBefore`, and `breakLine`-normalized paragraphs. Assert concatenated slices reproduce the exact normalized text and no slice starts/ends inside a cluster.

- [ ] **Step 2: Write failing style/spacing tests**

Measure rich runs with 10/20pt overrides, negative/positive character spacing, cell 18pt fallback, default 12pt, exact line spacing, multiple line spacing, before/after spacing, paragraph left/right margins, first-line indent, bullet indent, colspan width, default/custom margins, vertical direction, and fit. Assert the base line-height formula:

```ts
expect(single.lines[0]!.height)
  .toBe(Math.round(12 * 1.67 * 914_400 / 100));
```

Also reject zero/negative usable inline width and safe-sum overflow without mutating the cell or paragraphs.

- [ ] **Step 3: Run the new tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts
```

Expected: FAIL because the measurement module does not exist.

- [ ] **Step 4: Implement clusterization, wrapping, and line boxes**

Implement focused helpers in the new file:

```ts
interface TextCluster {
  readonly text: string;
  readonly codePoints: readonly number[];
}

function textClusters(value: string): readonly Readonly<TextCluster>[];
function clusterUnits(cluster: Readonly<TextCluster>): number;
function availableLineWidth(
  paragraph: Readonly<NormalizedRichTextParagraph>,
  width: number,
  margins: Readonly<Required<TextBoxMargins>>,
  firstLine: boolean,
): number;
function naturalLineHeight(fontSize: number, lineWeight: number): number;
```

Use the exact code-point ranges and precedence from the design spec. Wrap on whitespace first and split only an overlong non-whitespace token. Resolve run font/spacing and cell/table weights with nullish checks, never truthiness. Freeze every returned array/object and check all additions before overflow.

- [ ] **Step 5: Run focused/type gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts \
  packages/model/src/table-create.internal.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/rich-text.internal.ts \
  packages/model/src/table-content-measurement.internal.ts \
  packages/model/src/table-content-measurement.internal.test.ts
git commit -m "feat: measure table cell content"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review exact text reproduction, cluster safety, paragraph spacing precedence, no font/environment dependency, and immutable outputs. Expected divergence: `0 0`.

### Task 4: Row Bands, Rowspan Constraints, and Height Materialization

**Files:**
- Modify: `packages/model/src/table-content-measurement.internal.ts`
- Modify: `packages/model/src/table-content-measurement.internal.test.ts`

**Interfaces:**
- Consumes: measured cells in normalized physical rows, column widths, and optional row minimums.
- Produces:

```ts
export interface MeasuredTableRowLayout {
  readonly sourceRowIndex: number;
  readonly bands: readonly number[];
  readonly contentHeight: number;
  readonly height: number;
  readonly fragmentable: boolean;
  readonly cells: readonly (Readonly<MeasuredTableCellContent> | undefined)[];
}

export function measureTableAutoPageRows(
  definition: Readonly<NormalizedTableDefinition>,
): readonly Readonly<MeasuredTableRowLayout>[];

export function materializeMeasuredTableRows(
  definition: Readonly<NormalizedTableDefinition>,
  measuredRows: readonly Readonly<MeasuredTableRowLayout>[],
): Readonly<NormalizedTableDefinition>;
```

- [ ] **Step 1: Write failing row-band and minimum tests**

Build two-column rows whose cells have unequal line counts and font sizes. Assert the Nth band is the max Nth cell-line height, total content height is max top margin + band sum + max bottom margin, empty cells still contribute one default line, colspan anchors use the exact summed width, horizontal continuation cells contribute nothing, and positive row minimums never shrink.

Assert an all-positive/no-weight fixed definition is returned by identity from `materializeMeasuredTableRows()`, while automatic/zero/weight definitions return detached frozen rows, all-positive heights, `autoRowHeight: false`, and `height === rowHeights.reduce(...)`.

- [ ] **Step 2: Write failing rowspan constraint tests**

Create short ordinary cells plus 2-row, 3-row, adjacent, and nested-active rowspan anchors. Assert constraints are processed by span length/row/column and deficits distribute quotient/remainder from the first covered row:

```ts
expect(materialized.rowHeights).toEqual([50, 50, 11]);
expect(materialized.height).toBe(111);
```

Use a fixture whose 2-row span requires 80 and 3-row span requires 111 so the short constraint remains satisfied after the long constraint. Cover rows containing only merge continuations, 1-EMU automatic minimum, and safe-sum rejection.

- [ ] **Step 3: Run row tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  -t "row band|rowspan|materializes"
```

Expected: FAIL because row aggregation/materialization exports are absent.

- [ ] **Step 4: Implement deterministic row layout**

Measure only anchor cells. Build bands by ordinal line index, add effective row margins once, seed each row with `max(1, inputMinimum, ordinaryContentHeight)`, and then apply sorted rowspan lower-bound constraints with safe quotient/remainder distribution. `materializeMeasuredTableRows()` performs geometry replacement without page-capacity policy: preserve row/cell objects when no fragmentation is needed, replace only geometry fields, and freeze new containers. Public orchestration does not call this internal stage directly.

- [ ] **Step 5: Run focused/type gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/table-content-measurement.internal.ts \
  packages/model/src/table-content-measurement.internal.test.ts
git commit -m "feat: materialize measured table rows"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review band arithmetic, lower-bound ordering, fixed fast path, transform sum, continuation exclusion, and source immutability. Expected divergence: `0 0`.

### Task 5: Rich Text Row Fragmentation

**Files:**
- Modify: `packages/model/src/table-content-measurement.internal.ts`
- Modify: `packages/model/src/table-content-measurement.internal.test.ts`
- Modify: `packages/model/src/table-auto-page.internal.test.ts`

**Interfaces:**
- Consumes: measured body rows and `continuationCapacity - measuredHeaderHeight`.
- Produces:

```ts
interface FragmentedTableRow {
  readonly row: readonly Readonly<NormalizedTableCell>[];
  readonly height: number;
}

export function materializeTableAutoPageContent(
  definition: Readonly<NormalizedTableDefinition>,
  region: Readonly<TableAutoPageLayoutRegion>,
): Readonly<NormalizedTableDefinition>;
```

- [ ] **Step 1: Write failing plain/rich fragment tests**

Use a one-line body capacity to split one physical row into three fragments. Assert ordered content concatenation equals the source, each fragment height fits, margins repeat, empty peer cells remain legal, unequal cell line counts stay aligned by band index, colspan/hMerge topology is identical, and all page rows are frozen.

Use rich paragraphs containing before/after/exact spacing, a bullet, a soft break, run-local URL/internal links, outer cell hyperlink, font/color/bold/italic, surrogate emoji, combining text, and a ZWJ family. Assert:

- first fragment retains paragraph before/bullet, middle continuation clears both, and last retains paragraph after;
- no synthetic newline appears at a wrap-only boundary;
- a soft break is retained only when its preceding content remains in the same fragment;
- split runs preserve style/hyperlink and no cluster is cut;
- fragment `text` projections concatenate to the normalized source text after removing page boundaries.

- [ ] **Step 2: Write failing non-fragmentable/error tests**

Require no split for a row that merely does not fit the current first-page remainder but fits continuation capacity. Reject repeated-header overflow, rowspan/vMerge blocks, fixed/minimum blank height as the sole overflow cause, zero continuation body capacity, and a single content band plus effective margins larger than capacity. Assert no input objects change on every failure.

- [ ] **Step 3: Run fragment tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts \
  -t "fragment|oversized measured row"
```

Expected: FAIL because oversized measured rows still reject.

- [ ] **Step 4: Implement canonical fragment cloning**

Add focused helpers:

```ts
function fragmentMeasuredRow(
  definition: Readonly<NormalizedTableDefinition>,
  measured: Readonly<MeasuredTableRowLayout>,
  bodyCapacity: number,
): readonly Readonly<FragmentedTableRow>[];

function fragmentCellParagraphs(
  measured: Readonly<MeasuredTableCellContent>,
  firstBand: number,
  endBand: number,
): readonly Readonly<NormalizedRichTextParagraph>[];
```

Greedily select the largest non-empty band prefix that fits after repeated effective margins. Rejoin adjacent slices from the same source paragraph/run, clone style/hyperlink values, clear only boundary-owned paragraph fields, recompute projected text and fragment height, and reuse continuation cells/topology without measuring them twice.

- [ ] **Step 5: Run focused/type gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/table-content-measurement.internal.ts \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts
git commit -m "feat: fragment oversized table text rows"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review content completeness, boundary spacing/bullets, cluster safety, link values, merge prohibition, maximum-prefix termination, and positive fragment heights. Expected divergence: `0 0`.

### Task 6: Measured Auto-Page Orchestration

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `resolveTableAutoPageLayout()`, `materializeTableAutoPageContent()`, `planTableAutoPages()`, prepared links, and prepared slide insertion.
- Produces public fields:

```ts
export interface AddTableOptions {
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
}

export interface AddTableCellOptions {
  readonly autoPageCharWeight?: number;
  readonly autoPageLineWeight?: number;
}
```

- [ ] **Step 1: Write failing automatic public creation tests**

In `model.test.ts`, create automatic rows with plain, rich, CJK, soft-break, colspan, rowspan, custom margins, paragraph spacing, table weights, cell overrides, repeated headers, and an oversized fragmentable body row. Assert exact page text, positive row vectors, each transform height equals its row sum, each page stays within its resolved capacity, fixed explicit fixtures remain byte-identical, and `newAutoPagedSlides` keeps its existing lifecycle semantics.

Add a table-weight fixture where explicit zero and omitted produce equal numeric measurement but distinct normalized ownership, plus two cell overrides that create different line counts in equal-width columns.

- [ ] **Step 2: Write failing relationship and preflight tests**

Split a rich linked row across at least two pages. Assert every page owns exactly the URL/internal relationships its XML references, repeated header links are page-local, target slide part URIs remain stable after insertion, and caller hyperlink objects receive no `_rId`. For invalid weights, impossible width, oversized single band, header overflow, and rowspan block overflow, snapshot package bytes/journal/slide identities/runtime result and require exact zero mutation.

- [ ] **Step 3: Run integration tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/model.test.ts \
  -t "measures automatic table rows|fragments rich auto-page rows|measurement preflight"
```

Expected: FAIL because `SlideModel.addTable()` does not call the materializer and public weight fields are absent.

- [ ] **Step 4: Insert pure materialization before partition**

Use this exact order in `addTable()`:

```ts
const definition = normalizeTableDefinition(rows, options);
const region = definition.autoPage === undefined
  ? undefined
  : resolveTableAutoPageLayout(definition, slideSize, layoutMargins);
const materialized = region === undefined
  ? definition
  : materializeTableAutoPageContent(definition, region);
const pages = region === undefined
  ? Object.freeze([materialized])
  : planTableAutoPages(materialized, region);
```

Prepare page links and insertion plans only after fragmentation but before mutation. Keep the existing outer transaction and catch-time cache cleanup unchanged. Add the two public type fields without exporting measurement internals.

- [ ] **Step 5: Run lifecycle/type gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  -t "measure|fragment|auto-page|auto-paged|newAutoPagedSlides"
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: auto-page measured table content"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review pre-mutation ordering, fragment link preparation, fixed byte stability, runtime state timing, page geometry, and no change to ordinary automatic tables. Expected divergence: `0 0`.

### Task 7: Placeholder Auto-Page and Layout Recalculation

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-auto-page.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `ResolvedPlaceholderOwner.transform` and the shared layout region.
- Produces internal definition field:

```ts
export interface NormalizedTableDefinition {
  readonly placeholderAutoPage?: true;
}
```

- Page slices copy `placeholderAutoPage`; it is never serialized or publicly exported.

- [ ] **Step 1: Write failing placeholder geometry tests**

Install a named layout with a body placeholder whose X/Y/width/height differ from table defaults. Add automatic content using `placeholder` + `autoPage` and assert placeholder width changes wrapping, source Y is owner Y, continuation Y is owner Y or explicit start, bottom edge is `min(owner bottom, slide bottom margin)`, every page transform height equals row sum rather than owner height, and source/generated tables replace their own matching owner while retaining identity/name/shape ID.

Cover fixed positive rows with placeholder auto-page, repeated headers, rich links, sections, slide numbers, generated same-layout relationships, reopen, and non-auto placeholder byte behavior.

- [ ] **Step 2: Write failing placeholder preflight/rollback tests**

Add zero width/height, missing, ambiguous, filled, wrong-domain, malformed transform, owner bottom before Y, continuation Y at/after bottom, and layout mismatch cases. Snapshot parts/relationships/dates/journal/slides/shapes/sections/runtime state before each call. Inject failures at source owner replacement, generated placeholder materialization, generated owner replacement, relationship creation, and final outer commit; require exact restoration and no cached detached slide.

- [ ] **Step 3: Run placeholder tests and verify failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/model.test.ts \
  -t "placeholder auto-page|placeholder measured layout"
```

Expected: FAIL because normalization still rejects placeholder auto-page and commit still stretches placeholder rows.

- [ ] **Step 4: Implement effective definition and commit policy**

Remove the placeholder rejection only after `SlideModel.addTable()` preflights the source owner. Build a detached effective definition with owner X/Y/width, scaled column widths, `placeholderAutoPage: true`, and no row-height scaling. Resolve the region with `bottomEdgeOverride = owner.y + owner.height`; materialize/partition that definition.

In `commitNormalizedTable()`, retain legacy scaling when owner exists and `placeholderAutoPage !== true`. When the flag is true, use owner name/identity/rotation/flip but definition X/Y/width/height/columns/rows. Copy the flag in `pageDefinition()` and strip it only by omission from rendered XML.

- [ ] **Step 5: Run lifecycle/type gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  -t "placeholder|measure|fragment|auto-page|rollback"
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/table-create.internal.ts \
  packages/model/src/table-auto-page.internal.ts \
  packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: auto-page tables in placeholders"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review owner preflight, region bottom, no stretch, local owner replacement, non-auto regression, same-layout guarantee, and rollback cache cleanup. Expected divergence: `0 0`.

### Task 8: Public Surface, PptxGenJS Conformance, and Failure Hardening

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: the public table/cell weight fields, automatic/zero rows, placeholder auto-page, fragments, and `newAutoPagedSlides`.
- Produces: declaration/runtime/conformance proof without a new public symbol.

- [ ] **Step 1: Add SDK/root type and runtime closure**

In SDK/root tests, type-check table/cell weights at `-1`, `0`, and `1`; assert string/out-of-range runtime rejection; create/reopen automatic and zero/minimum rows; verify readonly generated slides; and execute all six `PRESENTATION_FORMAT_PROFILES`. Remove prior assertions that valid zero weights are unsupported.

Require declaration emit to expose both fields through `@pptx/model`, `@pptx/sdk`, and root `pptx`, while `Normalized*`/measurement types remain absent.

- [ ] **Step 2: Add PptxGenJS 4.0.1 boundary fixtures**

Generate legal PptxGenJS plain ASCII tables at exact wrap/page boundaries for weights `-1`, `0`, and `1`, import their output, and compare page text/header order with native automatic tables using the same logical content and geometry. Add native rich font/CJK/paragraph/cell-override cases and assert documented deterministic output rather than byte identity.

Keep explicit assertions that PptxGenJS mutates caller objects, clamps `9` to `1`, loses some cell overrides, may reuse a following slide, and has its colspan/rowspan estimator defects; native must stay detached, strict, and isolated.

- [ ] **Step 3: Extend failure matrix**

Inject failures at measurement return, fragment cloning, source table write, generated slide part, layout relationship, placeholder materialization, source/generated link creation, section/slide-number sync, and outer transaction. Require exact package snapshot/write bytes, mutation journal, presentation order, sections, stable model identities, earlier successful runtime result, and deterministic retry.

- [ ] **Step 4: Run public/full-surface focused gates**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  -t "measurement|automatic table|fragment|placeholder auto-page|auto-page"
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false build
git diff --check
```

Expected: all focused tests, typecheck, and build pass.

- [ ] **Step 5: Review, commit, push, and verify synchronization**

```bash
git add packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "test: verify measured table auto-page contracts"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review public declaration closure, six formats, PptxGenJS legal baseline, deliberate defect divergence, injected boundaries, and retry determinism. Expected divergence: `0 0`.

### Task 9: Documentation and Compatibility Matrix

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Document exact API, formula, and examples**

Add examples for omitted rows, mixed zero/minimum rows, table and cell weights, rich oversized fragmentation, repeated headers, placeholder auto-page, and `newAutoPagedSlides`. State EMU inputs, 12pt/2.3/1.67 formulas, deterministic Unicode cluster rules, margin/spacing behavior, fixed-mode preservation, fragment/merge limits, placeholder no-stretch behavior, runtime reset/reopen, and strict errors.

- [ ] **Step 2: Update supported/remaining statements**

Move automatic row measurement/layout recomputation, both weights, no-rowspan text fragmentation, and placeholder auto-page to supported in all six files. Keep `tableToSlides` and final peer/client audit pending, set current overall parity to approximately `99.7%`, and do not claim 100% or final PptxGenJS parity.

- [ ] **Step 3: Run consistency and focused review**

```bash
rg -n "autoPageCharWeight|autoPageLineWeight|automatic row|fragment|placeholder auto-page|tableToSlides|99.7" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "measurement|automatic table|fragment|placeholder auto-page"
git diff --check
```

Review formula consistency, exact supported boundary, no stale unsupported claims in current sections, historical checkpoint preservation, and absence of premature packed/full proof numbers.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: document table content measurement"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 10: Actual Packed and Real-Browser Verification

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

- [ ] **Step 1: Extend installed Node/NodeNext/type smoke**

From the actual tarball only, compile valid table/cell weight and automatic/zero/minimum declarations plus invalid type cases. Create a middle source/sentinel deck containing automatic rich/CJK/soft-break rows, a fragment spanning pages, repeated headers, a rowspan block, local/default links, and placeholder auto-page. Edit/move/delete generated pages, write/reopen, and emit stable `tableContentMeasurement` / `tableContentMeasurementInspect` booleans.

- [ ] **Step 2: Write and inspect the Node evidence deck**

Write `table-content-measurement-smoke.pptx`. Run installed CLI/Inspector doctor, package inspect, PowerPoint 2010 validate, slides list, and exact page/relationship part reads. Require positive row vectors, transform sums, same layout, contiguous sections, repeated headers, complete fragment text/style/link ownership, placeholder identity, fixed merge blocks, and zero orphan links.

- [ ] **Step 3: Extend real-Chrome lifecycle**

Use the packed browser conditional export for the same create/measure/fragment/placeholder/edit/move/delete/writeBlob/reopen flow. Download `browser-table-content-measurement.pptx`; require all state flags true and zero validation, console, page, and network errors.

- [ ] **Step 4: Build/pack twice and execute proof gates**

Compare two sorted dist SHA-256 manifests and two actual tarballs byte-for-byte, then run:

```bash
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-table-content-measurement-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Serve the extracted browser module over loopback, run the checked-in callback in installed Google Chrome, and retain JSON/PPTX/manifests/tarball hashes outside the repository.

- [ ] **Step 5: Review, commit, push, and verify synchronization**

```bash
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git commit -m "test: verify packed table measurement"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review actual-package-only imports, browser-safe deterministic measurement, fragment/link/placeholder evidence, fixed-mode regression, artifact reproducibility, and zero runtime errors. Expected divergence: `0 0`.

### Task 11: Full Gates and Recorded Proof

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Run final non-overlapping gates**

```bash
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false build
pnpm --filter @jiayunxie/pptx build
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-content-measurement.internal.test.ts \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  -t "measurement|automatic table|fragment|placeholder auto-page|auto-page"
pnpm --config.verify-deps-before-run=false test
pnpm --config.verify-deps-before-run=false test:performance
```

Record clean focused/full/performance file/test counts and durations from separate runs.

- [ ] **Step 2: Re-run authoritative package/client inspection**

Verify both clean dist manifests, both actual tarballs, installed Node/NodeNext/browser/CLI/Inspector JSON, real-Chrome JSON, Node/browser deck hashes, exact slide/table/relationship part reads, and PowerPoint 2010 diagnostics. Require stable formulas/page membership across two builds and do not infer broad proof from a narrow smoke flag.

- [ ] **Step 3: Record final evidence consistently**

In all six documents, add the same final counts/timings, dist/tarball hashes, installed consumer results, Chrome version/errors, deck part/relationship/slide/table counts, automatic/minimum/fixed row evidence, fragment content/style/link evidence, placeholder/layout/section evidence, diagnostics, commit chain, proof path, specialty `11/11`, overall `99.7%`, and next item `tableToSlides`.

- [ ] **Step 4: Review documentation and repository state**

```bash
rg -n "tableContentMeasurement|automatic row|fragment|placeholder auto-page|actual tarball|Chrome|PowerPoint 2010|11/11|99.7|tableToSlides" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --check
git status --short
```

Review identical cross-file numbers, no stale current unsupported claims, no generated artifacts staged, historical checkpoint wording unchanged, and only `.pnpm-store/` left untracked.

- [ ] **Step 5: Commit, push, and verify final synchronization**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: record table measurement proof"
git -c http.version=HTTP/1.1 push origin main
git rev-list --left-right --count origin/main...HEAD
git status --short
```

Expected divergence: `0 0`. The content measurement/layout recalculation specialty is complete only after all eleven tasks pass and retained evidence proves the full stated scope.

## Execution Selection

Execute inline in this task, one task at a time, because the user delegated implementation choices, requested no confirmation prompts, and requested review/commit/push after every completed item. After every task: report completed item, remaining tasks, and specialty/overall progress, then continue automatically.
