# Table Auto-Page and Repeated Headers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict explicit-row-height table pagination, repeated headers, continuation-slide creation, and `SlideModel.newAutoPagedSlides` while preserving layout, sections, merges, rich content, hyperlinks, transactions, and stable identities.

**Architecture:** Normalize auto-page controls into the existing detached table definition, partition physical rows with a pure EMU-based row-block planner, then commit ordinary page tables through one shared renderer. A presentation primitive inserts canonical blank slides directly after the source with the same layout and section; `SlideModel` resolves all hyperlink targets before insertion and publishes runtime-only generated-slide state only after the outer transaction succeeds.

**Tech Stack:** TypeScript 5, Vitest, lossless OOXML source spans, OPC transactional package graph, pnpm/tsup, PptxGenJS 4.0.1 conformance fixtures, installed-package Node/NodeNext/browser/CLI smoke, Google Chrome, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- Native geometry values are safe-integer EMU; `slideMargin` uses scalar or exact dense `[top, right, bottom, left]` EMU.
- This specialty accepts only explicit positive physical `rowHeights`; automatic measurement, `autoPageCharWeight`, `autoPageLineWeight`, and row text fragmentation remain the next specialty.
- Do not serialize auto-page metadata; each output page is an ordinary table with transform height equal to its included direct row-height sum.
- Do not split a rowspan region or let a rowspan cross the repeated-header/body boundary.
- Generated slides are new canonical blank slides inserted contiguously after the source; never reuse or modify a pre-existing following slide.
- Resolve internal-slide hyperlink targets before inserting slides so one-based input indices cannot drift.
- Keep source table creation, generated slides, relationships, order, sections, slide-number caches, model-cache cleanup, and runtime state failure-isolated under one outer transaction.
- Each task ends with focused review, commit, push, and `origin/main...HEAD` divergence `0 0`.

---

### Task 1: Auto-Page Option Normalization and Pure Partition Planner

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`
- Create: `packages/model/src/table-auto-page.internal.ts`
- Create: `packages/model/src/table-auto-page.internal.test.ts`

**Interfaces:**
- Consumes: `NormalizedTableDefinition`, normalized physical rows, exact `rowHeights`, `SlideSize`, and optional runtime layout margins.
- Produces:

```ts
export interface NormalizedTableAutoPageRequest {
  readonly repeatHeader: boolean;
  readonly headerRows: number;
  readonly slideStartY?: number;
  readonly slideMargin?: readonly [number, number, number, number];
}

export interface TableAutoPageMargins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

export function planTableAutoPages(
  definition: Readonly<NormalizedTableDefinition>,
  slideSize: Readonly<SlideSize>,
  layoutMargins?: Readonly<TableAutoPageMargins>,
): readonly Readonly<NormalizedTableDefinition>[];
```

- `NormalizedTableDefinition.autoPage` is absent when disabled and holds `NormalizedTableAutoPageRequest` when enabled.
- `NormalizedTableCell` becomes an exported internal interface so the partitioner can inspect `rowspan` without widening public exports.

- [ ] **Step 1: Write failing normalization tests**

Add exact cases to `table-create.internal.test.ts`:

```ts
it('normalizes detached strict auto-page controls', () => {
  const margin = [100, 200, 300, 400] as const;
  const definition = normalizeTableDefinition(
    [['H'], ['B']],
    {
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 1,
      autoPageSlideStartY: 500,
      slideMargin: margin,
      rowHeights: [40, 60],
    },
  );
  expect(definition.autoPage).toEqual({
    repeatHeader: true,
    headerRows: 1,
    slideStartY: 500,
    slideMargin: [100, 200, 300, 400],
  });
  expect(Object.isFrozen(definition.autoPage)).toBe(true);
  expect(Object.isFrozen(definition.autoPage!.slideMargin)).toBe(true);
});

it.each([
  [{ autoPage: 'yes' }, TypeError],
  [{ autoPage: false, autoPageRepeatHeader: true }, TypeError],
  [{ autoPage: true, autoPageHeaderRows: 1 }, TypeError],
  [{ autoPage: true, autoPageRepeatHeader: true, autoPageHeaderRows: 0 }, RangeError],
  [{ autoPage: true, rowHeights: [0, 10] }, RangeError],
  [{ autoPage: true, placeholder: 'Body', rowHeights: [10, 10] }, TypeError],
])('rejects invalid auto-page input before package access', (options, error) => {
  expect(() => normalizeTableDefinition([['A'], ['B']], options)).toThrow(error);
});
```

Also assert default header count 1, repeat false behavior, scalar margin broadcast, exact tuple length, sparse/accessor/symbol/class inputs, safe-integer rounding policy, caller detachment, unknown fields, row-count bounds, and control fields without `autoPage: true`.

- [ ] **Step 2: Run normalization tests and confirm the intended failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-create.internal.test.ts \
  -t "auto-page"
```

Expected: FAIL because the option reader rejects `autoPage` as unsupported.

- [ ] **Step 3: Implement strict detached normalization**

Extend `OPTION_KEYS`, export `NormalizedTableCell`, and add one small normalizer in `table-auto-page.internal.ts`. Use own data properties only, freeze the request and tuple, require `autoPage: true` for every related field, reject placeholder paging, require `autoRowHeight === false`, require all row heights positive, and leave `autoPage` absent when disabled.

The table definition return should preserve the existing mutable internal container shape and add only the detached request:

```ts
return {
  rows: physicalRows,
  ...existingGeometry,
  ...(autoPage === undefined ? {} : { autoPage }),
};
```

Do not alter the rendered table XML or public `AddTableOptions` yet.

- [ ] **Step 4: Write failing pure partition tests**

Create `table-auto-page.internal.test.ts` with a helper that calls `normalizeTableDefinition()` and `planTableAutoPages()`. Cover:

```ts
it('partitions exact EMU rows and repeats headers', () => {
  const pages = planTableAutoPages(definition({
    rows: [['H'], ['A'], ['B'], ['C']],
    y: 20,
    rowHeights: [20, 30, 30, 30],
    autoPageRepeatHeader: true,
    autoPageHeaderRows: 1,
    autoPageSlideStartY: 10,
    slideMargin: [0, 0, 10, 0],
  }), { width: 1000, height: 100 });
  expect(pages.map((page) => page.rows.map((row) => row[0]!.text)))
    .toEqual([['H', 'A'], ['H', 'B'], ['H', 'C']]);
  expect(pages.map(({ y, height }) => [y, height]))
    .toEqual([[20, 50], [10, 50], [10, 50]]);
});
```

Add exact-boundary fit, no overflow, header-only first page, repeat disabled, default 0.5-inch margins, layout-margin fallback, explicit margin precedence, startY precedence, first empty rejection, zero usable height, safe-sum overflow, oversized body block, all-header table, and frozen detached page definitions.

For merge coverage, use vertical, rectangular, adjacent, and nested-active rowspans. Assert each minimal row block stays intact; a merge crossing `headerRows` rejects; header-local and body-local merges retain their exact `rowspan`/continuation objects.

- [ ] **Step 5: Run partition tests and confirm the intended failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-auto-page.internal.test.ts
```

Expected: FAIL because `planTableAutoPages` does not exist.

- [ ] **Step 6: Implement the pure planner**

Implement helpers with no package access:

```ts
function rowBlocks(rows: readonly (readonly NormalizedTableCell[])[]):
  readonly { readonly start: number; readonly end: number }[];

function pageDefinition(
  source: Readonly<NormalizedTableDefinition>,
  rowIndexes: readonly number[],
  y: number,
): Readonly<NormalizedTableDefinition>;
```

Use safe integer addition for margins and row heights. Strip `autoPage` from returned page definitions, preserve all other fields, compute each page height from included rows, reuse normalized row/cell values without mutation, and return `Object.freeze(pages)`.

- [ ] **Step 7: Run focused/type gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-cell-merge.internal.test.ts
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-auto-page.internal.ts \
  packages/model/src/table-auto-page.internal.test.ts
git commit -m "feat: plan table auto pagination"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review strict input boundaries, no rendered XML changes, exact EMU arithmetic, merge-block completeness, immutable outputs, and absence of content measurement. Expected divergence: `0 0`.

### Task 2: Same-Layout Blank Slide Insertion Primitive

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: an attached source `SlideModel` with one safe direct internal slide-layout relationship.
- Produces:

```ts
export interface PreparedSlideInsertionAfter {
  readonly sourcePartUri: string;
  readonly sourceSlideId: number;
  readonly layoutPartUri: string;
  readonly materializeSlideNumber: boolean;
}

class PresentationModel {
  /** @internal */
  prepareSlideInsertionAfter(source: SlideModel): PreparedSlideInsertionAfter;
  /** @internal */
  insertPreparedBlankSlideAfter(
    after: SlideModel,
    prepared: Readonly<PreparedSlideInsertionAfter>,
  ): SlideModel;
  /** @internal */
  discardDetachedSlideModel(partUri: string): void;
}
```

- `attachSlide(slideUri, insertionIndex)` accepts an exact insertion index and writes the new `p:sldId` before that ordered element instead of appending.

- [ ] **Step 1: Write failing source/layout/order/section tests**

In `model.test.ts`, create three slides with two named layouts and sections. Prepare/insert after the middle slide and assert:

```ts
expect(model.slides.map(({ partUri }) => partUri)).toEqual([
  first.partUri,
  source.partUri,
  inserted.partUri,
  following.partUri,
]);
expect(inserted.relationships.filter(({ type }) =>
  type === SLIDE_LAYOUT_RELATIONSHIP).map(({ resolvedTarget }) => resolvedTarget))
  .toEqual([sourceLayoutPartUri]);
expect(model.sections?.find(({ id }) => id === sourceSectionId)?.slideIds)
  .toEqual([source.slideId, inserted.slideId]);
expect(model.slides[2]).toBe(inserted);
```

Verify layout placeholders and slide number materialize like `addSlide()`, following slides remain byte-identical, duplicate section titles do not affect exact membership, and section members follow presentation order.

- [ ] **Step 2: Write failing topology and rollback tests**

Add missing/repeated/external/wrong-content-type layout relationship cases and a detached source. Snapshot package bytes, relationships, mutations, slide/default-color caches, sections, and identities before each failure. Inject failure in `setPart()` after part allocation and require exact rollback plus explicit `discardDetachedSlideModel()` cleanup.

- [ ] **Step 3: Run focused tests and confirm the intended failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/model.test.ts \
  -t "prepared blank slide|slide insertion after"
```

Expected: FAIL because the insertion primitive is absent.

- [ ] **Step 4: Implement preparation and direct ordered insertion**

Factor only the canonical blank-slide commit shared with `addSlide()`:

```ts
private attachSlide(slideUri: string, insertionIndex = this.slides.length): SlideModel {
  // allocate relationship/id exactly as today
  // insert before ordered[insertionIndex] or append at end
}
```

`prepareSlideInsertionAfter()` performs all read-only source attachment, section, layout relationship, target content-type, and slide-number checks. `insertPreparedBlankSlideAfter()` revalidates the original source identity and the attached `after` slide, creates the blank part/layout relationship/placeholders, inserts immediately after `after`, copies exact membership from the original source slide ID, sorts section members, and synchronizes slide-number caches. Reusing one prepared source template with each newly generated `after` slide permits multi-page contiguous insertion without new topology guesses. Keep existing `addSlide()` and `duplicateSlide()` behavior unchanged.

- [ ] **Step 5: Run lifecycle gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/model.test.ts \
  -t "prepared blank slide|slide insertion after|slide lifecycle|sections|slide number"
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/presentation.ts packages/model/src/model.test.ts
git commit -m "feat: insert blank slides after source"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review no intermediate append state, same-layout ownership, exact section membership, placeholder/slide-number parity, cache cleanup, and unrelated slide isolation. Expected divergence: `0 0`.

### Task 3: Public AddTable Auto-Page Orchestration

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `planTableAutoPages()`, prepared hyperlink definitions, and prepared blank-slide insertion.
- Produces:

```ts
export type TableAutoPageMarginInput =
  | number
  | readonly [number, number, number, number];

export interface AddTableOptions {
  readonly autoPage?: boolean;
  readonly autoPageRepeatHeader?: boolean;
  readonly autoPageHeaderRows?: number;
  readonly autoPageSlideStartY?: number;
  readonly slideMargin?: TableAutoPageMarginInput;
}

class SlideModel {
  get newAutoPagedSlides(): readonly SlideModel[];
}
```

- A private `commitNormalizedTable(definition, preparedHyperlinks)` is the single table commit path for source and generated slides.

- [ ] **Step 1: Write failing public pagination tests**

Create a middle source slide followed by a sentinel slide. Add four explicit-height rows with repeat header and require:

```ts
const sourceTable = source.addTable(rows, {
  autoPage: true,
  autoPageRepeatHeader: true,
  autoPageHeaderRows: 1,
  autoPageSlideStartY: inches(0.5),
  slideMargin: inches(0.5),
  y: inches(5.5),
  rowHeights: [inches(0.5), inches(0.75), inches(0.75), inches(0.75)],
});
expect(sourceTable).toBeInstanceOf(TableModel);
expect(source.newAutoPagedSlides).toHaveLength(1);
expect(model.slides[model.slides.indexOf(source) + 1])
  .toBe(source.newAutoPagedSlides[0]);
expect(model.slides.at(-1)).toBe(sentinel);
expect(pageTables.map((table) => table.rows.map((row) => row.cells[0]!.text)))
  .toEqual([['Header', 'A'], ['Header', 'B', 'C']]);
```

Assert exact page transform heights, column widths, frozen getter array, no-overflow empty result, successful non-auto reset, move preservation, generated-slide deletion filtering, source duplication empty state, and reopen empty runtime state.

- [ ] **Step 2: Run public test and confirm the intended failure**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/model.test.ts \
  -t "auto-pages explicit table rows|newAutoPagedSlides"
```

Expected: FAIL because public options/getter and orchestration are absent.

- [ ] **Step 3: Refactor one normalized commit path**

Move the current owner resolution, shape ID allocation, relationship creation, render, insertion, model lookup, and error checks into:

```ts
private commitNormalizedTable(
  definition: Readonly<NormalizedTableDefinition>,
  prepared: readonly PreparedTableCellHyperlink[],
): TableModel;
```

The ordinary one-page path must remain byte-identical. Add an exact regression comparing package bytes and mutation journal for the same non-auto `addTable()` fixture before/after refactor.

- [ ] **Step 4: Implement outer orchestration and runtime state**

Before mutation:

1. normalize the complete table;
2. compute page definitions;
3. prepare slide insertion topology for every required continuation;
4. prepare hyperlinks for every page while original slide indices are unchanged.

Inside one outer transaction, commit the source page, insert each continuation after the previous page slide, and commit its page table. Track created part URIs for catch-time model-cache cleanup. Only after success replace the private frozen generated-part-URI list; the getter maps those URIs to currently attached stable slide models.

- [ ] **Step 5: Add six-format, layout, section, and reopen coverage**

For all six `PRESENTATION_FORMAT_PROFILES`, create at least two pages, write/reopen, and verify ordinary tables, row vectors, same layout relationships, section order, slide-number caches, headers, transforms, and zero retained auto-page OOXML metadata.

- [ ] **Step 6: Run gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  -t "auto-page|auto-paged|newAutoPagedSlides|table creation"
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
git add packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: auto-page tables across slides"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review ordinary-path byte stability, pre-insertion hyperlink resolution, result-state timing, generated slide order, stable identities, and six-format persistence. Expected divergence: `0 0`.

### Task 4: Merge, Relationship, and Failure-Isolation Hardening

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/table-auto-page.internal.ts`

**Interfaces:**
- Consumes: public auto-page fields and `newAutoPagedSlides` from Task 3.
- Produces: proven lifecycle behavior without a new public surface.

- [ ] **Step 1: Add merge/header boundary lifecycle tests**

Cover a header-local horizontal/vertical merge, a rectangular body merge that forces a whole row block to the next page, adjacent body merge blocks, hidden continuation text/style, and a rejected header/body-crossing rowspan. Inspect `mergeRegions` and raw `rowSpan/gridSpan/vMerge/hMerge` on every page and after write/reopen.

- [ ] **Step 2: Add relationship ownership tests**

Use header and body cells containing outer URL, local rich-run URL, and internal-slide hyperlinks. Put an internal target after the source so generated-slide insertion shifts its numeric index. Assert all page relationships still target the original part URI, each slide uses only its own `rId` values, repeated-header links exist on every page, tooltip/underline state survives, and no relationship is orphaned.

- [ ] **Step 3: Add comprehensive failure injection**

Snapshot parts, relationship parts, entry dates, presentation order, sections, slide model identities, shape collections, mutation journal, and an earlier successful `newAutoPagedSlides` result. Inject failures at source table `setPart`, generated slide part creation, layout relationship creation, placeholder materialization, generated table relationship creation, and final section/slide-number write. Require exact restoration and no cached detached generated slide.

- [ ] **Step 4: Add existing-following-slide and multiple-call tests**

Prove pre-existing following slides are never reused or modified, multiple auto-page calls insert independent contiguous groups, a later no-overflow/non-auto call resets only runtime result state, deleting one generated slide filters only that member, and moving source/generated/sentinel slides preserves stored model identities.

- [ ] **Step 5: Fix only observed defects and run regression gates**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-cell-merge.internal.test.ts \
  packages/model/src/table-cell-rich-text.internal.test.ts \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/model.test.ts \
  -t "auto-page|auto-paged|newAutoPagedSlides|table merge|table-cell hyperlink"
pnpm --config.verify-deps-before-run=false exec tsc \
  -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify synchronization**

```bash
git add packages/model/src/model.test.ts \
  packages/model/src/slide.ts packages/model/src/presentation.ts \
  packages/model/src/table-auto-page.internal.ts
git commit -m "test: harden table auto pagination"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Stage only files actually changed. Review merge atomicity, original hyperlink target identity, per-slide relationship ownership, rollback/cache isolation, and non-reuse of following slides. Expected divergence: `0 0`.

### Task 5: SDK Layout Margins, Root Types, and PptxGenJS Contracts

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: model default page margins and runtime SDK named-layout margin map.
- Produces:

```ts
class PresentationModel {
  /** @internal */
  tableAutoPageMarginsForSlide(
    slide: SlideModel,
  ): Readonly<TableAutoPageMargins> | undefined;
}

class PptxDocument extends PresentationModel {
  override tableAutoPageMarginsForSlide(
    slide: SlideModel,
  ): Readonly<TableAutoPageMargins> | undefined;
}
```

- Root/SDK declarations expose `TableAutoPageMarginInput`, the five `AddTableOptions` fields, and getter-only readonly `SlideModel.newAutoPagedSlides`.

- [ ] **Step 1: Write failing named-layout margin test**

Define a named layout with asymmetric EMU margin, add a slide from it, paginate without explicit `slideMargin`, and assert continuation Y/capacity uses runtime layout top/bottom. Then reopen, verify layout margin is absent as documented, and assert canonical 0.5-inch fallback until the margin is reapplied.

- [ ] **Step 2: Implement SDK override**

Resolve the source slide's unique internal layout relationship, look up its part URI in `#layoutMargins`, and return the frozen margin object. Base model returns `undefined`, allowing the pure planner to apply canonical fallback. Explicit table `slideMargin` remains higher priority.

- [ ] **Step 3: Add SDK/root positive and negative TypeScript contracts**

Compile examples using scalar/tuple margins, header controls, generated-slide readonly access, and all six formats. Add `@ts-expect-error` cases for assignment to `newAutoPagedSlides`, string booleans, malformed margin tuple, legacy aliases `addHeaderToEach`/`newSlideStartY`, and unsupported weights in this stage. Assert generated declarations contain exactly one getter and no setter.

- [ ] **Step 4: Add real PptxGenJS 4.0.1 public-output conformance**

Generate public fixtures for no overflow, multiple pages, one/two repeated header rows, continuation Y, rich cells, URL/internal links, and a source with a pre-existing following slide. Import the output and assert every emitted table is readable/editable and preserves final header/content/link state.

Create native explicit-height fixtures with the same final page table/header ordering where measurement is not under comparison. Document and test strict differences: native rejects invalid/coercible controls, does not mutate caller objects, does not clamp weights because weights are not yet accepted, and does not reuse a following slide.

- [ ] **Step 5: Run aggregate gates, review, commit, and push**

```bash
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "auto-page|auto-paged|newAutoPagedSlides|repeated header"
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false build
git diff --check
git add packages/model/src/presentation.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "test: verify table auto-page contracts"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review public names, readonly contract, runtime layout-margin precedence, generated declaration closure, PptxGenJS final-state evidence, and strict divergence. Expected divergence: `0 0`.

### Task 6: Documentation and Compatibility Matrix

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Document exact API and example**

Add a source-in-the-middle example using explicit `rowHeights`, `autoPage`, two repeated header rows, continuation Y, asymmetric slide margin, and `slide.newAutoPagedSlides`. Explain source/continuation geometry, same-layout insertion, section order, merge blocks, relationships, frozen runtime result, reset/reopen behavior, and strict input rules.

- [ ] **Step 2: Update support and remaining matrices**

Move explicit-row-height auto-page/repeated headers and `newAutoPagedSlides` to supported. Keep automatic row measurement, `autoPageCharWeight`, `autoPageLineWeight`, text-row fragmentation, placeholder auto-page, content/layout recomputation, and `tableToSlides` unsupported. Recalculate progress without claiming complete PptxGenJS parity.

- [ ] **Step 3: Run consistency and focused review**

```bash
rg -n "autoPage|auto-page|repeated header|newAutoPagedSlides|content measurement|tableToSlides" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
pnpm --config.verify-deps-before-run=false exec vitest run \
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "auto-page|auto-paged|newAutoPagedSlides|repeated header"
git diff --check
```

Review EMU units, explicit-height boundary, runtime-only state, PptxGenJS differences, and absence of premature actual-package/full-suite proof numbers.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: document table auto pagination"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 7: Actual Packed and Real-Browser Verification

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

- [ ] **Step 1: Extend actual-package declarations and Node smoke**

Assert installed Node/NodeNext types and negative contracts. From the actual tarball, create a source slide followed by a sentinel, paginate rich/linked/merged explicit-height rows with repeated headers, edit a generated-page cell, move/delete a generated slide, write/reopen, and emit stable `tableAutoPage` / `tableAutoPageState` booleans.

- [ ] **Step 2: Write and inspect a Node evidence deck**

Write `table-auto-page-smoke.pptx`. Run installed CLI/Inspector doctor, package inspect, PowerPoint 2010 validate, slides list, and exact source/generated/sentinel slide and relationship part reads. Require contiguous generated order, same layout targets, exact table row/height vectors, repeated header text/merge tokens, stable internal target part URIs, per-slide external relationships, and zero orphan links.

- [ ] **Step 3: Extend real-Chrome lifecycle**

Use the packed browser conditional export to perform the same create/paginate/edit/move/delete/writeBlob/reopen flow. Download `browser-table-auto-page.pptx`; require all state flags true and zero validation, console, page, and network errors.

- [ ] **Step 4: Build/pack twice and execute proof gates**

Compare two sorted 59-file dist SHA-256 manifests and two 62-file actual tarballs byte-for-byte, then run:

```bash
node --check scripts/smoke-npm-package.mjs
node --check scripts/playwright-browser-smoke.js
node scripts/smoke-npm-package.mjs \
  /tmp/pptx-table-auto-page-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Serve the extracted browser module over loopback, execute the checked-in callback in installed Google Chrome, and retain JSON/PPTX/manifest/tarball hashes outside the repository.

- [ ] **Step 5: Review, commit, push, and verify synchronization**

```bash
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js
git commit -m "test: verify packed table auto pagination"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Review actual-package-only imports, browser-safe code, same-layout evidence, header/merge/link ownership, deterministic artifacts, and zero runtime errors. Expected divergence: `0 0`.

### Task 8: Full Gates and Recorded Proof

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
  packages/model/src/table-auto-page.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/table-cell-merge.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  -t "auto-page|auto-paged|newAutoPagedSlides|repeated header"
pnpm --config.verify-deps-before-run=false test
pnpm --config.verify-deps-before-run=false test:performance
```

Record focused/full/performance file/test counts and durations from clean, non-overlapping runs.

- [ ] **Step 2: Re-run authoritative package/client inspection**

From the retained proof directory, verify both clean dist manifests, both actual tarballs, packed Node/NodeNext/browser/CLI/Inspector JSON, real-Chrome JSON, Node/browser deck hashes, exact part reads, and PowerPoint 2010 diagnostics. Do not infer broad proof from a narrow smoke flag.

- [ ] **Step 3: Record final evidence consistently**

In all six documents, add the same final test counts, performance timing, dist/tarball hashes, installed consumer results, Chrome version/errors, deck part/relationship/slide/table counts, layout/section/header/merge/link evidence, PowerPoint diagnostics, commit chain, proof path, specialty completion ratio, overall parity percentage, and next item `content measurement/layout recomputation`.

- [ ] **Step 4: Review documentation diff**

```bash
rg -n "tableAutoPage|auto-page|newAutoPagedSlides|actual tarball|Chrome|PowerPoint 2010" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git diff --check
git status --short
```

Review cross-file numbers, no stale “unsupported explicit auto-page” claims, no unsupported weight/measurement claims, and no unrelated tracked changes.

- [ ] **Step 5: Commit, push, and verify final synchronization**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: record table auto-page proof"
git push origin main
git rev-list --left-right --count origin/main...HEAD
git status --short
```

Expected divergence: `0 0`. The explicit-row-height auto-page specialty is complete only after all eight tasks pass and the retained evidence proves the full stated scope.

## Execution Selection

Execute inline in this thread, task by task, because the user delegated implementation choices and requested no confirmation prompts. After every task: review, commit, push, report completed/remaining/progress, then continue to the next task.
