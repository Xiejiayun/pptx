# Chart Creation and Semantic Editing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Create and semantically edit all nine PptxGenJS 4.0.1 standard chart types and compatible combination charts while keeping chart caches and embedded workbooks synchronized.

**Architecture:** Add strict chart definition/state modules to `@pptx/model`, generate deterministic XLSX payloads with the existing OPC/JSZip stack before synchronous presentation transactions, and render/patch only chart-owned XML spans. Reuse existing owned-dependency clone/delete behavior and make the advanced-charts plugin delegate standard data edits to the shared chart semantics.

**Tech Stack:** TypeScript 5.8, Vitest, lossless OOXML patching, OPC/JSZip, PptxGenJS 4.0.1 conformance, Playwright, pptx-inspect, PowerPoint, and LibreOffice.

## Global Constraints

- Execute inline in the current task; do not delegate implementation or review.
- Support `area`, `bar`, `bar3D`, `bubble`, `doughnut`, `line`, `pie`, `radar`, and `scatter` plus compatible combinations.
- Do not call or import PptxGenJS from production packages; it remains adapter/test evidence only.
- Create an embedded XLSX for every native chart and keep workbook cells, formulas, and chart caches synchronized.
- Generate workbook bytes before entering `OpcPackage.transaction()`; all presentation mutations remain synchronous and atomic.
- Preserve unowned chart XML, extensions, styles, and opaque dependencies; reject unsafe semantic edits before mutation.
- Keep `ChartModel.setXml()` as the explicit raw escape hatch and retain relationship-aware clone-on-write.
- Apply identical chart behavior to `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`.
- Keep modern `cx:*` charts, external workbooks, and chart animations preservation-only.
- Use fresh `/tmp/pptx-native-charts-*` directories for tarballs, generated fixtures, galleries, renders, and client round trips; never commit those artifacts or `.pnpm-store/`.
- End every task with focused tests, self-review, one commit, push, and local/remote divergence `0 0`.

---

## File Map

- Create `packages/model/src/chart.ts`: public chart catalog, series/definition/options/diagnostic types.
- Create `packages/model/src/chart-definition.internal.ts`: descriptor-safe normalization and deep-frozen snapshots.
- Create `packages/model/src/chart-definition.internal.test.ts`: input, detachment, range, combination, and no-accessor tests.
- Create `packages/model/src/chart-state.internal.ts`: strict standard-chart part and cache reader.
- Create `packages/model/src/chart-state.internal.test.ts`: categorical, scatter, bubble, combination, malformed, and no-mutation fixtures.
- Create `packages/model/src/chart-workbook.internal.ts`: deterministic worksheet column planning, XLSX rendering, and canonical readback.
- Create `packages/model/src/chart-workbook.internal.test.ts`: cell/formula/XML/determinism/invalid workbook tests.
- Create `packages/model/src/chart-render.internal.ts`: standard chart XML, axes, options, and slide graphic-frame rendering.
- Create `packages/model/src/chart-render.internal.test.ts`: exact group/series/cache/axis/frame output tests.
- Create `packages/model/src/chart-edit.internal.ts`: owned-span semantic replacement and chart/workbook relationship synchronization.
- Create `packages/model/src/chart-edit.internal.test.ts`: no-op, preservation, clone-on-write, and rollback tests.
- Modify `packages/model/src/shapes.ts`: upgrade `ChartModel` to the shared state and mutation APIs.
- Modify `packages/model/src/slide.ts`: async creation and chart deletion lifecycle.
- Modify `packages/model/src/index.ts`: export public chart types/catalog.
- Modify `packages/sdk/src/index.ts`: add the slide-index high-level chart creation overloads and write diagnostics.
- Modify `packages/model/src/dependency.internal.ts`: expose the existing owned-root collector needed by object deletion.
- Modify `plugins/advanced-charts/src/index.ts`: reuse shared standard-chart state and synchronized replacement.
- Extend model, SDK, adapter, root-package, packed Node, real-browser, validator, gallery, and public documentation evidence.

---

### Task 1: Strict chart definitions and read-only semantic state

**Files:**

- Create: `packages/model/src/chart.ts`
- Create: `packages/model/src/chart-definition.internal.ts`
- Create: `packages/model/src/chart-definition.internal.test.ts`
- Create: `packages/model/src/chart-state.internal.ts`
- Create: `packages/model/src/chart-state.internal.test.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/index.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: existing chart part bytes, chart-part relationships, `LosslessXmlDocument`, and `OpcPackage`.
- Produces:

```ts
export const CHART_TYPES = [
  'area', 'bar', 'bar3D', 'bubble', 'doughnut', 'line', 'pie', 'radar', 'scatter',
] as const;
export type ChartType = typeof CHART_TYPES[number];
export type ChartCategory = string | number;
export type ChartCategories = readonly ChartCategory[] | readonly (readonly string[])[];

export interface ChartSeriesInput {
  readonly name: string;
  readonly categories?: ChartCategories;
  readonly values: readonly number[];
  readonly xValues?: readonly number[];
  readonly sizes?: readonly number[];
}

export interface ChartSeries extends ChartSeriesInput {
  readonly categories?: ChartCategories;
  readonly values: readonly number[];
  readonly xValues?: readonly number[];
  readonly sizes?: readonly number[];
}

export interface ChartGroupInput {
  readonly type: ChartType;
  readonly series: readonly ChartSeriesInput[];
  readonly axis?: 'primary' | 'secondary';
}

export interface ChartDefinitionInput {
  readonly groups: readonly ChartGroupInput[];
}

export interface ChartDefinition {
  readonly groups: readonly Readonly<ChartGroupInput>[];
}

export type ChartStateStatus =
  | 'recognized'
  | 'cache-only'
  | 'modern'
  | 'unsupported'
  | 'ambiguous';

export interface ChartState {
  readonly status: ChartStateStatus;
  readonly definition?: Readonly<ChartDefinition>;
  readonly workbookPartUri?: string;
  readonly reason?: string;
}

export function normalizeChartDefinition(value: ChartDefinitionInput): Readonly<ChartDefinition>;
export function readChartState(pkg: OpcPackage, chartPartUri: string): Readonly<ChartState>;
```

- [ ] **Step 1: Add failing public catalog and normalization tests**

Create descriptor-safe fixtures and require exact type/catalog behavior, deep detachment, multi-level category shape,
scatter/bubble vectors, single-series pie/doughnut, and combination restrictions:

```ts
expect(CHART_TYPES).toEqual([
  'area', 'bar', 'bar3D', 'bubble', 'doughnut', 'line', 'pie', 'radar', 'scatter',
]);
expect(Object.isFrozen(CHART_TYPES)).toBe(true);

const input = {
  groups: [{
    type: 'bar',
    series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
  }],
} as const;
const normalized = normalizeChartDefinition(input);
expect(normalized.groups[0]?.series[0]).toEqual({
  name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
});
expect(Object.isFrozen(normalized.groups[0]?.series[0]?.values)).toBe(true);
```

Reject empty groups/series/names/vectors, unknown keys, sparse arrays, accessors, symbols, non-finite values, unequal
vectors, non-positive bubble sizes, categories on scatter/bubble, x/sizes on categorical charts, repeated secondary-
only groups, unsafe combinations, and extra series on pie/doughnut.

- [ ] **Step 2: Add failing strict chart-state fixtures**

Use independent compact XML fixture builders for one bar, one pie, one scatter, one bubble, and a bar+line combination.
Require cache order by `c:pt/@idx`, exact names/categories/values/X/sizes, axis assignment, workbook URI, frozen output,
and byte/journal stability:

```ts
const pkg = chartFixturePackage(BAR_CHART_XML, WORKBOOK_RELATIONSHIP_XML);
const before = packageSnapshot(pkg);
expect(readChartState(pkg, '/ppt/charts/chart1.xml')).toMatchObject({
  status: 'recognized',
  workbookPartUri: '/ppt/embeddings/workbook1.xlsx',
  definition: {
    groups: [{
      type: 'bar',
      axis: 'primary',
      series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
    }],
  },
});
expect(packageSnapshot(pkg)).toEqual(before);
```

Cover repeated roots/plot areas/group types/series containers/caches, wrong namespaces, missing/duplicate/negative/
decimal point indexes, point-count mismatch, non-finite numeric lexical values, malformed formulas, duplicate/dangling
axis IDs, wrong relationship type/mode, repeated external data, cache-only input, modern `cx:chart`, unknown standard
group type, and unrelated style/extension preservation.

- [ ] **Step 3: Run the focused tests and confirm red state**

```bash
pnpm exec vitest run packages/model/src/chart-definition.internal.test.ts \
  packages/model/src/chart-state.internal.test.ts
```

Expected: FAIL because the public chart module and strict state modules do not exist.

- [ ] **Step 4: Implement the public types and descriptor-safe normalizer**

Use existing model validators for ordinary/null-prototype objects, exact own-key sets, XML-safe strings, dense arrays,
and finite numbers. Normalize every series before checking type-specific vector rules:

```ts
export function normalizeChartDefinition(value: ChartDefinitionInput): Readonly<ChartDefinition> {
  const root = ownRecord(value, 'Chart definition', ['groups']);
  const groups = denseArray(root.groups, 'Chart groups').map((group, groupIndex) =>
    normalizeGroup(group, groupIndex));
  if (groups.length === 0) throw new RangeError('Chart definition requires at least one group');
  validateGroupCompatibility(groups);
  return deepFreeze({ groups });
}
```

Copy arrays and nested category levels; never retain caller objects or evaluate getters.

- [ ] **Step 5: Implement direct-structure state recognition**

Use direct-child helpers and namespace URIs, not broad local-name descendant ownership. Parse one exact cache at a time:

```ts
function directChildren(parent: XmlElement, namespaceUri: string, localName: string): readonly XmlElement[] {
  return parent.children.filter((child): child is XmlElement =>
    child.type === 'element' && child.namespaceUri === namespaceUri && child.localName === localName);
}

function readPoints(xml: LosslessXmlDocument, cache: XmlElement, numeric: boolean): readonly (string | number)[] {
  const points = directChildren(cache, CHART_NS, 'pt').map((point) => readPoint(xml, point, numeric));
  assertUniqueCanonicalIndexes(points);
  assertPointCount(xml, cache, points.length);
  return points.sort((left, right) => left.index - right.index).map(({ value }) => value);
}
```

Return a frozen status with a reason for every unsupported/ambiguous branch. Do not mutate or canonicalize imported XML.

- [ ] **Step 6: Route `ChartModel` through the strict state**

Replace the shallow `series` parser while retaining raw XML behavior:

```ts
get definition(): Readonly<ChartDefinition> | undefined {
  const uri = this.chartPartUri;
  return uri ? readChartState(this.slide.presentation.opcPackage, uri).definition : undefined;
}

get series(): readonly Readonly<ChartSeries>[] {
  return this.definition?.groups.flatMap(({ series }) => series) ?? EMPTY_CHART_SERIES;
}

get workbookPartUri(): string | undefined {
  const uri = this.chartPartUri;
  return uri ? readChartState(this.slide.presentation.opcPackage, uri).workbookPartUri : undefined;
}
```

Export `chart.ts` from `packages/model/src/index.ts` and remove the old local `ChartSeries` declaration/helper parser.

- [ ] **Step 7: Run focused and model regression tests**

```bash
pnpm exec vitest run packages/model/src/chart-definition.internal.test.ts \
  packages/model/src/chart-state.internal.test.ts packages/model/src/model.test.ts
pnpm typecheck
```

Expected: all tests pass; existing imported chart fixture still reports `Sales`, `Q1/Q2`, and `10/20`.

- [ ] **Step 8: Review, commit, and push Task 1**

```bash
git diff --check
git add packages/model/src/chart.ts packages/model/src/chart-definition.internal.ts \
  packages/model/src/chart-definition.internal.test.ts packages/model/src/chart-state.internal.ts \
  packages/model/src/chart-state.internal.test.ts packages/model/src/shapes.ts \
  packages/model/src/index.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: read strict chart semantics"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 2: Deterministic workbook plan and XLSX generation

**Files:**

- Create: `packages/model/src/chart-workbook.internal.ts`
- Create: `packages/model/src/chart-workbook.internal.test.ts`
- Modify: `packages/opc/src/index.ts`
- Modify: `packages/opc/src/index.test.ts`

**Interfaces:**

```ts
export interface ChartWorkbookPlan {
  readonly worksheetXml: string;
  readonly formulas: readonly Readonly<{
    groupIndex: number;
    seriesIndex: number;
    name: string;
    categories?: readonly string[];
    values: string;
    xValues?: string;
    sizes?: string;
  }>[];
}

export function planChartWorkbook(definition: Readonly<ChartDefinition>): Readonly<ChartWorkbookPlan>;
export function buildChartWorkbook(definition: Readonly<ChartDefinition>): Promise<Uint8Array>;
export function readChartWorkbookCells(bytes: Uint8Array): Promise<readonly (readonly (string | number)[])[]>;
export function chartWorkbookMatches(
  bytes: Uint8Array,
  definition: Readonly<ChartDefinition>,
): Promise<boolean>;
```

- [ ] **Step 1: Add a fixed-entry-date option to new OPC packages**

Write a failing OPC test requiring two packages created with the same date and parts to produce identical bytes:

```ts
const first = OpcPackage.create({ entryDate: new Date('1980-01-01T00:00:00Z') });
first.setPart('/data.xml', '<data/>', 'application/xml');
const second = OpcPackage.create({ entryDate: new Date('1980-01-01T00:00:00Z') });
second.setPart('/data.xml', '<data/>', 'application/xml');
expect(await first.write()).toEqual(await second.write());
```

Implement `OpcPackage.create(options?: { readonly entryDate?: Date })`, copy/validate the date, store it privately, and
pass it to every JSZip `file()` call made by a newly created package. Existing `create()` and opened packages retain
current behavior.

- [ ] **Step 2: Add failing workbook planner tests**

Require exact A1 ranges and sheet cells for categorical, multi-level, scatter, bubble, and combination definitions:

```ts
const plan = planChartWorkbook(normalizeChartDefinition({ groups: [{
  type: 'bar',
  series: [
    { name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] },
    { name: 'Cost', categories: ['Q1', 'Q2'], values: [7, 12] },
  ],
}] }));
expect(plan.formulas).toEqual([
  { groupIndex: 0, seriesIndex: 0, name: 'Sheet1!$B$1', categories: ['Sheet1!$A$2:$A$3'], values: 'Sheet1!$B$2:$B$3' },
  { groupIndex: 0, seriesIndex: 1, name: 'Sheet1!$C$1', categories: ['Sheet1!$A$2:$A$3'], values: 'Sheet1!$C$2:$C$3' },
]);
```

Require XML escaping, numeric zero preservation, independent X vectors, positive sizes, Excel columns beyond Z, and
stable group/series ordering.

- [ ] **Step 3: Add failing XLSX generation/readback tests**

Build twice and require byte identity, valid content types/relationships, one `Sheet1`, inline strings, no timestamps,
no shared strings, exact cell readback, and a successful definition match:

```ts
const first = await buildChartWorkbook(definition);
const second = await buildChartWorkbook(definition);
expect(second).toEqual(first);
expect(await chartWorkbookMatches(first, definition)).toBe(true);
expect(await readChartWorkbookCells(first)).toEqual(expectedCells);
const workbook = await OpcPackage.open(first);
expect(workbook.hasPart('/xl/worksheets/sheet1.xml')).toBe(true);
expect(workbook.hasPart('/xl/sharedStrings.xml')).toBe(false);
```

- [ ] **Step 4: Implement column planning and workbook XML**

Use one formula plan for both chart rendering and worksheet rendering. Render category levels first, then series values;
use X/Y and X/Y/size groups for scatter/bubble. Build the XLSX with fixed `1980-01-01T00:00:00Z` ZIP dates:

```ts
const workbook = OpcPackage.create({ entryDate: FIXED_ZIP_DATE });
workbook.setPart('/xl/workbook.xml', WORKBOOK_XML, WORKBOOK_CONTENT_TYPE);
workbook.setPart('/xl/worksheets/sheet1.xml', plan.worksheetXml, WORKSHEET_CONTENT_TYPE);
workbook.setPart('/xl/styles.xml', STYLES_XML, STYLES_CONTENT_TYPE);
workbook.addRelationship('/', { type: OFFICE_DOCUMENT_REL, target: 'xl/workbook.xml' });
workbook.addRelationship('/xl/workbook.xml', { id: 'rId1', type: WORKSHEET_REL, target: 'worksheets/sheet1.xml' });
workbook.addRelationship('/xl/workbook.xml', { id: 'rId2', type: STYLES_REL, target: 'styles.xml' });
return workbook.write();
```

Readback accepts only the generated direct workbook/sheet grammar and rejects formulas, merged cells, shared strings,
external links, or missing cells that cannot prove equality. `chartWorkbookMatches()` renders the expected cell matrix
from the normalized definition and compares exact typed cell values; it does not infer chart types from a worksheet.

- [ ] **Step 5: Run focused tests, review, commit, and push Task 2**

```bash
pnpm exec vitest run packages/opc/src/index.test.ts packages/model/src/chart-workbook.internal.test.ts
pnpm typecheck
git diff --check
git add packages/opc/src/index.ts packages/opc/src/index.test.ts \
  packages/model/src/chart-workbook.internal.ts packages/model/src/chart-workbook.internal.test.ts
git commit -m "feat: build deterministic chart workbooks"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

---

### Task 3: Categorical and axis-free chart creation

**Files:**

- Create: `packages/model/src/chart-render.internal.ts`
- Create: `packages/model/src/chart-render.internal.test.ts`
- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
export interface AddChartOptions {
  readonly name?: string;
  readonly altText?: string;
  readonly x?: Emu;
  readonly y?: Emu;
  readonly width?: Emu;
  readonly height?: Emu;
  readonly rotation?: OoxmlAngle;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
}

export interface NormalizedAddChartOptions {
  readonly name: string;
  readonly altText?: string;
  readonly x: Emu;
  readonly y: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation: OoxmlAngle;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

export function renderChartPart(
  definition: Readonly<ChartDefinition>,
  formulas: Readonly<ChartWorkbookPlan['formulas']>,
  workbookRelationshipId: string,
): string;

export function renderChartGraphicFrame(
  shapeId: number,
  relationshipId: string,
  options: Readonly<NormalizedAddChartOptions>,
): string;

PptxDocument.prototype.addChart(
  slideIndex: number,
  type: ChartType,
  series: readonly ChartSeriesInput[],
  options?: AddChartOptions,
): Promise<ChartModel>;
```

- [ ] **Step 1: Add failing exact XML tests for seven categorical/axis-free types**

Require `areaChart`, `barChart`, `bar3DChart`, `doughnutChart`, `lineChart`, `pieChart`, and `radarChart`; exact series
idx/order, name/category/value formulas and caches; axes only where required; and `c:externalData` using the passed ID.

```ts
const xml = renderChartPart(barDefinition, plan.formulas, 'rId1');
expect(xml).toContain('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>');
expect(xml).toContain('<c:f>Sheet1!$A$2:$A$3</c:f>');
expect(xml).toContain('<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>');
```

- [ ] **Step 2: Add failing slide creation and rollback tests**

For each of seven types, await `slide.addChart(type, series, options)`, require one chart part/workbook/frame, exact
relationship roles, stable live identity, placement, shape ordering before `p:extLst`, and reopen semantics. Inject
failures at workbook part, chart part, relationship, slide XML, and outer transaction boundaries and require the full
package snapshot unchanged.

- [ ] **Step 3: Implement canonical rendering and async creation**

Normalize and build the workbook before the transaction, then create all presentation resources atomically:

```ts
async addChart(type: ChartType, series: readonly ChartSeriesInput[], options: AddChartOptions = {}): Promise<ChartModel> {
  const definition = normalizeChartDefinition({ groups: [{ type, series }] });
  const normalizedOptions = normalizeAddChartOptions(options);
  const workbookBytes = await buildChartWorkbook(definition);
  const plan = planChartWorkbook(definition);
  return this.presentation.opcPackage.transaction(() =>
    this.addResolvedChart(definition, plan, workbookBytes, normalizedOptions));
}
```

Allocate `/ppt/charts/chartN.xml` and `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`, create chart/workbook and
slide/chart relationships, render the frame, and resolve the new `ChartModel` before commit.

- [ ] **Step 4: Verify six formats and public SDK access**

In model/SDK tests, create and reopen each type in all six presentation formats. Require `definition`, flattened
`series`, workbook URI/content type, transforms, name/alt text, zero chart diagnostics, and no orphan owned parts.

- [ ] **Step 5: Run gates, review, commit, and push Task 3**

```bash
pnpm exec vitest run packages/model/src/chart-render.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/chart.ts packages/model/src/chart-render.internal.ts \
  packages/model/src/chart-render.internal.test.ts packages/model/src/slide.ts \
  packages/model/src/shapes.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git commit -m "feat: create categorical native charts"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 4: Scatter and bubble chart creation

**Files:**

- Modify: `packages/model/src/chart-render.internal.ts`
- Modify: `packages/model/src/chart-render.internal.test.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/chart-state.internal.test.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:** Existing `addChart('scatter' | 'bubble', ...)`, `ChartSeriesInput.xValues`, and `sizes` become end-to-end.

- [ ] **Step 1: Add failing exact scatter and bubble tests**

Require independent per-series X/Y formulas and bubble X/Y/size formulas, exact numeric caches, zero preservation,
bubble size positivity, and value/value axes:

```ts
expect(scatterXml).toContain('<c:scatterChart>');
expect(scatterXml).toContain('<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>');
expect(scatterXml).toContain('<c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>');
expect(bubbleXml).toContain('<c:bubbleSize><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f>');
```

- [ ] **Step 2: Implement XY rendering and strict readback**

Render direct `xVal`, `yVal`, and `bubbleSize` containers from the workbook plan. Use two value axes for scatter and
bubble; do not reuse categorical-axis rendering. Extend strict state recognition to round-trip all vectors exactly.

- [ ] **Step 3: Verify six-format lifecycle and invalid isolation**

Create multiple scatter/bubble series with different X vectors, duplicate, write/reopen, and assert workbook/cache
equality. Reject NaN/infinity, unequal lengths, empty vectors, zero/negative sizes, and categorical fields without
consuming IDs or changing package bytes.

- [ ] **Step 4: Run gates, review, commit, and push Task 4**

```bash
pnpm exec vitest run packages/model/src/chart-render.internal.test.ts \
  packages/model/src/chart-state.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/chart-render.internal.ts packages/model/src/chart-render.internal.test.ts \
  packages/model/src/chart-state.internal.ts packages/model/src/chart-state.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git commit -m "feat: create native scatter and bubble charts"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 5: Combination charts and secondary axes

**Files:**

- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/chart-definition.internal.ts`
- Modify: `packages/model/src/chart-definition.internal.test.ts`
- Modify: `packages/model/src/chart-render.internal.ts`
- Modify: `packages/model/src/chart-render.internal.test.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
async addChart(groups: readonly ChartGroupInput[], options?: AddChartOptions): Promise<ChartModel>;
async addChart(slideIndex: number, groups: readonly ChartGroupInput[], options?: AddChartOptions): Promise<ChartModel>;
```

- [ ] **Step 1: Add failing compatibility and axis-allocation tests**

Accept bar+line, area+line, multiple same-type groups, primary-only, and primary+secondary groups. Reject empty groups,
pie/doughnut/bar3D/bubble combinations, scatter mixed with non-scatter, and secondary-only definitions. Require four
deterministic axis IDs when secondary is used and exact group references.

- [ ] **Step 2: Implement overload normalization and combo rendering**

Convert a single type/series call and a group-array call to the same `ChartDefinition`. Allocate axes after groups are
normalized, reuse primary IDs across primary groups, and emit secondary category/value axes only when referenced.

- [ ] **Step 3: Verify workbook/group order and reopen semantics**

Require worksheet columns, formula plans, plot group elements, flattened `ChartModel.series`, and `definition.groups`
to retain caller order. Duplicate and reopen all six formats without axis or relationship drift.

- [ ] **Step 4: Run gates, review, commit, and push Task 5**

```bash
pnpm exec vitest run packages/model/src/chart-definition.internal.test.ts \
  packages/model/src/chart-render.internal.test.ts packages/model/src/chart-state.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/chart.ts packages/model/src/chart-definition.internal.ts \
  packages/model/src/chart-definition.internal.test.ts packages/model/src/chart-render.internal.ts \
  packages/model/src/chart-render.internal.test.ts packages/model/src/chart-state.internal.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git commit -m "feat: create native combination charts"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 6: Atomic semantic data and type replacement

**Files:**

- Create: `packages/model/src/chart-edit.internal.ts`
- Create: `packages/model/src/chart-edit.internal.test.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
export function replaceChartDefinition(
  pkg: OpcPackage,
  slide: SlideModel,
  shapeId: number,
  current: Readonly<ChartState>,
  next: Readonly<ChartDefinition>,
  workbookBytes: Uint8Array,
): void;

ChartModel.prototype.replaceDefinition(value: ChartDefinitionInput): Promise<this>;
ChartModel.prototype.replaceSeries(value: readonly ChartSeriesInput[]): Promise<this>;
```

- [ ] **Step 1: Add failing no-op and synchronized replacement tests**

Require same-definition exact byte/journal no-op; name/category/value/X/size changes update cache, formulas, and XLSX;
single-type category conversions preserve unowned chart bytes; axis-free/XY/combo conversions replace only supported
plot/axis spans; and reopened state equals the requested definition.

- [ ] **Step 2: Add failing clone-on-write and rollback tests**

Create imported shared chart/workbook targets, edit one handle, and require a private chart+workbook subgraph while the
peer remains byte-identical. Inject failures after clone allocation, workbook write, chart write, relationship retarget,
slide XML write, validation, and outer transaction; require full package and object identity rollback.

- [ ] **Step 3: Implement preflight, workbook generation, and transactional patching**

Return before workbook generation for proven synchronized no-ops. Otherwise build bytes, re-resolve state inside the
transaction, clone owned dependencies when shared, patch group/axis/data spans, replace or add the workbook, and
retarget only the selected frame relationship.

```ts
async replaceDefinition(value: ChartDefinitionInput): Promise<this> {
  const next = normalizeChartDefinition(value);
  const current = this.requireEditableState();
  if (chartDefinitionsEqual(current.definition, next) && current.status === 'recognized') return this;
  const workbook = await buildChartWorkbook(next);
  this.slide.replaceChartDefinition(this.id, current, next, workbook);
  return this;
}
```

- [ ] **Step 4: Verify imported PptxGenJS normalization**

Open PptxGenJS bar/scatter/bubble/combo output, perform one semantic edit, and require the chart style/extensions outside
owned spans unchanged while workbook/cache/formulas become canonical and synchronized.

- [ ] **Step 5: Run gates, review, commit, and push Task 6**

```bash
pnpm exec vitest run packages/model/src/chart-edit.internal.test.ts \
  packages/model/src/chart-state.internal.test.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/chart-edit.internal.ts packages/model/src/chart-edit.internal.test.ts \
  packages/model/src/chart-state.internal.ts packages/model/src/shapes.ts packages/model/src/slide.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "feat: synchronize chart data and workbooks"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 7: Chart options and semantic option replacement

**Files:**

- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/chart-definition.internal.ts`
- Modify: `packages/model/src/chart-definition.internal.test.ts`
- Modify: `packages/model/src/chart-render.internal.ts`
- Modify: `packages/model/src/chart-render.internal.test.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/chart-edit.internal.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

Add strict public `ChartOptions`, `ChartGroupOptions`, `ChartAxisOptions`, `ChartTitleOptions`, `ChartLegendOptions`,
`ChartDataLabelOptions`, `ChartDataTableOptions`, and type-specific option unions exactly as defined in the design.

- [ ] **Step 1: Add failing option normalization matrices**

For every field, cover omitted/default, explicit zero/false, boundary values, detached colors/lines/fills, unknown keys,
accessors, wrong type, non-finite/range overflow, and incompatible type/group combinations. Require canonical defaults
for frame, title, legend, data labels, axes, table, colors, grouping, gap/overlap, hole/angle, markers, radar, and 3D.

- [ ] **Step 2: Add failing exact option XML/readback tests**

Render representative option-rich bar, bar3D, doughnut, line, pie, radar, scatter, bubble, and combo charts. Require
direct title/legend/axis/gridline/label/data-table/chart-area/plot-area/series/type-specific state and strict readback.

- [ ] **Step 3: Implement option rendering and owned-state projection**

Reuse existing strict rich color, simple fill, and simple line normalizers. Render only schema-legal children in fixed
order. State projection returns supported direct option values while preserving unexposed extensions and rich content.

- [ ] **Step 4: Extend `replaceDefinition()` to supported options**

Whole replacement updates supported option spans and frame metadata/transform, preserves unowned chart children, and
returns exact no-op for equal normalized options plus synchronized data. Workbook bytes remain unchanged for option-only
edits.

- [ ] **Step 5: Lock PptxGenJS valid-option conformance**

Compare public output for representative chart area/plot area, colors, title, legend, axis, labels, data table,
grouping, line, pie/doughnut, radar, and 3D options. Compare final semantics and formulas, not random UUID/timestamps or
redundant XML.

- [ ] **Step 6: Run gates, review, commit, and push Task 7**

```bash
pnpm exec vitest run packages/model/src/chart-definition.internal.test.ts \
  packages/model/src/chart-render.internal.test.ts packages/model/src/chart-edit.internal.test.ts \
  packages/model/src/model.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/chart.ts packages/model/src/chart-definition.internal.ts \
  packages/model/src/chart-definition.internal.test.ts packages/model/src/chart-render.internal.ts \
  packages/model/src/chart-render.internal.test.ts packages/model/src/chart-state.internal.ts \
  packages/model/src/chart-edit.internal.ts packages/model/src/model.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "feat: manage native chart options"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 8: Deletion, diagnostics, validation, and plugin integration

**Files:**

- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/dependency.internal.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/validator/src/index.ts`
- Modify: `plugins/advanced-charts/src/index.ts`
- Modify: `plugins/advanced-charts/src/index.test.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**

```ts
ChartModel.prototype.remove(): void;
SlideModel.prototype.deleteChart(shapeId: number): void;
export function chartDiagnostics(
  pkg: OpcPackage,
  slidePartUri: string,
): Promise<readonly ChartDiagnostic[]>;
```

- [ ] **Step 1: Add failing object deletion and GC tests**

Delete exclusive, duplicated, and imported shared charts. Require frame/relationship removal, recursive chart/workbook/
style/color collection only after the final incoming reference, preservation of shared images/themes/external targets,
no ID renumbering, stable peer identity, six-format reopen, and complete rollback on injected failure.

- [ ] **Step 2: Add failing diagnostic and strict-write tests**

Require precise codes for invalid relationships, unsupported/ambiguous structures, invalid caches/axes, missing
workbooks, divergence, and modern charts. `PptxDocument.write()` must reject error diagnostics in strict mode and retain
warnings/info in permissive mode without mutation.

- [ ] **Step 3: Implement deletion and write-time diagnostics**

Resolve the entire chart relationship root without inferring chart internals, remove the frame and relationship in one
transaction, then call owned-root garbage collection. Await per-slide chart diagnostics during SDK write alongside
the synchronous media and gradient checks.

- [ ] **Step 4: Integrate the advanced-charts plugin**

Replace its descendant-wide standard-series decoder with the shared strict state. Make `setSeriesValues()` call the
synchronized core replacement path; keep trendline/error-bar/data-label appends and modern inspection. Require raw XML
edits to report divergence until a semantic replacement or explicit workbook replacement proves equality.

- [ ] **Step 5: Run gates, review, commit, and push Task 8**

```bash
pnpm exec vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/validator/src/index.test.ts plugins/advanced-charts/src/index.test.ts
pnpm typecheck
git diff --check
git add packages/model/src/chart.ts packages/model/src/shapes.ts packages/model/src/slide.ts \
  packages/model/src/dependency.internal.ts packages/model/src/model.test.ts \
  packages/sdk/src/index.ts packages/sdk/src/index.test.ts packages/validator/src/index.ts \
  plugins/advanced-charts/src/index.ts plugins/advanced-charts/src/index.test.ts
git commit -m "feat: manage native chart lifecycle"
git push origin main
git rev-list --left-right --count HEAD...origin/main
```

---

### Task 9: PptxGenJS, packed/browser/client evidence, and documentation

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/plugins.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/compatibility/cross-client-testing.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:** Complete chart behavior through only the actual packed `@jiayunxie/pptx` package.

- [ ] **Step 1: Complete PptxGenJS public-output conformance**

Generate all nine valid public chart types plus bar+line primary/secondary combo with PptxGenJS 4.0.1. Import, project,
edit, write, reopen, and compare type/groups/series/categories/X/Y/sizes/formulas/workbook cells/relationships/options.
Record intentional strict corrections for invalid runtime passthrough without copying defects.

- [ ] **Step 2: Extend root, packed Node, declaration, and browser smoke**

Create/edit/type-convert/combine/duplicate/delete/reopen charts from the installed tarball and real browser bundle.
Return top-level `nativeCharts: true` only when nine types, workbook/cache equality, unique IDs, relationship isolation,
zero chart errors, and reopen all hold. Add compile-time public API positive/negative cases and static browser-import scan.

- [ ] **Step 3: Run final repository and package gates**

```bash
pnpm exec vitest run packages/model/src/chart-definition.internal.test.ts \
  packages/model/src/chart-state.internal.test.ts packages/model/src/chart-workbook.internal.test.ts \
  packages/model/src/chart-render.internal.test.ts packages/model/src/chart-edit.internal.test.ts \
  packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts packages/pptx/src/index.test.ts \
  plugins/advanced-charts/src/index.test.ts
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
pnpm --filter @jiayunxie/pptx build
git diff --check
```

Expected: every focused/full/type/performance/build gate passes.

- [ ] **Step 4: Pack and run installed Node/real-Chrome/CLI smoke**

```bash
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-native-charts-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-native-charts-pack/jiayunxie-pptx-0.1.0.tgz
```

Serve the built browser bundle locally and execute `scripts/playwright-browser-smoke.js` in installed Google Chrome.
Require both outputs to contain top-level `nativeCharts: true`.

- [ ] **Step 5: Generate and inspect the real chart gallery**

Through only the installed tarball, generate type, option, combo, data-edit, type-conversion, duplicate/delete, and
PptxGenJS-adoption slides with real business data. Verify exact chart/workbook part counts, relationships, cache/cell
equality, formulas, unique IDs, diagnostics, and zero orphan owned parts using `pptx-inspect` plus package scripts.

- [ ] **Step 6: Render and run cross-client evidence**

Render every slide at 180 DPI, run overflow detection, inspect each page at full size, and correct unintended overlap,
clipping, wrapping, or chart/data mismatch. Open/save/reopen in available PowerPoint and LibreOffice clients; compare
semantic data/options/workbooks/rendering and record client normalization or degradation exactly.

- [ ] **Step 7: Update public support and remaining-work documentation**

Document creation/editing examples, strict semantics, supported types/options, diagnostics, PptxGenJS conformance,
packed/browser/client results, and remaining modern/external/chart-animation gaps. Move standard chart creation and
semantic editing out of unsupported lists and set the next roadmap item to slide background/number/default color.

- [ ] **Step 8: Review, commit, push, and verify Task 9**

```bash
git diff --check
git add packages/pptxgenjs-adapter/src/index.test.ts packages/pptx/src/index.test.ts \
  scripts/build-npm-package-types.mjs scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js README.md packages/pptx/README.md docs/api/README.md \
  docs/plugins.md docs/compatibility/pptxgenjs-baseline.md \
  docs/compatibility/cross-client-testing.md docs/implementation-progress.md CHANGELOG.md
git diff --cached --check
git commit -m "docs: document native chart support"
git push origin main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected divergence: `0 0`; only `.pnpm-store/` remains untracked.

---

## Plan Self-Review Checklist

- Every design requirement maps to one or more tasks: types/state (1), workbook (2), standard types (3-4), combo (5),
  semantic edit (6), options (7), lifecycle/plugin/validation (8), and conformance/release evidence (9).
- Public names are consistent across tasks: `ChartDefinitionInput`, `ChartDefinition`, `ChartGroupInput`,
  `ChartSeriesInput`, `ChartModel.definition`, `replaceDefinition()`, `replaceSeries()`, and `remove()`.
- Every mutation path generates async workbook bytes before a synchronous presentation transaction.
- PptxGenJS is referenced only in adapter tests and evidence, never in production creation.
- No task claims modern chart creation, external workbook editing, or chart animations.
- Every task has a focused red/green test cycle, review gate, commit, push, and remote divergence check.
