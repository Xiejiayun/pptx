# Table Cell Text Direction Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user requires inline execution without subagents.

**Goal:** Let native `slide.addTable()` create cells with strict `horz`/`vert`/`vert270`/`wordArtVert` direction through `{ text, options: { textDirection } }`, matching PptxGenJS 4.0.1 creation wire semantics while preserving the existing direct editor.

**Architecture:** Reuse `TableCellTextDirection` and `normalizeTableCellTextDirection()`, carry an optional direction through the existing detached table definition, and add one narrow creation renderer for `tcPr@vert`. Keep explicit `horz` in the normalized value so it can later block a table-level default, but omit it from creation XML to match PptxGenJS final state.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span model, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice headless, Poppler.

## Global Constraints

- Public creation property is `AddTableCellOptions.textDirection?: TableCellTextDirection` with exact values `horz`, `vert`, `vert270`, and `wordArtVert`.
- Storage is the created physical cell's direct `a:tcPr@vert`; never write `a:bodyPr@vert`.
- Omitted/undefined and explicit `horz` create no direct direction attribute; the three non-horizontal values write their exact wire tokens.
- Explicit `horz` remains present in the normalized definition so the next table-level propagation item can distinguish it from omission.
- Attribute order is canonical margins, optional `anchor`, optional `vert`, then L/R/T/B borders and fill.
- Reuse the existing strict normalizer; do not add aliases, coercion, case folding, whitespace trimming, or the three text-box-only tokens.
- Existing `TableCell.textDirection` and `setCellTextDirection()` direct editing semantics remain unchanged; the editor may explicitly write `vert="horz"`.
- Table-level `textDirection`, fit creation, rich/multi-paragraph cells, hyperlink, merge, auto-page, and unrelated options remain outside this slice.
- Every successful task is reviewed, committed, pushed, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only gates do not create empty commits.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and render cell text direction

**Files:**
- Modify: `packages/model/src/table-cell-text-direction.internal.ts`
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: existing `TableCellTextDirection`, `normalizeTableCellTextDirection()`, table-cell margin/vertical-alignment/border/fill renderers.
- Produces:

```ts
export function renderTableCellTextDirectionAttribute(
  value: TableCellTextDirection | undefined,
): string;
```

- [ ] **Step 1: Add failing normalization and exact-render tests**

In `table-create.internal.test.ts`, add next to cell vertical-alignment creation:

```ts
it('normalizes and renders strict table cell text direction', () => {
  const nullOptions = Object.assign(Object.create(null), {
    textDirection: 'wordArtVert',
  });
  const rows = [[
    'String',
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { textDirection: undefined } },
    { text: 'Horizontal', options: { textDirection: 'horz' } },
    { text: 'Vertical', options: { textDirection: 'vert' } },
    { text: 'Rotate 270', options: { textDirection: 'vert270' } },
    { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
    { text: 'Null prototype', options: nullOptions },
    { text: 'Combined', options: {
      align: 'center',
      textDirection: 'vert270',
      valign: 'middle',
      margin: { top: 4, left: 8 },
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
      },
    } },
  ]];
  const definition = normalizeTableDefinition(rows, undefined);
  expect(definition.rows[0]!.map(({ textDirection }) => textDirection)).toEqual([
    undefined,
    undefined,
    undefined,
    'horz',
    'vert',
    'vert270',
    'wordArtVert',
    'wordArtVert',
    'vert270',
  ]);

  const equivalent = [
    [['Same']],
    [[{ text: 'Same' }]],
    [[{ text: 'Same', options: {} }]],
    [[{ text: 'Same', options: { textDirection: undefined } }]],
    [[{ text: 'Same', options: { textDirection: 'horz' } }]],
  ].map((input) => renderTableGraphicFrame(
    50,
    normalizeTableDefinition(input, undefined),
  ));
  expect(new Set(equivalent).size).toBe(1);
  expect(equivalent[0]).not.toMatch(/<a:tcPr[^>]*\svert=/);

  const xml = renderTableGraphicFrame(51, definition);
  expect([...xml.matchAll(/<a:tcPr([^>]*)>/g)].map((match) =>
    match[1]!.match(/\svert="([^"]+)"/)?.[1])).toEqual([
    undefined,
    undefined,
    undefined,
    undefined,
    'vert',
    'vert270',
    'wordArtVert',
    'wordArtVert',
    'vert270',
  ]);
  expect(xml).toMatch(
    /marB="45720" anchor="ctr" vert="vert270"><a:lnL[\s\S]*<\/a:lnB><a:solidFill>/,
  );
  expect(xml).not.toMatch(/<a:bodyPr[^>]*\svert=/);
});
```

Extend the malformed-options test with a getter-free accessor and invalid values:

```ts
const accessorTextDirectionOptions = {};
Object.defineProperty(accessorTextDirectionOptions, 'textDirection', {
  get() {
    cellAccessorCalls += 1;
    return 'vert';
  },
  enumerable: true,
  configurable: true,
});

const invalidTextDirections = [
  null,
  false,
  true,
  0,
  '',
  'Vert',
  ' vert ',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
  [],
  {},
  Symbol('vert'),
];
```

Add `{ text: 'A', options: accessorTextDirectionOptions }` and each value wrapped as `{ text: 'A', options: { textDirection } }` to the existing invalid-cell matrix. Preserve `expect(cellAccessorCalls).toBe(0)`.

- [ ] **Step 2: Run focused tests and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "text direction|malformed matrices"
```

Expected: `textDirection` is rejected as an unsupported cell option or the normalized field is absent.

- [ ] **Step 3: Add the narrow renderer and normalized field**

In `table-cell-text-direction.internal.ts` add:

```ts
export function renderTableCellTextDirectionAttribute(
  value: TableCellTextDirection | undefined,
): string {
  return value === undefined || value === 'horz'
    ? ''
    : ` vert="${escapeXmlAttribute(value)}"`;
}
```

In `table-create.internal.ts`, extend imports and the normalized cell:

```ts
import {
  normalizeTableCellTextDirection,
  renderTableCellTextDirectionAttribute,
} from './table-cell-text-direction.internal.js';
import type {
  TableCellBorders,
  TableCellFill,
  TableCellTextDirection,
} from './shapes.js';

interface NormalizedTableCell {
  readonly text: string;
  readonly alignment?: TextAlignment;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
  readonly textDirection?: TableCellTextDirection;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}
```

Extend the supported keys and normalize only supplied values:

```ts
function normalizeTableCellOptions(
  value: unknown,
  context: string,
): Pick<
  NormalizedTableCell,
  | 'alignment'
  | 'borders'
  | 'fill'
  | 'margins'
  | 'textDirection'
  | 'verticalAlignment'
> {
  const options = readDataObject(
    value,
    `${context} options`,
    ['align', 'border', 'fill', 'margin', 'textDirection', 'valign'],
  );
  const textDirection = options.textDirection === undefined
    ? undefined
    : normalizeTableCellTextDirection(
        options.textDirection,
        `${context} textDirection`,
      );

  return {
    ...(alignment === undefined ? {} : { alignment }),
    ...(borders === undefined ? {} : { borders }),
    ...(fill === undefined ? {} : { fill }),
    ...(margins === undefined ? {} : { margins }),
    ...(textDirection === undefined ? {} : { textDirection }),
    ...(verticalAlignment === undefined ? {} : { verticalAlignment }),
  };
}
```

Render after vertical alignment and before child elements:

```ts
const textDirectionAttribute = renderTableCellTextDirectionAttribute(
  cell.textDirection,
);
return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody>`
  + `<a:tcPr${marginAttributes}${verticalAlignmentAttribute}${textDirectionAttribute}>`
  + `${borders}${fill}</a:tcPr></a:tc>`;
```

- [ ] **Step 4: Run focused regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: all tests pass; omitted/undefined/horz remain byte-identical and existing direct editing tests remain green.

- [ ] **Step 5: Review, commit, push, and verify synchronization**

Review exact attribute ownership/order, explicit-horz normalized retention, no aliases, getter-free rejection, and no public API change yet. Stage only the three files and commit:

```text
feat: render table cell text direction during creation
```

Push/fetch with the verified SSH-over-443 commands and require `origin/main...HEAD = 0 0`; `.pnpm-store/` remains untracked.

---

### Task 2: Expose the typed API and prove the native lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 normalized/rendered `textDirection` and the existing `TableCellTextDirection` snapshot/editor.
- Produces:

```ts
export interface AddTableCellOptions {
  readonly textDirection?: TableCellTextDirection;
}
```

- [ ] **Step 1: Add failing public type and lifecycle assertions**

In the existing strict table creation tests, type source options as `AddTableCellOptions` and add directions:

```ts
const sourceOptions: AddTableCellOptions = {
  align: 'left',
  border: sourceBorder,
  fill: sourceFill,
  margin: sourceMargin,
  textDirection: 'vert',
  valign: 'top',
};
```

Use all four values across the first row and assert public snapshots plus raw tokens:

```ts
expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
  'vert',
  'vert270',
  'wordArtVert',
]);
expect([...createdTableXml.matchAll(/<a:tcPr[^>]*(?:\svert="([^"]+)")?[^>]*>/g)]
  .map((match) => match[1])).toEqual(expect.arrayContaining([
    'vert',
    'vert270',
    'wordArtVert',
  ]));
```

Add a dedicated SDK lifecycle table with five cells:

```ts
const table = slide.addTable([[
  'Omitted',
  { text: 'Horizontal', options: { textDirection: 'horz' } },
  { text: 'Vertical', options: { textDirection: 'vert' } },
  { text: 'Rotate 270', options: { textDirection: 'vert270' } },
  { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
]], {
  name: 'Cell text direction creation',
  columnWidths: inches(1.5),
  rowHeights: inches(1),
  align: 'center',
  margin: { top: 4, left: 8 },
  valign: 'middle',
});
expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
  undefined,
  undefined,
  'vert',
  'vert270',
  'wordArtVert',
]);
```

Duplicate before mutations; edit source text/alignment/margins/border/fill/fit/column widths/row heights; use `setCellTextDirection(0, 0, 'horz')` to prove the existing editor writes a direct horizontal token; rollback one temporary created table; write/reopen and assert source/clone isolation, stable `TableModel` identity, non-target part hash, geometry, snapshots, and package validation.

Add strict public invalid values:

```ts
const invalidDirections = [
  null,
  false,
  true,
  0,
  '',
  'Vert',
  ' vert ',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
  [],
  {},
  Symbol('table cell text direction'),
];
for (const textDirection of invalidDirections) {
  expect(() => slide.addTable([[{
    text: 'Invalid',
    options: { textDirection } as never,
  }]])).toThrow(TypeError);
}
```

Compare slide bytes and mutation journal before/after the loop.

- [ ] **Step 2: Run typecheck/focused tests and confirm red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  -t "text direction creation|strict basic tables"
```

Expected: TypeScript rejects `AddTableCellOptions.textDirection`.

- [ ] **Step 3: Add the public property**

In `slide.ts`, import the existing type from `shapes.ts` and extend only the cell interface:

```ts
import type {
  TableCellBorderInput,
  TableCellFill,
  TableCellTextDirection,
} from './shapes.js';

export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}
```

Do not add `textDirection` to `AddTableOptions` in this task.

- [ ] **Step 4: Run public regressions**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
```

Expected: all suites pass, including duplicate/rollback/reopen and invalid-value zero mutation.

- [ ] **Step 5: Review, commit, push, and verify synchronization**

Review public naming, cell-only scope, declaration propagation, immediate snapshots, explicit-horz collapse, editor interoperability, rollback and non-target isolation. Stage only the three files and commit:

```text
feat: expose table cell text direction creation
```

Push/fetch and require `0 0`.

---

### Task 3: Prove PptxGenJS 4.0.1 final-state conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: Tasks 1–2 public native API and PptxGenJS 4.0.1 public `addTable()`/`write()` output.
- Produces: conformance evidence for cell-level omission, explicit horz, three direct non-horizontal values, and strict invalid divergence.

- [ ] **Step 1: Add a real PptxGenJS/native comparison**

Add next to the existing imported table direction test:

```ts
it('matches native table-cell text direction creation to PptxGenJS final state', async () => {
  const generatedRows = [[
    { text: 'Omitted', options: {} },
    { text: 'Horizontal', options: { textDirection: 'horz' } },
    { text: 'Vertical', options: { textDirection: 'vert' } },
    { text: 'Rotate 270', options: { textDirection: 'vert270' } },
    { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
  ]] satisfies PptxGenJS.TableRow[];
  const generated = new PptxGenJS();
  expect(generated.version).toBe('4.0.1');
  generated.layout = 'LAYOUT_WIDE';
  generated.addSlide().addTable(generatedRows, {
    x: 0.5,
    y: 0.5,
    w: 10,
    h: 1,
    colW: 2,
    rowH: 1,
    margin: 0.1,
    valign: 'middle',
  });
  const imported = await importPptxGenJS(generated);

  const native = PptxDocument.create({ slideSize: 'wide' });
  const nativeTable = native.addSlide().addTable([[
    { text: 'Omitted' },
    { text: 'Horizontal', options: { textDirection: 'horz' } },
    { text: 'Vertical', options: { textDirection: 'vert' } },
    { text: 'Rotate 270', options: { textDirection: 'vert270' } },
    { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
  ]], {
    x: inches(0.5),
    y: inches(0.5),
    width: inches(10),
    height: inches(1),
    columnWidths: inches(2),
    rowHeights: inches(1),
    margin: 7.2,
    valign: 'middle',
  });
  const importedTable = imported.slides[0]!.shapes[0] as TableModel;
  const expected = [undefined, undefined, 'vert', 'vert270', 'wordArtVert'];
  expect(nativeTable.rows[0]!.cells.map(({ textDirection }) => textDirection))
    .toEqual(expected);
  expect(importedTable.rows[0]!.cells.map(({ textDirection }) => textDirection))
    .toEqual(expected);
  expect(nativeTable.transform).toMatchObject(importedTable.transform);
  expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
  expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);
});
```

Decode both public outputs and extract exactly one cell at a time; assert direct token vector `[undefined, undefined, 'vert', 'vert270', 'wordArtVert']`, no `bodyPr@vert`, equal text/margins/anchors, and reopen stability.

Generate a separate PptxGenJS runtime-invalid cell with `textDirection: 'eaVert' as never`; assert public output contains `tcPr@vert="eaVert"` and imports as `undefined`. Native creation with the same runtime value must throw `TypeError` before mutation.

- [ ] **Step 2: Run adapter and related suites**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/sdk/src/index.test.ts
```

Expected: all pass; the comparison reads only public PptxGenJS output.

- [ ] **Step 3: Review, commit, push, and verify synchronization**

Review final-state comparison scope, explicit-horz collapse, invalid passthrough distinction, no private PptxGenJS fields, exact geometry and reopen. Stage only the adapter test and commit:

```text
test: compare table cell text direction creation with pptxgenjs
```

Push/fetch and require `0 0`.

---

### Task 4: Prove packed Node/browser/type surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: installed tarball only, aggregate `@jiayunxie/pptx` exports, browser conditional bundle and generated declarations.
- Produces: packed creation/reopen smoke flags for cell text direction.

- [ ] **Step 1: Extend the Node smoke matrix**

Add typed directions to the existing table cells:

```ts
{ text: 'Region', options: {
  align: 'left',
  border: creationBorder,
  fill: creationFill,
  margin: creationMargin,
  textDirection: 'vert',
  valign: 'top',
} }
```

Use `vert270` and `wordArtVert` in neighboring cells and a separate explicit `horz` cell. Capture initial and reopened vectors:

```ts
const initialCreatedDirections = createdTable.rows
  .map(({ cells }) => cells.map(({ textDirection }) => textDirection));
const tableCellTextDirectionCreation =
  JSON.stringify(initialCreatedDirections) === JSON.stringify([
    ['vert', 'vert270'],
    ['wordArtVert', undefined],
  ]);
```

Add `tableCellTextDirectionCreation` to the emitted JSON result and throw on a false flag. Add declaration-only assignments:

```ts
const typedCellDirection: AddTableCellOptions = { textDirection: 'vert270' };
void typedCellDirection;
```

Import `type AddTableCellOptions` from the installed package in the generated typecheck file.

- [ ] **Step 2: Extend the browser conditional-bundle smoke**

Mirror the four-value creation matrix in the browser script string. Assert initial values, explicit-horz absence, direct editor interoperability, and write/reopen values without Node globals.

- [ ] **Step 3: Build, pack, and run the actual package smoke**

```sh
pnpm build
PACK_DIR="$(mktemp -d)"
pnpm --filter @jiayunxie/pptx pack --pack-destination "$PACK_DIR"
node scripts/smoke-npm-package.mjs "$PACK_DIR"/*.tgz
```

Expected: Node/browser/type/CLI smoke succeeds and reports `tableCellTextDirectionCreation: true`.

- [ ] **Step 4: Review, commit, push, and verify synchronization**

Review installed-package-only imports, Node/browser parity, strict types, creation→edit→write→reopen, and unchanged CLI. Stage only the smoke script and commit:

```text
test: smoke packed table cell text direction creation
```

Push/fetch and require `0 0`.

---

### Task 5: Document the supported contract and remaining boundary

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: final implementation, tests and packed public names.
- Produces: accurate public examples, compatibility wording and unsupported list.

- [ ] **Step 1: Add the release note**

Add one Unreleased bullet:

```md
- Added strict table-cell text-direction creation for `horz`, `vert`, `vert270`, and `wordArtVert`, with PptxGenJS 4.0.1 horizontal-collapse and direct non-horizontal final-state conformance.
```

- [ ] **Step 2: Update API and package examples**

Change the documented cell shape to:

```ts
{
  text: string;
  options?: {
    align?: TextAlignment;
    border?: TableCellBorderInput;
    fill?: TableCellFill;
    margin?: TextBoxMarginInput;
    textDirection?: TableCellTextDirection;
    valign?: TextBoxVerticalAlignment;
  };
}
```

Document direct `tcPr@vert`, four exact values, strict input, omitted/undefined/horz creation collapse, three non-horizontal snapshots, direct editor explicit-horz distinction, detachment, lifecycle and table-level pending.

- [ ] **Step 3: Update the compatibility matrix precisely**

Change the table-cell direction row to:

```md
| table-cell `textDirection` | `AddTableCellOptions.textDirection` / `TableCell.textDirection` / `TableModel.setCellTextDirection()` | 已支持 cell-level 创建与 direct physical-cell 读取/编辑；table-level default 创建尚未支持 |
```

Remove direction from the unsupported cell-creation list, retain table-level direction, fit creation, rich/multi-paragraph, hyperlink, merge and layout items. State that valid native/PptxGenJS final states match, native rejects invalid passthrough, and creation `horz`/omitted both lack a direct token while the direct editor can write explicit `horz`.

- [ ] **Step 4: Run documentation, type and focused checks**

```sh
rg -n "direction/fit creation|cell options other than.*direction|textDirection.*unsupported" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
```

Expected: no stale claim says cell-level direction creation is unsupported; table-level remains visible.

- [ ] **Step 5: Review, commit, push, and verify synchronization**

Review public name, four values, horizontal collapse, OOXML ownership, editor distinction, PptxGenJS wording and remaining boundary. Stage only the four docs and commit:

```text
docs: document table cell text direction creation
```

Push/fetch and require `0 0`.

---

### Task 6: Final regression and real-deck QA

**Files:**
- Verify only; create temporary artifacts under `/tmp`, never in the repository.

**Interfaces:**
- Consumes: implementation, tests, docs, actual tarball, repository CLI, PptxGenJS, LibreOffice, Poppler and the presentation overflow checker.
- Produces: final evidence; no empty QA commit.

- [ ] **Step 1: Run complete automated gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts
git diff --check
```

Expected: zero TypeScript errors, all non-pending tests pass, performance gate passes, no whitespace errors.

- [ ] **Step 2: Rebuild and rerun final tarball/CLI smoke**

```sh
pnpm build
PACK_DIR="$(mktemp -d)"
pnpm --filter @jiayunxie/pptx pack --pack-destination "$PACK_DIR"
node scripts/smoke-npm-package.mjs "$PACK_DIR"/*.tgz
node packages/cli/dist/main.js --version
node packages/cli/dist/main.js doctor --offline
```

Expected: packed Node/browser/types succeed, CLI version is `0.1.0`, offline doctor succeeds.

- [ ] **Step 3: Generate native and PptxGenJS real decks**

Create `/tmp/pptx-table-cell-text-direction-create-qa.mjs` using public APIs only. Generate:

- `native-source.pptx`: omitted, horz, vert, vert270, wordArtVert cells, non-equal column widths/row heights, borders/fills/margins/valign/alignment;
- `native-edited.pptx`: change text and unrelated cell properties while preserving directions, then use the direct editor on one cell;
- `native-reopened.pptx`: reopen and write the edited deck;
- `pptxgenjs-baseline.pptx`: equivalent supported PptxGenJS 4.0.1 table.

Assert public snapshots `[undefined, undefined, 'vert', 'vert270', 'wordArtVert']`, expected direct tokens, editor state and reopen stability before writing files.

- [ ] **Step 4: Validate packages and mutation isolation**

Use the repository PPTX inspector with `powerpoint-2010`; require zero errors and zero warnings for all four decks. Require source→edited changes only the owning slide part and edited→reopened zero changed parts. Omitted/horz equivalent native one-cell decks must have zero changed parts.

- [ ] **Step 5: Render and inspect every page**

Use isolated LibreOffice profiles to export native edited and PptxGenJS baseline to PDF, rasterize with Poppler, and inspect original-resolution PNGs. Verify vertical/270/stacked directions, horizontal collapse, text, borders, fill, margins, vertical/horizontal alignment, unequal grid geometry, no repair, clipping, overlap or abnormal wrapping. LibreOffice's known `wordArtVert` horizontal-display limitation is acceptable only when native and PptxGenJS baselines behave identically and the OOXML token remains correct.

- [ ] **Step 6: Run overflow checks and final repository audit**

Run the presentation overflow checker against native edited and PptxGenJS baseline. Then require:

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Expected: both overflow checks pass; status contains only `?? .pnpm-store/`; synchronization is `0 0`. Do not create an empty QA commit.
