# Table Fill Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to execute this plan inline task-by-task. The user requires inline execution without subagents. Steps use checkbox (- [ ]) syntax for tracking.

**Goal:** Add strict table-level fill creation that materializes solid or none defaults onto unfilled physical cells with PptxGenJS 4.0.1 solid-fill final-state parity.

**Architecture:** Extend the descriptor-safe table option normalizer with one optional TableCellFill, overlay the normalized value only where a normalized cell has no direct fill, and reuse the existing cell fill renderer, snapshot, and editor. Retain only final direct-cell state; do not add table metadata, a table getter, or a table editor.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler.

## Global Constraints

- Public creation property is AddTableOptions.fill?: TableCellFill.
- Table solid or none is copied only to physical cells whose normalized cell fill is undefined; a cell solid or none completely overrides it.
- Omitted or runtime-undefined table fill preserves existing bytes; an empty table fill object is invalid.
- Omitted, runtime-undefined, and empty cell options inherit a provided table fill.
- Normalized table and cell fills are descriptor-safe, getter-free, ordinary/null-prototype-only, and deeply detached from caller state.
- Creation retains only direct cell state; clearing a cell later does not reapply the table default.
- Serialization remains margins, optional anchor, L/R/T/B borders, then fill.
- Native none writes direct a:noFill; native explicit zero transparency writes alpha=100000. These remain intentional differences from PptxGenJS 4.0.1.
- No TableModel.fill, table-level fill editor, effective style lookup, advanced fill, table border, horizontal alignment, direction/fit creation, merge, hyperlink, rich text, auto-page, repeated headers, or layout recomputation enters this slice.
- Every successful small item is reviewed, committed, pushed, fetched, and verified with origin/main...HEAD = 0 0; QA-only success creates no empty commit.
- Never add, stage, modify, or remove .pnpm-store/; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and materialize table fill

**Files:**

- Modify: packages/model/src/table-create.internal.ts
- Modify: packages/model/src/table-create.internal.test.ts

**Interfaces:**

- Consumes: normalizeTableCellFill(), normalized cell fill, existing margin and valign overlays, and renderTableCellFill().
- Produces: NormalizedTableDefinition.rows whose fill fields contain final table/cell precedence.

- [ ] **Step 1: Add failing precedence, detachment, and exact-output tests**

Add a focused test after current cell-fill creation coverage:

~~~ts
it('materializes a detached strict table fill under cell fills', () => {
  const sourceColor = {
    kind: 'scheme' as const,
    value: 'accent1' as 'accent1' | 'accent6',
  };
  const sourceFill = {
    kind: 'solid' as const,
    color: sourceColor,
    transparency: 33.3334,
  };
  const definition = normalizeTableDefinition([[
    'String',
    { text: 'Object' },
    { text: 'Empty', options: {} },
    { text: 'Undefined', options: { fill: undefined } },
    { text: 'Cell none', options: { fill: { kind: 'none' as const } } },
    { text: 'Cell solid', options: { fill: {
      kind: 'solid' as const,
      color: { kind: 'srgb' as const, value: 'FFFF00' },
      transparency: 25,
    } } },
  ]], {
    fill: sourceFill,
    margin: { top: 9 },
    valign: 'middle',
  });

  const tableFill = {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 33.333,
  };
  expect(definition.rows[0]!.map(({ fill }) => fill)).toEqual([
    tableFill,
    tableFill,
    tableFill,
    tableFill,
    { kind: 'none' },
    {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFFF00' },
      transparency: 25,
    },
  ]);

  const xml = renderTableGraphicFrame(41, definition);
  expect(xml.match(/<a:schemeClr val="accent1">/g)).toHaveLength(4);
  expect(xml).toContain(
    '<a:tcPr marL="91440" marR="91440" marT="114300" marB="45720" anchor="ctr">' +
    '<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR>' +
    '<a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB>' +
    '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="66667"/></a:schemeClr></a:solidFill>' +
    '</a:tcPr>',
  );
  expect(xml).toContain('</a:lnB><a:noFill/></a:tcPr>');

  sourceColor.value = 'accent6';
  sourceFill.transparency = 1;
  expect(definition.rows[0]![0]!.fill).toEqual(tableFill);
});
~~~

Add no-fill, omission, and override cases:

~~~ts
const tableNone = normalizeTableDefinition(
  [['Inherited', { text: 'Solid override', options: { fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: '00FF00' },
    transparency: 0,
  } } }]],
  { fill: { kind: 'none' } },
);
expect(tableNone.rows[0]!.map(({ fill }) => fill)).toEqual([
  { kind: 'none' },
  {
    kind: 'solid',
    color: { kind: 'srgb', value: '00FF00' },
    transparency: 0,
  },
]);

const omitted = renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], {}),
);
const runtimeUndefined = renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], { fill: undefined }),
);
expect(runtimeUndefined).toBe(omitted);
expect(() => normalizeTableDefinition([['Invalid']], { fill: {} })).toThrow();
~~~

Extend malformed table option coverage with this corpus:

~~~ts
let tableFillAccessorCalls = 0;
const accessorTableOptions = {};
Object.defineProperty(accessorTableOptions, 'fill', {
  get() {
    tableFillAccessorCalls += 1;
    return { kind: 'none' };
  },
  enumerable: true,
});
const accessorFill = {};
Object.defineProperty(accessorFill, 'kind', {
  get() {
    tableFillAccessorCalls += 1;
    return 'none';
  },
  enumerable: true,
});
const accessorColor = { kind: 'srgb' };
Object.defineProperty(accessorColor, 'value', {
  get() {
    tableFillAccessorCalls += 1;
    return 'FF0000';
  },
  enumerable: true,
});
class TableFillClass {
  kind = 'none';
}
const invalidTableFills: unknown[] = [
  null,
  false,
  'FF0000',
  [],
  {},
  accessorFill,
  new TableFillClass(),
  Object.create({ kind: 'none' }),
  { kind: 'none', transparency: 0 },
  { kind: 'none', extra: true },
  { kind: 'none', [Symbol('fill')]: true },
  { kind: 'solid' },
  { kind: 'solid', color: accessorColor },
  { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
  { kind: 'solid', color: { kind: 'scheme', value: 'Accent1' } },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: NaN },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: Infinity },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: -1 },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: 101 },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, extra: true },
];
expect(() => normalizeTableDefinition(
  [['Accessor']],
  accessorTableOptions,
)).toThrow();
for (const fill of invalidTableFills) {
  expect(() => normalizeTableDefinition([['Invalid']], { fill })).toThrow();
}
expect(tableFillAccessorCalls).toBe(0);
~~~

Require these failures before any later cell getter side effect.

- [ ] **Step 2: Run focused tests and confirm red**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "table fill|malformed matrices"
~~~

Expected: table fill is rejected as an unsupported option.

- [ ] **Step 3: Implement the minimal normalization overlay**

Add fill to OPTION_KEYS. Immediately after readOptions(options), normalize and resolve it:

~~~ts
const tableFill = normalizeTableCellFill(
  normalizedOptions.fill,
  'Table fill',
);
const fillResolvedRows = tableFill === undefined
  ? normalizedRows
  : normalizedRows.map((row) => row.map((cell) =>
      cell.fill === undefined
        ? { ...cell, fill: tableFill }
        : cell));
~~~

Change the existing table margin overlay to consume fillResolvedRows:

~~~ts
const marginResolvedRows = tableMargins === undefined
  ? fillResolvedRows
  : fillResolvedRows.map((row) => row.map((cell) => ({
    ...cell,
    margins: { ...tableMargins, ...(cell.margins ?? {}) },
  })));
~~~

Keep the valign overlay on marginResolvedRows. Do not change NormalizedTableDefinition or renderTableCell().

- [ ] **Step 4: Run focused regressions and typecheck**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
~~~

Require table/cell precedence, table none, explicit-zero/fractional/full transparency, exact child order, omission byte equality, deep detachment, getter-free rejection, and all existing margin/valign/border/fill tests to pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only the two files and commit:

~~~text
feat: propagate table fill during creation
~~~

Push, fetch, and require origin/main...HEAD = 0 0; .pnpm-store/ remains untracked.

---

### Task 2: Expose the public option and prove the model lifecycle

**Files:**

- Modify: packages/model/src/slide.ts
- Modify: packages/model/src/model.test.ts

**Interfaces:**

- Consumes: Task 1's direct-cell materialization.
- Produces: public AddTableOptions.fill?: TableCellFill and create/read/edit/duplicate/rollback/write/reopen evidence.

- [ ] **Step 1: Add the typed option and lifecycle test first**

Add a model test using:

~~~ts
const tableFill = {
  kind: 'solid' as const,
  color: { kind: 'scheme' as const, value: 'accent1' as const },
  transparency: 33.3334,
};
const options: AddTableOptions = {
  name: 'Table fill lifecycle',
  fill: tableFill,
  columnWidths: inches(2),
  rowHeights: inches(1),
};
const table = slide.addTable([[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'None override', options: { fill: { kind: 'none' } } },
  { text: 'Solid override', options: { fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: 'FFFF00' },
    transparency: 25,
  } } },
]], options);
const original = table.rows[0]!.cells.map(({ fill }) => fill);
expect(original).toEqual([
  {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 33.333,
  },
  {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 33.333,
  },
  { kind: 'none' },
  {
    kind: 'solid',
    color: { kind: 'srgb', value: 'FFFF00' },
    transparency: 25,
  },
]);
~~~

Mutate tableFill and its nested color after addTable() and require snapshots and XML unchanged. Duplicate the slide, then clear source cell 0 and replace source cell 1:

~~~ts
table.setCellFill(0, 0, undefined);
table.setCellFill(0, 1, {
  kind: 'solid',
  color: { kind: 'srgb', value: '00FF00' },
  transparency: 0,
});
const edited = table.rows[0]!.cells.map(({ fill }) => fill);
expect(table.rows[0]!.cells[0]!.fill).toBeUndefined();
expect(duplicateTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);
~~~

Exercise the outer transaction with:

~~~ts
const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
const rollbackJournal = [...pkg.mutations];
let rolledBack: TableModel | undefined;
expect(() => pkg.transaction(() => {
  table.setCellFill(0, 2, {
    kind: 'solid',
    color: { kind: 'srgb', value: 'FF0000' },
  });
  rolledBack = slide.addTable(
    [['Temporary']],
    { fill: { kind: 'none' } },
  );
  throw new Error('restore table fill defaults');
})).toThrow('restore table fill defaults');
expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
expect(pkg.mutations).toEqual(rollbackJournal);
expect(slide.shapes[0]).toBe(table);
expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(edited);
expect(duplicateTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);
expect(() => rolledBack!.rows).toThrow(ModelParseError);
~~~

Write/reopen and require exact source and duplicate fill matrices plus width/height vectors.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
~~~

Expected: excess-property error for AddTableOptions.fill.

- [ ] **Step 3: Widen only the public creation option**

In packages/model/src/slide.ts add:

~~~ts
readonly fill?: TableCellFill;
~~~

Do not add a TableModel property or table editor.

- [ ] **Step 4: Run model gates**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
~~~

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only the two files and commit:

~~~text
feat: expose table fill creation
~~~

Push, fetch, and require origin/main...HEAD = 0 0.

---

### Task 3: Cover the SDK lifecycle and invalid public inputs

**Files:**

- Modify: packages/sdk/src/index.test.ts

**Interfaces:**

- Consumes: public AddTableOptions.fill and existing TableModel cell editor.
- Produces: package-facing inheritance, override, detachment, duplicate, rollback, reopen, invalid-input, stable-identity, and no-mutation evidence.

- [ ] **Step 1: Add a focused public lifecycle test**

Add:

~~~ts
it('materializes public table fills through duplicate, rollback, and reopen', async () => {
  const document = PptxDocument.create();
  const slide = document.addSlide();
  const sourceColor = {
    kind: 'scheme' as const,
    value: 'accent1' as 'accent1' | 'accent6',
  };
  const sourceFill = {
    kind: 'solid' as const,
    color: sourceColor,
    transparency: 33.3334,
  };
  const table = slide.addTable([[
    'Inherited string',
    { text: 'Inherited object', options: {} },
    { text: 'None override', options: { fill: { kind: 'none' } } },
    { text: 'Solid override', options: { fill: {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFFF00' },
      transparency: 25,
    } } },
  ]], {
    name: 'SDK table fill lifecycle',
    fill: sourceFill,
    columnWidths: inches(2),
    rowHeights: inches(1),
  });
  const original = table.rows[0]!.cells.map(({ fill }) => fill);
  sourceColor.value = 'accent6';
  sourceFill.transparency = 1;
  expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);

  const duplicate = document.duplicateSlide(0);
  const duplicateTable = duplicate.shapes[0] as TableModel;
  expect(duplicateTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);

  table.setCellFill(0, 0, undefined);
  table.setCellFill(0, 1, {
    kind: 'solid',
    color: { kind: 'srgb', value: '00FF00' },
    transparency: 0,
  });
  const edited = table.rows[0]!.cells.map(({ fill }) => fill);
  expect(edited[0]).toBeUndefined();

  const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
  const rollbackJournal = [...document.opcPackage.mutations];
  expect(() => document.transaction(() => {
    table.setCellFill(0, 2, {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
    });
    slide.addTable([['Temporary']], { fill: { kind: 'none' } });
    throw new Error('restore table fill defaults');
  })).toThrow('restore table fill defaults');
  expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
  expect(document.opcPackage.mutations).toEqual(rollbackJournal);
  expect(slide.shapes).toHaveLength(1);
  expect(slide.shapes[0]).toBe(table);
  expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(edited);
  expect(duplicateTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);

  const reopened = await PptxDocument.open(await document.write());
  const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
  const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
  expect(reopenedTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(edited);
  expect(reopenedDuplicate.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);
  expect(reopenedTable.columnWidths).toEqual(Array(4).fill(inches(2)));
  expect(reopenedTable.rowHeights).toEqual([inches(1)]);
});
~~~

- [ ] **Step 2: Add invalid table-fill coverage before mutation**

Create accessor-backed table options, fill, and nested color objects whose getter count must remain zero:

~~~ts
let sdkFillGetterCalls = 0;
const accessorOptions = {};
Object.defineProperty(accessorOptions, 'fill', {
  get() {
    sdkFillGetterCalls += 1;
    return { kind: 'none' };
  },
  enumerable: true,
});
const accessorFill = {};
Object.defineProperty(accessorFill, 'kind', {
  get() {
    sdkFillGetterCalls += 1;
    return 'none';
  },
  enumerable: true,
});
const accessorColor = { kind: 'srgb' };
Object.defineProperty(accessorColor, 'value', {
  get() {
    sdkFillGetterCalls += 1;
    return 'FF0000';
  },
  enumerable: true,
});
class SdkFillClass {
  kind = 'none';
}
const invalidFills: unknown[] = [
  null,
  false,
  'FF0000',
  [],
  {},
  accessorFill,
  new SdkFillClass(),
  Object.create({ kind: 'none' }),
  { kind: 'none', transparency: 0 },
  { kind: 'none', extra: true },
  { kind: 'none', [Symbol('fill')]: true },
  { kind: 'solid' },
  { kind: 'solid', color: accessorColor },
  { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
  { kind: 'solid', color: { kind: 'scheme', value: 'Accent1' } },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: NaN },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: Infinity },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: -1 },
  { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: 101 },
];
expect(() => slide.addTable(
  [['Accessor table']],
  accessorOptions as AddTableOptions,
)).toThrow();
for (const fill of invalidFills) {
  expect(() => slide.addTable(
    [['Invalid table']],
    { fill } as unknown as AddTableOptions,
  )).toThrow();
}
expect(sdkFillGetterCalls).toBe(0);
~~~

Require unchanged bytes, journal, slides/shapes identity, existing table text and fill.

- [ ] **Step 3: Run SDK and model suites**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts \
  packages/model/src/model.test.ts
~~~

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the SDK test and commit:

~~~text
test: cover table fill defaults
~~~

Push, fetch, and require origin/main...HEAD = 0 0.

---

### Task 4: Prove PptxGenJS 4.0.1 final-state parity

**Files:**

- Modify: packages/pptxgenjs-adapter/src/index.test.ts

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 runtime and native table creation.
- Produces: solid sRGB/theme/transparency final-state equality and strict none/explicit-zero difference evidence.

- [ ] **Step 1: Create equivalent PptxGenJS and native tables**

Use equal geometry, row/column sizes, margin, valign, borders, and these logical fill cases:

~~~ts
const generatedRows = [[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Cell solid', options: {
    fill: { color: 'FFFF00', transparency: 50 },
  } },
]];
generated.addSlide().addTable(generatedRows, {
  x: 0.5,
  y: 0.5,
  w: 9,
  h: 1,
  colW: [3, 3, 3],
  rowH: [1],
  fill: { color: generated.SchemeColor.accent1, transparency: 25 },
  margin: 0.1,
  valign: 'middle',
});

const nativeTable = native.addSlide().addTable([[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Cell solid', options: { fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: 'FFFF00' },
    transparency: 50,
  } } },
]], {
  x: inches(0.5),
  y: inches(0.5),
  width: inches(9),
  height: inches(1),
  columnWidths: inches(3),
  rowHeights: inches(1),
  fill: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
  margin: 7.2,
  valign: 'middle',
});
~~~

- [ ] **Step 2: Assert final state, round trip, and strict differences**

Open the generated deck through the adapter and require both tables to expose:

~~~ts
[
  {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
  {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
  {
    kind: 'solid',
    color: { kind: 'srgb', value: 'FFFF00' },
    transparency: 50,
  },
]
~~~

Require equal geometry, text, borders, margins, anchors, fill choice/color/alpha snapshots, and write/reopen matrices. Cover the remaining solid states with this case table:

~~~ts
const tableFillCases = [
  {
    label: 'sRGB opaque',
    generatedFill: { color: 'FF0000' },
    nativeFill: {
      kind: 'solid' as const,
      color: { kind: 'srgb' as const, value: 'FF0000' },
    },
    expectedTransparency: undefined,
  },
  {
    label: 'theme fractional',
    generatedFill: {
      color: generated.SchemeColor.accent1,
      transparency: 33.333,
    },
    nativeFill: {
      kind: 'solid' as const,
      color: { kind: 'scheme' as const, value: 'accent1' as const },
      transparency: 33.333,
    },
    expectedTransparency: 33.333,
  },
  {
    label: 'sRGB full transparency',
    generatedFill: { color: '445566', transparency: 100 },
    nativeFill: {
      kind: 'solid' as const,
      color: { kind: 'srgb' as const, value: '445566' },
      transparency: 100,
    },
    expectedTransparency: 100,
  },
];
~~~

For each case, create a one-cell table with that table fill, import the generated deck, and require the native and imported TableCell.fill snapshots to equal the case's nativeFill after write/reopen.

Create explicit difference tables:

~~~ts
const nativeDifferenceTable = nativeDifferences.addSlide().addTable([[
  'Inherited none',
  { text: 'Explicit zero override', options: { fill: {
    kind: 'solid',
    color: { kind: 'srgb', value: '00FF00' },
    transparency: 0,
  } } },
]], { fill: { kind: 'none' } });
generatedDifferences.addSlide().addTable([[
  'Collapsed none',
  { text: 'Collapsed zero alpha', options: {
    fill: { color: '00FF00', transparency: 0 },
  } },
]], { fill: { type: 'none' } });
~~~

Require native direct none and alpha=100000 to survive write/reopen, while PptxGenJS type none and zero transparency collapse to absent direct state/alpha. These are asserted differences, not parity failures. Require malformed native table fill to throw before mutation.

- [ ] **Step 3: Run adapter and model suites**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
~~~

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the adapter test and commit:

~~~text
test: compare table fill defaults with pptxgenjs
~~~

Push, fetch, and require origin/main...HEAD = 0 0.

---

### Task 5: Smoke the actual Node/browser/types package

**Files:**

- Modify: scripts/smoke-npm-package.mjs

**Interfaces:**

- Consumes: packed @jiayunxie/pptx Node/browser/type surfaces.
- Produces: tableFillCreation: true in smoke JSON.

- [ ] **Step 1: Extend Node smoke**

Add a detached source table fill:

~~~js
const creationTableFillColor = { kind: 'scheme', value: 'accent1' };
const creationTableFill = {
  kind: 'solid',
  color: creationTableFillColor,
  transparency: 50,
};
~~~

Pass fill: creationTableFill in created table options. Keep one empty-options cell to inherit, one cell solid override, and one cell none override. Mutate creationTableFillColor and creationTableFill after addTable() and require the live snapshot unchanged. Clear an inherited cell with setCellFill(..., undefined), require no re-inheritance, then require the same edited matrix after reopen.

Compute tableFillCreation only from real create/snapshot/detachment/override/clear/reopen assertions and expose it in the JSON result.

- [ ] **Step 2: Extend browser and declaration smoke**

Mirror table fill inheritance, source detachment, cell override, direct none, clear, and reopen in the browser bundle fixture. In the declaration fixture add:

~~~ts
const tableOptions: AddTableOptions = {
  fill: snapshotCellFill,
  margin: cellMargins,
  valign: cellAlignment,
};
~~~

Require browser table fill creation to fail loudly if any snapshot differs.

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

~~~sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_fill_package_dir=$(mktemp -d /tmp/pptx-table-fill-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_fill_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_fill_package_dir/jiayunxie-pptx-0.1.0.tgz"
~~~

Require every smoke flag true, including tableFillCreation, tableCellFillCreation, tableMarginCreation, tableVerticalAlignmentCreation, types, browser, and CLI 0.1.0.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the smoke script and commit:

~~~text
test: smoke packed table fill defaults
~~~

Push, fetch, and require origin/main...HEAD = 0 0.

---

### Task 6: Document the public contract and remaining boundary

**Files:**

- Modify: CHANGELOG.md
- Modify: docs/api/README.md
- Modify: docs/compatibility/pptxgenjs-baseline.md
- Modify: packages/pptx/README.md

**Interfaces:**

- Consumes: completed runtime/package evidence.
- Produces: accurate table-level creation and unsupported documentation.

- [ ] **Step 1: Update examples and contract**

Add a table fill default to examples while retaining cell solid and none overrides. Document:

- strict TableCellFill solid/none shape and sRGB/theme/transparency behavior;
- table then cell whole-value precedence and physical-cell materialization;
- omitted/runtime-undefined table bytes and invalid empty table fill;
- immediate deep-detached TableCell.fill access;
- clear/replace operates only on final direct cell state and never re-inherits;
- fixed margins/anchor/borders/fill serialization order.

- [ ] **Step 2: Update compatibility and unsupported lists**

Mark AddTableOptions.fill creation as supported and PptxGenJS solid final-state compatible. Retain native direct none and explicit-zero distinctions. Keep table fill getter/editor, advanced fills, table border, horizontal alignment, direction/fit creation, merge, hyperlink, rich text, auto-page, repeated headers, content measurement, and layout recomputation unsupported.

- [ ] **Step 3: Scan contradictions and run typecheck**

~~~sh
git diff --check
rg -n --pcre2 \
  'table-level (fill|填充).*(unsupported|尚未支持)|AddTableOptions.*fill.*unsupported' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
~~~

Review each hit in context: creation support is valid, while table-level read/edit and advanced fill remain unsupported.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the four documents and commit:

~~~text
docs: document table fill creation
~~~

Push, fetch, and require origin/main...HEAD = 0 0.

---

### Task 7: Run full gates and real-deck QA

**Files:**

- Review every Task 1–6 path; never stage or delete .pnpm-store/.

**Interfaces:**

- Consumes: implementation, tests, docs, actual tarball, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces: a verified set of pushed commits; any defect gets its own focused fix/review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-fill-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-fill-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
~~~

- [ ] **Step 2: Rebuild and smoke the actual tarball**

~~~sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_fill_qa_package_dir=$(mktemp -d /tmp/pptx-table-fill-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_fill_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_fill_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
~~~

- [ ] **Step 3: Generate nine real native and PptxGenJS decks**

Create /tmp/pptx-table-fill-qa.mjs with apply_patch using this complete source:

~~~js
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import {
  PptxDocument,
  TableModel,
  inches,
} from 'file:///Users/jeremy/workspace/pptx/packages/pptx/dist/index.js';

const require = createRequire(
  '/Users/jeremy/workspace/pptx/packages/pptxgenjs-adapter/package.json',
);
const PptxGenJS = require('pptxgenjs');
const output = '/tmp/pptx-table-fill';
await mkdir(output, { recursive: true });

const baseRows = [
  ['Inherited A', 'Inherited B', 'Inherited C', 'Inherited D'],
  ['Inherited E', 'Inherited F', 'Inherited G', 'Inherited H'],
];
const tableOptions = {
  name: 'Table fill creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(1.5), inches(2.5), inches(2), inches(2)],
  rowHeights: [inches(0.8), inches(1.6)],
  margin: 7.2,
  valign: 'middle',
};

function build(rows, options = {}) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const table = document.addSlide().addTable(rows, { ...tableOptions, ...options });
  return { document, table };
}

const omitted = build(baseRows);
const runtimeUndefined = build(baseRows, { fill: undefined });
await writeFile(output + '/omitted.pptx', await omitted.document.write());
await writeFile(output + '/undefined.pptx', await runtimeUndefined.document.write());

const none = build(baseRows, { fill: { kind: 'none' } });
assert.ok(none.table.rows.every(({ cells }) =>
  cells.every(({ fill }) => fill?.kind === 'none')));
await writeFile(output + '/table-none.pptx', await none.document.write());

const solidFill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'D9EAF7' },
  transparency: 25,
};
const solid = build(baseRows, { fill: solidFill });
assert.ok(solid.table.rows.every(({ cells }) =>
  cells.every(({ fill }) => fill?.kind === 'solid' &&
    fill.color.kind === 'srgb' &&
    fill.color.value === 'D9EAF7' &&
    fill.transparency === 25)));
await writeFile(output + '/table-solid.pptx', await solid.document.write());

const theme = build(baseRows, { fill: {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 33.3334,
} });
assert.equal(theme.table.rows[0].cells[0].fill.transparency, 33.333);
await writeFile(output + '/table-theme.pptx', await theme.document.write());

const mixedRows = [
  [
    'Inherited string',
    { text: 'Inherited object', options: {} },
    { text: 'Yellow override', options: { fill: {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFFF00' },
      transparency: 50,
    } } },
    { text: 'Direct none', options: { fill: { kind: 'none' } } },
  ],
  [
    { text: 'Inherited border', options: { border: {
      bottom: {
        kind: 'line',
        color: { kind: 'srgb', value: '70AD47' },
        width: 3,
        style: 'dash',
      },
    } } },
    { text: 'Explicit zero', options: { fill: {
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
      transparency: 0,
    } } },
    { text: 'Full transparency', options: { fill: {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 100,
    } } },
    { text: 'Inherited H' },
  ],
];
const source = build(mixedRows, { fill: solidFill });
const sourceFills = source.table.rows.map(({ cells }) =>
  cells.map(({ fill }) => fill));
assert.deepEqual(sourceFills[0][0], solidFill);
assert.deepEqual(sourceFills[0][1], solidFill);
assert.equal(sourceFills[0][2].color.value, 'FFFF00');
assert.equal(sourceFills[0][3].kind, 'none');
assert.equal(sourceFills[1][1].transparency, 0);
assert.equal(sourceFills[1][2].transparency, 100);
await writeFile(output + '/mixed-source.pptx', await source.document.write());

source.table.setCellFill(0, 0, undefined);
source.table.setCellFill(0, 1, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 40,
});
source.table.setCellText(0, 1, 'Edited inherited fill');
const editedBytes = await source.document.write();
await writeFile(output + '/mixed-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[0].cells[0].fill, undefined);
assert.deepEqual(reopenedTable.rows[0].cells[1].fill, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent2' },
  transparency: 40,
});
assert.equal(reopenedTable.rows[0].cells[1].text, 'Edited inherited fill');
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/mixed-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      'Inherited string',
      { text: 'Inherited object', options: {} },
      { text: 'Yellow override', options: {
        fill: { color: 'FFFF00', transparency: 50 },
      } },
      { text: 'Inherited D', options: {} },
    ],
    [
      { text: 'Inherited border', options: {
        border: [
          { type: 'none' },
          { type: 'none' },
          { type: 'dash', color: '70AD47', pt: 3 },
          { type: 'none' },
        ],
      } },
      { text: 'Green opaque', options: {
        fill: { color: '00FF00' },
      } },
      { text: 'Full transparency', options: {
        fill: { color: 'FF0000', transparency: 100 },
      } },
      { text: 'Inherited H', options: {} },
    ],
  ],
  {
    x: 0.75,
    y: 0.75,
    w: 8,
    h: 2.4,
    colW: [1.5, 2.5, 2, 2],
    rowH: [0.8, 1.6],
    margin: 0.1,
    valign: 'middle',
    fill: { color: 'D9EAF7', transparency: 25 },
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceFills,
  reopenedFills: reopenedTable.rows.map(({ cells }) =>
    cells.map(({ fill }) => fill)),
  editedText: reopenedTable.rows[0].cells[1].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
~~~

Run node /tmp/pptx-table-fill-qa.mjs and require all assertions to pass.

- [ ] **Step 4: Validate and diff packages**

~~~sh
for deck in /tmp/pptx-table-fill/*.pptx; do
  node packages/pptx/dist/cli.js --json package validate \
    "$deck" --profile powerpoint-2010
done
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-fill/omitted.pptx \
  /tmp/pptx-table-fill/undefined.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-fill/mixed-source.pptx \
  /tmp/pptx-table-fill/mixed-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-fill/mixed-edited.pptx \
  /tmp/pptx-table-fill/mixed-reopened.pptx
~~~

Require zero errors/warnings. Omitted/undefined must have zero changed parts; source/edited may change only /ppt/slides/slide1.xml; edited/reopened must have zero changed parts.

- [ ] **Step 5: Render and inspect**

Convert mixed-source.pptx, mixed-edited.pptx, and pptxgenjs-baseline.pptx to isolated PDFs with LibreOffice. Rasterize every page with Poppler at 180 DPI. Run slides_test.py on edited and baseline decks using the bundled presentations Python environment.

Inspect every PNG at original detail. Require visible table fill inheritance and cell overrides, clear/edit behavior, preserved vertical alignment, margins, borders, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict normalization, descriptor safety, whole-value precedence, direct ownership, serialization order, omission, declarations, lifecycle isolation, PptxGenJS differences, smoke flags, docs boundaries, and invalid-input no-mutation assertions.

~~~sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
~~~

Require 0 0 and only ?? .pnpm-store/. QA-only success creates no empty commit.
