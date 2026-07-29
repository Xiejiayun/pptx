# Table Border Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Execute this plan inline task-by-task. The user requires no subagents and requires every independently reviewable item to be reviewed, committed, pushed, fetched, and synchronized before continuing. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict table-level border creation that materializes scalar, TRBL, or named defaults onto cells without an effective cell border, with PptxGenJS 4.0.1 final-state parity where its behavior is well-defined.

**Architecture:** Extend descriptor-safe table option normalization with one optional `TableCellBorderInput`, normalize it through the existing strict border codec, and overlay the whole normalized value only where a normalized cell has no border value. Reuse the existing cell border renderer, snapshot, and editor; retain only final direct-cell state and add no table metadata, getter, or editor.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler.

## Global Constraints

- Public creation property is `AddTableOptions.border?: TableCellBorderInput`.
- Table scalar, TRBL, or named border is copied only to physical cells whose normalized `borders` is `undefined`; any non-empty cell scalar, TRBL, named value, or explicit none completely blocks it.
- Omitted, runtime-undefined, empty, and all-undefined table border preserve existing bytes; empty and all-undefined cell border inherit a provided table border.
- A partial value defines only its own sides; existing rendering writes canonical direct noFill for the other sides and never merges them from the other input layer.
- Normalized table and cell borders are descriptor-safe, getter-free, ordinary/null-prototype-only, and deeply detached from caller state.
- Creation retains only direct cell state; clearing or replacing a cell later does not reapply the table default.
- Serialization remains margins, optional anchor, L/R/T/B borders, then fill.
- Native empty values mean no effective input, style omission writes no direct dash token, and tuple zero width remains zero. These remain intentional differences from PptxGenJS 4.0.1.
- No `TableModel.border`, table-level border editor, diagonal or advanced line, border transparency, shared-edge/style resolution, horizontal alignment, direction/fit creation, merge, hyperlink, rich text, auto-page, repeated headers, or layout recomputation enters this slice.
- Every successful small item is reviewed, committed, pushed, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only success creates no empty commit.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Normalize and materialize table borders

**Files:**

- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**

- Consumes: `normalizeTableCellBorders(value, context)`, normalized cell `borders`, and `renderTableCellBorders()`.
- Produces: `NormalizedTableDefinition.rows` whose border fields contain final whole-value table/cell precedence.

- [ ] **Step 1: Add failing precedence, detachment, and exact-output tests**

Add focused coverage after the current cell-border creation cases:

~~~ts
it('materializes a detached strict table border under whole cell border values', () => {
  const sourceColor = {
    kind: 'scheme' as const,
    value: 'accent1' as 'accent1' | 'accent6',
  };
  const sourceBorder = {
    kind: 'line' as const,
    color: sourceColor,
    width: 1.50004,
    style: 'dash' as const,
  };
  const definition = normalizeTableDefinition([[
    'String',
    { text: 'Object' },
    { text: 'Empty options', options: {} },
    { text: 'Undefined', options: { border: undefined } },
    { text: 'Empty border', options: { border: {} } },
    { text: 'All undefined', options: {
      border: [undefined, undefined, undefined, undefined],
    } },
    { text: 'Cell partial', options: {
      border: { left: { kind: 'none' as const } },
    } },
    { text: 'Cell none', options: { border: { kind: 'none' as const } } },
  ]], {
    border: sourceBorder,
    fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    margin: { top: 9 },
    valign: 'middle',
  });

  const tableBorders = {
    top: {
      kind: 'line', color: { kind: 'scheme', value: 'accent1' },
      width: 1.5, style: 'dash',
    },
    right: {
      kind: 'line', color: { kind: 'scheme', value: 'accent1' },
      width: 1.5, style: 'dash',
    },
    bottom: {
      kind: 'line', color: { kind: 'scheme', value: 'accent1' },
      width: 1.5, style: 'dash',
    },
    left: {
      kind: 'line', color: { kind: 'scheme', value: 'accent1' },
      width: 1.5, style: 'dash',
    },
  };
  expect(definition.rows[0]!.slice(0, 6).map(({ borders }) => borders))
    .toEqual(Array(6).fill(tableBorders));
  expect(definition.rows[0]![6]!.borders).toEqual({ left: { kind: 'none' } });
  expect(definition.rows[0]![7]!.borders).toEqual({
    top: { kind: 'none' },
    right: { kind: 'none' },
    bottom: { kind: 'none' },
    left: { kind: 'none' },
  });

  const xml = renderTableGraphicFrame(41, definition);
  expect(xml.match(/<a:schemeClr val="accent1"\/>/g)).toHaveLength(24);
  expect(xml).toContain(
    '<a:tcPr marL="91440" marR="91440" marT="114300" marB="45720" anchor="ctr">' +
    '<a:lnL w="19050" cap="flat" cmpd="sng" algn="ctr">' +
    '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>' +
    '<a:prstDash val="sysDash"/><a:round/>' +
    '<a:headEnd type="none" w="med" len="med"/>' +
    '<a:tailEnd type="none" w="med" len="med"/></a:lnL>',
  );
  expect(xml).toContain(
    '</a:lnL><a:lnR w="19050" cap="flat" cmpd="sng" algn="ctr">',
  );
  expect(xml).toContain('</a:lnB><a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>');

  sourceColor.value = 'accent6';
  sourceBorder.width = 9;
  expect(definition.rows[0]![0]!.borders).toEqual(tableBorders);
});
~~~

Add scalar none, TRBL/named, zero/fractional/max width, omission, and all-undefined cases:

~~~ts
const none = normalizeTableDefinition([['Inherited']], {
  border: { kind: 'none' },
});
expect(none.rows[0]![0]!.borders).toEqual({
  top: { kind: 'none' }, right: { kind: 'none' },
  bottom: { kind: 'none' }, left: { kind: 'none' },
});

const tuple = normalizeTableDefinition([['Tuple']], {
  border: [
    { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 0 },
    undefined,
    { kind: 'line', color: { kind: 'scheme', value: 'accent2' }, width: 0.33334 },
    { kind: 'line', color: { kind: 'srgb', value: '00FF00' }, width: 1584, style: 'solid' },
  ],
});
expect(tuple.rows[0]![0]!.borders).toEqual({
  top: { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 0 },
  bottom: { kind: 'line', color: { kind: 'scheme', value: 'accent2' }, width: 0.333 },
  left: {
    kind: 'line', color: { kind: 'srgb', value: '00FF00' },
    width: 1584, style: 'solid',
  },
});

const named = normalizeTableDefinition([['Named']], {
  border: {
    right: { kind: 'none' },
    bottom: {
      kind: 'line', color: { kind: 'scheme', value: 'accent3' },
      width: 2, style: 'dash',
    },
  },
});
expect(named.rows[0]![0]!.borders).toEqual({
  right: { kind: 'none' },
  bottom: {
    kind: 'line', color: { kind: 'scheme', value: 'accent3' },
    width: 2, style: 'dash',
  },
});

const omitted = renderTableGraphicFrame(42, normalizeTableDefinition([['Same']], {}));
expect(renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], { border: undefined }),
)).toBe(omitted);
expect(renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], { border: {} }),
)).toBe(omitted);
expect(renderTableGraphicFrame(
  42,
  normalizeTableDefinition([['Same']], {
    border: [undefined, undefined, undefined, undefined],
  }),
)).toBe(omitted);
~~~

Extend malformed-table-option coverage with descriptor-safe invalid values:

~~~ts
let tableBorderAccessorCalls = 0;
const accessorTableOptions = {};
Object.defineProperty(accessorTableOptions, 'border', {
  get() {
    tableBorderAccessorCalls += 1;
    return { kind: 'none' };
  },
  enumerable: true,
});
const accessorBorder = {};
Object.defineProperty(accessorBorder, 'kind', {
  get() {
    tableBorderAccessorCalls += 1;
    return 'none';
  },
  enumerable: true,
});
const accessorSide = {};
Object.defineProperty(accessorSide, 'top', {
  get() {
    tableBorderAccessorCalls += 1;
    return { kind: 'none' };
  },
  enumerable: true,
});
const accessorColor = { kind: 'srgb' };
Object.defineProperty(accessorColor, 'value', {
  get() {
    tableBorderAccessorCalls += 1;
    return 'FF0000';
  },
  enumerable: true,
});
class TableBorderClass { kind = 'none'; }
const sparse = Array(4);
const invalidTableBorders: unknown[] = [
  null, false, 'FF0000', [], [undefined], sparse,
  accessorBorder, accessorSide, new TableBorderClass(),
  Object.create({ kind: 'none' }),
  { kind: 'none', width: 0 },
  { kind: 'none', extra: true },
  { kind: 'none', [Symbol('border')]: true },
  { top: undefined, extra: true },
  { kind: 'line' },
  { kind: 'line', color: accessorColor, width: 1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FFF' }, width: 1 },
  { kind: 'line', color: { kind: 'scheme', value: 'Accent1' }, width: 1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: NaN },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: Infinity },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: -1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1584.001 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1, style: 'dot' },
];
expect(() => normalizeTableDefinition([['Accessor']], accessorTableOptions)).toThrow();
for (const border of invalidTableBorders) {
  expect(() => normalizeTableDefinition([['Invalid']], { border })).toThrow();
}
expect(tableBorderAccessorCalls).toBe(0);
~~~

- [ ] **Step 2: Run focused tests and confirm red**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  -t "table border|malformed matrices"
~~~

Expected: table border is rejected as an unsupported table option.

- [ ] **Step 3: Implement the minimal whole-value overlay**

Add `border` to `OPTION_KEYS`. Immediately after `readOptions(options)`, normalize and resolve it before fill, margin, and valign:

~~~ts
const tableBorders = normalizeTableCellBorders(
  normalizedOptions.border,
  'Table border',
);
const borderResolvedRows = tableBorders === undefined
  ? normalizedRows
  : normalizedRows.map((row) => row.map((cell) =>
      cell.borders === undefined
        ? { ...cell, borders: tableBorders }
        : cell));
const tableFill = normalizeTableCellFill(normalizedOptions.fill, 'Table fill');
const fillResolvedRows = tableFill === undefined
  ? borderResolvedRows
  : borderResolvedRows.map((row) => row.map((cell) =>
      cell.fill === undefined
        ? { ...cell, fill: tableFill }
        : cell));
~~~

Do not change `NormalizedTableDefinition`, `normalizeTableCellOptions()`, or `renderTableCell()`.

- [ ] **Step 4: Run focused regressions and typecheck**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
~~~

Require whole-value precedence, scalar/TRBL/named/none, sRGB/theme, omitted/solid/dash, zero/fractional/max width, exact L/R/T/B order, border-before-fill order, omission byte equality, deep detachment, getter-free rejection, and all existing table tests to pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

Stage only the two files and commit:

~~~text
feat: propagate table borders during creation
~~~

Push over SSH 443, fetch over SSH 443, and require `git rev-list --left-right --count origin/main...HEAD` to print `0 0`; `.pnpm-store/` remains untracked.

---

### Task 2: Expose the public option and prove the model lifecycle

**Files:**

- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**

- Consumes: Task 1 direct-cell border materialization.
- Produces: public `AddTableOptions.border?: TableCellBorderInput` and create/read/edit/duplicate/rollback/write/reopen evidence.

- [ ] **Step 1: Add the typed option and lifecycle test first**

Add a model test using the public type:

~~~ts
const sourceColor = {
  kind: 'scheme' as const,
  value: 'accent1' as 'accent1' | 'accent6',
};
const sourceBorder = {
  kind: 'line' as const,
  color: sourceColor,
  width: 1.50004,
  style: 'dash' as const,
};
const options: AddTableOptions = {
  name: 'Table border lifecycle',
  border: sourceBorder,
  columnWidths: inches(2),
  rowHeights: inches(1),
};
const table = slide.addTable([[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Partial override', options: {
    border: { left: { kind: 'none' } },
  } },
  { text: 'None override', options: { border: { kind: 'none' } } },
]], options);
const original = table.rows[0]!.cells.map(({ borders }) => borders);
expect(original[0]).toEqual({
  top: {
    kind: 'line', color: { kind: 'scheme', value: 'accent1' },
    width: 1.5, style: 'dash',
  },
  right: {
    kind: 'line', color: { kind: 'scheme', value: 'accent1' },
    width: 1.5, style: 'dash',
  },
  bottom: {
    kind: 'line', color: { kind: 'scheme', value: 'accent1' },
    width: 1.5, style: 'dash',
  },
  left: {
    kind: 'line', color: { kind: 'scheme', value: 'accent1' },
    width: 1.5, style: 'dash',
  },
});
expect(original[1]).toEqual(original[0]);
expect(original[2]).toEqual({
  top: { kind: 'none' }, right: { kind: 'none' },
  bottom: { kind: 'none' }, left: { kind: 'none' },
});
expect(original[3]).toEqual(original[2]);

sourceColor.value = 'accent6';
sourceBorder.width = 9;
expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);
~~~

Duplicate the slide, clear source cell 0, and replace source cell 1:

~~~ts
const duplicate = presentation.duplicateSlide(0);
const duplicateTable = duplicate.shapes[0] as TableModel;
table.setCellBorders(0, 0, undefined);
table.setCellBorders(0, 1, {
  right: {
    kind: 'line', color: { kind: 'srgb', value: '00FF00' },
    width: 0, style: 'solid',
  },
});
const edited = table.rows[0]!.cells.map(({ borders }) => borders);
expect(edited[0]).toBeUndefined();
expect(edited[1]).toEqual({
  right: {
    kind: 'line', color: { kind: 'srgb', value: '00FF00' },
    width: 0, style: 'solid',
  },
});
expect(duplicateTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);
~~~

Exercise outer rollback and stale-model invalidation:

~~~ts
const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
const rollbackJournal = [...pkg.mutations];
let rolledBack: TableModel | undefined;
expect(() => pkg.transaction(() => {
  table.setCellBorders(0, 2, {
    kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 2,
  });
  rolledBack = slide.addTable(
    [['Temporary']],
    { border: { kind: 'none' } },
  );
  throw new Error('restore table border defaults');
})).toThrow('restore table border defaults');
expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
expect(pkg.mutations).toEqual(rollbackJournal);
expect(slide.shapes[0]).toBe(table);
expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(edited);
expect(duplicateTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);
expect(() => rolledBack!.rows).toThrow(ModelParseError);
~~~

Write/reopen and require exact source and duplicate border matrices plus the original width/height vectors.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
~~~

Expected: excess-property error for `AddTableOptions.border`.

- [ ] **Step 3: Widen only the public creation option**

In `packages/model/src/slide.ts`, add:

~~~ts
readonly border?: TableCellBorderInput;
~~~

Do not add a `TableModel` property or table-level editor.

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
feat: expose table border creation
~~~

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 3: Cover the SDK lifecycle and invalid public inputs

**Files:**

- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: public `AddTableOptions.border` and existing `TableModel.setCellBorders()`.
- Produces: package-facing inheritance, override, detachment, duplicate, rollback, reopen, invalid-input, stable-identity, and no-mutation evidence.

- [ ] **Step 1: Add a focused public lifecycle test**

Add this SDK-level setup and assert the same matrices before and after reopen:

~~~ts
it('materializes public table borders through duplicate, rollback, and reopen', async () => {
  const document = PptxDocument.create();
  const slide = document.addSlide();
  const sourceColor = {
    kind: 'scheme' as const,
    value: 'accent1' as 'accent1' | 'accent6',
  };
  const sourceBorder = {
    kind: 'line' as const,
    color: sourceColor,
    width: 1.50004,
    style: 'dash' as const,
  };
  const table = slide.addTable([[
    'Inherited string',
    { text: 'Inherited empty', options: { border: {} } },
    { text: 'Partial override', options: {
      border: { bottom: { kind: 'none' } },
    } },
    { text: 'None override', options: { border: { kind: 'none' } } },
  ]], {
    name: 'SDK table border lifecycle',
    border: sourceBorder,
    columnWidths: inches(2),
    rowHeights: inches(1),
  });
  const original = table.rows[0]!.cells.map(({ borders }) => borders);
  sourceColor.value = 'accent6';
  sourceBorder.width = 9;
  expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);

  const duplicate = document.duplicateSlide(0);
  const duplicateTable = duplicate.shapes[0] as TableModel;
  table.setCellBorders(0, 0, undefined);
  table.setCellBorders(0, 1, {
    right: {
      kind: 'line', color: { kind: 'srgb', value: '00FF00' },
      width: 0, style: 'solid',
    },
  });
  const edited = table.rows[0]!.cells.map(({ borders }) => borders);
  expect(edited[0]).toBeUndefined();

  const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
  const rollbackJournal = [...document.opcPackage.mutations];
  expect(() => document.transaction(() => {
    table.setCellBorders(0, 2, {
      kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 2,
    });
    slide.addTable([['Temporary']], { border: { kind: 'none' } });
    throw new Error('restore table border defaults');
  })).toThrow('restore table border defaults');
  expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
  expect(document.opcPackage.mutations).toEqual(rollbackJournal);
  expect(slide.shapes).toHaveLength(1);
  expect(slide.shapes[0]).toBe(table);
  expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(edited);
  expect(duplicateTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);

  const reopened = await PptxDocument.open(await document.write());
  const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
  const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
  expect(reopenedTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(edited);
  expect(reopenedDuplicate.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);
  expect(reopenedTable.columnWidths).toEqual(Array(4).fill(inches(2)));
  expect(reopenedTable.rowHeights).toEqual([inches(1)]);
});
~~~

- [ ] **Step 2: Add invalid table-border coverage before mutation**

Create accessor-backed table options, border, side, tuple item, and nested color objects whose getter count must remain zero. Use this corpus in the SDK test:

~~~ts
let sdkBorderGetterCalls = 0;
const accessorOptions = {};
Object.defineProperty(accessorOptions, 'border', {
  get() { sdkBorderGetterCalls += 1; return { kind: 'none' }; },
  enumerable: true,
});
const accessorBorder = {};
Object.defineProperty(accessorBorder, 'kind', {
  get() { sdkBorderGetterCalls += 1; return 'none'; },
  enumerable: true,
});
const accessorSide = {};
Object.defineProperty(accessorSide, 'top', {
  get() { sdkBorderGetterCalls += 1; return { kind: 'none' }; },
  enumerable: true,
});
const accessorColor = { kind: 'srgb' };
Object.defineProperty(accessorColor, 'value', {
  get() { sdkBorderGetterCalls += 1; return 'FF0000'; },
  enumerable: true,
});
class SdkBorderClass { kind = 'none'; }
const invalidBorders: unknown[] = [
  null, false, 'FF0000', [], [undefined], Array(4),
  accessorBorder, accessorSide, new SdkBorderClass(),
  Object.create({ kind: 'none' }),
  { kind: 'none', width: 0 }, { kind: 'none', extra: true },
  { kind: 'none', [Symbol('border')]: true },
  { kind: 'line' },
  { kind: 'line', color: accessorColor, width: 1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FFF' }, width: 1 },
  { kind: 'line', color: { kind: 'scheme', value: 'Accent1' }, width: 1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: NaN },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: Infinity },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: -1 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1584.001 },
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1, style: 'dot' },
];
expect(() => slide.addTable(
  [['Accessor table']], accessorOptions as AddTableOptions,
)).toThrow();
for (const border of invalidBorders) {
  expect(() => slide.addTable(
    [['Invalid table']], { border } as unknown as AddTableOptions,
  )).toThrow();
}
expect(sdkBorderGetterCalls).toBe(0);
~~~

Capture package bytes, mutation journal, slide/shape identities, and the existing table matrix before the corpus; require all of them unchanged afterward.

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
test: cover table border defaults
~~~

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 4: Prove PptxGenJS 4.0.1 final-state parity

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 runtime, adapter import, and native table creation.
- Produces: scalar/TRBL/none/zero/solid/dash/cell-override final-state equality and documented strict differences.

- [ ] **Step 1: Create equivalent scalar-default and cell-override tables**

Use equal geometry, sizes, margins, valign, and fill:

~~~ts
const generated = new PptxGenJS();
generated.layout = 'LAYOUT_WIDE';
generated.addSlide().addTable([[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Cell partial', options: {
    border: [
      { type: 'none' }, { type: 'none' },
      { type: 'dash', color: '70AD47', pt: 3 }, { type: 'none' },
    ],
  } },
  { text: 'Cell none', options: { border: { type: 'none' } } },
]], {
  x: 0.5, y: 0.5, w: 12, h: 1,
  colW: [3, 3, 3, 3], rowH: [1],
  border: { type: 'dash', color: '4472C4', pt: 1.5 },
  fill: { color: 'D9EAF7' }, margin: 0.1, valign: 'middle',
});

const native = PptxDocument.create({ slideSize: 'wide' });
native.addSlide().addTable([[
  'Inherited string',
  { text: 'Inherited object', options: {} },
  { text: 'Cell partial', options: { border: {
    bottom: {
      kind: 'line', color: { kind: 'srgb', value: '70AD47' },
      width: 3, style: 'dash',
    },
  } } },
  { text: 'Cell none', options: { border: { kind: 'none' } } },
]], {
  x: inches(0.5), y: inches(0.5),
  width: inches(12), height: inches(1),
  columnWidths: inches(3), rowHeights: inches(1),
  border: {
    kind: 'line', color: { kind: 'srgb', value: '4472C4' },
    width: 1.5, style: 'dash',
  },
  fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
  margin: 7.2, valign: 'middle',
});
~~~

Import generated bytes and require native/imported equality for transform, grid, rows, text, fill, margins, anchor, and final L/R/T/B state. The first two cells must have four 1.5pt dashed `4472C4` lines; the partial cell must have three none sides and one 3pt dashed `70AD47` bottom; the none cell must have four none sides. Require the same matrices after write/reopen.

- [ ] **Step 2: Cover TRBL states and explicit differences**

Use this PptxGenJS/native case table for one-cell tables:

~~~ts
const borderCases = [
  {
    label: 'scalar zero solid',
    generated: { type: 'solid', color: 'FF0000', pt: 0 },
    native: {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: 'FF0000' },
      width: 0,
      style: 'solid' as const,
    },
  },
  {
    label: 'scalar none',
    generated: { type: 'none' },
    native: { kind: 'none' as const },
  },
  {
    label: 'fractional dash',
    generated: { type: 'dash', color: '00FF00', pt: 0.333 },
    native: {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '00FF00' },
      width: 0.333,
      style: 'dash' as const,
    },
  },
];
~~~

Create a full TRBL table with top solid red 1pt, right none, bottom dashed green 2pt, and left solid blue 3pt in both libraries; require exact final border snapshots and direct XML order. Then assert the intentional differences using real generated output:

~~~ts
generatedDifferences.addSlide().addTable([[
  'Empty table default',
]], { border: {} });
generatedDifferences.addSlide().addTable([[
  { text: 'Empty cell blocks', options: { border: {} } },
]], { border: { type: 'solid', color: 'FF0000', pt: 2 } });
generatedDifferences.addSlide().addTable([['Short tuple']], {
  border: [{ type: 'solid', color: '00FF00', pt: 0 }],
});
~~~

Require PptxGenJS empty values to become four gray `666666`, 1pt solid lines; require native `{}` to remain absent/inherit; require PptxGenJS short-tuple missing sides to become none and its TRBL item `pt: 0` to become 1pt; require native tuple length to be exactly four and native zero to remain zero. Also require native omitted style to lack `prstDash`, while PptxGenJS omitted type/default line has direct solid. These are asserted differences, not parity failures.

- [ ] **Step 3: Run adapter and model suites**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts
~~~

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the adapter test and commit:

~~~text
test: compare table border defaults with pptxgenjs
~~~

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 5: Smoke the actual Node/browser/types package

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**

- Consumes: packed `@jiayunxie/pptx` Node/browser/type surfaces.
- Produces: `tableBorderCreation: true` in smoke JSON.

- [ ] **Step 1: Extend Node smoke**

Add a detached source border and pass it in the existing created-table options:

~~~js
const tableCreationBorderColor = { kind: 'scheme', value: 'accent4' };
const tableCreationBorder = {
  kind: 'line',
  color: tableCreationBorderColor,
  width: 1.5,
  style: 'dash',
};
// add to the existing table options:
border: tableCreationBorder,
~~~

Keep one empty-options cell that inherits, a cell scalar override, a cell TRBL override, and an explicit none cell. Capture all snapshots, mutate `tableCreationBorderColor.value` and `tableCreationBorder.width`, and require the live state unchanged. Call `setCellBorders(..., undefined)` on an inherited cell and require no re-inheritance; require the edited matrix after reopen.

Compute the flag from real assertions:

~~~js
const tableBorderCreation =
  allCreationLines(initialTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') &&
  allCreationLines(detachedTableDefaultBorders, 'scheme', 'accent4', 1.5, 'dash') &&
  createdTable.rows[1].cells[1].borders === undefined &&
  reopenedCreatedTable instanceof TableModel &&
  reopenedCreatedTable.rows[1].cells[1].borders === undefined;
~~~

Expose `tableBorderCreation` beside `tableCellBorderCreation` in the smoke JSON.

- [ ] **Step 2: Extend browser and declaration smoke**

Mirror table-border inheritance, detachment, whole-value cell override, clear, and reopen in the browser fixture. In the declaration fixture use:

~~~ts
const tableOptions: AddTableOptions = {
  border: snapshotCellBorders,
  fill: snapshotCellFill,
  margin: cellMargins,
  valign: cellAlignment,
};
~~~

Require the browser fixture to throw on any table-border snapshot mismatch.

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

~~~sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_border_package_dir=$(mktemp -d /tmp/pptx-table-border-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_border_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_border_package_dir/jiayunxie-pptx-0.1.0.tgz"
~~~

Require every smoke flag true, including `tableBorderCreation`, `tableCellBorderCreation`, `tableFillCreation`, `tableMarginCreation`, `tableVerticalAlignmentCreation`, types, browser, and CLI 0.1.0.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the smoke script and commit:

~~~text
test: smoke packed table border defaults
~~~

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 6: Document the public contract and remaining boundary

**Files:**

- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**

- Consumes: completed runtime/package evidence.
- Produces: accurate table-level border creation and remaining unsupported documentation.

- [ ] **Step 1: Update examples and contract**

Use this public example in the API/package documentation:

~~~ts
slide.addTable([[
  'Inherits all four table sides',
  { text: 'Cell blocks the table default', options: {
    border: { bottom: {
      kind: 'line',
      color: { kind: 'srgb', value: '70AD47' },
      width: 3,
      style: 'dash',
    } },
  } },
]], {
  border: {
    kind: 'line',
    color: { kind: 'scheme', value: 'accent1' },
    width: 1.5,
    style: 'solid',
  },
});
~~~

Document scalar/TRBL/named forms; sRGB/theme, finite 0..1584pt width, optional solid/dash; table-then-cell whole-value precedence; physical-cell materialization; empty/all-undefined input semantics; deep detachment; direct clear/replace without re-inheritance; and fixed margins/anchor/L/R/T/B/fill serialization order.

- [ ] **Step 2: Update compatibility and unsupported lists**

Mark `AddTableOptions.border` creation supported and PptxGenJS scalar/TRBL final-state compatible for explicit none/solid/dash/color/width/cell override. Document empty default, style omission, tuple zero, short tuple, named/theme native extensions. Keep table border getter/editor, diagonal/advanced line, transparency, shared-edge/style resolution, horizontal alignment, direction/fit creation, merge, hyperlink, rich text, auto-page, repeated headers, content measurement, and layout recomputation unsupported.

- [ ] **Step 3: Scan contradictions and run typecheck**

~~~sh
git diff --check
rg -n --pcre2 \
  'table(-level)? (border|边框).*(unsupported|尚未支持)|AddTableOptions.*border.*unsupported' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
~~~

Review every hit in context: creation support is valid; table-level read/edit and advanced border semantics remain unsupported.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

Stage only the four documents and commit:

~~~text
docs: document table border creation
~~~

Push, fetch, and require `origin/main...HEAD = 0 0`.

---

### Task 7: Run full gates and real-deck QA

**Files:**

- Review every Task 1–6 path; never stage or delete `.pnpm-store/`.

**Interfaces:**

- Consumes: implementation, tests, docs, actual tarball, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces: a verified set of pushed commits; any defect receives its own focused fix/review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-border-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-border-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
~~~

- [ ] **Step 2: Rebuild and smoke the actual tarball**

~~~sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_border_qa_package_dir=$(mktemp -d /tmp/pptx-table-border-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_border_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_border_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
~~~

- [ ] **Step 3: Generate nine real native and PptxGenJS decks**

Create `/tmp/pptx-table-border-qa.mjs` with `apply_patch` using this complete source:

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
const output = '/tmp/pptx-table-border';
await mkdir(output, { recursive: true });

const rows = [
  ['Inherited A', 'Inherited B', 'Inherited C', 'Inherited D'],
  ['Inherited E', 'Inherited F', 'Inherited G', 'Inherited H'],
];
const baseOptions = {
  name: 'Table border creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(1.5), inches(2.5), inches(2), inches(2)],
  rowHeights: [inches(0.8), inches(1.6)],
  fill: { kind: 'solid', color: { kind: 'srgb', value: 'F2F2F2' } },
  margin: 7.2,
  valign: 'middle',
};

function build(tableRows, options = {}) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const table = document.addSlide().addTable(
    tableRows,
    { ...baseOptions, ...options },
  );
  return { document, table };
}

function matrix(table) {
  return table.rows.map(({ cells }) => cells.map(({ borders }) => borders));
}

const omitted = build(rows);
const runtimeUndefined = build(rows, { border: undefined });
await writeFile(output + '/omitted.pptx', await omitted.document.write());
await writeFile(output + '/undefined.pptx', await runtimeUndefined.document.write());

const scalarBorder = {
  kind: 'line',
  color: { kind: 'srgb', value: '4472C4' },
  width: 1.5,
  style: 'dash',
};
const scalar = build(rows, { border: scalarBorder });
assert.ok(scalar.table.rows.every(({ cells }) => cells.every(({ borders }) =>
  ['top', 'right', 'bottom', 'left'].every((side) =>
    borders?.[side]?.kind === 'line' &&
    borders[side].color.kind === 'srgb' &&
    borders[side].color.value === '4472C4' &&
    borders[side].width === 1.5 &&
    borders[side].style === 'dash'))));
await writeFile(output + '/table-scalar.pptx', await scalar.document.write());

const tupleBorder = [
  { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 0, style: 'solid' },
  { kind: 'none' },
  { kind: 'line', color: { kind: 'srgb', value: '70AD47' }, width: 2, style: 'dash' },
  { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 3 },
];
const tuple = build(rows, { border: tupleBorder });
assert.deepEqual(tuple.table.rows[0].cells[0].borders, {
  top: tupleBorder[0], right: tupleBorder[1],
  bottom: tupleBorder[2], left: tupleBorder[3],
});
await writeFile(output + '/table-tuple.pptx', await tuple.document.write());

const mixedRows = [
  [
    'Inherited string',
    { text: 'Inherited empty', options: { border: {} } },
    { text: 'Green bottom only', options: { border: {
      bottom: {
        kind: 'line', color: { kind: 'srgb', value: '70AD47' },
        width: 3, style: 'dash',
      },
    } } },
    { text: 'Direct none', options: { border: { kind: 'none' } } },
  ],
  [
    { text: 'Zero top only', options: { border: {
      top: {
        kind: 'line', color: { kind: 'srgb', value: 'FF0000' },
        width: 0, style: 'solid',
      },
    } } },
    { text: 'Inherited F', options: { fill: { kind: 'none' } } },
    { text: 'Inherited G', options: { margin: [1, 2, 3, 4] } },
    { text: 'Inherited H', options: { valign: 'bottom' } },
  ],
];
const source = build(mixedRows, { border: scalarBorder });
const sourceBorders = matrix(source.table);
assert.deepEqual(sourceBorders[0][0], sourceBorders[0][1]);
assert.equal(sourceBorders[0][2].bottom.color.value, '70AD47');
assert.equal(sourceBorders[0][2].top.kind, 'none');
assert.ok(Object.values(sourceBorders[0][3]).every(({ kind }) => kind === 'none'));
assert.equal(sourceBorders[1][0].top.width, 0);
assert.equal(sourceBorders[1][0].right.kind, 'none');
await writeFile(output + '/mixed-source.pptx', await source.document.write());

source.table.setCellBorders(0, 0, undefined);
source.table.setCellBorders(0, 1, {
  right: {
    kind: 'line', color: { kind: 'scheme', value: 'accent2' },
    width: 4, style: 'solid',
  },
});
source.table.setCellText(0, 1, 'Edited inherited border');
const editedBytes = await source.document.write();
await writeFile(output + '/mixed-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[0].cells[0].borders, undefined);
assert.deepEqual(reopenedTable.rows[0].cells[1].borders, {
  right: {
    kind: 'line', color: { kind: 'scheme', value: 'accent2' },
    width: 4, style: 'solid',
  },
});
assert.equal(reopenedTable.rows[0].cells[1].text, 'Edited inherited border');
assert.deepEqual(reopenedTable.columnWidths, baseOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, baseOptions.rowHeights);
await writeFile(output + '/mixed-reopened.pptx', await reopened.write());

const generatedScalar = new PptxGenJS();
generatedScalar.layout = 'LAYOUT_WIDE';
generatedScalar.addSlide().addTable(rows, {
  x: 0.75, y: 0.75, w: 8, h: 2.4,
  colW: [1.5, 2.5, 2, 2], rowH: [0.8, 1.6],
  border: { type: 'dash', color: '4472C4', pt: 1.5 },
  fill: { color: 'F2F2F2' }, margin: 0.1, valign: 'middle',
});
await writeFile(
  output + '/pptxgenjs-scalar.pptx',
  await generatedScalar.write({ outputType: 'uint8array', compression: true }),
);

const generatedTuple = new PptxGenJS();
generatedTuple.layout = 'LAYOUT_WIDE';
generatedTuple.addSlide().addTable(rows, {
  x: 0.75, y: 0.75, w: 8, h: 2.4,
  colW: [1.5, 2.5, 2, 2], rowH: [0.8, 1.6],
  border: [
    { type: 'solid', color: 'FF0000', pt: 1 },
    { type: 'none' },
    { type: 'dash', color: '70AD47', pt: 2 },
    { type: 'solid', color: '4472C4', pt: 3 },
  ],
  fill: { color: 'F2F2F2' }, margin: 0.1, valign: 'middle',
});
await writeFile(
  output + '/pptxgenjs-tuple.pptx',
  await generatedTuple.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceBorders,
  reopenedBorders: matrix(reopenedTable),
  editedText: reopenedTable.rows[0].cells[1].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
~~~

Run `node /tmp/pptx-table-border-qa.mjs` and require all assertions to pass.

- [ ] **Step 4: Validate and diff packages**

~~~sh
for deck in /tmp/pptx-table-border/*.pptx; do
  node packages/pptx/dist/cli.js --json package validate \
    "$deck" --profile powerpoint-2010
done
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-border/omitted.pptx \
  /tmp/pptx-table-border/undefined.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-border/mixed-source.pptx \
  /tmp/pptx-table-border/mixed-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-border/mixed-edited.pptx \
  /tmp/pptx-table-border/mixed-reopened.pptx
~~~

Require zero errors/warnings. Omitted/undefined must have zero changed parts; source/edited may change only `/ppt/slides/slide1.xml`; edited/reopened must have zero changed parts.

- [ ] **Step 5: Render, overflow-check, and inspect**

Convert `table-scalar.pptx`, `table-tuple.pptx`, `mixed-source.pptx`, `mixed-edited.pptx`, `pptxgenjs-scalar.pptx`, and `pptxgenjs-tuple.pptx` to isolated PDFs with LibreOffice. Rasterize every page with Poppler at 180 DPI. Run `slides_test.py` on mixed-edited and both PptxGenJS decks using the bundled presentations Python environment.

Inspect every PNG at original detail. Require visible scalar/TRBL inheritance, whole-value cell overrides, clear/edit behavior, zero/none semantics where visible, preserved fill/vertical alignment/margins, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

Review strict normalization, descriptor safety, whole-value precedence, direct ownership, serialization order, omission, declarations, lifecycle isolation, PptxGenJS differences, smoke flags, docs boundaries, and invalid-input no-mutation assertions.

~~~sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
~~~

Require `0 0` and only `?? .pnpm-store/`. QA-only success creates no empty commit.
