# Existing Table Column Width Editing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, detached table-column-width reading and atomic lossless editing that keeps tblGrid and table transform width synchronized.

**Architecture:** Put input normalization, strict direct-child OOXML discovery, and minimal attribute patching in one focused internal helper. TableModel exposes the public getter/method and owns the package transaction; SDK, PptxGenJS, packed-consumer, validator, and visual tests prove the full lifecycle without changing existing setTransform() semantics.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless source-span OOXML editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, repository pptx-inspect CLI, LibreOffice headless, Poppler.

## Global Constraints

- Public API is exactly get columnWidths(): readonly number[] | undefined and setColumnWidths(value: number | readonly number[]): void.
- All public values are EMU; scalar broadcasts, array length exactly matches the current physical tblGrid column count.
- Input arrays are detached, dense, descriptor-safe Array.isArray() values; accessors, holes, extra keys, symbols, typed arrays, invalid numbers, unsafe rounding, and unsafe sums are rejected before package writes.
- Getter follows only the unique direct graphicFrame -> graphic -> graphicData -> tbl -> tblGrid path and returns undefined for any incomplete or ambiguous grid.
- Setter requires a unique direct graphicFrame -> xfrm -> ext@cx path and throws ModelParseError for unsafe grid or transform XML.
- Grid widths are positive safe-integer decimal values; existing cx may be zero but must otherwise be an unqualified non-negative safe-integer decimal token.
- Setter changes only differing gridCol@w and ext@cx attribute values, preserving all unrelated XML bytes and package parts.
- Numeric no-op preserves leading-zero tokens, slide bytes, and the mutation journal.
- A valid grid/transform mismatch is repaired by setting cx to the normalized grid sum.
- Merge markup never affects the physical column count and is never mutated.
- Do not change setTransform() behavior, add row-height editing, resize ratios, merge APIs, auto-layout, measurement, or pagination in this small item.
- Do not stage, remove, inspect recursively, or otherwise modify the untracked .pnpm-store/.
- Use inline execution in the current task because the user explicitly delegated decisions but did not authorize subagent work.
- Commit and push only after source review, typecheck, focused/full tests, performance, actual tarball smoke, CLI validation, PowerPoint 2010 validation, LibreOffice rendering, overflow checks, and original-resolution visual inspection pass.

---

### Task 1: Build the strict lossless column-grid helper with TDD

**Files:**
- Create: packages/model/src/table-column-widths.internal.ts
- Create: packages/model/src/table-column-widths.internal.test.ts

**Interfaces:**
- Consumes: LosslessXmlDocument, XmlElement, XmlAttribute, and ModelParseError.
- Produces:

~~~ts
export type NormalizedTableColumnWidthInput =
  | { readonly kind: 'scalar'; readonly value: number }
  | { readonly kind: 'array'; readonly values: readonly number[]; readonly sum: number };

export function normalizeTableColumnWidthInput(
  value: unknown,
): NormalizedTableColumnWidthInput;

export function readTableColumnWidths(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): readonly number[] | undefined;

export function replaceTableColumnWidths(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  input: NormalizedTableColumnWidthInput,
  partUri: string,
): boolean;
~~~

- [ ] **Step 1: Write focused failing tests for strict reads**

Use a local frame fixture:

~~~ts
function parseFrame(
  grid = '<a:tblGrid><a:gridCol w="914400"/><a:gridCol w="1828800"/></a:tblGrid>',
  transform = '<p:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></p:xfrm>',
): { xml: LosslessXmlDocument; frame: XmlElement } {
  const source =
    '<p:graphicFrame xmlns:p="p" xmlns:a="a">' + transform +
    '<a:graphic><a:graphicData><a:tbl>' + grid +
    '<a:tr h="914400"><a:tc gridSpan="2"><a:tcPr/></a:tc>' +
    '<a:tc hMerge="1"><a:tcPr/></a:tc></a:tr>' +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  const xml = LosslessXmlDocument.parse(source);
  return { xml, frame: xml.roots[0]! };
}

const valid = parseFrame();
expect(readTableColumnWidths(valid.xml, valid.frame)).toEqual([914_400, 1_828_800]);
const snapshot = readTableColumnWidths(valid.xml, valid.frame) as number[];
snapshot[0] = 1;
expect(readTableColumnWidths(valid.xml, valid.frame)).toEqual([914_400, 1_828_800]);
~~~

Require undefined for empty/missing/repeated/nested-only tblGrid and each malformed grid below:

~~~ts
[
  '<a:tblGrid/>',
  '<a:tblGrid><a:gridCol/></a:tblGrid>',
  '<a:tblGrid><a:gridCol w="0"/></a:tblGrid>',
  '<a:tblGrid><a:gridCol w="-1"/></a:tblGrid>',
  '<a:tblGrid><a:gridCol w="+1"/></a:tblGrid>',
  '<a:tblGrid><a:gridCol w="1.5"/></a:tblGrid>',
  '<a:tblGrid><a:gridCol w="1e3"/></a:tblGrid>',
  '<a:tblGrid><a:gridCol w="9007199254740992"/></a:tblGrid>',
  '<a:tblGrid><a:gridCol x:w="914400" xmlns:x="x"/></a:tblGrid>',
  '<a:tblGrid><a:gridCol w="914400" w="1828800"/></a:tblGrid>',
  '<a:tblGrid><x:keep xmlns:x="urn:test"><a:gridCol w="914400"/></x:keep></a:tblGrid>',
]
~~~

Add repeated direct graphic, graphicData, and tbl cases. Include one alternate-prefix valid path, one unknown direct tblGrid child, and gridSpan/hMerge/vMerge cells; only direct gridCol elements define the returned vector.

- [ ] **Step 2: Write focused failing normalization tests**

~~~ts
expect(normalizeTableColumnWidthInput(914_400.4)).toEqual({
  kind: 'scalar',
  value: 914_400,
});
expect(normalizeTableColumnWidthInput([914_400.4, 1_828_800.6])).toEqual({
  kind: 'array',
  values: [914_400, 1_828_801],
  sum: 2_743_201,
});
~~~

Reject undefined, null, strings, booleans, objects, typed arrays, empty arrays, holes, extra own keys, symbol keys, accessors, NaN, Infinity, zero, negative values, unsafe rounded values, and [Number.MAX_SAFE_INTEGER, 1]. Assert accessor call count is zero.

- [ ] **Step 3: Write focused failing minimal-write tests**

~~~ts
const changed = parseFrame(
  '<a:tblGrid keep="GRID"><a:gridCol w="0914400" keep="A"/>' +
  '<x:opaque xmlns:x="urn:test">KEEP</x:opaque>' +
  '<a:gridCol w="1828800" keep="B"/></a:tblGrid>',
  '<p:xfrm keep="XFRM"><a:off x="0" y="0"/>' +
  '<a:ext cx="2743200" cy="914400" keep="EXT"/></p:xfrm>',
);
expect(replaceTableColumnWidths(
  changed.xml,
  changed.frame,
  normalizeTableColumnWidthInput([914_400, 2_743_200]),
  '/ppt/slides/slide1.xml',
)).toBe(true);
expect(changed.xml.serialize()).toContain(
  '<a:gridCol w="0914400" keep="A"/><x:opaque xmlns:x="urn:test">KEEP</x:opaque>' +
  '<a:gridCol w="2743200" keep="B"/>',
);
expect(changed.xml.serialize()).toContain(
  '<a:ext cx="3657600" cy="914400" keep="EXT"/>',
);
~~~

Cover scalar broadcast, exact arrays, one-item arrays not broadcasting, wrong lengths, scalar sum overflow, cx="0" repair, valid mismatch repair, and leading-zero semantic no-op. Require no-op false and xml.changed false.

Require ModelParseError with the slide URI for missing/repeated/nested-only direct xfrm, missing/repeated/nested-only direct ext, and cx values -1, 1.5, 1e3, 9007199254740992, repeated cx, or namespaced-only x:cx.

- [ ] **Step 4: Run the new test file and confirm red**

~~~sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-column-widths.internal.test.ts
~~~

Expected: FAIL because the helper module or exports do not exist.

- [ ] **Step 5: Implement descriptor-safe normalization**

Use Reflect.ownKeys, Object.getOwnPropertyDescriptor, and Object.hasOwn(descriptor, 'value'). Normalize numbers with:

~~~ts
function normalizePositiveDimension(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(context + ' must be finite');
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(context + ' must round to a safe integer EMU value');
  }
  if (rounded <= 0) {
    throw new RangeError(context + ' must be greater than zero');
  }
  return rounded;
}
~~~

Copy arrays only from own data descriptors. Check sum overflow before each addition. Never retain the caller array.

- [ ] **Step 6: Implement strict direct-path reading**

~~~ts
function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement =>
      child.type === 'element' && child.localName === localName,
  );
}

function exactUnqualifiedAttribute(
  element: XmlElement,
  name: string,
): XmlAttribute | undefined {
  const matches = element.attributes.filter((attribute) => attribute.name === name);
  return matches.length === 1 ? matches[0] : undefined;
}

function parseUnsignedSafeInteger(value: string, positive: boolean): number | undefined {
  if (!/^\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) return undefined;
  return positive ? (parsed > 0 ? parsed : undefined) : parsed;
}
~~~

Resolve a level only when its direct-child count is exactly one. Require at least one direct gridCol and one exact unqualified valid w on every column. Ignore but preserve unknown non-gridCol children.

- [ ] **Step 7: Implement validation-before-patching replacement**

Resolve the complete grid, materialize scalar or exact vector, compute a safe sum, then resolve and validate the full transform. Do not patch until every check passes.

~~~ts
let changed = false;
for (let index = 0; index < columns.length; index += 1) {
  if (widths[index] !== target[index]) {
    xml.replaceAttribute(widthAttributes[index]!, String(target[index]));
    changed = true;
  }
}
if (currentCx !== targetSum) {
  xml.replaceAttribute(cxAttribute, String(targetSum));
  changed = true;
}
return changed;
~~~

Throw ModelParseError for malformed existing XML, TypeError for vector length, and RangeError for scalar expansion overflow.

- [ ] **Step 8: Run focused helper tests**

~~~sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-column-widths.internal.test.ts
~~~

Expected: all read, descriptor, lossless patch, no-op, mismatch, merge, and malformed cases pass.

---

### Task 2: Wire the public TableModel lifecycle

**Files:**
- Modify: packages/model/src/shapes.ts
- Modify: packages/model/src/model.test.ts

**Interfaces:**
- Consumes: Task 1 helper exports, SlideModel.resolveShape(), SlideModel.setXml(), and OpcPackage.transaction().
- Produces: live public TableModel.columnWidths and TableModel.setColumnWidths().

- [ ] **Step 1: Add failing public lifecycle tests**

~~~ts
expect(table.columnWidths).toEqual([inches(1), inches(2), inches(3)]);
const detached = table.columnWidths as number[];
detached[0] = 1;
expect(table.columnWidths).toEqual([inches(1), inches(2), inches(3)]);

table.setColumnWidths([inches(1.5), inches(2.5), inches(2)]);
expect(table.columnWidths).toEqual([inches(1.5), inches(2.5), inches(2)]);
expect(table.transform.width).toBe(inches(6));
expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
~~~

Save bytes/journal, repeat the same setter, and require a no-op. In an outer package transaction set scalar inches(2), throw, and require bytes, journal, widths, transform, and identity to restore.

Make ext@cx one inch while preserving a valid grid. Getter must return grid values; setting the same vector repairs only cx. Add gridSpan="2", hMerge="1", vMerge="1", and an opaque grid child; edit widths and require every merge/opaque token to survive.

- [ ] **Step 2: Add invalid-input and malformed-XML isolation**

Reject these values without bytes/journal changes:

~~~ts
[
  undefined,
  null,
  [],
  [1],
  [1, 2],
  [1, 2, 3, 4],
  new Uint32Array([1, 2, 3]),
  Number.NaN,
  Number.POSITIVE_INFINITY,
  0,
  -1,
]
~~~

Add hole, accessor, extra-key, symbol-key, unsafe-item, and sum-overflow arrays. For malformed grid require getter undefined and setter ModelParseError; for malformed transform require valid getter and setter ModelParseError.

- [ ] **Step 3: Run model tests and confirm the public API is red**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-column-widths.internal.test.ts \
  packages/model/src/model.test.ts
~~~

- [ ] **Step 4: Add the getter and transactional setter**

~~~ts
get columnWidths(): readonly number[] | undefined {
  const { xml, element } = this.resolve();
  return readTableColumnWidths(xml, element);
}

setColumnWidths(value: number | readonly number[]): void {
  const input = normalizeTableColumnWidthInput(value);
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableColumnWidths(xml, element, input, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
~~~

Normalize before resolve; re-resolve inside the transaction.

- [ ] **Step 5: Run model tests and typecheck**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-column-widths.internal.test.ts \
  packages/model/src/model.test.ts
~~~

Expected: declarations expose both members and all lifecycle/isolation assertions pass.

---

### Task 3: Prove SDK, PptxGenJS, browser, and packed-package behavior

**Files:**
- Modify: packages/sdk/src/index.test.ts
- Modify: packages/pptxgenjs-adapter/src/index.test.ts
- Modify: scripts/smoke-npm-package.mjs

**Interfaces:**
- Consumes: Task 2 API, native creation, PptxGenJS 4.0.1 output, aggregate Node/browser exports, and declarations.
- Produces: existing-deck editing, mismatch repair, reopen proof, and installed-consumer coverage.

- [ ] **Step 1: Extend the SDK table lifecycle**

Assert initial/duplicate [2, 4, 2] inch snapshots. Edit only the original:

~~~ts
table.setColumnWidths([inches(1.5), inches(4.5), inches(2)]);
expect(table.columnWidths).toEqual([inches(1.5), inches(4.5), inches(2)]);
expect(table.transform.width).toBe(inches(8));
expect(duplicateTable.columnWidths).toEqual([inches(2), inches(4), inches(2)]);
~~~

Inside rollback, set scalar inches(1) before throwing and verify restoration. After reopen require edited original and untouched duplicate vectors, exact transform sums, and unchanged non-slide parts.

- [ ] **Step 2: Add PptxGenJS imported-grid editing**

In basic table conformance:

~~~ts
expect(importedTable.columnWidths).toEqual([
  inches(1),
  inches(2),
  inches(3),
]);
importedTable.setColumnWidths([inches(1.5), inches(1.5), inches(3)]);
expect(importedTable.columnWidths).toEqual([
  inches(1.5),
  inches(1.5),
  inches(3),
]);
expect(importedTable.transform.width).toBe(inches(6));
~~~

Require reopen persistence. Add a public PptxGenJS table with options { x: 1, y: 1, w: 5, h: 1, colW: [1, 2, 3] }. Require grid sum 6 inches but transform 5 inches; setting the same getter snapshot repairs transform to 6 while grid tokens remain unchanged.

- [ ] **Step 3: Extend Node smoke**

~~~js
if (
  createdTable.columnWidths?.join(',') !==
  [inches(1), inches(3)].join(',')
) throw new Error('Table column-width read failed');
createdTable.setColumnWidths([inches(1.5), inches(2.5)]);
if (
  createdTable.columnWidths?.join(',') !==
    [inches(1.5), inches(2.5)].join(',') ||
  createdTable.transform.width !== inches(4)
) throw new Error('Table column-width edit failed');
~~~

Require reopen widths and existing cell/row-height/isolation behavior. Add tableColumnWidthEditing: true to emitted JSON.

- [ ] **Step 4: Extend browser and declaration smoke**

Browser:

~~~js
createdTable.setColumnWidths([inches(1), inches(1.5)]);
~~~

Require transform 2.5 inches and reopen vector [914400, 1371600].

Installed TypeScript:

~~~ts
const widthSnapshot: readonly number[] | undefined = typedTable.columnWidths;
typedTable.setColumnWidths(inches(2));
typedTable.setColumnWidths([inches(1.5), inches(2.5)]);
~~~

Keep widthSnapshot in the final void expression.

- [ ] **Step 5: Run focused conformance**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-column-widths.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
~~~

---

### Task 4: Document the editing slice accurately

**Files:**
- Modify: CHANGELOG.md
- Modify: docs/api/README.md
- Modify: docs/compatibility/pptxgenjs-baseline.md
- Modify: packages/pptx/README.md

**Interfaces:**
- Consumes: final names, malformed behavior, synchronized transform invariant, and verified PptxGenJS difference.
- Produces: public docs while row-height editing and broader table gaps remain explicitly unsupported.

- [ ] **Step 1: Update API example and behavior**

~~~ts
console.log(table.columnWidths?.map(emuToInches));
table.setColumnWidths([
  inches(2),
  inches(3),
  inches(3),
]);
~~~

Document detached snapshots, scalar broadcast, exact length, descriptor safety, safe sums, atomic synchronization, mismatch repair, semantic no-op, getter undefined, and setter ModelParseError. State that setTransform({ width }) changes only transform.

- [ ] **Step 2: Update compatibility baseline**

Move existing-table column widths from unsupported to supported. State that PptxGenJS has creation-only colW, while native lossless editing repairs cx mismatch. Preserve scalar-floor, single-array, length fallback, and coercion differences.

- [ ] **Step 3: Update README and changelog**

Add one Unreleased bullet and getter/setter examples. Preserve row-height reading/editing, merge editing, rich/multi-paragraph cell text, auto-page, repeated headers, hyperlinks, and content measurement in unsupported lists.

- [ ] **Step 4: Run docs checks**

~~~sh
git diff --check
rg -n "columnWidths|setColumnWidths|row-height|row height|existing-table|已有表格|部分支持" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
~~~

---

### Task 5: Full gates, real-deck QA, review, commit, and push

**Files:**
- Review every Task 1-4 path; never stage or delete .pnpm-store/.

**Interfaces:**
- Consumes: source/tests/docs, aggregate package, CLI, PptxGenJS 4.0.1, LibreOffice, Poppler, and overflow helper.
- Produces: reviewed feat: edit table column widths commit synchronized to origin/main.

- [ ] **Step 1: Run full functional and performance gates**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
~~~

Record exact passed/skipped counts and inspect any changed snapshot or timing.

- [ ] **Step 2: Build and smoke the actual npm tarball**

~~~sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_column_edit_package_dir=$(mktemp -d /tmp/pptx-table-column-edit-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_column_edit_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_column_edit_package_dir/jiayunxie-pptx-0.1.0.tgz"
~~~

Expected JSON includes tableCreation, tableColumnWidths, tableColumnWidthEditing, tableRowHeights, types, browser success, and CLI 0.1.0.

- [ ] **Step 3: Verify inspection runtime**

~~~sh
command -v pptx-inspect
pptx-inspect --json doctor
~~~

- [ ] **Step 4: Generate source, edited, reopened, and baseline decks**

Create /tmp/pptx-table-column-width-editing-qa.mjs with apply_patch and this complete program:

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
const output = '/tmp/pptx-table-column-width-editing';
await mkdir(output, { recursive: true });

const rows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];
const sourceWidths = [inches(2), inches(4), inches(2)];
const targetWidths = [inches(1.5), inches(4.5), inches(2)];

const native = PptxDocument.create({ slideSize: 'wide' });
const slide = native.addSlide();
const table = slide.addTable(rows, {
  name: 'Column width editing QA',
  x: inches(1),
  y: inches(1),
  columnWidths: sourceWidths,
  rowHeights: [inches(0.6), inches(0.9), inches(0.9)],
});
table.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
});
table.setCellBorders(0, 0, {
  kind: 'line',
  color: { kind: 'srgb', value: 'FFFFFF' },
  width: 1,
});
table.setCellTextDirection(1, 2, 'vert270');
table.setCellVerticalAlignment(2, 1, 'bottom');
await writeFile(output + '/source.pptx', await native.write());

table.setColumnWidths(targetWidths);
assert.deepEqual(table.columnWidths, targetWidths);
assert.equal(table.transform.width, inches(8));
const editedBytes = await native.write();
await writeFile(output + '/native-edited.pptx', editedBytes);

const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.deepEqual(reopenedTable.columnWidths, targetWidths);
assert.equal(reopenedTable.transform.width, inches(8));
assert.equal(reopenedTable.rows[1].cells[2].textDirection, 'vert270');
assert.equal(reopenedTable.rows[2].cells[2].text, '');
await writeFile(output + '/native-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  rows.map((row, rowIndex) => row.map((text, columnIndex) => ({
    text,
    options: {
      fill: rowIndex === 0 && columnIndex === 0 ? '4472C4' : 'FFFFFF',
      color: rowIndex === 0 && columnIndex === 0 ? 'FFFFFF' : '000000',
      border: rowIndex === 0 && columnIndex === 0
        ? { type: 'solid', color: 'FFFFFF', pt: 1 }
        : { type: 'none' },
      valign: rowIndex === 2 && columnIndex === 1 ? 'bottom' : 'middle',
      textDirection: rowIndex === 1 && columnIndex === 2 ? 'vert270' : 'horz',
    },
  }))),
  {
    x: 1,
    y: 1,
    w: 8,
    h: 2.4,
    colW: [1.5, 4.5, 2],
    rowH: [0.6, 0.9, 0.9],
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  widths: reopenedTable.columnWidths,
  transformWidth: reopenedTable.transform.width,
  emptyCell: reopenedTable.rows[2].cells[2].text === '',
  verticalCell: reopenedTable.rows[1].cells[2].textDirection,
}));
~~~

Run node /tmp/pptx-table-column-width-editing-qa.mjs. Expected widths are [1371600,4114800,1828800], transform width 7315200, emptyCell true, verticalCell vert270.

- [ ] **Step 5: Validate packages and mutation isolation**

~~~sh
pptx-inspect --json package validate /tmp/pptx-table-column-width-editing/source.pptx --profile powerpoint-2010
pptx-inspect --json package validate /tmp/pptx-table-column-width-editing/native-edited.pptx --profile powerpoint-2010
pptx-inspect --json package validate /tmp/pptx-table-column-width-editing/native-reopened.pptx --profile powerpoint-2010
pptx-inspect --json package validate /tmp/pptx-table-column-width-editing/pptxgenjs-baseline.pptx --profile powerpoint-2010
pptx-inspect --json package diff /tmp/pptx-table-column-width-editing/source.pptx /tmp/pptx-table-column-width-editing/native-edited.pptx
pptx-inspect --json package diff /tmp/pptx-table-column-width-editing/native-edited.pptx /tmp/pptx-table-column-width-editing/native-reopened.pptx
~~~

Require zero errors/warnings. Source-to-edited may change only /ppt/slides/slide1.xml; edited-to-reopened decompressed parts must be unchanged.

- [ ] **Step 6: Render and inspect all output**

~~~sh
mkdir -p /tmp/pptx-table-column-width-editing/rendered-native
mkdir -p /tmp/pptx-table-column-width-editing/rendered-baseline
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-column-width-editing/rendered-native \
  /tmp/pptx-table-column-width-editing/native-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-column-width-editing/rendered-baseline \
  /tmp/pptx-table-column-width-editing/pptxgenjs-baseline.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 /tmp/pptx-table-column-width-editing/rendered-native/native-edited.pdf \
  /tmp/pptx-table-column-width-editing/rendered-native/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 /tmp/pptx-table-column-width-editing/rendered-baseline/pptxgenjs-baseline.pdf \
  /tmp/pptx-table-column-width-editing/rendered-baseline/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-column-width-editing/native-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-column-width-editing/pptxgenjs-baseline.pptx
~~~

Open every PNG at original detail. Require visible 1.5:4.5:2 proportions, exact bounds, readable headers, vertical cell, bottom alignment, empty final cell, and no repair, clipping, overlap, blur, missing cell, or off-slide content.

- [ ] **Step 7: Review complete diff**

~~~sh
git diff --check
git status --short
git diff --stat
git diff -- \
  packages/model/src/table-column-widths.internal.ts \
  packages/model/src/table-column-widths.internal.test.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
~~~

Map every design requirement to a test or QA artifact. Confirm .pnpm-store/ is absent from staged changes.

- [ ] **Step 8: Commit only reviewed implementation**

~~~sh
git add \
  packages/model/src/table-column-widths.internal.ts \
  packages/model/src/table-column-widths.internal.test.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "feat: edit table column widths"
~~~

- [ ] **Step 9: Push and prove synchronization**

~~~sh
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o UserKnownHostsFile=/tmp/pptx-github-known-hosts -o StrictHostKeyChecking=yes' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
git status --short --branch
~~~

Expected: 0 0 and only ?? .pnpm-store/.

---

After synchronization, begin the separate existing-table row-height reading/editing design, plan, implementation, validation, review, commit, and push cycle without asking for routine decisions.
