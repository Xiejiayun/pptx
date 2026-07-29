# Existing Table Row Height Editing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict, detached table-row-height reading and atomic lossless editing that synchronizes transform height for fully explicit rows while preserving transform height when any row remains automatic.

**Architecture:** Put non-negative input normalization, strict direct-child OOXML discovery, and minimal `tr@h`/`ext@cy` patching in one focused internal helper. `TableModel` exposes the public getter/method and owns the package transaction; SDK, PptxGenJS, packed-consumer, validator, diff, and visual tests prove explicit and automatic-row lifecycles without changing creation or `setTransform()` semantics.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless source-span OOXML editing, OPC transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, repository pptx-inspect CLI, LibreOffice headless, Poppler.

## Global Constraints

- Public API is exactly `get rowHeights(): readonly number[] | undefined` and `setRowHeights(value: number | readonly number[]): void`.
- All public values are EMU; scalar broadcasts and array length exactly matches the current direct physical row count.
- Zero means OOXML automatic row height. Raw negative inputs are rejected even if rounding would produce JavaScript `-0`.
- Input arrays are detached, dense, descriptor-safe `Array.isArray()` values; accessors, holes, extra keys, symbols, typed arrays, invalid numbers, and unsafe rounding are rejected before package writes.
- Getter follows only the unique direct `graphicFrame -> graphic -> graphicData -> tbl -> tr@h` path and returns `undefined` for any incomplete or ambiguous row vector.
- Setter always requires a unique direct `graphicFrame -> xfrm -> ext@cy` path and throws `ModelParseError` for unsafe rows or transform XML.
- Existing `h` and `cy` tokens must be unqualified non-negative safe-integer decimals; leading zeros are valid and preserved on numeric no-op.
- Fully positive targets require a safe exact sum and set `cy` to that sum. Any target containing zero leaves valid `cy` byte-for-byte unchanged.
- Setter changes only differing direct `tr@h` and eligible `ext@cy` attribute values, preserving all unrelated XML bytes and package parts.
- Merge markup never affects the physical row count and is never mutated.
- `addTable({ rowHeights })` continues to reject explicit zero; omitting creation `rowHeights` remains the only creation-time automatic-row path.
- Do not change `setTransform()` behavior, add resize ratios, row insertion/deletion/reorder, merge APIs, auto-layout, measurement, repeated headers, or pagination in this small item.
- Do not stage, remove, inspect recursively, or otherwise modify the untracked `.pnpm-store/`.
- Use inline execution in the current task; the user explicitly requested autonomous progress without subagents.
- Commit and push only after source review, typecheck, focused/full tests, performance, actual tarball smoke, CLI validation, PowerPoint 2010 validation, LibreOffice rendering, overflow checks, and original-resolution visual inspection pass.

---

### Task 1: Build the strict lossless row-height helper with TDD

**Files:**
- Create: `packages/model/src/table-row-heights.internal.ts`
- Create: `packages/model/src/table-row-heights.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `XmlAttribute`, and `ModelParseError`.
- Produces:

~~~ts
export type NormalizedTableRowHeightInput =
  | { readonly kind: 'scalar'; readonly value: number }
  | { readonly kind: 'array'; readonly values: readonly number[] };

export function normalizeTableRowHeightInput(
  value: unknown,
): NormalizedTableRowHeightInput;

export function readTableRowHeights(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): readonly number[] | undefined;

export function replaceTableRowHeights(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  input: NormalizedTableRowHeightInput,
  partUri: string,
): boolean;
~~~

- [ ] **Step 1: Write focused failing tests for strict reads**

Use this local frame fixture:

~~~ts
function parseFrame(
  rows =
    '<a:tr h="0" keep="R1"><a:tc rowSpan="2"><a:tcPr/></a:tc></a:tr>' +
    '<a:tr h="1828800" keep="R2"><a:tc vMerge="1"><a:tcPr/></a:tc></a:tr>',
  transform =
    '<p:xfrm keep="X"><a:off x="0" y="0"/>' +
    '<a:ext cx="2743200" cy="1828800" keep="E"/></p:xfrm>',
): { xml: LosslessXmlDocument; frame: XmlElement } {
  const source =
    '<p:graphicFrame xmlns:p="p" xmlns:a="a">' + transform +
    '<a:graphic><a:graphicData><a:tbl>' +
    '<a:tblGrid><a:gridCol w="2743200"/></a:tblGrid>' + rows +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  const xml = LosslessXmlDocument.parse(source);
  return { xml, frame: xml.roots[0]! };
}

const valid = parseFrame();
expect(readTableRowHeights(valid.xml, valid.frame)).toEqual([0, 1_828_800]);
const snapshot = readTableRowHeights(valid.xml, valid.frame) as number[];
snapshot[0] = 1;
expect(readTableRowHeights(valid.xml, valid.frame)).toEqual([0, 1_828_800]);
~~~

Require `undefined` for no direct rows and every malformed row below:

~~~ts
[
  '<a:tr/>',
  '<a:tr h="-1"/>',
  '<a:tr h="+1"/>',
  '<a:tr h="1.5"/>',
  '<a:tr h="1e3"/>',
  '<a:tr h="9007199254740992"/>',
  '<a:tr x:h="914400" xmlns:x="x"/>',
  '<a:tr h="914400" h="1828800"/>',
]
~~~

Add missing/repeated/nested-only direct `graphic`, `graphicData`, and `tbl` cases. Include an alternate-prefix valid path, leading-zero heights, an unknown direct table child containing a nested fake `tr`, and `rowSpan`/`hMerge`/`vMerge` cells; only direct `tbl` rows define the returned vector.

- [ ] **Step 2: Write focused failing normalization tests**

~~~ts
expect(normalizeTableRowHeightInput(0)).toEqual({
  kind: 'scalar',
  value: 0,
});
expect(normalizeTableRowHeightInput(914_400.4)).toEqual({
  kind: 'scalar',
  value: 914_400,
});
expect(normalizeTableRowHeightInput([0.4, 914_400.6])).toEqual({
  kind: 'array',
  values: [0, 914_401],
});
~~~

Reject `undefined`, `null`, strings, booleans, objects, typed arrays, empty arrays, holes, extra own keys, symbol keys, accessors, `NaN`, infinities, `-1`, `-0.4`, unsafe rounded values, and arrays containing any invalid item. Assert accessor invocation count is zero.

- [ ] **Step 3: Write focused failing minimal-write tests**

~~~ts
const changed = parseFrame(
  '<a:tr h="0914400" keep="R1"><a:tc rowSpan="2"><a:tcPr/></a:tc></a:tr>' +
  '<x:opaque xmlns:x="urn:test"><a:tr h="999"/></x:opaque>' +
  '<a:tr h="1828800" keep="R2"><a:tc vMerge="1"><a:tcPr/></a:tc></a:tr>',
  '<p:xfrm keep="X"><a:off x="0" y="0"/>' +
  '<a:ext cx="2743200" cy="2743200" keep="E"/></p:xfrm>',
);
expect(replaceTableRowHeights(
  changed.xml,
  changed.frame,
  normalizeTableRowHeightInput([914_400, 2_743_200]),
  '/ppt/slides/slide1.xml',
)).toBe(true);
expect(changed.xml.serialize()).toContain('<a:tr h="0914400" keep="R1">');
expect(changed.xml.serialize()).toContain('<a:tr h="2743200" keep="R2">');
expect(changed.xml.serialize()).toContain('<a:ext cx="2743200" cy="3657600" keep="E"/>');
expect(changed.xml.serialize()).toContain('<x:opaque xmlns:x="urn:test"><a:tr h="999"/></x:opaque>');
~~~

Cover positive scalar broadcast, zero scalar broadcast, exact arrays, mixed zero/positive arrays, one-item arrays not broadcasting, wrong lengths, all-positive array and scalar expansion overflow, and leading-zero semantic no-op. Require numeric no-op to return false and leave `xml.changed` false.

For `[0, Number.MAX_SAFE_INTEGER, 1]`, require success when row count matches and require `cy` unchanged because zero disables sum derivation. For all-positive `[Number.MAX_SAFE_INTEGER, 1]`, require `RangeError` before any XML change.

For a mixed target, require every changed `tr@h` to update while `cy="02743200"` remains byte-identical. For an all-positive target whose rows already match but `cy` is different or zero, require only `cy` repair.

Require `ModelParseError` containing the slide URI for missing/repeated/nested-only direct `xfrm`, missing/repeated/nested-only direct `ext`, and `cy` values `-1`, `+1`, `1.5`, `1e3`, `9007199254740992`, repeated `cy`, or namespaced-only `x:cy`.

- [ ] **Step 4: Run the new test file and confirm red**

~~~sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-row-heights.internal.test.ts
~~~

Expected: FAIL because the helper module or exports do not exist.

- [ ] **Step 5: Implement descriptor-safe non-negative normalization**

Copy arrays only from own data descriptors using `Reflect.ownKeys()`, `Object.getOwnPropertyDescriptor()`, and `Object.hasOwn(descriptor, 'value')`. Normalize each item with:

~~~ts
function normalizeNonNegativeDimension(
  value: unknown,
  context: string,
): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(context + ' must be finite');
  }
  if (value < 0) {
    throw new RangeError(context + ' must be non-negative');
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(
      context + ' must round to a safe integer EMU value',
    );
  }
  return rounded === 0 ? 0 : rounded;
}
~~~

Never retain the caller array and never read an accessor value.

- [ ] **Step 6: Implement strict direct-path row reading**

Use direct-child local-name matching and exact unqualified attributes:

~~~ts
function directChildren(
  element: XmlElement,
  localName: string,
): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement =>
      child.type === 'element' && child.localName === localName,
  );
}

function exactUnqualifiedAttribute(
  element: XmlElement,
  name: string,
): XmlAttribute | undefined {
  const matches = element.attributes.filter(
    (attribute) => attribute.name === name,
  );
  return matches.length === 1 ? matches[0] : undefined;
}

function parseNonNegativeSafeInteger(value: string): number | undefined {
  if (!/^\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}
~~~

Resolve a path level only when its direct-child count is exactly one. Require at least one direct `tr` and one exact valid `h` on every row. Ignore but preserve unknown non-row direct children.

- [ ] **Step 7: Implement validation-before-patching replacement**

Resolve the complete row vector, materialize scalar or exact array, detect zero, and calculate a safe sum only when every target is positive. Resolve and validate the full transform before patching any attribute.

~~~ts
let changed = false;
for (let index = 0; index < rows.length; index += 1) {
  if (currentHeights[index] !== target[index]) {
    xml.replaceAttribute(heightAttributes[index]!, String(target[index]));
    changed = true;
  }
}
if (targetTotal !== undefined && currentCy !== targetTotal) {
  xml.replaceAttribute(cyAttribute, String(targetTotal));
  changed = true;
}
return changed;
~~~

Throw `ModelParseError` for malformed existing XML, `TypeError` for vector length, and `RangeError` for all-positive sum overflow. A zero-containing target must not call the sum helper and must not patch `cy`.

- [ ] **Step 8: Run focused helper tests**

~~~sh
node node_modules/vitest/vitest.mjs run packages/model/src/table-row-heights.internal.test.ts
~~~

Expected: all strict-read, descriptor, explicit, automatic, lossless patch, no-op, mismatch, merge, and malformed cases pass.

---

### Task 2: Wire the public TableModel lifecycle

**Files:**
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 helper exports, `SlideModel.resolveShape()`, `SlideModel.setXml()`, and `OpcPackage.transaction()`.
- Produces: live public `TableModel.rowHeights` and `TableModel.setRowHeights()`.

- [ ] **Step 1: Add failing public explicit-row lifecycle tests**

~~~ts
expect(table.rowHeights).toEqual([
  inches(0.5),
  inches(0.75),
  inches(1),
]);
const detached = table.rowHeights as number[];
detached[0] = 1;
expect(table.rowHeights).toEqual([
  inches(0.5),
  inches(0.75),
  inches(1),
]);

table.setRowHeights([
  inches(0.75),
  inches(1.25),
  inches(0.5),
]);
expect(table.transform.height).toBe(inches(2.5));
expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
~~~

Save bytes/journal, repeat the same setter, and require a no-op. In an outer package transaction set scalar `inches(1)`, throw, and require bytes, journal, heights, transform, and identity to restore. Make `ext@cy` one inch while preserving valid positive rows; getter must return rows and setting the same vector must repair only `cy`.

- [ ] **Step 2: Add failing public automatic-row lifecycle tests**

Create a table without `rowHeights` and require `[0, 0, 0]`. Set `[0, inches(1), 0]`, require row tokens to change, and require the pre-call transform height to remain exact. Set scalar zero and require every row zero with transform height still unchanged. Repeat numeric zero to prove leading-zero/no-op isolation.

Add `rowSpan="2"`, `hMerge="1"`, `vMerge="1"`, and an opaque direct table child containing a fake row; edit physical row heights and require every merge/opaque token to survive.

- [ ] **Step 3: Add invalid-input and malformed-XML isolation**

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
  -1,
  -0.4,
]
~~~

Add hole, accessor, extra-key, symbol-key, unsafe-item, and all-positive sum-overflow arrays. For malformed rows require getter `undefined` and setter `ModelParseError`; for malformed transform require valid getter and setter `ModelParseError` with zero mutation.

- [ ] **Step 4: Run model tests and confirm the public API is red**

~~~sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-row-heights.internal.test.ts \
  packages/model/src/model.test.ts
~~~

Expected: FAIL because `TableModel` does not expose the new members.

- [ ] **Step 5: Add the getter and transactional setter**

~~~ts
get rowHeights(): readonly number[] | undefined {
  const { xml, element } = this.resolve();
  return readTableRowHeights(xml, element);
}

setRowHeights(value: number | readonly number[]): void {
  const input = normalizeTableRowHeightInput(value);
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableRowHeights(xml, element, input, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
~~~

Normalize before resolve; re-resolve inside the transaction. Keep inherited `setTransform({ height })` unchanged and add a regression assertion that it changes only `cy`, not direct `tr@h`.

- [ ] **Step 6: Run model tests and typecheck**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-row-heights.internal.test.ts \
  packages/model/src/model.test.ts
~~~

Expected: declarations expose both members and every explicit/automatic lifecycle and isolation assertion passes.

---

### Task 3: Prove SDK, PptxGenJS, browser, and packed-package behavior

**Files:**
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 2 API, native creation, PptxGenJS 4.0.1 output, aggregate Node/browser exports, and declarations.
- Produces: existing-deck explicit/automatic editing, reopen proof, isolation, and installed-consumer coverage.

- [ ] **Step 1: Extend the SDK table lifecycle**

Assert initial and duplicate `[0.5, 0.75, 1]` inch snapshots. Edit only the original:

~~~ts
table.setRowHeights([
  inches(0.75),
  inches(1.25),
  inches(0.5),
]);
expect(table.rowHeights).toEqual([
  inches(0.75),
  inches(1.25),
  inches(0.5),
]);
expect(table.transform.height).toBe(inches(2.5));
expect(duplicateTable.rowHeights).toEqual([
  inches(0.5),
  inches(0.75),
  inches(1),
]);
~~~

Inside rollback, set scalar zero before throwing and verify restoration. Then set the original to `[0, inches(1), 0]` after capturing transform height, require that height unchanged, write/reopen, and require edited original and untouched duplicate vectors, exact transform behavior, and unchanged non-slide parts.

- [ ] **Step 2: Add PptxGenJS imported-row editing**

In basic table conformance:

~~~ts
expect(importedTable.rowHeights).toEqual([
  inches(0.75),
  inches(1.25),
]);
importedTable.setRowHeights([inches(1), inches(1.5)]);
expect(importedTable.rowHeights).toEqual([inches(1), inches(1.5)]);
expect(importedTable.transform.height).toBe(inches(2.5));
~~~

Require reopen persistence. Add a public PptxGenJS table with `{ x: 1, y: 1, w: 5, h: 1, rowH: [0.5, 1.5] }`; setting the same positive getter snapshot must repair transform height to two inches. Then set `[0, inches(1.5)]` and require the repaired two-inch transform height to remain unchanged.

- [ ] **Step 3: Extend Node smoke**

~~~js
const initialTableRowHeights = createdTable.rowHeights;
createdTable.setRowHeights([inches(0.75), inches(1.25)]);
if (
  createdTable.rowHeights?.join(',') !==
    [inches(0.75), inches(1.25)].join(',') ||
  createdTable.transform.height !== inches(2)
) throw new Error('Table row-height edit failed');
~~~

After positive synchronization, capture height, set `[0, inches(1.25)]`, require height preservation, then restore the positive vector before existing render/reopen checks. Require reopened heights and add `tableRowHeightEditing: true` to emitted JSON without weakening `tableRowHeights` creation coverage.

- [ ] **Step 4: Extend browser and declaration smoke**

Browser:

~~~js
createdTable.setRowHeights([inches(0.5), inches(1)]);
if (
  createdTable.rowHeights?.join(',') !==
    [inches(0.5), inches(1)].join(',') ||
  createdTable.transform.height !== inches(1.5)
) throw new Error('Browser table row-height editing failed');
~~~

Require the same vector after reopen. Installed TypeScript:

~~~ts
const heightSnapshot: readonly number[] | undefined = typedTable.rowHeights;
typedTable.setRowHeights(inches(1));
typedTable.setRowHeights([0, inches(1.5)]);
~~~

Keep `heightSnapshot` in the final `void` expression.

- [ ] **Step 5: Run focused conformance**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-row-heights.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
~~~

---

### Task 4: Document the editing slice accurately

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: final names, zero semantics, malformed behavior, conditional transform invariant, and verified PptxGenJS differences.
- Produces: public docs while broader table gaps remain explicitly unsupported.

- [ ] **Step 1: Update API example and behavior**

~~~ts
console.log(table.rowHeights?.map(emuToInches));
table.setRowHeights([
  inches(0.75),
  inches(1.25),
  inches(0.5),
]);
table.setRowHeights([0, inches(1), 0]);
~~~

Document detached snapshots, scalar broadcast, exact length, descriptor safety, raw-negative rejection, zero automatic rows, conditional safe sum, all-positive synchronization, zero-containing transform preservation, semantic no-op, getter `undefined`, and setter `ModelParseError`. State that `setTransform({ height })` changes only transform.

- [ ] **Step 2: Update compatibility baseline**

Move existing-table row heights from unsupported to supported. State that PptxGenJS has creation-only `rowH`, while native lossless editing repairs positive `cy` mismatch and safely preserves `cy` for automatic rows. Preserve single-array broadcast, short/falsy fallback, long-array truncation, coercion, and mismatch differences.

- [ ] **Step 3: Update README and changelog**

Add one Unreleased bullet and getter/setter examples. Remove only the existing-row-height gap; keep rich/multi-paragraph cell text, merge editing, row insertion/deletion, table/cell creation styles, auto-page, repeated headers, hyperlinks, and content measurement in unsupported lists.

- [ ] **Step 4: Run docs checks**

~~~sh
git diff --check
rg -n "rowHeights|setRowHeights|automatic row|row height|已有表格|部分支持" \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
~~~

---

### Task 5: Full gates, real-deck QA, review, commit, and push

**Files:**
- Review every Task 1-4 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: source/tests/docs, aggregate package, repository CLI, PptxGenJS 4.0.1, LibreOffice, Poppler, and presentations overflow helper.
- Produces: reviewed `feat: edit table row heights` commit synchronized to `origin/main`.

- [ ] **Step 1: Run full functional and performance gates**

~~~sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts
~~~

Record exact passed/skipped counts and inspect every failure, changed snapshot, or timing result.

- [ ] **Step 2: Build and smoke the actual npm tarball**

~~~sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js \
  --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_row_edit_package_dir=$(mktemp -d \
  /tmp/pptx-table-row-edit-package.XXXXXX)
npm pack --ignore-scripts \
  --pack-destination "$table_row_edit_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_row_edit_package_dir/jiayunxie-pptx-0.1.0.tgz"
~~~

Expected JSON includes `tableCreation`, `tableColumnWidths`, `tableColumnWidthEditing`, `tableRowHeights`, `tableRowHeightEditing`, declarations, browser success, and CLI `0.1.0`.

- [ ] **Step 3: Verify repository inspection runtime**

~~~sh
node packages/pptx/dist/cli.js --json doctor
~~~

Require successful JSON. Use this repository CLI path for validation and diff because no global `pptx-inspect` binary is assumed.

- [ ] **Step 4: Generate source, edited, reopened, and baseline decks**

Create `/tmp/pptx-table-row-height-editing-qa.mjs` with `apply_patch` and this complete program:

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
const output = '/tmp/pptx-table-row-height-editing';
await mkdir(output, { recursive: true });

const rows = [
  ['Region', 'Revenue', 'Growth'],
  ['East', '$1.2M', 'Vertical'],
  ['West', '$980K', ''],
];
const sourceHeights = [inches(0.6), inches(0.9), inches(0.9)];
const targetHeights = [inches(0.5), inches(1.25), inches(0.65)];

const native = PptxDocument.create({ slideSize: 'wide' });
const slide = native.addSlide();
const table = slide.addTable(rows, {
  name: 'Explicit row height editing QA',
  x: inches(1),
  y: inches(0.6),
  columnWidths: [inches(2), inches(4), inches(2)],
  rowHeights: sourceHeights,
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

const automatic = slide.addTable(rows, {
  name: 'Mixed automatic row height editing QA',
  x: inches(1),
  y: inches(4),
  columnWidths: [inches(2), inches(4), inches(2)],
  height: inches(2.4),
});
automatic.setCellFill(0, 0, {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent2' },
});
automatic.setCellTextDirection(1, 2, 'vert270');
automatic.setCellVerticalAlignment(2, 1, 'bottom');
await writeFile(output + '/source.pptx', await native.write());

table.setRowHeights(targetHeights);
assert.deepEqual(table.rowHeights, targetHeights);
assert.equal(table.transform.height, inches(2.4));

const automaticTransformHeight = automatic.transform.height;
automatic.setRowHeights([0, inches(0.8), 0]);
assert.deepEqual(automatic.rowHeights, [0, inches(0.8), 0]);
assert.equal(automatic.transform.height, automaticTransformHeight);

const editedBytes = await native.write();
await writeFile(output + '/native-edited.pptx', editedBytes);

const reopened = await PptxDocument.open(editedBytes);
const reopenedExplicit = reopened.slides[0].shapes.find(
  (shape) => shape.name === 'Explicit row height editing QA',
);
const reopenedAutomatic = reopened.slides[0].shapes.find(
  (shape) => shape.name === 'Mixed automatic row height editing QA',
);
assert.ok(reopenedExplicit instanceof TableModel);
assert.ok(reopenedAutomatic instanceof TableModel);
assert.deepEqual(reopenedExplicit.rowHeights, targetHeights);
assert.equal(reopenedExplicit.transform.height, inches(2.4));
assert.deepEqual(reopenedAutomatic.rowHeights, [0, inches(0.8), 0]);
assert.equal(reopenedAutomatic.transform.height, automaticTransformHeight);
assert.equal(reopenedExplicit.rows[1].cells[2].textDirection, 'vert270');
assert.equal(reopenedExplicit.rows[2].cells[2].text, '');
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
      textDirection: rowIndex === 1 && columnIndex === 2
        ? 'vert270'
        : 'horz',
    },
  }))),
  {
    x: 1,
    y: 0.6,
    w: 8,
    h: 2.4,
    colW: [2, 4, 2],
    rowH: [0.5, 1.25, 0.65],
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  explicitHeights: reopenedExplicit.rowHeights,
  explicitTransformHeight: reopenedExplicit.transform.height,
  automaticHeights: reopenedAutomatic.rowHeights,
  automaticTransformHeight: reopenedAutomatic.transform.height,
  emptyCell: reopenedExplicit.rows[2].cells[2].text === '',
  verticalCell: reopenedExplicit.rows[1].cells[2].textDirection,
}));
~~~

Run `node /tmp/pptx-table-row-height-editing-qa.mjs`. Expected explicit heights are `[457200,1143000,594360]`, explicit transform height is `2194560`, automatic heights are `[0,731520,0]`, automatic transform height is `2194560`, `emptyCell` is true, and `verticalCell` is `vert270`.

- [ ] **Step 5: Validate packages and mutation isolation**

~~~sh
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-row-height-editing/source.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-row-height-editing/native-edited.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-row-height-editing/native-reopened.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate \
  /tmp/pptx-table-row-height-editing/pptxgenjs-baseline.pptx \
  --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-row-height-editing/source.pptx \
  /tmp/pptx-table-row-height-editing/native-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-row-height-editing/native-edited.pptx \
  /tmp/pptx-table-row-height-editing/native-reopened.pptx
~~~

Require zero errors/warnings. Source-to-edited may change only `/ppt/slides/slide1.xml`; edited-to-reopened decompressed parts must be unchanged.

- [ ] **Step 6: Render and inspect all output**

~~~sh
mkdir -p /tmp/pptx-table-row-height-editing/rendered-native
mkdir -p /tmp/pptx-table-row-height-editing/rendered-baseline
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-row-height-editing/rendered-native \
  /tmp/pptx-table-row-height-editing/native-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice \
  --headless --convert-to pdf \
  --outdir /tmp/pptx-table-row-height-editing/rendered-baseline \
  /tmp/pptx-table-row-height-editing/pptxgenjs-baseline.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-row-height-editing/rendered-native/native-edited.pdf \
  /tmp/pptx-table-row-height-editing/rendered-native/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm \
  -png -r 180 \
  /tmp/pptx-table-row-height-editing/rendered-baseline/pptxgenjs-baseline.pdf \
  /tmp/pptx-table-row-height-editing/rendered-baseline/slide
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-row-height-editing/native-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-row-height-editing/pptxgenjs-baseline.pptx
~~~

Open every PNG at original detail. Require visible `0.5:1.25:0.65` explicit row proportions, readable headers, stable mixed automatic rows, vertical cell text, bottom alignment, empty final cell, and no repair, clipping, unintended overlap, blur, missing cell, or off-slide content.

- [ ] **Step 7: Review the complete diff**

~~~sh
git diff --check
git status --short
git diff --stat
git diff -- \
  packages/model/src/table-row-heights.internal.ts \
  packages/model/src/table-row-heights.internal.test.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
~~~

Map every design requirement to a test or QA artifact. Confirm `.pnpm-store/` is absent from staged changes and no unrelated source or generated file is included.

- [ ] **Step 8: Commit only the reviewed implementation**

~~~sh
git add \
  packages/model/src/table-row-heights.internal.ts \
  packages/model/src/table-row-heights.internal.test.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  scripts/smoke-npm-package.mjs \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "feat: edit table row heights"
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

Expected: `0 0` and only `?? .pnpm-store/`.

---

After synchronization, select the next narrow table capability gap, then repeat the separate design, plan, implementation, validation, review, commit, and push cycle without asking for routine decisions.
