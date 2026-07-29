# Table Cell Horizontal Alignment Editing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user requires autonomous inline execution without routine questions or subagents.

**Goal:** Add strict direct horizontal-alignment snapshots and lossless physical-cell editing for existing single-paragraph PPTX table cells.

**Architecture:** Add a dedicated internal codec that owns only the selected cell's unique direct `txBody/p/pPr@algn`, reusing the existing public `TextAlignment` union and `normalizeTextAlignment()`. `TableModel.rows` exposes a detached optional `horizontalAlignment`, while `setCellHorizontalAlignment()` performs canonical source-span add/replace/clear inside the existing OPC transaction boundary.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, OPC transactions, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarballs, repository JSON CLI, LibreOffice headless, Poppler, and the presentation overflow checker.

## Global Constraints

- Public snapshot is `TableCell.horizontalAlignment?: TextAlignment`.
- Public editor is `TableModel.setCellHorizontalAlignment(rowIndex, columnIndex, value)` with `TextAlignment | undefined`.
- Exact public values are `left`, `center`, `right`, and `justify`; canonical wire values are `l`, `ctr`, `r`, and `just`.
- Storage is the selected physical cell's unique direct `a:txBody/a:p/a:pPr@algn`, never `tcPr`, `bodyPr`, table metadata, style, or effective inheritance.
- Snapshot/editor support only cells with exactly one direct text body and one direct paragraph; rich/multi-paragraph alignment remains a separate future value-model slice.
- Getter returns `undefined` for direct absence, malformed/ambiguous structure, namespaced-only attributes, unsupported tokens, or non-single-paragraph cells; getter never mutates or throws for these states.
- Setter requires exactly one direct text body/paragraph, at most one direct `pPr`, and at most one unqualified `algn`; unsafe structure throws `ModelParseError` before mutation.
- Setter creates a same-prefix `pPr` only for a valid non-undefined value, preserves an empty `pPr` on clear, and never removes unrelated paragraph properties.
- Same canonical value and absent clear are exact no-ops; one unknown unqualified token may be replaced or cleared.
- Reuse `normalizeTextAlignment()`; do not add aliases, coercion, case folding, trimming, or PptxGenJS invalid fallback.
- Preserve all paragraph runs/fields/breaks/end properties, cell properties, neighbor cells, geometry, relationships, package parts, and stable model identity.
- Every successful small item is tested, reviewed, committed, pushed over SSH port 443, fetched, and verified with `origin/main...HEAD = 0 0`; QA-only gates do not create empty commits.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Add the strict direct-paragraph codec and public model surface

**Files:**
- Create: `packages/model/src/table-cell-horizontal-alignment.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `TextAlignment`, `normalizeTextAlignment(value, context)`, `LosslessXmlDocument`, `ModelParseError`, `TableModel.resolve()`, and OPC transactions.
- Produces: `readTableCellHorizontalAlignment()`, `replaceTableCellHorizontalAlignment()`, `TableCell.horizontalAlignment`, and `TableModel.setCellHorizontalAlignment()`.

- [ ] **Step 1: Add failing strict-read and lossless-edit tests**

Add two tests beside the table-cell vertical-alignment model tests. The strict-read fixture must use one physical row with these direct paragraph cases:

```ts
const paragraph = (properties: string, text: string): string =>
  `<a:p>${properties}<a:r><a:t>${text}</a:t></a:r></a:p>`;
const cell = (body: string, properties = '<a:tcPr/>'): string =>
  `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${body}</a:txBody>${properties}</a:tc>`;
const cells = [
  cell(paragraph('<a:pPr algn="l"/>', 'Left')),
  cell(paragraph('<a:pPr algn="ctr"/>', 'Center')),
  cell(paragraph('<a:pPr algn="r"/>', 'Right')),
  cell(paragraph('<a:pPr algn="just"/>', 'Justify')),
  cell(paragraph('', 'Missing pPr')),
  cell(paragraph('<a:pPr/>', 'Missing algn')),
  cell(paragraph('<a:pPr algn=""/>', 'Empty')),
  cell(paragraph('<a:pPr algn="L"/>', 'Case')),
  cell(paragraph('<a:pPr algn=" ctr "/>', 'Whitespace')),
  cell(paragraph('<a:pPr algn="left"/>', 'Long form')),
  cell(paragraph('<a:pPr algn="dist"/>', 'Distributed')),
  cell(paragraph('<a:pPr algn="thaiDist"/>', 'Thai distributed')),
  cell(paragraph('<a:pPr algn="justLow"/>', 'Low justify')),
  cell(paragraph('<a:pPr algn="unknown"/>', 'Unknown')),
  cell(paragraph('<a:pPr xmlns:x="urn:test" x:algn="ctr"/>', 'Namespaced')),
  cell(paragraph('<a:pPr algn="l" algn="r"/>', 'Repeated attribute')),
  cell(paragraph('<a:pPr algn="l"/><a:pPr keep="SECOND"/>', 'Repeated pPr')),
  cell(`${paragraph('<a:pPr algn="l"/>', 'First')}${paragraph('<a:pPr algn="r"/>', 'Second')}`),
  '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/></a:txBody><a:tcPr/></a:tc>',
  '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First body</a:t></a:r></a:p></a:txBody><a:txBody><a:bodyPr/><a:p><a:r><a:t>Second body</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>',
  cell(paragraph('', 'Impostors'), '<a:tcPr algn="ctr"><x:pPr xmlns:x="urn:test" algn="r"/></a:tcPr>'),
].join('');
```

After patching the existing table row with `cells`, assert:

```ts
const snapshot = table.rows;
expect(snapshot[0]!.cells.slice(0, 4).map(({ horizontalAlignment }) =>
  horizontalAlignment)).toEqual(['left', 'center', 'right', 'justify']);
expect(snapshot[0]!.cells.slice(4).every(({ horizontalAlignment }) =>
  horizontalAlignment === undefined)).toBe(true);
expect(pkg.mutations).toEqual(journal);
(snapshot[0]!.cells[0] as { horizontalAlignment?: string }).horizontalAlignment = 'right';
expect(table.rows[0]!.cells[0]!.horizontalAlignment).toBe('left');
expect(pkg.mutations).toEqual(journal);
```

The edit fixture must contain: same `ctr` with single quotes and opaque properties, missing `pPr`, self-closing `pPr`, expanded `pPr` with runs/fields/breaks/extensions, clear `just`, unknown `dist`, namespaced-only `x:algn`, a merged placeholder, repeated attribute, repeated `pPr`, multiple paragraphs, zero paragraphs, and repeated text bodies. Assert:

```ts
table.setCellHorizontalAlignment(0, 0, 'center'); // exact no-op
table.setCellHorizontalAlignment(0, 1, 'left');
table.setCellHorizontalAlignment(0, 2, 'right');
table.setCellHorizontalAlignment(0, 3, 'justify');
table.setCellHorizontalAlignment(0, 4, undefined);
table.setCellHorizontalAlignment(0, 5, 'center');
table.setCellHorizontalAlignment(0, 6, undefined);
table.setCellHorizontalAlignment(0, 7, 'left');
```

Require same-prefix `pPr` insertion, exact preservation of unrelated `pPr` attributes/children and all run/field/break XML, preserved `x:algn`, unchanged neighbor cell, four public snapshots, clear to `undefined`, and same-value journal identity. Invalid coordinates throw `RangeError`; repeated/ambiguous structure columns throw `ModelParseError`; invalid runtime values throw `TypeError`; all failures preserve exact bytes and journal. An outer package transaction that edits two cells and throws must restore bytes, snapshots, identity, and journal.

- [ ] **Step 2: Run the focused tests and confirm red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t "table-cell horizontal alignment"
```

Expected: TypeScript/runtime failure because `horizontalAlignment` and `setCellHorizontalAlignment()` do not exist.

- [ ] **Step 3: Implement the dedicated codec**

Create `table-cell-horizontal-alignment.internal.ts` with this complete structure:

```ts
import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type { TextAlignment } from './text.js';

const FROM_OOXML = new Map<string, TextAlignment>([
  ['l', 'left'],
  ['ctr', 'center'],
  ['r', 'right'],
  ['just', 'justify'],
]);
const TO_OOXML: Readonly<Record<TextAlignment, string>> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
};

export function readTableCellHorizontalAlignment(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
): TextAlignment | undefined {
  const textBodies = directChildren(cell, 'txBody');
  if (textBodies.length !== 1) return undefined;
  const paragraphs = directChildren(textBodies[0]!, 'p');
  if (paragraphs.length !== 1) return undefined;
  const properties = directChildren(paragraphs[0]!, 'pPr');
  if (properties.length !== 1) return undefined;
  const attributes = properties[0]!.attributes.filter(({ name }) => name === 'algn');
  if (attributes.length !== 1) return undefined;
  return FROM_OOXML.get(attributes[0]!.value);
}

export function replaceTableCellHorizontalAlignment(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  value: TextAlignment | undefined,
  partUri: string,
): boolean {
  const textBodies = directChildren(cell, 'txBody');
  if (textBodies.length !== 1) {
    throw new ModelParseError(
      'Table cell must contain exactly one direct text body',
      partUri,
    );
  }
  const paragraphs = directChildren(textBodies[0]!, 'p');
  if (paragraphs.length !== 1) {
    throw new ModelParseError(
      'Table cell must contain exactly one direct text paragraph',
      partUri,
    );
  }
  const paragraphElement = paragraphs[0]!;
  const paragraphXml = LosslessXmlDocument.parse(xml.original(paragraphElement));
  const root = paragraphXml.roots[0];
  if (!root || root.localName !== 'p' || root.selfClosing) {
    throw new ModelParseError('Invalid table cell paragraph template', partUri);
  }
  const properties = directChildren(root, 'pPr');
  if (properties.length > 1) {
    throw new ModelParseError(
      'Table cell paragraph contains repeated direct properties elements',
      partUri,
    );
  }
  const token = value === undefined ? undefined : TO_OOXML[value];
  const propertiesElement = properties[0];
  if (!propertiesElement) {
    if (token === undefined) return false;
    const firstElement = root.children.find(
      (child): child is XmlElement => child.type === 'element',
    );
    const insertionPoint = firstElement?.start ?? root.endTagStart;
    const separator = root.name.includes(':')
      ? `${root.name.slice(0, root.name.indexOf(':'))}:`
      : '';
    paragraphXml.replace(
      insertionPoint,
      insertionPoint,
      `<${separator}pPr algn="${escapeXmlAttribute(token)}"/>`,
    );
    xml.replaceElement(paragraphElement, paragraphXml.serialize());
    return true;
  }
  const attributes = propertiesElement.attributes.filter(({ name }) => name === 'algn');
  if (attributes.length > 1) {
    throw new ModelParseError(
      'Table cell paragraph contains repeated direct horizontal alignment attributes',
      partUri,
    );
  }
  const attribute = attributes[0];
  if (attribute?.value === token || (!attribute && token === undefined)) return false;
  if (token !== undefined) {
    if (attribute) paragraphXml.replaceAttribute(attribute, token);
    else {
      const insertionPoint = propertiesElement.selfClosing
        ? paragraphXml.source.lastIndexOf('/', propertiesElement.startTagEnd - 1)
        : propertiesElement.startTagEnd - 1;
      paragraphXml.replace(
        insertionPoint,
        insertionPoint,
        ` algn="${escapeXmlAttribute(token)}"`,
      );
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > propertiesElement.start && /[\t ]/.test(paragraphXml.source[start - 1] ?? '')) {
      start -= 1;
    }
    paragraphXml.replace(start, attribute.end, '');
  }
  xml.replaceElement(paragraphElement, paragraphXml.serialize());
  return true;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
```

Do not export the codec from the public model barrel.

- [ ] **Step 4: Wire the public snapshot and setter**

In `shapes.ts`, import the codec and `normalizeTextAlignment`, add the type-only `TextAlignment` import, and extend `TableCell`:

```ts
readonly horizontalAlignment?: TextAlignment;
```

In `TableModel.rows`, read and conditionally spread the snapshot:

```ts
const horizontalAlignment = readTableCellHorizontalAlignment(xml, cell);
// ...
...(horizontalAlignment !== undefined ? { horizontalAlignment } : {}),
```

Add the physical-cell editor beside `setCellVerticalAlignment()`:

```ts
setCellHorizontalAlignment(
  rowIndex: number,
  columnIndex: number,
  value: TextAlignment | undefined,
): void {
  const alignment = value === undefined
    ? undefined
    : normalizeTextAlignment(value, 'Table cell horizontal alignment');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    if (replaceTableCellHorizontalAlignment(xml, cell, alignment, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

- [ ] **Step 5: Run model regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/model/src/table-create.internal.test.ts
git diff --check
```

Expected: all model and creation tests pass; created table/cell `align` now appears immediately in `horizontalAlignment` snapshots without changing creation XML.

- [ ] **Step 6: Review, commit, push, fetch, and prove synchronization**

Review single-paragraph gating, direct ownership, namespace handling, unknown replacement, exact no-ops, prefix insertion, rollback, and unrelated XML preservation. Stage only the three files and commit:

```text
feat: edit table cell horizontal alignment
```

Then run:

```sh
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git HEAD:main
git fetch --force ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `?? .pnpm-store/` outside tracked state.

---

### Task 2: Prove the public SDK lifecycle and strict failures

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: public `TableCell.horizontalAlignment` and `TableModel.setCellHorizontalAlignment()` from Task 1.
- Produces: end-to-end Node SDK evidence for creation/import, edit, duplicate isolation, rollback, reopen, invalid values, and invalid physical coordinates.

- [ ] **Step 1: Add a dedicated strict single-paragraph fixture**

Add `tableHorizontalAlignmentFixture()` after `tableTextDirectionFixture()`. Open the shared fixture and insert direct paragraph properties into the five cells of row 2 without altering their `tcPr` or body properties:

```ts
async function tableHorizontalAlignmentFixture(): Promise<Uint8Array> {
  const document = await PptxDocument.open(await tableTextDirectionFixture());
  const slide = document.slides[0]!;
  const part = document.opcPackage.requirePart(slide.partUri);
  let source = new TextDecoder().decode(part.bytes);
  for (const [text, properties] of [
    ['Explicit none', '<a:pPr algn="l" keep="LEFT"/>'],
    ['Calculated shrink', '<a:pPr algn="ctr" keep="CENTER"><a:buNone/></a:pPr>'],
    ['Resize', '<a:pPr algn="r" keep="RIGHT"/>'],
    ['Absent fit', '<a:pPr algn="just" keep="JUSTIFY"/>'],
    ['Merged fit', '<a:pPr keep="ABSENT"/>'],
  ] as const) {
    source = source.replace(
      `<a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t>`,
      `<a:p>${properties}<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t>`,
    );
  }
  document.opcPackage.setPart(slide.partUri, source, part.contentType);
  return document.write();
}
```

Assert each replacement occurred exactly once before returning.

- [ ] **Step 2: Cover edit, duplicate, rollback, and reopen**

Add `it('edits table-cell horizontal alignments through duplicate, rollback, and reopen lifecycles', ...)` using the new fixture. Initial row-2 snapshots must be:

```ts
['left', 'center', 'right', 'justify', undefined]
```

Duplicate the slide, then edit the source:

```ts
table.setCellHorizontalAlignment(2, 0, 'center');
table.setCellHorizontalAlignment(2, 1, 'right');
table.setCellHorizontalAlignment(2, 2, undefined);
table.setCellHorizontalAlignment(2, 3, 'left');
table.setCellHorizontalAlignment(2, 4, 'justify');
```

Also call text, fit, direction, vertical-alignment, margin, border/fill, transform, column-width, and row-height editors. Require source snapshots `['center','right',undefined,'left','justify']`, unchanged duplicate snapshots, correct neighbor properties, preserved `keep` markers, direct `pPr@algn` only, stable source identity, outer transaction rollback, and identical state after write/reopen.

- [ ] **Step 3: Cover invalid runtime values and coordinates before mutation**

Add a separate test using this exact invalid set:

```ts
const invalidValues = [
  null,
  false,
  true,
  0,
  1,
  '',
  'Left',
  ' center ',
  'l',
  'ctr',
  'r',
  'just',
  'dist',
  'thaiDist',
  'justLow',
  {},
  [],
  Symbol('horizontal alignment'),
];
```

Use this exact invalid-coordinate vector:

```ts
const invalidCoordinates = [
  [-1, 0],
  [0, -1],
  [0.5, 0],
  [0, 0.5],
  [Number.NaN, 0],
  [0, Number.NaN],
  [Number.POSITIVE_INFINITY, 0],
  [0, Number.NEGATIVE_INFINITY],
  [3, 0],
  [2, 5],
];
```

Every call must throw `TypeError` or `RangeError` before mutation. Require exact slide bytes, journal, slide/table identity, all horizontal snapshots, all other cell snapshots, and text to remain unchanged.

- [ ] **Step 4: Run SDK/model regressions and typecheck**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  -t "table-cell horizontal alignment|table creation"
git diff --check
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review public-only imports, created/existing table behavior, physical coordinates, duplicate isolation, package-part isolation, rollback, and reopen. Stage only `packages/sdk/src/index.test.ts` and commit:

```text
test: cover sdk table cell horizontal alignment editing
```

Push/fetch with the Global Constraints commands and require `0 0`.

---

### Task 3: Prove PptxGenJS 4.0.1 materialized-state interoperability

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS output import plus Task 1 snapshots/editor.
- Produces: evidence that supported PptxGenJS table/cell values import to exact public snapshots and can be losslessly edited by the native model.

- [ ] **Step 1: Extend both existing horizontal-alignment parity fixtures**

In `matches native table-cell horizontal alignment to PptxGenJS final state`, add:

```ts
const expectedAlignments = [undefined, 'left', 'center', 'right', 'justify'];
expect(nativeTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
  horizontalAlignment)).toEqual(expectedAlignments);
expect(importedTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
  horizontalAlignment)).toEqual(expectedAlignments);
```

After reopen, assert both snapshot arrays again. The invalid PptxGenJS `dist` cell must import with `horizontalAlignment === undefined`.

In `matches native table horizontal alignment to PptxGenJS final state`, assert exact public values corresponding to `['ctr','ctr','ctr','l','r','just']`:

```ts
['center', 'center', 'center', 'left', 'right', 'justify']
```

- [ ] **Step 2: Edit imported PptxGenJS direct state**

On `importedTable`, edit one inherited cell to right, clear the cell override at index 3, and change justify to center:

```ts
importedTable.setCellHorizontalAlignment(0, 0, 'right');
importedTable.setCellHorizontalAlignment(0, 3, undefined);
importedTable.setCellHorizontalAlignment(0, 5, 'center');
```

Require public snapshots and direct tokens to become:

```ts
['right', 'center', 'center', undefined, 'right', 'center']
['r', 'ctr', 'ctr', undefined, 'r', 'ctr']
```

Require text, margins, valign, borders, fill, geometry, and every non-slide package part unchanged; write/reopen must preserve the edited final state.

- [ ] **Step 3: Run adapter/model/type regressions**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/model.test.ts \
  -t "horizontal alignment"
git diff --check
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Review that tests use only public PptxGenJS output, distinguish omitted/malformed direct state, and do not claim PptxGenJS has an existing-deck editor. Stage only the adapter test and commit:

```text
test: compare table cell horizontal alignment editing with pptxgenjs
```

Push/fetch and require `0 0`.

---

### Task 4: Prove the packed Node, browser, declaration, and CLI surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: the actual packed `@jiayunxie/pptx` Node/browser bundles and declarations.
- Produces: `tableCellHorizontalAlignmentEditing: true` in the smoke JSON and compile-time proof of the new public field/method.

- [ ] **Step 1: Extend Node runtime smoke**

Read the created table's existing cell/table alignment snapshots, then exercise all values and clear:

```js
const initialHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'left');
const leftHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'center');
const centerHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'right');
const rightHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, 'justify');
const justifyHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, undefined);
const clearedHorizontalAlignment = table?.rows[0]?.cells[0]?.horizontalAlignment;
```

Add a check that the neighbor is unchanged and emit:

```js
tableCellHorizontalAlignmentEditing:
  initialHorizontalAlignment === 'center' &&
  leftHorizontalAlignment === 'left' &&
  centerHorizontalAlignment === 'center' &&
  rightHorizontalAlignment === 'right' &&
  justifyHorizontalAlignment === 'justify' &&
  clearedHorizontalAlignment === undefined,
```

Ensure creation→edit→write→reopen also reads exact final state from the live and reopened table.

- [ ] **Step 2: Extend browser runtime smoke**

On the existing strict single-paragraph `Browser table` fixture, add:

```js
if (table.rows[0].cells[0].horizontalAlignment !== undefined ||
    table.rows[0].cells[1].horizontalAlignment !== undefined) {
  throw new Error('Browser table-cell horizontal alignment initial read failed');
}
table.setCellHorizontalAlignment(0, 0, 'left');
if (table.rows[0].cells[0].horizontalAlignment !== 'left') {
  throw new Error('Browser table-cell left alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'center');
if (table.rows[0].cells[0].horizontalAlignment !== 'center') {
  throw new Error('Browser table-cell center alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'right');
if (table.rows[0].cells[0].horizontalAlignment !== 'right') {
  throw new Error('Browser table-cell right alignment failed');
}
table.setCellHorizontalAlignment(0, 0, 'justify');
if (table.rows[0].cells[0].horizontalAlignment !== 'justify') {
  throw new Error('Browser table-cell justify alignment failed');
}
table.setCellHorizontalAlignment(0, 0, undefined);
if (table.rows[0].cells[0].horizontalAlignment !== undefined ||
    table.rows[0].cells[1].horizontalAlignment !== undefined ||
    table.rows[0].cells[0].text !== 'Browser target' ||
    table.rows[0].cells[1].text !== 'Browser neighbor') {
  throw new Error('Browser table-cell horizontal alignment clear failed');
}
```

The later browser write/reopen must assert both cells still read `undefined` and all direction/fit/vertical/margin/border/fill checks remain green.

- [ ] **Step 3: Extend declaration compilation**

Add:

```ts
const snapshotHorizontalAlignment: TextAlignment | undefined =
  table?.rows[0]?.cells[0]?.horizontalAlignment;
table?.setCellHorizontalAlignment(0, 0, cellHorizontalAlignment);
table?.setCellHorizontalAlignment(0, 0, undefined);
```

Include `snapshotHorizontalAlignment` in the final `void [...]` expression so strict compilation proves the declaration is usable.

- [ ] **Step 4: Build and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_align_edit_package_dir=$(mktemp -d /tmp/pptx-table-cell-align-edit-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_align_edit_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_cell_align_edit_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
```

Require every API flag true, including `tableCellHorizontalAlignmentEditing`, declarations true, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review installed-tarball-only imports, Node/browser parity, strict types, reopen behavior, and unchanged CLI. Stage only the smoke script and commit:

```text
test: smoke packed table cell horizontal alignment editing
```

Push/fetch and require `0 0`.

---

### Task 5: Document the supported boundary

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: final public behavior and compatibility evidence.
- Produces: accurate documentation for direct strict single-paragraph snapshot/editing and remaining rich/multi-paragraph limits.

- [ ] **Step 1: Update release and API documentation**

Document this exact contract in the repository's existing prose style:

```text
TableCell.horizontalAlignment reads only a strict single paragraph's direct
pPr@algn as left, center, right, or justify. TableModel.setCellHorizontalAlignment()
uses physical zero-based row/cell coordinates, writes l/ctr/r/just, and clears
only the direct unqualified token with undefined. It never resolves effective
alignment or retained table defaults.
```

State that missing/malformed/multi-paragraph cells read `undefined`, unsafe structure rejects editing without mutation, same-value edits are exact no-ops, and all other paragraph/cell state is preserved.

- [ ] **Step 2: Update PptxGenJS compatibility wording**

Mark PptxGenJS 4.0.1 table/cell valid materialized direct state as import-compatible. State that native editing is a lossless extension because PptxGenJS has no existing-deck editor. Preserve the invalid fallback difference and do not claim rich/multi-paragraph alignment support.

- [ ] **Step 3: Run documentation scans and full typecheck**

```sh
rg -n 'existing-cell horizontal-alignment.*(unsupported|pending)|已有.*horizontal-alignment.*尚未支持|TableCell.*horizontalAlignment' \
  CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
node node_modules/typescript/bin/tsc -b --pretty false
git diff --check
```

Review each result in context: strict single-paragraph snapshot/editing is supported, while rich/multi-paragraph alignment remains pending.

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

Review public names, four values, physical coordinates, direct ownership, single-paragraph boundary, no-op/clear semantics, PptxGenJS extension wording, and remaining unsupported scope. Stage only the four docs and commit:

```text
docs: document table cell horizontal alignment editing
```

Push/fetch and require `0 0`.

---

### Task 6: Run full gates and real-deck QA

**Files:**
- Review every Task 1–5 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: implementation, tests, docs, actual tarball, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces: a verified pushed slice; defects receive their own focused fix/review/commit/push cycle, while green QA creates no empty commit.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-align-edit-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-cell-align-edit-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
table_cell_align_edit_qa_package_dir=$(mktemp -d /tmp/pptx-table-cell-align-edit-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$table_cell_align_edit_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$table_cell_align_edit_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
```

Require offline doctor success and CLI `0.1.0`.

- [ ] **Step 3: Generate native and PptxGenJS real decks**

Create `/tmp/pptx-table-cell-align-edit-qa.mjs` with `apply_patch` using this fixed fixture:

```js
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
const output = '/tmp/pptx-table-cell-align-edit';
await mkdir(output, { recursive: true });

const expectedSource = ['left', 'center', 'right', 'justify', undefined];
const expectedSourceTokens = ['l', 'ctr', 'r', 'just', undefined];
const expectedEdited = ['left', 'right', undefined, 'justify', 'left'];
const expectedEditedTokens = ['l', 'r', undefined, 'just', 'l'];
const columnWidths = [
  inches(1.4),
  inches(2.2),
  inches(1.8),
  inches(1.8),
  inches(1.8),
];
const rowHeights = [inches(0.8), inches(1.4)];

function slideXml(document) {
  return new TextDecoder().decode(
    document.opcPackage.requirePart(document.slides[0].partUri).bytes,
  );
}

function tokens(document) {
  return [...slideXml(document).matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
    .map((match) => match[0].match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
}

function firstTable(document, slideIndex = 0) {
  const table = document.slides[slideIndex].shapes.find(
    (shape) => shape instanceof TableModel,
  );
  assert.ok(table instanceof TableModel);
  return table;
}

function alignments(table) {
  return table.rows[0].cells.map(({ horizontalAlignment }) => horizontalAlignment);
}

const native = PptxDocument.create({ slideSize: 'wide' });
const nativeTable = native.addSlide().addTable([
  [
    { text: 'Left direct', options: {
      align: 'left',
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'C00000' },
        width: 2,
      },
    } },
    { text: 'Center direct', options: {
      align: 'center',
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    } },
    { text: 'Right direct', options: { align: 'right' } },
    { text: 'Justify this sentence', options: {
      align: 'justify',
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
    } },
    { text: 'Absent direct', options: { margin: { left: 18 } } },
  ],
  ['Second row A', 'Second row B', 'Second row C', 'Second row D', 'Second row E'],
], {
  name: 'Horizontal alignment editing QA',
  x: inches(0.6),
  y: inches(0.7),
  columnWidths,
  rowHeights,
  valign: 'middle',
});
assert.deepEqual(alignments(nativeTable), expectedSource);
assert.deepEqual(tokens(native).slice(0, 5), expectedSourceTokens);
assert.deepEqual(nativeTable.columnWidths, columnWidths);
assert.deepEqual(nativeTable.rowHeights, rowHeights);
const nativeDuplicate = native.duplicateSlide(0);
const nativeDuplicateTable = firstTable(native, 1);
assert.deepEqual(alignments(nativeDuplicateTable), expectedSource);
assert.equal(nativeDuplicate.shapes[0], nativeDuplicateTable);
const nativeSourceBytes = await native.write();
await writeFile(output + '/native-source.pptx', nativeSourceBytes);

const noOp = await PptxDocument.open(nativeSourceBytes);
const noOpSlide = noOp.slides[0];
const noOpBefore = noOp.opcPackage.requirePart(noOpSlide.partUri).bytes.slice();
firstTable(noOp).setCellHorizontalAlignment(0, 0, 'left');
assert.deepEqual(noOp.opcPackage.requirePart(noOpSlide.partUri).bytes, noOpBefore);
const noOpBytes = await noOp.write();
await writeFile(output + '/native-noop.pptx', noOpBytes);

nativeTable.setCellHorizontalAlignment(0, 1, 'right');
nativeTable.setCellHorizontalAlignment(0, 2, undefined);
nativeTable.setCellHorizontalAlignment(0, 4, 'left');
nativeTable.setCellText(0, 0, 'Edited left direct');
nativeTable.setCellMargins(0, 1, { bottom: 9 });
nativeTable.setCellVerticalAlignment(0, 3, 'top');
nativeTable.setCellFill(1, 4, {
  kind: 'solid',
  color: { kind: 'srgb', value: 'E2F0D9' },
});
assert.deepEqual(alignments(nativeTable), expectedEdited);
assert.deepEqual(tokens(native).slice(0, 5), expectedEditedTokens);
assert.equal(nativeTable.rows[0].cells[0].text, 'Edited left direct');
assert.deepEqual(alignments(nativeDuplicateTable), expectedSource);
const nativeEditedBytes = await native.write();
await writeFile(output + '/native-edited.pptx', nativeEditedBytes);

const reopened = await PptxDocument.open(nativeEditedBytes);
const reopenedTable = firstTable(reopened);
assert.deepEqual(alignments(reopenedTable), expectedEdited);
assert.deepEqual(tokens(reopened).slice(0, 5), expectedEditedTokens);
assert.deepEqual(reopenedTable.columnWidths, columnWidths);
assert.deepEqual(reopenedTable.rowHeights, rowHeights);
await writeFile(output + '/native-reopened.pptx', await reopened.write());

const generated = new PptxGenJS();
generated.layout = 'LAYOUT_WIDE';
generated.addSlide().addTable([[
  { text: 'Left direct', options: { align: 'left', border: { type: 'solid', color: 'C00000', pt: 2 } } },
  { text: 'Center direct', options: { align: 'center', fill: { color: 'D9EAF7' } } },
  { text: 'Right direct', options: { align: 'right' } },
  { text: 'Justify this sentence', options: { align: 'justify', fill: { color: 'FFF2CC' } } },
  { text: 'Absent direct', options: { margin: [3.6, 7.2, 3.6, 18] } },
]], {
  x: 0.6,
  y: 0.7,
  w: 9,
  h: 0.8,
  colW: [1.4, 2.2, 1.8, 1.8, 1.8],
  rowH: [0.8],
  valign: 'middle',
});
const baselineBytes = await generated.write({ outputType: 'uint8array', compression: true });
const baseline = await PptxDocument.open(baselineBytes);
assert.deepEqual(alignments(firstTable(baseline)), expectedSource);
assert.deepEqual(tokens(baseline), expectedSourceTokens);
await writeFile(output + '/pptxgenjs-baseline.pptx', baselineBytes);

const hand = PptxDocument.create({ slideSize: 'wide' });
hand.addSlide().addTable([['Unknown direct', 'Namespaced direct']], {
  x: inches(1),
  y: inches(1),
  columnWidths: [inches(3), inches(3)],
  rowHeights: [inches(1)],
});
const handPart = hand.opcPackage.requirePart(hand.slides[0].partUri);
let patchedCount = 0;
const patchedXml = slideXml(hand).replace(
  /<a:pPr indent="0" marL="0"><a:buNone\/><\/a:pPr>/g,
  () => {
    patchedCount += 1;
    return patchedCount === 1
      ? '<a:pPr algn="dist" keep="UNKNOWN"><a:buNone/><x:keep xmlns:x="urn:test">OPAQUE</x:keep></a:pPr>'
      : '<a:pPr xmlns:x="urn:test" x:algn="ctr" keep="NAMESPACED"><a:buNone/></a:pPr>';
  },
);
assert.equal(patchedCount, 2);
hand.opcPackage.setPart(handPart.uri, patchedXml, handPart.contentType);
const handTable = firstTable(hand);
assert.deepEqual(alignments(handTable), [undefined, undefined]);
await writeFile(output + '/hand-source.pptx', await hand.write());
handTable.setCellHorizontalAlignment(0, 0, 'center');
handTable.setCellHorizontalAlignment(0, 1, 'left');
handTable.setCellHorizontalAlignment(0, 1, undefined);
assert.deepEqual(alignments(handTable), ['center', undefined]);
assert.match(slideXml(hand), /<a:pPr algn="ctr" keep="UNKNOWN">/);
assert.match(slideXml(hand), /x:algn="ctr" keep="NAMESPACED"/);
assert.match(slideXml(hand), /<x:keep xmlns:x="urn:test">OPAQUE<\/x:keep>/);
await writeFile(output + '/hand-edited.pptx', await hand.write());

process.stdout.write(JSON.stringify({
  source: expectedSource,
  edited: expectedEdited,
  reopened: alignments(reopenedTable),
  hand: alignments(handTable),
}));
```

Run `node /tmp/pptx-table-cell-align-edit-qa.mjs`. Require every assertion and the final JSON to pass.

- [ ] **Step 4: Validate and diff packages**

Run:

```sh
for deck in /tmp/pptx-table-cell-align-edit/*.pptx; do
  node packages/pptx/dist/cli.js --json package validate \
    "$deck" --profile powerpoint-2010
done
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-align-edit/native-source.pptx \
  /tmp/pptx-table-cell-align-edit/native-noop.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-align-edit/native-source.pptx \
  /tmp/pptx-table-cell-align-edit/native-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-align-edit/native-edited.pptx \
  /tmp/pptx-table-cell-align-edit/native-reopened.pptx
```

Require zero errors/warnings. The first and third diffs must contain zero changed parts. The second must contain only `/ppt/slides/slide1.xml`; unchanged `/ppt/slides/slide2.xml` and slide relationships prove duplicate isolation.

- [ ] **Step 5: Render and inspect at original resolution**

Use isolated LibreOffice profiles to export native edited, PptxGenJS baseline, and hand-patched edited decks, then rasterize every page:

```sh
table_cell_align_edit_render_dir=$(mktemp -d /tmp/pptx-table-cell-align-edit-render.XXXXXX)
mkdir -p "$table_cell_align_edit_render_dir/native" \
  "$table_cell_align_edit_render_dir/baseline" \
  "$table_cell_align_edit_render_dir/hand"
table_cell_align_edit_native_profile=$(mktemp -d /tmp/pptx-table-cell-align-edit-native-profile.XXXXXX)
table_cell_align_edit_baseline_profile=$(mktemp -d /tmp/pptx-table-cell-align-edit-baseline-profile.XXXXXX)
table_cell_align_edit_hand_profile=$(mktemp -d /tmp/pptx-table-cell-align-edit-hand-profile.XXXXXX)
soffice -env:UserInstallation="file://$table_cell_align_edit_native_profile" \
  --headless --convert-to pdf \
  --outdir "$table_cell_align_edit_render_dir/native" \
  /tmp/pptx-table-cell-align-edit/native-edited.pptx
soffice -env:UserInstallation="file://$table_cell_align_edit_baseline_profile" \
  --headless --convert-to pdf \
  --outdir "$table_cell_align_edit_render_dir/baseline" \
  /tmp/pptx-table-cell-align-edit/pptxgenjs-baseline.pptx
soffice -env:UserInstallation="file://$table_cell_align_edit_hand_profile" \
  --headless --convert-to pdf \
  --outdir "$table_cell_align_edit_render_dir/hand" \
  /tmp/pptx-table-cell-align-edit/hand-edited.pptx
pdftoppm -png -r 180 \
  "$table_cell_align_edit_render_dir/native/native-edited.pdf" \
  "$table_cell_align_edit_render_dir/native/slide"
pdftoppm -png -r 180 \
  "$table_cell_align_edit_render_dir/baseline/pptxgenjs-baseline.pdf" \
  "$table_cell_align_edit_render_dir/baseline/slide"
pdftoppm -png -r 180 \
  "$table_cell_align_edit_render_dir/hand/hand-edited.pdf" \
  "$table_cell_align_edit_render_dir/hand/slide"
```

Inspect every PNG at original detail. Require visible left/center/right/justify and clear-to-default behavior, preserved border/fill/margin/valign, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing cell, or off-slide content.

Run both overflow checks and require no overflow:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-align-edit/native-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-align-edit/pptxgenjs-baseline.pptx
```

- [ ] **Step 6: Final static review and synchronization proof**

Review strict single-paragraph gating, direct ownership, namespace behavior, normalization reuse, source-span isolation, invalid-input no-mutation, declarations, adapter evidence, smoke flags, and docs boundaries.

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `?? .pnpm-store/`. If QA finds a defect, fix only responsible files, rerun focused/full gates, review, commit, push, fetch, and re-prove synchronization; otherwise create no empty commit.
