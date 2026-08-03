# Table-Cell Hyperlink Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict plain table-cell URL/slide hyperlink creation and readonly direct-state snapshots, with PptxGenJS 4.0.1 final-state conformance and actual-package/browser proof.

**Architecture:** Extend table-cell normalization with the existing `Hyperlink` codec, prepare every relationship input before mutation, allocate one independent relationship per linked cell inside the existing table-creation transaction, and pass a relationship-ID matrix to the one-run table renderer. Add a focused strict direct-cell reader that delegates click decoding to the existing text-run hyperlink reader and expose it through `TableCell.hyperlink`.

**Tech Stack:** TypeScript, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public API, tsup, npm pack, installed CLI, `pptx-inspect`, Playwright, and Google Chrome.

## Global Constraints

- Public creation input: `AddTableCellOptions.hyperlink?: Hyperlink`.
- Public snapshot: `TableCell.hyperlink?: Hyperlink`.
- No `AddTableOptions.hyperlink`, table graphic-frame hyperlink, cell hyperlink setter, or rich/multi-run cell support in this item.
- URL and slide branches remain mutually exclusive; URL is non-empty/XML-safe; slide is a current positive safe one-based ordinal; tooltip absence and explicit empty remain distinct.
- Every linked cell owns one independent relationship, including equal-target cells.
- All table cells and internal slide targets are validated before the first hyperlink relationship is added.
- Cell links are direct single-run `a:rPr/a:hlinkClick` values and use the existing automatic `u="sng"` rendering.
- Snapshot reading accepts only one safe direct `txBody/p/r/rPr` path and never guesses through multiple paragraphs/runs, fields, graphic-frame clicks, or malformed relationships.
- Existing transaction, duplication, reorder, deletion cleanup, write/reopen, and six-format relationship lifecycle rules apply unchanged.
- Existing table creation bytes remain identical when every cell hyperlink is omitted/runtime-`undefined`.
- Every task ends with review, one scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`, generated workspace tarballs, or retained `/tmp` evidence.

---

### Task 1: Core creation, readonly snapshot, lifecycle, and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/table-cell-hyperlink.internal.ts`
- Create: `packages/model/src/table-cell-hyperlink.internal.test.ts`
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: existing `Hyperlink`, `NormalizedHyperlink`, `normalizeHyperlink()`, `readTextRunHyperlink()`, `RelationshipInput`, `HYPERLINK_RELATIONSHIP_TYPE`, `SLIDE_RELATIONSHIP_TYPE`, `relativeRelationshipTarget()`, `renderRichTextParagraphs()`, and table creation transaction.
- Produces: `readTableCellHyperlink(xml, cell, context)`, `TableCellHyperlinkRelationshipIds`, `AddTableCellOptions.hyperlink`, conditional cell-run rendering, and `TableCell.hyperlink` snapshots.

- [ ] **Step 1: Add focused failing direct-cell reader tests**

Create `packages/model/src/table-cell-hyperlink.internal.test.ts` with a namespace-complete cell fixture and relationship context:

```ts
const parseCell = (contents: string) => {
  const xml = LosslessXmlDocument.parse(
    `<a:tc xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">` +
    `<a:txBody>${contents}</a:txBody><a:tcPr/></a:tc>`,
  );
  return { xml, cell: xml.roots[0]! };
};

const context = (relationships: readonly Relationship[]) => ({
  relationships,
  slidePartUris: ['/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml'],
});
```

Import the not-yet-existing helper:

```ts
import { readTableCellHyperlink } from './table-cell-hyperlink.internal.js';
```

Require these exact supported cases:

- one direct paragraph/run/rPr external click with absent tooltip returns `{ url }`;
- explicit `tooltip=""` remains `{ url, tooltip: '' }`;
- internal slide relationship plus `ppaction://hlinksldjump` returns `{ slide: 2 }`;
- PptxGenJS external `action="" invalidUrl="" history="1"` imports;
- returned objects are recursively detached/frozen across repeated reads.

Require `undefined` for no click, missing/repeated/wrong-namespace `txBody`, paragraph, run, or `rPr`; `fld`/`br`; multiple paragraphs/runs; descendant lookalikes; duplicate clicks; empty/dangling/duplicate IDs; wrong relationship type/mode; unsupported action; unresolved or ambiguously resolved slide target; and malformed click children.

- [ ] **Step 2: Add failing table normalization and rendering tests**

Extend `packages/model/src/table-create.internal.test.ts` with this normalized matrix:

```ts
const url = { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' };
const slide = Object.assign(Object.create(null), { slide: 2, tooltip: '' });
const definition = normalizeTableDefinition([[
  { text: 'URL', options: { hyperlink: url } },
  { text: 'Slide', options: Object.assign(Object.create(null), { hyperlink: slide }) },
  { text: 'Plain', options: { hyperlink: undefined } },
]], {});
```

Require detached normalized hyperlinks and no mutation after changing the source objects. Render with:

```ts
const xml = renderTableGraphicFrame(7, definition, undefined, undefined, [[
  'rId7',
  'rId8',
  undefined,
]]);
```

Require exactly two direct run clicks, one escaped tooltip, one internal action, two `u="sng"` tokens, `xmlns:r` on the graphic frame, and no whole-frame click. Require empty linked text to retain one linked run. Require no-link, empty-options, and runtime-`undefined` inputs to remain byte-identical to the current baseline without `xmlns:r`.

Require `TypeError` for row count, column count, and hyperlink/ID presence mismatches. Add invalid-option cases for unknown keys, class/accessor/inherited/symbol properties, both/neither targets, empty/invalid URL, invalid tooltip, invalid slide number, and `_rId`.

- [ ] **Step 3: Run the new focused tests and verify expected failure**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  --reporter=dot
```

Expected: fail because `readTableCellHyperlink` is missing and `hyperlink` is not an accepted table-cell option or rendered property.

- [ ] **Step 4: Implement the strict reader and table renderer**

Create `packages/model/src/table-cell-hyperlink.internal.ts` with:

```ts
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

export function readTableCellHyperlink(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): NormalizedHyperlink | undefined {
  if (cell.localName !== 'tc' || namespaceUri(cell) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const bodies = directChildren(cell).filter((child) =>
    child.localName === 'txBody' && namespaceUri(child) === DRAWING_NAMESPACE);
  if (bodies.length !== 1) return undefined;
  const paragraphs = directChildren(bodies[0]!).filter((child) =>
    child.localName === 'p' && namespaceUri(child) === DRAWING_NAMESPACE);
  if (paragraphs.length !== 1) return undefined;

  const paragraphChildren = directChildren(paragraphs[0]!);
  if (paragraphChildren.some((child) =>
    namespaceUri(child) !== DRAWING_NAMESPACE
    || !['pPr', 'r', 'endParaRPr'].includes(child.localName))) {
    return undefined;
  }
  const runs = paragraphChildren.filter(({ localName }) => localName === 'r');
  if (runs.length !== 1) return undefined;
  const runChildren = directChildren(runs[0]!);
  if (runChildren.some((child) =>
    namespaceUri(child) !== DRAWING_NAMESPACE
    || !['rPr', 't'].includes(child.localName))) {
    return undefined;
  }
  const properties = runChildren.filter(({ localName }) => localName === 'rPr');
  const texts = runChildren.filter(({ localName }) => localName === 't');
  if (properties.length !== 1 || texts.length !== 1) return undefined;
  return readTextRunHyperlink(properties[0]!, context);
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function namespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declaration = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const matches = current.attributes.filter(({ name }) => name === declaration);
    if (matches.length > 1) return undefined;
    if (matches[0]) return matches[0].value;
  }
  return undefined;
}
```

Use direct-child and namespace checks; reject text-bearing structural ambiguity and never scan descendants for a fallback run.

In `packages/model/src/table-create.internal.ts`:

```ts
export type TableCellHyperlinkRelationshipIds =
  readonly (readonly (string | undefined)[])[];

interface NormalizedTableCell {
  readonly hyperlink?: NormalizedHyperlink;
}
```

Add `hyperlink` to the exact cell-option key list, normalize it with:

```ts
const hyperlink = options.hyperlink === undefined
  ? undefined
  : normalizeHyperlink(options.hyperlink, `${context} hyperlink`);
```

Extend `renderTableGraphicFrame()` with an optional final relationship-ID matrix. Validate the matrix shape against `definition.rows`. Pass each optional ID to `renderTableCell()`, and call:

```ts
renderRichTextParagraphs(paragraphs, {
  ...(cell.alignment === undefined ? {} : { defaultAlign: cell.alignment }),
  ...(cell.hyperlink === undefined
    ? {}
    : {
        defaultHyperlink: cell.hyperlink,
        hyperlinkRelationshipId: relationshipId!,
      }),
});
```

Conditionally append ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` to the `p:graphicFrame` root when at least one cell is linked. Keep the no-link output byte-identical.

- [ ] **Step 5: Add public types and two-phase relationship creation**

In `packages/model/src/slide.ts`, add `readonly hyperlink?: Hyperlink` to `AddTableCellOptions` and a `PreparedTableCellHyperlink` record with row, column, and `RelationshipInput`.

Before any table-cell relationship mutation, walk `definition.rows` and prepare all inputs:

```ts
const relationship = hyperlink.url !== undefined
  ? {
      type: HYPERLINK_RELATIONSHIP_TYPE,
      target: hyperlink.url,
      targetMode: 'External' as const,
    }
  : {
      type: SLIDE_RELATIONSHIP_TYPE,
      target: relativeRelationshipTarget(this.partUri, target.partUri),
      targetMode: 'Internal' as const,
    };
```

If `this.presentation.slides[hyperlink.slide - 1]` is absent, throw:

```ts
new RangeError(`Table cell ${rowIndex},${columnIndex} hyperlink slide ${hyperlink.slide} is out of range`)
```

Only after the full walk succeeds, allocate one relationship per prepared cell and build the exact optional ID matrix. Pass it to `renderTableGraphicFrame()`. Keep all work in the existing `opcPackage.transaction()` so later rendering/shape-resolution failures roll back relationships and XML together.

In `packages/model/src/shapes.ts`, add `readonly hyperlink?: Hyperlink` to `TableCell`. In `TableModel.rows`, call `readTableCellHyperlink()` with current slide relationships and current presentation slide-part order and include the property only when supported.

- [ ] **Step 6: Run focused helpers and current hyperlink/table regressions**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts \
  packages/model/src/table-cell-borders.internal.test.ts \
  packages/model/src/table-cell-fill.internal.test.ts \
  packages/model/src/table-cell-margins.internal.test.ts \
  --reporter=dot
```

Expected: all pass and the no-link table creation snapshots remain unchanged.

- [ ] **Step 7: Add model lifecycle and atomicity coverage**

Extend `packages/model/src/model.test.ts` with native creation that uses two equal-target external cells, one internal cell, and one plain cell. Require:

```ts
expect(table.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
  { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' },
  { url: 'https://example.com?a=1&b=2', tooltip: '' },
  { slide: 2 },
  undefined,
]);
```

Require three unique relationship IDs, two independent equal external targets, one internal target, exact run ownership, automatic underline, absent versus empty tooltip, frozen snapshots, and input detachment. Then cover:

- `setCellText()` preserving the first link and relationship;
- border/fill/margin/align/direction/fit and table-level bulk edits preserving all links;
- target slide move changing `{ slide: 2 }` to `{ slide: 1 }` and moving back restoring it;
- source duplication preserving external links and retargeting a self-link to the duplicate;
- deleting a non-self internal target clearing incoming cell clicks/relationships while retaining text and direct underline;
- outer transaction rollback, stable table identity, write/reopen, and all six `PRESENTATION_FORMAT_PROFILES`;
- a valid early URL plus a later out-of-range slide rejecting with unchanged slide bytes, relationships, parts, diagnostics, and mutation journal;
- an injected post-relationship render/shape-tree failure rolling back every added relationship.

- [ ] **Step 8: Add SDK and root-package type/runtime coverage**

Extend `packages/sdk/src/index.test.ts` and `packages/pptx/src/index.test.ts` with public `Hyperlink`, `AddTableCellOptions`, and `TableCell` assignments:

```ts
const options: AddTableCellOptions = {
  hyperlink: { url: 'https://example.com', tooltip: '' },
};
const table = slide.addTable([[{ text: 'Linked', options }]]);
const cell: TableCell = table.rows[0]!.cells[0]!;
const hyperlink: Hyperlink | undefined = cell.hyperlink;
void hyperlink;
```

Inside an unreachable type block require errors for `{}`, both branches, numeric URL, string slide, numeric tooltip, `_rId`, target shorthand, and table-level `hyperlink`. At runtime cover immediate/reopened snapshots, relation ownership, text-edit preservation, invalid-target failure isolation, duplicate/reorder/delete behavior, and zero package validation errors.

- [ ] **Step 9: Add real PptxGenJS final-state conformance**

Extend `packages/pptxgenjs-adapter/src/index.test.ts`. Generate with public PptxGenJS 4.0.1:

```ts
slide.addTable([[
  { text: 'URL', options: {
    hyperlink: { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' },
  } },
  { text: 'Slide', options: { hyperlink: { slide: 2 } } },
  { text: 'Plain' },
]], { x: 1, y: 1, w: 8, h: 1 });
```

Require imported snapshots `{ url, tooltip }`, `{ slide: 2, tooltip: '' }`, and `undefined`; exact external/internal relationship targets; URL extra attributes accepted; internal action; `u="sng"`; and write/reopen preservation. Create the native equivalent and compare final target/action/tooltip/underline semantics while recording native omitted-tooltip absence versus PptxGenJS explicit empty. Assert PptxGenJS caller `_rId` mutation separately and require native input detachment. Keep nested/rich table-cell text outside the conformance assertion.

- [ ] **Step 10: Run focused, type, bundle, full, and performance gates**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts \
  packages/model/src/table-cell-borders.internal.test.ts \
  packages/model/src/table-cell-fill.internal.test.ts \
  packages/model/src/table-cell-margins.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Inspect generated `model/slide.d.ts` and `model/shapes.d.ts` for the exact two optional properties.

- [ ] **Step 11: Review, commit, push, and verify core behavior**

Review direct ownership, namespace safety, complete prevalidation, one-relationship-per-cell allocation, absent/empty tooltip, input detachment, no-link byte identity, rollback, snapshot strictness, lifecycle, public types, and real PptxGenJS output. Stage only the ten declared core files:

```sh
git add \
  packages/model/src/table-cell-hyperlink.internal.ts \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/slide.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "feat: create table-cell hyperlinks"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0` before Task 2.

---

### Task 2: Actual-package declarations, Node/browser/CLI, and real-Chrome proof

**Files:**
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: Task 1 table-cell input/snapshot, actual tarball declarations, root/browser exports, installed CLI, and checked-in Chrome callback.
- Produces: stable `tableCellHyperlinks: true`, detailed `tableCellHyperlinkState`, `tableCellHyperlinksInspect: true`, and one retained packed PPTX.

- [ ] **Step 1: Lock the two packed declaration blocks**

In `scripts/build-npm-package-types.mjs`, slice `export interface AddTableCellOptions` through `export interface AddTableCell` in `model/slide.d.ts` and `export interface TableCell` through `export interface TableRow` in `model/shapes.d.ts`. Require:

```ts
readonly hyperlink?: Hyperlink;
```

in both scoped blocks. Throw feature-specific errors so the check cannot pass from `AddTextOptions`, `ShapeModel`, or `RichTextRunStyle`.

- [ ] **Step 2: Add installed Node runtime evidence**

In `scripts/smoke-npm-package.mjs`, create two slides and a one-by-four table with two equal URL targets (one normal tooltip, one explicit empty), one slide target, and one plain cell. Retain the caller objects, mutate them after creation, and record this canonical state:

```ts
const tableCellHyperlinkState = {
  immediate: [
    { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' },
    { url: 'https://example.com?a=1&b=2', tooltip: '' },
    { slide: 2 },
    null,
  ],
  inputDetached: true,
  snapshotsFrozen: true,
  independentRelationships: true,
  textEditPreserved: { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' },
  movedInternal: { slide: 1 },
  restoredInternal: { slide: 2 },
  reopened: [
    { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' },
    { url: 'https://example.com?a=1&b=2', tooltip: '' },
    { slide: 2 },
    null,
  ],
  invalidError: {
    name: 'RangeError',
    message: 'Table cell 0,1 hyperlink slide 99 is out of range',
  },
  failureIsolation: true,
  validationErrors: 0,
};
```

Use `null` only in serialized smoke state for absent optional properties. Require three unique direct click IDs, two independent hyperlink relationships with the same target, one internal slide relationship, matching direct `u="sng"`, no graphic-frame click, and no caller `_rId`. Move the target to index 0 and back, call `setCellText(0, 0, 'URL edited')`, reject a later invalid linked cell with unchanged bytes/relationships/journal, write `table-cell-hyperlinks-smoke.pptx`, and reopen.

Add the boolean/state to `apiChecks` and top-level JSON. Copy the retained deck beside the input tarball before deleting the temporary consumer workspace.

- [ ] **Step 3: Add installed TypeScript and browser-condition evidence**

Add `type AddTableCellOptions`, `type Hyperlink`, and `type TableCell` to the generated consumer. Require valid creation/read assignments and these compile-time failures:

```ts
// @ts-expect-error table-cell hyperlink requires exactly one target
const missingTarget: AddTableCellOptions = { hyperlink: {} };
// @ts-expect-error table-cell hyperlink branches are mutually exclusive
const bothTargets: AddTableCellOptions = {
  hyperlink: { url: 'https://example.com', slide: 2 },
};
// @ts-expect-error relationship IDs are internal
const relationshipEscape: AddTableCellOptions = {
  hyperlink: { url: 'https://example.com', _rId: 'rId9' },
};
// @ts-expect-error there is no table-level hyperlink default
slide.addTable([['A']], { hyperlink: { url: 'https://example.com' } });
```

In the generated browser consumer, reproduce the exact Node state with `writeBlob()` and require identical snapshots, independent relationships, target reorder behavior, reopen state, invalid error, failure isolation, and zero diagnostics.

- [ ] **Step 4: Add installed CLI and `pptx-inspect` evidence**

Use the installed CLI to inspect, validate, list slides, and read the retained table slide and its relationship part. Require two slides, one table, four physical cells, exactly three direct run `hlinkClick` elements, exactly two hyperlink relationships with equal external targets, exactly one internal slide relationship, no `p:cNvPr/a:hlinkClick`, one absent tooltip, one explicit empty tooltip, one internal action, and three `u="sng"` run properties.

Run the host broad-to-narrow workflow:

```sh
pptx-inspect --json package inspect table-cell-hyperlinks-smoke.pptx
pptx-inspect --json package validate table-cell-hyperlinks-smoke.pptx --profile powerpoint-2010
pptx-inspect --json slides list table-cell-hyperlinks-smoke.pptx
pptx-inspect --json part read table-cell-hyperlinks-smoke.pptx /ppt/slides/slide1.xml
pptx-inspect --json part read table-cell-hyperlinks-smoke.pptx /ppt/slides/_rels/slide1.xml.rels
```

Require zero validation errors and warnings before accepting narrow reads. Add `tableCellHyperlinksInspect: true` to the compact result.

- [ ] **Step 5: Extend the real-Chrome callback**

Add the same `tableCellHyperlinkState` and expected JSON to `scripts/playwright-browser-smoke.js`. Require immediate/read/frozen state, input detachment, independent relationships, text-edit preservation, target move/restore, reopen, invalid failure isolation, zero document diagnostics, and global `errorCounts: { console: 0, page: 0, network: 0 }`.

- [ ] **Step 6: Build, pack, install, and run all proof gates**

Run:

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
table_cell_hyperlink_artifacts=$(mktemp -d /tmp/pptx-table-cell-hyperlink-artifacts.XXXXXX)
(cd packages/pptx && npm pack --ignore-scripts --pack-destination "$table_cell_hyperlink_artifacts")
node scripts/smoke-npm-package.mjs "$table_cell_hyperlink_artifacts/jiayunxie-pptx-0.1.0.tgz"
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Extract the actual tarball beneath the same evidence directory, serve its browser module over loopback, and run the checked-in callback in installed Google Chrome through the bundled Playwright runtime. Retain tarball file count, SHA-256, installed Node/types/browser/CLI output, `pptx-inspect` JSON, full/compact Chrome state, and evidence directory outside the repository.

- [ ] **Step 7: Review, commit, push, and verify package proof**

Review declaration scoping, input detachment, equal-target independence, exact run ownership, absent/empty tooltip, target identity, Node/browser parity, CLI relationship inspection, zero browser errors, and absence of generated repository artifacts. Stage only the three scripts:

```sh
git add \
  scripts/build-npm-package-types.mjs \
  scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed table-cell hyperlinks"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0` before Task 3.

---

### Task 3: Public documentation and progress closeout

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: final Task 1 semantics and Task 2 measured source/package/browser/CLI evidence.
- Produces: consistent public creation/snapshot documentation, PptxGenJS boundary, final evidence record, revised remaining-work list, and table-cell hyperlink editing as the next focused item.

- [ ] **Step 1: Add public examples and exact semantics**

In all three public README/API surfaces, add one compact example with URL, internal slide, absent tooltip, explicit empty tooltip, immediate `TableCell.hyperlink`, target reorder, `setCellText()` preservation, and write/reopen. State that links are attached to the plain cell's one direct run, equal-target cells own independent relationships, and no table-level default or editor is included.

- [ ] **Step 2: Update compatibility and progress records**

Move plain single-run table-cell hyperlink creation/read snapshots to supported in the compatibility matrix and detailed table section. Keep rich/multi-paragraph cell links and editing pending. Record PptxGenJS legal URL/slide final-state import, explicit-empty tooltip and extra-attribute differences, input `_rId` mutation difference, focused/full/performance totals, actual tarball count/SHA, installed Node/types/browser/CLI, `pptx-inspect`, real Chrome, zero error counts, evidence path, and implementation commit hashes.

- [ ] **Step 3: Update changelog and remaining-work statements**

Add three changelog bullets for core relationship/snapshot semantics, PptxGenJS boundary, and actual-package/browser proof. Remove stale statements that all table-cell hyperlinks are unsupported. Keep overall parity approximately 97% until the final peer/client audit and select table-cell hyperlink editing/clearing as the next independently testable item before rich cell text.

- [ ] **Step 4: Run final documentation and regression gates**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 5: Review, commit, push, verify, and report**

Review API types, direct run ownership, one relationship per cell, absent/empty tooltip, internal target identity, readonly strictness, PptxGenJS differences, measured totals/hashes/evidence, next-item consistency, and `git diff --check`. Stage only the six documentation files:

```sh
git add \
  README.md \
  packages/pptx/README.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  docs/implementation-progress.md \
  CHANGELOG.md
git diff --cached --check
git commit -m "docs: document table-cell hyperlinks"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, report completed/remaining items and overall progress, then begin the table-cell hyperlink editing design item.
