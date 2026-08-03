# Table-Cell Rich Text Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. This repository task executes inline and performs a review, commit, and push after each independently testable deliverable.

**Goal:** Support complete structured rich/multi-paragraph table-cell text creation, detached snapshots, paragraph-aware plain projection, and lossless whole-replacement editing with the existing `RichTextParagraph` value model.

**Architecture:** Add one table-cell ownership boundary that validates the exact direct DrawingML `txBody`, delegates rich parsing/rendering/replacement to the existing rich-text engine, and exposes a safe plain projection. Carry normalized paragraphs and per-run hyperlink relationship matrices through table creation, then reuse the established shape-rich-text relationship algorithm for indexed cell replacement while preserving cell/table structure.

**Tech Stack:** TypeScript 5.8, Vitest 3, lossless OOXML/OPC transactions, pnpm workspaces, PptxGenJS 4.0.1 conformance fixtures, packed ESM/CJS/browser exports, CLI/Inspector, Google Chrome, and PowerPoint 2010 validation.

## Global Constraints

- Public creation is exactly `AddTableCell.text: string | readonly RichTextParagraph[]`; bare arrays are not cells.
- Public snapshots expose `TableCell.text: string` and detached `TableCell.richText: readonly RichTextParagraph[]`.
- Public replacement is exactly `setCellRichText(rowIndex: number, columnIndex: number, value: readonly RichTextParagraph[]): void` with zero-based physical coordinates.
- Single-line string cell output remains byte-identical; CR/LF strings normalize to `\n`, split into direct paragraphs, and preserve empty lines.
- Reuse the existing rich-text normalizer, reader, renderer, equality check, relationship binding, and whole-replacement engine; do not introduce table-only paragraph/run public types.
- Cell-option hyperlinks are inherited only by runs without a local override; explicit run hyperlinks use independent relationships and `hyperlink: false` suppresses inheritance.
- All normalization and internal-slide target resolution complete before observable package mutation; caller objects and returned snapshots remain detached.
- Same-value rich replacement is an exact no-op; replacement preserves `bodyPr`, `lstStyle`, `tcPr`, table geometry, neighboring cells, and table model identity.
- Relationship target equality reuses IDs, unique target changes update in place, shared target changes clone on write, and unreferenced replaced relationships are removed.
- `setCellText()` remains valid only for an exact safe single-paragraph/single-run cell and rejects rich/multi structures with `ModelParseError`.
- Do not add merge/span, row/column CRUD, auto-page, repeated headers, content measurement, outer table/cell font defaults, or `tableToSlides` in this item.
- Every deliverable is reviewed, committed, pushed, and followed by local/remote divergence verification of `0 0`.

---

## File Structure

- Create `packages/model/src/table-cell-rich-text.internal.ts`: exact cell text-body ownership, rich snapshot/state delegation, plain projection, and safe plain-text target resolution.
- Create `packages/model/src/table-cell-rich-text.internal.test.ts`: namespace, ownership, projection, detachment, malformed input, and safe plain-edit fixtures.
- Modify `packages/model/src/table-create.internal.ts`: normalized paragraph storage, CR/LF splitting, rich rendering, and cell-default/run-local relationship matrices.
- Modify `packages/model/src/table-create.internal.test.ts`: single-line parity, rich creation, style families, hyperlink inheritance/suppression, and input validation.
- Modify `packages/model/src/shapes.ts`: `TableCell.richText`, paragraph-aware `text`, safe `setCellText()`, and public `setCellRichText()` delegation.
- Modify `packages/model/src/slide.ts`: public creation union, prepared table run hyperlinks, transactional indexed replacement, ID reuse/COW/GC, and slide lifecycle integration.
- Modify `packages/model/src/model.test.ts`: public snapshot/editor, duplicate, rollback, relationship, slide reorder/delete, six-format, and write/reopen coverage.
- Modify `packages/sdk/src/index.test.ts`: SDK/root runtime and compile-time API contract.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: legal PptxGenJS rich-cell import and documented divergence coverage.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, and `CHANGELOG.md`: support status, API examples, parity boundary, proof, and remaining work.

### Task 1: Exact Table-Cell Rich-Text Boundary

**Files:**
- Create: `packages/model/src/table-cell-rich-text.internal.ts`
- Create: `packages/model/src/table-cell-rich-text.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, `ModelParseError`, `ShapeHyperlinkReadContext`, `RichTextParagraph`, `ReadRichTextState`, `readRichTextState()`, and DrawingML namespace ownership.
- Produces: `readTableCellRichText()`, `readTableCellText()`, `requireEditableTableCellRichTextState()`, and `requireEditablePlainTableCellText()`.

- [ ] **Step 1: Write exact-ownership and projection tests**

Create fixtures with namespace-correct direct `a:tc/a:txBody/a:p`, multiple runs, `a:fld`, `a:br`, two paragraphs, empty paragraphs, a run hyperlink, and detached mutations. Assert the public helper contract:

```ts
expect(readTableCellText(xml, cell, context)).toBe('A\nB\n\nC');
expect(readTableCellRichText(xml, cell, context)).toEqual([
  { runs: [{ text: 'A' }, { text: 'B', softBreakBefore: true }] },
  { runs: [] },
  { runs: [{ text: 'C', style: { bold: true } }] },
]);
expect(requireEditablePlainTableCellText(xml, plainCell, '/ppt/slides/slide1.xml')
  .localName).toBe('t');
expect(() => requireEditablePlainTableCellText(
  xml,
  richCell,
  '/ppt/slides/slide1.xml',
)).toThrow(ModelParseError);
```

Also assert missing/repeated/wrong-namespace direct `txBody` returns `[]`/`''` for snapshots, descendant impostors are ignored, and edit-state resolution rejects zero or repeated direct paragraphs while physical index handling remains outside this helper.

- [ ] **Step 2: Run the focused tests and verify the red state**

Run:

```bash
pnpm exec vitest run packages/model/src/table-cell-rich-text.internal.test.ts
```

Expected: failure because `table-cell-rich-text.internal.ts` and its exports do not exist.

- [ ] **Step 3: Implement the narrow ownership boundary**

Implement these exact signatures:

```ts
export function readTableCellRichText(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): readonly RichTextParagraph[];

export function readTableCellText(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): string;

export function requireEditableTableCellRichTextState(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
  partUri: string,
): ReadRichTextState;

export function requireEditablePlainTableCellText(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  partUri: string,
): XmlElement;
```

Resolve exactly one direct namespace-correct `txBody`; snapshot helpers return empty state for unsafe ownership. `requireEditableTableCellRichTextState()` additionally requires at least one direct namespace-correct paragraph before delegating to `readRichTextState()`. `requireEditablePlainTableCellText()` accepts only one direct paragraph whose DrawingML children are optional `pPr`, one `r` containing one optional `rPr` plus one `t`, and optional `endParaRPr`; reject fields, breaks, multiple runs/paragraphs, repeated children, foreign namespaces, and descendant substitutes.

Build `readTableCellText()` from the structured snapshot: join paragraphs with `\n`, join runs in order, and insert `\n` before a run with `softBreakBefore: true`. Freeze the outer snapshot through the existing reader behavior and never return source XML nodes.

- [ ] **Step 4: Run tests, typecheck the model package, and review the boundary**

Run:

```bash
pnpm exec vitest run packages/model/src/table-cell-rich-text.internal.test.ts
pnpm exec tsc -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
```

Review for direct-child ownership, namespace correctness, no descendant scans, exact error part URI, and no duplicated rich-text parsing.

- [ ] **Step 5: Commit, push, and confirm synchronization**

```bash
git add packages/model/src/table-cell-rich-text.internal.ts \
  packages/model/src/table-cell-rich-text.internal.test.ts
git commit -m "feat: read table-cell rich text"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 2: Rich and Multi-Paragraph Table Creation

**Files:**
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`
- Modify: `packages/model/src/slide.ts`

**Interfaces:**
- Consumes: Task 1 snapshot semantics, `normalizeRichText()`, `renderRichTextParagraphs()`, `NormalizedHyperlink`, and `RichTextRunHyperlinkRelationshipIds`.
- Produces: normalized cell paragraphs plus cell-default/run-local relationship bindings accepted by `renderTableGraphicFrame()`.

- [ ] **Step 1: Add failing normalization and rendering tests**

Add assertions for these inputs without mutating the originals:

```ts
const rows = [[
  'one\r\ntwo\rempty\n',
  { text: [{
    align: 'right',
    bullet: { kind: 'unordered' },
    runs: [
      { text: 'Bold', style: { bold: true } },
      { text: ' local', style: { hyperlink: { url: 'https://run.example' } } },
      { text: ' suppressed', style: { hyperlink: false, softBreakBefore: true } },
    ],
  }], options: { hyperlink: { url: 'https://cell.example' } } },
]] as const;
```

Assert CR/LF output has four direct paragraphs including the trailing empty paragraph; explicit paragraph alignment overrides cell/table alignment; every existing run/paragraph style survives; inherited runs share the cell-default ID; the local run uses its own ID; the false run has no click; and a single-line string produces exactly the pre-change XML bytes.

Add strict rejection cases for bare arrays, empty rich arrays, non-array runs, unknown/accessor/symbol/inherited fields, invalid style values, invalid hyperlink targets, and dangling internal slide targets. Verify definition normalization does not mutate rows, runs, styles, or hyperlink objects.

- [ ] **Step 2: Run the creation tests and verify focused failures**

Run:

```bash
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  -t "table-cell rich text|table creation"
```

Expected: failures because object-cell `text` still requires a one-line string and rendering has no run-link relationship matrix.

- [ ] **Step 3: Store normalized paragraphs and render them directly**

Change the internal cell shape to:

```ts
interface NormalizedTableCell {
  readonly paragraphs: ReturnType<typeof normalizeRichText>;
  readonly alignment?: TextAlignment;
  readonly hyperlink?: NormalizedHyperlink;
  // existing borders/fill/margins/direction/fit/vertical alignment remain
}

export interface TableCellRichTextRelationshipIds {
  readonly defaultHyperlinkRelationshipId?: string;
  readonly runHyperlinkRelationshipIds: RichTextRunHyperlinkRelationshipIds;
}

export type TableCellRichTextRelationshipMatrix =
  readonly (readonly TableCellRichTextRelationshipIds[])[];
```

Normalize strings with XML-character validation, `value.replace(/\r\n?/g, '\n')`, and one plain paragraph per split segment. Normalize structured `text` with `normalizeRichText()` and retain the detached normalized result. Preserve the previous single-line normalized paragraph shape so `renderRichTextParagraphs()` emits byte-identical XML.

Change `renderTableGraphicFrame()` and `renderTableCell()` to consume the relationship matrix, pass `defaultAlign`, the used cell-default hyperlink/ID, and per-run IDs to `renderRichTextParagraphs()`. Only require/create a default ID when at least one run inherits it. Add `xmlns:r` when any default or explicit run link will render.

- [ ] **Step 4: Prepare every table-cell relationship before mutation**

In `slide.ts`, replace the plain-cell prepared record with a positioned record that distinguishes default and run-local links:

```ts
interface PreparedTableCellHyperlink {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly paragraphIndex?: number;
  readonly runIndex?: number;
  readonly relationship: RelationshipInput;
}
```

`prepareTableCellHyperlinks()` walks normalized paragraphs, resolves all internal slides, prepares one cell-default relationship only when inherited by at least one run, and prepares one relationship for each explicit run hyperlink. `createTableCellHyperlinkRelationships()` allocates the exact matrix inside the existing add-table transaction. It never allocates a relationship for `hyperlink: false` or an unused cell default.

- [ ] **Step 5: Run focused and regression gates, then review**

Run:

```bash
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts
pnpm exec vitest run packages/model/src/model.test.ts -t "table creation"
pnpm typecheck
git diff --check
```

Review single-line byte parity, empty-line preservation, style coverage, no input mutation, relationship cardinality, and validation-before-observable-mutation.

- [ ] **Step 6: Commit, push, and confirm synchronization**

```bash
git add packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts packages/model/src/slide.ts
git commit -m "feat: create rich table-cell text"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 3: Public Snapshot and Safe Indexed Replacement

**Files:**
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 readers/ownership helpers, Task 2 normalized creation, `readDirectTablePhysicalCellMatrix()`, `normalizeRichText()`, `richTextParagraphsEqual()`, `replaceRichText()`, and hyperlink reference counting.
- Produces: `TableCell.richText`, paragraph-aware `TableCell.text`, safe `TableModel.setCellText()`, `TableModel.setCellRichText()`, and `SlideModel.setTableCellRichText()`.

- [ ] **Step 1: Add failing public model and transaction tests**

Create one table containing plain, multi-paragraph, multi-run, soft-break, field, linked, and empty cells. Assert:

```ts
expect(table.rows[0]![0]!.text).toBe('First\nSecond\nsoft');
expect(table.rows[0]![0]!.richText).toEqual(expectedParagraphs);
table.setCellRichText(0, 0, replacement);
expect(table.rows[0]![0]!.richText).toEqual(replacement);
expect(() => table.setCellText(0, 0, 'unsafe')).toThrow(ModelParseError);
table.setCellText(0, 1, 'safe');
expect(table.rows[0]![1]!.text).toBe('safe');
```

Mutate the source arrays and returned snapshots and prove the model is unchanged. Capture slide XML, relationships, ZIP dates, mutation journal, table identity, body properties, list style, cell properties, transform, and neighbor XML before a same-value edit; require exact equality afterward. Add out-of-range and malformed direct-ownership cases with the slide part URI in `ModelParseError`.

- [ ] **Step 2: Run focused model tests and verify failures**

Run:

```bash
pnpm exec vitest run packages/model/src/model.test.ts \
  -t "table-cell rich text|safe table-cell plain text"
```

Expected: failures because `richText` and `setCellRichText()` are absent and `setCellText()` accepts unsafe descendant state.

- [ ] **Step 3: Expose snapshots and public methods**

Change the public declarations exactly:

```ts
export interface TableCell {
  readonly text: string;
  readonly richText: readonly RichTextParagraph[];
  // existing scalar properties remain
}

export class TableModel extends BaseShapeModel {
  setCellRichText(
    rowIndex: number,
    columnIndex: number,
    value: readonly RichTextParagraph[],
  ): void;
}
```

In `rows`, use direct physical rows/cells and Task 1 readers with the slide relationship context. Keep existing strict scalar hyperlink/alignment readers unchanged. Route `setCellText()` through `requireEditablePlainTableCellText()` inside an OPC transaction and only serialize when text changes. Delegate `setCellRichText()` to `SlideModel.setTableCellRichText()`.

- [ ] **Step 4: Implement transactional cell rich replacement**

Add:

```ts
setTableCellRichText(
  id: number,
  rowIndex: number,
  columnIndex: number,
  value: readonly RichTextParagraph[],
): void;
```

Normalize and resolve all requested run hyperlink targets before mutation. Inside one transaction, resolve the exact table physical cell, require safe rich ownership, normalize current paragraphs, and return before any write when values are equal. Reuse same-position relationship IDs for equal targets; update a unique relationship in place; allocate a COW relationship for shared target changes; call `replaceRichText()` on the exact cell; save once; then remove only previous DrawingML hyperlink relationships whose slide reference count is zero.

Preserve direct `bodyPr`, `lstStyle`, `tcPr`, table transform, unrelated paragraph template children, neighboring cells, and model identity. Let the outer OPC transaction roll back `addRelationship`, `updateRelationship`, `removeRelationship`, or `setXml` failures.

- [ ] **Step 5: Run model, type, and build gates and review**

Run:

```bash
pnpm exec vitest run packages/model/src/table-cell-rich-text.internal.test.ts \
  packages/model/src/table-create.internal.test.ts packages/model/src/model.test.ts \
  -t "table-cell rich text|safe table-cell plain text|table creation"
pnpm typecheck
pnpm build
git diff --check
```

Review exact physical indexing, current/request equality, relationship ownership, mutation isolation, preserved source spans, and stable public snapshots.

- [ ] **Step 6: Commit, push, and confirm synchronization**

```bash
git add packages/model/src/shapes.ts packages/model/src/slide.ts \
  packages/model/src/model.test.ts
git commit -m "feat: edit table-cell rich text"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 4: Relationship and Slide Lifecycle Conformance

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: Tasks 2-3 public creation/read/edit API and existing duplicate/insert/delete/reorder/write/reopen helpers.
- Produces: regression proof for default inheritance, explicit links, ID reuse/COW/GC, six formats, SDK/root declarations, and PptxGenJS legal-output import.

- [ ] **Step 1: Add complete hyperlink and slide lifecycle tests**

Cover these exact transitions:

```ts
table.setCellRichText(0, 0, [{ runs: [
  { text: 'same', style: { hyperlink: { url: 'https://same.example' } } },
  { text: 'new', style: { hyperlink: { slide: 2, tooltip: 'Target' } } },
] }]);
```

Require same-position/same-target ID reuse, tooltip-only reuse, unique URL-to-slide update in place, shared ID target-change COW, reference-safe clear, and last-reference GC across rich cells and text shapes. Test internal target ordinal after insert/delete/reorder, duplicate self-link retargeting, target deletion cleanup for source and duplicate, and unrelated link survival.

Inject failures into relationship add/update/remove and slide XML write, including an outer transaction rollback; compare all slide parts, relationships, ZIP directory/file dates, model identity, and mutation journal with the pre-call state.

- [ ] **Step 2: Add SDK/root TypeScript contract tests**

Add positive compile/runtime coverage:

```ts
const cell: AddTableCell = { text: [{ runs: [{ text: 'Rich' }] }] };
const snapshot: readonly RichTextParagraph[] = table.rows[0]![0]!.richText;
table.setCellRichText(0, 0, snapshot);
```

Add `@ts-expect-error` cases for bare rich arrays as cells, mutable assignment to `richText`, string input to `setCellRichText()`, and unsupported PptxGenJS recursive `TableCell[]` objects. Exercise the same imports from `@pptx/sdk` and the root `pptx-ooxml` package.

- [ ] **Step 3: Add PptxGenJS 4.0.1 semantic conformance**

Generate a public PptxGenJS table with `breakLine`, CR/LF, paragraph alignment/bullet, bold/italic/font/size/color/underline, cell-default hyperlink, explicit run hyperlink, and false-equivalent plain runs. Import its actual `write({ outputType: 'uint8array' })` result and assert native `TableCell.text`/`richText` snapshots and post-import replacement.

Assert PptxGenJS may mutate caller options with `_rId`, materialize empty tooltip, and emit repeated paragraph properties, while native input remains detached and native canonical output contains one owned paragraph-properties block. Compare legal final semantics rather than proprietary implementation artifacts.

- [ ] **Step 4: Exercise duplicate, all six formats, and write/reopen**

For `pptx`, `pptm`, `potx`, `potm`, `ppsx`, and `ppsm`, create a rich linked table, duplicate its slide, edit source and duplicate independently, write, reopen, and assert structured snapshots plus URL/internal targets. Require duplicate self-links to point at the duplicate and non-self targets to preserve identity.

- [ ] **Step 5: Run focused, full, performance, type, and build gates**

Run:

```bash
pnpm exec vitest run packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts \
  -t "table-cell rich text"
pnpm typecheck
pnpm build
pnpm exec vitest run --maxWorkers=2
pnpm test:performance
git diff --check
```

Review test assertions against the design, confirm no skipped rich-cell case, record full/pass counts and core duration, and inspect the diff for unrelated changes.

- [ ] **Step 6: Commit, push, and confirm synchronization**

```bash
git add packages/model/src/model.test.ts packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "test: verify table-cell rich text lifecycle"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 5: Actual-Package, CLI, Inspector, Browser, and OOXML Proof

**Files:**
- Modify: `packages/sdk/src/index.test.ts` only if the packed smoke is checked in there; otherwise keep generated proof scripts under a fresh `/tmp/pptx-table-cell-rich-text-*` directory and commit no temporary artifacts.

**Interfaces:**
- Consumes: Task 4 built workspace and package exports.
- Produces: tarball hash/file count, installed Node/TypeScript/browser/CLI/Inspector evidence, real Chrome diagnostics, PowerPoint 2010 diagnostics, and exact part/relationship evidence.

- [ ] **Step 1: Pack and install the actual root tarball**

Create a fresh evidence directory with `mktemp -d`, run the repository’s established root-package packing workflow, record `shasum -a 256` and `tar -tf | wc -l`, and install the tarball into isolated Node and browser smoke projects. Do not include `.pnpm-store/` or `/tmp` evidence in git.

- [ ] **Step 2: Run installed Node and declaration smokes**

From the installed package, create a table with CR/LF and structured paragraphs, read `richText`, call `setCellRichText()`, write/reopen, and emit JSON containing:

```json
{
  "tableCellRichTextCreation": true,
  "tableCellRichTextSnapshot": true,
  "tableCellRichTextEditing": true,
  "tableCellRichTextLinks": true
}
```

Compile an installed-consumer TypeScript file using `moduleResolution: "NodeNext"`; verify the same types through the root export. Run the installed CLI package validate/inspect/diff commands and the `pptx-inspect` skill CLI against the generated source, edited, reopened, and PptxGenJS baseline files.

- [ ] **Step 3: Inspect OOXML and compatibility diagnostics**

Require direct table-cell paragraph/run counts, paragraph-aware text, click count, relationship type/target/mode, hyperlink ID reuse, no dangling IDs, preserved cell properties, and unchanged neighbor XML. Validate every generated file with the PowerPoint 2010 profile; accept only documented external-relationship portability warnings and require zero errors.

- [ ] **Step 4: Run the installed browser export in real Chrome**

Serve the isolated browser fixture, load the conditional browser export, perform rich table creation/read/edit/write in Google Chrome, and capture its JSON plus console/page/network diagnostics. Require all four capability flags to be true and console/page/network error counts to be zero.

- [ ] **Step 5: Review evidence and commit only checked-in proof changes**

Record the evidence directory, tarball SHA-256, file count, test counts, validation counts, relationship counts, Chrome result, and run duration. If a packed smoke assertion was added to `packages/sdk/src/index.test.ts`, run its focused test, review, commit, and push:

```bash
git add packages/sdk/src/index.test.ts
git commit -m "test: verify packed table-cell rich text"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

If no tracked file changed, retain the evidence for the documentation commit and do not create an empty commit. Expected divergence after any commit: `0 0`.

### Task 6: Documentation, Capability Status, and Final Synchronization

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: completed public API and Task 5 measured evidence.
- Produces: user-facing examples, exact support boundary, PptxGenJS semantic comparison, updated remaining-work list, and final progress percentage.

- [ ] **Step 1: Document the public API with runnable examples**

Add equivalent Chinese root and English package examples:

```ts
const table = slide.addTable([[
  { text: [
    { align: 'center', runs: [{ text: 'Title', style: { bold: true } }] },
    { runs: [{ text: 'OpenAI', style: { hyperlink: { url: 'https://openai.com' } } }] },
  ] },
]]);

console.log(table.rows[0]![0]!.text);
console.log(table.rows[0]![0]!.richText);
table.setCellRichText(0, 0, [{ runs: [{ text: 'Replaced' }] }]);
```

Explain CR/LF normalization, empty-line preservation, soft break versus paragraph break, styles, default/local hyperlinks, physical indexes, detached snapshots, exact no-op, safe `setCellText()` boundary, and relationship lifecycle.

- [ ] **Step 2: Update compatibility and remaining-work status**

Move rich/multi-paragraph table-cell text/style and run hyperlinks from partial/pending to supported. Record semantic conformance with PptxGenJS 4.0.1 and the intentional differences: no caller mutation, strict descriptors/types, no duplicate `pPr`, preserved omitted versus empty tooltip, no dangling/coercible links, and canonical native rich type.

Leave these explicit table gaps: outer table/cell font defaults, merge/colspan/rowspan, row/column CRUD, auto-page/repeated headers, `tableToSlides`, and final peer/client audit. Recalculate overall parity progress from the capability checklist rather than automatically claiming 100%.

- [ ] **Step 3: Record measured final proof**

Update the changelog and both READMEs with the actual full Vitest counts, performance result, tarball SHA-256/file count, validation errors/warnings, part/relationship counts, Chrome diagnostics, evidence directory, and the core/lifecycle/package-proof commit hashes created by Tasks 1-5.

- [ ] **Step 4: Review documentation and repository state**

Run:

```bash
rg -n "rich table-cell|table-cell rich|表格单元格富文本|remaining|仍.*支持" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md CHANGELOG.md
rg -n "rich/multi-paragraph table-cell.*(pending|unsupported)|富文本.*待" \
  README.md packages/pptx/README.md docs CHANGELOG.md
git diff --check
pnpm typecheck
pnpm exec vitest run --maxWorkers=2
pnpm test:performance
```

Review examples against declarations, replace stale support statements, verify evidence numbers from artifacts, and confirm `.pnpm-store/` remains untracked and unstaged.

- [ ] **Step 5: Commit, push, and report completion/remaining/progress**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md CHANGELOG.md
git commit -m "docs: document table-cell rich text"
git push origin main
git rev-list --left-right --count origin/main...HEAD
git status --short --branch
```

Expected divergence: `0 0`; the only allowed unrelated worktree entry is the existing untracked `.pnpm-store/`. Report the completed rich-cell item, the remaining ordered table items, measured overall progress, test/proof results, commit hashes, and evidence path, then start the next item without waiting for an execution choice.
