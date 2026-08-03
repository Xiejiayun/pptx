# Table-Cell Hyperlink Editing and Clearing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict lossless `TableModel.setCellHyperlink()` editing and clearing for supported plain single-run cells, including relationship reuse, clone-on-write, garbage collection, rollback, and actual-package/browser proof.

**Architecture:** Extend the focused cell reader with an editable-state boundary, add a local text-run click replacer to the shared hyperlink codec, and let `SlideModel` coordinate relationship mutation inside one OPC transaction. `TableModel` exposes only the physical-cell editor; readonly snapshots and creation remain unchanged.

**Tech Stack:** TypeScript, Vitest, lossless OOXML, JSZip-backed OPC transactions, PptxGenJS 4.0.1, tsup, npm pack, `pptx-inspect`, Playwright, and Google Chrome.

## Global Constraints

- Public API: `TableModel.setCellHyperlink(rowIndex: number, columnIndex: number, value: Hyperlink | undefined): void`.
- `TableCell.hyperlink` remains readonly and detached; no live cell wrapper is introduced.
- No `TableModel.hyperlink`, `AddTableOptions.hyperlink`, table-level default, or bulk hyperlink editor.
- Only exact direct plain single-run/single-paragraph cells are editable.
- URL and slide targets remain mutually exclusive; tooltip omission and explicit empty remain distinct.
- Validate internal targets and current editable state before relationship/XML mutation.
- Equal values and absent clear are exact part/relationship/journal no-ops.
- Same target keeps the relationship ID; unique target changes update in place; shared target changes clone on write.
- Clear removes only the direct click and collects only an unreferenced click relationship.
- Link insertion adds `u="sng"` only when no direct underline exists; edits and clear preserve existing run formatting.
- Rich/multi-run/multi-paragraph cells, hover/action-only/custom-show/macro/sound/history editing, and other hyperlink owners stay out of scope.
- Every task ends with review, a scoped commit, push, fetch, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`, generated workspace tarballs, or retained `/tmp` evidence.

---

### Task 1: Core editable state, relationship lifecycle, and public API

**Files:**
- Modify: `packages/model/src/shape-hyperlink.internal.ts`
- Modify: `packages/model/src/shape-hyperlink.internal.test.ts`
- Modify: `packages/model/src/table-cell-hyperlink.internal.ts`
- Modify: `packages/model/src/table-cell-hyperlink.internal.test.ts`
- Modify: `packages/model/src/table-physical-cells.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: `Hyperlink`, `NormalizedHyperlink`, `normalizeHyperlink()`, `readTextRunHyperlinkBinding()`, `renderShapeHyperlink()`, `shapeHyperlinksEqual()`, `relationshipReferenceCount()`, `HYPERLINK_RELATIONSHIP_TYPE`, `SLIDE_RELATIONSHIP_TYPE`, `relativeRelationshipTarget()`, and OPC transactions.
- Produces: `replaceTextRunHyperlinkElement()`, `EditableTableCellHyperlinkState`, `requireEditableTableCellHyperlinkState()`, `readDirectTablePhysicalCellMatrix()`, `SlideModel.setTableCellHyperlink()`, and `TableModel.setCellHyperlink()`.

- [ ] **Step 1: Add failing local text-run replacement tests**

Extend `shape-hyperlink.internal.test.ts` to import the not-yet-existing helper:

```ts
import { replaceTextRunHyperlinkElement } from './shape-hyperlink.internal.js';
```

Use one namespace-complete `a:rPr` root and require:

- inserting URL and internal-slide clicks with local `xmlns:r` when absent;
- adding `u="sng"` only when direct `u` is absent;
- preserving explicit `u="none"`, style children, PptxGenJS `history`, `tgtFrame`, highlight, sound, and extension content;
- tooltip absent/empty/value replacement and URL/internal action canonicalization;
- relationship-ID replacement;
- click clear while retaining underline and unrelated properties;
- absent clear and equal replacement returning `false` with byte-identical XML;
- `ModelParseError` for wrong root, repeated/wrong-namespace/invalid direct click, repeated underline, or unsafe action/relationship attributes;
- `TypeError` when hyperlink and relationship ID presence do not match.

- [ ] **Step 2: Add failing editable cell-state tests**

Extend `table-cell-hyperlink.internal.test.ts` with:

```ts
import { requireEditableTableCellHyperlinkState } from './table-cell-hyperlink.internal.js';
```

Require a supported unlinked cell to return its direct `rPr` with no binding, and a supported URL/internal cell to return:

```ts
{
  properties: expect.objectContaining({ localName: 'rPr' }),
  hyperlink: { url: 'https://example.com' },
  relationshipId: 'rId7',
}
```

Require `ModelParseError` for every structural case already rejected by `readTableCellHyperlink()`, plus present-but-invalid clicks, dangling/duplicate relationships, wrong type/mode/action, ambiguous slide targets, and wrong namespaces. An unlinked valid cell must remain distinguishable from an unsafe linked cell.

- [ ] **Step 3: Add failing model lifecycle tests**

Extend `model.test.ts` with one created table containing URL, internal, self, shared-import, unlinked, and styled cells. Require:

```ts
table.setCellHyperlink(0, 0, { url: 'https://edited.example', tooltip: '' });
table.setCellHyperlink(0, 1, { slide: 3 });
table.setCellHyperlink(0, 4, { url: 'https://added.example' });
table.setCellHyperlink(0, 2, undefined);
```

Cover:

- immediate snapshots and target ordinal changes after reorder;
- exact no-op bytes, relationships, journal, and stable table identity;
- tooltip-only update retaining the ID;
- unique target update retaining the ID;
- shared relationship target update allocating a new ID and leaving the peer unchanged;
- clear of unique and shared relationships with reference-aware GC;
- new link insertion with default underline and explicit underline preservation;
- clear preserving direct underline and all unrelated table/cell XML;
- `setCellText()` and every existing cell-property editor preserving the edited link;
- duplicate self-link retargeting, target deletion cleanup, rollback, write/reopen, and all six presentation formats;
- physical row/cell range errors;
- malformed current state, invalid input, dangling internal target, and injected relationship/add/update/remove/slide-write failures with zero package mutation.

- [ ] **Step 4: Add failing SDK, root-package, and PptxGenJS tests**

In `packages/sdk/src/index.test.ts`, require the public method, its runtime behavior, write/reopen state, and TypeScript rejection of invalid values:

```ts
table.setCellHyperlink(0, 0, { url: 'https://sdk-edit.example' });
table.setCellHyperlink(0, 1, { slide: 2, tooltip: '' });
table.setCellHyperlink(0, 2, undefined);
```

In `packages/pptx/src/index.test.ts`, require the same method through `@jiayunxie/pptx` exports and a `TableModel` typed value.

In `packages/pptxgenjs-adapter/src/index.test.ts`, generate legal PptxGenJS URL/slide cells, import them, edit URL/slide/tooltip/clear through native, and require:

- final snapshot and relationship semantics;
- retained unrelated `history`, `tgtFrame`, highlight, sound/extension-compatible content;
- owned external empty action canonicalization;
- no new caller `_rId` mutation by native;
- write/reopen preservation.

- [ ] **Step 5: Run focused tests and verify failure**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
```

Expected: fail because the replacement helper, editable state, and public editor do not exist.

- [ ] **Step 6: Implement the local text-run replacer**

In `shape-hyperlink.internal.ts`, export:

```ts
export function replaceTextRunHyperlinkElement(
  xml: LosslessXmlDocument,
  properties: XmlElement,
  hyperlink: NormalizedHyperlink | undefined,
  relationshipId: string | undefined,
  partUri: string,
): boolean;
```

Require `rPr` in the DrawingML namespace. Inspect direct children whose local name is `hlinkClick`; zero is safe, exactly one must pass the existing text-run click inspector, and every other case throws `ModelParseError`. Reject repeated direct `u` attributes.

For insertion, render with the effective drawing/relationship prefixes, declare the relationship namespace locally when needed, insert before direct `hlinkMouseOver` or `extLst`, and add `u="sng"` only when absent. For update, reuse `patchExistingHyperlink()` so relationship ID, tooltip, and action are locally canonicalized without replacing unrelated attributes/children. For clear, remove only the click. Return `false` for exact no-op or absent clear.

- [ ] **Step 7: Implement the strict editable cell boundary**

In `table-cell-hyperlink.internal.ts`, factor the existing direct structure walk into a private helper and export:

```ts
export interface EditableTableCellHyperlinkState {
  readonly properties: XmlElement;
  readonly hyperlink?: NormalizedHyperlink;
  readonly relationshipId?: string;
}

export function requireEditableTableCellHyperlinkState(
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
  partUri: string,
): EditableTableCellHyperlinkState;
```

Return `{ properties }` when no direct click exists. When a direct click-like child exists, require exactly one valid `readTextRunHyperlinkBinding()` result and return its hyperlink/relationship ID. Convert every unsafe structure or present-but-undecodable click into `ModelParseError('Table cell hyperlink state is not safely editable', partUri)`.

- [ ] **Step 8: Add exact physical-cell matrix resolution**

In `table-physical-cells.internal.ts`, export:

```ts
export function readDirectTablePhysicalCellMatrix(
  frame: XmlElement,
): readonly (readonly XmlElement[])[] | undefined;
```

Keep the existing exact direct `graphicFrame/graphic/graphicData/tbl/tr/tc` path and non-empty row/cell requirements. Refactor `readDirectTablePhysicalCells()` to flatten the matrix so existing bulk editors retain identical behavior.

- [ ] **Step 9: Implement transactional relationship editing**

In `slide.ts`, add:

```ts
setTableCellHyperlink(
  id: number,
  rowIndex: number,
  columnIndex: number,
  value: Hyperlink | undefined,
): void;
```

Normalize before the transaction and resolve internal slide target before mutation. Inside the transaction:

1. Resolve the table shape and exact physical matrix; throw `ModelParseError` for an unsafe table and `RangeError` for a missing physical cell.
2. Read `requireEditableTableCellHyperlinkState()` and return on `shapeHyperlinksEqual()`.
3. On clear, remove the click, write the slide only when changed, and remove the old relationship only when the updated XML reference count is zero.
4. On same target, patch the click using the current ID without relationship mutation.
5. On target change, build a strict external hyperlink or internal slide relationship input. Update the old relationship in place only when its reference count is one; otherwise add a new relationship.
6. Patch the selected click, write the slide, then collect the replaced relationship only if unreferenced.

Use the existing `shapeHyperlinkTargetsEqual()` target comparator and the rollback-safe OPC transaction.

- [ ] **Step 10: Expose the physical-cell editor**

In `TableModel` add:

```ts
setCellHyperlink(
  rowIndex: number,
  columnIndex: number,
  value: Hyperlink | undefined,
): void {
  this.slide.setTableCellHyperlink(this.id, rowIndex, columnIndex, value);
}
```

Do not change `TableCell`, `TableRow`, or creation input types.

- [ ] **Step 11: Run focused, type, full, and performance gates**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
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

- [ ] **Step 12: Review, commit, push, and verify**

Review API naming, zero-based physical indexes, strict editable boundary, exact no-op, tooltip absence/empty, target identity, ID reuse, shared COW, reference-aware GC, underline behavior, PptxGenJS preservation/canonicalization, rollback, and no public table default.

Stage only Task 1 files, commit, and push:

```sh
git diff --check
git add \
  packages/model/src/shape-hyperlink.internal.ts \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/table-cell-hyperlink.internal.ts \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-physical-cells.internal.ts \
  packages/model/src/shapes.ts \
  packages/model/src/slide.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit table-cell hyperlinks"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0` before Task 2.

---

### Task 2: Actual npm package, CLI, inspector, and browser proof

**Files:**
- Modify: `scripts/build-npm-package-types.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: packed root/SDK/model declarations and `TableModel.setCellHyperlink()` from Task 1.
- Produces: installed Node/types/browser/CLI, `pptx-inspect`, and real-Chrome evidence plus retained `/tmp` artifacts.

- [ ] **Step 1: Extend packed declaration checks**

Require the generated `TableModel` declaration to contain:

```ts
setCellHyperlink(
  rowIndex: number,
  columnIndex: number,
  value: Hyperlink | undefined,
): void;
```

Keep the existing `AddTableCellOptions.hyperlink` and readonly `TableCell.hyperlink` checks. Extend the packed TypeScript consumer with legal URL/slide/clear calls and `@ts-expect-error` cases for invalid targets, tooltip, and a table-level hyperlink property.

- [ ] **Step 2: Extend installed Node and browser-conditional smoke**

Create URL, internal, self, and plain cells from the installed tarball. Exercise no-op, tooltip-only update, URL/slide switch, add, clear, reorder, text/style coexistence, duplicate self-link, target deletion, invalid failure isolation, and write/reopen. Require one relationship per native cell, ID reuse for unique edits, COW for a deliberately shared imported ID, and reference-aware GC.

Expose `tableCellHyperlinkEditing: true` plus a stable JSON state from both Node and browser-conditional runs.

- [ ] **Step 3: Extend real Chrome coverage**

In `playwright-browser-smoke.js`, perform the same public editing states from the browser bundle. Require:

- `tableCellHyperlinkEditing: true`;
- immediate/reopen snapshots matching Node;
- exact no-op and invalid-failure package isolation;
- zero validation, console, page, and network errors.

- [ ] **Step 4: Pack and run actual-package gates**

Run the repository's existing pack/smoke commands from a new `mktemp -d` evidence directory. Retain the tarball, SHA-256, pack listing, installed declarations, Node/browser JSON, generated PPTX, CLI JSON, inspector JSON, Chrome JSON, and final evidence summary outside the repository.

- [ ] **Step 5: Inspect the generated PPTX from wide to narrow**

Use `pptx-inspect` to record package summary, validation, slide list, exact slide XML, and exact relationship part. Require final physical cells to contain the expected direct click/action/tooltip states, independent relationship IDs, preserved text/style, and no stale cleared relationship.

Validate against the PowerPoint 2010 profile. External hyperlink portability warnings are expected; package errors are not.

- [ ] **Step 6: Run final clean gates**

Run:

```sh
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 7: Review, commit, push, and verify**

Review actual tarball file count/hash, declarations, installed runtime, browser conditional export, CLI, `pptx-inspect`, exact relationship ownership, PowerPoint diagnostics, Chrome error counts, and evidence path.

Stage only the three scripts:

```sh
git add \
  scripts/build-npm-package-types.mjs \
  scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed table-cell hyperlink editing"
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
- Consumes: final Task 1 behavior and Task 2 measured evidence.
- Produces: public editing examples, exact ownership/clear semantics, PptxGenJS boundary, final evidence, and rich/multi-paragraph cell text as the next item.

- [ ] **Step 1: Update the three public README/API surfaces**

Extend the compact table-cell hyperlink example with URL/slide replacement, tooltip omission/empty, add, clear, no-op, target reorder, `setCellText()` preservation, and write/reopen. State:

- zero-based physical indexes;
- one independent relationship per native cell;
- exact no-op, unique ID reuse, shared COW, and reference-aware GC;
- link insertion underline default and clear style preservation;
- strict plain single-run boundary;
- no table-level default or rich-cell support.

- [ ] **Step 2: Update compatibility and progress records**

Move table-cell hyperlink editing/clearing to supported. Record that PptxGenJS 4.0.1 has no existing-deck editor, but its legal created state can be imported and edited natively. Include measured focused/full/performance totals, tarball file count/SHA, installed Node/types/browser/CLI, `pptx-inspect`, Chrome, PowerPoint warnings/errors, evidence path, and commit hashes.

- [ ] **Step 3: Update changelog and remaining work**

Add core lifecycle, PptxGenJS boundary, and actual-package/browser proof bullets. Keep overall parity approximately 97%. Select rich/multi-paragraph cell text/style as the next independently testable item, followed by merge, row/column CRUD, auto-page/repeated headers, `tableToSlides`, and final peer/client audit.

- [ ] **Step 4: Run final gates**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/shape-hyperlink.internal.test.ts \
  packages/model/src/table-cell-hyperlink.internal.test.ts \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 5: Review, commit, push, verify, and report**

Stage only the six documentation files:

```sh
git add \
  README.md \
  packages/pptx/README.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  docs/implementation-progress.md \
  CHANGELOG.md
git diff --cached --check
git commit -m "docs: document table-cell hyperlink editing"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, report completed/remaining items and overall progress, then begin the rich/multi-paragraph cell text design item.
