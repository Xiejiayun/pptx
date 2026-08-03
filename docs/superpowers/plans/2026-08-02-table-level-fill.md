# Table-Level Fill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict live `TableModel.fill` consensus read and atomic all-physical-cell replacement/clear, prove supported PptxGenJS final-state conformance, and ship verified packed Node/browser/types/CLI behavior.

**Architecture:** Extend the existing table-cell fill codec with table aggregation over `readDirectTablePhysicalCells()` and bulk mutation through `replaceTableCellFill()`. `TableModel` normalizes one `TableCellFill` before an OPC transaction, mutates one in-memory slide XML snapshot, and commits at most once; no table-default metadata or effective-style resolution is introduced.

**Tech Stack:** TypeScript 5.8, Vitest 3, lossless OOXML/XML helpers, OPC transactions, PptxGenJS 4.0.1 fixtures, tsup, npm tarballs, installed Node/browser consumers, `pptx-inspect`, Playwright, and real Google Chrome.

## Global Constraints

- Native input remains strict `TableCellFill | undefined`: direct none or solid sRGB/theme fill with optional finite `0..100` transparency.
- Getter exposes only one uniform supported direct physical-cell value; absent, mixed, malformed, advanced, or unsafe state is `undefined`.
- Setter whole-replaces or clears every physical cell, including merge continuations, in one package transaction and one slide commit.
- Omitted alpha and explicit zero transparency remain distinct; direct none remains distinct from absence.
- Existing `TableCell.fill`, `setCellFill()`, creation precedence/output, and unsupported advanced fill preservation semantics do not change.
- No table-style/effective fill, mixed sentinel, creation metadata, table-level border/fit, merge CRUD, auto-page, content measurement, or layout recomputation.
- Each implementation, package-proof, and documentation item is reviewed, committed, pushed, fetched, and verified at divergence `0 0` before the next item.
- Stage only files declared by the current task; never stage `.pnpm-store/`, generated tarballs, installed smoke workspaces, or `/tmp` evidence.

---

### Task 1: Core table-level fill API, lifecycle, and PptxGenJS conformance

**Files:**
- Create: `packages/model/src/table-cell-fill.internal.test.ts`
- Modify: `packages/model/src/table-cell-fill.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: `TableCellFill`, `normalizeTableCellFill()`, `readTableCellFill()`, `replaceTableCellFill()`, `simpleFillsEqual()`, and `readDirectTablePhysicalCells()`.
- Produces: `readTableFill()`, `replaceTableFill()`, and live `TableModel.fill: TableCellFill | undefined` getter/setter behavior used by packed consumers and documentation.

- [ ] **Step 1: Record the clean baseline**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
```

Expected: the existing table creation/cell fill, table-level alignment/direction/margins, SDK, and adapter suites pass before source edits.

- [ ] **Step 2: Write focused failing helper tests**

Create `packages/model/src/table-cell-fill.internal.test.ts` with a direct-table fixture whose cells preserve text, body properties, paragraph alignment, margins, border, direction, merge markers, transform, grid, row geometry, extensions, and relationships. The first tests must require uniform supported consensus and no read mutation:

```ts
const solid = {
  kind: 'solid' as const,
  color: { kind: 'scheme' as const, value: 'accent1' as const },
  transparency: 25,
};

expect(readTableFill(uniform.xml, uniform.frame)).toEqual(solid);
expect(uniform.xml.changed).toBe(false);
expect(readTableFill(none.xml, none.frame)).toEqual({ kind: 'none' });
expect(readTableFill(explicitZero.xml, explicitZero.frame)).toEqual({
  kind: 'solid',
  color: { kind: 'srgb', value: '00FF00' },
  transparency: 0,
});
```

Add explicit cases for:

```ts
expect(readTableFill(allAbsent.xml, allAbsent.frame)).toBeUndefined();
expect(readTableFill(presentAbsent.xml, presentAbsent.frame)).toBeUndefined();
expect(readTableFill(noneSolid.xml, noneSolid.frame)).toBeUndefined();
expect(readTableFill(differentColor.xml, differentColor.frame)).toBeUndefined();
expect(readTableFill(differentAlpha.xml, differentAlpha.frame)).toBeUndefined();
expect(readTableFill(omittedVsZero.xml, omittedVsZero.frame)).toBeUndefined();
expect(readTableFill(malformed.xml, malformed.frame)).toBeUndefined();
expect(readTableFill(gradient.xml, gradient.frame)).toBeUndefined();
expect(readTableFill(repeatedChoice.xml, repeatedChoice.frame)).toBeUndefined();
expect(readTableFill(missingProperties.xml, missingProperties.frame)).toBeUndefined();
expect(readTableFill(repeatedProperties.xml, repeatedProperties.frame)).toBeUndefined();
```

Include empty rows/table, ambiguous `graphic → graphicData → tbl` paths, `gridSpan`/`hMerge`/`vMerge` continuation cells, same local-name wrong-prefix children, border/text descendant fill lookalikes, and canonical uppercase sRGB equality.

- [ ] **Step 3: Run helper tests and verify RED**

Run:

```sh
node_modules/.bin/vitest run packages/model/src/table-cell-fill.internal.test.ts --reporter=verbose
```

Expected: compile/test failure because `readTableFill()` and `replaceTableFill()` do not exist.

- [ ] **Step 4: Implement strict table consensus and bulk replacement**

In `packages/model/src/table-cell-fill.internal.ts`, import `readDirectTablePhysicalCells` and add:

```ts
export function readTableFill(
  xml: LosslessXmlDocument,
  frame: XmlElement,
): TableCellFill | undefined {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) return undefined;
  const first = readTableCellFill(xml, cells[0]!);
  if (!first) return undefined;
  return cells.slice(1).every((cell) =>
    simpleFillsEqual(readTableCellFill(xml, cell), first)
  ) ? first : undefined;
}

export function replaceTableFill(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  fill: TableCellFill | undefined,
  partUri: string,
): boolean {
  const cells = readDirectTablePhysicalCells(frame);
  if (!cells) {
    throw new ModelParseError(
      'Table must contain one complete set of direct physical cells',
      partUri,
    );
  }
  let changed = false;
  for (const cell of cells) {
    changed = replaceTableCellFill(xml, cell, fill, partUri) || changed;
  }
  return changed;
}
```

Do not change `readTableCellFill()`, `replaceTableCellFill()`, `renderTableCellFill()`, `normalizeTableCellFill()`, or creation behavior.

- [ ] **Step 5: Complete bulk-helper GREEN cases**

Require exact behavior for all four physical cells:

```ts
expect(replaceTableFill(target.xml, target.frame, solid, PART_URI)).toBe(true);
expect(readTableFill(target.xml, target.frame)).toEqual(solid);
expect(replaceTableFill(target.xml, target.frame, solid, PART_URI)).toBe(false);
expect(replaceTableFill(target.xml, target.frame, { kind: 'none' }, PART_URI)).toBe(true);
expect(readTableFill(target.xml, target.frame)).toEqual({ kind: 'none' });
expect(replaceTableFill(target.xml, target.frame, undefined, PART_URI)).toBe(true);
expect(readTableFill(target.xml, target.frame)).toBeUndefined();
expect(replaceTableFill(target.xml, target.frame, undefined, PART_URI)).toBe(false);
```

Assert exact solid/noFill counts, alpha tokens, insertion before `extLst`, and byte preservation for every unrelated fixture token. Replace and clear one malformed/gradient choice successfully. Put repeated choices only in the final physical cell and require `ModelParseError` after earlier in-memory edits; then prove the caller's package bytes remain unchanged in the model test.

Run:

```sh
node_modules/.bin/vitest run packages/model/src/table-cell-fill.internal.test.ts --reporter=dot
```

Expected: all focused helper tests pass.

- [ ] **Step 6: Add the live `TableModel.fill` property**

In `packages/model/src/shapes.ts`, import both table helpers and add after `horizontalAlignment`:

```ts
get fill(): TableCellFill | undefined {
  const { xml, element } = this.resolve();
  return readTableFill(xml, element);
}

set fill(value: TableCellFill | undefined) {
  const fill = normalizeTableCellFill(value, 'Table fill');
  this.slide.presentation.opcPackage.transaction(() => {
    const { xml, element } = this.resolve();
    if (replaceTableFill(xml, element, fill, this.slide.partUri)) {
      this.slide.setXml(xml.serialize());
    }
  });
}
```

The normalizer must run before `resolve()` or package mutation.

- [ ] **Step 7: Add model and SDK lifecycle coverage**

Add `projects and atomically edits uniform table fill` to `packages/model/src/model.test.ts` and `projects and edits table-level fill through the public root API` to `packages/sdk/src/index.test.ts`. Use a two-by-two table with uniform scheme fill and unrelated border/margin/alignment/direction/fit/geometry state. Require:

```ts
expect(table.fill).toEqual({
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
});
const detached = table.fill!;
(detached as { transparency?: number }).transparency = 99;
expect(table.fill).toEqual({
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
});
table.setCellFill(0, 1, { kind: 'none' });
expect(table.fill).toBeUndefined();
table.fill = { kind: 'none' };
expect(table.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
  .toEqual(Array(4).fill({ kind: 'none' }));
table.fill = {
  kind: 'solid',
  color: { kind: 'srgb', value: 'D9EAF7' },
  transparency: 0,
};
expect(table.fill).toEqual({
  kind: 'solid',
  color: { kind: 'srgb', value: 'D9EAF7' },
  transparency: 0,
});
table.fill = undefined;
expect(table.fill).toBeUndefined();
```

Capture slide bytes/journal around same-value and all-absent clear no-ops. Require one real mutation commit, caller detachment, stable identity, duplicate isolation, move, user-transaction rollback, write/reopen, all six presentation formats, zero diagnostics, and preservation of text/borders/margins/alignment/direction/fit/grid/rows/transform/relationships. Inject repeated direct fill choices into the final cell, require read `undefined`, assignment `ModelParseError`, and zero package partial mutation. Prove legal assignment repairs one malformed or advanced direct fill.

Add compile-only checks:

```ts
const fill: TableCellFill | undefined = table.fill;
table.fill = { kind: 'none' };
table.fill = { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } };
table.fill = undefined;
// @ts-expect-error table fill rejects null
table.fill = null;
// @ts-expect-error table fill rejects PptxGenJS-shaped input
table.fill = { color: 'D9EAF7' };
void fill;
```

- [ ] **Step 8: Extend real PptxGenJS 4.0.1 conformance**

In `packages/pptxgenjs-adapter/src/index.test.ts`, add `projects and normalizes PptxGenJS table-level fill output` next to the existing table-level margins projection test. Generate three tables:

```ts
slide.addTable([[{ text: 'Uniform A', options: {} }, { text: 'Uniform B', options: {} }]], {
  fill: { color: generated.SchemeColor.accent1, transparency: 25 },
});
slide.addTable([[{ text: 'Omitted A', options: {} }, { text: 'Omitted B', options: {} }]], {});
slide.addTable([[
  { text: 'Inherited', options: {} },
  { text: 'Override', options: { fill: { color: 'FFFF00', transparency: 50 } } },
]], { fill: { color: '0000FF' } });
```

Require uniform scheme consensus, omitted `undefined`, mixed `undefined`, exact per-cell snapshots, then set the mixed imported table to native `{ kind: 'none' }`, write/reopen, and require uniform direct noFill. Keep the existing creation test's PptxGenJS none-collapse and explicit-zero-alpha differences unchanged.

- [ ] **Step 9: Run focused, type, bundle, full, and performance gates**

Run:

```sh
node_modules/.bin/vitest run \
  packages/model/src/table-cell-fill.internal.test.ts \
  packages/model/src/table-cell-margins.internal.test.ts \
  packages/model/src/table-cell-horizontal-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
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

Record exact focused/full totals and performance time. Inspect `packages/pptx/dist/types/model/shapes.d.ts` and require both fill accessors inside `TableModel`.

- [ ] **Step 10: Review, commit, push, and verify core behavior**

Review strict getter safety, none/solid/alpha equality, absence versus explicit direct intent, advanced-choice replacement, no-ops, late-cell atomicity, input errors, unrelated-state preservation, duplicate/rollback/reopen, public types, and real PptxGenJS output. Stage only the six declared files and run:

```sh
git add \
  packages/model/src/table-cell-fill.internal.test.ts \
  packages/model/src/table-cell-fill.internal.ts \
  packages/model/src/shapes.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "feat: edit table-level fill"
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
- Consumes: Task 1 `TableModel.fill`, packed `model/shapes.d.ts`, root/browser exports, installed CLI, and the checked-in Chrome callback.
- Produces: stable `tableFill: true`, detailed `tableFillState`, and `tableFillInspect: true` evidence.

- [ ] **Step 1: Lock the packed `TableModel` declaration block**

In both declaration checks, slice from `export declare class TableModel` to `export declare class ChartModel` and require:

```ts
get fill(): TableCellFill | undefined;
set fill(value: TableCellFill | undefined);
```

Throw a table-level-fill-specific error if either accessor is absent. The check must not pass from `TableCell.fill` or `ShapeModel.fill`.

- [ ] **Step 2: Add installed Node runtime evidence**

Create a dedicated two-by-two table and canonicalize snapshots to `null`, `{ kind: 'none' }`, or ordered solid objects. Record this exact state:

```ts
const tableFillState = {
  uniform: {
    kind: 'solid',
    color: { kind: 'scheme', value: 'accent1' },
    transparency: 25,
  },
  readIsolation: true,
  noOp: true,
  mixed: null,
  none: { kind: 'none' },
  noneCells: Array(4).fill({ kind: 'none' }),
  solid: {
    kind: 'solid',
    color: { kind: 'srgb', value: 'D9EAF7' },
    transparency: 0,
  },
  solidCells: Array(4).fill({
    kind: 'solid',
    color: { kind: 'srgb', value: 'D9EAF7' },
    transparency: 0,
  }),
  cleared: null,
  clearedCells: [null, null, null, null],
  reopened: { kind: 'none' },
  reopenedCells: Array(4).fill({ kind: 'none' }),
  invalidError: {
    name: 'TypeError',
    message: 'Table fill must be an object',
  },
  failureIsolation: true,
  validationErrors: 0,
};
```

Mutate the first detached solid snapshot's color value and preserve bytes/journal for read isolation. Assign the equivalent initial solid for no-op, make one cell none for mixed state, broadcast none, broadcast explicit-zero solid, clear, reject `null`, set final none, write `table-fill-smoke.pptx`, and reopen. Add boolean/state to `apiChecks` and top-level JSON.

Before the smoke workspace is removed, copy `table-fill-smoke.pptx` beside the input tarball as retained evidence so the host `pptx-inspect` workflow can inspect the actual packed-consumer output.

- [ ] **Step 3: Add installed TypeScript and browser-condition evidence**

In the generated TypeScript consumer:

```ts
const typedTableFill: TableCellFill | undefined = typedTable.fill;
typedTable.fill = { kind: 'none' };
typedTable.fill = {
  kind: 'solid',
  color: { kind: 'scheme', value: 'accent1' },
  transparency: 25,
};
typedTable.fill = undefined;
// @ts-expect-error table fill rejects null
typedTable.fill = null;
// @ts-expect-error table fill rejects PptxGenJS-shaped input
typedTable.fill = { color: 'D9EAF7' };
void typedTableFill;
```

In the generated browser consumer, reproduce the exact Node state with `writeBlob()` and require identical values, invalid error, failure isolation, and zero diagnostics.

- [ ] **Step 4: Add installed CLI and `pptx-inspect` evidence**

Use the installed CLI to validate, list slides, and read the final slide part from `table-fill-smoke.pptx`. Require one table shape and exactly four direct physical-cell `a:noFill` choices, with no direct cell `solidFill`, `gradFill`, `blipFill`, `pattFill`, or `grpFill` remaining. Reject border/text descendant false positives and add `tableFillInspect: true`.

Then run the required broad-to-narrow workflow against the same output:

```sh
pptx-inspect --json package inspect table-fill-smoke.pptx
pptx-inspect --json package validate table-fill-smoke.pptx --profile powerpoint-2010
pptx-inspect --json slides list table-fill-smoke.pptx
pptx-inspect --json part read table-fill-smoke.pptx /ppt/slides/slide1.xml
```

Require zero validation errors before accepting the exact slide-part read.

- [ ] **Step 5: Extend the real-Chrome callback**

Add the same `tableFillState` and expected JSON to `scripts/playwright-browser-smoke.js`. Require detached/read/no-op isolation, mixed projection, none/solid overwrite, clear, none reopen, invalid failure isolation, zero document diagnostics, and global `errorCounts: { console: 0, page: 0, network: 0 }`.

- [ ] **Step 6: Build, pack, install, and run all proof gates**

Run:

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
(cd packages/pptx && ../../node_modules/.bin/tsup)
(cd packages/pptx && ../../node_modules/.bin/tsup --config tsup.browser.config.ts)
node scripts/build-npm-package-types.mjs
table_fill_artifacts=$(mktemp -d /tmp/pptx-table-fill-artifacts.XXXXXX)
(cd packages/pptx && npm pack --ignore-scripts --pack-destination "$table_fill_artifacts")
node scripts/smoke-npm-package.mjs "$table_fill_artifacts/jiayunxie-pptx-0.1.0.tgz"
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

Extract the actual tarball beneath the same fresh evidence directory, serve its browser module over loopback, and run the checked-in callback in installed Google Chrome through the bundled Playwright runtime. Retain tarball file count, SHA-256, installed Node/types/browser/CLI output, `pptx-inspect` JSON, full and compact Chrome state, and evidence directory outside the repository.

- [ ] **Step 7: Review, commit, push, and verify package proof**

Review declaration scoping, stable state/key order, detached getter proof, browser/Node parity, exact direct-cell ownership, prior-field preservation, zero browser errors, and absence of generated repository artifacts. Stage only the three scripts, then:

```sh
git add \
  scripts/build-npm-package-types.mjs \
  scripts/smoke-npm-package.mjs \
  scripts/playwright-browser-smoke.js
git diff --cached --check
git commit -m "test: verify packed table-level fill"
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
- Produces: one consistent public support statement, PptxGenJS direct-state boundary, final evidence record, revised remaining-work list, and table-level border as the next advanced-table item.

- [ ] **Step 1: Add public examples and direct-state semantics**

In all three public README/API surfaces, add a compact example with uniform solid read, one-cell mixed state, none broadcast, explicit-zero solid broadcast, and `undefined` clear. State that the getter returns only one uniform supported safe direct value, never resolves table styles/defaults, and that mixed detail remains in `rows[].cells[].fill`.

- [ ] **Step 2: Update compatibility and progress records**

Move table-level fill consensus/bulk editing to supported in the compatibility matrix and detailed table section. Record PptxGenJS legal uniform/mixed/omitted final states, native existing-deck extension, none/explicit-zero differences, focused/full/performance totals, actual tarball count/SHA, installed Node/types/browser/CLI, `pptx-inspect`, real Chrome, zero error counts, evidence path, and all implementation commit hashes.

- [ ] **Step 3: Update changelog and remaining-work statements**

Add three changelog bullets for core semantics, PptxGenJS direct-state boundary, and actual-package/browser proof. Remove every stale statement that table-level fill getter/editor is pending. Select table-level border as the next item while keeping overall parity approximately 97% until the final peer/client audit.

- [ ] **Step 4: Run final documentation and regression gates**

Run:

```sh
git diff --check
node_modules/.bin/vitest run \
  packages/model/src/table-cell-fill.internal.test.ts \
  packages/model/src/table-cell-margins.internal.test.ts \
  packages/model/src/table-cell-horizontal-alignment.internal.test.ts \
  packages/model/src/table-cell-vertical-alignment.internal.test.ts \
  packages/model/src/table-cell-text-direction.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  --reporter=dot
node_modules/.bin/vitest run --reporter=dot --maxWorkers=2
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
```

- [ ] **Step 5: Review, commit, push, verify, and report**

Review examples, exact getter/setter type, direct none versus absence, omitted versus explicit-zero alpha, PptxGenJS boundary, measured totals/hashes/evidence, remaining-work consistency, and `git diff --check`. Stage only the six documentation files, then:

```sh
git add \
  README.md \
  packages/pptx/README.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  docs/implementation-progress.md \
  CHANGELOG.md
git diff --cached --check
git commit -m "docs: document table-level fill"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`, report completed/remaining items and overall progress, then begin the table-level border design item.
