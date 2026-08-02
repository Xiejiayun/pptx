# Rich Text `breakLine` Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict `RichTextRun.breakLine?: boolean` paragraph splitting across every native rich-text creation/editing owner, preserve canonical paragraph and relationship semantics, and prove PptxGenJS 4.0.1 parity through source, packed runtime, browser, CLI, and real PPTX compatibility gates.

**Architecture:** Treat `breakLine` as transient input syntax at the shared `normalizeRichText()` boundary. Validate every input first, split each input paragraph after flagged non-final runs, copy its normalized paragraph properties to every segment, and remove the marker before downstream rendering. Existing renderers, relationship matrices, getters, setters, owner wrappers, and six-format writers consume only canonical explicit paragraphs.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, actual npm tarball smoke, real Chrome, `pptx-inspect` PowerPoint 2010 validation, LibreOffice, and PDF/PNG visual QA.

## Global Constraints

- `RichTextRun.breakLine` is exactly `boolean | undefined`; only primitive true splits.
- All input paragraphs/runs/styles/properties validate before returning a normalized result or mutating any package state.
- A flagged non-final run ends the current paragraph. A flagged final run is consumed without producing a trailing empty paragraph.
- Empty flagged runs and consecutive flags preserve corresponding empty canonical paragraphs.
- Every split segment copies the original input paragraph's normalized align/rtl/margins/indent/bullet/level/spacing/tabStops properties.
- The normalized run type contains text/style/`softBreakBefore` only; transient `breakLine` never reaches rendering, relationship allocation, equality, or getter snapshots.
- `softBreakBefore` remains attached to its run even when that run becomes the first run of a split paragraph; do not weaken the existing reversible native soft-break contract.
- Getter snapshots expose explicit `RichTextParagraph[]` only and never guess `breakLine` markers from OOXML.
- Run hyperlink relationship matrices are allocated after splitting and must match canonical paragraph/run indexes without leaks or target drift.
- Plain slide, layout, master, placeholders, declarative masters, setters, duplicates, moves, rollback, and all six formats share the same normalizer.
- Rich-text table cells are not yet a public native API and remain in the advanced-table scope; do not widen table creation/editing in this slice.
- Do not add `breakLine` to outer `AddTextOptions`, accept CR/LF inside run text, modify the PptxGenJS adapter production path, or claim remaining full parity.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, browser artifacts, LibreOffice profiles, render output, or build output.
- Use repository-local binaries. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Strict canonical paragraph splitting

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Modify: `packages/model/src/rich-text.internal.test.ts`

**Interfaces:**
- Consumes: public `RichTextRun`, `normalizeRichText()`, current strict ordinary-object/key validation, and `NormalizedRichTextParagraph`.
- Produces: public `breakLine?: boolean`, a normalized run shape without the transient marker, and deterministic paragraph flattening.

- [ ] **Step 1: Add failing split-matrix unit tests**

Require exact normalized output for:

- omitted/undefined/false/true;
- middle and multiple non-final flags;
- final and single-run flags without trailing paragraph;
- empty flagged runs, consecutive empty flags, a following empty unflagged run, and original `runs: []`;
- multiple input paragraphs, proving splits never reorder or merge across owners;
- input immutability, distinct output arrays, and marker removal.

- [ ] **Step 2: Add failing property and soft-break tests**

Build one paragraph containing every direct property plus styles and `softBreakBefore`. Require every split segment to receive the same detached normalized property values and each run to retain its original text/style/soft-break state. Include a soft break on the first run after a split and require it to remain explicit.

- [ ] **Step 3: Add failing strict-validation tests**

Reject string, number, null, object, symbol, boxed boolean, and function values. Reject run accessors without executing them, unknown/symbol keys, prototype objects, and class instances. Verify valid earlier runs do not hide an invalid later run and no input object/array is modified.

- [ ] **Step 4: Implement normalized run separation and splitting**

Add the public field in `text.ts`. In the internal module, use an explicit normalized run interface that omits `breakLine`. Have `normalizeRun()` return the validated canonical run plus a detached split flag, then flatten each normalized paragraph through a focused helper. Keep property normalization once per input paragraph and do not invoke the renderer during splitting.

The core shape should remain equivalent to:

```ts
interface NormalizedRunInput {
  readonly run: NormalizedRichTextRun;
  readonly breakLine: boolean;
}

function splitParagraph(
  properties: Omit<NormalizedRichTextParagraph, 'runs'>,
  runs: readonly NormalizedRunInput[],
): readonly NormalizedRichTextParagraph[] {
  const output: NormalizedRichTextParagraph[] = [];
  let current: NormalizedRichTextRun[] = [];
  for (const [index, value] of runs.entries()) {
    current.push(value.run);
    if (value.breakLine && index + 1 < runs.length) {
      output.push({ ...properties, runs: current });
      current = [];
    }
  }
  output.push({ ...properties, runs: current });
  return output;
}
```

- [ ] **Step 5: Run focused gates**

```sh
node_modules/.bin/vitest run packages/model/src/rich-text.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review final-run guard, empty/consecutive behavior, flatten order, property copying, transient-marker removal, soft-break preservation, input immutability, strict validation, and unchanged renderer output for legacy inputs. Commit `feat: normalize rich text line breaks`, push, fetch, and require divergence `0 0`.

---

### Task 2: Shape creation, live editing, relationships, and lifecycle

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify if defects are exposed: `packages/model/src/slide.ts`
- Modify if defects are exposed: `packages/model/src/rich-text.internal.ts`

**Interfaces:**
- Consumes: Task 1 canonical paragraphs, `addRichText()`, `setShapeRichText()`, relationship preparation/allocation/cleanup, and existing owner transactions.
- Produces: permanent create/edit/lifecycle evidence and any minimal integration correction required by canonical reindexing.

- [ ] **Step 1: Add failing direct create/getter/edit tests**

Create a slide rich-text shape with middle/trailing/empty/consecutive flags and all paragraph properties. Require exact `a:p` count/order/content, immediate canonical getter without `breakLine`, write/reopen equality, and setting the getter snapshot back as a semantic/bytes no-op. Edit an existing shape to add, remove, and move paragraph boundaries.

- [ ] **Step 2: Add failing hyperlink matrix tests**

Combine outer URL/internal-slide hyperlink, per-run URL/internal-slide override, `false` suppression, empty linked runs, and consecutive splits. Require every canonical run to hold the intended target and relationship kind after create/edit/reopen. When boundaries or targets change, require no stale relationship and no accidental target reuse.

- [ ] **Step 3: Add failing ownership and lifecycle tests**

Combine splits with style, paragraph settings, text-body layout, geometry, `rectRadius`, `isTextBox`, fill, line, arrows, shadow, and outer hyperlink. Require only text paragraphs and necessary relationships to change. Cover duplicate independence, move across slides, outer transaction rollback, stable live identity, and all six presentation formats.

- [ ] **Step 4: Add atomic invalid-input tests**

Snapshot part bytes, relationships, part URIs, mutations, shape order, and a live handle. Exercise invalid `breakLine` on creator and setter, including a valid hyperlink before an invalid later run. Require exact zero mutation and accessor non-execution.

- [ ] **Step 5: Apply only exposed integration corrections**

The expected implementation is already centralized in Task 1. If relationship reuse or equality logic assumes pre-split indexes, correct it at the canonical paragraph boundary; do not add a second splitter or renderer branch. Preserve reference-count cleanup and outer transaction rollback.

- [ ] **Step 6: Run model gates**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts -t "breakLine|rich text line break" --reporter=dot
node_modules/.bin/vitest run packages/model/src/rich-text.internal.test.ts packages/model/src/model.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 7: Review, commit, push, and verify**

Review exact OOXML paragraph boundaries, getter canonicalization, same-semantic no-op, hyperlink matrix indexes/cleanup, invalid-input atomicity, six-format persistence, duplicate/move/rollback, and isolation from non-text state. Commit `test: cover rich text line break lifecycle` if evidence alone is sufficient, otherwise `feat: integrate rich text line breaks`; push, fetch, and require divergence `0 0`.

---

### Task 3: All public owners and root declarations

**Files:**
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify if a closed-key defect is exposed: `packages/sdk/src/master-layout.ts`

**Interfaces:**
- Consumes: Task 1 public interface and shared normalizer through slide/layout/master/placeholder/declarative routes.
- Produces: all-owner runtime coverage, declarative cloning evidence, and root declaration/type proof.

- [ ] **Step 1: Add failing owner matrix tests**

Exercise direct slide/layout/master rich text, layout/master placeholder prompts, named placeholder population, and declarative master text/placeholder definitions. Require matching canonical paragraph counts/properties and owner-source isolation after create/write/reopen.

- [ ] **Step 2: Add placeholder lifecycle tests**

Use split placeholder source content and split population content, including empty segments and opposite owner defaults. Require materialization/population to preserve placeholder identity/transform/`isTextBox`, leave layout/master source bytes unchanged, and avoid duplicate or missing paragraphs through duplicate/rollback/reopen.

- [ ] **Step 3: Add declarative detachment and timing tests**

Mutate source arrays/objects after invoking declarative master creation and while asynchronous resources are delayed. Require the stored definition to use detached normalized state. Invalid/accessor `breakLine` must fail before observable package mutation or resource-side effects.

- [ ] **Step 4: Add root type/runtime tests**

Import `RichTextRun`, `RichTextParagraph`, and the owner APIs from the root package. Compile legal boolean usage and add `@ts-expect-error` cases for string, number, null, object, symbol, and outer `AddTextOptions.breakLine`. Verify runtime create/read/edit/write/reopen through root exports.

- [ ] **Step 5: Apply only exposed public-surface corrections**

No closed outer option key is expected because `breakLine` belongs inside run data. If declarative cloning bypasses the shared boundary, route it through the existing canonical helper without duplicating field lists or semantics. Do not add rich-text table-cell creation or editing here; that remains advanced-table work.

- [ ] **Step 6: Run public gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "breakLine|rich text line break" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 7: Review, commit, push, and verify**

Review every public rich-text owner, placeholder source/population isolation, declarative detachment timing, root declaration surface, compile-time negatives, and the absence of an outer `AddTextOptions.breakLine`. Commit `test: cover rich text line break owners`, push, fetch, and require divergence `0 0`.

---

### Task 4: PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()`/`write()`, package import, raw OOXML helpers, and native Tasks 1–3 APIs.
- Produces: permanent paragraph/run/hyperlink parity evidence and explicit strict/native-soft-break divergence records.

- [ ] **Step 1: Add public-only break matrix**

Generate middle, trailing, empty, consecutive, false, undefined, multiple-run, and multiple-paragraph-equivalent cases through public PptxGenJS. Compare imported native shapes by canonical paragraph count, run grouping, visible text, style, and raw `a:p` order.

- [ ] **Step 2: Compare paragraph properties and relationships**

Use PptxGenJS run-local align/tabStops/spacing/bullet options to produce independent paragraphs, then express the equivalent native state through explicit paragraph properties plus `breakLine`. Compare final paragraph semantics and URL/internal-slide relationship targets rather than incidental relationship IDs.

- [ ] **Step 3: Record soft-break and strict-input boundaries**

Lock PptxGenJS's first-run `softBreakBefore` suppression and Native's intentional explicit-soft-break preservation. Lock upstream truthy string/object and falsy numeric runtime behavior, then require Native invalid values to throw with package state unchanged. Do not normalize away either documented difference.

- [ ] **Step 4: Run conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "breakLine|rich text line break" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "breakLine|rich text line break" --reporter=dot
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only upstream generation, exact paragraph grouping, empty/trailing guards, semantic property comparison, relationship target comparison, intentional strict/soft-break differences, production adapter non-modification, and native zero-mutation evidence. Commit `test: compare rich text line breaks with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 5: Packed Node, declarations, browser, and CLI proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual packed `@jiayunxie/pptx`, Node/browser conditional exports, emitted declarations, installed CLI, and existing rich-text smoke sections.
- Produces: `richTextBreakLine: true` across installed Node/browser smoke and a validated representative deck.

- [ ] **Step 1: Extend actual-tarball Node runtime coverage**

Using only installed artifacts, create middle/trailing/empty/consecutive splits across slide/layout/master/placeholder/declarative owners, combine per-run hyperlinks and soft breaks, live-edit, duplicate, write/reopen, and cover six formats. Require canonical getter state, source isolation, relationship targets, and add `richTextBreakLine` to final JSON.

- [ ] **Step 2: Extend installed declaration checks**

Compile legal `RichTextRun.breakLine` usage plus `@ts-expect-error` invalid run values and outer option usage. Require emitted declarations to contain the public run field and not add it to `AddTextOptions`.

- [ ] **Step 3: Extend installed CLI inspection and validation**

Run installed `pptx-inspect` validation/list/part-read commands against the generated deck. Require zero errors/warnings, expected slide count, exact canonical `a:p` grouping, preserved empty paragraph, hyperlink relationships, and absence of any private marker.

- [ ] **Step 4: Extend real Chrome coverage**

Through the browser bundle, create/edit/reopen representative splits across owners and validate. Add detailed browser state plus `richTextBreakLine: true`; require zero page, console, network, and validation errors.

- [ ] **Step 5: Run packed and build gates**

```sh
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
node_modules/.bin/tsup --config packages/pptx/tsup.config.ts
node_modules/.bin/tsup --config packages/pptx/tsup.browser.config.ts
pnpm --filter @jiayunxie/pptx run build
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review installed-only imports, type negatives, Node/browser parity, every owner, canonical getter/reopen state, relationship evidence, CLI raw-part checks, real Chrome error capture, temporary artifact containment, and preservation of all existing smoke fields. Commit `test: verify packed rich text line breaks`, push, fetch, and require divergence `0 0`.

---

### Task 6: Compatibility, visual QA, documentation, and completion audit

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: Tasks 1–5 implementation/evidence plus the repository `pptx-inspect` and presentation verification workflows.
- Produces: user-facing API contract, compatibility boundary, exact progress accounting, release note, and reproducible final QA record.

- [ ] **Step 1: Run focused PowerPoint compatibility validation**

Read and follow the local `pptx-inspect` skill before this step. Generate a representative deck with middle/trailing/empty/consecutive splits, all paragraph properties, style/soft break/hyperlinks, slide/layout/master/placeholder/declarative owners, and live editing. Validate under PowerPoint 2010, inspect exact slide/layout/master parts and relationships, and prove a boundary edit changes only its owning part plus necessary relationship entries.

- [ ] **Step 2: Run LibreOffice round-trip and visual QA**

Open/save the representative deck through LibreOffice with an isolated temporary profile. Reopen natively and require every `(owner, canonical paragraphs, visible text, hyperlink target, soft-break count)` tuple to survive. Render source and round-tripped decks to PDF/PNG; inspect every page for correct paragraph boundaries, empty-line behavior, bullets/alignment/tab stops, visible hyperlinks, no clipping, and no overflow. Keep all artifacts outside the repository.

- [ ] **Step 3: Update public documentation and progress accounting**

Document:

- `RichTextRun.breakLine?: boolean` strict input contract and post-run meaning;
- middle/trailing/empty/consecutive examples and canonical getter output;
- paragraph-property copying, explicit `softBreakBefore` composition, hyperlink reindexing, and all supported owners;
- no outer `AddTextOptions.breakLine`, no CR/LF shortcut, and no persistent/private OOXML marker;
- PptxGenJS valid boolean paragraph parity plus strict truthiness and first-run soft-break differences;
- remaining advanced line/effect/text/table, `tableToSlides`, output/runtime helpers, peer-range audit, and full-parity work.

Move `breakLine` from pending to supported without claiming remaining advanced text or full PptxGenJS parity complete. Select the next smallest remaining subitem after final gap review.

- [ ] **Step 4: Run complete release gates**

```sh
node_modules/.bin/vitest run --reporter=dot
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
node_modules/.bin/tsup --config packages/pptx/tsup.config.ts
node_modules/.bin/tsup --config packages/pptx/tsup.browser.config.ts
pnpm --filter @jiayunxie/pptx run build
node scripts/smoke-npm-package.mjs
node scripts/playwright-browser-smoke.js
git diff --check
```

Record full test counts, performance duration, PowerPoint/LibreOffice diagnostics, tarball file count/hash, Node/types/browser/CLI booleans, Chrome validation/console/page/network counts, tuple preservation, overflow results, and page-by-page visual outcome.

- [ ] **Step 5: Review completion claims**

Cross-check every supported claim against a permanent test, smoke assertion, raw OOXML record, or compatibility artifact. Search for stale statements that list `breakLine` as missing. Ensure remaining full-parity work stays explicit and no generated artifact is staged.

- [ ] **Step 6: Commit, push, and verify**

Stage only the six documentation files, commit `docs: document rich text line breaks`, push, fetch, and require divergence `0 0`.

- [ ] **Step 7: Report and continue**

Report this slice complete only after every gate passes. Summarize commits, exact test/validation/package/browser evidence, completed capability, remaining roadmap, and updated overall progress. Continue inline with the next selected parity slice; do not stop the full-parity program.
