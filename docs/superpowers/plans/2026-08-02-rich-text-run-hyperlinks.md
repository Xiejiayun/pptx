# Rich Text Run Hyperlinks Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict per-run rich-text hyperlink creation, reading, editing, clearing, and relationship lifecycle support across every public rich-text owner, with PptxGenJS 4.0.1 valid-output parity.

**Architecture:** Extend `RichTextRunStyle` with a run-local `Hyperlink | false` direct state. Keep normalization/rendering in `rich-text.internal.ts`, relationship interpretation in the existing shape-hyperlink codec, and relationship allocation/update/GC in `SlideModel`; explicit run links own independent relationships, while omitted run values may inherit the existing outer shared relationship during creation.

**Tech Stack:** TypeScript strict mode, lossless OOXML source spans, OPC transactions/relationships, Vitest, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, real Google Chrome, and `pptx-inspect`.

## Global Constraints

- `RichTextRunStyle.hyperlink` is exactly `Hyperlink | false | undefined`; no aliases, raw relationship IDs, action strings, or coercion.
- Explicit linked runs require non-empty text and each receive an independent relationship, even for identical targets.
- Omitted run hyperlink inherits `AddTextOptions.hyperlink`; `false` suppresses that default; explicit run value overrides it.
- `ShapeModel.hyperlink` remains whole-shape-only; `ShapeModel.richText` owns run-local state.
- Native accepts PptxGenJS URL `action=""` during run reads but does not copy orphan, `rIdundefined`, dangling, falsy-underline, or console-only defects.
- All internal slide targets resolve against current presentation identity before observable mutation.
- Every independently reviewable task ends with review, commit, push, and remote divergence verification.
- Execute inline in the current task; do not dispatch subagents.

---

### Task 1: Public Run Value, Strict Normalization, and Pure Rendering

**Files:**
- Modify: `packages/model/src/text.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Test: `packages/model/src/rich-text.internal.test.ts`

**Interfaces:**
- Consumes: existing public `Hyperlink`, internal `NormalizedHyperlink`, `normalizeHyperlink()`, and `renderShapeHyperlink()`.
- Produces: `RichTextRunStyle.hyperlink?: Hyperlink | false`, normalized run-local values, `RichTextRunHyperlinkRelationshipIds`, and pure run rendering with local/default precedence.

- [ ] **Step 1: Write failing normalization and render tests**

Add cases proving explicit URL/internal/tooltip normalization, nested caller detachment, `false`, inherited outer default, explicit override, suppression, independent supplied IDs, empty-linked-run rejection, and underline precedence:

```ts
const paragraphs = normalizeRichText([{
  runs: [
    { text: 'Inherited' },
    { text: 'Local', style: { hyperlink: { url: 'https://local.example' } } },
    { text: 'Suppressed', style: { hyperlink: false } },
    { text: 'No underline', style: {
      hyperlink: { slide: 2, tooltip: '' },
      underline: false,
    } },
  ],
}]);
const rendered = renderRichTextParagraphs(paragraphs, {
  defaultHyperlink: normalizeHyperlink({ url: 'https://outer.example' }, 'outer'),
  hyperlinkRelationshipId: 'rIdOuter',
  runHyperlinkRelationshipIds: [[undefined, 'rIdLocal', undefined, 'rIdInternal']],
});
expect(rendered.match(/r:id="rIdOuter"/g)).toHaveLength(1);
expect(rendered).toContain('r:id="rIdLocal"');
expect(rendered).not.toMatch(/Suppressed[\s\S]*?<a:hlinkClick/);
expect(rendered).toMatch(/u="none"[^>]*>[\s\S]*r:id="rIdInternal"/);
```

- [ ] **Step 2: Run the focused test and verify failure**

Run:

```bash
./node_modules/.bin/vitest run packages/model/src/rich-text.internal.test.ts
```

Expected: FAIL because `RichTextRunStyle.hyperlink` and `runHyperlinkRelationshipIds` are not implemented.

- [ ] **Step 3: Add the public type and strict normalizer**

In `text.ts`:

```ts
import type { Hyperlink } from './hyperlink.js';

export interface RichTextRunStyle {
  readonly hyperlink?: Hyperlink | false;
}
```

In `normalizeStyle()` add `hyperlink` to the closed key list and normalize it without getters:

```ts
const hyperlink = candidate.hyperlink === undefined
  ? undefined
  : candidate.hyperlink === false
    ? false
    : normalizeHyperlink(candidate.hyperlink, `${context} hyperlink`);
if (hyperlink !== undefined && hyperlink !== false && runText.length === 0) {
  throw new TypeError(`${context} hyperlink requires non-empty text`);
}
```

Pass `run.text` into style normalization so empty-link rejection occurs during input validation.

- [ ] **Step 4: Add a parallel relationship-ID matrix to the pure renderer**

Define:

```ts
export type RichTextRunHyperlinkRelationshipIds =
  readonly (readonly (string | undefined)[])[];
```

Resolve each run with explicit-local > false suppression > outer default, require paired IDs, and keep explicit underline authoritative:

```ts
const local = run.style?.hyperlink;
const hyperlink = local === false ? undefined : local ?? options.defaultHyperlink;
const relationshipId = local && local !== false
  ? options.runHyperlinkRelationshipIds?.[paragraphIndex]?.[runIndex]
  : hyperlink === undefined ? undefined : options.hyperlinkRelationshipId;
```

- [ ] **Step 5: Run model internal tests and typecheck**

Run:

```bash
./node_modules/.bin/vitest run packages/model/src/rich-text.internal.test.ts
./node_modules/.bin/tsc -b --pretty false
```

Expected: all focused tests pass and TypeScript reports no errors.

- [ ] **Step 6: Review, commit, push, and verify remote parity**

```bash
git diff --check
git add packages/model/src/text.ts packages/model/src/rich-text.internal.ts packages/model/src/rich-text.internal.test.ts
git commit -m "feat: define rich text run hyperlinks"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.

### Task 2: Model Creation and Outer-Default Interaction

**Files:**
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: normalized run hyperlinks and `RichTextRunHyperlinkRelationshipIds` from Task 1, existing `createTextHyperlinkRelationship()`, and outer text-shape creation transactions.
- Produces: per-run relationships and binding matrices for `addRichText()` / rich `addPlaceholder()` on slide/layout/master owners.

- [ ] **Step 1: Write failing creation tests**

Cover local URL/internal links, identical targets with distinct IDs, outer inheritance, local override, false suppression, explicit tooltip empty, style coexistence, and zero-mutation invalid target/empty run:

```ts
const shape = source.addRichText([{
  runs: [
    { text: 'Outer' },
    { text: 'Local', style: { hyperlink: { url: 'https://local.example' } } },
    { text: 'None', style: { hyperlink: false } },
    { text: 'Target', style: { hyperlink: { slide: 2, tooltip: '' } } },
  ],
}], { hyperlink: { url: 'https://outer.example' } });
expect(shape.hyperlink).toEqual({ url: 'https://outer.example' });
expect(shape.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
  { url: 'https://outer.example' },
  { url: 'https://local.example' },
  undefined,
  { slide: 2, tooltip: '' },
]);
```

- [ ] **Step 2: Run the focused model test and verify failure**

```bash
./node_modules/.bin/vitest run packages/model/src/model.test.ts -t "rich text run hyperlink creation"
```

Expected: FAIL because SlideModel does not allocate run-local relationships.

- [ ] **Step 3: Add target preparation without observable mutation**

Create a private preparation helper that walks normalized paragraphs, resolves every internal target to a current `SlideModel`, and returns relationship inputs before adding any relationship:

```ts
interface PreparedRunHyperlink {
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly hyperlink: NormalizedHyperlink;
  readonly relationship: RelationshipInput;
}
```

Reuse `relativeRelationshipTarget(this.partUri, target.partUri)` and the existing hyperlink/slide relationship constants.

- [ ] **Step 4: Allocate one relationship per explicit run inside the outer transaction**

Build the matrix after preparation:

```ts
const ids = paragraphs.map(({ runs }) => runs.map(() => undefined as string | undefined));
for (const prepared of runHyperlinks) {
  ids[prepared.paragraphIndex]![prepared.runIndex] =
    this.presentation.opcPackage.addRelationship(this.partUri, prepared.relationship).id;
}
```

Keep outer omitted runs bound to the existing `hyperlinkRelationshipId`; do not allocate for `false`.

- [ ] **Step 5: Pass bindings through every rich creation surface**

Update rich `addPlaceholder()` and `addRichText()` calls to `renderRichTextParagraphs()` with `runHyperlinkRelationshipIds`. Because layout/master wrappers call the same SlideModel methods, no parallel renderer is added.

- [ ] **Step 6: Run focused and complete model suites**

```bash
./node_modules/.bin/vitest run packages/model/src/rich-text.internal.test.ts packages/model/src/model.test.ts
./node_modules/.bin/tsc -b --pretty false
```

- [ ] **Step 7: Review, commit, push, and verify remote parity**

```bash
git diff --check
git add packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: create rich text run hyperlinks"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

### Task 3: Strict Run Reader

**Files:**
- Modify: `packages/model/src/shape-hyperlink.internal.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `ShapeHyperlinkReadContext`, exact relationship/action parsing, `readRichText()`.
- Produces: `readTextRunHyperlink()`, strict direct run snapshots, and PptxGenJS URL `action=""` compatibility.

- [ ] **Step 1: Add failing imported-run read tests**

Create namespace-correct fixtures for URL, internal, tooltip absent/empty, PptxGenJS URL attributes/extension, wrong mode/type/action, dangling IDs, duplicate clicks, and wrong namespace. Require only valid direct state in `shape.richText`.

- [ ] **Step 2: Run the reader tests and verify failure**

```bash
./node_modules/.bin/vitest run packages/model/src/model.test.ts -t "reads strict rich text run hyperlinks"
```

- [ ] **Step 3: Generalize the hyperlink element decoder**

Add an exported run-specific reader that accepts a direct `rPr` and context. Preserve existing shape-reader strictness while allowing absent or empty action for external run URLs:

```ts
export function readTextRunHyperlink(
  properties: XmlElement,
  context: ShapeHyperlinkReadContext,
): NormalizedHyperlink | undefined;
```

Require one direct DrawingML click. Accept legal `snd`/`extLst` children and PptxGenJS extra attributes without exposing them.

- [ ] **Step 4: Thread read context into rich-text decoding**

Change:

```ts
readRichText(xml, element, {
  relationships: this.relationships,
  slidePartUris: this.presentation.slides.map(({ partUri }) => partUri),
});
```

Add `hyperlink` to `readStyle()` only when `readTextRunHyperlink()` succeeds.

- [ ] **Step 5: Run reader, model, and type tests**

```bash
./node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/rich-text.internal.test.ts
./node_modules/.bin/tsc -b --pretty false
```

- [ ] **Step 6: Review, commit, push, and verify remote parity**

```bash
git diff --check
git add packages/model/src/shape-hyperlink.internal.ts packages/model/src/rich-text.internal.ts packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: read rich text run hyperlinks"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

### Task 4: Whole-Rich-Text Editing and Relationship Lifecycle

**Files:**
- Modify: `packages/model/src/shape-hyperlink.internal.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Modify: `packages/model/src/slide.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: strict reader/bindings from Task 3, OPC relationship add/update/remove, `relationshipReferenceCount()`.
- Produces: same-value no-op, index-stable ID reuse, unique update, shared clone-on-write, clear/GC, rollback, and reopen behavior through `ShapeModel.richText`.

- [ ] **Step 1: Write failing lifecycle tests**

Require:

- `shape.richText = shape.richText` leaves exact package bytes, relationships, and journal unchanged;
- tooltip-only change reuses ID;
- unique target change updates relationship in place;
- outer/shared target change clones on write and preserves whole-shape/other-run target;
- clear one run keeps still-referenced relationships; clear final reference removes it;
- invalid later run rolls back all earlier prepared changes;
- write/reopen returns edited hyperlinks.

- [ ] **Step 2: Run lifecycle tests and verify failure**

```bash
./node_modules/.bin/vitest run packages/model/src/model.test.ts -t "edits rich text run hyperlinks"
```

- [ ] **Step 3: Expose strict existing run binding IDs**

Add an internal matrix containing value plus relationship ID by paragraph/run index:

```ts
export interface ReadRichTextRunHyperlinkBinding {
  readonly hyperlink: NormalizedHyperlink;
  readonly relationshipId: string;
}
```

Do not expose IDs through public snapshots.

- [ ] **Step 4: Prepare relationship mutations before text replacement**

For each target run:

```ts
if (old && shapeHyperlinkTargetsEqual(old.hyperlink, next)) reuse old.relationshipId;
else if (old && relationshipReferenceCount(xml, old.relationshipId) === 1)
  updateRelationship(...);
else addRelationship(...);
```

Collect all prior text-body relationship IDs before replacement. After `setXml(updated)`, remove only IDs with global reference count zero.

- [ ] **Step 5: Add exact same-value fast path**

Normalize input, compare to the current public snapshot including tooltip property presence and `false` semantics, and return before relationship allocation or XML rendering when equal.

- [ ] **Step 6: Run full model and transaction suites**

```bash
./node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/shape-hyperlink.internal.test.ts packages/model/src/rich-text.internal.test.ts
./node_modules/.bin/tsc -b --pretty false
```

- [ ] **Step 7: Review, commit, push, and verify remote parity**

```bash
git diff --check
git add packages/model/src/shape-hyperlink.internal.ts packages/model/src/rich-text.internal.ts packages/model/src/slide.ts packages/model/src/model.test.ts
git commit -m "feat: edit rich text run hyperlinks"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

### Task 5: SDK, Root Package, Declarative Owners, and Six-Format Lifecycle

**Files:**
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`
- Modify if required by closed cloning: `packages/sdk/src/master-layout.ts`

**Interfaces:**
- Consumes: public model types and runtime from Tasks 1–4.
- Produces: verified aggregate exports, slide/layout/master/placeholder/declarative behavior, and six-format lifecycle coverage.

- [ ] **Step 1: Add failing SDK lifecycle tests**

Cover rich text on slide/layout/master, rich placeholder prompt/population, declarative text/placeholder, async caller detachment, duplicate external/other/self, move, target deletion, rollback, and six formats.

- [ ] **Step 2: Add root type/runtime tests**

Compile and execute:

```ts
const style: RichTextRunStyle = {
  hyperlink: { url: 'https://example.com', tooltip: '' },
};
const suppressed: RichTextRunStyle = { hyperlink: false };
```

Also retain `@ts-expect-error` cases for both target branches, aliases, non-string tooltip, and true suppression.

- [ ] **Step 3: Run focused SDK/root tests and verify failure or pass-through**

```bash
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "rich text run hyperlink"
```

Expected: tests fail until any closed clone/key boundary is updated; if runtime already passes through, the new tests still prove the surface.

- [ ] **Step 4: Update only required closed key/clone boundaries**

If `master-layout.ts` clones run styles by explicit keys, add `hyperlink` and preserve nested `Hyperlink` detachment. Do not add parallel SDK-only types.

- [ ] **Step 5: Run complete SDK/root/model suites**

```bash
./node_modules/.bin/vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts
./node_modules/.bin/tsc -b --pretty false
```

- [ ] **Step 6: Review, commit, push, and verify remote parity**

```bash
git diff --check
git add packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/sdk/src/master-layout.ts
git commit -m "test: cover public rich text run hyperlinks"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Stage `master-layout.ts` only if modified.

### Task 6: PptxGenJS 4.0.1 Public-Output Conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()` / `write()`, adapter import, and native run snapshot/editing.
- Produces: permanent evidence for valid parity and intentional strict corrections.

- [ ] **Step 1: Add public-only PptxGenJS fixtures**

Generate run URL/internal/tooltip, identical-target runs, mixed linked/unlinked, color extension, empty linked run, outer+local, underline false, and invalid runtime values without reading PptxGenJS private fields.

- [ ] **Step 2: Assert exact imported final state**

Require `ShapeModel.richText` to expose valid run links, identical target runs to retain distinct relationship IDs in raw XML, and whole-shape hyperlink to remain absent for per-run-only output.

- [ ] **Step 3: Add native comparison and defect-correction assertions**

Assert native valid output matches target/type/mode/action/run selection while preserving omitted tooltip, explicit underline false, zero-mutation rejection, valid outer+local relationships, and no empty-run orphan.

- [ ] **Step 4: Run adapter and cross-package focused gates**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts
./node_modules/.bin/vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts -t "rich text run hyperlink"
```

- [ ] **Step 5: Review, commit, push, and verify remote parity**

```bash
git diff --check
git add packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "test: compare rich text run hyperlinks with pptxgenjs"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

### Task 7: Packed Node, Types, Browser, CLI, and Documentation

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: completed public run hyperlink API and lifecycle.
- Produces: `richTextRunHyperlinks: true` release evidence and updated parity ledger.

- [ ] **Step 1: Extend actual-tarball Node/type/CLI smoke**

Using only the installed tarball, create outer/local/suppressed URL/internal/self links across slide/layout/master/placeholder/declarative owners; read/edit/clear, duplicate/move/delete, write/reopen, and report:

```js
richTextRunHyperlinks: true
```

Add declaration consumer checks for `Hyperlink | false` plus invalid compile-time cases. Save internal-only and external smoke decks; require inspect/slides/part-read and PowerPoint 2010 validation behavior.

- [ ] **Step 2: Extend real-Chrome smoke**

Create/read/edit/clear/reopen run links through the browser bundle and add exact state, relationship independence, action, validation, and `errorCounts: { console: 0, page: 0, network: 0 }` expectations.

- [ ] **Step 3: Run full verification**

```bash
./node_modules/.bin/vitest run
RUN_PERF=1 ./node_modules/.bin/vitest run packages/testkit/src/performance.test.ts
./node_modules/.bin/tsc -b --pretty false
./node_modules/.bin/tsc -b
```

Then run both tsup configs, declaration build, pack to a fresh `mktemp -d` directory, `node scripts/smoke-npm-package.mjs <tarball>`, and the real-Google-Chrome payload.

- [ ] **Step 4: Update all six public/progress documents**

Document public type, direct read/edit semantics, outer inheritance/false suppression, independent explicit relationships, underline precedence, lifecycle, PptxGenJS differences, final test/build/tarball/browser/CLI evidence, and select the next actual parity gap.

- [ ] **Step 5: Run stale-language and diff review**

```bash
rg -n -i 'per-run rich-text hyperlink.*(next|pending|尚未|下一)' --glob '*.md' .
git diff --check
git status --short
```

Historical design/plan scope statements may remain; public current-state documents must contain no stale gap claim.

- [ ] **Step 6: Review, commit, push, and verify remote parity**

```bash
git add scripts/smoke-npm-package.mjs scripts/playwright-browser-smoke.js README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
git commit -m "docs: document rich text run hyperlink support"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; the only untracked workspace item remains the pre-existing `.pnpm-store/`.
