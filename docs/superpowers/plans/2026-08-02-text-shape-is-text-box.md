# Text Shape `isTextBox` Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict `AddTextOptions.isTextBox?: boolean` creation across every text owner, expose safe direct-state editing through `ShapeModel.isTextBox`, preserve placeholder source semantics, and prove PptxGenJS 4.0.1 parity through source, packed runtime, browser, CLI, and real PPTX compatibility gates.

**Architecture:** Normalize `isTextBox` as a descriptor-safe own boolean before package mutation and carry the detached value through the existing one-pass text renderer. A focused internal codec owns strict reading and canonical replacement of `p:cNvSpPr@txBox`; the same reader supplies layout-placeholder materialization. Geometry and text/style codecs remain unchanged and orthogonal.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC transactions, PptxGenJS 4.0.1 public output, tsup, actual npm tarball smoke, real Chrome, `pptx-inspect` PowerPoint 2010 validation, LibreOffice, and PDF/PNG visual QA.

## Global Constraints

- `AddTextOptions.isTextBox` is exactly `boolean | undefined`; omitted, inherited, own `undefined`, and false normalize to false.
- Own accessors are rejected without execution. String/number/object truthiness is not copied from PptxGenJS runtime behavior.
- Canonical false is attribute absence; canonical true is unqualified `txBox="1"`.
- Plain/rich text, `addPlaceholder()`, named placeholder population, layout/master wrappers, and declarative master text/placeholder objects share one normalizer and renderer.
- `ShapeModel.isTextBox` reads false/true aliases, returns `undefined` for unsafe existing state, sets canonical boolean state, repairs a single malformed token, and rejects ambiguous structure before mutation.
- Layout placeholder materialization copies source direct semantic. Population uses the population call's value and does not mutate or inherit the layout/master source value.
- `isTextBox` changes only `p:cNvSpPr@txBox`; geometry, adjustments/`rectRadius`, transform, text, fill/line/arrows/shadow/hyperlinks, placeholder identity, relationships, parts, ordering, and cache remain independently owned.
- Changing the default from historical fixed true to false is an intentional PptxGenJS-compatible creation correction. Opening/writing an existing deck without editing must preserve source bytes.
- Do not add `breakLine`, generic non-visual properties, geometry inference, or a second text-shape renderer.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, browser artifacts, LibreOffice profiles, render output, or build output.
- Use repository-local binaries. Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0` before continuing.

---

### Task 1: Strict direct `txBox` codec

**Files:**
- Create: `packages/model/src/text-shape-is-text-box.internal.ts`
- Create: `packages/model/src/text-shape-is-text-box.internal.test.ts`

**Interfaces:**
- Consumes: `LosslessXmlDocument`, `XmlElement`, namespace-aware direct-child resolution, and `ModelParseError`.
- Produces: `readTextShapeIsTextBox()` and `replaceTextShapeIsTextBox()` for one `p:sp/p:nvSpPr/p:cNvSpPr` state.

- [ ] **Step 1: Add failing reader fixtures**

Build compact real-namespace `p:sp` fixtures and require:

- attribute absence -> false;
- `1` / `true` / `on` -> true;
- `0` / `false` / `off` -> false;
- invalid token -> undefined;
- alternate presentation prefix with unqualified `txBox` remains readable;
- foreign `sp`, qualified lookalike attribute, duplicate attribute, missing/repeated direct `nvSpPr` or `cNvSpPr`, and descendant-only lookalikes -> undefined.

- [ ] **Step 2: Add failing replacement fixtures**

Require false->true insertion, true/false/invalid alias canonicalization, true->false removal, exact canonical same-value bytes no-op, and preservation of all neighboring attributes/children/comments/whitespace. Ambiguous states must throw before changing the XML object.

- [ ] **Step 3: Implement the focused codec**

Resolve namespace-correct direct owners only. Record zero or one semantic `txBox` attribute, treat qualified or repeated lookalikes as unsafe, and use source offsets for surgical insertion/replacement/removal. Return a boolean from replacement so callers avoid `setPart()` and mutation journal changes on exact no-op.

- [ ] **Step 4: Run focused gates**

```sh
node_modules/.bin/vitest run packages/model/src/text-shape-is-text-box.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review namespace ownership, direct-child cardinality, boolean alias set, qualified/duplicate attribute handling, source-offset removal, no-op behavior, malformed recovery, and unrelated-byte isolation. Commit `feat: add text box state codec`, push, fetch, and require divergence `0 0`.

---

### Task 2: Core creator and live `ShapeModel` integration

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 codec, `validateAddTextOptions()`, `NormalizedTextInput`, `NormalizedAddTextOptions`, `addTextShape()`, and `textShapeXml()`.
- Produces: public `AddTextOptions.isTextBox`, one-pass canonical creation, `SlideModel` read/set methods, and `ShapeModel.isTextBox`.

- [ ] **Step 1: Add failing normalization and atomicity tests**

Snapshot part bytes, relationships, part URIs, mutations, shape order, and an existing live handle. Require exact zero mutation for string, number, null, object, symbol, boxed boolean, function, and own accessor inputs. Require accessor non-execution, inherited value ignored, own undefined equal to omitted, and primitive false/true accepted.

- [ ] **Step 2: Add failing creation and live-edit tests**

Create omitted/undefined/false/true plain, rich, empty, multiline, and direct placeholder text. Require exact `cNvSpPr` XML and immediate getter. Cover canonical same-value bytes/journal no-op, alias canonicalization, malformed-token recovery, ambiguous-state rejection, stable live identity, and unrelated sibling isolation.

- [ ] **Step 3: Add failing ownership/lifecycle tests**

Combine `isTextBox` with every implemented text geometry/radius/style/text-body surface. Edit `presetType`, adjustments, transform, fill, line, arrows, shadow, hyperlinks, margins, valign, direction, fit, wrap, plain/rich text, and `isTextBox` in both directions; require only the owning state to change. Cover duplicate independence, move, outer rollback, write/reopen, and all six presentation formats.

- [ ] **Step 4: Implement strict normalization and one-pass rendering**

Add `readonly isTextBox?: boolean` and a descriptor-safe normalizer:

```ts
function normalizeTextShapeIsTextBox(options: AddTextOptions): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(options, 'isTextBox');
  if (!descriptor) return false;
  if (!Object.hasOwn(descriptor, 'value')) {
    throw new TypeError('Text shape isTextBox must be a data property');
  }
  const value = descriptor.value;
  if (value === undefined) return false;
  if (typeof value !== 'boolean') {
    throw new TypeError('Text shape isTextBox must be a boolean');
  }
  return value;
}
```

Thread the normalized boolean through every plain/rich/placeholder `addTextShape()` call and render `<p:cNvSpPr${isTextBox ? ' txBox="1"' : ''}/>` during initial creation. Do not call a live setter after shape creation.

- [ ] **Step 5: Expose the live model surface**

Add `SlideModel.getShapeIsTextBox()` / `setShapeIsTextBox()` around Task 1. The setter validates primitive boolean before entering the package transaction and calls `setXml()` only when replacement reports a change. Add getter/setter accessors on `ShapeModel`.

- [ ] **Step 6: Run core gates**

```sh
node_modules/.bin/vitest run packages/model/src/text-shape-is-text-box.internal.test.ts packages/model/src/model.test.ts -t "text box state|isTextBox" --reporter=dot
node_modules/.bin/vitest run packages/model/src/model.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 7: Review, commit, push, and verify**

Review validation-before-mutation, all renderer call sites, default behavior correction, exact child order, public getter/setter typing, same-value no-op, malformed/ambiguous boundaries, outer rollback, six formats, and geometry/style isolation. Commit `feat: create and edit text box state`, push, fetch, and require divergence `0 0`.

---

### Task 3: Placeholder materialization and all public owners

**Files:**
- Modify: `packages/model/src/placeholder.internal.ts`
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 reader, Task 2 public option/model accessors, `MaterializedPlaceholderDescriptor`, `TEXT_OPTION_KEYS`, layout/master wrappers, and declarative master cloning.
- Produces: source-faithful materialization, all-owner lifecycle evidence, closed-key support, and root declaration/type coverage.

- [ ] **Step 1: Add failing materialization tests**

Define layout placeholders with omitted/false/true and boolean aliases. Add a slide and require materialized owners to preserve semantic false/true while using canonical output. Invalid or ambiguous source state must abort slide creation atomically. A non-`p:sp` placeholder materializes false without a forged attribute.

- [ ] **Step 2: Add failing population tests**

Populate source-false with true, source-true with false, and both with omitted. Require population-call default/explicit state, stable owner id/name/identity/transform, and exact layout/master source bytes. Cover plain and rich population plus duplicate/rollback/reopen.

- [ ] **Step 3: Add failing public-owner/declarative tests**

Exercise slide/layout/master direct plain/rich text, layout/master placeholders, declarative text/placeholder objects, and delayed asynchronous master resources. Mutate source option objects after invocation and require detached values. Invalid/accessor inputs must fail before observable package mutation.

- [ ] **Step 4: Add root type/runtime tests**

Import `AddTextOptions` and `ShapeModel` from the root package. Compile legal boolean creation and live assignment; add `@ts-expect-error` cases for string, number, null, object, and assigning undefined to the setter. Verify runtime create/read/edit/write/reopen.

- [ ] **Step 5: Implement materialization and declarative support**

Read source `isTextBox` into `MaterializedPlaceholderDescriptor`; reject undefined unsafe state and render canonical false/true. Add exactly `'isTextBox'` to `TEXT_OPTION_KEYS`; `PLACEHOLDER_OPTION_KEYS` inherits it. Do not clone or validate through a second path.

- [ ] **Step 6: Run public/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text box state|isTextBox" --reporter=dot
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -p packages/pptx/tsconfig.json --pretty false
git diff --check
```

- [ ] **Step 7: Review, commit, push, and verify**

Review layout-source reading, non-shape fallback, population precedence, owner identity/transform stability, source isolation, declarative clone timing, closed-key expansion, root declarations, duplicate/rollback/six-format persistence, and no unrelated master behavior changes. Commit `feat: preserve text box owner semantics`, push, fetch, and require divergence `0 0`.

---

### Task 4: PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()`/`defineSlideMaster()`/`write()`, package import, raw OOXML helpers, and native Task 2/3 APIs.
- Produces: permanent parity evidence and explicit strict-divergence records.

- [ ] **Step 1: Add public-only creation fixtures**

Generate omitted/undefined/false/true plain and rich text through public PptxGenJS. Compare imported `ShapeModel.isTextBox` and raw `cNvSpPr`; false/true controls must differ only by ` txBox="1"`.

- [ ] **Step 2: Add master/placeholder fixtures**

Generate master text, omitted/false/true placeholders, empty slide materialization, and false<->true population. Compare layout and slide owner state with native equivalents, including source isolation.

- [ ] **Step 3: Prove geometry/style orthogonality and strict divergence**

Combine `shape`, `rectRadius`, fill/line/arrows/shadow/hyperlink, and rich text; normalize only `txBox` and require remaining XML equality. Lock PptxGenJS truthy string/object and falsy numeric runtime behavior, then require native invalid inputs to throw with package state unchanged.

- [ ] **Step 4: Run conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text box state|isTextBox" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text box state|isTextBox" --reporter=dot
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review public-only upstream generation, exact raw attribute assertions, all master/placeholder surfaces, supported boolean final semantics, intentional truthiness divergence, production adapter non-modification, and native zero-mutation evidence. Commit `test: compare text box state with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 5: Packed Node, declarations, browser, and CLI proof

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual packed `@jiayunxie/pptx`, Node/browser conditional exports, emitted declarations, installed CLI, and existing text-shape smoke sections.
- Produces: `textShapeIsTextBox: true` across installed Node/browser smoke and a validated representative deck.

- [ ] **Step 1: Extend actual-tarball Node runtime coverage**

Using only installed artifacts, create omitted/false/true plain/rich/placeholder/layout/master/declarative text, materialize and populate placeholders, live-edit in both directions, duplicate, write/reopen, and cover six formats. Require immediate state, same-value no-op, source isolation, and add `textShapeIsTextBox` to final JSON.

- [ ] **Step 2: Extend installed declaration checks**

Compile legal `AddTextOptions.isTextBox` and `ShapeModel.isTextBox` usage plus `@ts-expect-error` invalid inputs/assignment. Require emitted declarations to contain both public surfaces with the intended read/write types.

- [ ] **Step 3: Extend installed CLI inspection and validation**

Run installed `pptx-inspect` validation/list/part-read commands against the generated deck. Require zero errors/warnings, expected slide count, and exact true/false `cNvSpPr` fragments in slide/layout/master parts.

- [ ] **Step 4: Extend real Chrome coverage**

Through the browser bundle, create omitted/false/true text across owners, materialize/populate, edit/reopen, and validate. Add detailed browser state plus `textShapeIsTextBox: true`; require zero page, console, network, and validation errors.

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

Review installed-only imports, type negatives, Node/browser parity, placeholder source semantics, CLI raw-part evidence, real Chrome error capture, temporary artifact containment, and preservation of all existing smoke fields. Commit `test: verify packed text box state`, push, fetch, and require divergence `0 0`.

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

Read and follow the local `pptx-inspect` skill before this step. Generate a representative deck with omitted/false/true across plain/rich text, geometry/radius/styles, layout/master placeholders, materialized owners, and populated owners. Validate under PowerPoint 2010, inspect exact slide/layout/master parts, and prove a live toggle changes only its owning part and `txBox` attribute.

- [ ] **Step 2: Run LibreOffice round-trip and visual QA**

Open/save the representative deck through LibreOffice with an isolated temporary profile. Reopen natively and require every `(owner, text, presetType, adjustments, isTextBox)` tuple to survive. Render source and round-tripped decks to PDF/PNG; inspect every page for visible text, stable geometry/style, no clipping, and no overflow. Keep all artifacts outside the repository.

- [ ] **Step 3: Update public documentation and progress accounting**

Document:

- `AddTextOptions.isTextBox?: boolean` default/strict input contract;
- canonical false absence and canonical true `txBox="1"` output;
- `ShapeModel.isTextBox` readable aliases, unsafe `undefined`, canonical setter, and ambiguity behavior;
- all supported owners, materialization copy, population precedence, and source isolation;
- geometry/radius/style orthogonality and the intentional historical-default correction;
- PptxGenJS valid boolean parity plus strict runtime truthiness divergence;
- remaining `breakLine`, advanced line/effect/text/table, `tableToSlides`, output/runtime helpers, peer-range audit, and full-parity work.

Move `isTextBox` from pending to supported without claiming advanced text or full PptxGenJS parity complete. Select the next smallest remaining subitem after final gap review.

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

Cross-check every supported claim against a permanent test, smoke assertion, raw OOXML record, or compatibility artifact. Search for stale statements that list `isTextBox` as missing. Ensure `breakLine` and full parity remain explicitly outstanding and no generated artifact is staged.

- [ ] **Step 6: Commit, push, and verify**

Stage only the six documentation files, commit `docs: document text box state`, push, fetch, and require divergence `0 0`.

- [ ] **Step 7: Report and continue**

Report this slice complete only after every gate passes. Summarize commits, exact test/validation/package/browser evidence, completed capability, remaining roadmap, and updated overall progress. Continue inline with the next selected parity slice; do not stop the full-parity program.
