# Text Shape Hyperlink Creation Implementation Plan

> **For agentic workers:** Execute this plan inline task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking; do not dispatch subagents. Every repository-changing task ends with an independent review, commit, push, fetch, and remote-divergence check.

**Goal:** Add strict uniform hyperlink creation to every public text-shape creation surface, matching valid PptxGenJS 4.0.1 shape/run output and proving actual packed Node/browser/type/CLI behavior.

**Architecture:** Reuse the existing `Hyperlink` union and `normalizeHyperlink()` / `renderShapeHyperlink()` codec. Normalize once inside `validateAddTextOptions()`, resolve one target relationship per text shape, and render the same ID into `p:cNvPr` plus every materialized `a:rPr`. Existing `ShapeModel.hyperlink` remains the whole-shape editor; run read/edit stays a later vertical slice.

**Tech Stack:** TypeScript strict mode, Vitest, lossless OOXML, OPC relationships and transactions, PptxGenJS 4.0.1 public output, tsup, npm tarball smoke, Playwright Chromium smoke, and `pptx-inspect` PowerPoint 2010 validation.

## Global Constraints

- `AddTextOptions.hyperlink` accepts only the existing mutually exclusive `Hyperlink`: one non-empty XML-safe `url` or one positive safe-integer one-based `slide`, plus optional XML-safe `tooltip`.
- Omitted/runtime-`undefined` hyperlink creates neither click XML nor relationship and must preserve the published text-shape bytes.
- Normalize and detach before mutation. Accessors, symbols, arrays, custom prototypes, inherited fields, unknown keys, aliases, missing/both targets, coercions, empty URL, invalid tooltip, and invalid/out-of-range slides leave bytes, relationships, parts, journal, shape order, and live caches unchanged.
- Plain/rich text, empty/multiline content, placeholder creation/population, layout/master wrappers, and declarative `defineSlideMaster()` text/placeholder objects share one implementation.
- One linked text shape owns one relationship. Its direct `p:cNvPr/a:hlinkClick` and every emitted non-empty text run's `a:rPr/a:hlinkClick` share that ID.
- URL relationships are exact external hyperlink relationships. Slide relationships are exact internal slide relationships to a real target part, with `ppaction://hlinksldjump` on every click element.
- Run hyperlinks default to single underline; an explicit rich-run underline state remains authoritative. Existing run color/font/effects/highlight/outline children stay intact and valid.
- Existing `ShapeModel.hyperlink` edits only `p:cNvPr`. Shared run references force target edits to clone-on-write and prevent clear from prematurely collecting the original relationship.
- Duplicate preserves external/other-slide targets and retargets self-links to the duplicate. Move/insert/delete update ordinal reporting without changing target identity. Target deletion removes all matching DrawingML click/hover references.
- Do not add per-run hyperlink public values, run hyperlink reading/editing, hover, table-cell, image, chart, media, group, graphic-frame, custom-show, macro, program, sound, or action-only navigation.
- PptxGenJS `rIdundefined`, dangling targets, double targets, coercions, console-only failures, and always-empty tooltip attributes are evidence, not behavior to copy.
- Never stage `.pnpm-store/`, tarballs, temporary consumers, generated decks, browser artifacts, or build output.
- Every implementation/evidence commit is pushed to `main`, fetched, and verified at divergence `0 0`.

---

### Task 1: Add strict text hyperlink normalization, relationship allocation, and dual rendering

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/rich-text.internal.ts`
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/model/src/rich-text.internal.test.ts`

**Interfaces:**
- Consumes: public `Hyperlink`, `NormalizedHyperlink`, `normalizeHyperlink()`, `renderShapeHyperlink()`, `HYPERLINK_RELATIONSHIP_TYPE`, `SLIDE_RELATIONSHIP_TYPE`, `relativeRelationshipTarget()`, and existing text renderers.
- Produces: `AddTextOptions.hyperlink?: Hyperlink`, normalized `hyperlink?: NormalizedHyperlink`, one shared relationship ID, canonical whole-shape click, and canonical click children in plain/rich runs.

- [ ] **Step 1: Write failing public-model contract tests**

Create URL, internal-slide, self, tooltip-absent, tooltip-empty, Unicode/metacharacter, empty text, multiline plain, and multi-run rich fixtures:

```ts
const source = {
  url: 'https://example.com/path?a=1&b=2',
  tooltip: 'Visit <now>',
} as const;
const linked = slide.addText('First\nSecond', { hyperlink: source });
const rich = slide.addRichText([{
  runs: [
    { text: 'One' },
    { text: 'Two', style: { underline: false, color: { kind: 'scheme', value: 'accent2' } } },
  ],
}], { hyperlink: { slide: 2 } });
```

Require:

- immediate detached/frozen `ShapeModel.hyperlink` snapshot;
- one relationship per shape, exact type/mode/target, and caller detachment;
- one `cNvPr` click plus one click per non-empty run, all sharing the relationship ID;
- internal action on shape and runs, exact tooltip absence/empty, and XML escaping;
- default `u="sng"`, with explicit rich underline preserved;
- empty text retains only the shape click; multiline emits linked runs for non-empty lines;
- combined fill/line/arrows/shadow output remains unchanged outside hyperlink-owned bytes.

- [ ] **Step 2: Write failing invalid-input zero-mutation tests**

Snapshot slide bytes, relationships, package part URIs, mutation journal, shape array, and existing model identity before each call. Reject malformed values through `addText()`, `addRichText()`, and `addPlaceholder()`:

```ts
{}
{ url: '', slide: undefined }
{ url: 'https://example.com', slide: 2 }
{ slide: 0 }
{ slide: 1.5 }
{ slide: 99 }
{ url: 123 }
{ target: 'https://example.com' }
```

Also reject null/primitives/arrays/dates/class instances, inherited-only properties, all symbols, own accessors without invocation, unknown keys, invalid XML controls, unsafe integers, NaN/infinity, and `_rId`/`kind` aliases. Require runtime `hyperlink: undefined` to equal omitted output.

- [ ] **Step 3: Write failing rich-text renderer tests**

Extend focused internal tests to require an optional uniform default hyperlink and relationship ID to:

- append canonical click XML after existing run font children;
- add single underline only when the run has no explicit underline;
- leave empty style-less runs absent;
- preserve color, transparency, outline, glow, highlight, underline fill, and font XML;
- reject supplying a hyperlink without an ID or an ID without a hyperlink;
- preserve current exact output when both are absent.

- [ ] **Step 4: Run focused tests and confirm intended failure**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/rich-text.internal.test.ts -t "text shape hyperlink|uniform text hyperlink" --reporter=dot
```

Expected: compile/runtime failure because `AddTextOptions.hyperlink` and dual rendering are absent.

- [ ] **Step 5: Implement minimal normalization and renderer changes**

In `packages/model/src/slide.ts`:

```ts
export interface AddTextOptions extends Partial<Transform> {
  readonly hyperlink?: Hyperlink;
  // retain every existing field
}

interface NormalizedAddTextOptions {
  readonly hyperlink?: NormalizedHyperlink;
  // retain every existing field
}

const hyperlink = options.hyperlink === undefined
  ? undefined
  : normalizeHyperlink(options.hyperlink, 'Text shape hyperlink');
```

Thread it through `NormalizedTextInput` and all plain/rich/placeholder creation paths. After all input/owner/identity validation, resolve internal target existence and add exactly one relationship to the current owner part. Pass its ID to paragraph and shape rendering. A small focused helper may remove plain/rich duplication but must not refactor preset-shape creation in this task.

Render the whole-shape click with the existing codec and conditionally bind the relationship namespace. Extend `textParagraphXml()` / `defaultTextRunXml()` and `renderRichTextParagraphs()` / `renderRun()` to append the same canonical click. Enforce `(hyperlink === undefined) === (relationshipId === undefined)`. Do not invoke `ShapeModel.hyperlink` after creation and do not create separate run relationships.

- [ ] **Step 6: Prove whole-shape editor ownership with shared run references**

On a created text hyperlink:

- same whole-shape assignment is exact bytes/journal no-op;
- tooltip-only edit keeps the shared relationship;
- target replacement clone-on-writes `cNvPr` and leaves all run IDs/target untouched;
- whole-shape clear removes only `cNvPr` click and keeps the run relationship;
- fill/line/arrows/shadow edits preserve both layers.

- [ ] **Step 7: Run model regression gates**

```sh
node_modules/.bin/vitest run packages/model/src/model.test.ts packages/model/src/rich-text.internal.test.ts packages/model/src/shape-hyperlink.internal.test.ts packages/model/src/shape-shadow.internal.test.ts packages/model/src/shape-arrows.internal.test.ts packages/model/src/simple-line.internal.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 8: Review, commit, push, and verify**

Review validation-before-mutation, relationship pairing, owner part selection, target identity, namespace binding, DrawingML child order, empty/multiline behavior, explicit underline precedence, omitted-byte stability, and shared-reference ownership. Stage only the four files, commit `feat: create text shape hyperlinks`, push, fetch, and require divergence `0 0`.

---

### Task 2: Prove placeholder, master/layout, root-package, and lifecycle coverage

**Files:**
- Modify: `packages/sdk/src/master-layout.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptx/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddTextOptions.hyperlink`, `AddPlaceholderOptions`, `SlideMasterObject`, layout/master wrappers, and `PptxDocument` lifecycle APIs.
- Produces: declarative option acceptance plus end-to-end public coverage for every text owner and presentation format.

- [ ] **Step 1: Add failing SDK surface tests**

Create external/internal/self links across slide, layout, master, rich text, placeholders, placeholder population, and declarative named layouts. Require:

- relationship ownership by the correct slide/layout/master part;
- layout/master source isolation when populating a slide placeholder;
- stable wrapper and shape identity;
- duplicate external/other-slide/self-link behavior;
- move/insert/delete-before-target ordinal updates;
- target deletion cleanup for both `cNvPr` and `rPr` references;
- outer transaction rollback, asynchronous declarative detachment, write/reopen, and all `PRESENTATION_FORMAT_PROFILES` formats.

- [ ] **Step 2: Add strict declarative and root-package tests**

Add `'hyperlink'` to `TEXT_OPTION_KEYS`. Pass invalid text/placeholder links through `defineSlideMaster()` with and without delayed image resolution and require rejection before observable mutation.

In `packages/pptx/src/index.test.ts`, import `AddTextOptions` and `Hyperlink`, compile both branches, create/reopen them from `@jiayunxie/pptx`, and add exact negatives:

```ts
const typedTextLink: Hyperlink = { slide: 2, tooltip: '' };
const typedTextOptions: AddTextOptions = { hyperlink: typedTextLink };
// @ts-expect-error exactly one target is required
const invalidBoth: AddTextOptions = { hyperlink: { url: 'https://example.com', slide: 2 } };
// @ts-expect-error target aliases are intentionally unsupported
const invalidAlias: AddTextOptions = { hyperlink: { target: 'https://example.com' } };
```

- [ ] **Step 3: Run focused tests and confirm intended failure**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape hyperlink|text hyperlink" --reporter=dot
```

Expected: declarative definition rejects `hyperlink` until the closed key set is extended; other surfaces exercise Task 1.

- [ ] **Step 4: Implement declarative acceptance only**

Add exactly one key to `TEXT_OPTION_KEYS`:

```ts
'hyperlink',
```

Do not add custom cloning. `readOptions()` already detaches nested data before asynchronous preparation, and Task 1 performs strict semantic normalization before package mutation.

- [ ] **Step 5: Run SDK/root/lifecycle gates**

```sh
node_modules/.bin/vitest run packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts --reporter=dot
node_modules/.bin/tsc -b --pretty false
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review owner relationship placement, placeholder source isolation, target identity through duplicate/move/delete, asynchronous detachment, compile-time exclusions, six-format persistence, and no unrelated key expansion. Commit `test: cover public text shape hyperlinks`, push, fetch, and require divergence `0 0`.

---

### Task 3: Lock PptxGenJS 4.0.1 public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public PptxGenJS `addText()` / `write()`, raw imported package parts, native `PptxDocument`, and live `ShapeModel.hyperlink`.
- Produces: exact supported-case and intentional-divergence evidence for text-shape hyperlink creation.

- [ ] **Step 1: Add a public-only PptxGenJS fixture**

Create plain URL/internal/self links, tooltip omitted/empty/custom, XML metacharacters, multiline text, explicit color/underline, rich outer options, valid heterogeneous rich-run links, and invalid target/coercion cases using public PptxGenJS only.

Inspect exact shape XML and relationships. Lock the valid baseline:

- plain outer options write both `p:cNvPr` and all `a:rPr` clicks;
- all references share one relationship ID;
- internal action/type/target and external mode/target are exact;
- omitted tooltip becomes direct empty and linked runs receive underline;
- rich outer options produce `rIdundefined` with no relationship in 4.0.1;
- per-run rich links produce one valid relationship per linked run and no whole-shape click.

- [ ] **Step 2: Add native comparison and strict-divergence assertions**

Create supported native equivalents. Compare name/text/transform, number and location of click elements, shared IDs, target/action/type/mode, tooltip visible semantics, and default underline. Require native rich outer options to produce a valid uniform shared link; require invalid/both/out-of-range/coercible values to reject with exact zero mutation.

Do not import PptxGenJS's malformed rich-outer deck as a valid semantic fixture. Record it as raw evidence only.

- [ ] **Step 3: Run focused conformance gates**

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "text shape hyperlink" --reporter=dot
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptx/src/index.test.ts -t "text shape hyperlink|text hyperlink" --reporter=dot
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

Review public-only PptxGenJS generation, exact raw evidence, supported-vs-defect separation, strict native invalid semantics, and absence of adapter production changes. Commit `test: compare text shape hyperlinks with pptxgenjs`, push, fetch, and require divergence `0 0`.

---

### Task 4: Prove actual packed Node, browser, declaration, and CLI behavior

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**
- Consumes: actual `@jiayunxie/pptx` tarball, Node/browser conditional exports, generated declarations, and installed `pptx-inspect` CLI.
- Produces: `textShapeHyperlinks: true` in Node and real-Chrome smoke output plus a validated packed deck.

- [ ] **Step 1: Extend packed Node and declaration smoke**

Using only the installed tarball, create plain/rich/placeholder/layout/master/declarative text with URL/internal/self/tooltip/multiline and combined fill+line+arrows+shadow+hyperlink. Assert immediate detached snapshots, mutate caller input, duplicate, move, whole-shape replace/clear with shared run references, write/reopen, and save `text-shape-hyperlinks-smoke.pptx`. Add `textShapeHyperlinks` to final JSON.

In the generated TypeScript consumer, import `AddTextOptions` and `Hyperlink`, compile both valid branches, and add `@ts-expect-error` cases matching Task 2.

- [ ] **Step 2: Extend real-Chrome smoke**

In `page.evaluate`, create/reopen a Blob-backed deck with external/internal, plain/rich, placeholder, and combined-style links. Return exact immediate/detached/reopen snapshots plus validation count and `textShapeHyperlinks: true`. Require zero console, page, and network errors.

- [ ] **Step 3: Add installed CLI validation**

Run installed `pptx-inspect package inspect`, `package validate --profile powerpoint-2010`, `slides list`, and `part read` on the packed smoke deck. Require no package errors, expected portability diagnostics only for external URLs, valid internal/self relationships, and raw dual shape/run click evidence. Also validate an internal-only deck at 0 errors / 0 warnings.

- [ ] **Step 4: Build and pack without repository artifacts**

```sh
cd packages/pptx
../../node_modules/.bin/tsup
../../node_modules/.bin/tsup --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
text_link_pack_dir="$(mktemp -d /tmp/pptx-text-link-pack.XXXXXX)"
npm pack --ignore-scripts --pack-destination "$text_link_pack_dir"
cd ../..
node scripts/smoke-npm-package.mjs "$text_link_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require exactly 57 tarball files unless a reviewed declaration closure change explains a new count. The installed smoke must reject workspace protocol/internal runtime imports and report `textShapeHyperlinks: true`.

- [ ] **Step 5: Run type/build gates**

```sh
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review actual-tarball isolation, Node/browser parity, declaration negatives, external-warning expectations, internal-only zero-warning proof, existing smoke fields, and absence of temporary repository artifacts. Commit `test: verify packed text shape hyperlinks`, push, fetch, and require divergence `0 0`.

---

### Task 5: Close documentation and full release gates

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: implementation, conformance, packed, browser, validator, and final-suite evidence from Tasks 1–4.
- Produces: accurate public support language and selection of the next smallest unsupported PptxGenJS parity item.

- [ ] **Step 1: Update public API examples and contracts**

Add URL/internal/tooltip/placeholder examples, strict `Hyperlink` rules, uniform plain/rich behavior, single shared relationship semantics, default underline, lifecycle behavior, and layout/master/declarative coverage. State clearly that `ShapeModel.hyperlink` edits only the whole-shape click and that per-run rich-text links remain a separate gap.

- [ ] **Step 2: Update compatibility and progress ledgers**

Move outer text-shape hyperlink creation from the `slide.addText` gap column into supported behavior. Record PptxGenJS dual output, omitted-tooltip collapse, rich-outer `rIdundefined`, permissive/dangling invalid behavior, native strictness, and final test/packed/browser/CLI results. Retain per-run hyperlink read/edit, heterogeneous rich-run creation, hover, and non-text-owner gaps. Select the next small item from the remaining parity ledger.

- [ ] **Step 3: Run stale-language and Markdown checks**

Search all six documents for statements that still list uniform text hyperlink creation as unsupported or next:

```sh
git diff --check
rg -n -i "text[- ]shape hyperlink.*(unsupported|pending|未支持)|next.*text[- ]shape hyperlink|下一.*text[- ]shape hyperlink" README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md CHANGELOG.md
```

- [ ] **Step 4: Run final release gates**

```sh
node_modules/.bin/vitest run
RUN_PERF=1 node_modules/.bin/vitest run packages/testkit/src/performance.test.ts
node_modules/.bin/tsc -b --pretty false
node_modules/.bin/tsc -b
```

Rerun both tsup builds, declaration closure, final actual tarball smoke, real Chromium smoke, and installed CLI validation. Require `textShapeHyperlinks: true`, Chromium zero console/page/network errors, internal-only PowerPoint 2010 validation 0 errors / 0 warnings, and only the documented portability diagnostics for external links.

- [ ] **Step 5: Review, commit, push, and verify**

Review every claim against current command output and raw OOXML evidence. Stage only the six documentation files, commit `docs: document text shape hyperlink support`, push, fetch, require divergence `0 0`, and report completed item, remaining advanced-text gaps, updated progress, and the next selected item.
