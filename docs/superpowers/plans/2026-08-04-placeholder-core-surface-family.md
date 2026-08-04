# Placeholder Core Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> every repository edit, review, commit, and push. Steps use checkbox (`- [ ]`)
> syntax for tracking.

**Goal:** Close exactly eight PptxGenJS placeholder definition and population
atoms by reusing the existing runtime, native, package, browser, and OOXML
vertical slices.

**Architecture:** One frozen ID group produces eight deliberate-difference
entries. Existing controls plus one aggregate runtime test prove the observable
PptxGenJS/native boundary; shared package and browser states gain one geometry
check without adding another build or browser session.

**Tech Stack:** TypeScript 5.8, Node ESM, Vitest 3, PptxGenJS 4.0.1, stable JSON
and Markdown audit artifacts.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-04-placeholder-core-surface-family-design.md`
  exactly.
- Close only the eight target atoms, all as deliberate differences.
- Keep style-only placeholder properties and catalog union atoms outside the
  batch.
- Evidence agents do not modify repository files.
- Run focused tests once. Keep npm pack, persistent Chrome, OOXML package
  inspection, and full Vitest at the image-family batch boundary.
- Do not stage `.pnpm-store/` or temporary package/browser/PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the exact runtime boundary

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS master placeholder definitions and text/image population.
- Produces: one stable aggregate control title covering all eight atoms.

- [x] **Step 1: Lock definition identity and geometry**

Create a placeholder with distinct lookup/object names and percentage
`x/y/w/h`; compare the public output with native explicit-unit state and assert
the known `pic`/`tbl` type differences.

- [x] **Step 2: Lock population ownership**

Populate text and image owners using conflicting caller geometry. Require the
PptxGenJS output and native output to expose the known selector, identity, and
image-extent differences. Require native missing/ambiguous/domain failures to
leave the package unchanged.

### Task 2: Strengthen shared packed and browser states

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Consumes: existing `masterLayoutChecks` and `masterLayoutState` flows.
- Produces: one aggregate transform equality boolean in each existing state.

- [x] **Step 1: Add installed-package geometry state**

Compare reopened placeholder transforms to the original layout transforms for
the representative text and image owners. Do not create a new package build.

- [x] **Step 2: Add persistent-Chrome geometry state**

Return and assert the same representative transform equality through the
existing Blob/reopen browser flow. Do not create a second browser session.

### Task 3: Generate the exact matrix entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the three existing controls, the new aggregate, native lifecycle
  tests, and shared packed/browser state.
- Produces: exactly eight deliberate-difference entries.

- [x] **Step 1: Define ID and evidence mappings**

Create one frozen eight-ID group. Map definition fields to
`AddPlaceholderOptions`/`Transform`, text population to `AddTextOptions`, and
image population to `AddImageOptions`; attach owner-specific notes and controls.

- [x] **Step 2: Lock exact migration and totals**

Assert eight unique IDs, all deliberate differences, exact exclusions, 1,381
manifest entries, and totals 620/311/91/359/393.

- [x] **Step 3: Regenerate deterministically**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: generated JSON/Markdown are deterministic, diagnostics remain empty,
and exactly 393 atoms remain unverified.

### Task 4: Focused verification, review, commit, and push

**Files:**

- Review and stage only files owned by this capability family.

**Interfaces:**

- Consumes: the verified family diff.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Run focused controls once**

```bash
./node_modules/.bin/vitest run \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptx/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/validator/src/index.test.ts \
  -t "placeholder identity|empty layout placeholders|empty placeholder geometry|placeholder population|populate text shape placeholder|populate image chart placeholder|creates, selects, populates, and reopens placeholders|healthy master layout placeholder graph|placeholder core surface"
```

Expected: all selected placeholder controls pass.

- [x] **Step 2: Run TypeScript and exact migration review**

Run the repository TypeScript check, compare generated atoms against HEAD, and
require exactly eight `unverified` to `deliberate-difference` transitions with
no other atom content change. Require `git diff --check` and no unrelated
staged path.

- [x] **Step 3: Commit and synchronize**

```bash
git commit -m "docs: close placeholder core surface family"
git -c http.version=HTTP/1.1 push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
