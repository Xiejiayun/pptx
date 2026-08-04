# Text Box Fit Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> all shared-file edits, review, commit, and push. Steps use checkbox
> (`- [ ]`) syntax for tracking.

**Goal:** Close exactly six PptxGenJS text-box fit atoms with direct native,
runtime/control, package, browser, and OOXML evidence.

**Architecture:** Generate one frozen manifest family for the modern fit union
and two deprecated aliases. Reuse the existing native and PptxGenJS aggregate
tests, add one real-Chrome lifecycle state, regenerate deterministic audit
artifacts, and deliver the family as one reviewed commit.

**Tech Stack:** Node.js 22 ESM, TypeScript 5.9, Vitest 3.2, PptxGenJS 4.0.1,
Playwright/Chrome, stable audit JSON and Markdown.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-05-text-box-fit-surface-family-design.md`
  exactly.
- Close only the six fit atoms in this batch.
- Classify modern `fit` plus its three tokens as supported.
- Classify `autoFit` and `shrinkText` as deprecated aliases of `fit`.
- Evidence agents never modify shared files.
- The focused tests have already run once for this family; do not rerun them.
  Run package, Chrome, OOXML, and full gates once at the batch boundary.
- Never stage `.pnpm-store/` or temporary PPTX/package/browser artifacts.
- Finish with review, one family commit, push, fetch, and divergence `0 0`.

---

### Task 1: Lock the exact matrix family

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`

- [x] **Step 1: Add the exact six entries and evidence mappings**

Generate four supported entries and two deprecated-alias entries. Require the
aliases to point to `interface:TextPropsOptions@property:fit`.

- [x] **Step 2: Lock family and global totals**

Assert the exact IDs/statuses, a 4/2 status split, canonical alias targets,
1,430 classified entries, and 624/354/93/359 totals.

### Task 2: Add the missing real-Chrome aggregate

**Files:**

- Modify: `scripts/playwright-browser-smoke.js`

- [x] **Step 1: Cover create, OOXML, edit, and failure isolation once**

Create omitted/none/shrink/resize shapes, assert exact fit children, reject an
invalid token without mutation, duplicate, and edit all modes.

- [x] **Step 2: Cover browser write/reopen once**

Use one `writeBlob()` and reopen cycle to assert edited state, duplicate
isolation, PPTX MIME, and zero warnings/errors.

### Task 3: Run the focused family gate and regenerate artifacts

**Files:**

- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

- [x] **Step 1: Run focused evidence tests once**

The read-only artifact line ran:

```sh
node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts packages/sdk/src/index.test.ts -t "imports public PptxGenJS output and continues editing|creates, edits, duplicates, and reopens text-box fit modes|rejects malformed text-box fit"
```

Result: 2 files passed, 3 tests passed, 0 failed, and 361 skipped. The matrix
and syntax gates run after the main-agent edits below; the focused Vitest gate
is not repeated.

- [x] **Step 2: Regenerate deterministic artifacts**

Run `node scripts/pptxgenjs-surface-audit.mjs --write` twice and require a
clean second-generation diff plus `unverified=344`.

- [x] **Step 3: Run TypeScript and the shared artifact batch once**

Run the root TypeScript build, one actual npm tarball smoke, and the repository
browser-smoke command that loads `scripts/playwright-browser-smoke.js` in one
persistent Chrome session. Require `textBoxFitFamily=true` and zero console,
page, network, or validation errors. Retain `browser-text-box-fit.pptx`, inspect
its package and exact slide part with `pptx-inspect`, and require the
PowerPoint 2010 profile to report zero errors and zero warnings.

- [x] **Step 4: Run the full suite once**

Run the repository full Vitest command once and require no failures.

### Task 4: Review, commit, and synchronize

- [x] **Step 1: Review exact migration and worktree scope**

Compare the previous and generated audit JSON by atom ID. Require exactly four
`unverified -> supported` and two `unverified -> deprecated-alias`
transitions, with no other status or metadata regressions; run
`git diff --check`.

- [x] **Step 2: Commit and push the family**

Stage only the two docs, manifest/test, browser smoke, and generated audit
artifacts. Commit as `docs: close text box fit surface family`, push `main`,
fetch `origin/main`, and require `git rev-list --left-right --count
HEAD...origin/main` to print `0 0`.
