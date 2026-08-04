# Text Paragraph/Layout Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> all shared-file edits, review, commit, and push. Steps use checkbox
> (`- [ ]`) syntax for tracking.

**Goal:** Close exactly ten PptxGenJS text paragraph/layout atoms with direct
native, runtime/control, package, browser, and OOXML evidence.

**Architecture:** Generate one frozen manifest family for four supported
properties, five deliberate structured/native differences, and the deprecated
`inset` alias. Reuse existing native, package, and PptxGenJS aggregate evidence,
add only the missing `inset` probes and one real-Chrome lifecycle state,
regenerate deterministic audit artifacts, and deliver one reviewed family
commit.

**Tech Stack:** Node.js 22 ESM, TypeScript 5.9, Vitest 3.2, PptxGenJS 4.0.1,
Playwright/Chrome, stable audit JSON and Markdown.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-05-text-paragraph-layout-surface-family-design.md`
  exactly.
- Close only the ten `TextPropsOptions` atoms in this batch.
- Classify align, RTL, vertical alignment, and wrap as supported; structured
  paragraph spacing and margins as deliberate differences; `inset` as a
  deprecated alias of `margin`.
- Evidence agents never modify shared files.
- Run focused tests once for the family. Run npm, Chrome, OOXML, and full gates
  once at the batch boundary.
- Never stage `.pnpm-store/` or temporary PPTX/package/browser artifacts.
- Finish with review, one family commit, push, fetch, and divergence `0 0`.

---

### Task 1: Lock the PptxGenJS control and exact matrix family

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`

- [x] **Step 1: Add only the missing `inset` runtime controls**

Extend the existing PptxGenJS aggregate with uniform `inset: 1`, zero inset,
`inset: 1` plus `margin: 10` precedence, and inert run-level inset. Assert
914400 EMU for one inch, zero for zero, 127000 EMU for the winning ten-point
margin, and no run-level override.

- [x] **Step 2: Add the exact ten manifest entries**

Generate four supported entries, five deliberate-difference entries, and one
deprecated-alias entry. Require the alias canonical target to be
`interface:TextPropsOptions@property:margin`.

- [x] **Step 3: Lock family and global totals**

Assert exact IDs/statuses, a 4/5/1 status split, 1,440 classified entries, and
628/359/94/359 supported/difference/deprecated/defect totals.

### Task 2: Add one real-Chrome family aggregate

**Files:**

- Modify: `scripts/playwright-browser-smoke.js`

- [x] **Step 1: Cover all canonical mappings in one document**

Create rich text with all four alignments, RTL true/false, exact and multiple
line spacing, before/after spacing, plus text boxes with scalar/tuple margins,
all vertical alignments, and both wrap values. Assert immediate model values
and exact `a:pPr`/`a:bodyPr` OOXML.

- [x] **Step 2: Cover edits and failure isolation once**

Duplicate the slide, edit paragraph and text-box values, then attempt one
invalid paragraph spacing and one invalid margin mutation. Require unchanged
bytes and mutation journal after each rejection.

- [x] **Step 3: Cover browser write/reopen once**

Use one `writeBlob()` and reopen cycle to assert edited state, duplicate
isolation, and zero error/warning diagnostics. Expose one
`textParagraphLayoutFamilyState` with immediate, OOXML, edited, reopened,
duplicate-isolation, rollback, and diagnostics fields.

### Task 3: Run the family and batch gates

**Files:**

- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

- [x] **Step 1: Run focused evidence tests once**

Run the audit library plus the single PptxGenJS aggregate and native SDK tests
for paragraph alignment, RTL, spacing, margins, vertical alignment, and wrap.

- [x] **Step 2: Regenerate deterministic artifacts**

Run `node scripts/pptxgenjs-surface-audit.mjs --write` twice and require a
clean second-generation diff plus `unverified=334`.

- [x] **Step 3: Run shared batch validation once**

Run TypeScript, one actual npm tarball smoke, the persistent real-Chrome gate,
and package/OOXML validation for the retained browser evidence deck. Require
`textParagraphLayoutFamily=true` and zero console, page, network, validation,
or compatibility-profile diagnostics.

- [x] **Step 4: Run the full suite once**

Run the repository full Vitest command once and require no failures.

### Task 4: Review, commit, and synchronize

- [x] **Step 1: Review exact migration and worktree scope**

Compare the previous and generated audit JSON by atom ID. Require exactly four
`unverified -> supported`, five `unverified -> deliberate-difference`, and one
`unverified -> deprecated-alias` transitions, with no other status or metadata
regressions; run `git diff --check`.

- [x] **Step 2: Commit and push the family**

Stage only the two design/plan docs, control, manifest/test, browser smoke, and
generated audit artifacts. Commit as
`docs: close text paragraph layout surface family`, push `main`, fetch
`origin/main`, and require `git rev-list --left-right --count
HEAD...origin/main` to print `0 0`.
