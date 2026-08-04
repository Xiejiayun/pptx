# Rich-Text Effects Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> all shared-file edits, review, commit, and push. Steps use checkbox
> (`- [ ]`) syntax for tracking.

**Goal:** Close exactly sixteen PptxGenJS rich-text effect atoms with direct
native, runtime/control, package, browser, and OOXML evidence.

**Architecture:** Generate one frozen manifest family for nine supported
atoms and seven deliberate structured/native differences. Reuse existing
native, package, and PptxGenJS evidence, add one six-format SDK aggregate and
one real-Chrome lifecycle state, regenerate deterministic audit artifacts,
and deliver one reviewed family commit.

**Tech Stack:** Node.js 22 ESM, TypeScript 5.9, Vitest 3.2, PptxGenJS 4.0.1,
Playwright/Chrome, stable audit JSON and Markdown.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-05-rich-text-effects-surface-family-design.md`
  exactly.
- Close only the sixteen rich-text effect atoms in this batch.
- Classify size/opacity/effect owners, transparency, and strike string tokens
  as supported; structured colors, structured baseline/spacing names, and
  boolean strike as deliberate differences.
- Evidence agents never modify shared files.
- Run focused tests once for the family. Run npm, Chrome, OOXML, and full gates
  once at the batch boundary.
- Never stage `.pnpm-store/` or temporary PPTX/package/browser artifacts.
- Finish with review, one family commit, push, fetch, and divergence `0 0`.

---

### Task 1: Lock the combined native lifecycle

**Files:**

- Modify: `packages/sdk/src/index.test.ts`

- [x] **Step 1: Add one all-format rich-effect aggregate**

For each key in `PRESENTATION_FORMAT_PROFILES`, create one rich-text shape
containing outline, glow, baseline, character spacing, strike, and
transparency values. Assert immediate detached state and exact effect OOXML.

- [x] **Step 2: Prove edit and duplicate isolation**

Duplicate the source slide, replace the source run with alternate values, and
assert the duplicate retains the initial run-effect snapshot.

- [x] **Step 3: Prove write/reopen and diagnostics**

Write and reopen each format once. Assert the edited source, unchanged
duplicate, and zero package error/warning diagnostics.

### Task 2: Add one real-Chrome family aggregate

**Files:**

- Modify: `scripts/playwright-browser-smoke.js`

- [x] **Step 1: Cover all mappings in one browser document**

Create runs that combine sRGB and scheme outline/glow colors, zero and
non-zero sizes/opacities, named and numeric baselines, positive/negative/zero
character spacing, both strike tokens plus explicit false, and zero/fractional
transparency. Assert immediate model values and exact `a:rPr` OOXML.

- [x] **Step 2: Cover edits and failure isolation once**

Duplicate the slide, edit the source effect values, then attempt one malformed
effect mutation. Require unchanged part bytes and mutation journal after the
rejection.

- [x] **Step 3: Cover browser write/reopen once**

Use one `writeBlob()` and reopen cycle to assert edited state, duplicate
isolation, MIME, and zero error/warning diagnostics. Expose one
`richTextEffectsFamilyState` object.

### Task 3: Bind the exact matrix family

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

- [x] **Step 1: Add the exact sixteen manifest entries**

Generate nine supported and seven deliberate-difference entries. Bind each
entry to the existing implementation, PptxGenJS control, focused SDK OOXML
test, installed-package smoke, and unified Chrome state.

- [x] **Step 2: Lock family and global totals**

Assert exact IDs/statuses, a 9/7 status split, 1,456 classified entries, and
637/366/94/359 supported/difference/deprecated/defect totals.

- [x] **Step 3: Regenerate deterministic artifacts**

Run `node scripts/pptxgenjs-surface-audit.mjs --write` twice and require a
clean second generation with `unverified=318`.

### Task 4: Run gates and publish

- [x] **Step 1: Run focused evidence once**

Run the audit library, the two existing PptxGenJS controls, the six existing
native effect tests, and the new six-format aggregate.

- [x] **Step 2: Run shared batch validation once**

Run TypeScript, one actual npm tarball smoke, one real-Chrome gate, and
package/OOXML validation for the retained browser evidence deck. Require every
`richTextEffectsFamilyState` field to pass and zero browser or compatibility
diagnostics.

- [x] **Step 3: Run the full suite once**

Run the repository full Vitest command once and require no failures.

- [x] **Step 4: Review, commit, and synchronize**

Require exactly sixteen status transitions, run `git diff --check`, stage only
the family files, commit as `docs: close rich text effects surface family`,
push `main`, fetch `origin/main`, and require divergence `0 0`.
