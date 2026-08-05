# Data/Path Inheritance Surface Family Implementation Plan

> **For agentic workers:** Evidence agents write only to `/tmp`; the main
> agent owns repository edits, review, matrix generation, commit, and push.

**Goal:** Close four inherited data/path declaration atoms with one aggregate
runtime control and evidence-backed matrix update.

**Architecture:** Preserve native typed image/media sources and reject inert
text aliases. Reuse existing package/browser source-owner gates because no
packed implementation changes.

**Tech Stack:** TypeScript, Vitest, PptxGenJS 4.0.1, lossless OOXML, and the
surface-audit generator.

## Global Constraints

- Close exactly four atoms as two deliberate differences / two defect exclusions.
- Do not modify product implementation or add `data`/`path` text aliases.
- Evidence agents do not modify repository files.
- Run focused tests once and reuse the latest npm/browser batch evidence.
- Review, commit, and push the complete capability family as one unit.

---

### Task 1: Add the aggregate runtime control

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`

- [x] Lock the exact four IDs and status split.
- [x] Verify image/media data-only, path-only, and data-first behavior.
- [x] Verify the image-only losing-path description leak.
- [x] Verify plain/rich text byte isolation and package sentinel absence.
- [x] Verify native typed-source write/reopen state without aliases.

### Task 2: Backfill the matrix

**Files:**
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.json`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.md`

- [x] Register exact evidence and status rules for all four atoms.
- [x] Lock 1,744 classified, 30 unverified, and zero diagnostics.
- [x] Regenerate twice and require byte-identical artifacts.

### Task 3: Document, review, and synchronize

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/implementation-progress.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`

- [x] Run the aggregate and four existing source-owner focused controls once.
- [x] Run TypeScript, audit tests, and diff checks.
- [x] Complete Agent A/B/C evidence review and main-agent review.
- [ ] Stage only this family, exclude `.pnpm-store/`, commit, push, and verify
  local/remote divergence `0 0`.
