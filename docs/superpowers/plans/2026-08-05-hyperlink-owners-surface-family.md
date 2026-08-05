# Hyperlink Owners Surface Family Implementation Plan

> **For agentic workers:** Evidence agents write only to `/tmp`; the main
> agent owns repository edits, review, matrix generation, commit, and push.

**Goal:** Close six PptxGenJS hyperlink-owner atoms by adding strict raster and
SVG image hyperlink creation/editing and validating the shared owner model.

**Architecture:** Reuse the existing shape hyperlink codec from `ImageModel`
and bind normalized image hyperlinks inside the current image transaction.
Detach SDK inputs before async work and batch runtime, package, browser, OOXML,
and matrix evidence once for the family.

**Tech Stack:** TypeScript, Vitest, lossless OOXML, PptxGenJS 4.0.1, npm pack,
persistent Chrome, and `pptx-inspect`.

## Global Constraints

- Close exactly six atoms as five supported / one deliberate difference.
- Add no image-specific hyperlink codec and do not expose hyperlink on unrelated owners.
- Reject invalid internal targets before any media or relationship mutation.
- Evidence agents do not modify repository files.
- Run focused tests once and expensive npm/Chrome/full-suite gates once.
- Review, commit, and push the complete capability family as one unit.

---

### Task 1: Add failing model and SDK contracts

**Files:**
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

- [ ] Add `creates, edits, duplicates, rolls back, and reopens image hyperlinks`.
- [ ] Cover raster/SVG URL, internal and self targets, tooltip intent, editing,
  no-op, relationship reuse/COW/GC, duplicate/move/delete/rollback, six formats,
  write/reopen, and unchanged image appearance/media state.
- [ ] Add `creates raster and SVG image hyperlinks through the public SDK surface`.
- [ ] Prove synchronous caller detachment and strict invalid-input zero mutation.
- [ ] Run only those two titles and require the missing public surface to fail.

### Task 2: Implement the minimal shared image owner

**Files:**
- Modify: `packages/model/src/image.ts`
- Modify: `packages/model/src/image-create.internal.ts`
- Modify: `packages/model/src/svg-image-create.internal.ts`
- Modify: `packages/model/src/shapes.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/sdk/src/raster-image-source.ts`

- [ ] Add typed raster/SVG `hyperlink` options and normalize them with
  `normalizeHyperlink()` into frozen detached state.
- [ ] Pass the field through the SVG whitelist and SDK model-option whitelist;
  call `detachDataObject()` before any source await.
- [ ] Add only `ImageModel.hyperlink`, delegating to the slide shared reader/editor.
- [ ] Preflight internal targets before image allocation and bind the link within
  the current raster/SVG outer transaction.
- [ ] Run the two focused titles and TypeScript; require both to pass.

### Task 3: Lock upstream control and shared artifact proof

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Create: `scripts/hyperlink-owners-6-lifecycle-probe.mjs`
- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`

- [ ] Add `closes PptxGenJS hyperlink owners through shared strict native state`.
- [ ] Assert the exact six IDs, the 5/1 split, upstream valid output, and the
  rich outer-link `rIdundefined` deliberate difference.
- [ ] Reuse one lifecycle probe for workspace, installed tarball, and browser
  exports and retain both package decks.
- [ ] Validate exact picture `cNvPr/a:hlinkClick` ownership, relationship targets,
  unchanged payload/appearance parts, and PowerPoint 2010 diagnostics.

### Task 4: Backfill the matrix and documentation

**Files:**
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.json`
- Regenerate: `docs/compatibility/pptxgenjs-surface-audit.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`

- [ ] Register exact code/test/package/OOXML/client/control evidence for all six atoms.
- [ ] Lock 1,740 classified, 34 unverified, 98.08%, and zero unsupported,
  stale, or diagnostic entries.
- [ ] Regenerate the audit twice and require byte-identical artifacts.

### Task 5: Batch review, commit, and synchronize

- [ ] Run TypeScript, the three focused titles, full Vitest, npm pack, persistent
  Chrome, tarball CLI inspection, PowerPoint 2010 validation, and audit 8/8 once.
- [ ] Complete Agent A/B/C and main-agent review with zero findings.
- [ ] Stage only this family, exclude `.pnpm-store/`, commit, push `main`, fetch,
  and require local/remote divergence `0 0`.
