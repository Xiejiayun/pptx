# Slide and Section Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns all
> repository edits, review, commit, and push. Steps use checkbox (`- [ ]`) syntax
> for tracking.

**Goal:** Close exactly 26 PptxGenJS slide and section declaration atoms by
reusing the existing native lifecycle and adding one aggregate upstream runtime
control.

**Architecture:** Frozen status groups generate the exact manifest family. One
adapter test locks callable ownership, method returns, caller mutation,
fallbacks, alias precedence, and auto-page behavior; existing native, package,
browser, and OOXML probes provide the positive lifecycle evidence. Focused
gates run for this family and expensive artifact gates run once after all three
lifecycle families are committed.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, stable JSON
and Markdown audit artifacts, npm pack, persistent Chrome, and `pptx-inspect`.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-04-slide-section-surface-family-design.md` exactly.
- Close only the 26 target atoms and require the exact 10/15/1 classification.
- Reuse existing slide, section, notes, hidden, background, color, slide-number,
  auto-page, master, package, browser, and OOXML evidence; add only the missing
  aggregate PptxGenJS control and matrix assertions.
- Run focused tests and TypeScript once before review; defer npm pack, Chrome,
  OOXML, and full Vitest to the three-family lifecycle batch boundary.
- Do not stage `.pnpm-store/` or temporary package, browser, or PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the upstream slide and section boundaries

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 slide methods, continuation projection, master and
  section selectors, section options, and native `PptxDocument` equivalents.
- Produces: the stable title `locks the slide and section declarations against
  PptxGenJS 4.0.1`.

- [x] **Step 1: Lock the exact 26-ID inventory**

Build three explicit ID groups and assert 10 supported, 15 deliberate
differences, one deprecated alias, 26 unique IDs, and no nested option atom.

- [x] **Step 2: Lock methods and callable ownership**

Require all five upstream slide methods to return their owning slide, native
notes to return the slide, native image/shape/text/media operations to return
typed models, six `PresSlide` callables to exist, generic native media to be
absent, and PptxGenJS's internal master function slots to be null.

- [x] **Step 3: Lock state, mutation, and fallback differences**

Require caller hyperlink/background mutation, `bkgd` precedence, auto-page
continuation reuse, legal hidden state, known and unknown master/section paths,
positive section order, and the `order: 0` divergence.

- [x] **Step 4: Run the aggregate control once**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks the slide and section declarations against PptxGenJS 4.0.1"
```

Expected: one aggregate control passes without retaining temporary artifacts.

### Task 2: Generate and verify the 26 manifest entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the aggregate control plus existing property-specific adapter,
  native, package, browser, and OOXML anchors.
- Produces: exactly 26 new manifest entries and global counts
  616/269/87/355/447.

- [x] **Step 1: Generate entries from frozen status groups**

Define exact IDs, native mappings, family evidence, status-specific controls,
serialization/client flags, notes, and the `Slide.bkgd` canonical target.

- [x] **Step 2: Lock exact family and global counts**

Assert 26 distinct IDs, 1,327 manifest entries, exact 10/15/1 family counts,
the one alias mapping, and global totals 616 supported, 269
deliberate-difference, 87 deprecated-alias, 355 defect-excluded, and 447
unverified.

- [x] **Step 3: Regenerate and inspect the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: deterministic artifacts, zero stale/unsupported/diagnostic entries,
and 447 unverified atoms.

### Task 3: Verify, review, commit, push, and synchronize

**Files:**

- Review and stage only the seven capability-family files from Tasks 1-2 plus
  this design and plan.

**Interfaces:**

- Consumes: the verified diff and programmatic matrix comparison against HEAD.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Run family gates once**

Run the focused adapter controls for section, hidden, notes, backgrounds,
default colors, slide numbers, auto-page, and this aggregate; run the audit
tests and both TypeScript checks once.

- [x] **Step 2: Review exact migration**

Require exactly 26 changed atoms, all from `unverified`, exact 10/15/1
transitions, no other atom content changes, no unrelated staged paths, and no
`.pnpm-store/` files. Run `git diff --check`.

- [x] **Step 3: Commit and push**

```bash
git commit -m "docs: close slide section surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
