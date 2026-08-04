# Master, Background, and Slide-Number Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns all
> repository edits, review, commit, and push. Steps use checkbox (`- [ ]`) syntax
> for tracking.

**Goal:** Close exactly 34 PptxGenJS master, background, and slide-number
declaration atoms with one aggregate runtime control and reused native,
package, browser, and OOXML evidence.

**Architecture:** Frozen status groups generate the exact manifest family. One
adapter test locks declared ownership, valid branches, aliases, defects,
mutation, defaults, and strict native differences. Focused gates run for this
family; expensive artifact gates run once at the lifecycle batch boundary.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, stable JSON
and Markdown audit artifacts, npm pack, persistent Chrome, and `pptx-inspect`.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-04-master-background-slide-number-surface-family-design.md`
  exactly.
- Close only the 34 target atoms and require the exact 4/22/4/4
  classification.
- Reuse existing master, background, slide-number, package, browser, and OOXML
  evidence; add only the missing aggregate PptxGenJS control and matrix
  assertions.
- Run focused tests and TypeScript once before review; defer npm pack, Chrome,
  OOXML, and full Vitest to the lifecycle batch boundary.
- Do not stage `.pnpm-store/` or temporary package, browser, or PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the upstream master, background, and slide-number boundaries

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 master definitions, `BackgroundProps`,
  `SlideNumberProps`, and native strict equivalents.
- Produces: the stable title `locks master, background, and slide-number
  declarations against PptxGenJS 4.0.1`.

- [x] **Step 1: Lock the exact 34-ID inventory**

Build four explicit ID groups and assert four supported, 22 deliberate
differences, four deprecated aliases, four defect exclusions, 34 unique IDs,
and no nested foreign-family member.

- [x] **Step 2: Lock master branches and aliases**

Require canonical color/data backgrounds and the legacy string background to
write, every `bkgd: BackgroundProps` branch to remain inert, the declared chart
shape to fail, all other master-object branches to write in order, and native
strict objects to return an editable live layout.

- [x] **Step 3: Lock direct background behavior**

Require caller identity, `fill` and `alpha` aliases, caller mutation, color,
transparency, data/path images, inert `src`, malformed PptxGenJS none output,
valid solid output, and native explicit no-fill plus detached strict state.

- [x] **Step 4: Lock slide-number behavior**

Require caller identity, zero geometry fallbacks, four horizontal alignments,
three vertical alignments, zero/scalar/tuple margins, ignored top-level
transparency, and native detached state with explicit zero and nested style
transparency.

- [x] **Step 5: Run the aggregate control once**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks master, background, and slide-number declarations against PptxGenJS 4.0.1"
```

Expected: one aggregate control passes without retaining temporary artifacts.

### Task 2: Generate and verify the 34 manifest entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the aggregate control plus existing native, package, browser, and
  OOXML anchors.
- Produces: exactly 34 new manifest entries and global counts
  620/291/91/359/413.

- [x] **Step 1: Generate entries from frozen status groups**

Define exact IDs, native mappings, evidence categories, status-specific
controls, serialization/client flags, notes, alias canonical targets, and
empty native mappings for the four defective atoms.

- [x] **Step 2: Lock exact family and global counts**

Assert 34 distinct IDs, 1,361 manifest entries, exact 4/22/4/4 family counts,
the four alias mappings, the four defect IDs, and global totals 620 supported,
291 deliberate-difference, 91 deprecated-alias, 359 defect-excluded, and 413
unverified.

- [x] **Step 3: Regenerate and inspect the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: deterministic artifacts, zero stale/unsupported/diagnostic entries,
and 413 unverified atoms.

### Task 3: Verify, review, commit, push, and synchronize

**Files:**

- Review and stage only the seven capability-family files from Tasks 1-2 plus
  this design and plan.

**Interfaces:**

- Consumes: the verified diff and programmatic matrix comparison against HEAD.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Run family gates once**

Run the aggregate control plus the focused existing master, background, and
slide-number controls; run the audit tests and both TypeScript checks once.

- [x] **Step 2: Review exact migration**

Require exactly 34 changed atoms, all from `unverified`, exact 4/22/4/4
transitions, no other atom content changes, no unrelated staged paths, and no
`.pnpm-store/` files. Run `git diff --check`.

- [x] **Step 3: Commit and push**

```bash
git commit -m "docs: close master background slide number surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
