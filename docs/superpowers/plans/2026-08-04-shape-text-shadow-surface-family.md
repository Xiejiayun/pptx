# Shape and Text Shadow Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> every repository edit, review, commit, and push. Steps use checkbox (`- [ ]`)
> syntax for tracking.

**Goal:** Close exactly 12 PptxGenJS shape/text simple-shadow declaration atoms
by reusing the existing runtime, native, package, browser, and OOXML evidence.

**Architecture:** A frozen ID group produces 12 deliberate-difference entries.
Existing owner-specific adapter controls plus one cross-owner aggregate
control provide the upstream runtime evidence; no production behavior changes.

**Tech Stack:** Node ESM, Vitest 3, PptxGenJS 4.0.1, TypeScript 5.8, stable JSON
and Markdown audit artifacts.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-04-shape-text-shadow-surface-family-design.md`
  exactly.
- Close only the 12 target atoms with the exact 0/12/0/0 status split.
- Reuse the existing shape and text shadow runtime controls and release
  evidence; add only the missing cross-owner control and do not change runtime
  shadow behavior.
- Run focused tests once for the family. Keep npm pack, persistent Chrome, and
  full Vitest at the multi-family batch boundary.
- Do not stage `.pnpm-store/` or temporary package/browser/PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the shared declaration across owners

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 shape, text, image, and chart shadow writers.
- Produces: one stable aggregate control title for all 12 atoms.

- [x] **Step 1: Lock the exact atom inventory**

Assert the seven properties, three union members, and two owner properties are
12 unique IDs and exclude `ImageProps.shadow` from this family boundary.

- [x] **Step 2: Lock cross-owner behavior**

Require shape/text outer and none behavior from the existing controls, legal
image/chart inner output, chart rotate behavior, and the known shape/text
inner and rotate differences. Do not create persistent fixture files.

### Task 2: Generate the exact shadow family entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`

**Interfaces:**

- Consumes: the two existing adapter control titles and existing native,
  package, OOXML, and browser anchors.
- Produces: `SHAPE_TEXT_SHADOW_FAMILY_ENTRIES` containing exactly 12 entries.

- [x] **Step 1: Define exact status groups**

Create one frozen deliberate-difference ID group with exactly 12 unique IDs.

- [x] **Step 2: Map owner-aware evidence**

Map `ShapeProps.shadow` to the shape control, `TextPropsOptions.shadow` to the
text control, and shared `ShadowProps` atoms to the cross-owner control. Map
all entries to `ShapeShadow`/`ShapeModel.shadow`, packed shadow states, SDK
OOXML tests, and the persistent browser state.

### Task 3: Lock migration and regenerate the matrix

**Files:**

- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: `SHAPE_TEXT_SHADOW_FAMILY_ENTRIES`.
- Produces: exact family/global count assertions and deterministic artifacts.

- [x] **Step 1: Add exact family assertions**

Assert 12 unique IDs, all with deliberate-difference status, and no
image-shadow atom.

- [x] **Step 2: Lock global totals**

Assert 1,373 closed entries and totals 620/303/91/359/401 with zero
unsupported and stale entries.

- [x] **Step 3: Regenerate deterministically**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: generated JSON/Markdown are deterministic, diagnostics remain empty,
and exactly 401 atoms remain unverified.

### Task 4: Focused verification, review, commit, and push

**Files:**

- Review and stage only the manifest, audit test, generated matrix, design,
  and plan files for this family.

**Interfaces:**

- Consumes: the verified matrix diff and existing shadow controls.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Run focused controls once**

```bash
./node_modules/.bin/vitest run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  packages/model/src/simple-shadow.internal.test.ts \
  packages/model/src/shape-shadow.internal.test.ts \
  packages/sdk/src/index.test.ts \
  -t "shadow"
```

Expected: all selected shadow tests pass.

- [x] **Step 2: Review exact migration**

Compare generated atoms against HEAD. Require exactly 12 transitions from
`unverified`, all 12 to deliberate-difference, no other atom content change, no
unrelated staged path, and clean `git diff --check`.

- [x] **Step 3: Commit and synchronize**

```bash
git commit -m "docs: close shape text shadow surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
