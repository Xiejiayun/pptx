# Shape Custom Path Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> every repository edit, review, commit, and push. Steps use checkbox (`- [ ]`)
> syntax for tracking.

**Goal:** Close exactly fifteen PptxGenJS `ShapeProps.points` custom-path atoms
using the existing native custom-geometry vertical slice plus one real-Chrome
aggregate gate.

**Architecture:** One frozen manifest group maps the points bag to the native
custom-geometry command union. Existing PptxGenJS controls, native lifecycle,
and tarball tests provide most evidence; the browser smoke adds all six command
kinds once, and one audit assertion locks scope and totals.

**Tech Stack:** TypeScript 5.8, Node ESM, Vitest 3, PptxGenJS 4.0.1,
`pptx-inspect`, stable JSON and Markdown audit artifacts.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-05-shape-custom-path-surface-family-design.md`
  exactly.
- Close only the fifteen target atoms, all as deliberate differences.
- Keep shared `Coord`, transforms, preset shortcuts, formulas, handles, sites,
  rectangles, evaluator state, and path appearance outside this batch.
- Evidence agents do not modify repository files.
- Run focused tests once. Run npm pack, persistent Chrome, OOXML inspection,
  and full Vitest once at this family boundary.
- Do not stage `.pnpm-store/` or temporary package/browser/PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the exact evidence-backed matrix entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`

- [x] **Step 1: Define the fifteen frozen IDs**

Add the `points` root, five direct inline fields, and nine curve fields as one
deliberate-difference family.

- [x] **Step 2: Map native and aggregate evidence**

Map each atom to the strict custom-geometry tree and existing PptxGenJS,
tarball, SDK, and browser anchors. Require control, serialization, and client
evidence on every entry.

- [x] **Step 3: Lock scope and totals**

Assert exact IDs, exclusions, 1,424 manifest entries, and totals
620/354/91/359/350.

### Task 2: Add the missing real-Chrome custom-path aggregate

**Files:**

- Modify: `scripts/playwright-browser-smoke.js`

- [x] **Step 1: Create all command kinds once**

Create one native custom path with ordered move, line, quadratic, cubic, arc,
second move/line, and close commands.

- [x] **Step 2: Prove live, XML, and reopen state**

Require a deep-frozen snapshot, writeBlob/reopen equality, exact command tag
counts/order, and zero error diagnostics. Return a stable
`browserCustomPathState` evidence object.

### Task 3: Regenerate and run the focused family gate once

**Files:**

- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

- [x] **Step 1: Regenerate and audit**

Run the audit writer and Node surface/runtime tests. Require exactly 350
unverified atoms and empty diagnostics.

- [x] **Step 2: Run focused tests once**

Run the custom-geometry model, SDK, adapter, and browser-smoke focused set in
one family gate.

- [x] **Step 3: Run TypeScript and exact migration review**

Require exactly fifteen `unverified` to `deliberate-difference` transitions and
no other status change.

### Task 4: Run shared artifact gates once

- [x] **Step 1: Run the actual npm tarball smoke once**

Require all-command/multi-path create, edit, convert, write, and reopen checks
from the installed package and browser-condition bundle.

- [x] **Step 2: Run persistent Chrome and inspect OOXML once**

Reuse one Chrome session, retain its PPTX, inspect the package and slide part,
and validate the six custom-path command tags and ordering.

- [x] **Step 3: Run the full suite once**

Run the repository full Vitest command and require no failures.

### Task 5: Review, commit, and synchronize

- [x] **Step 1: Complete evidence-line and main-agent review**

Require matrix scope, runtime classification, package/browser/OOXML validity,
generated artifacts, and `git diff --check` to pass with no unrelated tracked
path.

- [x] **Step 2: Commit and synchronize**

Create one `docs: close shape custom path surface family` commit, push `main`,
fetch `origin/main`, and require divergence `0 0`.
