# Bullet And Numbering Surface Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the PptxGenJS 4.0.1 bullet and numbering capability family in one reviewed batch using reusable runtime, native, package, browser, and OOXML evidence.

**Architecture:** Three file-isolated workers produce the surface manifest, packed-package smoke, and real-Chrome smoke while the main worker owns focused native/control tests, evidence review, generated matrix files, and Git integration. One aggregate fixture covers all 16 numbering styles and every effective owner so a single lifecycle run can support many manifest atoms without rebuilding per atom.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, JSZip, a persistent Chromium browser, OOXML package inspection, npm pack.

## Global Constraints

- Treat the user's four-lane proposal as the approved design and execute without additional design questions.
- Close exactly the declared bullet/numbering owner and token atoms supported by the collected runtime evidence; do not infer behavior from TypeScript declarations alone.
- Keep PptxGenJS runtime defects as `defect-excluded`, intentional native API improvements as `deliberate-difference`, and working legacy names as `deprecated-alias`.
- Use one focused adapter run and one focused native SDK run for the family.
- Use one packed tarball build and one persistent real-Chrome run for the family.
- Each worker owns one file; only the main worker updates generated matrix/count files and performs Git operations.
- Do not stage `.pnpm-store/` or generated temporary packages.
- Finish with review, one family commit, push, fetch, and zero local/remote divergence.

---

### Task 1: Lock control and native lifecycle behavior

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS public output, `PptxDocument`, `ShapeModel`, `TableModel`, and six presentation format profiles.
- Produces: stable test titles referenced by surface-manifest evidence.

- [x] **Step 1: Add the PptxGenJS owner/control matrix**

Add `locks bullet and numbering behavior across every declared owner` and assert boolean/object bullets, canonical and legacy field names, all 16 numbering tokens, and inert owner output.

- [x] **Step 2: Run the focused adapter test**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks bullet and numbering behavior across every declared owner"
```

Expected: one matching test passes.

- [x] **Step 3: Add the six-format native lifecycle test**

Add `creates and reopens bullet and numbering owners in all six formats` with this owner set:

```ts
const owners = ['text', 'placeholder', 'table-cell'] as const;
const lifecycle = ['create', 'write', 'reopen', 'edit', 'clear', 'write', 'reopen'] as const;
```

The test must check all 16 `NumberingStyle` tokens, exact `a:buAutoNum`/`a:buChar` XML, and zero error diagnostics.

- [x] **Step 4: Run the focused SDK test**

```bash
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts \
  -t "creates and reopens bullet and numbering owners in all six formats"
```

Expected: one matching test passes.

---

### Task 2: Add one packed-package aggregate probe

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**

- Consumes: the installed tarball's ESM/CJS SDK exports and existing ZIP/XML helpers.
- Produces: a named bullet/numbering lifecycle check in the final smoke result.

- [x] **Step 1: Add the aggregate packed-package fixture**

Create text, placeholder, and table-cell owners with all numbering styles plus custom character, start value, and indent values. Verify the first snapshot and exact OOXML, then edit/clear, write, reopen, and verify the second snapshot.

- [x] **Step 2: Strengthen browser-condition package assertions**

Require the browser-condition import to return the expected bullet snapshot rather than merely accepting bullet input.

- [x] **Step 3: Expose the check and run one tarball smoke**

```bash
node --check scripts/smoke-npm-package.mjs
cd packages/pptx
npm pack --ignore-scripts --pack-destination /tmp/pptx-bullet-pack
cd ../..
node scripts/smoke-npm-package.mjs /tmp/pptx-bullet-pack/jiayunxie-pptx-0.1.0.tgz
```

Expected: the final `checks` object includes a truthy bullet/numbering lifecycle marker.

---

### Task 3: Add one persistent real-Chrome aggregate probe

**Files:**

- Modify: `scripts/playwright-browser-smoke.js`

**Interfaces:**

- Consumes: the browser bundle, persistent Playwright page, `writeBlob()`, and browser-side ZIP/XML inspection.
- Produces: a real-Chrome bullet/numbering lifecycle marker and diagnostic count.

- [x] **Step 1: Add the browser lifecycle**

In the existing page evaluation, create all 16 numbering styles across text, placeholder, and table-cell owners, add custom character/start/indent cases, and assert the initial public snapshots.

- [x] **Step 2: Verify OOXML and the second reopen**

Assert exact `a:buAutoNum` and `a:buChar` fragments, edit/clear the owners, run `writeBlob()`, reopen it, and require zero error diagnostics.

- [x] **Step 3: Run the persistent browser smoke once**

```bash
node --check scripts/playwright-browser-smoke.js
# Execute scripts/playwright-browser-smoke.js through the persistent browser harness.
```

Expected: the returned report includes the bullet/numbering marker and no browser errors.

---

### Task 4: Close the capability atoms and integrate evidence

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify generated compatibility/progress matrix files selected by `scripts/pptxgenjs-surface-audit.mjs --write`

**Interfaces:**

- Consumes: stable focused test titles, packed-smoke marker, Chrome marker, native code anchors, and runtime-control findings.
- Produces: 175 newly closed atoms and updated exact category totals.

- [x] **Step 1: Add the manifest entries**

Generate entries only from explicit owner/property/token lists. Include `interface:TextPropsOptions@property:indentLevel` in the family and attach shared evidence references to every entry.

- [x] **Step 2: Derive and lock exact totals**

```bash
node scripts/pptxgenjs-surface-audit.mjs --check
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Update count assertions from the audit output, never from an estimate.

- [x] **Step 3: Generate and inspect the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
git diff --check
git diff --stat
```

Confirm all 175 target atoms are closed and no unrelated surface changed.

---

### Task 5: Batch verification, review, commit, and push

**Files:**

- Review all files changed by Tasks 1-4.

**Interfaces:**

- Consumes: all family evidence and generated audit totals.
- Produces: one reviewed commit synchronized with `origin/main`.

- [x] **Step 1: Run family and repository verification**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks bullet and numbering behavior across every declared owner"
./node_modules/.bin/vitest run packages/sdk/src/index.test.ts \
  -t "creates and reopens bullet and numbering owners in all six formats"
pnpm typecheck
pnpm test
```

Run npm tarball and Chrome smoke once each after the focused tests.

- [x] **Step 2: Review the final diff**

Check runtime expectations against OOXML evidence, category labels against control behavior, generated counts against the manifest, and staged paths against the family scope.

- [x] **Step 3: Commit and push the family**

```bash
git add <only-family-files>
git commit -m "docs: close bullet and numbering surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`; `.pnpm-store/` remains untracked and unstaged.
