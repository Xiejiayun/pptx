# Presentation Root, Output, and Theme Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns all
> repository edits, review, commit, and push. Steps use checkbox (`- [ ]`) syntax
> for tracking.

**Goal:** Close exactly 35 PptxGenJS presentation-root, layout, theme, and output
declaration atoms by reusing the existing native lifecycle and adding one
aggregate upstream runtime control.

**Architecture:** Frozen status groups generate the exact manifest family. One
adapter test locks method returns and declaration defects; existing native,
package, browser, and OOXML probes provide the positive lifecycle evidence.
Focused gates run for this family and expensive artifact gates run once after
all three lifecycle families are committed.

**Tech Stack:** TypeScript 5.8, Vitest 3, Node ESM, PptxGenJS 4.0.1, stable JSON
and Markdown audit artifacts, npm pack, persistent Chrome, and `pptx-inspect`.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-04-presentation-root-output-surface-family-design.md` exactly.
- Close only the 35 target atoms and require the exact 18/14/3 classification.
- Reuse existing metadata, theme, layout, output, stream, package, browser, and
  OOXML evidence; add only the missing aggregate PptxGenJS control and matrix
  assertions.
- Run focused tests once before review; defer npm pack, Chrome, OOXML, TypeScript,
  and full Vitest to the three-family lifecycle batch boundary.
- Do not stage `.pnpm-store/` or temporary package, browser, or PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the upstream root and output boundaries

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 presentation methods, output selectors, metadata,
  theme, layout projection, and native `PptxDocument` equivalents.
- Produces: the stable title `locks the presentation root, output, and theme
  declarations against PptxGenJS 4.0.1`.

- [x] **Step 1: Lock the exact 35-ID inventory**

Build three explicit ID groups and assert 18 supported, 14 deliberate
differences, three defect exclusions, 35 unique IDs, and no ChartType or slide
atom.

- [x] **Step 2: Lock method returns and layout differences**

Require upstream section/layout/master definition returns, the positional
master-name overload, custom layout name retention, and native editable return
identities, object-only slide options, and normalized custom layout projection.

- [x] **Step 3: Lock write, file, stream, and compression behavior**

Require all six JSZip write representations, `STREAM` returning a Buffer,
`stream()` returning only a Buffer under every compression setting,
`writeFile({ fileName })` returning its path, native positional file output,
native real Node readable, and the known explicit-output compression defect.

- [x] **Step 4: Run the aggregate control once**

```bash
./node_modules/.bin/vitest run packages/pptxgenjs-adapter/src/index.test.ts \
  -t "locks the presentation root, output, and theme declarations against PptxGenJS 4.0.1"
```

Expected: the aggregate control passes and removes all temporary output files.

### Task 2: Generate and verify the 35 manifest entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the aggregate control plus existing property-specific adapter,
  native, package, browser, and OOXML anchors.
- Produces: exactly 35 new manifest entries and global counts
  606/254/86/355/473.

- [x] **Step 1: Generate entries from frozen status groups**

Define exact IDs, native mappings, family evidence, status-specific controls,
serialization/client flags, and notes. Keep the three impossible stream return
branches native-empty and controlled by the aggregate test.

- [x] **Step 2: Lock exact family and global counts**

Assert 35 distinct IDs, 1,301 manifest entries, exact 18/14/3 family counts,
the three defect IDs, and global totals 606 supported, 254
deliberate-difference, 86 deprecated-alias, 355 defect-excluded, and 473
unverified.

- [x] **Step 3: Regenerate and inspect the matrix**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: deterministic artifacts, zero stale/unsupported/diagnostic entries,
and 473 unverified atoms.

### Task 3: Review, commit, push, and synchronize

**Files:**

- Review and stage only the seven capability-family files from Tasks 1-2 plus
  this design and plan.

**Interfaces:**

- Consumes: the verified diff and programmatic matrix comparison against HEAD.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Review exact migration**

Require exactly 35 changed atoms, all from `unverified`, exact 18/14/3
transitions, no other atom content changes, no unrelated staged paths, and no
`.pnpm-store/` files.

- [x] **Step 2: Commit and push**

```bash
git commit -m "docs: close presentation root output surface family"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
