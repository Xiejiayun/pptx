# Image Source, Sizing, and Transform Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> every repository edit, review, commit, and push. Steps use checkbox (`- [ ]`)
> syntax for tracking.

**Goal:** Close exactly fourteen PptxGenJS image source, sizing, and transform
atoms by binding the existing runtime, native, package, browser, and OOXML
vertical slices and one aggregate divergence control to the declaration matrix.

**Architecture:** One frozen deliberate-difference ID group produces the
fourteen entries. Existing image controls and artifact states provide most of
the evidence; one aggregate runtime control locks intrinsic-dimension,
percentage, coercion, and failure-timing differences, and one aggregate audit
assertion locks scope, mappings, exclusions, and totals.

**Tech Stack:** TypeScript 5.8, Node ESM, Vitest 3, PptxGenJS 4.0.1,
`pptx-inspect`, stable JSON and Markdown audit artifacts.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-04-image-source-sizing-transform-surface-family-design.md`
  exactly.
- Close only the fourteen target atoms, all as deliberate differences.
- Keep metadata, interaction, appearance, direct geometry, placeholder, and
  inherited cross-owner atoms outside this batch.
- Evidence agents do not modify repository files.
- Run focused tests once. Run npm pack, persistent Chrome, OOXML inspection,
  and full Vitest once at this image-family boundary.
- Do not stage `.pnpm-store/` or temporary package/browser/PPTX artifacts.
- Finish with review, one family commit, push, fetch, and zero divergence.

---

### Task 1: Lock the exact evidence-backed matrix entries

**Files:**

- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`

**Interfaces:**

- Consumes: existing raster/SVG source, sizing, transform, package, browser,
  and OOXML evidence.
- Produces: exactly fourteen immutable manifest entries and one aggregate scope
  assertion.

- [x] **Step 1: Define frozen ID groups**

Add fourteen deliberate-difference IDs for source, sizing object shape, nested
geometry, sizing modes, rotation, and flips.

- [x] **Step 2: Map native APIs and evidence**

Map source IDs to `ImageSource`/`resolveImageSource`; sizing IDs to
`ImageSizing`/`calculateImageSizing`; crop coordinates to `ImageCropRegion`;
and transform IDs to `Transform`. Attach the existing and aggregate runtime
controls plus actual package, browser, and OOXML anchors.

- [x] **Step 3: Lock scope and totals**

Assert fourteen unique IDs, exact statuses, all mandatory evidence, the
metadata/interaction/appearance exclusions, 1,395 manifest entries, and totals
620/325/91/359/379.

### Task 2: Lock aggregate runtime differences

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 and native image source/sizing/transform APIs.
- Produces: one stable control title covering gaps not isolated by the existing
  seven image controls.

- [x] **Step 1: Prove sizing and source differences**

Compare nested percentage sizing and an intrinsic/outer aspect-ratio mismatch;
lock unknown/missing sizing behavior and permissive empty or malformed base64
input against native pre-mutation rejection.

- [x] **Step 2: Prove transform coercion differences**

Lock PptxGenJS rotation wrapping/string/infinity behavior plus truthy flip
coercion, and require native invalid transforms to leave the package unchanged.

### Task 3: Regenerate and run the focused family gate once

**Files:**

- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the updated immutable manifest.
- Produces: deterministic compatibility artifacts and focused family evidence.

- [x] **Step 1: Regenerate and audit**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: generated JSON/Markdown are deterministic, diagnostics remain empty,
and exactly 379 atoms remain unverified.

- [x] **Step 2: Run focused tests once**

```bash
./node_modules/.bin/vitest run \
  packages/model/src/image-create.internal.test.ts \
  packages/model/src/svg-image-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/raster-image-sizing.test.ts \
  packages/sdk/src/prepared-image-source.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "embedded raster image|path and data images|contain, cover, and crop|SVG data/path pictures|sizing fallbacks|image divergences|ImageProps source|image sources|image sizing|source rectangle"
```

Expected: every selected source, sizing, transform, reopen, and control case
passes in one run.

- [x] **Step 3: Run TypeScript and exact migration review**

Run `./node_modules/.bin/tsc -b --pretty false`, compare generated atoms against
HEAD, and require exactly fourteen `unverified` to `deliberate-difference`
transitions with no other atom content change.

### Task 4: Run the shared image-batch artifact gates

**Files:**

- Verify existing package, browser, and generated PPTX artifacts without adding
  per-atom runs.

**Interfaces:**

- Consumes: the existing unified npm and browser probes.
- Produces: one actual-tarball result, one persistent-Chrome result, one OOXML
  validation result, and one full-suite result.

- [x] **Step 1: Run the actual npm tarball smoke once**

Use the repository's existing npm package smoke command and require raster/SVG
source, sizing, transforms, relationships, write, and reopen checks to pass.

- [x] **Step 2: Run the persistent Chrome smoke once**

Use the existing persistent-browser smoke state and require Blob/data SVG plus
a browser-URL raster source, contain/cover/crop sizing, rotation, flips,
writeBlob, reopen, and validation state to pass.

- [x] **Step 3: Inspect OOXML and run the full suite once**

Inspect the generated image deck with `pptx-inspect package inspect` and
`package validate --profile powerpoint-2010`; then run the repository full
Vitest command once.

### Task 5: Review, commit, and synchronize

**Files:**

- Review and stage only files owned by this capability family.

**Interfaces:**

- Consumes: all verified source, runtime, artifact, and migration evidence.
- Produces: one synchronized capability-family commit.

- [x] **Step 1: Complete three-line and main-agent review**

Require matrix scope, runtime parity, artifact validity, generated artifacts,
and `git diff --check` to be approved with no unrelated tracked path.

- [x] **Step 2: Commit and synchronize**

```bash
git commit -m "docs: close image source sizing transform surface family"
git -c http.version=HTTP/1.1 push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
