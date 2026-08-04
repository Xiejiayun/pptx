# Media Core Surface Family Implementation Plan

> **For agentic workers:** Evidence agents are read-only. The main agent owns
> every repository edit, review, commit, and push. Steps use checkbox (`- [ ]`)
> syntax for tracking.

**Goal:** Close exactly fourteen PptxGenJS media source, type, metadata, and
geometry atoms using the existing native media vertical slice plus one focused
online-media runtime control.

**Architecture:** Frozen manifest groups produce the fourteen entries. Existing
embedded media controls provide most evidence; one aggregate runtime case and
one retained external case in each package/browser probe lock the `online/link`
branch, while one audit assertion locks scope, exclusions, mappings, and
totals.

**Tech Stack:** TypeScript 5.8, Node ESM, Vitest 3, PptxGenJS 4.0.1,
`pptx-inspect`, stable JSON and Markdown audit artifacts.

## Global Constraints

- Implement
  `docs/superpowers/specs/2026-08-04-media-core-surface-family-design.md`
  exactly.
- Close only the fourteen target atoms, all as deliberate differences.
- Keep `DataOrPathProps`, playback/timing, hyperlink, placeholder, and
  image-only atoms outside this batch.
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

**Interfaces:**

- Consumes: existing media runtime, native, package, browser, and OOXML
  evidence.
- Produces: exactly fourteen immutable manifest entries and one aggregate scope
  assertion.

- [x] **Step 1: Define frozen media ID groups**

Add the eleven `MediaProps` IDs and three `MediaType` union IDs as deliberate
differences, and exclude inherited `DataOrPathProps` explicitly.

- [x] **Step 2: Map real native APIs and evidence**

Map source/type fields to `MediaSource`, `SlideModel.addAudio/addVideo`, and
`PptxDocument.addAudio/addVideo`; map poster, metadata, and transform fields to
`AddMediaOptions` and `MediaModel`. Correct the existing nonexistent
`addOnlineVideo` label.

- [x] **Step 3: Lock scope and totals**

Assert fourteen unique IDs, exact statuses, mandatory evidence, exclusions,
1,409 manifest entries, and totals 620/339/91/359/365.

### Task 2: Lock the missing online-media runtime branch

**Files:**

- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**

- Consumes: PptxGenJS 4.0.1 `type: 'online'` behavior and native external-file
  `MediaSource` behavior.
- Produces: one stable control title for `MediaProps.link` and
  `MediaType.online`.

- [x] **Step 1: Prove PptxGenJS online output**

Create a public online media object with explicit geometry and name, then lock
its embed relationships, target, picture state, write/reopen behavior, and
failure/coercion boundaries.

- [x] **Step 2: Prove strict native external-video behavior**

Create native external video through `addVideo` with an HTTP(S) source, verify
its shared core relationship/poster state and richer playback/timing state,
prove there is no separate `addOnlineVideo` API, and require invalid external
sources to leave the package unchanged.

### Task 3: Regenerate and run the focused family gate once

**Files:**

- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**

- Consumes: the updated immutable manifest.
- Produces: deterministic compatibility artifacts and focused media evidence.

- [x] **Step 1: Regenerate and audit**

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: deterministic artifacts, empty diagnostics, and exactly 365
unverified atoms.

- [x] **Step 2: Run focused tests once**

```bash
./node_modules/.bin/vitest run \
  packages/codecs/src/media-*.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "media|audio|video|external mode|online"
```

Expected: selected media source, descriptor, creation, edit, relationship,
reopen, and runtime controls pass in one run.

- [x] **Step 3: Run TypeScript and exact migration review**

Run `./node_modules/.bin/tsc -b --pretty false`, compare generated atoms against
HEAD, and require exactly fourteen `unverified` to `deliberate-difference`
transitions with no other status change.

### Task 4: Extend and run the shared media artifact gates once

**Files:**

- Modify: `scripts/smoke-npm-package.mjs`
- Modify: `scripts/playwright-browser-smoke.js`
- Verify the generated PPTX probes without adding per-atom runs.

**Interfaces:**

- Consumes: the unified npm and browser media probes, each extended with one
  retained external online-video case.
- Produces: one actual-tarball result, one persistent-Chrome result, one OOXML
  validation result, and one full-suite result.

- [x] **Step 1: Run the actual npm tarball smoke once**

Require embedded/external sources, MIME and extension state, posters,
transforms, lifecycle, relationships, write, and reopen checks to pass.

- [x] **Step 2: Run the persistent Chrome smoke once**

Require the existing browser media deck, external transition, Blob/data
sources, transforms, lifecycle, writeBlob/reopen, and diagnostics to pass.

- [x] **Step 3: Inspect OOXML and run the full suite once**

Inspect the generated media deck with `pptx-inspect package inspect` and
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
git commit -m "docs: close media core surface family"
git -c http.version=HTTP/1.1 push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected: divergence is `0 0`.
