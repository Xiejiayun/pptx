# Chart Residual 23 Implementation Plan

> **For agentic workers:** Repository edits are owned by the main agent. Evidence agents may only write `/tmp/chart-residual-*`; they must not modify the worktree. Steps use checkbox (`- [x]`) syntax for tracking.

**Goal:** Close the final 23 PptxGenJS 4.0.1 chart atoms with strict native create, import, edit, rollback, reopen, package, browser, and OOXML evidence.

**Architecture:** Implement four sequential slices because the chart model, normalizer, renderer, reader, and editor overlap. Run matrix analysis, upstream runtime controls, and artifact planning in parallel as read-only evidence lines; register the 23 atoms and run expensive artifact gates only after all slices pass focused review.

**Tech Stack:** TypeScript 5.9, Vitest, Node.js 22 ESM probes, OOXML, pnpm/npm pack, persistent Chrome, and the deterministic PptxGenJS surface audit.

## Global Constraints

- PptxGenJS `4.0.1` is the only compatibility baseline.
- Slice order is A → B → C → D; overlapping chart source files are never edited concurrently.
- Agent A owns matrix/reuse evidence, Agent B owns runtime/control evidence, Agent C owns native/package/browser/OOXML evidence, and the main agent owns every repository edit.
- Evidence agents write only `/tmp`; the main agent reviews and merges evidence.
- Do not register any of the 23 atoms until all four slices and the final artifact gate pass.
- Run focused tests once per slice and npm pack, persistent Chrome, PowerPoint 2010 validation, and the full suite once for the complete capability family.
- Commit and push once for the complete Chart Residual 23 capability family after final review.
- Never stage `.pnpm-store/`.

---

### Task 1: Slice A — Catalog, Layout, Title, and Effective Language

**Files:**
- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/chart-definition.internal.ts`
- Modify: `packages/model/src/chart-options.internal.ts`
- Modify: `packages/model/src/chart-render.internal.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/chart-edit.internal.ts`
- Modify: `packages/model/src/chart-definition.internal.test.ts`
- Modify: `packages/model/src/chart-render.internal.test.ts`
- Modify: `packages/model/src/chart-state.internal.test.ts`
- Modify: `packages/model/src/chart-edit.internal.test.ts`
- Modify as required by the bubble workbook contract: `packages/model/src/chart-workbook.internal.ts`

**Interfaces:**
- Produces: `ChartType` member `bubble3D`; `ChartBar3DShape`; `ChartLayoutOptions`; `ChartTitleAlignment`; `ChartBar3DGroupOptions.shape`; `ChartOptions.layout`; `ChartTitleOptions.align`.
- Preserves: title position, nested categories, untouched `dPt`, `dLbl`, `c:tx`, `spPr`, `effectLst`, and `extLst` spans.

- [x] Add failing normalization tests for the exact catalog order, detached/frozen `bubble3D`, strict six-token bar shape, complete `[0,1]` layout, title alignment, and invalid zero-mutation cases.
- [x] Run `pnpm exec vitest run packages/model/src/chart-definition.internal.test.ts` and confirm the new assertions fail for missing state.
- [x] Add failing render/read tests for `<c:bubble3D val="1"/>`, `<c:shape>`, inner edge-mode manual layout, direct title paragraph alignment, mixed bubble flags, and conservative uniform run-language promotion.
- [x] Run `pnpm exec vitest run packages/model/src/chart-render.internal.test.ts packages/model/src/chart-state.internal.test.ts` and confirm the semantic gaps fail.
- [x] Implement the public types and strict normalization, treating `bubble3D` as the bubble data/workbook contract while retaining a distinct type.
- [x] Implement canonical rendering and strict import projection; absent/default state remains absent and malformed or ambiguous OOXML remains unsupported.
- [x] Make Slice A edits owner-scoped so an unrelated option change does not canonicalize unowned chart XML.
- [x] Add edit tests for same-value no-op, local replacement, rollback, COW, reopen, and untouched custom spans.
- [x] Run `pnpm exec vitest run packages/model/src/chart-definition.internal.test.ts packages/model/src/chart-render.internal.test.ts packages/model/src/chart-state.internal.test.ts packages/model/src/chart-edit.internal.test.ts packages/model/src/chart-workbook.internal.test.ts`.
- [x] Run `node /tmp/chart-residual-slice-a-runtime-control.mjs --refresh --expect-closure` and review all closure assertions.

### Task 2: Slice B — Shared Line Cap

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/shape-line.internal.ts`
- Modify: every line wrapper identified by `/tmp/chart-residual-23-slice-b-file-map.md`
- Modify: chart source and focused tests that read, render, or patch series/grid lines.

**Interfaces:**
- Produces: `ShapeLineCap = 'flat' | 'round' | 'square'` and `ShapeLine.cap` on the `kind: 'line'` branch.
- Serializes: `flat → flat`, `round → rnd`, and `square → sq` as `<a:ln cap="…">` attributes.

- [x] Add failing strict normalization tests, including rejection of `cap` on `kind: 'none'` and preservation of cap-only lines.
- [x] Add failing renderer/reader/editor tests for all three tokens, cap clearing, cap-only edits, and preservation of width/fill/dash/join/arrows/extensions.
- [x] Implement one shared cap attribute mapper and apply it to every `<a:ln>` wrapper; never render cap as a child element.
- [x] Update semantic equality and surgical chart line editing so unrelated series/grid state survives.
- [x] Run the focused line, shape, chart render/state/edit, model, and SDK tests listed in `/tmp/chart-residual-23-slice-b-file-map.md` once.
- [x] Run `node /tmp/chart-residual-slice-b-runtime-control.mjs --refresh --expect-closure` and review the OOXML and zero-mutation assertions.

### Task 3: Slice C — Indexed Point Styles and Effects

**Files:**
- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/chart-options.internal.ts`
- Modify: `packages/model/src/chart-render.internal.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/chart-edit.internal.ts`
- Modify: corresponding chart focused tests.

**Interfaces:**
- Produces: indexed `ChartPointOptions`; `ChartSeriesOptions.points`; `ChartSeriesOptions.shadow`; series-scoped data-label fill.
- Represents compatibility aliases only through strict native point/series state; no flat `dataNoEffects`, `dataBorder`, `invertedColors`, or `dataLabelBkgrdColors` API is added.

- [x] Add failing tests for dense detached point arrays, unique safe indices within each series value count, point fill/line/shadow, series shadow, and series data-label fill.
- [x] Add failing OOXML and edit-isolation tests for direct `c:dPt/c:spPr`, series `c:spPr` effects, and `c:dLbls/c:spPr` schema order.
- [x] Implement normalization with series value counts and strict frozen state.
- [x] Implement render/read/edit of the selected semantic owner only; preserve untouched `effectLst`, `extLst`, `dPt`, and `dLbl` bytes.
- [x] Run the chart definition/render/state/edit focused tests once.
- [x] Run the Agent B Slice C runtime/control probe with `--refresh --expect-closure` and review all assertions.

### Task 4: Slice D — Scatter Point Labels

**Files:**
- Modify: `packages/model/src/chart.ts`
- Modify: `packages/model/src/chart-options.internal.ts`
- Modify: `packages/model/src/chart-render.internal.ts`
- Modify: `packages/model/src/chart-state.internal.ts`
- Modify: `packages/model/src/chart-edit.internal.ts`
- Modify: corresponding chart focused tests.

**Interfaces:**
- Produces: `ChartDataLabelFieldKind`; indexed `ChartPointDataLabelOptions`; `ChartSeriesDataLabelOptions.pointLabels`.
- Preserves source field UUIDs and extension payload for unrelated edits; canonical field IDs are deterministic only when exact point-label content is replaced.

- [x] Add failing normalization tests for unique in-range indices, literal/field combinations, XML-safe text, dense detached arrays, and cardinality.
- [x] Add failing create/import/edit/rollback/reopen tests for `XY`, `custom`, and `customXY`, including empty customXY labels with field runs.
- [x] Implement deterministic chart/series/point/field identity without exposing OOXML UUIDs in the public model.
- [x] Implement owner-local rendering and editing that preserves unrelated `c16:uniqueId`, UUID, `extLst`, and custom label bytes.
- [x] Run the chart definition/render/state/edit focused tests once.
- [x] Run the Agent B Slice D runtime/control probe with `--refresh --expect-closure` and review all assertions.

### Task 5: Complete-Family Evidence, Matrix Closure, and Delivery

**Files:**
- Create: `scripts/chart-residual-23-lifecycle-probe.mjs`
- Modify: adapter aggregate tests and lifecycle/package/browser test entry points.
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: deterministic audit JSON/Markdown and progress documentation.

**Interfaces:**
- Produces: one shared native/npm/browser lifecycle contract and final audit counts `1,774/1,774`, supported `764`, deliberate-difference `542`, deprecated `94`, defect-excluded `374`, unverified `0`.

- [x] Add the aggregate adapter control titled `closes chart residual visual options through strict native state` and the shared lifecycle probe with exact OOXML assertions.
- [x] Run the complete chart focused test set once and run `pnpm typecheck`.
- [x] Run one npm-pack consumer pass using the shared lifecycle probe.
- [x] Run one persistent-Chrome pass using the same lifecycle state contract.
- [x] Generate native six-format create/edit/reopen artifacts and validate the PPT2010 profile at `0 errors / 0 warnings`.
- [x] Run `node_modules/.bin/vitest run --maxWorkers 1 --minWorkers 1` once for the complete repository.
- [x] Register all 23 atoms together, regenerate the audit twice, compare the outputs byte-for-byte, and run `pnpm audit:pptxgenjs`.
- [x] Review Agent A/B/C final reports, `git diff --check`, the exact staged paths, and confirm `.pnpm-store/` is untracked and unstaged.
- [x] Commit with `git commit -m "feat: complete chart residual surface family"`.
- [x] Push with `GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=10' git push git@github.com:Xiejiayun/pptx.git main`.
- [x] Fetch the remote tracking ref and require `git rev-list --left-right --count HEAD...origin/main` to print `0 0`.
