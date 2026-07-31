# Custom Geometry Evaluator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a pure DrawingML custom-geometry evaluator and a live `ShapeModel` entry point that resolve built-ins, all 17 formulas, guides, handles, sites, text rectangles, and path values into a detached deep-frozen numeric tree.

**Architecture:** Keep direct OOXML parsing/editing unchanged. Add the public evaluated type/error family beside the existing direct types in `custom-geometry.ts`, put the pure entry point in `custom-geometry-evaluator.ts`, keep scalar/dependency/tree logic in a focused internal module, and let `ShapeModel.evaluateCustomGeometry()` provide only live transform context. Evaluation is read-only and source-order strict; it never rewrites package bytes.

**Tech Stack:** TypeScript 5.9, Vitest, `@pptx/model`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, Node/browser package smoke tests, Artifact Tool, LibreOffice headless, Poppler, `pptx-inspect`, PowerPoint 2010 Open XML validation.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-01-custom-geometry-evaluator-design.md` exactly.
- Public names are `CustomGeometryEvaluationContext`, `CustomGeometryEvaluationErrorCode`, `CustomGeometryEvaluationError`, every `EvaluatedCustomGeometry*` type, `evaluateCustomGeometry()`, and `ShapeModel.evaluateCustomGeometry()`.
- Support exactly the 37 built-ins and all 17 formula operators listed in the design.
- Evaluate adjustments first and guides second, both in source order; pre-audit dependencies only to classify cycle, forward reference, and unknown token errors.
- Use IEEE-754 finite numbers without integer rounding; normalize negative zero; division by zero returns zero.
- Keep source path coordinate spaces unchanged; do not scale paths, compute arc endpoints, tessellate curves, or calculate bounds.
- Materialize the effective default text rectangle in evaluated output.
- Never mutate package bytes, relationships, shape identity, source snapshots, or mutation journals during evaluation.
- Do not add preset-geometry expansion, adjustment overrides, handle dragging, connector snapping/creation, or a renderer.
- Execute inline in the root task; do not dispatch subagents.
- Never stage, delete, or modify `.pnpm-store/`.
- End every task with source review, isolated commit, push to `main`, fetch, and `HEAD...origin/main == 0 0`.

---

### Task 1: Public evaluated types and scalar DrawingML semantics

**Files:**
- Modify: `packages/model/src/custom-geometry.ts`
- Create: `packages/model/src/custom-geometry-evaluator.internal.ts`
- Create: `packages/model/src/custom-geometry-evaluator.internal.test.ts`

**Interfaces:**
- Produces: all public `CustomGeometryEvaluation*` and `EvaluatedCustomGeometry*` types from the design.
- Produces: `CustomGeometryEvaluationError` with stable `code`, optional `guideName`, and optional `token`.
- Produces internally: `evaluateBuiltInGuide(token, context)` and `evaluateFormulaValue(formula, resolve, guideName)`.
- Consumes later: Task 2 guide environment and Task 3 tree resolver.

- [ ] **Step 1: Add failing tests for all 37 built-ins**

Create a table-driven test with `{ width: 120, height: 80 }`. Assert exact values for zero edges, width/height,
centers, short/long sides, every `wdN`/`hdN`/`ssdN`, and every circle fraction. Include an odd context
`{ width: 101, height: 67 }` and assert fractional values such as `wd2 === 50.5` and `ssd32 === 67 / 32`.

```ts
expect(evaluateBuiltInGuide('l', context)).toBe(0);
expect(evaluateBuiltInGuide('r', context)).toBe(120);
expect(evaluateBuiltInGuide('3cd4', context)).toBe(270 * 60_000);
expect(evaluateBuiltInGuide('unknown', context)).toBeUndefined();
```

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry-evaluator.internal.test.ts
```

Expected: FAIL because the evaluator module does not exist.

- [ ] **Step 2: Add failing tests for every formula operator**

Use a resolver backed by `{ x: 3, y: 4, z: 5, rightAngle: 90 * 60_000 }`. Cover:

```ts
expect(formula('val', ['x'])).toBe(3);
expect(formula('*/', ['x', 'y', 2])).toBe(6);
expect(formula('*/', ['x', 'y', 0])).toBe(0);
expect(formula('+/', ['x', 'y', 0])).toBe(0);
expect(formula('?:', [0, 10, 20])).toBe(20);
expect(formula('mod', ['x', 'y', 12])).toBe(13);
expect(formula('sin', [2, 'rightAngle'])).toBeCloseTo(2);
expect(formula('cat2', [10, 3, 4])).toBeCloseTo(6);
expect(formula('sat2', [10, 3, 4])).toBeCloseTo(8);
```

Add negative-zero normalization, quadrant-sensitive `at2`, negative `sqrt`, overflow, and non-finite-result
assertions. Negative sqrt must throw code `invalid-domain`; a finite-input operation producing infinity must throw
`non-finite-result`.

- [ ] **Step 3: Define the complete public evaluated type family**

Copy the exact context, error code, guide, point, text rectangle, command, XY handle, polar handle, handle union,
connection-site, path, and root interfaces from the design. Implement the error as:

```ts
export class CustomGeometryEvaluationError extends Error {
  constructor(
    readonly code: CustomGeometryEvaluationErrorCode,
    message: string,
    readonly guideName?: string,
    readonly token?: string,
  ) {
    super(message);
    this.name = 'CustomGeometryEvaluationError';
  }
}
```

Add these declarations to `custom-geometry.ts`, which is already exported from the model package root. Do not
add the public `evaluateCustomGeometry()` function until Task 3; this keeps the internal module from importing a
runtime class from a facade that imports the internal module back.

- [ ] **Step 4: Implement built-ins and scalar formula evaluation**

In the internal module, use a closed built-in `switch` and these helpers:

```ts
export function evaluateBuiltInGuide(
  token: string,
  context: Readonly<CustomGeometryEvaluationContext>,
): number | undefined;

export function evaluateFormulaValue(
  formula: Readonly<CustomGeometryFormula>,
  resolve: (value: CustomGeometryValue, location: string) => number,
  guideName: string,
): number;
```

Implement the exact formulas from the design. Convert between radians and OOXML angles only at trig/`at2`
boundaries. Route every result through one helper that normalizes `-0`, rejects `NaN`/infinity, and distinguishes
negative sqrt from general non-finite output.

- [ ] **Step 5: Run focused and workspace type gates**

```bash
pnpm vitest run packages/model/src/custom-geometry-evaluator.internal.test.ts
pnpm typecheck
git diff --check
```

Expected: scalar tests pass and the public types compile without exporting internal helpers from the package root.

- [ ] **Step 6: Review, commit, push, and verify Task 1**

Review operator arity, angle units, all 37 built-ins, finite-result handling, and public type names. Then run:

```bash
git add packages/model/src/custom-geometry.ts \
  packages/model/src/custom-geometry-evaluator.internal.ts \
  packages/model/src/custom-geometry-evaluator.internal.test.ts
git diff --cached --check
git commit -m "feat: add custom geometry scalar evaluation"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

### Task 2: Ordered guide environment and dependency diagnostics

**Files:**
- Modify: `packages/model/src/custom-geometry-evaluator.internal.ts`
- Test: `packages/model/src/custom-geometry-evaluator.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 `evaluateBuiltInGuide()`, `evaluateFormulaValue()`, and public error/types.
- Produces internally: `evaluateGuideEnvironment(geometry, context)` returning evaluated guide arrays and a stable value resolver.
- Guarantees: custom values already evaluated shadow built-ins; later same-name built-ins do not shadow early built-ins.

- [ ] **Step 1: Add failing source-order and precedence tests**

Use a normalized geometry containing:

```ts
adjustments: [
  { name: 'adj', formula: { operator: 'val', operands: [25_000] } },
],
guides: [
  { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj', 100_000] } },
  { name: 'w', formula: { operator: 'val', operands: ['w'] } },
  { name: 'x2', formula: { operator: '+-', operands: ['w', 'x1', 0] } },
],
```

At width `200_000`, require `x1=50_000`, custom `w=200_000`, and `x2=250_000`. Assert adjustments and guides
remain separate ordered arrays and `resolve('x2')` returns the stored custom result.

- [ ] **Step 2: Add failing dependency diagnostic tests**

Lock exact error codes and details for:

```ts
// forward-reference
g1 = val g2; g2 = val 1
// self-cycle
g1 = val g1
// multi-node cycle
g1 = val g2; g2 = val g3; g3 = val g1
// unknown token
g1 = val missingGuide
```

Also assert `name="w" fmla="val w"` is not a cycle because unresolved `w` is the built-in width, while a prior
custom `w` shadows the built-in for later guides.

- [ ] **Step 3: Implement dependency pre-audit**

Pre-index ordered adjustment and guide names. Walk only string operands that are not already-resolved values or
built-ins at the point of use. Use DFS colors to collect cycle members, but retain the original list order for
actual evaluation. Classification order must be cycle, forward reference, then unknown token; include the active
guide and token in the thrown error.

- [ ] **Step 4: Implement the ordered environment**

Add:

```ts
export interface EvaluatedGuideEnvironment {
  readonly adjustments?: readonly EvaluatedCustomGeometryGuide[];
  readonly guides?: readonly EvaluatedCustomGeometryGuide[];
  resolve(value: CustomGeometryValue, location: string): number;
}

export function evaluateGuideEnvironment(
  geometry: Readonly<CustomGeometry>,
  context: Readonly<CustomGeometryEvaluationContext>,
): EvaluatedGuideEnvironment;
```

Evaluate adjustment entries, then guide entries. Store each result only after its own formula succeeds. Resolver
priority is direct number, stored custom guide, built-in, classified semantic error. Freeze returned arrays and
entries; keep the resolver closure private to the returned internal environment.

- [ ] **Step 5: Run focused and regression tests**

```bash
pnpm vitest run packages/model/src/custom-geometry-evaluator.internal.test.ts
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm typecheck
git diff --check
```

Expected: all guide-order and error-code cases pass; the direct-state codec remains unchanged.

- [ ] **Step 6: Review, commit, push, and verify Task 2**

```bash
git add packages/model/src/custom-geometry-evaluator.internal.ts \
  packages/model/src/custom-geometry-evaluator.internal.test.ts
git diff --cached --check
git commit -m "feat: resolve custom geometry guides"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

### Task 3: Full numeric geometry-tree resolver and public pure API

**Files:**
- Create: `packages/model/src/custom-geometry-evaluator.ts`
- Modify: `packages/model/src/custom-geometry-evaluator.internal.ts`
- Test: `packages/model/src/custom-geometry-evaluator.internal.test.ts`
- Modify: `packages/model/src/index.ts`

**Interfaces:**
- Consumes: Task 2 `evaluateGuideEnvironment()`.
- Produces: public `evaluateCustomGeometry(geometry, context): EvaluatedCustomGeometry`.
- Guarantees: strict input normalization, effective default rect materialization, recursive freeze, and no input mutation.

- [ ] **Step 1: Add a failing full-tree fixture**

Create one geometry containing adjustment/shape guides, mixed XY/polar handles, duplicate connection sites,
token/numeric text rectangle, two paths, every command kind, present/omitted path flags, and an arc whose radii
come from guides. Evaluate at `{ width: 200_000, height: 100_000 }` and assert the complete numeric tree.

The expected text rectangle for a second default-only fixture is exactly:

```ts
{ left: 0, top: 0, right: 200_000, bottom: 100_000 }
```

Assert path `width/height` and numeric direct coordinates remain unchanged rather than being scaled to context.

- [ ] **Step 2: Add detachment, freeze, and optional-presence tests**

Mutate the caller geometry and context after evaluation; require the result to stay unchanged. Recursively assert
frozen root/context, arrays, guide entries, handles, positions, sites, rectangle, paths, commands, controls, and
endpoints. Omitted guide/handle/site arrays must remain absent; text rectangle must always be present.

- [ ] **Step 3: Add strict context and semantic-domain tests**

Reject inherited/accessor/unknown/symbol context fields, zero/negative/fractional/unsafe width or height, malformed
geometry input, guide-resolved zero/negative arc radii, and non-finite formula output. Structural failures use
`TypeError`; semantic failures use `CustomGeometryEvaluationError` with the design's stable code.

- [ ] **Step 4: Implement strict context normalization and tree mapping**

In the public module, normalize context using exact descriptor-safe own data properties, call existing
`normalizeCustomGeometry(geometry, 'Custom geometry evaluation')`, then call the internal tree resolver. In the
internal module add focused point, handle, site, rectangle, command, and path mapping helpers. Preserve optional
property presence with `Object.hasOwn()` and validate evaluated arc radii after resolution.

```ts
export function evaluateCustomGeometry(
  geometry: CustomGeometry,
  context: CustomGeometryEvaluationContext,
): EvaluatedCustomGeometry {
  const normalizedGeometry = normalizeCustomGeometry(geometry, 'Custom geometry evaluation');
  const normalizedContext = normalizeEvaluationContext(context);
  return evaluateCustomGeometryTree(normalizedGeometry, normalizedContext);
}
```

- [ ] **Step 5: Run evaluator, codec, and type gates**

```bash
pnpm vitest run packages/model/src/custom-geometry-evaluator.internal.test.ts
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm typecheck
git diff --check
```

Expected: complete numeric-tree assertions pass without changing direct codec snapshots or bytes.

- [ ] **Step 6: Review, commit, push, and verify Task 3**

Review every direct `CustomGeometryValue` location against the mapper and verify no path scaling or source mutation.

```bash
git add packages/model/src/custom-geometry-evaluator.ts packages/model/src/index.ts \
  packages/model/src/custom-geometry-evaluator.internal.ts \
  packages/model/src/custom-geometry-evaluator.internal.test.ts
git diff --cached --check
git commit -m "feat: evaluate custom geometry trees"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

### Task 4: Live `ShapeModel` and SDK lifecycle

**Files:**
- Modify: `packages/model/src/shapes.ts`
- Test: `packages/model/src/model.test.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 3 public `evaluateCustomGeometry()` and `EvaluatedCustomGeometry`.
- Produces: `ShapeModel.evaluateCustomGeometry(): EvaluatedCustomGeometry | undefined`.
- Guarantees: live transform sizing, read-only behavior, and stable package identity across lifecycle operations.

- [ ] **Step 1: Add failing model tests for live evaluation**

Create a custom shape with formulas using `w/h`, a token text rectangle, handle, site, and path. Capture slide-part
bytes and journal length, evaluate, and require no change. Then call `setTransform({ width: 300_000,
height: 150_000 })`; a new evaluation must use the new context and produce new guide/text-rectangle results.

- [ ] **Step 2: Add lifecycle and failure tests**

Cover duplicate, move, write/reopen, and all six formats. A preset shape and a malformed custom geometry must
return `undefined`. A lexically safe geometry containing an unknown token must still be readable through
`customGeometry` but throw `CustomGeometryEvaluationError` when evaluated, with exact bytes/journal unchanged.

- [ ] **Step 3: Implement the live method**

Import the pure function under a non-conflicting alias and add:

```ts
evaluateCustomGeometry(): EvaluatedCustomGeometry | undefined {
  const geometry = this.customGeometry;
  if (!geometry) return undefined;
  const { width, height } = this.transform;
  return evaluateGeometry(geometry, { width, height });
}
```

Do not cache results; each call must observe current XML and transform state.

- [ ] **Step 4: Add SDK public-surface tests**

From `@pptx/sdk`, import the pure evaluator, error class, and evaluated types. Test zero-input creation, live shape
evaluation, browser-safe export reachability, reopen, recursive freeze, and unknown-token failure without using
workspace-internal paths.

- [ ] **Step 5: Run lifecycle gates**

```bash
pnpm vitest run packages/model/src/model.test.ts -t "custom geometry evaluation"
pnpm vitest run packages/sdk/src/index.test.ts -t "custom geometry evaluation"
pnpm typecheck
git diff --check
```

Expected: focused public lifecycle passes for pure and shape entry points.

- [ ] **Step 6: Review, commit, push, and verify Task 4**

```bash
git add packages/model/src/shapes.ts packages/model/src/model.test.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: evaluate live custom geometry"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

### Task 5: PptxGenJS conformance and actual tarball coverage

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 4 package exports and shape method.
- Produces: PptxGenJS legal numeric-path evaluation evidence.
- Produces: smoke JSON `customGeometryEvaluator: true` for Node/browser/types.

- [ ] **Step 1: Add PptxGenJS public-output evaluation coverage**

Generate a legal PptxGenJS 4.0.1 `ShapeType.custGeom + points` presentation through public `write()`, import it,
and evaluate the resulting shape. Assert its numeric path commands equal the imported direct snapshot, its
effective text rectangle is `{ left: 0, top: 0, right: width, bottom: height }`, and evaluation does not alter
adapter/package bytes. Record that PptxGenJS exposes no formula or evaluator input.

- [ ] **Step 2: Add Node actual-package smoke**

Using only the installed tarball export, create a formula geometry covering built-ins, guides, handle/site/rect,
and arc/path commands. Call both pure and shape evaluators; assert guide values, effective default rect, frozen
nested output, live transform recomputation, typed unknown-token error, write/reopen consistency, and 0 error
diagnostics.

- [ ] **Step 3: Add browser runtime smoke**

Extend the generated browser script to import `evaluateCustomGeometry` and `CustomGeometryEvaluationError` from
`dist/browser.js`, evaluate through pure and live APIs, write a Blob, reopen, and assert the same numeric/frozen
results. Do not import Node built-ins.

- [ ] **Step 4: Add consumer type coverage**

Import every public evaluator type. Add valid type declarations for context, evaluated guide/point/handles/site/
rectangle/path/command/root and both entry points. Add `@ts-expect-error` cases for missing context height,
string context width, an extra context field, string evaluated coordinates, invalid error code, token-bearing
evaluated paths, and nonexistent shape arguments. Keep fractional/unsafe context checks in runtime tests because
TypeScript's `number` type cannot encode safe-integer constraints.

- [ ] **Step 5: Pack and run package gates**

Add nested/top-level JSON `customGeometryEvaluator`. Then run:

```bash
pnpm typecheck
pnpm build
git diff --exit-code -- packages/pptx/dist
mkdir -p /tmp/pptx-geometry-evaluator-pack
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-geometry-evaluator-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-geometry-evaluator-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected JSON includes `"customGeometryEvaluator":true`, `"types":true`, and CLI `0.1.0`.

- [ ] **Step 6: Review, commit, push, and verify Task 5**

```bash
git add packages/pptxgenjs-adapter/src/index.test.ts scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify packaged custom geometry evaluator"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

### Task 6: Real-file gallery, documentation, and release gate

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Temporary only: scripts/output under `pptx_geometry_eval_dir=$(mktemp -d /tmp/pptx-custom-geometry-evaluator.XXXXXX)`

**Interfaces:**
- Documents: evaluator API, built-ins/operators, source-order semantics, numeric precision, errors, tree output, and remaining connector boundary.
- Proves: native and LibreOffice-normalized real files strict-reopen and evaluate through the actual tarball.

- [ ] **Step 1: Load presentation tooling and verify the offline stack**

Read the current Presentations skill before artifact work, load bundled workspace dependencies, then run:

```bash
command -v pptx-inspect
pptx-inspect --json doctor
```

Use the returned LibreOffice/Python/Poppler paths; do not assume system binaries.

- [ ] **Step 2: Generate a four-slide actual-tarball evaluator gallery**

Create four slides covering:

1. all 17 operators with visible labeled paths;
2. all 37 built-ins across portrait/landscape and odd extents;
3. evaluated XY/polar handles, ordered/duplicate sites, and token-backed text rectangles;
4. multiple paths, all command kinds, default rect materialization, and live transform recomputation.

Save:

```text
$pptx_geometry_eval_dir/custom-geometry-evaluator-gallery.pptx
$pptx_geometry_eval_dir/source-and-evaluated-snapshots.json
$pptx_geometry_eval_dir/custom-geometry-evaluator-gallery.sha256
```

The actual tarball must strict-reopen and evaluate every named target; JSON must contain exact source snapshots,
contexts, evaluated trees, and error-free shape lists.

- [ ] **Step 3: Render, inspect, and check overflow**

Render every page at 180 DPI, run the bundled overflow checker, create a four-page montage, and inspect every page
at original resolution. Require readable labels, intact paths/text, no clipping/overlap/off-slide objects, and
visually plausible formula quadrants. Iterate only temporary artifacts until all checks pass.

- [ ] **Step 4: LibreOffice round-trip and semantic comparison**

Save with an isolated LibreOffice profile to:

```text
$pptx_geometry_eval_dir/roundtrip/custom-geometry-evaluator-gallery.pptx
```

Reopen original and round-trip through the actual tarball. Write `roundtrip-comparison.json` with slide/shape
counts, names/order, source direct trees, evaluated guide arrays, handles/sites/rectangles, and numeric path
commands. Require strict evaluation of every target. Compare client-normalized numeric results with a documented
floating tolerance; record direct expression rewrites without weakening the reader/evaluator.

- [ ] **Step 5: Validate both packages**

```bash
pptx-inspect --json package validate "$pptx_geometry_eval_dir/custom-geometry-evaluator-gallery.pptx" --profile powerpoint-2010
pptx-inspect --json package validate "$pptx_geometry_eval_dir/roundtrip/custom-geometry-evaluator-gallery.pptx" --profile powerpoint-2010
pptx-inspect --json package diff "$pptx_geometry_eval_dir/custom-geometry-evaluator-gallery.pptx" "$pptx_geometry_eval_dir/roundtrip/custom-geometry-evaluator-gallery.pptx"
```

Required: 0 errors and 0 warnings for original and round-trip. Diff output is normalization evidence, not a
zero-diff requirement.

- [ ] **Step 6: Update public documentation**

Document the pure and shape APIs, all public types, 37 built-ins, 17 formulas, source-order/custom-shadowing
rules, finite fractional results, default rect materialization, recursive freeze, error codes, no-mutation
behavior, PptxGenJS default/numeric-only boundary, and LibreOffice evidence. Remove current claims that geometry
evaluation is pending. Keep path scaling/arc endpoint/bounds, handle dragging, and connector snapping/creation
explicitly outside this evaluator.

- [ ] **Step 7: Run the complete release gate**

```bash
rg -n -i "geometry evaluation.*(pending|not implemented|尚未|未支持)|geometry evaluator 为下一子项" \
  CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --exit-code -- packages/pptx/dist
node scripts/smoke-npm-package.mjs /tmp/pptx-geometry-evaluator-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

The stale-text search must return no matches. Expected: all tests pass with only the established performance skip
in the normal suite, the explicit 1,000-part performance test passes, dist is reproducible, and actual tarball
Node/browser/types/CLI smoke succeeds.

- [ ] **Step 8: Review, commit, push, and verify Task 6**

Stage only the five documentation files:

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document custom geometry evaluator"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

## Plan Self-Review

- Spec coverage: Tasks 1–3 implement scalar semantics, dependency diagnostics, complete numeric trees, and the pure API; Task 4 adds live model/SDK behavior; Task 5 proves PptxGenJS and packed runtimes/types; Task 6 proves real-file compatibility and documentation.
- Scope: preset expansion, path scaling, arc endpoints, tessellation, bounds, handle dragging, and connector snapping/creation remain explicitly excluded.
- Type consistency: every task uses the exact `CustomGeometryEvaluation*`, `EvaluatedCustomGeometry*`, and `evaluateCustomGeometry` names from the design.
- Numeric consistency: all formula results are finite doubles, division by zero is zero, negative zero is normalized, and evaluated arc radii remain positive.
- State consistency: direct geometry remains lossless; evaluated output always materializes effective text rectangles and preserves optional list/field presence elsewhere.
- Mutation safety: pure/model success and every semantic failure assert exact bytes, relationships, identity, and journal stability.
- Existing compatibility: no evaluator logic enters `custom-geometry.internal.ts`; direct read/edit behavior and existing snapshots remain unchanged.
- Placeholder scan: every task has concrete files, signatures, fixtures, commands, expected results, review scope, commit, push, fetch, and divergence checks.
- Execution: standing user direction selects inline execution and excludes subagent dispatch.
