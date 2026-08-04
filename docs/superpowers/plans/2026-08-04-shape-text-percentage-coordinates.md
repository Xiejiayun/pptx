# Shape and Text Percentage Coordinates Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict slide-relative percentage coordinates to native shape, custom-shape, and text creation, then close the 12 corresponding PptxGenJS 4.0.1 audit atoms with direct control, package, and OOXML evidence.

**Architecture:** Keep stored and editable transforms as concrete EMU, introduce a separate `SlideCoordinate`/`TransformInput` creation boundary, and resolve percentage strings through one internal axis-aware function. Shape and text creation receive the current `PresentationModel.slideSize`; PptxGenJS control and packed-consumer tests compare the final serialized transforms rather than pretending the APIs use identical field names or numeric units.

**Tech Stack:** TypeScript 5.9, Vitest, Node.js 22, PptxGenJS 4.0.1, OOXML `a:off`/`a:ext`, workspace package smoke, deterministic surface manifest generation, pnpm, and Git.

## Global Constraints

- The approved design is `docs/superpowers/specs/2026-08-04-shape-text-percentage-coordinates-design.md`.
- PptxGenJS 4.0.1 is the only control baseline.
- Public `Transform` and `BaseShapeModel.setTransform()` remain absolute EMU APIs.
- Existing numeric creation inputs remain absolute EMU; callers use `inches()` for explicit inch conversion.
- Percentages resolve against current slide width for `x`/`width` and height for `y`/`height`.
- Position may be negative or exceed the slide; width and height must resolve to positive safe integers.
- Reject malformed, partial, non-finite, whitespace-padded, accessor-backed, and unsafe coordinate inputs before serialization.
- Placeholder-owner transforms continue to override caller geometry.
- Every task ends with focused review, exact-file staging, commit, SSH-443 push, remote-tracking refresh, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`.

---

### Task 1: Shared Percentage Coordinate Type and Resolver

**Files:**
- Modify: `packages/model/src/units.ts`
- Create: `packages/model/src/slide-coordinate.internal.ts`
- Create: `packages/model/src/slide-coordinate.internal.test.ts`

**Interfaces:**
- Consumes: `Emu`, `OoxmlAngle`, `SlideSize`, an input candidate, axis, fallback, and context.
- Produces:

```ts
export type SlideCoordinate = Emu | `${number}%`;

export interface TransformInput {
  readonly x: SlideCoordinate;
  readonly y: SlideCoordinate;
  readonly width: SlideCoordinate;
  readonly height: SlideCoordinate;
  readonly rotation: OoxmlAngle;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

export function resolveSlideCoordinate(
  value: unknown,
  axis: 'horizontal' | 'vertical',
  slideSize: Readonly<SlideSize> | undefined,
  fallback: Emu,
  context: string,
): Emu;
```

- [ ] **Step 1: Write failing resolver tests**

Create `slide-coordinate.internal.test.ts` with a 10 × 8 inch size and exact assertions:

```ts
const SIZE = Object.freeze({ width: inches(10), height: inches(8) });

expect(resolveSlideCoordinate('10%', 'horizontal', SIZE, inches(1), 'x'))
  .toBe(inches(1));
expect(resolveSlideCoordinate('20%', 'vertical', SIZE, inches(1), 'y'))
  .toBe(inches(1.6));
expect(resolveSlideCoordinate('1.25e1%', 'horizontal', SIZE, 0 as Emu, 'width'))
  .toBe(inches(1.25));
expect(resolveSlideCoordinate('-25%', 'horizontal', SIZE, 0 as Emu, 'x'))
  .toBe(inches(-2.5));
expect(resolveSlideCoordinate('125%', 'vertical', SIZE, 0 as Emu, 'y'))
  .toBe(inches(10));
expect(resolveSlideCoordinate(inches(2), 'horizontal', SIZE, 0 as Emu, 'x'))
  .toBe(inches(2));
expect(resolveSlideCoordinate(undefined, 'horizontal', SIZE, inches(3), 'x'))
  .toBe(inches(3));
```

Assert `'-0%'` returns positive zero. Reject `''`, `'10'`, `'%'`, `' 10%'`, `'10% '`,
`'10%%'`, `'10%garbage'`, `'NaN%'`, `'Infinity%'`, objects, arrays, booleans,
non-finite numbers, percentages without a slide size, and values whose resolved EMU is not a
safe integer.

- [ ] **Step 2: Run the focused test and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false vitest run packages/model/src/slide-coordinate.internal.test.ts
```

Expected: FAIL because `slide-coordinate.internal.ts` does not exist.

- [ ] **Step 3: Add public input types and the minimal parser**

Add `SlideCoordinate` and `TransformInput` beside `Transform` in `units.ts`. In the internal
module, use one anchored decimal/exponent percentage expression, require the original string to
equal its trimmed form, require a supplied slide size for strings, multiply by the axis dimension,
round once, reject a non-safe result, and normalize `-0`:

```ts
const PERCENTAGE = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?%$/iu;
```

Do not export the resolver through `packages/model/src/index.ts`; it is an implementation detail.
The public types are already exported because `units.ts` is re-exported by the model index.

- [ ] **Step 4: Verify the resolver and declarations**

```bash
pnpm --config.verify-deps-before-run=false vitest run packages/model/src/slide-coordinate.internal.test.ts
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
```

Review decimal/exponent parsing, axis selection, safe rounding, no inch heuristic, immutable public
types, and no changes to output `Transform`.

- [ ] **Step 5: Commit, push, and verify synchronization**

```bash
git add packages/model/src/units.ts \
  packages/model/src/slide-coordinate.internal.ts \
  packages/model/src/slide-coordinate.internal.test.ts
git commit -m "feat: resolve slide percentage coordinates"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 2: Shape, Custom-Shape, and Text Creation Integration

**Files:**
- Modify: `packages/model/src/preset-shape.ts`
- Modify: `packages/model/src/preset-shape.internal.ts`
- Modify: `packages/model/src/preset-shape.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `TransformInput` and `resolveSlideCoordinate()`, plus
  `this.presentation.slideSize`.
- Produces: percentage-aware `AddShapeOptions`, `AddCustomShapeOptions`, `AddTextOptions`,
  `addRichText()`, and `addPlaceholder()` creation paths while returning concrete EMU transforms.

- [ ] **Step 1: Add failing normalizer and end-to-end tests**

In `preset-shape.internal.test.ts`, add a test named
`resolves strict slide-relative percentage transforms for preset and custom shapes`. Use a 10 × 8
inch slide size, assert the exact normalized x/y/width/height for decimal percentages, verify
absolute `inches()` inputs remain unchanged, and assert zero/negative percentage extents throw.

In `packages/sdk/src/index.test.ts`, add a test named
`creates and reopens shape and text percentage coordinates against the current slide size`:

```ts
const document = PptxDocument.create({
  slideSize: { width: inches(10), height: inches(8) },
});
const slide = document.addSlide();
const shape = slide.addShape('rect', {
  x: '10%', y: '20%', width: '30%', height: '40%',
});
const custom = slide.addCustomShape(customTriangleGeometry, {
  x: '-10%', y: '125%', width: '20%', height: '25%',
});
const text = slide.addText('Percentage text', {
  x: '12.5%', y: '25%', width: '37.5%', height: '50%',
});
```

Assert all three concrete transforms, exact slide XML `a:off`/`a:ext`, reopened transforms, no
validation errors after write, and byte/journal stability when invalid width, height, malformed
string, or coordinate accessor is supplied.

- [ ] **Step 2: Run focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false vitest run \
  packages/model/src/preset-shape.internal.test.ts \
  packages/sdk/src/index.test.ts \
  -t "percentage coordinate|percentage transforms"
```

Expected: FAIL because public option types and creation normalizers accept only concrete EMU.

- [ ] **Step 3: Wire shapes to the shared resolver**

Change `AddShapeOptions` to extend `Partial<TransformInput>`. Add an optional
`Readonly<SlideSize>` parameter to `normalizePresetShape()` and `normalizeCustomShape()`, pass
`this.presentation.slideSize` from `SlideModel.addShape()` and `addCustomShape()`, and resolve
`x/y/width/height` with the shared helper. Keep existing defaults, option-key checks, flip/rotation
normalization, positive extent checks, placeholder override, and transaction order unchanged.

- [ ] **Step 4: Normalize text transform once**

Change `AddTextOptions` to extend `Partial<TransformInput>`. Add a concrete `transform: Transform`
field to `NormalizedAddTextOptions` and `NormalizedTextInput`. Build it in
`validateAddTextOptions(options, slideSize)` by reading each coordinate through own data-property
descriptors and resolving against the supplied slide size. Validate rotation and flips once.

Pass the concrete transform into `addTextShape()`/`textShapeXml()` for plain text, rich text, and
placeholder creation. If a placeholder owner exists, pass `owner.transform` instead. Remove raw
coordinate reads and rounding from `textShapeXml()` so accessors cannot run during rendering.

- [ ] **Step 5: Run integration and regression gates**

```bash
pnpm --config.verify-deps-before-run=false vitest run \
  packages/model/src/slide-coordinate.internal.test.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/sdk/src/index.test.ts
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
```

Review all `addTextShape()` call sites, placeholder precedence, custom geometry, absolute transform
behavior, error-before-mutation, reopened output, and `Transform` editing compatibility.

- [ ] **Step 6: Commit, push, and verify synchronization**

```bash
git add packages/model/src/preset-shape.ts \
  packages/model/src/preset-shape.internal.ts \
  packages/model/src/preset-shape.internal.test.ts \
  packages/model/src/slide.ts \
  packages/sdk/src/index.test.ts
git commit -m "feat: create shapes and text with percentage coordinates"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 3: PptxGenJS Control and Packed-Package Proof

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: locked PptxGenJS 4.0.1 public API and packed `@jiayunxie/pptx` declarations/runtime.
- Produces: direct control comparison plus real consumer proof for shape/text percentage geometry.

- [ ] **Step 1: Add the failing PptxGenJS control**

Add an adapter test named
`matches PptxGenJS shape and text percentage coordinate output with explicit native units`.
Define a 10 × 8 PptxGenJS layout, create the two approved control objects, import public output, and
assert their transforms equal the exact values from the design. Create the native equivalent with
`width`/`height` and percentage strings, assert identical final transforms and zero error
diagnostics after write/reopen. Include one native absolute `inches(1)` coordinate to lock the
deliberate numeric-unit difference.

- [ ] **Step 2: Add the packed-consumer assertion**

In the generated consumer inside `smoke-npm-package.mjs`, create a custom 10 × 8 presentation,
add percentage shape and text objects through packed declarations, write/reopen them, and set a
`shapeTextPercentageCoordinates` boolean only when all eight transform values match. Add that
boolean to the required `apiChecks` object and final summary.

- [ ] **Step 3: Run focused control and packed smoke**

```bash
pnpm --config.verify-deps-before-run=false vitest run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "percentage coordinate output"
pnpm --config.verify-deps-before-run=false typecheck
node scripts/smoke-npm-package.mjs
git diff --check
```

Review real PptxGenJS use, no private-field reads, exact custom layout dimensions, packed type
acceptance, reopened transforms, diagnostics, and no source-workspace fallback.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add packages/pptxgenjs-adapter/src/index.test.ts scripts/smoke-npm-package.mjs
git commit -m "test: prove shape and text percentage parity"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 4: Close 12 Audit Atoms and Regenerate Truthful Progress

**Files:**
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**
- Consumes: Task 1–3 code, tests, PptxGenJS control, package smoke, and OOXML assertions.
- Produces: 12 direct `deliberate-difference` entries and counts
  `supported=7`, `deliberate-difference=12`, `defect-excluded=1`,
  `unverified=1754`, `unsupported=0`, `stale=0`.

- [ ] **Step 1: Add a manifest helper and the exact atom list**

Add a `deliberateDifference()` helper that sets `serialization: true`, supplies code/test/package/
OOXML links, and uses the adapter test as `control`. Define the exact IDs once:

```js
const SHAPE_TEXT_COORDINATE_ATOMS = Object.freeze([
  'interface:PositionProps@property:x',
  'interface:PositionProps@property:y',
  'interface:PositionProps@property:w',
  'interface:PositionProps@property:h',
  'interface:ShapeProps@property:x',
  'interface:ShapeProps@property:y',
  'interface:ShapeProps@property:w',
  'interface:ShapeProps@property:h',
  'interface:TextPropsOptions@property:x',
  'interface:TextPropsOptions@property:y',
  'interface:TextPropsOptions@property:w',
  'interface:TextPropsOptions@property:h',
]);
```

Map position atoms to `SlideCoordinate`/`TransformInput`, shape atoms to `SlideModel.addShape`, and
text atoms to `SlideModel.addText`. Point code evidence at `slide-coordinate.internal.ts`, test and
control evidence at the exact adapter title, package evidence at
`shapeTextPercentageCoordinates`, and OOXML evidence at the SDK title.

- [ ] **Step 2: Update the initial-manifest test and confirm generated counts**

Change the immutable-manifest test from 8 to 20 entries and assert its sorted status vector contains
7 `supported`, 12 `deliberate-difference`, and 1 `defect-excluded`. Run:

```bash
pnpm --config.verify-deps-before-run=false test:pptxgenjs-surface-audit
pnpm --config.verify-deps-before-run=false audit:pptxgenjs:write
```

Inspect the JSON and Markdown for the exact expected counts, zero diagnostics, unchanged hashes,
and the absence of all 12 IDs from `incompleteIds`.

- [ ] **Step 3: Run final family gates**

```bash
pnpm --config.verify-deps-before-run=false test:pptxgenjs-surface-audit
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false test
node scripts/smoke-npm-package.mjs
git diff --check
```

The default completeness command remains expected to fail only because 1,754 unrelated atoms are
still open:

```bash
pnpm --config.verify-deps-before-run=false audit:pptxgenjs
```

Review generated artifact determinism by running `audit:pptxgenjs:write` twice and requiring a clean
second diff.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add scripts/pptxgenjs-surface-manifest.mjs \
  scripts/pptxgenjs-surface-audit-lib.test.mjs \
  docs/compatibility/pptxgenjs-surface-audit.json \
  docs/compatibility/pptxgenjs-surface-audit.md
git commit -m "docs: close shape and text percentage coordinates"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected divergence: `0 0`; only `.pnpm-store/` remains untracked.
