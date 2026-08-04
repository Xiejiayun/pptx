# Image Percentage Coordinates Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict slide-relative percentage coordinates to raster and SVG image creation, then close the four direct PptxGenJS 4.0.1 `ImageProps.x/y/w/h` audit atoms with runtime, package, OOXML, and control evidence.

**Architecture:** Reuse `resolveSlideCoordinate()` inside the existing raster appearance normalizer that SVG already delegates to, and pass the current `PresentationModel.slideSize` from both slide image creation paths. Preserve concrete EMU transforms, placeholder precedence, source sizing behavior, and async source-loader mutation isolation; nested image sizing coordinates remain a separate audit item.

**Tech Stack:** TypeScript 5.9, Vitest, Node.js 22, PptxGenJS 4.0.1, OOXML `p:pic` / `a:off` / `a:ext`, pnpm packed-consumer smoke, and deterministic surface audit generation.

## Global Constraints

- The approved design is `docs/superpowers/specs/2026-08-04-image-percentage-coordinates-design.md`.
- PptxGenJS 4.0.1 is the only control baseline.
- Public `Transform`, `ImageModel.transform`, and `setTransform()` remain absolute EMU APIs.
- Existing numeric image inputs remain absolute EMU; callers use `inches()` for explicit inch conversion.
- Percentages resolve against current slide width for `x`/`width` and height for `y`/`height`.
- Position may be negative or exceed the slide; width and height must resolve to positive safe integers.
- Reject malformed, non-finite, whitespace-padded, accessor-backed, and unsafe coordinate inputs before package mutation.
- Placeholder-owner transforms continue to override caller geometry.
- `ImageSizing` frame values remain absolute EMU, crop source values remain source-image pixels, and sizing continues to own final width/height.
- Every task ends with focused review, exact-file staging, commit, push, remote-tracking refresh, and `HEAD...origin/main` divergence `0 0`.
- Never stage `.pnpm-store/`.

---

### Task 1: Raster and SVG Model Creation

**Files:**
- Modify: `packages/model/src/image.ts`
- Modify: `packages/model/src/image-create.internal.ts`
- Modify: `packages/model/src/image-create.internal.test.ts`
- Modify: `packages/model/src/svg-image-create.internal.ts`
- Modify: `packages/model/src/svg-image-create.internal.test.ts`
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `TransformInput`, `SlideSize`, and `resolveSlideCoordinate()` from the completed shared coordinate work.
- Produces:

```ts
export interface AddImageOptions extends Partial<TransformInput> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly sourceRectangle?: ImageSourceRectangle;
}

export interface AddSvgImageOptions extends Partial<TransformInput> {
  readonly name?: string;
  readonly altText?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly sourceRectangle?: ImageSourceRectangle;
}

export function normalizeEmbeddedRasterImage(
  bytes: unknown,
  options: unknown,
  slideSize?: Readonly<SlideSize>,
): NormalizedEmbeddedRasterImage;

export function normalizeEmbeddedSvgImage(
  svgBytes: unknown,
  fallbackPngBytes: unknown,
  options?: unknown,
  slideSize?: Readonly<SlideSize>,
): NormalizedEmbeddedSvgImage;
```

- [ ] **Step 1: Write failing raster and SVG normalizer tests**

Add a 10 × 8 inch size and exact assertions:

```ts
const SIZE = Object.freeze({ width: inches(10), height: inches(8) });

expect(normalizeEmbeddedRasterImage(Uint8Array.of(1), {
  contentType: 'image/png',
  x: '10%',
  y: '20%',
  width: '30%',
  height: '40%',
}, SIZE)).toMatchObject({
  x: inches(1),
  y: inches(1.6),
  width: inches(3),
  height: inches(3.2),
});

expect(normalizeEmbeddedSvgImage(
  Uint8Array.of(1),
  Uint8Array.of(2),
  { x: '12.5%', y: '25%', width: '37.5%', height: '50%' },
  SIZE,
)).toMatchObject({
  x: inches(1.25),
  y: inches(2),
  width: inches(3.75),
  height: inches(4),
});
```

Also assert absolute `inches()` values remain unchanged, percentage input without a slide size throws, and zero/negative percentage extents, malformed strings, and coordinate accessors throw without reading accessors.

- [ ] **Step 2: Run the focused tests and confirm failure**

```bash
pnpm --config.verify-deps-before-run=false vitest run \
  packages/model/src/image-create.internal.test.ts \
  packages/model/src/svg-image-create.internal.test.ts \
  -t "percentage coordinates"
```

Expected: FAIL because image options accept only `Transform` and the normalizers do not receive a slide size.

- [ ] **Step 3: Integrate the shared resolver once**

Change both public image option interfaces to `Partial<TransformInput>`. In
`normalizeEmbeddedRasterImage()`, replace the four transform `normalizeInteger()` calls with:

```ts
const x = resolveSlideCoordinate(
  values.x, 'horizontal', slideSize, 0 as Emu, 'Embedded raster image x',
);
const y = resolveSlideCoordinate(
  values.y, 'vertical', slideSize, 0 as Emu, 'Embedded raster image y',
);
const width = resolveSlideCoordinate(
  values.width, 'horizontal', slideSize, EMU_PER_INCH, 'Embedded raster image width',
);
const height = resolveSlideCoordinate(
  values.height, 'vertical', slideSize, EMU_PER_INCH, 'Embedded raster image height',
);
```

Keep the existing positive extent checks. Pass `slideSize` through
`normalizeEmbeddedSvgImage()` to the raster normalizer. Pass
`this.presentation.slideSize` from both `SlideModel.addImage()` and `addSvgImage()`.

- [ ] **Step 4: Lock placeholder precedence with percentage caller geometry**

In the existing model test `populate image chart placeholder owners in place through the model API`, replace the caller image `x` and `width` overrides with `'90%'` while keeping assertions against the layout owner transform. This proves percentage normalization does not outrank placeholder geometry.

- [ ] **Step 5: Run model regression gates**

```bash
pnpm --config.verify-deps-before-run=false vitest run \
  packages/model/src/image-create.internal.test.ts \
  packages/model/src/svg-image-create.internal.test.ts \
  packages/model/src/model.test.ts \
  -t "percentage coordinates|placeholder owners"
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
```

Review default one-inch extents, raster/SVG shared behavior, placeholder precedence, strict option descriptors, source rectangle preservation, and absence of percentage changes to output `Transform`.

- [ ] **Step 6: Commit, push, and verify synchronization**

```bash
git add packages/model/src/image.ts \
  packages/model/src/image-create.internal.ts \
  packages/model/src/image-create.internal.test.ts \
  packages/model/src/svg-image-create.internal.ts \
  packages/model/src/svg-image-create.internal.test.ts \
  packages/model/src/slide.ts \
  packages/model/src/model.test.ts
git commit -m "feat: create images with percentage coordinates"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 2: High-Level Raster and SVG Source Loader

**Files:**
- Modify: `packages/sdk/src/raster-image-source.ts`
- Modify: `packages/sdk/src/raster-image-source.test.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 1 `AddSvgImageOptions` percentage fields and existing `commitPreparedImage()` flow.
- Produces: `AddImageSourceOptions` top-level `width?: SlideCoordinate` and
  `height?: SlideCoordinate` when `sizing` is absent.

- [ ] **Step 1: Add failing public-type and normalization tests**

Require the source-option normalizer to preserve detached percentage values:

```ts
expect(normalizeAddImageSourceOptions({
  x: '10%', y: '20%', width: '30%', height: '40%',
}).imageOptions).toEqual({
  x: '10%', y: '20%', width: '30%', height: '40%',
});
```

In the SDK declaration block, require this to compile:

```ts
const percentageImage: AddImageSourceOptions = {
  x: '10%',
  y: '20%',
  width: '30%',
  height: '40%',
};
```

Keep the existing `@ts-expect-error` that forbids any top-level width/height together with
`sizing`.

- [ ] **Step 2: Add a failing end-to-end source-loader test**

Create a 10 × 8 inch presentation. Add one PNG and one SVG through
`PptxDocument.addImage()` using different percentage rectangles and an explicit PNG fallback for
SVG. Assert exact immediate transforms, exact slide XML `a:off`/`a:ext`, zero PowerPoint 2010
errors, reopened transforms, and unchanged package snapshot/output after malformed and non-positive
percentage extents.

- [ ] **Step 3: Widen only the explicit top-level rectangle type**

Import `SlideCoordinate` into `raster-image-source.ts` and change only the no-sizing union branch:

```ts
| {
    readonly sizing?: undefined;
    readonly width?: SlideCoordinate;
    readonly height?: SlideCoordinate;
  }
```

Do not change `ImageSizing`, crop source coordinates, or the conflict check. The runtime normalizer
continues to copy data properties without evaluating coordinate strings; Task 1 resolves them after
the source is prepared and before package mutation.

- [ ] **Step 4: Run source-loader and SDK gates**

```bash
pnpm --config.verify-deps-before-run=false vitest run \
  packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/index.test.ts \
  -t "percentage coordinates|source options"
pnpm --config.verify-deps-before-run=false typecheck
git diff --check
```

Review raster/SVG coverage, source loader detachment, sizing exclusion, error-after-source but
before-mutation behavior, reopened SVG pairing, and exact current slide dimensions.

- [ ] **Step 5: Commit, push, and verify synchronization**

```bash
git add packages/sdk/src/raster-image-source.ts \
  packages/sdk/src/raster-image-source.test.ts \
  packages/sdk/src/index.test.ts
git commit -m "feat: load images with percentage coordinates"
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
- Consumes: locked PptxGenJS 4.0.1 `ImageProps`, valid PNG data, and packed public declarations/runtime.
- Produces: direct output comparison and `imagePercentageCoordinates: true` packed evidence.

- [ ] **Step 1: Add the PptxGenJS image control**

Add a test named
`matches PptxGenJS image percentage coordinate output with explicit native units`. Use a 10 × 8
layout and the same valid PNG for both libraries. PptxGenJS receives
`x/y/w/h = 10%/20%/30%/40%`; native receives `x: inches(1)` plus percentage
`y/width/height`. Require both final transforms to equal:

```ts
{
  x: 914_400,
  y: 1_463_040,
  width: 2_743_200,
  height: 2_926_080,
  rotation: 0,
  flipHorizontal: false,
  flipVertical: false,
}
```

Write/reopen native output and require zero PowerPoint 2010 error diagnostics.

- [ ] **Step 2: Add packed runtime and declaration proof**

In the packed consumer, create a custom 10 × 8 document, add percentage PNG and SVG images,
write/reopen, and define `imagePercentageCoordinates` only when all immediate and reopened
transforms match. Add it to required `apiChecks` and the final summary. Compile all three public
types:

```ts
const typedImageCoordinate: SlideCoordinate = '12.5%';
const typedImageSourcePercent: AddImageSourceOptions = {
  x: '10%', y: '20%', width: '30%', height: '40%',
};
const typedLowLevelImagePercent: AddImageOptions = {
  contentType: 'image/png',
  x: typedImageCoordinate,
  y: '20%',
  width: '30%',
  height: '40%',
};
const typedSvgImagePercent: AddSvgImageOptions = {
  x: '12.5%', y: '25%', width: '37.5%', height: '50%',
};
```

- [ ] **Step 3: Run control, type, and actual tarball gates**

```bash
pnpm --config.verify-deps-before-run=false vitest run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "image percentage coordinate output"
pnpm --config.verify-deps-before-run=false typecheck
image_pack_dir=$(mktemp -d /tmp/pptx-image-percentage.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$image_pack_dir"
image_tarball=$(find "$image_pack_dir" -maxdepth 1 -name '*.tgz' -print -quit)
node scripts/smoke-npm-package.mjs "$image_tarball"
git diff --check
```

Review public PptxGenJS API use, exact custom layout, valid PNG source, no private-field reads,
packed type acceptance, raster/SVG reopen transforms, diagnostics, and no workspace-source fallback.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add packages/pptxgenjs-adapter/src/index.test.ts scripts/smoke-npm-package.mjs
git commit -m "test: prove image percentage parity"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 4: Close Four Image Audit Atoms

**Files:**
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `scripts/pptxgenjs-surface-audit-lib.test.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`

**Interfaces:**
- Consumes: Tasks 1–3 code, SDK OOXML test, PptxGenJS control, and package smoke.
- Produces: four `deliberate-difference` entries and counts
  `supported=7`, `deliberate-difference=16`, `defect-excluded=1`,
  `unverified=1750`, `unsupported=0`, `stale=0`.

- [ ] **Step 1: Add the exact image coordinate entries**

Define the atom list once:

```js
const IMAGE_COORDINATE_ATOMS = Object.freeze([
  'interface:ImageProps@property:x',
  'interface:ImageProps@property:y',
  'interface:ImageProps@property:w',
  'interface:ImageProps@property:h',
]);
```

Map every atom to `SlideModel.addImage`, `SlideModel.addSvgImage`, and
`PptxDocument.addImage`. Point code evidence to `normalizeEmbeddedRasterImage`, test/control to the
exact adapter title, package evidence to `const imagePercentageCoordinates =`, and OOXML evidence to
the exact SDK test title. State that native uses explicit names/units and keeps nested sizing as a
separate open item.

- [ ] **Step 2: Update manifest expectations and regenerate artifacts**

Change the immutable manifest test from 20 to 24 entries and require 7 `supported`, 16
`deliberate-difference`, and 1 `defect-excluded`. Run:

```bash
pnpm --config.verify-deps-before-run=false test:pptxgenjs-surface-audit
pnpm --config.verify-deps-before-run=false audit:pptxgenjs:write
```

Inspect both artifacts for exact expected counts, zero diagnostics, unchanged declaration/runtime
hashes, and absence of all four image IDs from `incompleteIds`.

- [ ] **Step 3: Run final family gates**

```bash
pnpm --config.verify-deps-before-run=false test:pptxgenjs-surface-audit
pnpm --config.verify-deps-before-run=false typecheck
pnpm --config.verify-deps-before-run=false test
final_image_pack_dir=$(mktemp -d /tmp/pptx-image-final.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$final_image_pack_dir"
final_image_tarball=$(find "$final_image_pack_dir" -maxdepth 1 -name '*.tgz' -print -quit)
node scripts/smoke-npm-package.mjs "$final_image_tarball"
pnpm --config.verify-deps-before-run=false audit:pptxgenjs:write
pnpm --config.verify-deps-before-run=false audit:pptxgenjs:write
git diff --check
```

The default completeness check must exit 1 only because 1,750 unrelated atoms remain unverified:

```bash
pnpm --config.verify-deps-before-run=false audit:pptxgenjs
```

Review generated artifact determinism, all focused/full gates, package output
`imagePercentageCoordinates: true`, and unchanged statuses for nested image sizing atoms.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add scripts/pptxgenjs-surface-manifest.mjs \
  scripts/pptxgenjs-surface-audit-lib.test.mjs \
  docs/compatibility/pptxgenjs-surface-audit.json \
  docs/compatibility/pptxgenjs-surface-audit.md
git commit -m "docs: close image percentage coordinates"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git refs/heads/main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short --branch
```

Expected divergence: `0 0`; only `.pnpm-store/` remains untracked.
