# Custom Geometry Connection Sites Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict ordered DrawingML custom-geometry connection sites for zero-input creation and lossless existing-deck editing.

**Architecture:** Extend the root `CustomGeometry` direct-state tree with an ordered `connectionSites` list whose entries own one required angle and position. Keep normalization, rendering, parsing, semantic equality, and unsupported-state gating in `custom-geometry.internal.ts`; reuse the existing `SlideModel.addCustomShape()` / `ShapeModel.customGeometry` transaction path for lifecycle behavior.

**Tech Stack:** TypeScript 5.9, Vitest, `@pptx/lossless-xml`, OPC package transactions, PptxGenJS 4.0.1 public-output fixtures, Node/browser package smoke tests, LibreOffice headless, PowerPoint 2010 Open XML validation.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-01-custom-geometry-connection-sites-design.md` exactly.
- Public names are `CustomGeometryConnectionSite` and optional ordered `CustomGeometry.connectionSites`.
- Every site requires exact own data `position` and `angle`; preserve source order and duplicates.
- Reuse `CustomGeometryValue` and `CustomGeometryPoint`; do not infer guide existence, angle range, resolved position, connector snapping, or relationships.
- Omitted and empty lists normalize to no own `connectionSites` property.
- Do not add custom text rectangles, geometry evaluation, connector creation, snapping, or resolved bounds.
- Keep unsupported OOXML lossless and reject unsafe edits before package mutation.
- Execute inline in the root task; do not dispatch subagents.
- Do not stage `.pnpm-store/`.
- End every task with review, isolated commit, push to `main`, fetch, and `HEAD...origin/main == 0 0`.

---

### Task 1: Public site type, strict normalization, rendering, and equality

**Files:**
- Modify: `packages/model/src/custom-geometry.ts`
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Produces: `CustomGeometryConnectionSite`.
- Produces: optional ordered `CustomGeometry.connectionSites`.
- Preserves: `normalizeCustomGeometry()`, `renderCustomGeometry()`, and `customGeometryEqual()` signatures.

- [ ] **Step 1: Add failing typed normalization fixtures**

Add beside `handleGeometry`:

```ts
const connectionGeometry: CustomGeometry = {
  adjustments: [
    { name: 'adjX', formula: { operator: 'val', operands: [25_000] } },
    { name: 'adjY', formula: { operator: 'val', operands: [60_000] } },
    { name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } },
  ],
  handles: [{
    kind: 'xy',
    position: { x: 'adjX', y: 'adjY' },
    xGuide: 'adjX',
    yGuide: 'adjY',
  }],
  connectionSites: [
    { angle: 0, position: { x: 'hc', y: 't' } },
    { angle: 'adjAng', position: { x: 'r', y: 'adjY' } },
    { angle: -5_400_000, position: { x: 25_000, y: 100_000 } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 0, y: 0 } },
      { kind: 'lineTo', point: { x: 100_000, y: 100_000 } },
    ],
  }],
};
```

Assert normalization detaches and recursively freezes the root, list, each site, and each position. Mutate
caller-owned arrays/sites/positions afterward and prove the snapshot stays unchanged. Assert omitted and
`connectionSites: []` normalize without an own property. Add compile-only declarations for the public type.

- [ ] **Step 2: Add failing strict input cases**

Cover non-object site/list, missing position, missing angle, unknown/inherited/accessor/symbol fields,
sparse/subclass arrays, point accessors, unsafe/NaN/fractional numbers, empty/decimal/whitespace/invalid-XML
tokens, and runtime `undefined`. Assert accessors are never invoked. Prove zero, negative and large numeric
angles, token angles, token/numeric coordinates, duplicate sites and arbitrary list order are accepted.

- [ ] **Step 3: Verify the red state**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: compile/assertion failure because the public site type and root property do not exist.

- [ ] **Step 4: Add the public type**

In `custom-geometry.ts`, after the handle union add:

```ts
export interface CustomGeometryConnectionSite {
  readonly position: CustomGeometryPoint;
  readonly angle: CustomGeometryValue;
}
```

Add `readonly connectionSites?: readonly CustomGeometryConnectionSite[]` after `handles` and before `paths`.

- [ ] **Step 5: Implement descriptor-safe normalization**

Import the type. Add `connectionSites` to `ROOT_KEYS`, plus:

```ts
const CONNECTION_SITE_KEYS = new Set(['position', 'angle']);
```

Normalize after handles and before paths. `normalizeConnectionSiteList()` must call `readArray()`, require
exact required keys through `readObject()`, normalize `position` through `normalizePoint()`, normalize `angle`
through `normalizeCustomGeometryValue(value, context, false)`, and freeze sites/list. Include the property only
when non-empty.

- [ ] **Step 6: Render sites canonically**

Replace fixed `<prefix:cxnLst/>` with `renderConnectionSiteList()`. Empty/absent stays self-closing. Non-empty
renders in exact order:

```xml
<a:cxnLst><a:cxn ang="0"><a:pos x="hc" y="t"/></a:cxn></a:cxnLst>
```

Use `renderCustomGeometryValue()` for angle and the existing focused point helper for `pos`.

- [ ] **Step 7: Extend semantic equality**

Compare optional list presence, length, order, every angle, and both position coordinates before paths.
Test absent list, reversed order, removed item, changed angle, changed x, and duplicate count.

- [ ] **Step 8: Run focused gates**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm typecheck
git diff --check
```

Expected: focused suite/typecheck pass and existing path/formula/handle snapshots stay unchanged.

- [ ] **Step 9: Review, commit, and push Task 1**

```bash
git add packages/model/src/custom-geometry.ts packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: add custom geometry connection site codec"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 2: Strict OOXML reader and live whole-replacement editing

**Files:**
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 site type, normalizer and renderer.
- Produces: `readCustomGeometry()` snapshots for supported non-empty `a:cxnLst`.
- Preserves: exact no-op, alternate prefix, supported handles, and non-default-rect isolation.

- [ ] **Step 1: Add failing supported-reader fixtures**

Render `connectionGeometry` with `a:` and `d:` prefixes and assert identical readback. Add lexical `+0`,
escaped angle/coordinate tokens, duplicate sites, whitespace, and empty/absent list fixtures. Assign the
normalized snapshot back and prove serialized bytes are unchanged.

- [ ] **Step 2: Add failing malformed fixtures**

Cover repeated list, list attribute/text, wrong-namespace/unknown child, missing/qualified/extra/repeated `ang`,
missing/repeated/wrong-namespace `pos`, extra child/text, missing/qualified/extra `x/y`, unsafe integer,
fractional number, whitespace/empty/decimal token, and invalid XML character. Retain non-default `rect`
rejection. For every case assert getter `undefined`, replacement throws `ModelParseError`, and bytes remain
unchanged.

- [ ] **Step 3: Verify non-empty sites are still red**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: non-empty connection lists return `undefined` before parser implementation.

- [ ] **Step 4: Parse the list and entries**

Remove `EMPTY_CUSTOM_LIST_NAMES` because all three optional lists now have focused parsers. In
`parseCustomGeometryElement()` call `parseConnectionSiteList(children)` and reject `undefined`. The parser
must accept absent/empty, reject repeated/malformed lists, preserve direct child order, require DrawingML
`cxn`, require exact unqualified `ang`, and require exactly one direct `pos`.

Parse `ang` with `parseCustomGeometryValue(value, false)` and `pos` with
`parsePointElement(element, 'pos')`. Return `{ angle, position }`. Pass a non-empty list to
`normalizeCustomGeometry()` as `connectionSites`.

- [ ] **Step 5: Run reader/edit isolation gates**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm vitest run packages/model/src/model.test.ts -t "custom geometry"
pnpm typecheck
git diff --check
```

Expected: supported sites read/edit/no-op; malformed site and custom-rect states reject without mutation.

- [ ] **Step 6: Review, commit, and push Task 2**

```bash
git add packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: edit custom geometry connection sites"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 3: Public lifecycle, six-format, SDK, and PptxGenJS boundary

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public connection-site tree through existing custom geometry APIs.
- Proves: create/read/edit/reorder/conversion/duplicate/rollback/write/reopen behavior.

- [ ] **Step 1: Add model lifecycle tests**

Create mixed numeric/token sites on a blank deck, assert immediate frozen snapshot, whole-replace with reordered
and changed sites, convert custom → preset → custom, and prove stable shape identity/name/transform/fill/line/
arrows/shadow/hyperlink/text/effects/relationships/sibling order. Duplicate and mutate the copy; prove source
isolation. Roll back an outer transaction and verify bytes/journal. Repeat write/reopen across
`.pptx/.pptm/.ppsx/.ppsm/.potx/.potm`.

- [ ] **Step 2: Replace the old unsupported-connection test**

Change the existing test that injects
`<a:cxn ang="0"><a:pos x="0" y="0"/></a:cxn>` to expect a supported snapshot. Keep the non-default text
rectangle case as `undefined` plus zero-mutation replacement rejection.

- [ ] **Step 3: Add SDK lifecycle coverage**

Use `PptxDocument.create()`, `addSlide()`, `addCustomShape()`, live `ShapeModel.customGeometry`, `write()`, and
`PptxDocument.open()` only. Verify create/edit/reorder/reopen, deep freeze, exact same-value no-op, custom/preset
conversion and 0 error diagnostics.

- [ ] **Step 4: Lock the PptxGenJS public-output boundary**

Generate legal `ShapeType.custGeom` through PptxGenJS 4.0.1 public API and assert imported snapshots do not own
`connectionSites`; `cxnLst` remains empty. Assert adapter production source contains no private-field lookup.

- [ ] **Step 5: Run lifecycle gates**

```bash
pnpm vitest run packages/model/src/model.test.ts -t "connection"
pnpm vitest run packages/sdk/src/index.test.ts -t "connection"
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "custom geometry"
pnpm typecheck
git diff --check
```

- [ ] **Step 6: Review, commit, and push Task 3**

```bash
git add packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: verify custom geometry connection site lifecycle"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 4: Actual tarball Node/browser/types/CLI coverage

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: public package exports and existing custom geometry APIs.
- Produces: smoke JSON `customGeometryConnectionSites: true`.

- [ ] **Step 1: Add Node runtime smoke**

Create/edit/reorder/reopen a mixed site list from the installed tarball. Assert angle/position token and numeric
values, deep freeze, exact snapshots, preset/custom conversion and diagnostics. Do not import workspace source.

- [ ] **Step 2: Add browser runtime smoke**

Extend the browser custom-geometry block to create/edit/reopen connection sites through `dist/browser.js` and
Blob output. Verify `Object.isFrozen()` for list/site/position and exact JSON state.

- [ ] **Step 3: Add consumer type fixture**

Import `CustomGeometryConnectionSite`; add valid numeric/token values and `@ts-expect-error` cases for missing
angle, missing position, extra field, unsafe angle type and runtime `undefined`. Include sites in typed geometry.

- [ ] **Step 4: Expose the smoke result and run packed gates**

Add `customGeometryConnectionSites` to nested `api` and top-level JSON. Run:

```bash
pnpm typecheck
pnpm build
git diff --exit-code -- packages/pptx/dist
mkdir -p /tmp/pptx-connection-site-pack
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-connection-site-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-connection-site-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

If pnpm emits another exact tarball filename, use that path. Expected JSON includes
`"customGeometryConnectionSites":true`, `"types":true`, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, and push Task 4**

```bash
git add scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify packaged custom geometry connection sites"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 5: Real-file gallery, documentation, validation, and release gate

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Temporary only: gallery scripts/output under a new `/tmp/pptx-custom-geometry-connections-*` directory

**Interfaces:**
- Documents: site type, ordered list, mappings, units/tokens, strict lifecycle and remaining boundary.
- Proves: actual tarball output is valid, renderable, reopenable, and client normalization is understood.

- [ ] **Step 1: Generate a four-slide actual-tarball gallery**

Use Artifact Tool for the visual base and the packed library for native site injection. Cover cardinal numeric
angles/positions; token angle/position references; ordered duplicate/mixed sites; create/edit/reorder and
preset/custom lifecycle. Audience-facing labels must explain that sites are editing metadata and are not
visible in slideshow output. Save source snapshots and SHA-256.

- [ ] **Step 2: Validate and inspect every slide**

Run overflow, render all pages, create a montage, and inspect every full-size slide. Fix clipping, overlap,
wrapping, unreadable labels and broken visible custom paths.

- [ ] **Step 3: Round-trip and structurally compare**

LibreOffice-save to a separate output directory. Reopen original and copy with the actual tarball; compare
slide/shape counts, names, adjustments, guides, handles, connection-site order/angle/position, paths and
commands. Inspect raw OOXML and record any rect insertion, token evaluation or metadata flattening without
broadening strict ownership.

- [ ] **Step 4: Validate both PPTX packages**

```bash
pptx-inspect --json package validate <original> --profile powerpoint-2010
pptx-inspect --json package validate <round-trip> --profile powerpoint-2010
pptx-inspect --json package diff <original> <round-trip>
```

Required: 0 errors and 0 warnings for both.

- [ ] **Step 5: Update public documentation**

Document `CustomGeometryConnectionSite`, ordered `connectionSites`, public-to-OOXML mapping, coordinate/angle
numeric units and token meaning, required fields, empty-list absence, strict parsing, detachment/freeze/no-op,
PptxGenJS 4.0.1 native-extension status, client normalization, and remaining custom-rect/evaluator boundary.
Remove every stale statement that connection sites are unsupported.

- [ ] **Step 6: Run the complete release gate**

```bash
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --exit-code -- packages/pptx/dist
node scripts/smoke-npm-package.mjs /tmp/pptx-connection-site-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected: established skip count only, 1,000-part performance under five seconds, reproducible dist, actual
tarball Node/browser/types/CLI success, and no stale unsupported-connection claim.

- [ ] **Step 7: Final review, commit, and push Task 5**

Stage only the five documentation files:

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document custom geometry connection sites"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

## Plan Self-Review

- Spec coverage: Tasks 1–2 implement exact type/codec/reader/editor; Task 3 proves lifecycle and PptxGenJS
  boundary; Task 4 proves packaged runtimes/types; Task 5 proves real-file compatibility and docs.
- Scope: custom text rectangle, evaluator, connector creation/snapping, relationships and resolved positions are
  explicitly excluded.
- Type consistency: every task uses `CustomGeometryConnectionSite` and `connectionSites` exactly.
- State consistency: each site requires angle + position; omitted/empty root list disappears; order and duplicates
  remain exact.
- Mutation safety: invalid input and unsafe OOXML have zero-mutation assertions before public release claims.
- Execution: standing user delegation selects inline execution and forbids subagent dispatch.
