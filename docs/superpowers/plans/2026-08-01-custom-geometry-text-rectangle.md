# Custom Geometry Text Rectangle Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict DrawingML custom-geometry text rectangles for zero-input creation and lossless existing-deck editing without changing existing default snapshots.

**Architecture:** Extend the root `CustomGeometry` direct-state tree with one optional `textRectangle` object whose four required fields own `a:rect@l/t/r/b`. Keep normalization, default folding, rendering, parsing, semantic equality, and malformed-state gating in `custom-geometry.internal.ts`; reuse `SlideModel.addCustomShape()` / `ShapeModel.customGeometry` for all lifecycle behavior.

**Tech Stack:** TypeScript 5.9, Vitest, `@pptx/lossless-xml`, OPC package transactions, PptxGenJS 4.0.1 public-output fixtures, Node/browser package smoke tests, Artifact Tool, LibreOffice headless, Poppler, `pptx-inspect`, PowerPoint 2010 Open XML validation.

## Global Constraints

- Implement `docs/superpowers/specs/2026-08-01-custom-geometry-text-rectangle-design.md` exactly.
- Public names are `CustomGeometryTextRectangle` and optional `CustomGeometry.textRectangle`.
- Every rectangle requires exact own data `left`, `top`, `right`, and `bottom`.
- Reuse `CustomGeometryValue`; do not infer guide existence, resolved coordinates, rectangle ordering, text layout, or geometry bounds.
- Omitted and explicit `{ left: 'l', top: 't', right: 'r', bottom: 'b' }` normalize to no own `textRectangle` property.
- Render absent/default state with the existing canonical `<a:rect l="l" t="t" r="r" b="b"/>`.
- Do not add a geometry evaluator, text measurement/layout engine, connector creation/snapping, handle drag evaluation, or a partial rectangle editor.
- Keep malformed OOXML lossless and reject unsafe edits before package mutation.
- Execute inline in the root task; do not dispatch subagents.
- Never stage, delete, or modify `.pnpm-store/`.
- End every task with source review, isolated commit, push to `main`, fetch, and `HEAD...origin/main == 0 0`.

---

### Task 1: Public rectangle type, strict normalization, rendering, and equality

**Files:**
- Modify: `packages/model/src/custom-geometry.ts`
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Produces: `CustomGeometryTextRectangle`.
- Produces: optional `CustomGeometry.textRectangle`.
- Preserves: `normalizeCustomGeometry()`, `renderCustomGeometry()`, and `customGeometryEqual()` signatures.
- Preserves: every existing geometry without a non-default rect normalizes to the same snapshot and renders the same bytes.

- [ ] **Step 1: Add a failing typed rectangle fixture**

Add beside `connectionGeometry`:

```ts
const textRectangleGeometry: CustomGeometry = {
  guides: [
    { name: 'textLeft', formula: { operator: 'val', operands: [20_000] } },
    { name: 'textRight', formula: { operator: 'val', operands: [80_000] } },
  ],
  connectionSites: [{ angle: 0, position: { x: 'hc', y: 't' } }],
  textRectangle: {
    left: 'textLeft',
    top: 10_000,
    right: 'textRight',
    bottom: 90_000,
  },
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

Assert the normalized snapshot equals the fixture; the root and `textRectangle` are frozen; and mutation of the
caller-owned rectangle after normalization does not change the snapshot. Include a compile-only declaration of
`CustomGeometryTextRectangle` so the test is red until the public type exists.

- [ ] **Step 2: Lock default folding and accepted direct values**

Normalize all three cases and assert none owns `textRectangle`:

```ts
{ paths: [{ width: 1, height: 1, commands: [] }] }
{
  textRectangle: { left: 'l', top: 't', right: 'r', bottom: 'b' },
  paths: [{ width: 1, height: 1, commands: [] }],
}
```

The third case is the canonical geometry read from `renderCustomGeometry()`. Separately prove zero, negative,
large safe integers, guide/built-in tokens, mixed numeric/token fields, and XML metacharacters in tokens are
accepted. Do not assert `left <= right` or `top <= bottom`.

- [ ] **Step 3: Add failing strict input cases**

Cover non-object rectangle, missing each required field, extra/inherited/accessor/symbol fields, runtime
`undefined`, unsafe/NaN/fractional numbers, empty/decimal/whitespace/invalid-XML tokens, and object subclasses.
Assert accessors are never invoked and input failure does not mutate the caller. Follow the existing
descriptor-safe table-driven tests rather than adding a second validation harness.

- [ ] **Step 4: Verify the red state**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: TypeScript/assertion failure because `CustomGeometryTextRectangle` and `textRectangle` do not exist.

- [ ] **Step 5: Add the public type and root property**

In `custom-geometry.ts`, after `CustomGeometryConnectionSite`, add:

```ts
export interface CustomGeometryTextRectangle {
  readonly left: CustomGeometryValue;
  readonly top: CustomGeometryValue;
  readonly right: CustomGeometryValue;
  readonly bottom: CustomGeometryValue;
}
```

Add `readonly textRectangle?: CustomGeometryTextRectangle` after `connectionSites` and before `paths`.

- [ ] **Step 6: Implement descriptor-safe normalization and default folding**

Import the type, add `textRectangle` to `ROOT_KEYS`, and add:

```ts
const TEXT_RECTANGLE_KEYS = new Set(['left', 'top', 'right', 'bottom']);
const DEFAULT_TEXT_RECTANGLE: Readonly<CustomGeometryTextRectangle> = Object.freeze({
  left: 'l',
  top: 't',
  right: 'r',
  bottom: 'b',
});
```

Normalize after connection sites and before paths. `normalizeTextRectangle()` must call `readObject()` with the
same set for allowed and required keys, normalize every field with
`normalizeCustomGeometryValue(value, context, false)`, freeze the result, and return `undefined` when all four
values equal `DEFAULT_TEXT_RECTANGLE`. Include the property only when the normalized result is non-default.

- [ ] **Step 7: Render rectangle values canonically**

Replace the fixed rect literal with `renderTextRectangle(geometry.textRectangle, prefix)`. The helper must use
the default object when the property is absent, write attributes in `l/t/r/b` order, and call
`renderCustomGeometryValue()` for each value:

```xml
<a:rect l="textLeft" t="10000" r="textRight" b="90000"/>
```

Assert both `a:` and empty/alternate prefixes, numeric lexical output, XML escaping, and unchanged default
renderer bytes.

- [ ] **Step 8: Extend semantic equality**

Add `textRectanglesEqual()` after connection-site comparison. It must distinguish absent from non-default and
compare `left/top/right/bottom` in that order. Test absent, changed left/top/right/bottom, numeric-versus-token,
and explicit-default folding. The normalized default and omitted geometry must compare equal.

- [ ] **Step 9: Run focused gates**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm typecheck
git diff --check
```

Expected: focused suite and typecheck pass; existing path/formula/handle/connection snapshots and canonical XML
remain unchanged.

- [ ] **Step 10: Review, commit, and push Task 1**

```bash
git add packages/model/src/custom-geometry.ts packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: add custom geometry text rectangle codec"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 2: Strict OOXML reader and live whole-replacement editing

**Files:**
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 rectangle type, normalizer, renderer, default folding, and equality.
- Produces: `readCustomGeometry()` snapshots for valid non-default `a:rect`.
- Preserves: absent/default rect compatibility, exact no-op, alternate prefix, all existing guides/handles/sites.

- [ ] **Step 1: Add failing supported-reader fixtures**

Render `textRectangleGeometry` with `a:` and `d:` prefixes and assert identical readback. Add raw XML fixtures
for numeric `+0`, negative/large integers, guide/built-in tokens, escaped tokens, whitespace around elements,
absent rect, canonical default rect, and non-default mixed rect. Absent/default must not own `textRectangle`;
non-default must preserve all four normalized values.

- [ ] **Step 2: Add exact no-op and reset cases**

For lexical numeric and escaped-token fixtures, capture `xml.serialize()`, assign the normalized snapshot through
`replaceCustomGeometry()`, and require `false` plus identical bytes. Replace a non-default snapshot with an
omitted/default snapshot and require the canonical default rect; parse it again and require no own property.

- [ ] **Step 3: Add failing malformed rect fixtures**

Extend the malformed table with repeated rect, wrong-namespace rect, missing `l/t/r/b`, qualified lookalike
attribute, extra attribute, repeated lexical attribute, child element, non-whitespace text, unsafe integer,
fraction, empty/whitespace/decimal token, and invalid XML character. For each source assert reader `undefined`,
replacement throws `ModelParseError`, and serialized XML stays unchanged.

- [ ] **Step 4: Verify non-default rect is still red**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: non-default rect fixtures return `undefined` before parser implementation.

- [ ] **Step 5: Parse absent/default/non-default state**

Replace `isDefaultRectangle()` with:

```ts
function parseTextRectangle(
  children: readonly XmlElement[],
): CustomGeometryTextRectangle | null | undefined
```

Semantics are exact: `undefined` means malformed, `null` means absent/canonical default, and an object means
supported non-default state. Require at most one same-stage `rect`; one namespace-correct direct element; no
children/text; exact unqualified `l/t/r/b`; and `parseCustomGeometryValue(value, false)` for all four values.
Return `null` when the parsed result matches the Task 1 default helper.

In `parseCustomGeometryElement()`, parse after guides/handles/connections, reject `undefined`, and pass
`...(textRectangle ? { textRectangle } : {})` into `normalizeCustomGeometry()`.

- [ ] **Step 6: Run reader/edit isolation gates**

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm vitest run packages/model/src/model.test.ts -t "custom geometry"
pnpm typecheck
git diff --check
```

Expected: supported rect read/edit/reset/no-op passes; malformed ownership rejects with zero mutation.

- [ ] **Step 7: Review, commit, and push Task 2**

```bash
git add packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: edit custom geometry text rectangles"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 3: Public lifecycle, six-format, SDK, and PptxGenJS boundary

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public rectangle tree through existing custom geometry APIs.
- Proves: create/read/edit/reset/conversion/duplicate/rollback/write/reopen behavior.
- Proves: existing PptxGenJS output remains a default-without-own-property snapshot.

- [ ] **Step 1: Replace the old unsupported-rectangle model test**

Rename `reads custom connections and keeps non-default text rectangles unsupported` to describe supported
connections and text rectangles. After injecting:

```xml
<a:rect l="0" t="t" r="r" b="b"/>
```

expect:

```ts
{
  ...customHandleGeometry,
  textRectangle: { left: 0, top: 't', right: 'r', bottom: 'b' },
}
```

Whole-replace it with a mixed numeric/token rectangle, assert the new snapshot, then reset with omitted/default
and assert no own property.

- [ ] **Step 2: Add model lifecycle tests across six formats**

Create a blank-deck shape with guides, handles, sites, paths, and a non-default rect. Assert immediate frozen
snapshot, then edit all four fields and preserve live identity, name, transform, fill, line, arrows, shadow,
hyperlink, shape XML text/effects/ext, relationships, and sibling order. Duplicate and mutate the copy to prove
source isolation. Roll back create/edit/reset in throwing outer transactions and verify package/journal/next ID.
Convert custom → preset → custom. Repeat write/reopen for `.pptx/.pptm/.ppsx/.ppsm/.potx/.potm`.

- [ ] **Step 3: Add SDK zero-input lifecycle coverage**

Use only public `PptxDocument.create()`, `addSlide()`, `addCustomShape()`, live
`ShapeModel.customGeometry`, `write()`, and `PptxDocument.open()`. Verify non-default create, edit, reset,
deep freeze, same-value exact no-op, preset/custom conversion, reopen, and 0 error diagnostics.

- [ ] **Step 4: Lock the PptxGenJS 4.0.1 boundary**

Extend the public custom-geometry boundary test to require:

```ts
expect(Object.hasOwn(importedShape.customGeometry!, 'textRectangle')).toBe(false);
expect(shapeXml(imported, 0, importedShape.id)).toMatch(
  /<a:rect\s+l="l"\s+t="t"\s+r="r"\s+b="b"\s*\/>/,
);
```

Keep the existing private-field source scan. Native non-default text rectangles are a DrawingML extension, not
a claimed PptxGenJS option.

- [ ] **Step 5: Run lifecycle gates**

```bash
pnpm vitest run packages/model/src/model.test.ts -t "text rectangle"
pnpm vitest run packages/sdk/src/index.test.ts -t "text rectangle"
pnpm vitest run packages/pptxgenjs-adapter/src/index.test.ts -t "custom geometry"
pnpm typecheck
git diff --check
```

- [ ] **Step 6: Review, commit, and push Task 3**

```bash
git add packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: verify custom geometry text rectangle lifecycle"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 4: Actual tarball Node/browser/types/CLI coverage

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: public package exports and existing custom geometry APIs.
- Produces: smoke JSON `customGeometryTextRectangles: true`.

- [ ] **Step 1: Add Node runtime smoke**

Create a geometry containing guides, a connection site, and:

```js
textRectangle: {
  left: 'textLeft',
  top: 12_500,
  right: 'textRight',
  bottom: 87_500,
},
```

Create/edit/reset/reopen it from the installed tarball. Assert exact snapshots, frozen rectangle, caller
detachment, same-value no-op through the public mutation journal behavior already exercised by the smoke,
preset/custom conversion, and 0 error diagnostics. Do not import workspace source.

- [ ] **Step 2: Add browser runtime smoke**

Extend the browser custom-geometry block to create/edit/reopen the mixed rectangle through `dist/browser.js`
and Blob output. Verify `Object.isFrozen()` for root and rectangle, exact JSON state, and default reset omission.

- [ ] **Step 3: Add consumer type coverage**

Import `CustomGeometryTextRectangle`; add valid numeric and token declarations and include one in typed
`CustomGeometry`. Add `@ts-expect-error` declarations for missing left/top/right/bottom, extra field, boolean
value, and runtime `undefined`.

- [ ] **Step 4: Expose the smoke result and run packed gates**

Add `customGeometryTextRectangles` to nested `api` and top-level JSON. Run:

```bash
pnpm typecheck
pnpm build
git diff --exit-code -- packages/pptx/dist
mkdir -p /tmp/pptx-text-rectangle-pack
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-text-rectangle-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-text-rectangle-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected JSON includes `"customGeometryTextRectangles":true`, `"types":true`, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, and push Task 4**

```bash
git add scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify packaged custom geometry text rectangles"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 5: Real-file gallery, previous client-normalized decks, documentation, and release gate

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Temporary only: scripts/output under `pptx_text_rect_dir=$(mktemp -d /tmp/pptx-custom-geometry-text-rectangles.XXXXXX)`

**Interfaces:**
- Documents: rectangle type, mapping, default folding, units/tokens, strict lifecycle and evaluator boundary.
- Proves: actual tarball output is valid, visibly meaningful, reopenable, and compatible with client normalization.
- Proves: prior LibreOffice rect insertion no longer blocks otherwise supported custom geometry.

- [ ] **Step 1: Load presentation tooling and verify `pptx-inspect`**

Read the current Presentations skill before artifact work, load the bundled workspace dependency paths, then run:

```bash
command -v pptx-inspect
pptx-inspect --json doctor
```

Expected: installed CLI and healthy offline runtime. Use the returned LibreOffice/Python/Poppler paths for the
remaining steps; do not assume system binaries.

- [ ] **Step 2: Generate a four-slide actual-tarball gallery**

Use Artifact Tool for the visual base and actual tarball APIs to convert existing text-bearing preset shapes to
custom geometry while preserving their `p:txBody`. Create:

1. default versus numeric inset text rectangle;
2. guide/token-backed left/right with numeric top/bottom;
3. asymmetric and negative/oversized direct rectangles;
4. whole-replacement edit/reset plus preserved connection sites.

Use visible text long enough to demonstrate wrapping and usable-region changes. Add audience-facing labels that
distinguish geometry text rectangles from text-body margins. Save:

```text
$pptx_text_rect_dir/text-rectangle-gallery.pptx
$pptx_text_rect_dir/source-snapshots.json
$pptx_text_rect_dir/text-rectangle-gallery.sha256
```

The actual tarball must strict-reopen every named target shape and reproduce exact snapshots.

- [ ] **Step 3: Render, inspect, and check overflow**

Render every page to PNG at 180 DPI, run the bundled `slides_test.py` overflow checker, and create a four-page
montage. Inspect each page at original resolution. Require readable labels, intentional wrap differences, no
clipping/overlap/off-slide content, and visually intact custom paths. Iterate the temporary gallery until all
checks pass.

- [ ] **Step 4: LibreOffice round-trip and structural comparison**

Save a copy with an isolated LibreOffice profile to:

```text
$pptx_text_rect_dir/roundtrip/text-rectangle-gallery.pptx
```

Reopen original and copy with the actual tarball. Write `$pptx_text_rect_dir/roundtrip-comparison.json`
containing slide/shape counts, names/order, guides, handles, connection sites, text rectangles, paths, commands,
and direct text content. Record any guide insertion, numeric/token conversion, rect normalization, wrap change,
or shape transform change without weakening the strict reader.

- [ ] **Step 5: Reopen the four prior LibreOffice-normalized galleries**

From the actual tarball, strict-open these exact retained files and write one comparison JSON:

```text
/private/tmp/pptx-custom-geometry-release.z6h4hN/gallery/lo-roundtrip/custom-geometry-gallery.pptx
/private/tmp/pptx-custom-geometry-formulas.Mr2XOu/lo-roundtrip-final2/guide-formula-gallery.pptx
/private/tmp/pptx-custom-geometry-handles-6rl3dd/roundtrip/custom-geometry-handles-gallery.pptx
/private/tmp/pptx-custom-geometry-connections-dNKjNK/roundtrip/gallery-original.pptx
```

For each target previously blocked only by non-default rect, require a defined `customGeometry` snapshot with
`textRectangle`. If another independent malformed state exists, record the exact raw OOXML reason and keep that
shape rejected; do not broaden ownership.

- [ ] **Step 6: Validate original and round-trip packages**

```bash
pptx-inspect --json package validate "$pptx_text_rect_dir/text-rectangle-gallery.pptx" --profile powerpoint-2010
pptx-inspect --json package validate "$pptx_text_rect_dir/roundtrip/text-rectangle-gallery.pptx" --profile powerpoint-2010
pptx-inspect --json package diff "$pptx_text_rect_dir/text-rectangle-gallery.pptx" "$pptx_text_rect_dir/roundtrip/text-rectangle-gallery.pptx"
```

Required: original and round-trip both report 0 errors and 0 warnings. The diff is evidence of client
normalization, not a zero-diff requirement.

- [ ] **Step 7: Update public documentation**

Document `CustomGeometryTextRectangle`, optional `textRectangle`, `left/top/right/bottom` to `l/t/r/b`, direct
coordinate integers, token preservation, default folding, strict malformed-state handling, caller detachment,
deep freeze, exact no-op, whole replacement, PptxGenJS 4.0.1 default-only boundary, LibreOffice normalization,
and the remaining geometry evaluator/connector boundary. Remove stale claims that non-default custom text
rectangles are unsupported.

- [ ] **Step 8: Run the complete release gate**

```bash
rg -n -i "text rectangles?.*unsupported|text rectangles?.*尚未支持|text rectangles?.*pending|custom text rectangle 为下一子项" \
  CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
pnpm typecheck
pnpm test
pnpm test:performance
pnpm build
git diff --exit-code -- packages/pptx/dist
node scripts/smoke-npm-package.mjs /tmp/pptx-text-rectangle-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

The stale-text search must return no matches. Expected: established skip count only, 1,000-part performance
under five seconds, reproducible dist, and actual tarball Node/browser/types/CLI success.

- [ ] **Step 9: Final review, commit, and push Task 5**

Stage only the five documentation files:

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document custom geometry text rectangles"
git push git@github.com:Xiejiayun/pptx.git main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected: divergence `0 0`; only `.pnpm-store/` remains untracked.

## Plan Self-Review

- Spec coverage: Tasks 1–2 implement type/codec/default folding/reader/editor; Task 3 proves public lifecycle and
  PptxGenJS boundary; Task 4 proves packaged runtimes/types; Task 5 proves real-file compatibility and docs.
- Scope: evaluator, text layout engine, connector creation/snapping, handle drag, relationships, and resolved
  coordinates remain explicitly excluded.
- Type consistency: every task uses `CustomGeometryTextRectangle` and `textRectangle` exactly.
- State consistency: all four fields are required; omitted/explicit default disappears; non-default remains exact.
- Mutation safety: invalid input and malformed OOXML have zero-mutation assertions before public release claims.
- Existing compatibility: renderer default bytes and old snapshots remain unchanged; prior LibreOffice rect
  normalization is reopened without relaxing unrelated strict checks.
- Placeholder scan: every command, file, type, fixture, expected result, and commit boundary is concrete.
- Execution: standing user delegation selects inline execution and forbids subagent dispatch.
