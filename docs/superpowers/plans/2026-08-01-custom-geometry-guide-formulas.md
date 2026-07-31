# Custom Geometry Guide Formulas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict DrawingML custom-geometry adjustment/shape guide formulas and guide-referenced path values for zero-input creation and lossless existing-deck editing.

**Architecture:** Extend the existing `CustomGeometry` direct-state tree with typed formula tuples and numeric-or-token values. Keep all normalization, OOXML rendering/parsing, semantic equality, and unsupported-state gating in the focused custom-geometry codec; the existing public `SlideModel`/`ShapeModel` transaction flow then exposes creation and whole replacement without a second mutation path.

**Tech Stack:** TypeScript 5.9, Vitest, `@pptx/lossless-xml`, OPC package transactions, PptxGenJS 4.0.1 output fixtures, Node/browser package smoke tests, LibreOffice headless, PowerPoint 2010 Open XML validation.

## Global Constraints

- Implement the approved design in `docs/superpowers/specs/2026-08-01-custom-geometry-guide-formulas-design.md` exactly; do not add handles, connections, custom rect, or evaluation in this item.
- Execute inline in the root task because the user explicitly requested no subagents and delegated all implementation choices.
- Preserve numeric-only `CustomGeometry` source and snapshot compatibility: omitted/empty guide lists do not appear as own properties.
- Accept only finite safe-integer numbers or non-empty, XML-safe, no-whitespace, non-decimal string tokens.
- Require guide names to be globally unique across `adjustments` and `guides`; do not resolve dependencies or arithmetic domains yet.
- Keep existing unsupported OOXML lossless and reject unsafe edits before package mutation.
- Do not stage `.pnpm-store/`.
- End every task with review, an isolated commit, push to `main`, fetch, and `HEAD...origin/main == 0 0`.

---

### Task 1: Public formula types, strict normalization, rendering, and equality

**Files:**
- Modify: `packages/model/src/custom-geometry.ts`
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Produces: `CustomGeometryValue`, `CustomGeometryFormula`, `CustomGeometryGuide`.
- Produces: optional `CustomGeometry.adjustments` / `CustomGeometry.guides` and token-capable point/arc fields.
- Preserves: `normalizeCustomGeometry()`, `renderCustomGeometry()`, and `customGeometryEqual()` function signatures.

- [ ] **Step 1: Add failing public-type and normalization tests**

Add a `formulaGeometry` fixture covering all arities and path token consumers:

```ts
const formulaGeometry: CustomGeometry = {
  adjustments: [{
    name: 'adj1',
    formula: { operator: 'val', operands: [25_000] },
  }],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } },
    { name: 'y1', formula: { operator: '+-', operands: ['h', 0, 'x1'] } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 0 } },
      { kind: 'lineTo', point: { x: 'r', y: 'y1' } },
      {
        kind: 'arcTo',
        widthRadius: 'x1',
        heightRadius: 'hd2',
        startAngle: 0,
        sweepAngle: 'cd2',
      },
    ],
  }],
};
```

Assert normalization detaches and recursively freezes root/lists/guides/formulas/operands/paths;
empty guide lists normalize to absent properties; all 17 operators accept exact arity; duplicate names,
unknown/extra/accessor fields, sparse/subclass/wrong-length tuples, unsafe numbers, empty/whitespace/
decimal-string/XML-control tokens reject.

- [ ] **Step 2: Run focused tests and verify the red state**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: typecheck/transpile or assertions fail because formula types/root fields and token values are not implemented.

- [ ] **Step 3: Add the public type model**

In `custom-geometry.ts`, add:

```ts
export type CustomGeometryValue = number | string;

export type CustomGeometryFormula =
  | { readonly operator: 'val' | 'abs' | 'sqrt'; readonly operands: readonly [CustomGeometryValue] }
  | { readonly operator: 'at2' | 'cos' | 'max' | 'min' | 'sin' | 'tan'; readonly operands: readonly [CustomGeometryValue, CustomGeometryValue] }
  | { readonly operator: '*/' | '+-' | '+/' | '?:' | 'cat2' | 'mod' | 'pin' | 'sat2'; readonly operands: readonly [CustomGeometryValue, CustomGeometryValue, CustomGeometryValue] };

export interface CustomGeometryGuide {
  readonly name: string;
  readonly formula: CustomGeometryFormula;
}
```

Change point and arc value fields to `CustomGeometryValue`; add optional guide lists to
`CustomGeometry` without changing path `width/height`.

- [ ] **Step 4: Implement descriptor-safe formula normalization**

In `custom-geometry.internal.ts`:

```ts
const ROOT_KEYS = new Set(['adjustments', 'guides', 'paths']);
const FORMULA_KEYS = new Set(['operator', 'operands']);
const GUIDE_KEYS = new Set(['name', 'formula']);
const FORMULA_ARITIES = new Map<string, number>([
  ['val', 1], ['abs', 1], ['sqrt', 1],
  ['at2', 2], ['cos', 2], ['max', 2], ['min', 2], ['sin', 2], ['tan', 2],
  ['*/', 3], ['+-', 3], ['+/', 3], ['?:', 3], ['cat2', 3], ['mod', 3], ['pin', 3], ['sat2', 3],
]);
```

Add `normalizeGuideList`, `normalizeFormula`, and `normalizeCustomGeometryValue`. Reuse
`readObjectData`, `requireKeys`, and `readArray`; use the existing XML 1.0 character loop pattern.
Normalize empty lists to no own property and reject duplicate names with one `Set<string>` shared by
both lists.

- [ ] **Step 5: Render canonical guide XML and escaped token values**

Import `escapeXmlAttribute` from `@pptx/lossless-xml`. Render guide lists with:

```ts
function renderGuideList(
  name: 'avLst' | 'gdLst',
  guides: readonly Readonly<CustomGeometryGuide>[] | undefined,
  prefix: string,
): string {
  if (!guides?.length) return `<${prefix}${name}/>`;
  const children = guides.map((guide) =>
    `<${prefix}gd name="${escapeXmlAttribute(guide.name)}" ` +
    `fmla="${escapeXmlAttribute(renderFormula(guide.formula))}"/>`).join('');
  return `<${prefix}${name}>${children}</${prefix}${name}>`;
}
```

Use one ASCII space between operator/operands. Render number values as canonical decimal and string
values through `escapeXmlAttribute()` in point and arc attributes.

- [ ] **Step 6: Extend semantic equality**

Compare guide-list presence after empty normalization, ordered guide names, operator, operand tuples,
and every numeric-or-token path value. Ensure formulas differing only in a guide or operand compare false,
while separately normalized equal inputs compare true.

- [ ] **Step 7: Run focused tests and typecheck**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm typecheck
git diff --check
```

Expected: focused suite and typecheck pass; existing numeric-only snapshots remain unchanged.

- [ ] **Step 8: Review, commit, and push Task 1**

Review only the three task files and confirm every changed line maps to the formula codec. Then:

```bash
git add packages/model/src/custom-geometry.ts packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: add custom geometry guide formula codec"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 2: Strict OOXML reader and live whole-replacement editing

**Files:**
- Modify: `packages/model/src/custom-geometry.internal.ts`
- Test: `packages/model/src/custom-geometry.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 `CustomGeometryGuide`, `CustomGeometryFormula`, `CustomGeometryValue`.
- Produces: `readCustomGeometry()` snapshots for supported non-empty `avLst/gdLst` and token path values.
- Preserves: `replaceCustomGeometry()` exact semantic no-op and prefix/namespace behavior.

- [ ] **Step 1: Add failing strict-reader fixtures**

Create an alternate-prefix fixture containing non-empty `avLst/gdLst`, XML whitespace in `fmla`,
`+1` numeric lexical values, all path token positions, and escaped guide names. Assert the snapshot
normalizes to the Task 1 tree and assigning that snapshot through `replaceCustomGeometry()` does not
change serialized bytes.

Add malformed cases for duplicate/missing/qualified `name/fmla`, non-leaf `gd`, wrong namespace,
unknown operator, wrong arity, unsafe decimal, whitespace token, duplicate name across lists, token path
`w/h`, and extra list content. Retain explicit rejection tests for non-empty `ahLst/cxnLst` and non-default
`rect`.

- [ ] **Step 2: Run focused tests and verify failure**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
```

Expected: reader returns `undefined` for the new supported fixture before implementation.

- [ ] **Step 3: Parse guide lists and formulas**

Replace the blanket empty-list loop with separate handling:

```ts
for (const name of ['ahLst', 'cxnLst'] as const) {
  const lists = children.filter((child) => child.localName === name);
  if (lists.length > 1 || (lists[0] && !isEmptyElement(lists[0]))) return undefined;
}

const adjustments = parseGuideList(children, 'avLst');
const guides = parseGuideList(children, 'gdLst');
if (adjustments === undefined || guides === undefined) return undefined;
```

`parseGuideList()` must distinguish absent/empty success from malformed failure, parse exact direct `gd`
children, tokenize `fmla.trim().split(/[ \t\r\n]+/)`, validate operator arity, and convert signed decimal
operands with safe-integer checks.

- [ ] **Step 4: Parse numeric-or-token path values**

Add `parseCustomGeometryValue(value)`:

```ts
function parseCustomGeometryValue(value: string | undefined): CustomGeometryValue | undefined {
  if (value === undefined) return undefined;
  if (INTEGER_PATTERN.test(value)) return parseInteger(value, false);
  return isValidCustomGeometryToken(value) ? value : undefined;
}
```

Use it for point `x/y` and arc `wR/hR/stAng/swAng`; keep path `w/h` on positive `parseInteger`.

- [ ] **Step 5: Build the root snapshot and use normalization as the final gate**

Construct `{ ...(adjustments.length ? { adjustments } : {}), ...(guides.length ? { guides } : {}), paths }`
and call `normalizeCustomGeometry()` once so global uniqueness, detachment, and recursive freeze remain the
single source of truth. Catch failure and return `undefined` without mutation.

- [ ] **Step 6: Verify prefix-preserving edits and malformed-state isolation**

Run:

```bash
pnpm vitest run packages/model/src/custom-geometry.internal.test.ts
pnpm vitest run packages/model/src/model.test.ts -t "custom geometry"
pnpm typecheck
git diff --check
```

Expected: supported formulas read/edit/no-op; unsupported handles/connections/rect still reject with unchanged bytes.

- [ ] **Step 7: Review, commit, and push Task 2**

```bash
git add packages/model/src/custom-geometry.internal.ts packages/model/src/custom-geometry.internal.test.ts
git diff --cached --check
git commit -m "feat: edit custom geometry guide formulas"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 3: Public lifecycle, six-format, SDK, and PptxGenJS conformance coverage

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: public `SlideModel.addCustomShape()` and `ShapeModel.customGeometry` with formulas.
- Proves: create/read/edit/preset conversion/duplicate/rollback/write/reopen behavior and native-extension boundary.

- [ ] **Step 1: Add public create/edit and detachment tests**

Use a mutable formula geometry with at least one adjustment, two guides, all three arities, and guide
references in point and arc fields. Create a styled/hyperlinked custom shape, mutate every caller-owned
nested object afterward, and assert the live snapshot remains equal, detached, and recursively frozen.

Capture package bytes/journal before assigning `structuredClone(snapshot)` and prove exact no-op. Replace
one guide formula and path token, then prove shape identity, transform, fill, line, arrows, shadow, hyperlink,
text/effects/extensions, relationships, and neighboring shape bytes are preserved.

- [ ] **Step 2: Extend lifecycle coverage**

For every `PRESENTATION_FORMAT_PROFILES` entry, create a formula geometry, duplicate the slide, edit only
the source, move/delete the duplicate, write/reopen, and assert source/duplicate isolation and exact formulas.
Wrap formula creation and editing in throwing outer transactions and assert package snapshots and ID sequence
roll back.

- [ ] **Step 3: Cover preset/custom conversion and unsupported targets**

Convert a preset shape to formula custom geometry and back; assert live identity and unrelated state. Inject
non-empty handle, connection, and custom rect states separately; assert getter `undefined`, setter
`ModelParseError`, and no package mutation.

- [ ] **Step 4: Add SDK zero-input integration coverage**

In `packages/sdk/src/index.test.ts`, create `PptxDocument.create()`, add a formula custom shape, write,
reopen, edit a guide and token coordinate, write/reopen again, and assert `validatePackage()` has no errors.

- [ ] **Step 5: Preserve PptxGenJS 4.0.1 semantics**

In the existing custom-geometry adapter tests, assert legal PptxGenJS shapes still have no own
`adjustments/guides` properties and produce the same path snapshot. Keep malformed runtime outputs returning
`undefined`. Add an explicit assertion/comment that PptxGenJS exposes no public guide formula input, so no
production adapter change is expected.

- [ ] **Step 6: Run model, SDK, adapter, and full suites**

Run:

```bash
pnpm vitest run packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
pnpm test
pnpm typecheck
git diff --check
```

Expected: all tests pass with only the repository's established skip count.

- [ ] **Step 7: Review, commit, and push Task 3**

```bash
git add packages/model/src/model.test.ts packages/sdk/src/index.test.ts packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: verify custom geometry guide formula lifecycle"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 4: Actual tarball Node/browser/type/CLI verification

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves: aggregate package exports formula types and actual packed Node/browser builds preserve runtime behavior.
- Produces: `customGeometryGuideFormulas: true` in package-smoke JSON.

- [ ] **Step 1: Extend generated TypeScript consumer**

Import `CustomGeometryValue`, `CustomGeometryFormula`, and `CustomGeometryGuide`. Add valid typed values for
each arity and a token path. Add `@ts-expect-error` cases for unknown operator and incorrect tuple arity;
keep runtime-only lexical errors out of compile assertions.

- [ ] **Step 2: Extend actual Node package smoke**

Create a formula geometry through the installed tarball, verify recursive freeze/detachment, exact no-op,
formula/path edit, preset/custom conversion, write/reopen, and expected compact `avLst/gdLst` XML. Add
`customGeometryGuideFormulas` to the exported check object and final JSON summary.

- [ ] **Step 3: Extend browser-condition smoke**

Update the existing browser custom-geometry block to create/read/edit/reopen guide formulas and token arc
values through `dist/browser.js`. Require recursive frozen snapshots and formula equality.

- [ ] **Step 4: Run build and actual tarball smoke**

Run:

```bash
pnpm build
pnpm --filter @jiayunxie/pptx pack --pack-destination /tmp/pptx-guide-formula-pack
node scripts/smoke-npm-package.mjs /tmp/pptx-guide-formula-pack/jiayunxie-pptx-0.1.0.tgz
git diff --exit-code -- packages/pptx/dist
git diff --check
```

If the pack command emits a different exact tarball filename, pass the emitted path to the smoke script.
Expected JSON contains `"customGeometryGuideFormulas":true`, `"types":true`, and CLI version `0.1.0`.

- [ ] **Step 5: Review, commit, and push Task 4**

```bash
git add scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: verify packaged custom geometry guide formulas"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 5: Documentation, gallery, compatibility validation, and release gate

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Temporary only: a formula gallery script and outputs under a new `/tmp/pptx-custom-geometry-formulas-*` directory

**Interfaces:**
- Documents: formula/value types, units/tokens, supported operators, strict boundaries, and PptxGenJS extension status.
- Proves: generated PPTX is valid, renderable, reopenable, and round-trip behavior is understood.

- [ ] **Step 1: Generate a representative formula gallery from the actual tarball**

Create a 16:9 deck containing:

1. `val`, `*/`, `+-`, `+/`, `?:` with guide-referenced polygon points;
2. `abs`, `sqrt`, `max`, `min`, `pin`, `mod` with curves;
3. `at2`, `sin`, `cos`, `tan`, `cat2`, `sat2` with arc radii/angles;
4. formula edit and preset/custom conversion lifecycle evidence.

Use labels that distinguish native formula support from client rendering behavior. Save source snapshots and
SHA-256 next to the temporary deck.

- [ ] **Step 2: Validate and visually inspect every slide**

Run repository render and overflow helpers, inspect every full-size slide, and fix unintended clipping,
overlap, broken paths, or unreadable labels. Export PDF through LibreOffice, render all pages, and inspect them.

- [ ] **Step 3: Round-trip and structurally compare**

LibreOffice-save a copy, reopen both with the actual tarball, compare slide/shape counts, names, guide lists,
formulas, path tokens, and commands. Record any LibreOffice normalization separately from library behavior;
do not broaden strict parsing merely to accept semantically ambiguous client output.

- [ ] **Step 4: Run PowerPoint 2010 validation**

Validate original and LibreOffice round-trip PPTX with the same Open XML SDK/PowerPoint 2010 profile used by
the previous custom-path item. Required result: 0 errors / 0 warnings for both files. Treat any schema error
as a release blocker and fix the emitted OOXML before proceeding.

- [ ] **Step 5: Update public documentation**

Document `CustomGeometryValue`, `CustomGeometryFormula`, `CustomGeometryGuide`, operator arities,
adjustments/guides, token references, strict lexical rules, detachment/freeze/no-op, and remaining handles/
connections/rect/evaluation boundary. Update the compatibility row so formulas are complete and handles are
next; state that PptxGenJS 4.0.1 has no public guide-formula input.

- [ ] **Step 6: Run the complete release gate**

Run:

```bash
pnpm typecheck
pnpm test
RUN_PERF=1 pnpm vitest run packages/testkit/src/performance.test.ts
pnpm build
git diff --exit-code -- packages/pptx/dist
node scripts/smoke-npm-package.mjs /tmp/pptx-guide-formula-pack/jiayunxie-pptx-0.1.0.tgz
git diff --check
```

Expected: full suite passes with established skips, 1,000-part performance stays within its 5-second budget,
dist is reproducible, actual tarball Node/browser/types/CLI checks pass, and docs have no stale unsupported claim.

- [ ] **Step 7: Final review, commit, and push Task 5**

Stage only the five documentation files:

```bash
git add CHANGELOG.md README.md packages/pptx/README.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md
git diff --cached --check
git commit -m "docs: document custom geometry guide formulas"
git push git@github.com:Xiejiayun/pptx.git main:main
git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
git status --short
```

Expected divergence: `0 0`; only `.pnpm-store/` may remain untracked.

## Plan Self-Review

- Spec coverage: Tasks 1–2 implement the complete typed formula/value codec and reader/editor; Task 3 proves
  lifecycle and PptxGenJS boundaries; Task 4 proves packed runtimes/types; Task 5 proves real-file compatibility
  and documentation.
- Scope: handles, connections, custom rect, and evaluator are explicitly excluded and remain losslessly gated.
- Type consistency: all tasks use `CustomGeometryValue`, `CustomGeometryFormula`, `CustomGeometryGuide`,
  `adjustments`, and `guides` exactly as defined in the design.
- Mutation safety: every invalid-input and unsafe-existing-state path has a zero-mutation assertion before
  public release claims.
- Execution: standing user delegation selects inline execution; no additional choice or review pause is needed.
