# Table-Cell Text Style Defaults Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. This repository task executes inline without subagents and performs review, commit, and push after every independently testable deliverable.

**Goal:** Add strict table-level and cell-level font family, font size, bold, color, and paragraph-spacing creation defaults whose resolved direct state remains readable and editable.

**Architecture:** Extend the shared rich-text renderer with optional run defaults while preserving byte-identical output when they are omitted. Normalize table and cell fields independently, resolve table → cell defaults before rendering, then let explicit paragraph/run fields win through the existing rich-text precedence engine; do not retain creation metadata.

**Tech Stack:** TypeScript 5.8, Vitest 3, lossless DrawingML/OPC transactions, pnpm workspaces, PptxGenJS 4.0.1 conformance fixtures, packed Node/browser exports, CLI/Inspector, Google Chrome, and PowerPoint 2010 validation.

## Global Constraints

- Public fields are exactly `fontFamily?: string`, `fontSize?: number`, `bold?: boolean`, `color?: RichTextColor`, and `spacing?: ParagraphSpacing` on both `AddTableOptions` and `AddTableCellOptions`.
- Native API does not accept PptxGenJS aliases `fontFace`, `paraSpaceBefore`, `paraSpaceAfter`, `lineSpacing`, or `lineSpacingMultiple`.
- Resolution order is table defaults → cell defaults → explicit paragraph/run fields; explicit `bold: false`, paragraph `spacing: false`, and paragraph `line: false` must block inherited values.
- Spacing overlays independently by `before`, `after`, and `line`; scalar run fields overlay independently by field.
- All inputs are strict, descriptor-safe, detached, and validated before observable package mutation.
- Omitted and runtime-`undefined` defaults preserve all legacy bytes.
- Linked runs without an explicit run color skip a defined outer table/cell color; explicit run color still wins.
- Table-cell `endParaRPr` materializes resolved outer font family/font size; bold/color require a direct run.
- Creation stores only direct OOXML state. `setCellRichText()` never reapplies creation defaults; safe `setCellText()` preserves the current direct run template.
- Do not add retained style metadata, effective theme/table-style resolution, new table/cell font getters/editors, merge/span, row/column CRUD, auto-page, repeated headers, content measurement, or `tableToSlides`.
- Every task is reviewed, committed, pushed, and followed by local/remote divergence verification of `0 0`.

---

## File Structure

- Modify `packages/model/src/rich-text.internal.ts`: optional font family/font size/bold defaults, hyperlink-aware outer color suppression, and table-cell end-paragraph font defaults.
- Modify `packages/model/src/rich-text.internal.test.ts`: focused renderer precedence, hyperlink color, empty paragraph, and legacy parity tests.
- Modify `packages/model/src/table-create.internal.ts`: strict table/cell field normalization, table → cell resolution, spacing overlay, and renderer handoff.
- Modify `packages/model/src/table-create.internal.test.ts`: detached normalization, strict invalid input, precedence, rich/empty output, and byte-parity tests.
- Modify `packages/model/src/slide.ts`: public option types only; existing add-table transaction and relationship preparation remain unchanged.
- Modify `packages/model/src/model.test.ts`: immediate snapshot, editing, rollback, duplicate, six-format, and write/reopen lifecycle coverage.
- Modify `packages/sdk/src/index.test.ts`: SDK/root positive and negative TypeScript contract plus runtime creation.
- Modify `packages/pptxgenjs-adapter/src/index.test.ts`: PptxGenJS 4.0.1 final-state conformance and strict-difference coverage.
- Modify `README.md`, `packages/pptx/README.md`, `docs/api/README.md`, `docs/compatibility/pptxgenjs-baseline.md`, and `CHANGELOG.md`: API, support boundary, parity notes, proof, and remaining work.

### Task 1: Shared Rich-Text Renderer Defaults

**Files:**
- Modify: `packages/model/src/rich-text.internal.ts`
- Modify: `packages/model/src/rich-text.internal.test.ts`

**Interfaces:**
- Consumes: `NormalizedRichTextParagraph`, `RichTextColor`, `NormalizedParagraphSpacingUpdate`, existing hyperlink resolution, and the existing default-language/default-color renderer path.
- Produces: optional `RenderRichTextOptions.defaultFontFamily`, `defaultFontSize`, `defaultBold`, and `suppressDefaultColorForHyperlinks` behavior; no public exports change.

- [ ] **Step 1: Write failing run-default and legacy-parity tests**

Add one focused suite that normalizes these paragraphs and renders them with defaults:

```ts
const paragraphs = normalizeRichText([{
  runs: [
    { text: 'Inherited' },
    {
      text: 'Local false',
      style: {
        fontFamily: 'Arial',
        fontSize: 10,
        bold: false,
        color: { kind: 'srgb', value: 'FF0000' },
      },
    },
    { text: 'Linked', style: { hyperlink: { url: 'https://example.com' } } },
  ],
}, { runs: [] }]);

const rendered = renderRichTextParagraphs(paragraphs, {
  defaultFontFamily: 'Aptos',
  defaultFontSize: 18.25,
  defaultBold: true,
  defaultColor: { kind: 'scheme', value: 'accent1' },
  suppressDefaultColorForHyperlinks: true,
  runHyperlinkRelationshipIds: [[undefined, undefined, 'rId7'], []],
});
```

Assert inherited run `sz="1825" b="1"`, three Aptos font children, and accent1 fill; local run `sz="1000" b="0"`, Arial, and FF0000; linked run inherits Aptos/18.25/bold but has no direct `solidFill`; `endParaRPr` carries `sz="1825"` and three Aptos font children. Render the same paragraphs without new options and assert exact equality to a stored baseline generated before the feature.

- [ ] **Step 2: Run the focused test and verify the red state**

Run:

```bash
pnpm exec vitest run packages/model/src/rich-text.internal.test.ts \
  -t "run style defaults"
```

Expected: TypeScript/test failure because the four new renderer options do not exist or do not affect output.

- [ ] **Step 3: Extend the private renderer options and use one run-options object**

Add these fields to private `RenderRichTextOptions`:

```ts
readonly defaultFontFamily?: string;
readonly defaultFontSize?: number;
readonly defaultBold?: boolean;
readonly suppressDefaultColorForHyperlinks?: boolean;
```

Replace the expanding positional `renderRun()` parameters with one private object:

```ts
interface RenderRunOptions {
  readonly prefix: string;
  readonly defaultLanguage?: string;
  readonly defaultFontFamily?: string;
  readonly defaultFontSize?: number;
  readonly defaultBold?: boolean;
  readonly defaultColor?: Readonly<RichTextColor>;
  readonly hyperlink?: NormalizedHyperlink;
  readonly hyperlinkRelationshipId?: string;
  readonly declareHyperlinkRelationshipNamespace: boolean;
  readonly suppressDefaultColorForHyperlinks: boolean;
}
```

Resolve the three scalar defaults with nullish precedence so explicit `false` survives:

```ts
const fontFamily = style.fontFamily ?? options.defaultFontFamily;
const fontSize = style.fontSize ?? options.defaultFontSize;
const bold = style.bold ?? options.defaultBold;
```

For color, distinguish a suppressed linked default from ordinary canonical `tx1`:

```ts
const suppressOuterColor = options.suppressDefaultColorForHyperlinks
  && options.hyperlink !== undefined
  && style.color === undefined;
const color = style.color
  ?? (suppressOuterColor ? undefined : options.defaultColor)
  ?? (suppressOuterColor ? undefined : { kind: 'scheme' as const, value: 'tx1' });
const solidFill = color === undefined
  ? ''
  : `<${prefix}solidFill>${renderMainTextColorChoice(
      color,
      prefix,
      style.transparency,
    )}</${prefix}solidFill>`;
```

Render font children from the resolved family and attributes from resolved size/bold. Keep all other style and hyperlink semantics unchanged.

- [ ] **Step 4: Materialize end-paragraph font defaults without changing omitted output**

Create a private helper that returns the existing self-closing value when both defaults are absent and otherwise writes only resolved size/font family:

```ts
function renderDefaultEndParagraphProperties(
  prefix: string,
  language: string,
  fontFamily?: string,
  fontSize?: number,
): string;
```

Expected font-bearing shape:

```xml
<a:endParaRPr lang="en-US" sz="1825" dirty="0">
  <a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/>
</a:endParaRPr>
```

Continue to prefer `options.endParagraphProperties` during existing-deck replacement.

- [ ] **Step 5: Run tests, typecheck, and review the renderer**

Run:

```bash
pnpm exec vitest run packages/model/src/rich-text.internal.test.ts
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts -t "hyperlink|rich text"
pnpm exec tsc -p packages/model/tsconfig.json --noEmit --pretty false
git diff --check
```

Review nullish rather than truthy precedence, explicit false output, legacy byte parity, linked color suppression only when requested, schema order, escaping, and no behavior change for existing callers.

- [ ] **Step 6: Commit, push, and verify synchronization**

```bash
git add packages/model/src/rich-text.internal.ts \
  packages/model/src/rich-text.internal.test.ts
git commit -m "feat: render rich-text style defaults"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 2: Strict Table/Cell Default Normalization and Creation

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 renderer defaults, `normalizeRichTextColor()`, `normalizeParagraphSpacing()`, and existing descriptor-safe table/cell readers.
- Produces: the five public fields on both option interfaces and resolved direct run/paragraph output for every physical cell.

- [ ] **Step 1: Add failing normalization, precedence, and byte-parity tests**

Use a table default with all five fields and cells covering plain inheritance, runtime undefined, partial cell override, explicit cell false, rich run override, paragraph partial spacing, paragraph `spacing: false`, linked inherited color suppression, and an empty paragraph:

```ts
const definition = normalizeTableDefinition([[
  'Plain',
  { text: 'False', options: { bold: false, spacing: { before: 3 } } },
  { text: [{
    spacing: { after: 9, line: false },
    runs: [
      { text: 'Inherited' },
      { text: 'Local', style: { fontSize: 10, bold: false } },
    ],
  }, { spacing: false, runs: [] }], options: {
    fontFamily: 'Courier New',
    color: { kind: 'srgb', value: '00AA00' },
  } },
]], {
  fontFamily: 'Aptos',
  fontSize: 18.25,
  bold: true,
  color: { kind: 'scheme', value: 'accent1' },
  spacing: {
    before: 6,
    after: 8,
    line: { kind: 'multiple', factor: 1.5 },
  },
});
```

Assert the detached normalized cell defaults and direct OOXML: table values on plain cell; cell `bold: false`; cell partial `before: 3` plus inherited after/line; rich local size/bold; paragraph after/line override with inherited before; `spacing: false` emits no spacing; and no input object changes. Compare omitted, `{ field: undefined }`, and pre-feature cells byte-for-byte.

- [ ] **Step 2: Add strict invalid-input tests and verify the red state**

Cover both table and cell context for:

```ts
const invalidDefaults = [
  { fontFamily: '' },
  { fontFamily: 'bad\u0000font' },
  { fontSize: NaN },
  { fontSize: 0.99 },
  { fontSize: 4000.01 },
  { bold: 1 },
  { color: { kind: 'srgb', value: 'XYZ' } },
  { spacing: {} },
  { spacing: { before: -1 } },
];
```

Also test accessor, symbol, inherited, class-instance, PptxGenJS alias, and caller-mutation cases. Run:

```bash
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  -t "text style defaults"
```

Expected: failures because the option readers reject the new fields and the public types do not expose them.

- [ ] **Step 3: Add public fields and one narrow normalized defaults shape**

Add the five fields to both interfaces in `slide.ts`. In `table-create.internal.ts`, introduce:

```ts
interface NormalizedTableTextDefaults {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly color?: RichTextColor;
  readonly spacing?: NormalizedParagraphSpacingUpdate;
}
```

Extend `OPTION_KEYS`, the cell option supported-key list, `NormalizedTableCell`, and the `normalizeTableCellOptions()` pick. Implement one `normalizeTableTextDefaults(options, context)` helper using:

- direct non-empty/XML-safe string validation for font family;
- finite `1..4000` validation and `Math.round(value * 100) / 100` for font size;
- direct boolean validation for bold;
- `normalizeRichTextColor(value, `${context} color`)`;
- `normalizeParagraphSpacing(value, `${context} spacing`)`.

Return only present normalized fields so omission remains byte-identical.

- [ ] **Step 4: Resolve table → cell defaults with field-safe overlays**

Normalize table defaults after descriptor-safe `readOptions()`. Map every row once and resolve scalar fields with cell values first. Merge spacing subfields without mutating either source:

```ts
function resolveTableTextDefaults(
  table: NormalizedTableTextDefaults,
  cell: NormalizedTableCell,
): NormalizedTableCell {
  const fontFamily = cell.fontFamily ?? table.fontFamily;
  const fontSize = cell.fontSize ?? table.fontSize;
  const bold = cell.bold ?? table.bold;
  const color = cell.color ?? table.color;
  const spacing = table.spacing === undefined && cell.spacing === undefined
    ? undefined
    : { ...table.spacing, ...cell.spacing };
  return {
    ...cell,
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(bold === undefined ? {} : { bold }),
    ...(color === undefined ? {} : { color }),
    ...(spacing === undefined ? {} : { spacing }),
  };
}
```

Use explicit variables in production if that is clearer, but retain nullish handling so `false` survives. Apply this resolution before the existing align/border/fill/margin/direction/valign layers.

- [ ] **Step 5: Pass resolved defaults into the shared renderer**

Extend the `renderRichTextParagraphs()` options in `renderTableCell()`:

```ts
...(cell.fontFamily === undefined ? {} : { defaultFontFamily: cell.fontFamily }),
...(cell.fontSize === undefined ? {} : { defaultFontSize: cell.fontSize }),
...(cell.bold === undefined ? {} : { defaultBold: cell.bold }),
...(cell.color === undefined ? {} : {
  defaultColor: cell.color,
  suppressDefaultColorForHyperlinks: true,
}),
...(cell.spacing === undefined ? {} : { defaultSpacing: cell.spacing }),
```

Do not modify hyperlink relationship preparation or persist default metadata.

- [ ] **Step 6: Run focused/regression gates and review creation**

Run:

```bash
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  packages/model/src/rich-text.internal.test.ts
pnpm exec vitest run packages/model/src/model.test.ts -t "table creation"
pnpm typecheck
git diff --check
```

Review exact supported-key lists, validation contexts, detachment, scalar/partial spacing precedence, explicit false, no relationship changes, and omitted/undefined byte parity.

- [ ] **Step 7: Commit, push, and verify synchronization**

```bash
git add packages/model/src/slide.ts \
  packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts
git commit -m "feat: add table-cell text defaults"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 3: Public Lifecycle and PptxGenJS Conformance

**Files:**
- Modify: `packages/model/src/model.test.ts`
- Modify: `packages/sdk/src/index.test.ts`
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: Task 2 public fields and direct rich-text state.
- Produces: regression evidence for snapshots, edits, formats, transaction isolation, declarations, and PptxGenJS 4.0.1 final-state parity.

- [ ] **Step 1: Add model lifecycle tests**

Create one table with table defaults, plain/rich/empty/linked cells, and cell overrides. Assert immediately and after write/reopen:

```ts
expect(table.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style).toMatchObject({
  fontFamily: 'Aptos',
  fontSize: 18.25,
  bold: true,
  color: { kind: 'scheme', value: 'accent1' },
});
expect(table.rows[0]!.cells[1]!.richText[0]!.spacing).toEqual({
  before: 3,
  after: 8,
  line: { kind: 'multiple', factor: 1.5 },
});
```

Then verify `setCellText()` preserves the first run style, while `setCellRichText(0, 0, [{ runs: [{ text: 'Replacement' }] }])` produces no old font/size/bold/color/spacing defaults. Cover duplicate isolation, outer transaction rollback, unchanged sibling parts, and all six output formats.

- [ ] **Step 2: Add SDK/root compile-time and runtime tests**

Declare both option types with all five fields, create through the SDK/root package, and verify immediate/reopened snapshots. Include negative cases:

```ts
// @ts-expect-error native table options use fontFamily, not fontFace
const tableAlias: AddTableOptions = { fontFace: 'Aptos' };
// @ts-expect-error bold must be boolean
const cellBold: AddTableCellOptions = { bold: 1 };
// @ts-expect-error native spacing uses the structured ParagraphSpacing value
const cellSpacing: AddTableCellOptions = { paraSpaceAfter: 6 };
```

- [ ] **Step 3: Add real PptxGenJS 4.0.1 conformance fixtures**

Generate public output with table `fontFace/fontSize/bold/color`, cell exact/multiple spacing, cell overrides, rich run overrides, empty paragraphs, and local/default hyperlinks. Import it and compare final direct `richText` state against a native table using `fontFamily/fontSize/bold/color/spacing`.

Lock intentional differences: PptxGenJS table-level spacing does not propagate; its cell/run `bold: false` can be overwritten by outer true; native keeps false and never mutates caller objects. Verify both imported and native tables remain editable and stable after write/reopen.

- [ ] **Step 4: Run lifecycle, adapter, type, build, and review gates**

Run:

```bash
pnpm exec vitest run packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "table-cell text style defaults|PptxGenJS table text defaults"
pnpm typecheck
pnpm build
git diff --check
```

Review direct-state snapshots rather than retained metadata, set-text preservation, rich replacement no re-inheritance, six-format stability, invalid-input rollback, exact PptxGenJS version, and negative type assertions.

- [ ] **Step 5: Commit, push, and verify synchronization**

```bash
git add packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
git commit -m "test: verify table-cell text defaults"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 4: Documentation and Support Matrix

**Files:**
- Modify: `README.md`
- Modify: `packages/pptx/README.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: verified public API, PptxGenJS differences, and lifecycle results from Tasks 1–3.
- Produces: bilingual API guidance, accurate support status, remaining-work list, and release notes.

- [ ] **Step 1: Document the exact API and precedence**

Add a table-cell text-defaults section with a working example using table defaults, cell overrides, rich run overrides, spacing partial merge, explicit false, and hyperlink color behavior. State the exact table → cell → paragraph/run precedence, strict native names, direct-state-only persistence, `setCellText()` preservation, and `setCellRichText()` no re-inheritance.

- [ ] **Step 2: Update compatibility and remaining-work status**

Mark table/cell outer font family/font size/bold/color/spacing creation as supported. Record PptxGenJS `fontFace`/spacing aliases at import boundary, hyperlink color behavior, truthy false defect, caller mutation difference, and native table-level spacing extension. Keep merge/span, row/column CRUD, auto-page/repeated headers, `tableToSlides`, and final audit as remaining.

- [ ] **Step 3: Review documentation and run focused gates**

Run:

```bash
rg -n "fontFamily|fontSize|bold|spacing|fontFace|merge|tableToSlides" \
  README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md CHANGELOG.md
pnpm typecheck
pnpm exec vitest run packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t "text style defaults|table text defaults"
git diff --check
```

Review Chinese/English consistency, exact option names, no stale unsupported claim, and remaining-scope accuracy.

- [ ] **Step 4: Commit, push, and verify synchronization**

```bash
git add README.md packages/pptx/README.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md CHANGELOG.md
git commit -m "docs: document table-cell text defaults"
git push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

### Task 5: Full and Published-Package Proof

**Files:**
- Modify only if a verified defect is found; otherwise this task creates evidence outside the repository.

**Interfaces:**
- Consumes: committed implementation and documentation.
- Produces: final full-suite, performance, packed Node/browser/type/CLI/Inspector, PowerPoint 2010 validation, and real-Chrome evidence.

- [ ] **Step 1: Run clean project gates**

Run:

```bash
pnpm typecheck
pnpm build
pnpm exec vitest run --maxWorkers=2
pnpm test:performance
git diff --check
```

Require zero failures other than the repository's single documented skipped test, and record exact file/test counts and performance duration.

- [ ] **Step 2: Pack and verify installed exports**

Pack the root package into a fresh temporary directory, install the actual tarball, and run Node ESM, NodeNext TypeScript declaration, browser conditional-export, CLI, and Inspector probes. Create a table using all five fields at both levels, inspect direct XML, edit with both cell text APIs, write, reopen, and verify no dangling relationship or unsupported validation diagnostic.

- [ ] **Step 3: Verify real Chrome and PowerPoint compatibility**

Serve the packed browser bundle locally, run the create/snapshot/edit/write/reopen probe in real Google Chrome, and require zero console/page/network errors. Run PowerPoint 2010 validation over the final package and require zero errors; only intentional external-hyperlink warnings are accepted.

- [ ] **Step 4: Fix any discovered defect through a new reviewed commit**

If proof finds a defect, add the narrowest failing regression, implement the smallest fix, rerun affected/full gates, then review, commit, push, and verify `0 0`. Do not rewrite prior commits or include evidence/temp artifacts in git.

- [ ] **Step 5: Record final evidence in documentation if counts or hashes are release-visible**

If the established README/changelog pattern requires final counts, tarball SHA-256, browser version, or evidence path, update only those fields, review, run `git diff --check`, commit as `docs: record table-cell text defaults proof`, push, and verify `0 0`.

- [ ] **Step 6: Confirm clean synchronized completion**

Run:

```bash
git status --short --branch
git rev-list --left-right --count origin/main...HEAD
```

Expected: only the pre-existing ignored/untracked `.pnpm-store/` remains and divergence is `0 0`.
