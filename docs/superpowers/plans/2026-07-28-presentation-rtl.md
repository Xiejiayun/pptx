# Presentation RTL Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add native presentation-root RTL create/read/edit/clear support with strict validation and PptxGenJS 4.0.1 valid-output conformance.

**Architecture:** Extend `PresentationModel` with a direct tri-state `rtlMode` property that losslessly patches only `p:presentation@rtl`. Thread the same strict boolean through `CreatePresentationOptions` before package creation, then verify the public SDK, adapter, packed bundle, CLI, and LibreOffice paths without coupling global RTL to paragraph or text-body direction.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span model, PptxGenJS 4.0.1 conformance fixture, tsup, npm tarball smoke, repository CLI, LibreOffice headless.

## Global Constraints

- Public creation and document property names are both `rtlMode` for PptxGenJS parity.
- `document.rtlMode` is `boolean | undefined`: true writes `rtl="1"`, false writes `rtl="0"`, and undefined removes the root attribute.
- Read only direct root `p:presentation@rtl`: `1/true/on` are true, `0/false/off` are false, and missing/empty/unknown are undefined.
- Never read or modify paragraph `pPr@rtl`, text-body `bodyPr@rtlCol`, master/layout/default text styles, alignment, or run order.
- Validate creation input before `OpcPackage.create()` and setter input before any part mutation.
- PptxGenJS false/omitted both omit root RTL; native explicit false intentionally writes `0` with equivalent effective behavior.
- Implement inline without subagent delegation, as required for this repository session.
- Review, commit, and push only after typecheck, full tests, performance, package smoke, CLI validation, and LibreOffice comparison pass.

---

### Task 1: Direct model getter and lossless setter

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Test: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: `PresentationModel.parsePresentation()`, `presentationRoot()`, `LosslessXmlDocument.attribute()`, `replaceAttribute()`, `replace()`, `setXmlPart()`, and `OpcPackage.transaction()`.
- Produces: `PresentationModel.rtlMode: boolean | undefined`, strict root-token reads, and exact root-attribute replace/insert/remove behavior.

- [ ] **Step 1: Write a failing strict direct-read test**

Create or patch presentation parts whose root `rtl` value is `1`, `true`, `on`, `0`, `false`, `off`, missing, empty, and `yes`. For every case instantiate `PresentationModel`, read `model.rtlMode`, and assert:

```ts
expect(values.map(({ model }) => model.rtlMode)).toEqual([
  true,
  true,
  true,
  false,
  false,
  false,
  undefined,
  undefined,
  undefined,
]);
expect(pkg.mutations).toEqual(journal);
```

Put `rtl="1"` on a descendant `a:lvl1pPr` while omitting it from the root and assert the getter remains undefined.

- [ ] **Step 2: Run the model suite and confirm the getter test fails**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: the new test fails because `PresentationModel.rtlMode` does not exist.

- [ ] **Step 3: Add the strict getter**

In `PresentationModel`, read only the root attribute and preserve false:

```ts
get rtlMode(): boolean | undefined {
  const { xml } = this.parsePresentation();
  const root = presentationRoot(xml, this.presentationPartUri);
  const value = xml.attribute(root, 'rtl')?.value;
  if (value === undefined) return undefined;
  if (['1', 'true', 'on'].includes(value)) return true;
  if (['0', 'false', 'off'].includes(value)) return false;
  return undefined;
}
```

Re-run the model suite and expect the strict getter test to pass with no mutation journal changes.

- [ ] **Step 4: Write failing setter, clearing, preservation, and rollback tests**

Start with root XML containing `rtl="yes"`, `saveSubsetFonts`, a namespace declaration, and unknown children. Set true, false, and undefined in separate transactions. Assert exact root tokens `1`, `0`, and absent; assert all unrelated root attributes and children remain byte-identical. Then capture presentation bytes, journal, and slide identity and assert rollback restores all three:

```ts
const before = pkg.requirePart(model.presentationPartUri).bytes;
const journal = [...pkg.mutations];
const slide = model.slides[0];
expect(() => pkg.transaction(() => {
  model.rtlMode = true;
  throw new Error('restore presentation RTL');
})).toThrow('restore presentation RTL');
expect(pkg.requirePart(model.presentationPartUri).bytes).toEqual(before);
expect(pkg.mutations).toEqual(journal);
expect(model.slides[0]).toBe(slide);
```

Loop through null, number, string, object, array, and symbol values cast as never; each setter call must throw TypeError before bytes or journal change.

- [ ] **Step 5: Implement minimal setter and root attribute patching**

Normalize before parsing or mutating:

```ts
set rtlMode(value: boolean | undefined) {
  this.opcPackage.transaction(() => {
    if (value !== undefined && typeof value !== 'boolean') {
      throw new TypeError('Presentation RTL mode must be a boolean or undefined');
    }
    const { xml } = this.parsePresentation();
    const root = presentationRoot(xml, this.presentationPartUri);
    updatePresentationAttribute(xml, root, 'rtl', value === undefined ? undefined : value ? '1' : '0');
    this.setXmlPart(this.presentationPartUri, xml.serialize());
  });
}
```

Add a local helper that mirrors existing lossless attribute patching without introducing a new public codec:

```ts
function updatePresentationAttribute(
  xml: LosslessXmlDocument,
  root: XmlElement,
  name: string,
  value: string | undefined,
): void {
  const attribute = xml.attribute(root, name);
  if (value !== undefined) {
    if (attribute) xml.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? xml.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      xml.replace(insertionPoint, insertionPoint, ` ${name}="${value}"`);
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(xml.source[start - 1] ?? '')) start -= 1;
    xml.replace(start, attribute.end, '');
  }
}
```

- [ ] **Step 6: Run model tests and typecheck**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: typecheck exits 0 and all model tests pass.

### Task 2: Zero-input creation and native SDK lifecycle

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Test: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: `CreatePresentationOptions`, `createPresentationPackage()`, `presentationXml()`, `PptxDocument.create/open/write`, and Task 1 `PresentationModel.rtlMode`.
- Produces: `CreatePresentationOptions.rtlMode?: boolean`, canonical root creation output, strict create validation, and public write/reopen lifecycle coverage.

- [ ] **Step 1: Write failing omitted/true/false creation tests**

Create three documents and assert direct state plus root XML:

```ts
const readPresentationXml = (document: PptxDocument): string => new TextDecoder().decode(
  document.opcPackage.requirePart(document.presentationPartUri).bytes,
);
const omitted = PptxDocument.create();
const enabled = PptxDocument.create({ rtlMode: true });
const disabled = PptxDocument.create({ rtlMode: false });
expect([omitted.rtlMode, enabled.rtlMode, disabled.rtlMode]).toEqual([undefined, true, false]);
expect(readPresentationXml(enabled)).toMatch(/<p:presentation[^>]* rtl="1"/);
expect(readPresentationXml(disabled)).toMatch(/<p:presentation[^>]* rtl="0"/);
expect(readPresentationXml(omitted)).not.toMatch(/<p:presentation[^>]*\srtl=/);
```

Repeat true creation for all six `PresentationFormat` values and validate every package.

- [ ] **Step 2: Add strict creation option normalization**

Add `readonly rtlMode?: boolean` to `CreatePresentationOptions`. Before `OpcPackage.create()`, normalize with:

```ts
const rtlMode = options.rtlMode;
if (rtlMode !== undefined && typeof rtlMode !== 'boolean') {
  throw new TypeError('Presentation RTL mode must be a boolean');
}
```

Pass the normalized value to `presentationXml(cx, cy, rtlMode)` and include exactly one attribute fragment:

```ts
const rtlAttribute = rtlMode === undefined ? '' : ` rtl="${rtlMode ? '1' : '0'}"`;
```

Insert the fragment on `p:presentation`; do not change `defaultTextStyle` or any other package part.

- [ ] **Step 3: Cover edit/clear/write/reopen/rollback**

Create with true, add an RTL and LTR paragraph as independent evidence, then edit root false and clear it without changing either paragraph snapshot. Duplicate a slide, write/reopen, and assert root state plus both slides/shapes remain intact. Use `document.transaction()` to set root true and throw; assert exact presentation bytes, mutation journal, and slide identity restore. End with zero validator errors.

- [ ] **Step 4: Add invalid creation mutation-isolation coverage**

For null, number, string, object, array, and symbol values, assert `PptxDocument.create({ rtlMode: value as never })` throws TypeError before returning a document. Review `createPresentationPackage()` to confirm this validation remains textually above `OpcPackage.create()`.

- [ ] **Step 5: Run focused SDK tests and typecheck**

```sh
node node_modules/.pnpm/typescript@5.9.3/node_modules/typescript/lib/tsc.js -b --pretty false
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts packages/sdk/src/index.test.ts
```

Expected: both suites pass and declaration generation accepts the new create option/property.

### Task 3: PptxGenJS conformance, documentation, and packed surface

**Files:**
- Test: `packages/pptxgenjs-adapter/src/index.test.ts`
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Tasks 1–2 public property/create option and the existing real PptxGenJS import/write/reopen fixture.
- Produces: root-level conformance evidence, documented global/paragraph separation, and Node/browser/declaration tarball smoke.

- [ ] **Step 1: Add real PptxGenJS root RTL fixtures**

Extend the test-only instance interface with mutable `rtlMode: unknown`. Generate separate PptxGenJS presentations for omitted, true, false, and truthy string values. Import each through `importPptxGenJS()` and assert `[undefined, true, undefined, true]`. Read exact presentation XML and assert only true/truthy roots contain `rtl="1"`; confirm descendant `a:lvl1pPr@rtl="0"` never becomes the document getter. Write/reopen the true case and assert true survives without adapter-private reads.

- [ ] **Step 2: Update compatibility and API documentation**

Add a completed baseline row:

```md
| presentation `pptx.rtlMode` | `CreatePresentationOptions.rtlMode` / `document.rtlMode` | 已支持 |
```

Document direct six-token reads, explicit native false, clear semantics, strict rejection, and separation from paragraph `AddTextOptions.rtlMode` / `RichTextParagraph.rtl`. Add one changelog bullet. Update API/package README examples to create global RTL, set false, and clear it.

- [ ] **Step 3: Extend Node and browser packed smoke**

In Node smoke create `{ rtlMode: true }`, assert true, set false and assert false, set undefined and assert undefined, and verify paragraph RTL remains unchanged. In browser smoke perform the same root-state transitions on a created document and assert no exception.

- [ ] **Step 4: Extend TypeScript declaration smoke**

Compile this exact surface from the packed tarball:

```ts
const globalRtl: PptxDocument = PptxDocument.create({ rtlMode: true });
const globalRtlSnapshot: boolean | undefined = globalRtl.rtlMode;
globalRtl.rtlMode = false;
globalRtl.rtlMode = undefined;
```

Include `globalRtl` and `globalRtlSnapshot` alongside the existing values in the final `void` array so `noUnusedLocals`-style consumers see the complete surface.

- [ ] **Step 5: Run adapter, full, and performance suites**

```sh
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/pptxgenjs-adapter/src/index.test.ts
node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run
RUN_PERF=1 node node_modules/.pnpm/vitest@3.2.7_@types+node@24.13.3/node_modules/vitest/vitest.mjs run packages/testkit/src/performance.test.ts
```

Expected: all functional tests pass, only the default performance case is skipped in the full run, and the isolated performance test passes.

### Task 4: Package, compatibility, review, commit, and push

**Files:**
- Review all Task 1–3 files; never stage `.pnpm-store/`.

**Interfaces:**
- Consumes: completed model/create/docs/smoke implementation.
- Produces: reviewed `feat: support presentation rtl` commit synchronized to `origin/main`.

- [ ] **Step 1: Build and smoke the actual tarball**

From `packages/pptx`, run:

```sh
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js
node ../../node_modules/.pnpm/tsup@8.5.1_postcss@8.5.19_typescript@5.9.3/node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
presentation_rtl_package_dir=$(mktemp -d /tmp/pptx-presentation-rtl-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_rtl_package_dir"
node ../../scripts/smoke-npm-package.mjs "$presentation_rtl_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expect API/browser/types/CLI checks true.

- [ ] **Step 2: CLI and LibreOffice comparison**

Generate `/tmp/pptx-presentation-rtl-native/native.pptx` with `{ rtlMode: true }` and `/tmp/pptx-presentation-rtl-pptxgenjs/pptxgenjs.pptx` with `pptx.rtlMode = true`; give both identical Arabic/Hebrew and English-control paragraphs. Validate:

```sh
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-presentation-rtl-native/native.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json package validate /tmp/pptx-presentation-rtl-pptxgenjs/pptxgenjs.pptx --profile powerpoint-2010
node packages/pptx/dist/cli.js --json part read /tmp/pptx-presentation-rtl-native/native.pptx /ppt/presentation.xml
node packages/pptx/dist/cli.js --json part read /tmp/pptx-presentation-rtl-pptxgenjs/pptxgenjs.pptx /ppt/presentation.xml
```

Both validators must report 0 errors/0 warnings and both exact roots must contain `rtl="1"`. Export each with `/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice --headless --convert-to pdf`, render every PDF page with `pdftoppm -png`, and inspect for repair prompts, clipping, or unintended paragraph-direction changes.

- [ ] **Step 3: Final review and commit**

Run `git diff --check`, inspect the complete diff, and verify no code path edits `pPr@rtl`, `bodyPr@rtlCol`, master/layout/default text style, or alignment. Confirm status lists only intended files plus `.pnpm-store/`. Stage explicit files, then:

```sh
git commit -m "feat: support presentation rtl"
GIT_SSH_COMMAND='ssh -p 443 -o HostKeyAlias=github.com' \
git -c url."ssh://git@ssh.github.com:443/".insteadOf="https://github.com/" \
push origin main
git rev-list --left-right --count origin/main...HEAD
```

Expected final divergence is `0 0`.
