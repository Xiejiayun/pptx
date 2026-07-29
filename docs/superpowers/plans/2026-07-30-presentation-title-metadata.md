# Presentation Title Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation-title metadata creation, direct-state reading, lossless editing, clearing, and PptxGenJS 4.0.1 conformance.

**Architecture:** A focused model-internal core-properties helper owns root relationship resolution, namespace-aware `dc:title` discovery, minimal part creation, and source-span patches. `PresentationModel` exposes the live property, while `PptxDocument.create({ title })` applies explicit creation input through the same setter so create and edit semantics cannot diverge.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, OPC package transactions, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarballs, `pptx-inspect`, LibreOffice headless, Poppler.

## Global Constraints

- Public API is exactly `CreatePresentationOptions.title?: string` and `PresentationModel.title: string | undefined`.
- Omitted creation input writes no title; explicit empty string is direct empty state; assigning `undefined` clears direct state.
- Only the package root core-properties relationship and direct Dublin Core title are owned.
- Part URI and lexical namespace prefixes are not fixed; namespace URI and exact relationship/content type define ownership.
- Existing creator, subject, revision, timestamps, unknown children, comments, whitespace, unrelated parts, and relationships remain byte-preserved.
- Same-value assignment and absent clear are exact bytes/journal no-ops.
- Missing core-properties state may be created; dangling, external, wrong-content-type, wrong-root, duplicate-relationship, and duplicate-title states must not be destructively guessed.
- Native inputs are strict strings without trimming or coercion; invalid input fails before mutation.
- PptxGenJS is used only through public `write()` output; private fields are forbidden.
- `.pnpm-store/` is user-owned and must never be modified, removed, staged, or committed.
- Every independently reviewed task is committed, pushed over SSH port 443, fetched, and verified with `origin/main...HEAD = 0 0` before the next task.

---

### Task 1: Add the core-properties title helper

**Files:**
- Create: `packages/model/src/presentation-title.internal.ts`
- Create: `packages/model/src/presentation-title.internal.test.ts`

**Interfaces:**
- Consumes: `OpcPackage`, `Relationship`, `PackageError`, `LosslessXmlDocument`, `XmlElement`, `escapeXmlText`, and `ModelParseError`.
- Produces: `readPresentationTitle(pkg: OpcPackage): string | undefined`, `replacePresentationTitle(pkg: OpcPackage, value: string | undefined): void`, and internal strict normalization.

- [ ] **Step 1: Write focused package fixtures and failing snapshot tests**

Create a helper that starts from `OpcPackage.create()` and installs core state only when requested:

```ts
const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CORE_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';

function corePackage(xml?: string, uri = '/docProps/core.xml'): OpcPackage {
  const pkg = OpcPackage.create();
  if (xml !== undefined) {
    pkg.transaction(() => {
      pkg.setPart(uri, xml, CORE_TYPE);
      pkg.addRelationship('/', { id: 'rId1', type: CORE_REL, target: uri.slice(1) });
    });
  }
  return pkg;
}
```

Add snapshot cases for:

```ts
expect(readPresentationTitle(corePackage())).toBeUndefined();
expect(readPresentationTitle(corePackage(
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Quarterly &amp; Review</dc:title></cp:coreProperties>',
))).toBe('Quarterly & Review');
expect(readPresentationTitle(corePackage(
  '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:d="http://purl.org/dc/elements/1.1/"><d:title/></c:coreProperties>',
))).toBe('');
```

Also assert `undefined` and zero journal change for descendant-only, wrong namespace, two direct Dublin Core titles, title with an element child, title containing CDATA, wrong root, wrong content type, external relationship, dangling relationship, and two root core relationships.

- [ ] **Step 2: Run the new test and prove the helper is missing**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-title.internal.test.ts --reporter=dot
```

Expected: FAIL because `presentation-title.internal.js` does not exist.

- [ ] **Step 3: Implement strict relationship and namespace discovery**

Add exact constants and helpers:

```ts
const CORE_PROPERTIES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CORE_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml';
const CORE_PROPERTIES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DUBLIN_CORE_NAMESPACE = 'http://purl.org/dc/elements/1.1/';

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function namespaceUri(element: XmlElement): string | undefined {
  const prefix = lexicalPrefix(element.name);
  const declaration = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let scope: XmlElement | undefined = element; scope; scope = scope.parent) {
    const attribute = scope.attributes.find(({ name }) => name === declaration);
    if (attribute) return attribute.value;
  }
  return undefined;
}
```

`readPresentationTitle()` returns `undefined` rather than throwing for unreadable snapshot state. Require one internal root relationship, existing target part with exact content type, one root in the core-properties namespace, one direct `title` child in the Dublin Core namespace, no element children, and no `<![CDATA[` inside the original element. Return `xml.text(title)` so entities decode and empty state remains `''`.

- [ ] **Step 4: Write failing mutation and lifecycle tests**

Cover all of these exact outcomes:

```ts
replacePresentationTitle(pkg, 'Quarterly & <Review>');
expect(readPresentationTitle(pkg)).toBe('Quarterly & <Review>');
expect(new TextDecoder().decode(pkg.requirePart('/docProps/core.xml').bytes))
  .toContain('<dc:title>Quarterly &amp; &lt;Review&gt;</dc:title>');
```

- Missing relationship plus string creates `/docProps/core.xml`, exact content type, one root relationship, and only the title field.
- Occupied orphan `/docProps/core.xml` creates `/docProps/core1.xml` without modifying the orphan.
- Existing alternate `d:` prefix is reused.
- No Dublin Core binding inserts a locally declared canonical `dc:title`.
- Self-closing core root expands safely.
- Self-closing title expands for non-empty replacement.
- Explicit empty inserts `<dc:title></dc:title>` and snapshots `''`.
- Same decoded value preserves exact bytes and journal, including numeric entity spelling.
- `undefined` removes one valid direct title but preserves the part, relationship, creator, revision, whitespace, and opaque child.
- Missing title plus `undefined` is exact no-op.
- Invalid runtime values and XML controls throw `TypeError` before mutation.
- Unique malformed title, dangling/external/wrong-content-type/wrong-root state, duplicate relationships, and duplicate titles throw `PackageError` or `ModelParseError` with zero mutation.
- An outer `pkg.transaction()` rollback restores created part, root relationship, content-types bytes, source part bytes, and journal.

- [ ] **Step 5: Implement normalization, insert, replace, clear, and creation**

Use one strict normalizer:

```ts
function normalizePresentationTitle(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation title must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation title contains invalid XML characters');
  }
  return value;
}
```

For missing state, prefer `/docProps/core.xml`; if occupied, call `pkg.allocatePartUri('/docProps', 'core', '.xml')`. Create the part before the relationship inside the caller's transaction. Render the root with canonical core and Dublin Core namespace declarations and `escapeXmlText(value)`. Use the target URI without its leading slash for the root relationship target.

For an existing valid part, parse once, require mutation-safe state, compare the decoded current string before patching, and then:

- call `xml.removeElement(title)` for clear;
- call `xml.replaceText(title, value)` for non-self-closing replacement;
- expand a self-closing title by replacing its trailing `/>` with `>${escapeXmlText(value)}</${title.name}>`;
- call `xml.appendChildXml(root, encodedTitle)` for insertion.

- [ ] **Step 6: Run focused tests and full model typecheck**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-title.internal.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

Expected: all focused tests pass and the model project typechecks.

- [ ] **Step 7: Review, commit, push, fetch, and prove synchronization**

Review namespace URI ownership, malformed getter/setter split, allocation, relationship target, content type, escaping, semantic no-op, and rollback. Then:

```sh
git add -- packages/model/src/presentation-title.internal.ts \
  packages/model/src/presentation-title.internal.test.ts
git diff --cached --check
git commit -m 'feat: add presentation title core property codec'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch origin
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 2: Expose the live model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 1 `readPresentationTitle()` and `replacePresentationTitle()`.
- Produces: public `PresentationModel.title: string | undefined` with transactional mutation.

- [ ] **Step 1: Write failing public model lifecycle tests**

Build title state on top of `modelFixture()` without changing the shared fixture. Assert getter no mutation, same-value no-op, custom/empty values, invalid-value isolation, other-part hashes, rollback, clear preservation, missing-part creation/reopen, and absent clear no-op. Stable slide identity must survive every title operation.

- [ ] **Step 2: Run the targeted model test and prove the property is absent**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation title metadata' --reporter=dot
```

Expected: FAIL because `PresentationModel.title` is not declared.

- [ ] **Step 3: Add the public getter and setter**

Import Task 1 and add near `rtlMode`:

```ts
get title(): string | undefined {
  return readPresentationTitle(this.opcPackage);
}

set title(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationTitle(this.opcPackage, value);
  });
}
```

Do not export the internal helper through `packages/model/src/index.ts`.

- [ ] **Step 4: Run focused model tests and the complete model suite**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation title metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m 'feat: expose presentation title metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch origin
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 3: Add native title creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 2 `PresentationModel.title` and existing `PptxDocument.create/open/write`.
- Produces: `CreatePresentationOptions.title?: string` and full create/edit/reopen lifecycle.

- [ ] **Step 1: Write failing SDK creation tests**

Require omitted `undefined`, escaped custom title, direct empty, all six formats, immediate edit, same-value no-op, rollback, clear, write/reopen, and strict rejection of null, boolean, number, object, array, symbol, and XML controls.

- [ ] **Step 2: Run the SDK test and prove the create option is missing**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation title metadata' --reporter=dot
```

- [ ] **Step 3: Extend create options and apply the shared setter**

Add `readonly title?: string` to `CreatePresentationOptions`. Keep package scaffolding unchanged and update:

```ts
static create(options: CreatePresentationOptions = {}): PptxDocument {
  const document = new PptxDocument(createPresentationPackage(options));
  if (options.title !== undefined) document.title = options.title;
  return document;
}
```

Do not add a PptxGenJS-branded default to `CORE_PROPERTIES_XML`.

- [ ] **Step 4: Run SDK, validator, and full typecheck gates**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation title metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src \
  packages/validator/src --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m 'feat: create presentation title metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch origin
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 4: Prove PptxGenJS title conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: PptxGenJS 4.0.1 public `title` property and `write()`, adapter import, and Task 3 title API.
- Produces: real-output evidence for default, custom, and empty PptxGenJS titles plus native explicit-title parity.

- [ ] **Step 1: Add real public-output cases**

Create three PptxGenJS instances with one slide each:

```ts
const baseline = new PptxGenJS();
baseline.addSlide();
const custom = new PptxGenJS();
custom.title = 'Quarterly & <Review>';
custom.addSlide();
const empty = new PptxGenJS();
empty.title = '';
empty.addSlide();
```

Import only through `importPptxGenJS()` and require:

```ts
expect((await importPptxGenJS(baseline)).title).toBe('PptxGenJS Presentation');
expect((await importPptxGenJS(custom)).title).toBe('Quarterly & <Review>');
expect((await importPptxGenJS(empty)).title).toBe('');
```

Assert exact core XML escaping and empty state, then write/reopen all three and repeat snapshots. Create a native explicit title and require the same custom semantic state. Native omitted remains `undefined` as a documented default difference.

- [ ] **Step 2: Run the adapter test**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'presentation title metadata' --reporter=dot
```

- [ ] **Step 3: Review, commit, push, fetch, and prove synchronization**

Review public output use, exact version assertion, default/custom/empty distinction, native omitted difference, escaping, and reopen. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m 'test: compare presentation title metadata with pptxgenjs'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch origin
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 5: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: Task 3 aggregate runtime and declaration exports.
- Produces: actual-tarball `presentationTitle: true`, browser lifecycle proof, and compile-time `string | undefined` coverage.

- [ ] **Step 1: Add isolated Node title smoke**

Near presentation RTL, add:

```js
const metadata = PptxDocument.create({ title: 'Packed & <Title>' });
const createdPresentationTitle = metadata.title;
metadata.title = 'Edited title';
const editedPresentationTitle = metadata.title;
const reopenedMetadata = await PptxDocument.open(await metadata.write());
const reopenedPresentationTitle = reopenedMetadata.title;
metadata.title = '';
const emptyPresentationTitle = metadata.title;
metadata.title = undefined;
const clearedPresentationTitle = metadata.title;
```

Emit:

```js
presentationTitle:
  createdPresentationTitle === 'Packed & <Title>' &&
  editedPresentationTitle === 'Edited title' &&
  reopenedPresentationTitle === 'Edited title' &&
  emptyPresentationTitle === '' &&
  clearedPresentationTitle === undefined,
```

- [ ] **Step 2: Mirror the runtime proof through the browser condition**

Create with `{ title: 'Browser title' }`, require the getter, edit, write/reopen, set empty, and clear. The browser bundle must retain no static `node:` import.

- [ ] **Step 3: Compile the declaration surface**

In generated `smoke.ts`, add:

```ts
const titledDocument: PptxDocument = PptxDocument.create({ title: 'Typed title' });
const titleSnapshot: string | undefined = titledDocument.title;
titledDocument.title = 'Edited typed title';
titledDocument.title = '';
titledDocument.title = undefined;
```

Include these values in the final `void [...]` expression.

- [ ] **Step 4: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
presentation_title_package_dir=$(mktemp -d /tmp/pptx-presentation-title-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_title_package_dir"
presentation_title_tarball=$(find "$presentation_title_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_title_tarball"
cd ../..
```

Require every smoke field true, including `presentationTitle`, types true, browser resolution, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m 'test: smoke packed presentation title metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch origin
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 6: Document the contract and remaining metadata gaps

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: completed runtime, conformance, and packed evidence.
- Produces: accurate title documentation without claiming other metadata fields.

- [ ] **Step 1: Update public API documentation**

Add examples:

```ts
const document = PptxDocument.create({ title: 'Quarterly Review' });
document.title = 'Updated Review';
document.title = '';
document.title = undefined;
```

Document direct snapshot, omitted versus empty versus clear, strict validation, relationship-based part lookup, missing-part creation, same-value no-op, preservation of other metadata, and malformed ambiguity rejection.

- [ ] **Step 2: Update compatibility and remaining gaps**

Add matrix row:

```md
| presentation `pptx.title` | `CreatePresentationOptions.title` / `document.title` | 已支持 strict string 创建、direct core-property 读取/编辑/empty/clear；native omitted 不写 PptxGenJS 品牌默认值 |
```

State PptxGenJS default `PptxGenJS Presentation`, exact custom/empty import, and native editing/missing-part extensions. Keep author, company, revision, subject, theme, sections, masters, placeholders, and other metadata pending.

Add changelog bullet:

```md
- Added strict presentation-title metadata creation plus namespace-aware direct core-property reading, lossless editing, empty/clear states, and missing-part lifecycle support.
```

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'presentation title.*(unsupported|pending)|pptx\.title.*(未支持|尚未支持)|metadata.*全部支持' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md || true
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m 'docs: document presentation title metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch origin
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 7: Run full gates and real-deck QA

**Files:**
- Review every Task 1–6 path; never stage `.pnpm-store/`.

**Interfaces:**
- Consumes: implementation, tests, docs, actual tarball, repository CLI, LibreOffice, Poppler, and overflow checker.
- Produces: a fully verified pushed title feature; any defect gets a focused fix/review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-presentation-title-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-presentation-title-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Require every functional suite and test pass; the default pending performance case must pass separately below 5 seconds.

- [ ] **Step 2: Rebuild and smoke the final tarball plus doctor**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
presentation_title_qa_package_dir=$(mktemp -d /tmp/pptx-presentation-title-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_title_qa_package_dir"
presentation_title_qa_tarball=$(find "$presentation_title_qa_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_title_qa_tarball"
cd ../..
pptx-inspect --json doctor
```

- [ ] **Step 3: Generate real native and PptxGenJS QA decks**

Generate `/tmp/pptx-presentation-title-qa/` containing native omitted/source/edited/empty/cleared/reopened and PptxGenJS default/custom/empty decks. Each has one visible wide slide: native creation uses `slideSize: 'wide'` and every PptxGenJS fixture sets `layout = 'LAYOUT_WIDE'`. Native custom text is `Quarterly & <Review>`; edited text is `Edited title`.

- [ ] **Step 4: Validate packages and exact parts**

For every deck run:

```sh
pptx-inspect --json package validate "$deck" --profile powerpoint-2010
pptx-inspect --json part read "$deck" /docProps/core.xml
```

Require zero errors/warnings, native omitted/cleared absence, exact custom decoding and escaping, explicit empty state, and PptxGenJS default brand value.

- [ ] **Step 5: Prove package mutation isolation**

```sh
pptx-inspect --json package diff \
  /tmp/pptx-presentation-title-qa/native-source.pptx \
  /tmp/pptx-presentation-title-qa/native-edited.pptx
pptx-inspect --json package diff \
  /tmp/pptx-presentation-title-qa/native-edited.pptx \
  /tmp/pptx-presentation-title-qa/native-reopened.pptx
```

Require source→edited changes only `/docProps/core.xml`; edited→reopened has zero decompressed part changes.

- [ ] **Step 6: Render and check visual isolation**

Export native source, edited, and PptxGenJS custom through isolated LibreOffice profiles. Require one page, unchanged `960 × 540 pt` wide geometry, successful PNG rasterization, and no overflow using:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  "$deck"
```

Inspect every rendered page individually; metadata must not alter visible content.

- [ ] **Step 7: Final static review and synchronization proof**

```sh
git diff --check
git status --short --branch
git fetch origin
git rev-list --left-right --count origin/main...HEAD
```

Require only `?? .pnpm-store/`, `main...origin/main`, and `0 0`. If QA finds a defect, add the smallest failing regression test, fix it, review, commit, push over SSH 443, fetch, repeat affected gates, and do not create an empty completion commit.
