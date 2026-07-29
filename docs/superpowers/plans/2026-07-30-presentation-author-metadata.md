# Presentation Author Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation-author metadata creation, direct-state reading, lossless editing, clearing, and PptxGenJS 4.0.1 conformance without overwriting last-modified-by.

**Architecture:** Extract the proven title relationship/part/namespace/source-span lifecycle into one model-internal generic core text-property helper. Keep title behavior behind its existing wrapper, add a creator-specific author wrapper, expose a live `PresentationModel.author`, and apply explicit creation input through the same setter so create and edit cannot diverge.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, OPC package transactions, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarballs, `pptx-inspect`, LibreOffice headless, Poppler.

## Global Constraints

- Public API is exactly `CreatePresentationOptions.author?: string` and `PresentationModel.author: string | undefined`.
- Native zero-input creation keeps the current direct `@jiayunxie/pptx` creator; explicit empty string is direct empty state; assigning `undefined` clears direct creator.
- Author owns only direct Dublin Core `dc:creator`; it must preserve `cp:lastModifiedBy`, title, subject, revision, timestamps, unknown children, and all unrelated package state.
- PptxGenJS default/custom/empty creator states are imported exactly, but its author-to-lastModifiedBy mirroring is documented rather than copied by native mutation.
- Part URI and lexical namespace prefixes are not fixed; namespace URI and exact root relationship/content type define ownership.
- Same-value assignment and absent clear are exact bytes/journal no-ops.
- Missing core-properties state may be created; dangling, external, wrong-content-type, wrong-root, duplicate-relationship, and duplicate-creator states must not be destructively guessed.
- Native inputs are strict strings without trimming or coercion; invalid input fails before observable mutation.
- Existing title APIs, bytes, errors, rollback, conformance, packed surface, docs, and QA behavior must remain green after the generic refactor.
- PptxGenJS is used only through public `author`, `addSlide()`, and `write()` output; private fields are forbidden.
- `.pnpm-store/` is user-owned and must never be modified, removed, staged, or committed.
- Every independently reviewed task is committed, pushed over SSH port 443, fetched, and verified with `origin/main...HEAD = 0 0` before the next task.

---

### Task 1: Extract the generic core text-property lifecycle

**Files:**
- Create: `packages/model/src/presentation-core-properties.internal.ts`
- Modify: `packages/model/src/presentation-title.internal.ts`

**Interfaces:**
- Produces internal `CoreTextPropertyDescriptor`, `readCoreTextProperty(pkg, descriptor)`, and `replaceCoreTextProperty(pkg, descriptor, value)`.
- Preserves existing `readPresentationTitle()` and `replacePresentationTitle()` signatures and behavior.

- [ ] **Step 1: Record the green title baseline before refactoring**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-title.internal.test.ts \
  packages/model/src/model.test.ts -t 'presentation title metadata|presentation title core property' \
  --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

Require all selected title tests and model typecheck to pass before moving code.

- [ ] **Step 2: Move relationship, part, root, namespace, and patch mechanics into a generic helper**

Create `presentation-core-properties.internal.ts` with the exact internal contract:

```ts
export interface CoreTextPropertyDescriptor {
  readonly label: string;
  readonly localName: string;
  readonly namespace: string;
  readonly preferredPrefix: string;
}

export function readCoreTextProperty(
  pkg: OpcPackage,
  descriptor: CoreTextPropertyDescriptor,
): string | undefined;

export function replaceCoreTextProperty(
  pkg: OpcPackage,
  descriptor: CoreTextPropertyDescriptor,
  value: string | undefined,
): void;
```

Move the exact constants and state discovery from the title helper. `CorePropertiesState` contains relationship, part, XML document, and root; field discovery is descriptor-driven:

```ts
function propertyElements(
  root: XmlElement,
  descriptor: CoreTextPropertyDescriptor,
): XmlElement[] {
  return directChildren(root).filter(
    (child) => child.localName === descriptor.localName
      && namespaceUri(child) === descriptor.namespace,
  );
}
```

`readCoreTextProperty()` requires one safe root relationship/part/root, exactly one direct matching element, simple text, and no CDATA. `replaceCoreTextProperty()` must:

- validate zero/one relationship before field mutation;
- create a minimal core part with only the selected field when no relationship exists and value is a string;
- reject multiple matching direct elements using the stable message `Core properties contain multiple direct ${descriptor.label}s`;
- preserve same decoded value as an exact no-op;
- remove on `undefined`, insert with prefix reuse, replace text, or expand self-closing field/root safely;
- call `pkg.setPart()` only after a semantic change.

Render creation and insertion from the descriptor:

```ts
const qualifiedName = `${descriptor.preferredPrefix}:${descriptor.localName}`;
const field = `<${qualifiedName}>${escapeXmlText(value)}</${qualifiedName}>`;
```

When the root lacks a binding for `descriptor.namespace`, declare `xmlns:${descriptor.preferredPrefix}` locally on the inserted field. Keep the current `/docProps/core.xml` preference, collision allocation, content type, root relationship target, and all unsafe-state errors.

- [ ] **Step 3: Reduce the title file to a strict wrapper**

Keep title-specific normalization and exports in `presentation-title.internal.ts`:

```ts
const TITLE_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'title',
  localName: 'title',
  namespace: 'http://purl.org/dc/elements/1.1/',
  preferredPrefix: 'dc',
};

export function readPresentationTitle(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, TITLE_PROPERTY);
}

export function replacePresentationTitle(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, TITLE_PROPERTY, normalizePresentationTitle(value));
}
```

Keep the existing control-character source regex exactly as textual `\u` escapes:

```ts
/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/
```

- [ ] **Step 4: Prove title behavior and bytes did not regress**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-title.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

Require the existing focused cases, complete model suite, and typecheck to pass without modifying any expected title snapshots.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

Review that the generic helper is internal-only, descriptor-driven, no title behavior moved to the public model, error paths remain zero-mutation, and `.pnpm-store/` is untouched.

```sh
git add -- packages/model/src/presentation-core-properties.internal.ts \
  packages/model/src/presentation-title.internal.ts
git diff --cached --check
git commit -m 'refactor: share presentation core property lifecycle'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 2: Add the author core-property codec

**Files:**
- Create: `packages/model/src/presentation-author.internal.ts`
- Create: `packages/model/src/presentation-author.internal.test.ts`

**Interfaces:**
- Consumes Task 1 generic helper.
- Produces internal `readPresentationAuthor(pkg)` and `replacePresentationAuthor(pkg, value)`.

- [ ] **Step 1: Write focused author snapshots and lifecycle tests**

Use package fixtures with alternate core part URI/prefix and cover:

```ts
expect(readPresentationAuthor(corePackage())).toBeUndefined();
expect(readPresentationAuthor(corePackage(coreXml(
  '<d:creator>Alice &amp; Bob</d:creator><cp:lastModifiedBy>Editor</cp:lastModifiedBy>',
  'c:coreProperties',
  `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:cp="${CORE_NAMESPACE}"`,
)))).toBe('Alice & Bob');
expect(readPresentationAuthor(corePackage(coreXml('<dc:creator/>')))).toBe('');
```

Require getter zero mutation for absent, wrong namespace, descendant-only, duplicate direct creators, element-child creator, CDATA creator, wrong root/content type, dangling/external relationship, and duplicate root relationships.

Mutation tests must prove:

- missing relationship plus string creates the preferred or allocated core part containing only creator;
- alternate `d:` prefix is reused, and missing binding inserts a local canonical `dc:` declaration;
- creator insertion/replacement/empty/clear preserves title, lastModifiedBy, revision, timestamps, unknown children, comments, whitespace, and every other part;
- self-closing root and creator expand safely;
- numeric-entity same-value and absent clear are exact package snapshots;
- lastModifiedBy never changes when creator changes or clears;
- invalid runtime values/XML controls and unsafe ownership leave parts, relationships, content types, journal, and values unchanged;
- outer transaction rollback restores minimal part creation and existing-part edits.

- [ ] **Step 2: Run RED and prove the author helper is absent**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-author.internal.test.ts --reporter=dot
```

Expected: FAIL because `presentation-author.internal.js` does not exist.

- [ ] **Step 3: Implement the strict author wrapper**

```ts
const AUTHOR_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'author',
  localName: 'creator',
  namespace: 'http://purl.org/dc/elements/1.1/',
  preferredPrefix: 'dc',
};

export function readPresentationAuthor(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, AUTHOR_PROPERTY);
}

export function replacePresentationAuthor(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, AUTHOR_PROPERTY, normalizePresentationAuthor(value));
}
```

Normalization is strict and author-specific:

```ts
function normalizePresentationAuthor(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation author must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation author contains invalid XML characters');
  }
  return value;
}
```

- [ ] **Step 4: Run focused tests plus all title/model safety gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-author.internal.test.ts \
  packages/model/src/presentation-title.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation-author.internal.ts \
  packages/model/src/presentation-author.internal.test.ts
git diff --cached --check
git commit -m 'feat: add presentation author core property codec'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 3: Expose the live author model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes Task 2 author helper.
- Produces public `PresentationModel.author: string | undefined` with transactional mutation.

- [ ] **Step 1: Write failing public model lifecycle tests**

Build an alternate-URI/prefix core fixture containing creator, lastModifiedBy, title, revision, and opaque children. Assert getter no mutation, same-value no-op, custom/empty values, invalid-value isolation, other-part hashes, creator-only mutation, lastModifiedBy preservation, rollback, clear preservation, missing-part creation/reopen, absent clear no-op, and stable slide identity.

- [ ] **Step 2: Run the targeted model test and prove the property is absent**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation author metadata' --reporter=dot
```

Expected: FAIL because `PresentationModel.author` is not declared.

- [ ] **Step 3: Add the public getter and setter**

```ts
get author(): string | undefined {
  return readPresentationAuthor(this.opcPackage);
}

set author(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationAuthor(this.opcPackage, value);
  });
}
```

Place it with presentation title/RTL properties and do not export either author helper from `packages/model/src/index.ts`.

- [ ] **Step 4: Run focused model tests and the complete model suite**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation author metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m 'feat: expose presentation author metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 4: Add native author creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `CreatePresentationOptions.author?: string` and complete create/edit/reopen lifecycle.

- [ ] **Step 1: Write failing SDK creation tests**

Require native omitted author `@jiayunxie/pptx`, escaped custom author, direct empty creator, runtime undefined preservation, all six formats, immediate edit, same-value no-op, title coexistence, rollback, clear, write/reopen, creator-only mutation, lastModifiedBy preservation, and strict rejection of null, boolean, number, object, array, symbol, and XML controls.

- [ ] **Step 2: Run the SDK test and prove the create option is missing**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation author metadata' --reporter=dot
```

- [ ] **Step 3: Extend create options and apply the shared setter**

Add the exact optional property:

```ts
readonly author?: string;
```

Apply it before title while leaving package scaffolding unchanged:

```ts
const document = new PptxDocument(createPresentationPackage(options));
if (options.author !== undefined) document.author = options.author;
if (options.title !== undefined) document.title = options.title;
return document;
```

Do not replace the canonical creator or lastModifiedBy in `CORE_PROPERTIES_XML`; omitted creation must preserve existing bytes.

- [ ] **Step 4: Run SDK, validator, and full typecheck gates**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation author metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src \
  packages/validator/src --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m 'feat: create presentation author metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 5: Prove PptxGenJS author conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes PptxGenJS 4.0.1 public `author` and `write()`, adapter import, and native author API.
- Produces real-output evidence for default/custom/empty creator and the intentional lastModifiedBy ownership difference.

- [ ] **Step 1: Add real public-output cases**

Extend the local test interface with `author: string`. Create default, custom, and empty PptxGenJS instances with one slide each:

```ts
const baseline = new PptxGenJS();
baseline.addSlide();
const custom = new PptxGenJS();
custom.author = 'Alice & <Bob>';
custom.addSlide();
const empty = new PptxGenJS();
empty.author = '';
empty.addSlide();
```

Import only through `importPptxGenJS()` and require author snapshots `PptxGenJS`, `Alice & <Bob>`, and `''`; exact creator and lastModifiedBy XML for all three; getter zero mutation; write/reopen stability; PptxGenJS version `4.0.1`.

Create native custom author and require:

```ts
expect(native.author).toBe('Alice & <Bob>');
expect(nativeCore).toContain('<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>');
expect(nativeCore).toContain(
  '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
);
```

Native omitted remains `@jiayunxie/pptx`; document this default and mirroring difference without treating it as a failure.

- [ ] **Step 2: Run focused and complete adapter tests plus typecheck**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'presentation author metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 3: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m 'test: compare presentation author metadata with pptxgenjs'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 6: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Produces actual-tarball `presentationAuthor: true`, browser lifecycle proof, and compile-time `string | undefined` coverage.

- [ ] **Step 1: Add isolated Node author smoke**

```js
const authorship = PptxDocument.create({ author: 'Packed & <Author>' });
const createdPresentationAuthor = authorship.author;
authorship.author = 'Edited author';
const editedPresentationAuthor = authorship.author;
const reopenedAuthorship = await PptxDocument.open(await authorship.write());
const reopenedPresentationAuthor = reopenedAuthorship.author;
authorship.author = '';
const emptyPresentationAuthor = authorship.author;
authorship.author = undefined;
const clearedPresentationAuthor = authorship.author;
```

Emit `presentationAuthor: true` only when all five states match.

- [ ] **Step 2: Mirror the lifecycle through the browser condition**

Create with `{ author: 'Browser author' }`, read, edit, write/reopen, set empty, and clear. Keep the existing static `node:` import rejection.

- [ ] **Step 3: Compile the declaration surface**

```ts
const authoredDocument: PptxDocument = PptxDocument.create({ author: 'Typed author' });
const authorSnapshot: string | undefined = authoredDocument.author;
authoredDocument.author = 'Edited typed author';
authoredDocument.author = '';
authoredDocument.author = undefined;
```

Include both values in the generated `void [...]` expression.

- [ ] **Step 4: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
presentation_author_package_dir=$(mktemp -d /tmp/pptx-presentation-author-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_author_package_dir"
presentation_author_tarball=$(find "$presentation_author_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_author_tarball"
cd ../..
```

Require every field true, including presentation title, presentation author, types, browser resolution, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m 'test: smoke packed presentation author metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 7: Document the contract and remaining metadata gaps

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Produces accurate author documentation without claiming lastModifiedBy or other metadata fields.

- [ ] **Step 1: Update public API documentation**

Add examples for create, edit, empty, and clear. Document direct creator snapshots, native canonical default, strict validation, relationship-based lookup, missing-part creation, same-value no-op, creator-only ownership, lastModifiedBy preservation, and malformed ambiguity rejection.

- [ ] **Step 2: Update compatibility and remaining gaps**

Add matrix row:

```md
| presentation `pptx.author` | `CreatePresentationOptions.author` / `document.author` | 已支持 strict string 创建、direct creator 读取/编辑/empty/clear；native 保留独立 lastModifiedBy |
```

Document PptxGenJS default `PptxGenJS`, native default `@jiayunxie/pptx`, exact custom/empty import, PptxGenJS mirroring, and native preservation. Keep company, revision, subject, lastModifiedBy, timestamps, custom properties, theme, sections, masters, and placeholders pending.

Add changelog bullet:

```md
- Added strict presentation-author metadata creation plus namespace-aware direct creator reading, lossless editing, empty/clear states, and last-modified-by preservation.
```

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'presentation author.*(unsupported|pending)|pptx\.author.*(未支持|尚未支持)|metadata.*全部支持' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md || true
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m 'docs: document presentation author metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 8: Run full gates and real-deck QA

**Files:**
- Review every Task 1–7 path; never stage `.pnpm-store/`.

**Interfaces:**
- Produces a fully verified pushed author feature; any defect gets a focused fix/review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-presentation-author-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-presentation-author-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Require every functional suite/test pass and the independent performance case below five seconds.

- [ ] **Step 2: Rebuild and smoke the final tarball plus doctor**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
presentation_author_qa_package_dir=$(mktemp -d /tmp/pptx-presentation-author-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_author_qa_package_dir"
presentation_author_qa_tarball=$(find "$presentation_author_qa_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_author_qa_tarball"
cd ../..
command -v pptx-inspect
pptx-inspect --json doctor
```

- [ ] **Step 3: Generate real native and PptxGenJS QA decks**

Create and record one fresh QA directory before generation:

```sh
qa_dir=$(mktemp -d /tmp/pptx-presentation-author-qa.XXXXXX)
printf '%s\n' "$qa_dir" > /tmp/pptx-presentation-author-qa-dir
```

Generate native default/source/edited/empty/cleared/reopened and PptxGenJS default/custom/empty decks inside `$qa_dir`. Each has one visible `wide` slide with identical audience-facing text and no metadata rendered on-canvas. Native custom author is `Alice & <Bob>`; edited author is `Edited author`; native lastModifiedBy remains `@jiayunxie/pptx` throughout. Every later QA step must restore the exact directory with `qa_dir=$(cat /tmp/pptx-presentation-author-qa-dir)` before using it.

- [ ] **Step 4: Validate packages and exact parts**

For every deck run:

```sh
pptx-inspect --json package inspect "$deck"
pptx-inspect --json package validate "$deck" --profile powerpoint-2010
pptx-inspect --json slides list "$deck"
pptx-inspect --json part read "$deck" /docProps/core.xml
```

Require zero errors/warnings; exact default/custom/empty/cleared creator; exact XML escaping; PptxGenJS creator/lastModifiedBy mirroring; native lastModifiedBy preservation.

- [ ] **Step 5: Prove package mutation isolation**

```sh
pptx-inspect --json package diff \
  "$qa_dir/native-source.pptx" "$qa_dir/native-edited.pptx"
pptx-inspect --json package diff \
  "$qa_dir/native-edited.pptx" "$qa_dir/native-reopened.pptx"
```

Require source-to-edited changes only `/docProps/core.xml`; edited-to-reopened has zero decompressed part changes.

- [ ] **Step 6: Render and check visual isolation**

Export native source, edited, and PptxGenJS custom through separate LibreOffice user profiles. Require one page, `960 ± 0.1 × 540 ± 0.1 pt` wide geometry, successful PNG rasterization, identical native source/edited PNG hashes, and no overflow:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  "$deck"
```

Inspect every rendered page individually at full size.

- [ ] **Step 7: Final static review and synchronization proof**

```sh
git diff --check
git status --short --branch
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require only user-owned `.pnpm-store/` untracked and `0 0`.
