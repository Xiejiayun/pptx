# Presentation Company Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation-company metadata creation, direct-state reading, lossless editing, clearing, and PptxGenJS 4.0.1 conformance through the OOXML extended-properties part.

**Architecture:** Add one focused model-internal company helper that owns the package-root extended-properties relationship, validates the app-properties part/root, and source-span patches only direct `Company`. Expose a transactional `PresentationModel.company`, apply explicit creation input through the same setter, and prove the public behavior against PptxGenJS output, packed artifacts, and real PPTX package/render validation.

**Tech Stack:** TypeScript 5.9, Vitest 3.2, lossless OOXML source-span editing, OPC package transactions, PptxGenJS 4.0.1 public-output fixtures, tsup, npm tarballs, `pptx-inspect`, LibreOffice headless, Poppler.

## Global Constraints

- Public API is exactly `CreatePresentationOptions.company?: string` and `PresentationModel.company: string | undefined`.
- Native omitted creation keeps the current app-properties bytes and returns `undefined`; explicit empty string is direct empty state; assigning `undefined` clears direct Company.
- Company owns only direct `Company` in the extended-properties namespace; it must preserve Application, AppVersion, PresentationFormat, statistics, vectors, link state, unknown children, and all unrelated package state.
- PptxGenJS default/XML-safe-custom/empty Company states are imported exactly; its default `PptxGenJS`, coercion, and missing XML escaping are documented rather than copied by native behavior.
- Part URI and lexical namespace prefixes are not fixed; namespace URI and exact root relationship/content type define ownership.
- New Company is inserted before the first direct `LinksUpToDate`, `SharedDoc`, `HyperlinksChanged`, or `AppVersion`, otherwise appended, without reordering existing children.
- Same-value assignment and absent clear are exact bytes/journal no-ops.
- Missing extended-properties state may be created; dangling, external, wrong-content-type, wrong-root, duplicate-relationship, and duplicate-Company states must not be destructively guessed.
- Native inputs are strict strings without trimming or coercion; XML metacharacters are escaped and invalid XML controls fail before observable mutation.
- Existing title, author, RTL, slide size, format, creation, packed surface, and QA behavior must remain green.
- PptxGenJS is used only through public `company`, `version`, `addSlide()`, and `write()` output; private fields are forbidden.
- `.pnpm-store/` is user-owned and must never be modified, removed, staged, or committed.
- Every independently reviewed task is committed, pushed over SSH port 443, fetched, and verified with `origin/main...HEAD = 0 0` before the next task.

---

### Task 1: Add the extended-properties company codec

**Files:**
- Create: `packages/model/src/presentation-company.internal.ts`
- Create: `packages/model/src/presentation-company.internal.test.ts`

**Interfaces:**
- Produces internal `readPresentationCompany(pkg): string | undefined`.
- Produces internal `replacePresentationCompany(pkg, value): void` with strict normalization and source-span mutation.

- [ ] **Step 1: Write direct snapshot and namespace tests**

Build `OpcPackage.create()` fixtures with the exact extended-properties relationship/content type. Cover canonical default namespace, alternate `ep:` prefix, absent Company, `<Company/>`, escaped Unicode text, wrong-namespace direct Company, descendant-only Company, element-child Company, CDATA, duplicate direct Company, wrong root, wrong content type, dangling/external relationship, and multiple relationships.

Representative expectations:

```ts
expect(readPresentationCompany(appPackage())).toBeUndefined();
expect(readPresentationCompany(appPackage(
  '<ep:Properties xmlns:ep="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    + '<ep:Company>Acme &amp; 国际</ep:Company>'
    + '</ep:Properties>',
))).toBe('Acme & 国际');
expect(readPresentationCompany(appPackage(
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    + '<Company/>'
    + '</Properties>',
))).toBe('');
```

For every getter case snapshot all parts, relationships, content types, and `pkg.mutations`, then require zero mutation.

- [ ] **Step 2: Write lifecycle, isolation, and rollback tests**

Require all of these exact behaviors:

- missing relationship plus string creates `/docProps/app.xml`, its content-type override, and one root relationship;
- occupied `/docProps/app.xml` allocates `/docProps/app1.xml` without modifying the occupant;
- minimal creation contains only root plus Company and does not synthesize Application/AppVersion/statistics;
- default-namespace root inserts `<Company>`, prefixed root inserts the root prefix such as `<ep:Company>`;
- insertion occurs before the first following property, with separate cases for LinksUpToDate and AppVersion;
- self-closing root and self-closing Company expand safely;
- replacement correctly escapes `Acme & <Partners>` and reads back the decoded string;
- numeric-entity same-value and absent clear preserve exact bytes and journal;
- replace/empty/clear preserve Application, AppVersion, PresentationFormat, statistics, vectors, attributes, comments, whitespace, unknown children, relationships, and unrelated parts;
- duplicate direct Company and every unsafe relationship/part/root state reject before mutation;
- null, boolean, number, object, array, symbol, and `bad\u0001company` throw `TypeError` before package inspection or mutation;
- an outer `pkg.transaction()` rollback restores both existing-part edits and newly-created part/relationship/content-type state.

- [ ] **Step 3: Run RED and prove the helper is absent**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-company.internal.test.ts --reporter=dot
```

Expected: FAIL because `presentation-company.internal.js` does not exist.

- [ ] **Step 4: Implement strict normalization and ownership constants**

Use these exact constants and internal exports:

```ts
const EXTENDED_PROPERTIES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const EXTENDED_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const EXTENDED_PROPERTIES_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const FOLLOWING_COMPANY_PROPERTIES = new Set([
  'LinksUpToDate',
  'SharedDoc',
  'HyperlinksChanged',
  'AppVersion',
]);

export function readPresentationCompany(pkg: OpcPackage): string | undefined;

export function replacePresentationCompany(
  pkg: OpcPackage,
  value: string | undefined,
): void;
```

Normalize first:

```ts
function normalizePresentationCompany(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation company must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation company contains invalid XML characters');
  }
  return value;
}
```

- [ ] **Step 5: Implement read, create, replace, insert, and clear**

Parse only one internal root relationship whose target exists with the exact content type. Require one XML root whose local name is `Properties` and whose namespace URI is the extended-properties namespace. Company candidates are direct children with local name `Company` and the same namespace.

Read simple text only:

```ts
function readSimpleCompany(
  xml: LosslessXmlDocument,
  company: XmlElement,
): string | undefined {
  if (directChildren(company).length > 0) return undefined;
  if (/<!\[CDATA\[/i.test(xml.original(company))) return undefined;
  return xml.text(company);
}
```

Missing relationship plus string creates:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>VALUE</Company></Properties>
```

Use `escapeXmlText(value)`, `/docProps/app.xml` or `pkg.allocatePartUri('/docProps', 'app', '.xml')`, `pkg.setPart()`, and `pkg.addRelationship('/', { type, target: partUri.slice(1) })`.

For insertion, use the root lexical prefix to render either `Company` or `${prefix}:Company`; locate the first direct same-namespace following property and insert at its `start` offset, otherwise call `appendChildXml(root, companyXml)`. Replace non-self-closing content with `replaceText`, expand self-closing Company while preserving its original open tag/attributes, and clear with `removeElement`. Call `pkg.setPart()` only after a semantic change.

Stable unsafe-state errors are:

```text
Presentation has multiple extended-properties relationships
Extended-properties relationship must be internal
Extended-properties relationship target is missing
Extended-properties part has an unsupported content type
Extended-properties part must have one root
Extended-properties root is invalid
Extended properties contain multiple direct companies
Extended properties company is not simple text
```

- [ ] **Step 6: Run focused, model-wide, and type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-company.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 7: Review, commit, push, fetch, and prove synchronization**

Review that the helper is internal-only, no core-property code changed, normalization precedes package inspection, all unsafe paths are zero-mutation, prefix/default-namespace handling is URI-based, and `.pnpm-store/` is untouched.

```sh
git add -- packages/model/src/presentation-company.internal.ts \
  packages/model/src/presentation-company.internal.test.ts
git diff --cached --check
git commit -m 'feat: add presentation company extended property codec'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 2: Expose the live company model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes Task 1 company helper.
- Produces public `PresentationModel.company: string | undefined` with transactional mutation.

- [ ] **Step 1: Write failing public model lifecycle tests**

Build an alternate `/metadata/application.xml` fixture with prefixed root/Company and Application, AppVersion, PresentationFormat, statistics, vectors, opaque children, plus an unrelated slide part. Assert getter zero mutation, same-value exact no-op, escaped custom/empty values, other-property preservation, Company-only mutation, unrelated-part byte isolation, rollback, clear, missing-part creation, absent-clear no-op, write/reopen, and stable slide identity.

Use an invalid runtime assignment loop:

```ts
for (const company of [
  null,
  true,
  false,
  0,
  1,
  {},
  [],
  Symbol('company'),
  'bad\u0001company',
]) {
  expect(() => {
    presentation.company = company as never;
  }).toThrow(TypeError);
}
```

- [ ] **Step 2: Run RED and prove the model property is absent**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation company metadata' --reporter=dot
```

Expected: FAIL because `PresentationModel.company` is not declared.

- [ ] **Step 3: Add the public getter and transactional setter**

Import the Task 1 helpers and add:

```ts
get company(): string | undefined {
  return readPresentationCompany(this.opcPackage);
}

set company(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationCompany(this.opcPackage, value);
  });
}
```

Place it with author/title/RTL properties and do not export the internal helper from `packages/model/src/index.ts`.

- [ ] **Step 4: Run focused, complete model, and type gates**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation company metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m 'feat: expose presentation company metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 3: Add native company creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `CreatePresentationOptions.company?: string` and the complete create/edit/reopen lifecycle.

- [ ] **Step 1: Write failing SDK creation tests**

Require native omitted and runtime-undefined company to be `undefined` with byte-identical canonical app.xml; custom `Acme & <Partners>` to be escaped; direct empty Company; all six presentation formats; immediate edit; same-value no-op; coexistence with title/author/RTL; rollback; clear; write/reopen; app-part-only mutation; preservation of Application/AppVersion/PresentationFormat; and strict rejection of null, boolean, number, object, array, symbol, and XML controls.

- [ ] **Step 2: Run RED and prove the create option is absent**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation company metadata' --reporter=dot
```

- [ ] **Step 3: Extend create options and apply the shared setter**

Add the exact optional property:

```ts
readonly company?: string;
```

Apply it through the public model without changing `APP_PROPERTIES_XML`:

```ts
const document = new PptxDocument(createPresentationPackage(options));
if (options.author !== undefined) document.author = options.author;
if (options.company !== undefined) document.company = options.company;
if (options.title !== undefined) document.title = options.title;
return document;
```

Omitted creation must not add Company or alter any existing app-properties byte.

- [ ] **Step 4: Run focused SDK, SDK-wide, validator, and full type gates**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation company metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src \
  packages/validator/src --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m 'feat: create presentation company metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 4: Prove PptxGenJS company conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes PptxGenJS 4.0.1 public `company` and `write()`, adapter import, and native company API.
- Produces real-output evidence for default/XML-safe-custom/empty Company and the intentional native default/escaping differences.

- [ ] **Step 1: Extend the local public test interface**

Add only the public property needed by the test:

```ts
company: string;
```

- [ ] **Step 2: Add real public-output cases**

Create default, XML-safe custom, and empty PptxGenJS instances with one slide each:

```ts
const baseline = new PptxGenJS();
baseline.addSlide();
const custom = new PptxGenJS();
custom.company = 'Acme 国际';
custom.addSlide();
const empty = new PptxGenJS();
empty.company = '';
empty.addSlide();
```

Import only through `importPptxGenJS()` and require snapshots `PptxGenJS`, `Acme 国际`, and `''`; exact app.xml Company for all three; getter zero mutation; preservation of Application/AppVersion; write/reopen stability; PptxGenJS version `4.0.1`.

Create native custom and omitted company and require:

```ts
const native = PptxDocument.create({ company: 'Acme & <Partners>' });
const nativeOmitted = PptxDocument.create();
expect(native.company).toBe('Acme & <Partners>');
expect(nativeOmitted.company).toBeUndefined();
expect(nativeApp).toContain(
  '<Company>Acme &amp; &lt;Partners&gt;</Company>',
);
```

Keep malformed PptxGenJS markup outside the supported conformance cases, but prove the runtime limitation explicitly without private fields:

```ts
const unsafe = new PptxGenJS();
unsafe.company = 'A & <B>';
unsafe.addSlide();
const importedUnsafe = await importPptxGenJS(unsafe);
const unsafeApp = new TextDecoder().decode(
  importedUnsafe.opcPackage.requirePart('/docProps/app.xml').bytes,
);
expect(unsafeApp).toContain('<Company>A & <B></Company>');
expect(importedUnsafe.company).toBeUndefined();
```

This case proves raw interpolation is preserved but is not counted as a supported XML company state.

- [ ] **Step 3: Run focused, complete adapter, and type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'presentation company metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m 'test: compare presentation company metadata with pptxgenjs'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 5: Prove the packed Node, browser, and declaration surface

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Produces actual-tarball `presentationCompany: true`, browser lifecycle proof, and compile-time `string | undefined` coverage.

- [ ] **Step 1: Add isolated Node company smoke**

```js
const organization = PptxDocument.create({ company: 'Packed & <Company>' });
const createdPresentationCompany = organization.company;
organization.company = 'Edited company';
const editedPresentationCompany = organization.company;
const reopenedOrganization = await PptxDocument.open(await organization.write());
const reopenedPresentationCompany = reopenedOrganization.company;
organization.company = '';
const emptyPresentationCompany = organization.company;
organization.company = undefined;
const clearedPresentationCompany = organization.company;
```

Emit `presentationCompany: true` only when all five states match.

- [ ] **Step 2: Mirror the lifecycle through the browser condition**

Create with `{ company: 'Browser company' }`, read, edit, write/reopen, set empty, and clear. Keep the existing static `node:` import rejection and all previous browser assertions.

- [ ] **Step 3: Compile the declaration surface**

```ts
const companyDocument: PptxDocument = PptxDocument.create({ company: 'Typed company' });
const companySnapshot: string | undefined = companyDocument.company;
companyDocument.company = 'Edited typed company';
companyDocument.company = '';
companyDocument.company = undefined;
```

Include both values in the generated `void [...]` expression.

- [ ] **Step 4: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
presentation_company_package_dir=$(mktemp -d /tmp/pptx-presentation-company-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_company_package_dir"
presentation_company_tarball=$(find "$presentation_company_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_company_tarball"
cd ../..
```

Require every field true, including presentation title, author, company, types, browser resolution, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m 'test: smoke packed presentation company metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
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
- Produces accurate company documentation without claiming other extended/core/custom metadata fields.

- [ ] **Step 1: Update public API documentation**

Add create/edit/empty/clear examples and document direct Company snapshots, native omitted `undefined`, strict validation/escaping, relationship-based lookup, default/prefixed namespaces, missing-part creation, schema-friendly insertion, same-value no-op, Company-only ownership, other app-property preservation, and malformed/ambiguous rejection.

- [ ] **Step 2: Update compatibility and remaining gaps**

Add matrix row:

```md
| presentation `pptx.company` | `CreatePresentationOptions.company` / `document.company` | 已支持 strict string 创建、direct extended-properties Company 读取/编辑/empty/clear；native omitted 为 undefined |
```

Document PptxGenJS default `PptxGenJS`, native omitted `undefined`, exact XML-safe custom/empty import, native XML escaping, and PptxGenJS raw-interpolation limitation. Remove company from pending metadata lists while keeping revision, subject, lastModifiedBy, timestamps, custom properties, theme, sections, masters, and placeholders pending.

Add changelog bullet:

```md
- Added strict presentation-company metadata creation plus namespace-aware extended-properties reading, lossless editing, empty/clear states, and unrelated app-property preservation.
```

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'presentation company.*(unsupported|pending)|pptx\.company.*(未支持|尚未支持)|Company、revision|company, revision' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md || true
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m 'docs: document presentation company metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 7: Run full gates and real-deck QA

**Files:**
- Review every Task 1–6 path; never stage `.pnpm-store/`.

**Interfaces:**
- Produces a fully verified pushed company feature; any defect gets a focused fix/review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-presentation-company-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-presentation-company-vitest.json
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
presentation_company_qa_package_dir=$(mktemp -d /tmp/pptx-presentation-company-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_company_qa_package_dir"
presentation_company_qa_tarball=$(find "$presentation_company_qa_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_company_qa_tarball"
cd ../..
command -v pptx-inspect
pptx-inspect --json doctor
```

- [ ] **Step 3: Generate real native and PptxGenJS QA decks**

Create one fresh QA directory and retain the returned absolute path in the QA evidence:

```sh
qa_dir=$(mktemp -d /tmp/pptx-presentation-company-qa.XXXXXX)
```

Generate native omitted/source/edited/empty/cleared/reopened and PptxGenJS default/custom/empty decks inside it. Each has one visible `wide` slide with identical audience-facing text and no metadata rendered on-canvas. Native source company is `Source & <Company>`, edited company is `Edited & <Company>`, and PptxGenJS custom is the XML-safe `Acme 国际`. Every later QA command uses that exact absolute directory rather than discovering an older matching directory.

- [ ] **Step 4: Validate packages and exact app-properties parts**

For every deck run:

```sh
pptx-inspect --json package inspect "$deck"
pptx-inspect --json package validate "$deck" --profile powerpoint-2010
pptx-inspect --json slides list "$deck"
pptx-inspect --json part read "$deck" /docProps/app.xml
```

Require zero errors/warnings; exact omitted/custom/empty/cleared Company; exact native XML escaping; exact PptxGenJS default/custom/empty values; and Application/AppVersion/PresentationFormat preservation.

- [ ] **Step 5: Prove package mutation isolation**

```sh
pptx-inspect --json package diff \
  "$qa_dir/native-source.pptx" "$qa_dir/native-edited.pptx"
pptx-inspect --json package diff \
  "$qa_dir/native-edited.pptx" "$qa_dir/native-reopened.pptx"
```

Require source-to-edited changes only `/docProps/app.xml`; edited-to-reopened has zero decompressed part changes.

- [ ] **Step 6: Render and check visual isolation**

Export native source, edited, and PptxGenJS custom through separate LibreOffice user profiles. Require one page, `960 ± 0.1 × 540 ± 0.1 pt` wide geometry, successful PNG rasterization, identical native source/edited PNG SHA-256 hashes, and no overflow:

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
