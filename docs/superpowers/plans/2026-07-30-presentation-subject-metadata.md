# Presentation Subject Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation subject creation plus namespace-aware direct `dc:subject` reading, lossless editing, empty/clear states, PptxGenJS 4.0.1 conformance, packed-surface proof, and real-deck QA.

**Architecture:** Reuse the existing descriptor-driven core-properties helper and add one subject-specific internal wrapper. Expose `PresentationModel.subject`, route `CreatePresentationOptions.subject` through the same transactional setter, then prove the public contract independently at model, SDK, adapter, packed-package, documentation, and real-file layers.

**Tech Stack:** TypeScript strict mode, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/model`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, npm pack smoke, `pptx-inspect`, LibreOffice headless.

## Global Constraints

- Subject owns only the unique direct Dublin Core `dc:subject` in the root-related core-properties part.
- Native omitted/runtime-`undefined` creation remains `undefined`; explicit `''` writes direct empty subject; assignment `undefined` clears only subject.
- Input is strict XML-safe `string | undefined`; no coercion, trimming, fallback, inferred subject, or PptxGenJS brand default.
- Existing title, creator, lastModifiedBy, revision, timestamps, unknown children, lexical formatting, relationships, and unrelated parts remain preserved.
- Missing core-properties state is created through the existing generic helper; unsafe or ambiguous ownership rejects before mutation and rolls back.
- Only PptxGenJS 4.0.1 public `subject`, `addSlide()`, and `write()` output count as conformance evidence.
- Each task ends with focused review, explicit staging, commit, SSH port 443 push, fetch, and `origin/main...HEAD` equal to `0 0`.
- Execute inline in the current task because the user authorized autonomous continuation and repository instructions prohibit subagent dispatch.
- Never modify, delete, stage, or commit `.pnpm-store/`.

---

### Task 1: Add the subject core-property wrapper

**Files:**
- Create: `packages/model/src/presentation-subject.internal.ts`
- Create: `packages/model/src/presentation-subject.internal.test.ts`

**Interfaces:**
- Consumes `readCoreTextProperty(pkg, descriptor)` and `replaceCoreTextProperty(pkg, descriptor, value)` from `presentation-core-properties.internal.ts`.
- Produces internal `readPresentationSubject(pkg): string | undefined` and `replacePresentationSubject(pkg, value): void`.

- [ ] **Step 1: Write focused subject snapshots and lifecycle tests**

Define the fixture boundary explicitly, using `CORE_RELATIONSHIP`, `CORE_CONTENT_TYPE`, `CORE_NAMESPACE`, `DUBLIN_CORE_NAMESPACE`, `coreXml(children, root, namespaces)`, `corePackage(xml, options)`, `packageSnapshot(pkg)`, `partText(pkg, uri)`, and `setRootRelationships(pkg, relationships)`. The mutation wrapper must call `replacePresentationSubject()` inside `pkg.transaction()`. Begin with the exact import and descriptor-facing expectations:

```ts
import {
  readPresentationSubject,
  replacePresentationSubject,
} from './presentation-subject.internal.js';

expect(readPresentationSubject(corePackage())).toBeUndefined();
expect(readPresentationSubject(corePackage(coreXml(
  '<d:subject>主题 &amp; Forecast</d:subject><cp:revision>7</cp:revision>',
  'c:coreProperties',
  `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:cp="${CORE_NAMESPACE}"`,
)))).toBe('主题 & Forecast');
expect(readPresentationSubject(corePackage(coreXml('<dc:subject/>')))).toBe('');
```

Require getter zero mutation for absent, title-only, wrong namespace, descendant-only, duplicate direct subjects, element-child subject, CDATA subject, wrong root/content type, dangling/external relationship, and duplicate root relationships.

Mutation assertions must prove:

- missing relationship plus string creates only a minimal subject part at the preferred or allocated URI;
- alternate Dublin Core prefix is reused, while a missing binding inserts a local canonical `dc` declaration;
- insert/replace/empty/clear changes only direct subject and preserves title, creator, lastModifiedBy, revision, created/modified timestamps, comments, unknown children, unrelated bytes, and relationship identity;
- self-closing root and subject expand safely;
- numeric-entity same-value and absent clear are exact package snapshots;
- invalid runtime values and unsafe ownership leave parts, relationships, content types, mutation journal, and values unchanged;
- an outer transaction rollback restores both minimal part creation and existing-part edits.

- [ ] **Step 2: Run RED and prove the subject wrapper is absent**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-subject.internal.test.ts --reporter=dot
```

Expected: FAIL because `presentation-subject.internal.js` does not exist.

- [ ] **Step 3: Implement the strict subject wrapper**

```ts
import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const SUBJECT_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'subject',
  localName: 'subject',
  namespace: 'http://purl.org/dc/elements/1.1/',
  preferredPrefix: 'dc',
};

export function readPresentationSubject(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, SUBJECT_PROPERTY);
}

export function replacePresentationSubject(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, SUBJECT_PROPERTY, normalizePresentationSubject(value));
}

function normalizePresentationSubject(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation subject must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation subject contains invalid XML characters');
  }
  return value;
}
```

- [ ] **Step 4: Run focused tests and core-property regressions**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-subject.internal.test.ts \
  packages/model/src/presentation-title.internal.test.ts \
  packages/model/src/presentation-author.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation-subject.internal.ts \
  packages/model/src/presentation-subject.internal.test.ts
git diff --cached --check
git commit -m 'feat: add presentation subject core property codec'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 2: Expose the live subject model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes Task 1 subject wrapper.
- Produces public `PresentationModel.subject: string | undefined` with transactional mutation.

- [ ] **Step 1: Write failing public model lifecycle tests**

Build an alternate-URI/prefix core fixture containing subject, title, creator, lastModifiedBy, revision, and an opaque child. Assert:

```ts
expect(model.subject).toBe('Original & subject');
model.subject = 'Edited & <safe>';
expect(model.subject).toBe('Edited & <safe>');
expect(updated).toContain('<d:subject>Edited &amp; &lt;safe&gt;</d:subject>');
model.subject = '';
expect(model.subject).toBe('');
model.subject = undefined;
expect(model.subject).toBeUndefined();
```

Also require getter no mutation, same-value exact no-op, invalid-value isolation, non-core-part byte identity, stable slide identity, rollback, subject-only clear, missing-part creation, absent-clear no-op, write/reopen, and preservation of all adjacent core properties.

- [ ] **Step 2: Run RED and prove the public property is absent**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation subject metadata' --reporter=dot
```

Expected: TypeScript/Vitest failure because `PresentationModel.subject` is not declared.

- [ ] **Step 3: Add the public getter and setter**

Add the internal import and place the property with the other presentation metadata:

```ts
import {
  readPresentationSubject,
  replacePresentationSubject,
} from './presentation-subject.internal.js';

get subject(): string | undefined {
  return readPresentationSubject(this.opcPackage);
}

set subject(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationSubject(this.opcPackage, value);
  });
}
```

Do not export the subject helper from `packages/model/src/index.ts`.

- [ ] **Step 4: Run focused model tests and the complete model gates**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation subject metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m 'feat: expose presentation subject metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 3: Add native subject creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `CreatePresentationOptions.subject?: string` and complete create/edit/reopen lifecycle.

- [ ] **Step 1: Write failing SDK creation tests**

Require native omitted and runtime-undefined subject to be `undefined` with byte-identical core properties, escaped custom subject, direct empty subject, all six formats, same-value no-op, live edit, rollback, clear, write/reopen, stable slide identity, and strict invalid-input rejection.

Use a combined document to prove field independence:

```ts
const combined = PptxDocument.create({
  author: 'Combined author',
  company: 'Combined company',
  subject: 'Combined subject',
  title: 'Combined title',
});
expect([
  combined.author,
  combined.company,
  combined.subject,
  combined.title,
]).toEqual([
  'Combined author',
  'Combined company',
  'Combined subject',
  'Combined title',
]);
```

Mutation and clear assertions must prove creator, title, lastModifiedBy, revision, app properties, and every unrelated part remain unchanged.

- [ ] **Step 2: Run RED and prove the create option is missing**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation subject metadata' --reporter=dot
```

- [ ] **Step 3: Extend create options and apply the shared setter**

Add the exact optional property:

```ts
readonly subject?: string;
```

Apply subject after title so combined creation produces deterministic direct metadata order without changing any existing field's bytes when subject is absent:

```ts
const document = new PptxDocument(createPresentationPackage(options));
if (options.author !== undefined) document.author = options.author;
if (options.company !== undefined) document.company = options.company;
if (options.title !== undefined) document.title = options.title;
if (options.subject !== undefined) document.subject = options.subject;
return document;
```

Do not add subject to `CORE_PROPERTIES_XML`; omitted creation must preserve the current canonical bytes.

- [ ] **Step 4: Run SDK, validator, and full typecheck gates**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation subject metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src \
  packages/validator/src --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m 'feat: create presentation subject metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`.

---

### Task 4: Prove PptxGenJS subject conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes PptxGenJS 4.0.1 public `subject`, `addSlide()`, `write()`, adapter import, and native subject API.
- Produces real-output evidence for default/custom/empty subject plus the intentional native omitted-state difference.

- [ ] **Step 1: Add real public-output cases**

Add `subject: string` to the local `PptxGenJSInstance` interface, beside author/company/title. Create default, custom, and empty instances:

```ts
const baseline = new PptxGenJS();
baseline.addSlide();
const custom = new PptxGenJS();
custom.subject = 'Revenue & <Forecast>';
custom.addSlide();
const empty = new PptxGenJS();
empty.subject = '';
empty.addSlide();
```

Import only through `importPptxGenJS()` and require snapshots `PptxGenJS Presentation`, `Revenue & <Forecast>`, and `''`; exact escaped `dc:subject` XML; getter zero mutation; PptxGenJS version `4.0.1`; write/reopen stability; native explicit equality; and native omitted `undefined` without a direct subject element.

- [ ] **Step 2: Run focused and complete adapter tests plus typecheck**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'presentation subject metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src --reporter=dot
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 3: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m 'test: compare presentation subject metadata with pptxgenjs'
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
- Produces actual-tarball `presentationSubject: true`, browser lifecycle proof, and compile-time `string | undefined` coverage.

- [ ] **Step 1: Add isolated Node subject smoke**

```js
const subjectMatter = PptxDocument.create({ subject: 'Packed & <Subject>' });
const createdPresentationSubject = subjectMatter.subject;
subjectMatter.subject = 'Edited subject';
const editedPresentationSubject = subjectMatter.subject;
const reopenedSubjectMatter = await PptxDocument.open(await subjectMatter.write());
const reopenedPresentationSubject = reopenedSubjectMatter.subject;
subjectMatter.subject = '';
const emptyPresentationSubject = subjectMatter.subject;
subjectMatter.subject = undefined;
const clearedPresentationSubject = subjectMatter.subject;
```

Emit `presentationSubject: true` only when all five states match.

- [ ] **Step 2: Mirror the lifecycle through the browser condition**

Create with `{ subject: 'Browser subject' }`, read, edit, write/reopen, set empty, and clear. Keep the existing rejection of static `node:` imports in the browser bundle.

- [ ] **Step 3: Compile the declaration surface**

```ts
const subjectDocument: PptxDocument = PptxDocument.create({ subject: 'Typed subject' });
const subjectSnapshot: string | undefined = subjectDocument.subject;
subjectDocument.subject = 'Edited typed subject';
subjectDocument.subject = '';
subjectDocument.subject = undefined;
```

Include both values in the generated `void [...]` expression.

- [ ] **Step 4: Build and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
presentation_subject_package_dir=$(mktemp -d /tmp/pptx-presentation-subject-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_subject_package_dir"
presentation_subject_tarball=$(find "$presentation_subject_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_subject_tarball"
cd ../..
```

Require every field true, including presentation title/author/company/subject, types, browser resolution, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m 'test: smoke packed presentation subject metadata'
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
- Produces accurate subject documentation without claiming revision, lastModifiedBy, timestamps, custom properties, or other missing presentation features.

- [ ] **Step 1: Update public API documentation**

Add create/edit/empty/clear examples. Document direct subject snapshot, native omitted `undefined`, strict validation, relationship-based lookup, missing-part creation, same-value no-op, subject-only ownership, other-core-property preservation, and malformed ambiguity rejection.

- [ ] **Step 2: Update compatibility and remaining gaps**

Add the matrix row:

```md
| presentation `pptx.subject` | `CreatePresentationOptions.subject` / `document.subject` | 已支持 strict string 创建、direct subject 读取/编辑/empty/clear；native omitted 为 `undefined` |
```

Document PptxGenJS default `PptxGenJS Presentation`, native omitted `undefined`, exact custom/empty import, and XML escaping. Remove subject from pending metadata lists while keeping revision, lastModifiedBy, timestamps, custom properties, theme, sections, masters, and placeholders pending.

Add the changelog bullet:

```md
- Added strict presentation-subject metadata creation plus namespace-aware direct subject reading, lossless editing, empty/clear states, and adjacent core-property preservation.
```

- [ ] **Step 3: Scan contradictions and run typecheck**

```sh
git diff --check
rg -n --pcre2 \
  'presentation subject.*(unsupported|pending)|pptx\.subject.*(未支持|尚未支持)|metadata.*全部支持' \
  CHANGELOG.md docs/api docs/compatibility packages/pptx/README.md || true
node node_modules/typescript/bin/tsc -b --pretty false
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m 'docs: document presentation subject metadata'
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
- Produces a fully verified and pushed subject feature. Any defect receives its own focused fix, review, commit, push, fetch, and synchronization cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-presentation-subject-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-presentation-subject-vitest.json
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
presentation_subject_qa_package_dir=$(mktemp -d /tmp/pptx-presentation-subject-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$presentation_subject_qa_package_dir"
presentation_subject_qa_tarball=$(find "$presentation_subject_qa_package_dir" \
  -maxdepth 1 -name '*.tgz' -print -quit)
node ../../scripts/smoke-npm-package.mjs "$presentation_subject_qa_tarball"
cd ../..
command -v pptx-inspect
pptx-inspect --json doctor
```

- [ ] **Step 3: Generate real native and PptxGenJS QA decks**

Before artifact work, read and follow the available `presentations` and `pptx-inspect` skill instructions. Create one persisted temporary directory:

```sh
qa_dir=$(mktemp -d /tmp/pptx-presentation-subject-qa.XXXXXX)
printf '%s\n' "$qa_dir" > /tmp/pptx-presentation-subject-qa-dir
```

Generate native omitted/source/edited/empty/cleared/reopened and PptxGenJS default/custom/empty decks in that exact directory. Every deck has one visually identical `wide` slide; metadata is not rendered. Native custom subject is `Revenue & <Forecast>` and edited subject is `Edited subject`.

- [ ] **Step 4: Validate packages and exact parts**

For every deck run:

```sh
pptx-inspect --json package inspect "$deck"
pptx-inspect --json package validate "$deck" --profile powerpoint-2010
pptx-inspect --json slides list "$deck"
pptx-inspect --json part read "$deck" /docProps/core.xml
```

Require zero errors/warnings; exact native omitted/custom/edited/empty/cleared subject; exact XML escaping; PptxGenJS default/custom/empty subject; and preservation of title, creator, lastModifiedBy, revision, slide, and every unrelated part.

- [ ] **Step 5: Prove package mutation isolation**

```sh
pptx-inspect --json package diff \
  "$qa_dir/native-source.pptx" "$qa_dir/native-edited.pptx"
pptx-inspect --json package diff \
  "$qa_dir/native-edited.pptx" "$qa_dir/native-reopened.pptx"
```

Require source-to-edited changes only `/docProps/core.xml`; edited-to-reopened has zero decompressed part changes.

- [ ] **Step 6: Render and check visual isolation**

Export native source, native edited, and PptxGenJS custom through separate LibreOffice user profiles. Require one PDF page at `960 ± 0.1 × 540 ± 0.1 pt`, successful PNG rasterization, identical PNG hashes for all three files, no overflow, and full-size visual inspection of each rendered page.

Run the installed slide overflow checker for each deck:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  "$deck"
```

- [ ] **Step 7: Final static review and synchronization proof**

```sh
git diff --check
git status --short --branch
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require only user-owned `.pnpm-store/` untracked and `0 0`. Mark every checkbox complete only after the evidence exists.
