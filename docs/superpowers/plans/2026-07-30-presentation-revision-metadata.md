# Presentation Revision Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation revision creation plus namespace-aware direct `cp:revision` reading, lossless editing, clear state, PptxGenJS 4.0.1 safe-output conformance, packed-surface proof, and real-deck QA.

**Architecture:** Reuse the descriptor-driven core-properties helper and add one revision-specific lexical wrapper. Make the smallest generic helper correction needed for a descriptor that shares the canonical root `cp` namespace, expose `PresentationModel.revision`, route `CreatePresentationOptions.revision` through the same transactional setter, then prove the contract independently at model, SDK, adapter, packed-package, documentation, and real-file layers.

**Tech Stack:** TypeScript strict mode, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/model`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, npm pack smoke, `pptx-inspect`, LibreOffice headless.

## Global Constraints

- Revision owns only the unique direct core-properties namespace `revision` in the root-related core-properties part.
- Public values are `string | undefined`; accepted strings match `^[0-9]+$` exactly and preserve their lexical form, including leading zeros.
- Native zero-input and create-time `undefined` preserve canonical revision `'1'`; assignment `undefined` clears only direct revision.
- No coercion, trimming, arithmetic increment, timestamp update, lastModifiedBy update, fallback, range conversion, or PptxGenJS invalid-runtime emulation.
- Existing title, subject, creator, lastModifiedBy, timestamps, unknown children, lexical formatting, relationships, and unrelated parts remain preserved.
- Missing core-properties state is created through the generic helper; unsafe or ambiguous ownership rejects before mutation and rolls back.
- Only PptxGenJS 4.0.1 public `revision`, `addSlide()`, and `write()` output count as conformance evidence.
- Each task ends with focused review, explicit staging, commit, SSH port 443 push, fetch, and `origin/main...HEAD` equal to `0 0`.
- Execute inline in the current task because the user authorized autonomous continuation and repository instructions prohibit subagent dispatch.
- Never modify, delete, stage, or commit `.pnpm-store/`.

---

### Task 1: Add the revision core-property wrapper

**Files:**
- Create: `packages/model/src/presentation-revision.internal.ts`
- Create: `packages/model/src/presentation-revision.internal.test.ts`
- Modify: `packages/model/src/presentation-core-properties.internal.ts`

**Interfaces:**
- Consumes `readCoreTextProperty(pkg, descriptor)` and `replaceCoreTextProperty(pkg, descriptor, value)` from `presentation-core-properties.internal.ts`.
- Produces internal `readPresentationRevision(pkg): string | undefined` and `replacePresentationRevision(pkg, value): void`.
- Preserves every existing title/author/subject helper output while allowing minimal `cp:revision` part creation without a duplicate namespace declaration.

- [ ] **Step 1: Write focused revision snapshot and lifecycle tests**

Create a fixture boundary with exact constants for the core-properties relationship, content type, core namespace, and Dublin Core namespace. Reuse the same package snapshot dimensions as the subject tests: part URIs, content types, byte arrays, relationships, content-type XML, and mutation journal.

Start the test file with the absent, alternate-prefix, lexical-valid, and lexical-invalid read cases:

```ts
import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationRevision,
  replacePresentationRevision,
} from './presentation-revision.internal.js';

expect(readPresentationRevision(corePackage())).toBeUndefined();
expect(readPresentationRevision(corePackage(coreXml(
  '<c:revision>007</c:revision><d:title>Quarterly</d:title>',
  'c:coreProperties',
  `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}"`,
)))).toBe('007');
expect(readPresentationRevision(corePackage(coreXml('<cp:revision>0</cp:revision>'))))
  .toBe('0');
expect(readPresentationRevision(corePackage(coreXml('<cp:revision/>'))))
  .toBeUndefined();
expect(readPresentationRevision(corePackage(coreXml('<cp:revision>1.5</cp:revision>'))))
  .toBeUndefined();
```

Require getter zero mutation for missing relationship, missing revision, zero, leading zeros, long digits, empty/self-closing, whitespace, sign, decimal, exponent, Unicode digits, wrong namespace, descendant-only, duplicate direct revisions, element-child revision, CDATA revision, wrong root/content type, dangling/external relationship, and duplicate root relationships.

Mutation assertions must prove:

- missing relationship plus `'7'` creates only a minimal revision part at the preferred or allocated URI with exactly one `xmlns:cp` declaration;
- alternate core prefix is reused for insertion;
- insert/replace/clear changes only direct revision and preserves title, subject, creator, lastModifiedBy, created/modified timestamps, comments, unknown children, unrelated bytes, and relationship identity;
- valid lexical same-value and absent clear are exact package snapshots;
- self-closing root and self-closing invalid revision can be expanded or repaired safely;
- simple invalid revision can be replaced by digits or cleared, while nested/CDATA unsafe revision rejects mutation;
- invalid runtime values and unsafe ownership leave parts, relationships, content types, mutation journal, and values unchanged;
- outer transaction rollback restores both minimal part creation and existing-part edits.

Use this invalid input table exactly:

```ts
const invalidValues = [
  '',
  ' ',
  '+1',
  '-1',
  '1.0',
  '1e3',
  '１２',
  null,
  false,
  0,
  1n,
  {},
  [],
  Symbol('revision'),
] as const;
```

- [ ] **Step 2: Run RED and prove the wrapper is absent**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-revision.internal.test.ts --reporter=dot
```

Expected: FAIL because `presentation-revision.internal.js` does not exist.

- [ ] **Step 3: Implement the strict revision wrapper**

```ts
import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const REVISION_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'revision',
  localName: 'revision',
  namespace: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  preferredPrefix: 'cp',
};

export function readPresentationRevision(pkg: OpcPackage): string | undefined {
  const value = readCoreTextProperty(pkg, REVISION_PROPERTY);
  return value !== undefined && isPresentationRevision(value) ? value : undefined;
}

export function replacePresentationRevision(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, REVISION_PROPERTY, normalizePresentationRevision(value));
}

function normalizePresentationRevision(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !isPresentationRevision(value)) {
    throw new TypeError(
      'Presentation revision must be a non-empty ASCII whole-number string or undefined',
    );
  }
  return value;
}

function isPresentationRevision(value: string): boolean {
  return /^[0-9]+$/.test(value);
}
```

- [ ] **Step 4: Correct same-namespace minimal part creation**

In `createCoreProperties()`, preserve existing Dublin Core bytes while conditionally omitting a redundant descriptor namespace declaration:

```ts
const qualifiedName = `${descriptor.preferredPrefix}:${descriptor.localName}`;
const preferredNamespaceDeclaration =
  descriptor.preferredPrefix === 'cp' && descriptor.namespace === CORE_PROPERTIES_NAMESPACE
    ? ''
    : ` xmlns:${descriptor.preferredPrefix}="${descriptor.namespace}"`;
const xml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
  + `<cp:coreProperties xmlns:cp="${CORE_PROPERTIES_NAMESPACE}"${preferredNamespaceDeclaration}>`
  + `<${qualifiedName}>${escapeXmlText(value)}</${qualifiedName}>`
  + '</cp:coreProperties>';
```

Do not change relationship allocation, prefix lookup, insertion, replacement, clear, or error behavior.

- [ ] **Step 5: Run focused tests and core-property regressions**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-revision.internal.test.ts \
  packages/model/src/presentation-title.internal.test.ts \
  packages/model/src/presentation-author.internal.test.ts \
  packages/model/src/presentation-subject.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

Expected: all tests and typecheck pass; title/author/subject minimal-part byte assertions remain unchanged.

- [ ] **Step 6: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation-core-properties.internal.ts \
  packages/model/src/presentation-revision.internal.ts \
  packages/model/src/presentation-revision.internal.test.ts
git diff --cached --check
git commit -m 'feat: add presentation revision core property codec'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 2: Expose the live revision model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes Task 1 revision wrapper.
- Produces public `PresentationModel.revision: string | undefined` with transactional mutation.

- [ ] **Step 1: Write a failing public model lifecycle test**

Build an alternate-URI/prefix core fixture containing revision, title, subject, creator, lastModifiedBy, timestamps, and an opaque child. Assert direct lexical behavior:

```ts
expect(model.revision).toBe('007');
model.revision = '42';
expect(model.revision).toBe('42');
expect(updated).toContain('<c:revision>42</c:revision>');
model.revision = '0009';
expect(model.revision).toBe('0009');
model.revision = undefined;
expect(model.revision).toBeUndefined();
```

Also require getter no mutation, valid same-value exact no-op, invalid-value isolation, invalid-existing snapshot returning `undefined`, valid repair, invalid simple-text clear, non-core-part byte identity, stable slide identity, rollback, revision-only clear, missing-part creation with one namespace declaration, absent-clear no-op, and write/reopen.

- [ ] **Step 2: Run RED and prove the public property is absent**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation revision metadata' --reporter=dot
```

Expected: TypeScript/Vitest failure because `PresentationModel.revision` is not declared.

- [ ] **Step 3: Add the public getter and setter**

Add the internal import and place the property with the other presentation metadata:

```ts
import {
  readPresentationRevision,
  replacePresentationRevision,
} from './presentation-revision.internal.js';

get revision(): string | undefined {
  return readPresentationRevision(this.opcPackage);
}

set revision(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationRevision(this.opcPackage, value);
  });
}
```

Do not export the revision helper from `packages/model/src/index.ts`.

- [ ] **Step 4: Run focused model tests and complete model gates**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation revision metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m 'feat: expose presentation revision metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 3: Add native revision creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `CreatePresentationOptions.revision?: string`.
- Makes zero-input and explicit-undefined creation preserve canonical `'1'` while explicit digits use `document.revision`.

- [ ] **Step 1: Write failing SDK creation and lifecycle tests**

Add one focused test beside the subject metadata test:

```ts
const omitted = PptxDocument.create();
const explicitUndefined = PptxDocument.create({ revision: undefined } as never);
const custom = PptxDocument.create({ revision: '7' });
const leading = PptxDocument.create({ revision: '007' });

expect([
  omitted.revision,
  explicitUndefined.revision,
  custom.revision,
  leading.revision,
]).toEqual(['1', '1', '7', '007']);
expect(readCoreXml(explicitUndefined)).toBe(readCoreXml(omitted));
expect(readCoreXml(custom)).toContain('<cp:revision>7</cp:revision>');
expect(readCoreXml(leading)).toContain('<cp:revision>007</cp:revision>');
```

For every `PresentationFormat`, create revision `42`, validate the package, write, reopen, and assert format plus lexical revision. Add combined author/company/subject/title/revision creation to prove only the canonical revision changes. Add same-value no-op, live edit, leading-zero edit, rollback, clear, reopened clear, other-part isolation, stable slide identity, and invalid input rejection.

Use the same invalid table from Task 1 and additionally verify every failing create call throws `TypeError` synchronously.

- [ ] **Step 2: Run RED and prove create options are absent**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation revision metadata' --reporter=dot
```

Expected: FAIL because `CreatePresentationOptions.revision` does not exist or explicit revision remains canonical `'1'`.

- [ ] **Step 3: Extend the option and route explicit values through the setter**

In `packages/sdk/src/create.ts`:

```ts
export interface CreatePresentationOptions {
  readonly author?: string;
  readonly company?: string;
  readonly format?: PresentationFormat;
  readonly revision?: string;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly subject?: string;
  readonly title?: string;
}
```

In `PptxDocument.create()`:

```ts
if (options.author !== undefined) document.author = options.author;
if (options.company !== undefined) document.company = options.company;
if (options.revision !== undefined) document.revision = options.revision;
if (options.title !== undefined) document.title = options.title;
if (options.subject !== undefined) document.subject = options.subject;
```

- [ ] **Step 4: Run SDK, type, and workspace regressions**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation revision metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/sdk --pretty false
pnpm typecheck
```

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m 'feat: create presentation revision metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 4: Prove PptxGenJS 4.0.1 revision conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Uses only the public PptxGenJS instance `revision`, `version`, `addSlide()`, and `write()`.
- Proves adapter import and native output for safe digit strings without broadening the adapter implementation.

- [ ] **Step 1: Extend the public test double surface**

Add the public field to `PptxGenJSInstance`:

```ts
revision: string;
```

- [ ] **Step 2: Write the conformance test**

Create default, zero, custom, and leading-zero instances:

```ts
const baseline = new PptxGenJS();
baseline.addSlide();
const zero = new PptxGenJS();
zero.revision = '0';
zero.addSlide();
const custom = new PptxGenJS();
custom.revision = '42';
custom.addSlide();
const leading = new PptxGenJS();
leading.revision = '007';
leading.addSlide();

const expectedRevisions = ['1', '0', '42', '007'] as const;
const imported = await Promise.all([
  importPptxGenJS(baseline),
  importPptxGenJS(zero),
  importPptxGenJS(custom),
  importPptxGenJS(leading),
]);
expect(imported.map(({ revision }) => revision)).toEqual(expectedRevisions);
```

Read `/docProps/core.xml` from each imported package and assert exact direct tokens. Write/reopen all imported documents and preserve lexical values. Compare `PptxDocument.create({ revision: '42' })` with the imported custom state and compare native default with PptxGenJS default.

Create PptxGenJS invalid runtime outputs for `''`, `'-1'`, `'1.5'`, and `'abc'`; import them and assert `document.revision` is `undefined` while raw core XML remains byte-preserved through no-op write/reopen. Assert native create rejects those same values.

- [ ] **Step 3: Run focused and complete adapter tests**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'presentation revision metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/pptxgenjs-adapter --pretty false
```

- [ ] **Step 4: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m 'test: compare presentation revision metadata with pptxgenjs'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 5: Cover the packed Node, browser, and declaration surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves the actual `@jiayunxie/pptx` tarball exposes revision creation, reading, editing, clear, reopen, browser behavior, and declaration types.

- [ ] **Step 1: Add Node packed-runtime checks**

Insert after presentation subject checks:

```js
const revisioned = PptxDocument.create({ revision: '007' });
const createdPresentationRevision = revisioned.revision;
revisioned.revision = '42';
const editedPresentationRevision = revisioned.revision;
const reopenedRevisioned = await PptxDocument.open(await revisioned.write());
const reopenedPresentationRevision = reopenedRevisioned.revision;
revisioned.revision = undefined;
const clearedPresentationRevision = revisioned.revision;
```

Add a `presentationRevision` check requiring `'007'`, `'42'`, reopened `'42'`, and final `undefined`. Do not add an empty-string success case because empty revision is invalid by contract.

- [ ] **Step 2: Add browser packed-runtime checks**

```js
const browserRevisioned = PptxDocument.create({ revision: '007' });
if (browserRevisioned.revision !== '007') {
  throw new Error('Browser presentation revision create failed');
}
browserRevisioned.revision = '42';
if (browserRevisioned.revision !== '42') {
  throw new Error('Browser presentation revision edit failed');
}
const reopenedBrowserRevisioned = await PptxDocument.open(await browserRevisioned.write());
if (reopenedBrowserRevisioned.revision !== '42') {
  throw new Error('Browser presentation revision reopen failed');
}
browserRevisioned.revision = undefined;
if (browserRevisioned.revision !== undefined) {
  throw new Error('Browser presentation revision clear failed');
}
```

- [ ] **Step 3: Add declaration compile checks**

```ts
const revisionDocument: PptxDocument = PptxDocument.create({ revision: '007' });
const revisionSnapshot: string | undefined = revisionDocument.revision;
revisionDocument.revision = '42';
revisionDocument.revision = undefined;
```

Include both identifiers in the final `void [...]` expression.

- [ ] **Step 4: Build the package and run actual tarball smoke**

```sh
pnpm --filter @jiayunxie/pptx build
node scripts/smoke-npm-package.mjs
```

Require Node, browser, declarations, and CLI output to pass and the JSON result to include `presentationRevision: true`.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m 'test: smoke packed presentation revision metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 6: Document the public revision contract

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Documents the same public names and values implemented in Tasks 2–5.
- Removes revision from the metadata backlog while retaining lastModifiedBy, timestamps, custom properties, theme, sections, masters, and placeholders.

- [ ] **Step 1: Add examples to API and package READMEs**

Extend creation examples with:

```ts
revision: '7',
```

Extend editing examples with:

```ts
created.revision = '008';
created.revision = undefined;
```

Add a revision contract paragraph stating:

- values are lexical `string | undefined` and only non-empty ASCII digits are accepted;
- zero-input native and PptxGenJS 4.0.1 both expose `'1'`;
- leading zeros are preserved exactly;
- assignment `undefined` removes only direct revision;
- getter does not synthesize a default for missing or invalid existing state;
- valid repair and clear preserve every adjacent core property and unrelated part;
- invalid PptxGenJS runtime output remains byte-preserved but is not exposed as a supported revision snapshot.

- [ ] **Step 2: Update the compatibility matrix and backlog**

Add this matrix row after subject/company metadata:

```markdown
| presentation `pptx.revision` | `CreatePresentationOptions.revision` / `document.revision` | 已支持 strict whole-number string 创建、direct revision 读取/编辑/clear；native/PptxGenJS 默认均为 `'1'` |
```

Replace the metadata backlog sentence so it starts with `lastModifiedBy、timestamps、custom properties` rather than `Revision、lastModifiedBy...`. Document PptxGenJS's runtime lack of validation as a non-parity bug boundary.

- [ ] **Step 3: Add the changelog entry**

```markdown
- Added strict presentation-revision metadata creation plus namespace-aware direct revision reading, lexical digit preservation, lossless editing/clear, and PptxGenJS 4.0.1 safe-output conformance.
```

- [ ] **Step 4: Scan for stale claims and run documentation checks**

```sh
rg -n "Revision、lastModifiedBy|revision.*(仍待|尚未支持|unsupported)|metadata.*revision" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md \
  packages/pptx/README.md || true
git diff --check
```

Review every hit in context. Claims about preserving adjacent revision in title/author/subject edits remain correct; only claims that revision itself is unavailable must be removed.

- [ ] **Step 5: Review, commit, push, fetch, and prove synchronization**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m 'docs: document presentation revision metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 7: Run full gates and real-deck QA

**Files:**
- No repository file changes expected.
- Create all temporary artifacts outside the tracked tree or under a temporary directory that is moved out before final status.

**Interfaces:**
- Consumes the complete revision public surface.
- Produces final evidence for compatibility, isolation, package stability, rendering, and remote synchronization.

- [ ] **Step 1: Run complete automated gates**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
pnpm --filter @jiayunxie/pptx build
node scripts/smoke-npm-package.mjs
pptx-inspect --json doctor
```

Require every TypeScript project, Vitest suite, performance harness, actual tarball Node/browser/declaration/CLI smoke, and doctor check to pass.

- [ ] **Step 2: Generate nine real decks**

Use the library under test for native output and PptxGenJS 4.0.1 public output. Every deck uses wide layout and the same visible centered Aptos Display 36pt bold line `Revision metadata never changes slide content`.

Generate:

```text
native-default.pptx
native-source.pptx
native-edited.pptx
native-leading.pptx
native-cleared.pptx
native-reopened.pptx
pptxgenjs-default.pptx
pptxgenjs-custom.pptx
pptxgenjs-leading.pptx
```

Use source `'7'`, edited `'42'`, leading `'007'`, and custom PptxGenJS `'42'`.

- [ ] **Step 3: Validate exact metadata and mutation isolation**

For every deck:

```sh
pptx-inspect --json package validate DECK --profile powerpoint-2010
pptx-inspect --json part read DECK /docProps/core.xml
```

Require zero errors/warnings and the exact expected direct revision. Then run:

```sh
pptx-inspect --json package diff native-source.pptx native-edited.pptx
pptx-inspect --json package diff native-edited.pptx native-reopened.pptx
```

Require source→edited to change only `/docProps/core.xml` and edited→reopened to report zero changed, added, or removed parts.

- [ ] **Step 4: Render and visually inspect representative files**

Use separate LibreOffice user profiles for native source, native edited, and PptxGenJS custom. Require one page at approximately `960 × 540 pt`, rasterize each PDF at the same DPI, and require identical PNG SHA-256 hashes.

Run `slides_test.py` on the three PPTX files. Inspect every rendered image at original detail and require the line to be centered, single-line, fully visible, unblurred, and identical with no clipping, overlap, repair dialog, fallback artifact, or off-slide content.

- [ ] **Step 5: Final review and synchronization proof**

```sh
git diff --check
git status --short --branch
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`, no tracked changes, and only `.pnpm-store/` untracked. No empty commit is created for QA-only evidence because Tasks 1–6 already contain reviewed and pushed material commits.
