# Presentation Last Modified By Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation lastModifiedBy creation plus namespace-aware direct `cp:lastModifiedBy` reading, lossless editing, empty/clear states, PptxGenJS 4.0.1 author-mirror conformance, packed-surface proof, and real-deck QA.

**Architecture:** Reuse the descriptor-driven core-properties helper through one lastModifiedBy-specific wrapper. Expose an independent `PresentationModel.lastModifiedBy`, route `CreatePresentationOptions.lastModifiedBy` through the same transactional setter, and prove that native creator/lastModifiedBy ownership stays independent while PptxGenJS public author output imports as a mirrored final state.

**Tech Stack:** TypeScript strict mode, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/model`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, npm pack smoke, `pptx-inspect`, LibreOffice headless.

## Global Constraints

- LastModifiedBy owns only the unique direct core-properties namespace `lastModifiedBy` in the root-related core-properties part.
- Input is strict XML-safe `string | undefined`; empty string is a direct empty property, assignment `undefined` clears it.
- Native zero-input and create-time `undefined` preserve canonical `'@jiayunxie/pptx'`.
- Author and lastModifiedBy remain independently owned; neither setter changes the other.
- No automatic user lookup, save-time refresh, revision increment, timestamp update, fallback, coercion, or inferred audit policy.
- Existing title, subject, creator, revision, timestamps, unknown children, lexical formatting, relationships, and unrelated parts remain preserved.
- Missing core-properties state is created through the generic helper; unsafe or ambiguous ownership rejects before mutation and rolls back.
- PptxGenJS conformance uses only public `author`, `addSlide()`, and `write()` output because 4.0.1 has no public independent lastModifiedBy property.
- Each task ends with focused review, explicit staging, commit, SSH port 443 push, fetch, and `origin/main...HEAD` equal to `0 0`.
- Execute inline because the user authorized autonomous continuation and repository instructions prohibit subagent dispatch.
- Never modify, delete, stage, or commit `.pnpm-store/`.

---

### Task 1: Add the lastModifiedBy core-property wrapper

**Files:**
- Create: `packages/model/src/presentation-last-modified-by.internal.ts`
- Create: `packages/model/src/presentation-last-modified-by.internal.test.ts`

**Interfaces:**
- Consumes `readCoreTextProperty(pkg, descriptor)` and `replaceCoreTextProperty(pkg, descriptor, value)`.
- Produces internal `readPresentationLastModifiedBy(pkg): string | undefined` and `replacePresentationLastModifiedBy(pkg, value): void`.

- [ ] **Step 1: Write focused snapshots and lifecycle tests**

Use the same core relationship/content-type/package snapshot fixture boundary as revision. Begin with exact direct-state expectations:

```ts
expect(readPresentationLastModifiedBy(corePackage())).toBeUndefined();
expect(readPresentationLastModifiedBy(corePackage(coreXml(
  '<c:lastModifiedBy>编辑者 &amp; Reviewer</c:lastModifiedBy><c:revision>7</c:revision>',
  'c:coreProperties',
  `xmlns:c="${CORE_NAMESPACE}"`,
)))).toBe('编辑者 & Reviewer');
expect(readPresentationLastModifiedBy(corePackage(coreXml(
  '<cp:lastModifiedBy/>',
)))).toBe('');
```

Require getter zero mutation for absent, creator-only, empty, alternate prefix/URI, wrong namespace, descendant-only, duplicate direct fields, element child, CDATA, wrong root/content type, dangling/external relationship, and duplicate relationships.

Mutation tests must prove minimal missing-part creation with exactly one `xmlns:cp`, occupied URI allocation, prefix reuse, XML escaping, insert/replace/empty/clear, self-closing root/property expansion, semantic entity same-value no-op, absent clear no-op, preservation of creator/title/subject/revision/timestamps/comments/unknown/unrelated parts, invalid input isolation, unsafe ownership rejection, and outer transaction rollback.

Use this invalid table:

```ts
const invalidValues = [
  null,
  true,
  false,
  0,
  1n,
  {},
  [],
  Symbol('lastModifiedBy'),
  'bad\u0001editor',
] as const;
```

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-last-modified-by.internal.test.ts --reporter=dot
```

Expected: FAIL because the helper module does not exist.

- [ ] **Step 3: Implement the strict wrapper**

```ts
import { type OpcPackage } from '@pptx/opc';
import {
  readCoreTextProperty,
  replaceCoreTextProperty,
  type CoreTextPropertyDescriptor,
} from './presentation-core-properties.internal.js';

const LAST_MODIFIED_BY_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'last modified by',
  localName: 'lastModifiedBy',
  namespace: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  preferredPrefix: 'cp',
};

export function readPresentationLastModifiedBy(pkg: OpcPackage): string | undefined {
  return readCoreTextProperty(pkg, LAST_MODIFIED_BY_PROPERTY);
}

export function replacePresentationLastModifiedBy(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(
    pkg,
    LAST_MODIFIED_BY_PROPERTY,
    normalizePresentationLastModifiedBy(value),
  );
}

function normalizePresentationLastModifiedBy(value: unknown): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new TypeError('Presentation lastModifiedBy must be a string or undefined');
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError('Presentation lastModifiedBy contains invalid XML characters');
  }
  return value;
}
```

- [ ] **Step 4: Run focused and model regressions**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-last-modified-by.internal.test.ts \
  packages/model/src/presentation-revision.internal.test.ts \
  packages/model/src/presentation-title.internal.test.ts \
  packages/model/src/presentation-author.internal.test.ts \
  packages/model/src/presentation-subject.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, and verify**

```sh
git add -- packages/model/src/presentation-last-modified-by.internal.ts \
  packages/model/src/presentation-last-modified-by.internal.test.ts
git diff --cached --check
git commit -m 'feat: add presentation last modified by core property codec'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 2: Expose the live lastModifiedBy model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes Task 1 wrapper.
- Produces `PresentationModel.lastModifiedBy: string | undefined`.

- [ ] **Step 1: Write a failing public lifecycle test**

Build an alternate URI/prefix fixture with creator, lastModifiedBy, revision, title, subject, timestamps, and opaque content. Assert:

```ts
expect(model.lastModifiedBy).toBe('Original & editor');
model.lastModifiedBy = 'Edited & <safe>';
expect(model.lastModifiedBy).toBe('Edited & <safe>');
expect(updated).toContain('<c:lastModifiedBy>Edited &amp; &lt;safe&gt;</c:lastModifiedBy>');
model.lastModifiedBy = '';
expect(model.lastModifiedBy).toBe('');
model.lastModifiedBy = undefined;
expect(model.lastModifiedBy).toBeUndefined();
```

Also prove getter no mutation, same-value exact no-op, invalid input isolation, creator independence, unrelated part identity, stable slide identity, rollback, missing-part creation, absent clear no-op, and write/reopen.

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation last modified by metadata' --reporter=dot
```

Expected: initial getter returns `undefined` because the public accessor is absent.

- [ ] **Step 3: Add the model accessor**

```ts
import {
  readPresentationLastModifiedBy,
  replacePresentationLastModifiedBy,
} from './presentation-last-modified-by.internal.js';

get lastModifiedBy(): string | undefined {
  return readPresentationLastModifiedBy(this.opcPackage);
}

set lastModifiedBy(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationLastModifiedBy(this.opcPackage, value);
  });
}
```

Do not export the helper from `packages/model/src/index.ts`.

- [ ] **Step 4: Run focused, full model, and type gates**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation last modified by metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, and verify**

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m 'feat: expose presentation last modified by metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 3: Add native lastModifiedBy creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `CreatePresentationOptions.lastModifiedBy?: string`.
- Applies only explicit values through the live setter.

- [ ] **Step 1: Write failing SDK tests**

```ts
const omitted = PptxDocument.create();
const explicitUndefined = PptxDocument.create({ lastModifiedBy: undefined } as never);
const custom = PptxDocument.create({ lastModifiedBy: 'Editor & <Reviewer>' });
const empty = PptxDocument.create({ lastModifiedBy: '' });

expect([
  omitted.lastModifiedBy,
  explicitUndefined.lastModifiedBy,
  custom.lastModifiedBy,
  empty.lastModifiedBy,
]).toEqual([
  '@jiayunxie/pptx',
  '@jiayunxie/pptx',
  'Editor & <Reviewer>',
  '',
]);
```

Require exact XML escaping, six-format creation/reopen, author+lastModifiedBy independent combination, same-value no-op, edit/empty/clear, rollback, stable slide and unrelated parts, and invalid creation rejection.

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation last modified by metadata' --reporter=dot
```

- [ ] **Step 3: Extend create options and setter routing**

Add `readonly lastModifiedBy?: string` after `format`. In `PptxDocument.create()` apply it after author/company and before revision:

```ts
if (options.lastModifiedBy !== undefined) {
  document.lastModifiedBy = options.lastModifiedBy;
}
```

- [ ] **Step 4: Run SDK and workspace type gates**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation last modified by metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/sdk --pretty false
pnpm typecheck
```

- [ ] **Step 5: Review, commit, push, and verify**

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m 'feat: create presentation last modified by metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 4: Prove PptxGenJS author-mirror conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Uses only public PptxGenJS `author`, `version`, `addSlide()`, and `write()`.
- Proves imported lastModifiedBy mirrors creator without changing native independent ownership.

- [ ] **Step 1: Write default/custom/empty conformance**

```ts
const baseline = new PptxGenJS();
baseline.addSlide();
const custom = new PptxGenJS();
custom.author = 'Alice & <Bob>';
custom.addSlide();
const empty = new PptxGenJS();
empty.author = '';
empty.addSlide();

const expected = ['PptxGenJS', 'Alice & <Bob>', ''] as const;
const imported = await Promise.all([
  importPptxGenJS(baseline),
  importPptxGenJS(custom),
  importPptxGenJS(empty),
]);
expect(imported.map(({ lastModifiedBy }) => lastModifiedBy)).toEqual(expected);
expect(imported.map(({ author }) => author)).toEqual(expected);
```

Assert exact direct XML, no read mutation, write/reopen, native default difference, explicit native `{ author, lastModifiedBy }` final-state equality, and independent native edits where changing either field preserves the other.

- [ ] **Step 2: Run focused and complete adapter gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'last modified by metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/pptxgenjs-adapter --pretty false
```

- [ ] **Step 3: Review, commit, push, and verify**

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m 'test: compare presentation last modified by metadata with pptxgenjs'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 5: Cover packed Node, browser, and declarations

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves actual tarball create/read/edit/empty/clear/reopen and type declarations.

- [ ] **Step 1: Add Node runtime checks**

```js
const editorship = PptxDocument.create({ lastModifiedBy: 'Packed & <Editor>' });
const createdPresentationLastModifiedBy = editorship.lastModifiedBy;
editorship.lastModifiedBy = 'Edited editor';
const editedPresentationLastModifiedBy = editorship.lastModifiedBy;
const reopenedEditorship = await PptxDocument.open(await editorship.write());
const reopenedPresentationLastModifiedBy = reopenedEditorship.lastModifiedBy;
editorship.lastModifiedBy = '';
const emptyPresentationLastModifiedBy = editorship.lastModifiedBy;
editorship.lastModifiedBy = undefined;
const clearedPresentationLastModifiedBy = editorship.lastModifiedBy;
```

Add `presentationLastModifiedBy: true` only when all five states match.

- [ ] **Step 2: Add browser runtime checks**

Create `'Browser editor'`, edit, reopen, empty, and clear with explicit error messages for each state.

- [ ] **Step 3: Add declaration checks**

```ts
const lastModifiedDocument: PptxDocument = PptxDocument.create({
  lastModifiedBy: 'Typed editor',
});
const lastModifiedSnapshot: string | undefined = lastModifiedDocument.lastModifiedBy;
lastModifiedDocument.lastModifiedBy = 'Edited typed editor';
lastModifiedDocument.lastModifiedBy = '';
lastModifiedDocument.lastModifiedBy = undefined;
```

Include both identifiers in the final `void` expression.

- [ ] **Step 4: Build, pack, and smoke the exact artifact**

```sh
pnpm --filter @jiayunxie/pptx build
mkdir -p artifacts/npm
pnpm --filter @jiayunxie/pptx pack --pack-destination artifacts/npm
node scripts/smoke-npm-package.mjs \
  artifacts/npm/jiayunxie-pptx-0.1.0.tgz
```

Require `presentationLastModifiedBy: true`, types true, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, push, and verify**

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m 'test: smoke packed presentation last modified by metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 6: Document the contract

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Documents independent native ownership and PptxGenJS author-derived mirror state.
- Removes lastModifiedBy from the metadata backlog while retaining timestamps and custom properties.

- [ ] **Step 1: Add create/edit examples and contract prose**

Add `lastModifiedBy: 'Presentation Team'`, then edit/empty/clear examples. State canonical native default, strict XML-safe string, empty/clear distinction, relationship/prefix handling, missing-part lifecycle, exact no-ops, and adjacent-property preservation.

- [ ] **Step 2: Add compatibility row and update backlog**

```markdown
| PptxGenJS `author` mirrored `cp:lastModifiedBy`（无独立 public field） | `CreatePresentationOptions.lastModifiedBy` / `document.lastModifiedBy` | 已支持 strict string 创建、direct lastModifiedBy 读取/编辑/empty/clear；native 与 creator 独立 ownership |
```

Replace the backlog start `lastModifiedBy、timestamps、custom properties` with `timestamps、custom properties`.

- [ ] **Step 3: Add changelog and scan stale claims**

```markdown
- Added strict presentation-lastModifiedBy metadata creation plus namespace-aware direct reading, independent creator ownership, lossless empty/clear editing, and PptxGenJS author-mirror conformance.
```

```sh
rg -n "lastModifiedBy.*(仍待|尚未支持|unsupported|pending)|lastModifiedBy、timestamps" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md \
  packages/pptx/README.md || true
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m 'docs: document presentation last modified by metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 7: Run full gates and real-deck QA

**Files:**
- No tracked file changes expected.

**Interfaces:**
- Produces final automated, OOXML, isolation, rendering, and synchronization evidence.

- [ ] **Step 1: Run complete automated gates**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
pnpm --filter @jiayunxie/pptx build
node scripts/smoke-npm-package.mjs \
  artifacts/npm/jiayunxie-pptx-0.1.0.tgz
pptx-inspect --json doctor
```

- [ ] **Step 2: Generate nine real decks**

Create native default/source/edited/empty/cleared/reopened and PptxGenJS default/custom/empty-author decks. Use source `Source & <Editor>`, edited `Edited & <Editor>`, and PptxGenJS custom author `Custom & <Editor>`. Every deck has the same wide one-slide content: `LastModifiedBy metadata never changes slide content`, centered Aptos Display 36pt bold.

- [ ] **Step 3: Validate metadata and isolation**

Run PowerPoint 2010 validation and direct `/docProps/core.xml` reads on all nine. Require zero errors/warnings and exact escaped/empty/absent states. Source→edited may change only `/docProps/core.xml`; edited→reopened must have zero part changes.

- [ ] **Step 4: Render and inspect**

Use separate LibreOffice profiles for native source, native edited, and PptxGenJS custom. Require one approximately `960 × 540 pt` page, identical same-DPI PNG hashes, clean `slides_test.py`, and visually centered single-line content without clipping, overlap, blur, fallback, or repair.

- [ ] **Step 5: Final synchronization proof**

```sh
git diff --check
git status --short --branch
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`, no tracked changes, and only `.pnpm-store/` untracked. Do not create an empty QA commit.
