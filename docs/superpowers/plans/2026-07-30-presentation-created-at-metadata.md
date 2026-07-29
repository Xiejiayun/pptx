# Presentation Created-At Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation created-at creation plus namespace- and QName-aware typed `dcterms:created` reading, lossless editing/repair/clear, PptxGenJS 4.0.1 public-output conformance, packed-surface proof, and real-deck QA.

**Architecture:** Extend the descriptor-driven core-properties helper with an optional qualified-type contract, leaving every untyped descriptor byte-compatible. Add one created-at wrapper for strict W3CDTF lexical/semantic validation, expose `PresentationModel.createdAt`, route explicit `CreatePresentationOptions.createdAt` through the same transactional setter, and treat PptxGenJS write-time timestamps as importable final state rather than hidden native save behavior.

**Tech Stack:** TypeScript strict mode, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/model`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, npm pack smoke, `pptx-inspect`, LibreOffice headless.

## Global Constraints

- Created-at owns only the unique direct `{http://purl.org/dc/terms/}created` and its unique expanded-name-correct XSI type attribute in the root-related core-properties part.
- The required type QName resolves to `{http://purl.org/dc/terms/}W3CDTF`; lexical prefixes, part URI, declaration scope, and child order are not contracts.
- Public input is strict `string | undefined`, using `YYYY-MM-DDTHH:mm:ss[.fraction](Z|±HH:mm)` with valid Gregorian date/time and `00:00..14:00` timezone range.
- Empty, whitespace, date-only, timezone-less, invalid calendar, leap-second, lowercase-z, Unicode-digit, non-string, and coercible values reject before mutation.
- Native omitted and runtime-`undefined` creation preserve the canonical package and expose `undefined`; only explicit valid values write created-at.
- `write()` does not read the clock or refresh created-at; modified, revision, lastModifiedBy, file time, and current time are not fallbacks.
- Existing modified, title, subject, creator, lastModifiedBy, revision, company, unknown attributes/children, lexical formatting, relationships, and unrelated parts remain preserved.
- Missing core-properties state is created minimally; unsafe or ambiguous ownership rejects and outer transactions roll back.
- PptxGenJS conformance uses only public constructor, `addSlide()`, and `write()` because 4.0.1 exposes no created timestamp property.
- Each task ends with focused review, explicit staging, commit, SSH port 443 push, fetch, and `origin/main...HEAD` equal to `0 0`.
- Execute inline because the user authorized autonomous continuation and repository instructions prohibit subagent dispatch.
- Never modify, delete, stage, or commit `.pnpm-store/`.

---

### Task 1: Add qualified core-property typing and the created-at codec

**Files:**
- Modify: `packages/model/src/presentation-core-properties.internal.ts`
- Create: `packages/model/src/presentation-created-at.internal.ts`
- Create: `packages/model/src/presentation-created-at.internal.test.ts`

**Interfaces:**
- Extends `CoreTextPropertyDescriptor` with optional `qualifiedType?: CoreQualifiedTypeDescriptor`.
- Preserves existing `readCoreTextProperty(pkg, descriptor)` and `replaceCoreTextProperty(pkg, descriptor, value)` signatures.
- Produces `readPresentationCreatedAt(pkg): string | undefined` and `replacePresentationCreatedAt(pkg, value): void`.

- [ ] **Step 1: Write failing typed read and zero-mutation tests**

Reuse the existing core relationship/content-type/package snapshot fixture boundary. Add constants for DCTERMS and XSI and begin with exact cases:

```ts
expect(readPresentationCreatedAt(corePackage())).toBeUndefined();
expect(readPresentationCreatedAt(corePackage(coreXml(
  '<d:created i:type="d:W3CDTF">2024-02-29T23:59:59.123456+14:00</d:created>',
  'c:coreProperties',
  `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
)))).toBe('2024-02-29T23:59:59.123456+14:00');
```

Require getter `undefined` and exact package snapshot preservation for:

```ts
const unsupported = [
  '<cp:created xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">2026-07-30T00:00:00Z</cp:created>',
  '<d:created>2026-07-30T00:00:00Z</d:created>',
  '<d:created i:type="d:Other">2026-07-30T00:00:00Z</d:created>',
  '<d:created i:type="missing:W3CDTF">2026-07-30T00:00:00Z</d:created>',
  '<d:created i:type="d:W3CDTF"><x:child xmlns:x="urn:test"/></d:created>',
  '<d:created i:type="d:W3CDTF"><![CDATA[2026-07-30T00:00:00Z]]></d:created>',
] as const;
```

Also cover descendant-only created, duplicate direct created, two differently-prefixed attributes resolving to XSI type, wrong root/content type, dangling/external relationship, duplicate relationships, invalid QName shapes, and lexical-invalid timestamps.

- [ ] **Step 2: Write failing lexical grammar tests**

Accepted values must round-trip exactly:

```ts
const valid = [
  '0001-01-01T00:00:00Z',
  '2000-02-29T23:59:59.0Z',
  '2024-02-29T12:34:56.123456+05:30',
  '2026-07-30T00:00:00-00:30',
  '9999-12-31T23:59:59+14:00',
] as const;
```

Invalid inputs must throw `TypeError` before any package mutation:

```ts
const invalid = [
  '', ' 2026-07-30T00:00:00Z', '2026-07-30T00:00:00Z ',
  '0000-01-01T00:00:00Z', '1900-02-29T00:00:00Z',
  '2026-02-30T00:00:00Z', '2026-07-30',
  '2026-07-30T00:00Z', '2026-07-30T00:00:00',
  '2026-07-30T24:00:00Z', '2026-07-30T23:60:00Z',
  '2026-07-30T23:59:60Z', '2026-07-30T00:00:00z',
  '2026-07-30T00:00:00+14:01', '2026-07-30T00:00:00+15:00',
  null, true, 0, 1n, new Date(), {}, [], Symbol('createdAt'),
] as const;
```

- [ ] **Step 3: Write failing mutation, repair, preservation, and rollback tests**

Require:

- minimal missing-part creation with one `xmlns:cp`, one `xmlns:dcterms`, one `xmlns:xsi`, and only direct typed created;
- occupied `/docProps/core.xml` allocation;
- alternate element/XSI prefixes and QName value prefix reuse;
- property-local namespace declarations;
- insert into normal/self-closing root;
- expand self-closing created while preserving opaque attributes;
- same text + correct type exact bytes/journal no-op;
- same text + missing/wrong unique type repairs only type state;
- valid value repairs simple lexical-invalid text;
- wrong-namespace local `type` remains while a correct type is added;
- replace and clear preserve modified, creator, lastModifiedBy, revision, title, subject, comments, whitespace, unknown children, relationships, and an unrelated part hash;
- absent clear no-op, unsafe ownership zero mutation, and outer transaction rollback for creation and edit.

Use explicit repair assertions:

```ts
replaceCreatedAt(pkg, '2026-07-30T00:00:00Z');
expect(partText(pkg)).toContain('i:type="d:W3CDTF"');
expect(readPresentationCreatedAt(pkg)).toBe('2026-07-30T00:00:00Z');

replaceCreatedAt(pkg, undefined);
expect(partText(pkg)).not.toContain(':created');
```

- [ ] **Step 4: Run RED**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-created-at.internal.test.ts --reporter=dot
```

Expected: FAIL because the created-at module and qualified type contract do not exist.

- [ ] **Step 5: Add the qualified-type descriptor contract**

Add exactly:

```ts
export interface CoreQualifiedTypeDescriptor {
  readonly attributeNamespace: string;
  readonly attributePreferredPrefix: string;
  readonly valueNamespace: string;
  readonly valuePreferredPrefix: string;
  readonly valueLocalName: string;
}

export interface CoreTextPropertyDescriptor {
  readonly label: string;
  readonly localName: string;
  readonly namespace: string;
  readonly preferredPrefix: string;
  readonly qualifiedType?: CoreQualifiedTypeDescriptor;
}
```

Add internal helpers that resolve expanded names from the element scope rather than comparing lexical prefixes:

```ts
function attributeNamespaceUri(element: XmlElement, name: string): string | undefined;
function attributesByExpandedName(
  element: XmlElement,
  namespace: string,
  localName: string,
): readonly XmlAttribute[];
function nonEmptyPrefixForNamespace(element: XmlElement, uri: string): string | undefined;
function qualifiedTypeMatches(
  element: XmlElement,
  descriptor: CoreQualifiedTypeDescriptor,
): boolean;
```

`qualifiedTypeMatches()` must require one expanded-name-correct attribute, a single-colon QName value, an in-scope prefix that resolves to `valueNamespace`, and exact `valueLocalName`.

- [ ] **Step 6: Extend read and mutation without changing untyped behavior**

In `readCoreTextProperty()`, keep the current simple-text read first, then apply the optional type check:

```ts
const property = state.properties[0]!;
const value = readSimpleProperty(state.xml, property);
if (value === undefined) return undefined;
if (descriptor.qualifiedType && !qualifiedTypeMatches(property, descriptor.qualifiedType)) {
  return undefined;
}
return value;
```

For non-`undefined` replacement, determine type state before the no-op check. A correct type and identical text returns immediately; otherwise emit non-overlapping source patches that:

```ts
// one correct expanded attribute, wrong QName value
xml.replaceAttribute(attribute, `${valuePrefix}:${type.valueLocalName}`);

// no correct expanded attribute
xml.replace(insertionPoint, insertionPoint,
  `${namespaceDeclaration} ${attributePrefix}:type="${valuePrefix}:${type.valueLocalName}"`,
);
```

When expanding a self-closing property, render one replacement element containing preserved start-tag attributes, any missing namespace declarations, the correct type attribute, escaped text, and the original qualified closing name. Reject duplicate expanded type attributes before adding patches. Keep clear behavior scoped to the unique safe simple property.

Update `createCoreProperties()` and `renderInsertedProperty()` so typed descriptors declare required element/value/attribute namespaces exactly once and render the type attribute; untyped output must remain byte-identical under existing tests.

- [ ] **Step 7: Implement the strict created-at wrapper**

Create the descriptor and validator:

```ts
const CREATED_AT_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'created timestamp',
  localName: 'created',
  namespace: 'http://purl.org/dc/terms/',
  preferredPrefix: 'dcterms',
  qualifiedType: {
    attributeNamespace: 'http://www.w3.org/2001/XMLSchema-instance',
    attributePreferredPrefix: 'xsi',
    valueNamespace: 'http://purl.org/dc/terms/',
    valuePreferredPrefix: 'dcterms',
    valueLocalName: 'W3CDTF',
  },
};

export function readPresentationCreatedAt(pkg: OpcPackage): string | undefined {
  const value = readCoreTextProperty(pkg, CREATED_AT_PROPERTY);
  return value !== undefined && isPresentationCreatedAt(value) ? value : undefined;
}

export function replacePresentationCreatedAt(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(pkg, CREATED_AT_PROPERTY, normalizePresentationCreatedAt(value));
}
```

Implement `isPresentationCreatedAt()` with the exact regex and numeric calendar checks from the design, not `Date.parse()`. Use the Gregorian leap rule `year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)` and enforce offset hour `< 14 || minute === 0`.

- [ ] **Step 8: Run focused and regression gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-created-at.internal.test.ts \
  packages/model/src/presentation-title.internal.test.ts \
  packages/model/src/presentation-author.internal.test.ts \
  packages/model/src/presentation-subject.internal.test.ts \
  packages/model/src/presentation-last-modified-by.internal.test.ts \
  packages/model/src/presentation-revision.internal.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 9: Review, commit, push, and verify**

Review expanded-name matching, QName resolution, namespace declaration duplication, patch overlap, self-closing expansion, invalid zero-mutation, untyped byte compatibility, and rollback. Then:

```sh
git add -- packages/model/src/presentation-core-properties.internal.ts \
  packages/model/src/presentation-created-at.internal.ts \
  packages/model/src/presentation-created-at.internal.test.ts
git diff --cached --check
git commit -m 'feat: add presentation created-at core property codec'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0` and only `.pnpm-store/` untracked.

---

### Task 2: Expose the live created-at model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes Task 1 `readPresentationCreatedAt()` / `replacePresentationCreatedAt()`.
- Produces `PresentationModel.createdAt: string | undefined`.

- [ ] **Step 1: Write a failing public lifecycle test**

Build an alternate core URI/prefix fixture containing typed created and modified, creator, lastModifiedBy, revision, title, subject, comments, unknown content, plus one slide and unrelated part. Assert:

```ts
expect(model.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
model.createdAt = '2026-07-30T00:00:00Z';
expect(model.createdAt).toBe('2026-07-30T00:00:00Z');
expect(updated).toContain('<d:created i:type="d:W3CDTF">2026-07-30T00:00:00Z</d:created>');
model.createdAt = undefined;
expect(model.createdAt).toBeUndefined();
```

Require getter zero mutation, same-value exact no-op, valid repair, invalid input isolation, modified/lastModifiedBy/revision independence, unrelated part identity, stable slide identity, rollback, missing-part creation, absent clear no-op, and write/reopen.

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation created-at metadata' --reporter=dot
```

Expected: FAIL because `PresentationModel.createdAt` is absent.

- [ ] **Step 3: Add the transactional accessor**

```ts
import {
  readPresentationCreatedAt,
  replacePresentationCreatedAt,
} from './presentation-created-at.internal.js';

get createdAt(): string | undefined {
  return readPresentationCreatedAt(this.opcPackage);
}

set createdAt(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationCreatedAt(this.opcPackage, value);
  });
}
```

Do not export the internal helper from `packages/model/src/index.ts`.

- [ ] **Step 4: Run focused, full model, and type gates**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t 'presentation created-at metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/model/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
```

- [ ] **Step 5: Review, commit, push, and verify**

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m 'feat: expose presentation created-at metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 3: Add native created-at creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `CreatePresentationOptions.createdAt?: string`.
- Applies only explicit values through the live setter.

- [ ] **Step 1: Write failing SDK creation tests**

```ts
const omitted = PptxDocument.create();
const explicitUndefined = PptxDocument.create({ createdAt: undefined } as never);
const custom = PptxDocument.create({
  createdAt: '2024-02-29T12:34:56.123456+05:30',
});

expect([omitted.createdAt, explicitUndefined.createdAt, custom.createdAt]).toEqual([
  undefined,
  undefined,
  '2024-02-29T12:34:56.123456+05:30',
]);
```

Require exact typed XML, no canonical-byte change for omitted/runtime undefined, all six formats creation/reopen, combination with every existing metadata option, same-value no-op, edit/clear, rollback, stable slide/unrelated parts, and the complete invalid matrix from Task 1 rejected during creation.

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation created-at metadata' --reporter=dot
```

- [ ] **Step 3: Extend create options and explicit setter routing**

Add after `company`:

```ts
readonly createdAt?: string;
```

In `PptxDocument.create()` apply it after company and before lastModifiedBy:

```ts
if (options.createdAt !== undefined) document.createdAt = options.createdAt;
```

Do not change `CORE_PROPERTIES_XML` or read the clock.

- [ ] **Step 4: Run SDK and workspace type gates**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts \
  -t 'presentation created-at metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run packages/sdk/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/sdk --pretty false
pnpm typecheck
```

- [ ] **Step 5: Review, commit, push, and verify**

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts \
  packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m 'feat: create presentation created-at metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 4: Prove PptxGenJS write-time created conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Uses only public PptxGenJS `version`, constructor, `addSlide()`, and `write()`.
- Proves real `dcterms:created` imports as strict `createdAt` and remains stable through native write/reopen.

- [ ] **Step 1: Add public-output conformance**

Generate one real output and import it:

```ts
const generated = new PptxGenJS();
expect(generated.version).toBe('4.0.1');
generated.addSlide();
const bytes = await generated.write({ outputType: 'uint8array', compression: true });
const imported = await PptxDocument.open(bytes);

expect(imported.createdAt).toMatch(
  /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
);
```

Read raw core XML from the public output and require exact text equality, correct DCTERMS element/XSI attribute namespaces, `W3CDTF` QName resolution, and no getter mutation. Capture `/dcterms:modified` raw bytes and all non-core part hashes.

- [ ] **Step 2: Prove native parity, edit isolation, and reopen stability**

```ts
const createdAt = imported.createdAt!;
const native = PptxDocument.create({ createdAt });
expect(native.createdAt).toBe(createdAt);

imported.createdAt = '2024-02-29T12:34:56.123+05:30';
expect(imported.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
const reopened = await PptxDocument.open(await imported.write());
expect(reopened.createdAt).toBe(imported.createdAt);
```

Require modified raw XML unchanged by the native edit, only core.xml changed, native omitted remains `undefined`, and no assertion that two independent PptxGenJS writes share a timestamp.

- [ ] **Step 3: Run focused and complete adapter gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts \
  -t 'created-at metadata' --reporter=dot
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/pptxgenjs-adapter --pretty false
```

- [ ] **Step 4: Review, commit, push, and verify**

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m 'test: compare presentation created-at metadata with pptxgenjs'
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
- Proves the actual tarball exposes create/read/edit/clear/reopen and declaration types in Node and browser bundles.

- [ ] **Step 1: Add Node runtime checks**

```js
const chronology = PptxDocument.create({
  createdAt: '2024-02-29T12:34:56.123+05:30',
});
const createdPresentationCreatedAt = chronology.createdAt;
chronology.createdAt = '2026-07-30T00:00:00Z';
const editedPresentationCreatedAt = chronology.createdAt;
const reopenedChronology = await PptxDocument.open(await chronology.write());
const reopenedPresentationCreatedAt = reopenedChronology.createdAt;
chronology.createdAt = undefined;
const clearedPresentationCreatedAt = chronology.createdAt;
```

Add `presentationCreatedAt: true` only when the four states match exact expectations.

- [ ] **Step 2: Add browser runtime checks**

Create with an offset/fraction timestamp, edit to UTC seconds, write/reopen, clear, and use explicit error messages for each failure. Do not use `Date` or a browser clock.

- [ ] **Step 3: Add declaration checks**

```ts
const createdAtDocument: PptxDocument = PptxDocument.create({
  createdAt: '2026-07-30T00:00:00Z',
});
const createdAtSnapshot: string | undefined = createdAtDocument.createdAt;
createdAtDocument.createdAt = '2024-02-29T12:34:56.123+05:30';
createdAtDocument.createdAt = undefined;
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

Require `presentationCreatedAt: true`, `types: true`, and CLI `0.1.0`.

- [ ] **Step 5: Review, commit, push, and verify**

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m 'test: smoke packed presentation created-at metadata'
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

---

### Task 6: Document the created-at contract

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Documents strict lexical input, typed OOXML ownership, deterministic native behavior, and PptxGenJS write-time clock difference.
- Changes the pending timestamp wording from plural to modified-only while retaining custom properties and save-policy gaps.

- [ ] **Step 1: Add create/edit examples and contract prose**

Add:

```ts
createdAt: '2024-02-29T12:34:56.123+05:30',
```

Then show edit and clear:

```ts
document.createdAt = '2026-07-30T00:00:00Z';
document.createdAt = undefined;
```

State exact grammar, semantic calendar validation, required typed state, namespace/QName handling, relationship-based discovery, missing-part lifecycle, repair/no-op behavior, adjacent-property preservation, and no automatic clock/save policy.

- [ ] **Step 2: Add the compatibility row and update backlog**

```markdown
| PptxGenJS write-time `dcterms:created`（无独立 public field） | `CreatePresentationOptions.createdAt` / `document.createdAt` | 已支持 strict W3CDTF string创建、typed direct created读取/编辑/repair/clear；native write不自动刷新 |
```

Replace generic `timestamps` backlog wording with `modified timestamp` wherever created-at is now covered, without claiming modified support.

- [ ] **Step 3: Add changelog and scan stale claims**

```markdown
- Added strict presentation created-at metadata creation plus namespace/QName-aware typed reading, lossless repair/clear editing, and PptxGenJS write-time timestamp conformance.
```

```sh
rg -n "createdAt.*(仍待|尚未支持|unsupported|pending)|created/modified|timestamps" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md \
  packages/pptx/README.md
git diff --check
```

Classify every match; retain statements that specifically identify modified or automatic save policy as pending.

- [ ] **Step 4: Review, commit, push, and verify**

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m 'docs: document presentation created-at metadata'
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
- Produces final automated, OOXML typing, mutation isolation, rendering, and synchronization evidence.

- [ ] **Step 1: Run complete automated gates**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
pnpm --filter @jiayunxie/pptx build
pnpm --filter @jiayunxie/pptx pack --pack-destination artifacts/npm
node scripts/smoke-npm-package.mjs \
  artifacts/npm/jiayunxie-pptx-0.1.0.tgz
pptx-inspect --json doctor
```

- [ ] **Step 2: Generate seven real decks**

Use the built library under test and PptxGenJS 4.0.1 public output to create:

```text
native-default.pptx
native-source.pptx
native-edited.pptx
native-cleared.pptx
native-reopened.pptx
native-offset-fraction.pptx
pptxgenjs-default.pptx
```

Use source `2024-02-29T12:34:56.123456+05:30`, edited/reopened `2026-07-30T00:00:00Z`, and clear as absence. Every deck has the same wide one-slide content `CreatedAt metadata never changes slide content`, centered Aptos Display 36pt bold with wrapping disabled.

- [ ] **Step 3: Validate exact typed metadata and isolation**

For all seven:

```sh
pptx-inspect --json package validate FILE --profile powerpoint-2010
pptx-inspect --json part read FILE /docProps/core.xml
```

Require zero errors/warnings, exact lexical values, correct element/XSI/type QName namespaces, native default/cleared absence, and PptxGenJS UTC-seconds created. Then:

```sh
pptx-inspect --json package diff native-source.pptx native-edited.pptx
pptx-inspect --json package diff native-edited.pptx native-reopened.pptx
```

Require source→edited only `/docProps/core.xml`; edited→reopened zero decompressed part changes.

- [ ] **Step 4: Render and inspect**

Use separate LibreOffice profiles for native source, native edited, and PptxGenJS default. Export to PDF and rasterize all at the same DPI. Require one approximately `960 × 540 pt` page, identical PNG SHA-256 hashes, and clean:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  FILE
```

Inspect all three rendered pages individually at original detail; require one centered line with no clipping, overlap, blur, fallback, repair, or unexpected wrapping.

- [ ] **Step 5: Final synchronization proof**

```sh
git diff --check
git status --short --branch
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git \
  main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Require `0 0`, no tracked changes, and only `.pnpm-store/` untracked. Do not create an empty QA commit.
