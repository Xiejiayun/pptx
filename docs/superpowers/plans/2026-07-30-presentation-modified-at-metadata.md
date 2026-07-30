# Presentation Modified-At Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add strict native presentation modified-at creation plus namespace- and QName-aware typed `dcterms:modified` reading, lossless editing/repair/clear, PptxGenJS 4.0.1 public-output conformance, packed-surface proof, and real-deck QA.

**Architecture:** Extract the already-proven created-at W3CDTF lexical/calendar validator and qualified-type descriptor into one narrow internal timestamp helper. Keep created-at and modified-at in independent property wrappers, expose `PresentationModel.modifiedAt`, route explicit `CreatePresentationOptions.modifiedAt` through the same transactional setter, and import PptxGenJS write-time timestamps as final OOXML state without adding a native save-time clock side effect.

**Tech Stack:** TypeScript strict mode, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/model`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, npm pack smoke, `pptx-inspect`, LibreOffice headless.

## Global Constraints

- Modified-at owns only the unique direct `{http://purl.org/dc/terms/}modified` and its unique expanded-name-correct XSI type attribute in the root-related core-properties part.
- The required type QName resolves to `{http://purl.org/dc/terms/}W3CDTF`; lexical prefixes, part URI, declaration scope, and child order are not contracts.
- Public input is strict `string | undefined`, using `YYYY-MM-DDTHH:mm:ss[.fraction](Z|±HH:mm)` with valid Gregorian date/time and `00:00..14:00` timezone range.
- Empty, whitespace, date-only, timezone-less, invalid calendar, leap-second, lowercase-z, Unicode-digit, non-string, and coercible values reject before mutation.
- Native omitted and runtime-`undefined` creation preserve the canonical package and expose `undefined`; only explicit valid values write modified-at.
- `write()` does not read the clock or refresh modified-at; created, revision, lastModifiedBy, file time, and current time are not fallbacks.
- Existing created, title, subject, creator, lastModifiedBy, revision, company, unknown attributes/children, lexical formatting, relationships, and unrelated parts remain preserved.
- Missing core-properties state is created minimally; unsafe or ambiguous ownership rejects and outer transactions roll back.
- PptxGenJS conformance uses only public constructor, `addSlide()`, and `write()` because 4.0.1 exposes no modified timestamp property.
- Each task ends with focused review, explicit staging, commit, SSH port 443 push, fetch, and `origin/main...HEAD` equal to `0 0`.
- Execute inline because the user authorized autonomous continuation and repository instructions prohibit subagent dispatch.
- Never modify, delete, stage, or commit `.pnpm-store/`.

---

### Task 1: Extract the timestamp contract and add the modified-at codec

**Files:**
- Create: `packages/model/src/presentation-timestamp.internal.ts`
- Modify: `packages/model/src/presentation-created-at.internal.ts`
- Create: `packages/model/src/presentation-modified-at.internal.ts`
- Create: `packages/model/src/presentation-modified-at.internal.test.ts`

**Interfaces:**
- Produces `W3CDTF_QUALIFIED_TYPE: CoreQualifiedTypeDescriptor`.
- Produces `isPresentationTimestamp(value: string): boolean`.
- Produces `normalizePresentationTimestamp(value: unknown, propertyName: 'createdAt' | 'modifiedAt'): string | undefined`.
- Preserves `readPresentationCreatedAt(pkg)` and `replacePresentationCreatedAt(pkg, value)` signatures and behavior.
- Produces `readPresentationModifiedAt(pkg): string | undefined` and `replacePresentationModifiedAt(pkg, value): void`.

- [ ] **Step 1: Write failing namespace/type and zero-mutation reads**

Create `presentation-modified-at.internal.test.ts` with the same direct package fixture boundary used by the created-at codec. Require these exact positive cases:

```ts
expect(readPresentationModifiedAt(corePackage())).toBeUndefined();
expect(readPresentationModifiedAt(corePackage(coreXml(
  '<d:modified i:type="d:W3CDTF">2024-02-29T23:59:59.123456+14:00</d:modified>',
  'c:coreProperties',
  `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
)))).toBe('2024-02-29T23:59:59.123456+14:00');
```

Require `undefined` with exact package snapshot preservation for wrong element namespace, missing/wrong/unbound/XSI-wrong type, unprefixed QName, descendant-only modified, duplicate direct modified, duplicate expanded XSI type attributes, element children, CDATA, malformed XML, invalid root/content type, external/dangling/duplicate relationships, and lexical-invalid timestamp text.

- [ ] **Step 2: Write failing shared timestamp regression tests**

Test both wrappers against exact valid values:

```ts
const valid = [
  '0001-01-01T00:00:00Z',
  '2000-02-29T23:59:59.0Z',
  '2024-02-29T12:34:56.123456+05:30',
  '2026-07-30T00:00:00-00:30',
  '9999-12-31T23:59:59+14:00',
] as const;
```

Test modified-at rejection before mutation for:

```ts
const invalid = [
  '', ' 2026-07-30T00:00:00Z', '2026-07-30T00:00:00Z ',
  '0000-01-01T00:00:00Z', '1900-02-29T00:00:00Z',
  '2026-02-30T00:00:00Z', '2026-07-30',
  '2026-07-30T00:00Z', '2026-07-30T00:00:00',
  '2026-07-30T24:00:00Z', '2026-07-30T23:60:00Z',
  '2026-07-30T23:59:60Z', '2026-07-30T00:00:00z',
  '2026-07-30T00:00:00+14:01', '2026-07-30T00:00:00+15:00',
  null, true, 0, 1n, new Date(), {}, [], Symbol('modifiedAt'),
] as const;
```

Keep the complete existing created-at test file in the focused run so extraction cannot loosen or change its contract.

- [ ] **Step 3: Write failing mutation, repair, isolation, and rollback tests**

Require all of the following:

- missing-part creation emits one `xmlns:cp`, one `xmlns:dcterms`, one `xmlns:xsi`, and only direct typed modified;
- occupied `/docProps/core.xml` allocation and alternate part URI work;
- alternate element/XSI/QName prefixes and property-local declarations are reused correctly;
- normal/self-closing root insertion and self-closing modified expansion preserve opaque state;
- same text plus correct type is an exact bytes/journal no-op;
- same text plus missing/wrong unique type repairs only type state;
- a valid value repairs simple lexical-invalid text;
- a wrong-namespace local `type` remains while a correct XSI type is added;
- replacement and clear preserve created, creator, lastModifiedBy, revision, title, subject, comments, whitespace, unknown children, relationships, and unrelated part bytes;
- absent clear is a no-op; unsafe state rejects with zero mutation; outer transaction rollback restores part creation and edits.

Use exact isolation assertions:

```ts
replaceModifiedAt(pkg, '2026-07-30T01:02:03Z');
expect(readPresentationModifiedAt(pkg)).toBe('2026-07-30T01:02:03Z');
expect(partText(pkg)).toContain(
  '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-29T00:00:00Z</dcterms:created>',
);

replaceModifiedAt(pkg, undefined);
expect(partText(pkg)).not.toContain(':modified');
expect(readPresentationCreatedAt(pkg)).toBe('2026-07-29T00:00:00Z');
```

- [ ] **Step 4: Run RED**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-modified-at.internal.test.ts --reporter=dot
```

Expected: FAIL because `presentation-modified-at.internal.ts` does not exist.

- [ ] **Step 5: Add the shared timestamp helper**

Move the created-at qualified type object, regex, Gregorian checks, and normalizer into `presentation-timestamp.internal.ts`:

```ts
import { type CoreQualifiedTypeDescriptor } from './presentation-core-properties.internal.js';

export const W3CDTF_QUALIFIED_TYPE: CoreQualifiedTypeDescriptor = {
  attributeNamespace: 'http://www.w3.org/2001/XMLSchema-instance',
  attributePreferredPrefix: 'xsi',
  valueNamespace: 'http://purl.org/dc/terms/',
  valuePreferredPrefix: 'dcterms',
  valueLocalName: 'W3CDTF',
};

const PRESENTATION_TIMESTAMP_PATTERN =
  /^([0-9]{4})-([0-9]{2})-([0-9]{2})T([0-9]{2}):([0-9]{2}):([0-9]{2})(?:\.([0-9]+))?(Z|[+-][0-9]{2}:[0-9]{2})$/;

export function normalizePresentationTimestamp(
  value: unknown,
  propertyName: 'createdAt' | 'modifiedAt',
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !isPresentationTimestamp(value)) {
    throw new TypeError(
      `Presentation ${propertyName} must be a valid W3CDTF date-time string or undefined`,
    );
  }
  return value;
}
```

Implement `isPresentationTimestamp()` with the current created-at numeric logic exactly: year `>= 1`, Gregorian leap rule, month-day bounds, time bounds, and offset hour `< 14 || minute === 0`. Do not call `Date`, `Date.parse()`, or coerce values.

- [ ] **Step 6: Refactor created-at without behavior changes**

Change only its descriptor and wrappers:

```ts
const CREATED_AT_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'created timestamp',
  localName: 'created',
  namespace: 'http://purl.org/dc/terms/',
  preferredPrefix: 'dcterms',
  qualifiedType: W3CDTF_QUALIFIED_TYPE,
};

export function readPresentationCreatedAt(pkg: OpcPackage): string | undefined {
  const value = readCoreTextProperty(pkg, CREATED_AT_PROPERTY);
  return value !== undefined && isPresentationTimestamp(value) ? value : undefined;
}

export function replacePresentationCreatedAt(pkg: OpcPackage, value: string | undefined): void {
  replaceCoreTextProperty(
    pkg,
    CREATED_AT_PROPERTY,
    normalizePresentationTimestamp(value, 'createdAt'),
  );
}
```

Delete only the now-moved private regex/normalizer/predicate.

- [ ] **Step 7: Implement the modified-at wrapper**

Create:

```ts
const MODIFIED_AT_PROPERTY: CoreTextPropertyDescriptor = {
  label: 'modified timestamp',
  localName: 'modified',
  namespace: 'http://purl.org/dc/terms/',
  preferredPrefix: 'dcterms',
  qualifiedType: W3CDTF_QUALIFIED_TYPE,
};

export function readPresentationModifiedAt(pkg: OpcPackage): string | undefined {
  const value = readCoreTextProperty(pkg, MODIFIED_AT_PROPERTY);
  return value !== undefined && isPresentationTimestamp(value) ? value : undefined;
}

export function replacePresentationModifiedAt(
  pkg: OpcPackage,
  value: string | undefined,
): void {
  replaceCoreTextProperty(
    pkg,
    MODIFIED_AT_PROPERTY,
    normalizePresentationTimestamp(value, 'modifiedAt'),
  );
}
```

- [ ] **Step 8: Run focused and regression gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-modified-at.internal.test.ts \
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

Review shared-helper exactness, created-at regression, expanded-name/QName behavior, namespace declaration count, repair isolation, invalid zero mutation, and rollback. Then:

```sh
git add -- packages/model/src/presentation-timestamp.internal.ts \
  packages/model/src/presentation-created-at.internal.ts \
  packages/model/src/presentation-modified-at.internal.ts \
  packages/model/src/presentation-modified-at.internal.test.ts
git diff --cached --check
git commit -m "feat: add presentation modified-at core property codec"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Expose the live modified-at model property

**Files:**
- Modify: `packages/model/src/presentation.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes `readPresentationModifiedAt(pkg)` and `replacePresentationModifiedAt(pkg, value)`.
- Produces `PresentationModel.modifiedAt: string | undefined` with transactional setter semantics.

- [ ] **Step 1: Write the failing public model lifecycle test**

Add one test that opens an alternate-URI/prefix fixture containing independent valid created and modified values. Assert read-only access creates no mutations and retains stable slide identity. Then cover same-value no-op, invalid zero mutation, edit, unrelated part byte preservation, created-at isolation, lastModifiedBy/revision independence, outer rollback, clear, write/reopen, repair, absent clear, and missing-part creation.

Use these key assertions:

```ts
expect(model.createdAt).toBe('2024-02-29T12:34:56.123456+05:30');
expect(model.modifiedAt).toBe('2026-07-30T01:00:00Z');

model.modifiedAt = '2026-07-30T02:03:04.5+08:00';
expect(model.modifiedAt).toBe('2026-07-30T02:03:04.5+08:00');
expect(model.createdAt).toBe('2024-02-29T12:34:56.123456+05:30');

model.modifiedAt = undefined;
expect(model.modifiedAt).toBeUndefined();
expect(model.createdAt).toBe('2024-02-29T12:34:56.123456+05:30');
```

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts --reporter=dot
```

Expected: FAIL because `PresentationModel.modifiedAt` does not exist.

- [ ] **Step 3: Add the getter and transactional setter**

Import the wrapper and add adjacent to `createdAt`:

```ts
get modifiedAt(): string | undefined {
  return readPresentationModifiedAt(this.opcPackage);
}

set modifiedAt(value: string | undefined) {
  this.opcPackage.transaction(() => {
    replacePresentationModifiedAt(this.opcPackage, value);
  });
}
```

- [ ] **Step 4: Run focused, model, and type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-modified-at.internal.test.ts \
  packages/model/src/model.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/model --pretty false
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review getter purity, transaction nesting, stable identity, invalid zero mutation, created/modified isolation, rollback, missing-part lifecycle, and reopen. Then:

```sh
git add -- packages/model/src/presentation.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: expose presentation modified-at metadata"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Add native modified-at creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `CreatePresentationOptions.modifiedAt?: string`.
- Applies only an explicit non-`undefined` value through `document.modifiedAt = options.modifiedAt`.
- Keeps `PptxDocument.create()` synchronous and deterministic.

- [ ] **Step 1: Write failing create/edit/clear/rollback/reopen tests**

Cover omitted, runtime `undefined`, custom fractional/offset value, all six presentation formats, combined metadata options, same-value no-op, invalid values, edit, outer rollback, write/reopen, clear/reopen, and unchanged unrelated parts.

Use an independence case containing every current metadata field:

```ts
const combined = PptxDocument.create({
  author: 'Combined author',
  company: 'Combined company',
  createdAt: '2026-07-29T00:00:00Z',
  lastModifiedBy: 'Combined editor',
  modifiedAt: '2026-07-30T01:02:03.456+08:00',
  revision: '8',
  subject: 'Combined subject',
  title: 'Combined title',
});
expect([
  combined.createdAt,
  combined.modifiedAt,
  combined.lastModifiedBy,
  combined.revision,
]).toEqual([
  '2026-07-29T00:00:00Z',
  '2026-07-30T01:02:03.456+08:00',
  'Combined editor',
  '8',
]);
```

Verify two consecutive native `write()` calls preserve the exact modified value and core bytes rather than reading the clock.

- [ ] **Step 2: Write failing invalid creation tests**

Pass the same lexical/calendar/timezone/type-invalid vector from Task 1 through `PptxDocument.create({ modifiedAt })`. Every case must throw `TypeError`; creating a fresh valid document afterward must prove no shared state leaked.

- [ ] **Step 3: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts --reporter=dot
```

Expected: FAIL because `modifiedAt` is absent from `CreatePresentationOptions` and not applied.

- [ ] **Step 4: Add the create option and apply it explicitly**

In `create.ts`:

```ts
readonly modifiedAt?: string;
```

In `PptxDocument.create()`:

```ts
if (options.modifiedAt !== undefined) document.modifiedAt = options.modifiedAt;
```

Place the assignment adjacent to created-at/lastModifiedBy so source order reflects core metadata without coupling their state.

- [ ] **Step 5: Run focused, SDK, declaration, and package type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/presentation-modified-at.internal.test.ts \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/sdk packages/pptx --pretty false
pnpm typecheck
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review public type shape, omitted bytes, strict validation, all-format handling, combined option independence, deterministic writes, and no implicit clock. Then:

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create presentation modified-at metadata"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Prove PptxGenJS public-output conformance

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes only `new PptxGenJS()`, `addSlide()`, `write()`, and `importPptxGenJS()`.
- Proves `document.modifiedAt` imports the exact direct typed `dcterms:modified` final state.
- Proves native modification preserves the independent PptxGenJS-created `dcterms:created`.

- [ ] **Step 1: Add the public-output conformance test**

Generate a deck through public APIs, import it, and extract raw core XML from the imported package. Assert:

```ts
expect(imported.modifiedAt).toMatch(
  /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
);
const modifiedAt = imported.modifiedAt!;
const createdAt = imported.createdAt!;
expect(coreXml.match(
  /<dcterms:modified xsi:type="dcterms:W3CDTF">([^<]+)<\/dcterms:modified>/,
)?.[1]).toBe(modifiedAt);
```

Create a native document with `modifiedAt`, then edit the imported deck to an offset/fractional value. Assert only core.xml changes, created text remains exact, write/reopen preserves both timestamps, and PptxGenJS-origin slide content remains intact.

- [ ] **Step 2: Run focused adapter test**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
```

Expected: PASS after the prior tasks, proving the public surface rather than private PptxGenJS state.

- [ ] **Step 3: Review, commit, push, and verify**

Review that the test never reads `_slides` or mocks timestamp internals, compares raw and model state from one write, avoids cross-write timestamp equality assumptions, and proves created isolation. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare presentation modified-at metadata with pptxgenjs"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Prove the packed Node, browser, and declaration surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves packed runtime create/edit/write/reopen/clear in Node and browser-bundled paths.
- Proves emitted declarations accept `modifiedAt?: string` and expose `document.modifiedAt: string | undefined`.

- [ ] **Step 1: Extend Node packed runtime smoke**

Create chronology with both timestamps:

```js
const chronology = PptxDocument.create({
  createdAt: '2024-02-29T12:34:56.123+05:30',
  modifiedAt: '2024-03-01T01:02:03.456+08:00',
});
const createdPresentationModifiedAt = chronology.modifiedAt;
chronology.modifiedAt = '2026-07-30T01:02:03Z';
const editedPresentationModifiedAt = chronology.modifiedAt;
const reopenedChronology = await PptxDocument.open(await chronology.write());
const reopenedPresentationModifiedAt = reopenedChronology.modifiedAt;
chronology.modifiedAt = undefined;
const clearedPresentationModifiedAt = chronology.modifiedAt;
```

Add `presentationModifiedAt: true` to the stable JSON result only when all four states match and `createdAt` remains unchanged.

- [ ] **Step 2: Extend browser bundle smoke**

Use the same explicit modified values and assert create, edit, write/reopen, clear, and created isolation inside the generated browser module.

- [ ] **Step 3: Extend declaration compile smoke**

Add:

```ts
const modifiedAtDocument: PptxDocument = PptxDocument.create({
  modifiedAt: '2026-07-30T01:02:03Z',
});
const modifiedAtSnapshot: string | undefined = modifiedAtDocument.modifiedAt;
modifiedAtDocument.modifiedAt = '2024-02-29T12:34:56.123+05:30';
modifiedAtDocument.modifiedAt = undefined;
```

Include both variables in the final `void [...]` usage array.

- [ ] **Step 4: Run the real tarball smoke and full typecheck**

```sh
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
pptx_modified_pack_dir=$(mktemp -d /tmp/pptx-modified-at-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_modified_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_modified_pack_dir/jiayunxie-pptx-0.1.0.tgz"
git diff --check
```

Expected smoke JSON includes `"presentationModifiedAt":true`, browser assertions pass, declaration compilation succeeds, and CLI remains version `0.1.0`.

- [ ] **Step 5: Review, commit, push, and verify**

Review that smoke imports the packed tarball rather than workspace sources, exercises both environments, tests declarations, and leaves created-at independent. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed presentation modified-at metadata"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Document the modified-at contract and remaining gaps

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Documents `CreatePresentationOptions.modifiedAt` and `document.modifiedAt`.
- Records strict W3CDTF grammar, direct typed ownership, deterministic write behavior, and PptxGenJS clock difference.
- Removes modified timestamp from the remaining metadata gap without claiming custom properties, themes, sections, masters, or placeholders complete.

- [ ] **Step 1: Update examples and reference text**

Add `modifiedAt: '2024-03-01T01:02:03.456+08:00'` beside created-at in both public creation examples and show one edit plus clear. State that modified-at is independent from created, revision, and lastModifiedBy.

- [ ] **Step 2: Update compatibility and changelog**

Add one matrix row:

```md
| PptxGenJS write-time `dcterms:modified`（无独立 public field） | `CreatePresentationOptions.modifiedAt` / `document.modifiedAt` | 已支持 strict W3CDTF string创建、typed direct modified读取/编辑/repair/clear；native write不自动刷新 |
```

Document PptxGenJS public write-time behavior as an importable final state, native deterministic writes, namespace/QName resolution, exact no-ops, created isolation, and invalid-state policy. Change the remaining-gap paragraph from “modified timestamp、custom properties...” to “custom properties...”.

- [ ] **Step 3: Run documentation consistency checks**

```sh
rg -n "modifiedAt|dcterms:modified|modified timestamp" \
  CHANGELOG.md docs/api/README.md docs/compatibility/pptxgenjs-baseline.md \
  packages/pptx/README.md
git diff --check
pnpm typecheck
```

- [ ] **Step 4: Review, commit, push, and verify**

Review public names, examples, grammar, native/PptxGenJS distinction, no overclaim, and remaining-gap accuracy. Then:

```sh
git add -- CHANGELOG.md docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "docs: document presentation modified-at metadata"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run full gates and real-deck QA

**Files:**
- No tracked file changes expected.
- Store temporary evidence under `/tmp/pptx-modified-at-qa-20260730`.

**Interfaces:**
- Proves full repository compatibility, performance, real package semantics, validation, render stability, and no hidden save-time clock.

- [ ] **Step 1: Run full tests and performance gate**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
pnpm --filter @jiayunxie/pptx build
pptx_modified_qa_pack_dir=$(mktemp -d /tmp/pptx-modified-at-qa-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_modified_qa_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_modified_qa_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require all tests, the 5-second performance budget, packed Node/browser/types smoke, and CLI smoke to pass. A resource-only timeout must be isolated and then the original full command rerun successfully; do not accept only a focused retry.

- [ ] **Step 2: Generate deterministic real decks**

Create `/tmp/pptx-modified-at-qa-20260730` and generate native omitted, source, edited, cleared, reopened, second-write, and PptxGenJS public-output decks. Use one visible slide with stable text and these values:

```text
source createdAt  = 2024-02-29T12:34:56.123+05:30
source modifiedAt = 2024-03-01T01:02:03.456+08:00
edited modifiedAt = 2026-07-30T01:02:03Z
```

- [ ] **Step 3: Validate packages and exact core state**

Run `pptx-inspect package validate --profile powerpoint-2010` on every deck and require zero errors/warnings. Run `pptx-inspect part read --part /docProps/core.xml` on source/edited/cleared/reopened/PptxGenJS decks and assert exact typed modified state, offset/fraction preservation, created preservation, and clear removal.

- [ ] **Step 4: Prove mutation isolation and deterministic reopen**

Run package diff source→edited and require only the root-related core-properties part to change. Run edited→reopened and edited→second-write diffs and require zero part changes. Confirm no slide, relationship, content-type, app-property, or media part changes.

- [ ] **Step 5: Render and inspect visual stability**

Convert source, edited, and reopened decks to PDF with LibreOffice headless. Require the same page count and geometry. Render PDFs to PNG, require identical SHA-256 hashes, and run `slides_test.py` on all three without overflow findings. If Aptos Display is unavailable, use an explicit fontconfig mapping to the bundled Aptos font so hash comparison measures metadata isolation rather than font fallback.

- [ ] **Step 6: Final review and repository check**

Review implementation against every design section, inspect `git log`, and run:

```sh
git diff --check
git status --short --branch
git rev-list --left-right --count origin/main...HEAD
```

Expected: only user-owned `?? .pnpm-store/`, branch `main...origin/main`, and divergence `0 0`. Do not create an empty QA commit.
