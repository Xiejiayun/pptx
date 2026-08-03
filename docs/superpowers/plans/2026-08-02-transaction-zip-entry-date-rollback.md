# Transaction ZIP Entry-Date Rollback Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make failed OPC transactions restore ZIP entry dates so rollback output remains byte-identical across clock boundaries.

**Architecture:** Capture detached per-entry dates in the existing private transaction savepoint. Reuse those dates only while reconstructing saved entries during rollback; successful writes and default/fixed entry-date behavior remain unchanged.

**Tech Stack:** TypeScript, JSZip, Vitest, OPC package transactions.

## Global Constraints

- Add no public API.
- Preserve successful create/open/write/compression behavior.
- Preserve existing fixed `PackageCreateOptions.entryDate` precedence.
- Restore only metadata that existed at the savepoint; do not synthesize dates for new entries.
- Keep the implementation scoped to `packages/opc/src/index.ts` and its test.
- Do not stage the six in-progress table-cell hyperlink documentation files or `.pnpm-store/` with this fix.

---

### Task 1: Preserve ZIP entry dates across rollback

**Files:**
- Modify: `packages/opc/src/index.ts`
- Modify: `packages/opc/src/index.test.ts`

**Interfaces:**
- Consumes: private `PackageSavepoint`, `OpcPackage.#createSavepoint()`, `OpcPackage.#restoreSavepoint()`, and `OpcPackage.#writeZipFile()`.
- Produces: private `PackageSavepoint.zipEntryDates` and restore-only date forwarding; no public type or method changes.

- [ ] **Step 1: Add the deterministic failing regression test**

Import `vi` from Vitest and add this case beside the current transaction rollback test:

```ts
it('restores ZIP entry dates when a transaction rolls back across time', async () => {
  vi.useFakeTimers({ toFake: ['Date'] });
  try {
    vi.setSystemTime(new Date('2026-08-02T00:00:00.000Z'));
    const pkg = OpcPackage.create();
    pkg.setPart('/data.xml', '<data>original</data>', 'application/xml');
    const before = await pkg.write({ compression: false });

    vi.setSystemTime(new Date('2026-08-02T00:02:00.000Z'));
    expect(() => pkg.transaction(() => {
      pkg.setPart('/data.xml', '<data>temporary</data>', 'application/xml');
      throw new Error('rollback across time');
    })).toThrow('rollback across time');

    expect(await pkg.write({ compression: false })).toEqual(before);
  } finally {
    vi.useRealTimers();
  }
});
```

- [ ] **Step 2: Run the regression test and verify the current failure**

Run:

```sh
node_modules/.bin/vitest run packages/opc/src/index.test.ts \
  -t "restores ZIP entry dates when a transaction rolls back across time" \
  --reporter=dot
```

Expected: the post-rollback ZIP differs only in per-entry DOS timestamp bytes.

- [ ] **Step 3: Capture and restore entry dates**

Extend the private savepoint:

```ts
interface PackageSavepoint {
  readonly parts: readonly MutablePart[];
  readonly journal: readonly MutationRecord[];
  readonly defaults: readonly (readonly [string, string])[];
  readonly overrides: readonly (readonly [string, string])[];
  readonly zipEntryDates: readonly (readonly [string, Date])[];
}
```

Capture copied non-directory dates in `#createSavepoint()`:

```ts
zipEntryDates: Object.entries(this.#zip.files)
  .filter(([, entry]) => !entry.dir)
  .map(([name, entry]) => [name, new Date(entry.date.getTime())] as const),
```

In `#restoreSavepoint()`, create a map before removing entries and pass the saved date for each restored part:

```ts
const zipEntryDates = new Map(savepoint.zipEntryDates);
// existing remove and part restoration
this.#writeZipFile(
  part.uri.slice(1),
  part.bytes,
  zipEntryDates.get(part.uri.slice(1)),
);
```

Allow the private writer to accept the restore-only date while retaining package-level fixed-date behavior:

```ts
#writeZipFile(name: string, bytes: Uint8Array, restoredDate?: Date): void {
  const date = restoredDate ?? this.#entryDate;
  if (date) {
    this.#zip.file(name, bytes, { date: new Date(date.getTime()) });
  } else {
    this.#zip.file(name, bytes);
  }
}
```

- [ ] **Step 4: Run focused OPC verification**

Run:

```sh
node_modules/.bin/vitest run packages/opc/src/index.test.ts --reporter=dot
```

Require every OPC test to pass, including transaction rollback, nested savepoints, compression, and fixed/default entry dates.

- [ ] **Step 5: Review, commit, push, and verify**

Run:

```sh
git diff --check -- packages/opc/src/index.ts packages/opc/src/index.test.ts
git add packages/opc/src/index.ts packages/opc/src/index.test.ts
git diff --cached --check
git commit -m "fix: preserve ZIP entry dates on rollback"
git push origin main
git fetch origin main
git rev-list --left-right --count HEAD...origin/main
```

Require divergence `0 0`. Then resume the table-cell hyperlink documentation gates; the full suite must pass without clock-boundary snapshot failures.
