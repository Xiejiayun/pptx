# PptxGenJS 4.0.1 Full-Surface Audit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a deterministic, fail-closed audit that enumerates every public PptxGenJS 4.0.1 capability atom and reports exactly which atoms have direct native implementation and evidence.

**Architecture:** Resolve the locked package from the adapter's dependency context, extract a stable declaration graph with the workspace TypeScript compiler, and cross-check it with a public-only runtime probe. Feed both into a strict manifest/evidence verifier and a deterministic CLI that generates reviewer JSON and Markdown without treating missing entries as supported.

**Tech Stack:** Node.js 22 ESM, TypeScript 5.9 compiler API, `node:test`, SHA-256, Git object verification, PptxGenJS 4.0.1, stable JSON/Markdown generation, pnpm, and existing packed-package/browser/PPTX evidence.

## Global Constraints

- The only compatibility baseline is the `pptxgenjs@4.0.1` package selected by `packages/pptxgenjs-adapter/package.json` and `pnpm-lock.yaml`.
- Never hard-code a `.pnpm` store path. Resolve the package entry from the adapter, walk to the nearest package root whose `package.json` name is `pptxgenjs`, and then resolve `types/index.d.ts` from that package root.
- Parse declarations with the workspace `typescript` dependency. Do not add a parser dependency and do not use regular expressions as the declaration authority.
- Public roots are the default `PptxGenJS` class, `PptxGenJS.Slide`, their public members, every input/output type reachable from those members, and reachable public enum/string-literal catalogs.
- Ignore private fields, underscore-only runtime helpers, ambient platform types outside the PptxGenJS declaration file, and PptxGenJS defects that do not represent legal final-state capability.
- Atom IDs, declaration ordering, JSON keys, Markdown groups, hashes, and diagnostics are deterministic and contain no absolute paths, timestamps, temporary paths, or machine-specific values.
- A declaration atom missing from the manifest is `unverified`; a manifest ID missing from declarations is `stale`; neither state may be converted to supported by aggregate smoke results.
- `supported`, `deliberate-difference`, `deprecated-alias`, and `defect-excluded` entries must satisfy their status-specific direct-evidence rules.
- `--write` may persist a truthful incomplete report. The default `--check` mode fails on artifact drift, invalid evidence, `unsupported`, `unverified`, or `stale`.
- Do not add the completeness command to the root green `check` chain until all open atoms are closed; expose it as an explicit release gate from the first CLI task.
- Every task ends with focused review, commit, SSH-443 push, remote-tracking refresh, and `HEAD...origin/main` divergence `0 0`.
- Stage only the exact files listed by each task. Never stage `.pnpm-store/` or retained external proof artifacts.

---

### Task 1: Locked Dependency Resolver and Declaration Atom Extractor

**Files:**
- Create: `scripts/pptxgenjs-surface-declarations.mjs`
- Create: `scripts/pptxgenjs-surface-declarations.test.mjs`

**Interfaces:**
- Consumes: `packages/pptxgenjs-adapter/package.json`, the resolved PptxGenJS package, declaration source text, and the workspace TypeScript compiler.
- Produces:

```js
export async function resolvePptxGenJSPackage(adapterPackagePath)
// => Object.freeze({ root, packageJsonPath, declarationPath, entryPath,
//                    version: '4.0.1', declarationSha256 })

export function extractPptxGenJSPublicSurface({
  sourceText,
  fileName = 'types/index.d.ts',
  typescript,
})
// => Object.freeze({ schemaVersion: 1, atoms, roots, diagnostics })
```

Each atom has this frozen shape:

```js
{
  id: 'interface:TextPropsOptions@property:margin',
  kind: 'property',
  owner: 'TextPropsOptions',
  name: 'margin',
  declaredIn: 'TextPropsOptions',
  optional: true,
  readonly: false,
  deprecated: false,
  signatures: [],
  deprecatedSignatures: [],
  typeText: 'Margin',
}
```

Root methods use `class:<root>#<name>`, root properties use `class:<root>@property:<name>`, slide methods use `method:Slide#<name>`, and slide properties use `property:Slide#<name>`. Method overloads share one atom, store all normalized signatures, and separately retain deprecated signatures. Named enum/union members use `union:<name>#<canonical-value>`; direct option unions use `union:<owning-atom-id>#<canonical-value>` and add `@path:<dot-path>` when a method parameter or nested property must disambiguate the owner. Inline records use `inline:<owning-atom-id>@property:<dot-path>`, so declaration reformatting cannot change their IDs.

- [ ] **Step 1: Write failing resolver, AST, reachability, and stability tests**

Create one inline fixture containing an overloaded default class, namespace `Slide`, inherited interfaces, optional/readonly properties, an inline nested record, recursive references, one enum, and one string-literal union:

```js
const FIXTURE = `
export default Deck;
declare class Deck {
  readonly mode: Deck.Mode;
  addSlide(options?: Deck.AddSlideOptions): Deck.Slide;
  addSlide(masterName?: string): Deck.Slide;
}
declare namespace Deck {
  export enum Mode { wide = 'wide', standard = 'standard' }
  export type Align = 'left' | 'right';
  export interface BaseOptions { readonly x?: number }
  export interface AddSlideOptions extends BaseOptions {
    align?: Align;
    nested?: { label?: string; next?: AddSlideOptions };
  }
  export class Slide {
    hidden: boolean;
    addText(text: string, options?: AddSlideOptions): Slide;
  }
}
`;
```

Assert exact stable IDs, two signatures on `class:Deck#addSlide`, `method:Slide#addText`, inherited `x` with `declaredIn: 'BaseOptions'`, one recursive traversal, inline `nested.label` / `nested.next`, and both enum/union values. Assert the extracted object and nested arrays are frozen. Add malformed-source, missing-root, duplicate-ID, unresolved in-file type, external `Promise`/`Array`, comment-only private member, and declaration-order permutation cases.

For the real resolver, assert package name/version `pptxgenjs@4.0.1`, a 64-character declaration hash, a declaration path below the discovered package root, and an error when a temporary fake package reports another version. The test must not contain `.pnpm` in its expected paths.

- [ ] **Step 2: Run the focused test and confirm failure**

```bash
node --test scripts/pptxgenjs-surface-declarations.test.mjs
```

Expected: FAIL because the resolver and extractor module does not exist.

- [ ] **Step 3: Implement package discovery, indexed AST traversal, and stable atoms**

Use `createRequire(resolve(adapterPackagePath))` to resolve the package entry. Starting at the entry directory, inspect parent `package.json` files until one has `name === 'pptxgenjs'`; validate `version === '4.0.1'`, resolve its `types` entry inside that root, append `index.d.ts` when the target is a directory, read bytes once, and hash those bytes.

Build declaration indexes before walking roots:

```js
const declarationsByName = new Map();
const namespaceMembers = new Map();
const visitedTypes = new Set();
const atomsById = new Map();
```

Derive the presentation root name from the `export default` assignment, then locate its class and same-name namespace; do not hard-code the fixture or production class name. Use TypeScript node kinds for classes, namespaces, interfaces, aliases, enums, properties, methods, type literals, intersections, arrays, tuples, indexed records, and literal unions. Normalize whitespace with the TypeScript printer, not source offsets. Traverse only identifiers resolved to declarations in the same source file; treat built-ins such as `string`, `number`, `Array`, `ReadonlyArray`, `Promise`, `Blob`, and `Function` as leaf types. Emit an inherited member under the reachable consumer interface ID with `declaredIn` set to its base; do not emit the base member again unless that base type is independently reachable outside the inheritance edge.

Sort atoms by `id`, sort overload signatures by normalized text, freeze all output, and throw diagnostics for missing public roots, duplicate IDs with conflicting metadata, parse errors, or unresolved PptxGenJS-local types.

- [ ] **Step 4: Run real extraction and focused gates**

```bash
node --test scripts/pptxgenjs-surface-declarations.test.mjs
node --input-type=module -e "import ts from 'typescript'; import { readFile } from 'node:fs/promises'; import { extractPptxGenJSPublicSurface, resolvePptxGenJSPackage } from './scripts/pptxgenjs-surface-declarations.mjs'; const pkg = await resolvePptxGenJSPackage('./packages/pptxgenjs-adapter/package.json'); const sourceText = await readFile(pkg.declarationPath, 'utf8'); const result = extractPptxGenJSPublicSurface({ sourceText, fileName: 'types/index.d.ts', typescript: ts }); if (result.atoms.length !== 1774) throw new Error('unexpected PptxGenJS 4.0.1 public surface'); console.log(JSON.stringify({ version: pkg.version, atoms: result.atoms.length }));"
node --check scripts/pptxgenjs-surface-declarations.mjs
node --check scripts/pptxgenjs-surface-declarations.test.mjs
git diff --check
```

Review stable IDs, inheritance origin, overload merging, recursive termination, external-type exclusion, exact version enforcement, absence of store-path assumptions, and zero absolute paths in extracted atom metadata.

- [ ] **Step 5: Commit, push, and verify synchronization**

```bash
git add scripts/pptxgenjs-surface-declarations.mjs \
  scripts/pptxgenjs-surface-declarations.test.mjs
git commit -m "feat: extract PptxGenJS public surface"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 2: Public Runtime Probe and Dependency Drift Fingerprint

**Files:**
- Create: `scripts/pptxgenjs-runtime-probe.mjs`
- Create: `scripts/pptxgenjs-runtime-probe.test.mjs`

**Interfaces:**
- Consumes: the Task 1 package descriptor and declaration surface.
- Produces:

```js
export async function probePptxGenJSRuntime({ packageInfo, surface })
// => Object.freeze({ schemaVersion: 1, packageVersion, declarationSha256,
//                    runtimeEntrySha256, classMembers, slideMembers,
//                    catalogs, minimalCalls })

export function hashRuntimeProbe(probe)
// => lowercase SHA-256 of canonical JSON
```

- [ ] **Step 1: Write failing public-only and deterministic probe tests**

Assert the real probe reports version `4.0.1`, the six output catalog values, all declared presentation methods, all declared `Slide` methods, and stable minimal call results:

```js
assert.deepEqual(probe.minimalCalls, {
  addSection: 'undefined',
  addSlide: 'Slide',
  addNotesReturnsSlide: true,
});
```

Run the probe twice and require deep equality plus equal hashes. Walk every result key and fail if a key starts with `_`, contains an absolute workspace prefix, or contains a function/object identity string. Inject fixtures with a missing declared runtime member, a catalog mismatch, a package-version mismatch, and a runtime entry outside the discovered root; require named diagnostics.

- [ ] **Step 2: Run the focused test and confirm failure**

```bash
node --test scripts/pptxgenjs-runtime-probe.test.mjs
```

Expected: FAIL because the runtime probe module does not exist.

- [ ] **Step 3: Implement declared-member projection and canonical probe hashing**

Load the runtime through `createRequire(packageInfo.packageJsonPath)(packageInfo.entryPath)`. Derive member names only from declaration atoms, then read/check those public names on a real presentation and slide. Access declared catalogs through `AlignH`, `AlignV`, `ChartType`, `OutputType`, `SchemeColor`, `ShapeType`, and `PlaceholderType`; sort catalog entries by key.

Call only `addSection({ title: 'Audit' })`, `addSlide()`, and `slide.addNotes('Audit note')`. Do not inspect or serialize `_slides`, `_rels`, generated XML, or other private state. Hash the runtime entry bytes and canonical JSON with recursively sorted object keys.

- [ ] **Step 4: Run focused gates, review, commit, and push**

```bash
node --test scripts/pptxgenjs-surface-declarations.test.mjs \
  scripts/pptxgenjs-runtime-probe.test.mjs
node --check scripts/pptxgenjs-runtime-probe.mjs
node --check scripts/pptxgenjs-runtime-probe.test.mjs
git diff --check
git add scripts/pptxgenjs-runtime-probe.mjs \
  scripts/pptxgenjs-runtime-probe.test.mjs
git commit -m "test: probe PptxGenJS public runtime"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review declaration-derived projection, catalog values, absence of private-state reads, deterministic hashes, version/hash coupling, and platform-neutral output. Expected divergence: `0 0`.

### Task 3: Strict Manifest Schema and Direct-Evidence Verifier

**Files:**
- Create: `scripts/pptxgenjs-surface-manifest.mjs`
- Create: `scripts/pptxgenjs-surface-audit-lib.mjs`
- Create: `scripts/pptxgenjs-surface-audit-lib.test.mjs`

**Interfaces:**
- Consumes: declaration atoms, runtime probe, a plain-data manifest, repository root, and Git object lookup.
- Produces:

```js
export const PPTXGENJS_SURFACE_MANIFEST = Object.freeze({
  schemaVersion: 1,
  packageVersion: '4.0.1',
  entries: [],
  extensions: [],
});

export async function buildPptxGenJSAudit({
  surface,
  runtimeProbe,
  manifest,
  repositoryRoot,
  gitCommitExists,
})
// => frozen report with counts, atoms, diagnostics, incompleteIds, extensions
```

Manifest evidence links have one exact shape:

```js
{
  path: 'packages/sdk/src/index.test.ts',
  pattern: "it('writes a presentation'",
  commit: '86fb1ec',
}
```

Test links use `title` instead of `pattern`. All paths are repository-relative. A supported serialized entry sets `serialization: true`; one requiring visual/client proof sets `client: true`. `deliberate-difference` and `defect-excluded` include `control`; `deprecated-alias` includes `canonical` and `control`.

- [ ] **Step 1: Write failing schema, status, evidence, and fail-closed tests**

Build a three-atom fixture and assert:

```js
assert.deepEqual(report.counts, {
  supported: 1,
  'deliberate-difference': 0,
  'deprecated-alias': 0,
  'defect-excluded': 0,
  unsupported: 0,
  unverified: 2,
  stale: 0,
});
```

Cover every allowed status, duplicate entries, unknown fields, wrong schema/version, a stale manifest ID, extension-ID collision, extension used to offset a gap, missing native mapping, missing status-specific `canonical`/`control`, and missing code/test/package/OOXML/client evidence.

Use a temporary repository fixture to cover missing files, absolute paths, parent traversal, unreadable files, absent literal pattern, absent exact test title, a valid commit, and an invalid commit. Require diagnostics sorted by atom ID then diagnostic code. Require all returned arrays and entries to be frozen.

- [ ] **Step 2: Run the focused test and confirm failure**

```bash
node --test scripts/pptxgenjs-surface-audit-lib.test.mjs
```

Expected: FAIL because the manifest and verifier modules do not exist.

- [ ] **Step 3: Implement exact schema normalization and evidence checks**

Define fixed key sets for the manifest, entries, status fields, evidence categories, and evidence links. Read own data descriptors only and reject accessors, symbols, inherited fields, class instances, duplicate IDs, empty strings, non-boolean flags, and unknown keys.

Verify evidence by resolving each path under `repositoryRoot`, checking that the normalized relative path remains inside it, reading the file as UTF-8, and searching for the exact literal `pattern` or `title`. Verify an optional commit without a shell by passing these arguments to `execFile`:

```bash
git cat-file -e <commit>^{commit}
```

Expose command execution as the injected `gitCommitExists` function so tests never depend on the workspace history. Never execute evidence text as a regular expression or shell fragment.

For each declaration atom, merge its valid manifest entry or create an `unverified` result. Append each valid extra manifest entry as `stale`. Compute counts from atom results, never from manifest claims. Keep extensions in a separate list and never include them in the denominator.

- [ ] **Step 4: Run focused gates, review, commit, and push**

```bash
node --test scripts/pptxgenjs-surface-audit-lib.test.mjs
node --check scripts/pptxgenjs-surface-manifest.mjs
node --check scripts/pptxgenjs-surface-audit-lib.mjs
node --check scripts/pptxgenjs-surface-audit-lib.test.mjs
git diff --check
git add scripts/pptxgenjs-surface-manifest.mjs \
  scripts/pptxgenjs-surface-audit-lib.mjs \
  scripts/pptxgenjs-surface-audit-lib.test.mjs
git commit -m "feat: verify PptxGenJS surface evidence"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review fail-closed defaults, status-specific requirements, descriptor safety, path containment, literal evidence matching, Git argument isolation, frozen output, and extension denominator isolation. Expected divergence: `0 0`.

### Task 4: Deterministic Audit CLI and Reviewer Artifacts

**Files:**
- Create: `scripts/pptxgenjs-surface-audit.mjs`
- Create: `scripts/pptxgenjs-surface-audit.test.mjs`
- Create: `docs/compatibility/pptxgenjs-surface-audit.json`
- Create: `docs/compatibility/pptxgenjs-surface-audit.md`
- Modify: `package.json`

**Interfaces:**
- Consumes: Tasks 1–3 modules and manifest.
- Produces these commands:

```bash
node scripts/pptxgenjs-surface-audit.mjs --write
node scripts/pptxgenjs-surface-audit.mjs --check
pnpm test:pptxgenjs-surface-audit
pnpm audit:pptxgenjs
pnpm audit:pptxgenjs:write
```

`--write` writes a truthful report and succeeds when extraction/schema/evidence processing is structurally valid. `--check` writes nothing and fails on byte drift, diagnostics, or any `unsupported`, `unverified`, or `stale` atom.

- [ ] **Step 1: Write failing CLI, determinism, and exit-status tests**

Inject a temporary output directory and fixture modules. Require two `--write` runs to produce byte-identical JSON and Markdown. Assert JSON top-level key order and the absence of absolute paths/timestamps. Assert Markdown groups in this exact order:

```js
const GROUPS = [
  'presentation', 'slide-lifecycle', 'text', 'shape', 'image',
  'media', 'chart', 'table', 'master-layout', 'output-runtime', 'other',
];
```

Require `--check` success for a complete fixture, nonzero status for artifact drift, and nonzero status with named IDs for unsupported/unverified/stale fixtures. Require invalid flags and simultaneous `--write --check` to fail without modifying output.

- [ ] **Step 2: Run the focused test and confirm failure**

```bash
node --test scripts/pptxgenjs-surface-audit.test.mjs
```

Expected: FAIL because the CLI and renderers do not exist.

- [ ] **Step 3: Implement canonical JSON, Markdown grouping, and atomic output**

Export `createAuditReport()`, `renderAuditJson()`, `renderAuditMarkdown()`, and `runAuditCli()` for tests. Canonical JSON uses recursively sorted object keys, two-space indentation, one trailing newline, and repository-relative evidence paths. Markdown contains version/hash, counts, completion state, diagnostics, then one table per non-empty group with atom ID, status, native mapping, and evidence links.

Classify groups from atom owner/name through a fixed pure map. Unknown owners go to `other`; group assignment never changes status or denominator.

For `--write`, create temporary sibling files with `open(..., 'wx')`, write/sync/close them, and rename both only after both renders succeed. Remove temporary siblings after a handled failure. For `--check`, generate in memory and byte-compare both committed artifacts before evaluating completion.

Add exact root scripts:

```json
{
  "test:pptxgenjs-surface-audit": "node --test scripts/pptxgenjs-surface-*.test.mjs scripts/pptxgenjs-runtime-probe.test.mjs",
  "audit:pptxgenjs": "node scripts/pptxgenjs-surface-audit.mjs --check",
  "audit:pptxgenjs:write": "node scripts/pptxgenjs-surface-audit.mjs --write"
}
```

- [ ] **Step 4: Generate the real initial artifacts and run gates**

```bash
pnpm audit:pptxgenjs:write
pnpm test:pptxgenjs-surface-audit
node --check scripts/pptxgenjs-surface-audit.mjs
pnpm typecheck
git diff --check
```

Run `pnpm audit:pptxgenjs` separately and require a nonzero result that names real incomplete atoms; at this stage that failure is the expected release-gate result, not a unit-test failure. Run `pnpm audit:pptxgenjs:write` again and require zero Git diff for both generated artifacts.

Review byte stability, atomic writes, truthful counts, group stability, no machine paths/timestamps, no mutation in check mode, and no accidental addition to the root `check` chain.

- [ ] **Step 5: Commit, push, and verify synchronization**

```bash
git add package.json scripts/pptxgenjs-surface-audit.mjs \
  scripts/pptxgenjs-surface-audit.test.mjs \
  docs/compatibility/pptxgenjs-surface-audit.json \
  docs/compatibility/pptxgenjs-surface-audit.md
git commit -m "feat: report PptxGenJS surface gaps"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Expected divergence: `0 0`.

### Task 5: First Trustworthy Gap Matrix and Follow-Up Queue

**Files:**
- Modify: `scripts/pptxgenjs-surface-manifest.mjs`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.json`
- Modify: `docs/compatibility/pptxgenjs-surface-audit.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `docs/implementation-progress.md`

**Interfaces:**
- Consumes: the generated atom inventory, all repository evidence that can satisfy Task 3 rules, and the previous manually maintained baseline.
- Produces: an evidence-backed initial manifest plus an ordered queue whose counts come only from generated `unsupported` and `unverified` atoms.

- [ ] **Step 1: Establish a reviewed classification batch**

Start with the presentation root and `Slide` root atoms only. For each atom, inspect its reachable option atoms and add an explicit manifest entry only when every required evidence link resolves. Mark confirmed legal gaps `unsupported`; leave uncertain atoms absent so the generator reports `unverified`. Use `deliberate-difference`, `deprecated-alias`, or `defect-excluded` only with a real PptxGenJS control test and the required native evidence.

Do not use broad entries such as “all text options supported,” do not attach one aggregate smoke pattern to unrelated atoms, and do not reduce unresolved counts with extensions.

- [ ] **Step 2: Regenerate and review the first matrix**

```bash
pnpm audit:pptxgenjs:write
pnpm test:pptxgenjs-surface-audit
pnpm typecheck
git diff --check
```

Review every explicit status change in the generated JSON. Confirm `declarationTotal` equals the sum of all status counts except `stale`, the stale count equals extra manifest IDs, the incomplete ID list is sorted/unique, every evidence link resolves, and the report still says incomplete whenever any open atom remains.

- [ ] **Step 3: Replace the obsolete aggregate completion statement**

In `docs/compatibility/pptxgenjs-baseline.md`, label the former public-capability 100% figure as a historical manual-checklist checkpoint and link to `pptxgenjs-surface-audit.md` as the current authority. In `docs/implementation-progress.md`, record the exact generated counts and the next gap family, without estimating unsupported atoms that remain unverified.

Order follow-up families by this deterministic priority:

1. shared coordinates, sizing, object metadata, and links;
2. presentation/slide lifecycle and theme cascade;
3. text and shape styling;
4. image lifecycle and styling;
5. media and chart options;
6. table, master, layout, and placeholder residuals;
7. final package/browser/PPTX/client certification.

Each family receives its own design, plan, focused tests, implementation, evidence, review, commit, and push. A family is removed from the queue only after regenerated status counts prove its atoms closed.

- [ ] **Step 4: Run completion-safety gates, review, commit, and push**

```bash
pnpm audit:pptxgenjs:write
pnpm test:pptxgenjs-surface-audit
pnpm typecheck
git diff --check
git add scripts/pptxgenjs-surface-manifest.mjs \
  docs/compatibility/pptxgenjs-surface-audit.json \
  docs/compatibility/pptxgenjs-surface-audit.md \
  docs/compatibility/pptxgenjs-baseline.md docs/implementation-progress.md
git commit -m "docs: publish PptxGenJS surface gaps"
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git push git@github.com:Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -p 443 -o HostName=ssh.github.com' \
  git fetch git@github.com:Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count HEAD...origin/main
```

Review direct evidence per changed atom, exact generated counts, unresolved IDs, documentation consistency, false-completion removal, and the next selected family. Expected divergence: `0 0`.

## Phase Completion

This implementation plan is complete when Task 5 is pushed and the repository has a reproducible declaration inventory, runtime fingerprint, strict evidence verifier, deterministic artifacts, and a truthful ordered gap queue. It does not claim full PptxGenJS parity.

Full parity is complete only after the follow-up family loop drives `unsupported = 0`, `unverified = 0`, and `stale = 0`; the explicit audit gate passes; packed Node/NodeNext/browser/CLI consumers pass; native and PptxGenJS control decks pass PPTX structural checks; and the required client corpus is recorded without unresolved failures.
