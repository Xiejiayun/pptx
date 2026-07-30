# Presentation Theme Fonts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add PptxGenJS-compatible presentation heading/body theme-font creation plus relationship-accurate reading, whole replacement, surgical existing-theme editing, adapter conformance, packed-surface proof, and real-deck QA.

**Architecture:** Put strict DrawingML font-scheme parsing and source-span mutation in a focused codec-internal helper, then make `ThemeModel` the live mutation boundary and resolve the primary theme only through the presentation part's unique direct theme relationship. Map that model into PptxGenJS naming at the SDK layer, preserving the native Aptos zero-input template while applying PptxGenJS Calibri fallbacks only for an explicit theme object.

**Tech Stack:** TypeScript strict mode, Vitest, `@pptx/lossless-xml`, `@pptx/opc`, `@pptx/codecs`, `@pptx/sdk`, PptxGenJS 4.0.1 public output, npm pack smoke, `pptx-inspect`, LibreOffice headless.

## Global Constraints

- Theme fonts own only the presentation-direct theme part's unique DrawingML `majorFont/latin@typeface` and `minorFont/latin@typeface` values.
- Theme selection follows the unique internal presentation theme relationship; it never guesses from part order, master/layout/notes relationships, or the number of theme files.
- XML lookup is direct-child and expanded-namespace correct; part URI and lexical namespace prefix are not contracts.
- Font inputs are descriptor-safe non-whitespace strings without XML 1.0 illegal characters; no coercion, trimming, getter invocation, or retained caller reference occurs.
- Zero-input native creation remains byte-identical with `Aptos Display` / `Aptos`; explicit partial/empty theme objects materialize PptxGenJS 4.0.1 fallbacks `Calibri Light` / `Calibri`.
- `document.theme` is a detached full snapshot and a whole-replacement setter; `ThemeModel.setFonts()` is a partial direct editor.
- Same-value operations are exact package/journal no-ops; invalid or ambiguous state rejects before mutation and outer transactions roll back atomically.
- Script fonts, East Asian/complex-script fonts, panose, theme colors, format scheme, extensions, comments, whitespace, relationships, and unrelated parts remain preserved.
- Adapter tests use only PptxGenJS public constructor, `theme`, `addSlide()`, and `write()`.
- Each task ends with focused review, explicit staging, commit, SSH port 443 push, fetch, and `origin/main...HEAD` equal to `0 0`.
- Execute inline because the user authorized autonomous continuation and repository instructions prohibit subagent dispatch.
- Never modify, delete, stage, or commit `.pnpm-store/`.

---

### Task 1: Add the strict theme-font codec and live ThemeModel editor

**Files:**
- Create: `packages/codecs/src/theme-fonts.internal.ts`
- Create: `packages/codecs/src/theme-fonts.internal.test.ts`
- Modify: `packages/codecs/src/master.ts`
- Modify: `packages/codecs/src/codecs.test.ts`

**Interfaces:**
- Produces `ThemeFontSnapshot { majorLatin: string; minorLatin: string }`.
- Produces `ThemeFontUpdate { majorLatin?: string; minorLatin?: string }`.
- Produces `readThemeFonts(xml): ThemeFontSnapshot | undefined`.
- Produces `replaceThemeFonts(xml, value): void`.
- Produces `ThemeModel.fonts: ThemeFontSnapshot | undefined` and `ThemeModel.setFonts(value): void`.
- Produces `MasterLayoutThemeCodec.presentationTheme: ThemeModel | undefined`.

- [ ] **Step 1: Write RED tests for namespace-correct reads**

Create `theme-fonts.internal.test.ts` around `LosslessXmlDocument.parse()` and require:

```ts
expect(readThemeFonts(themeXml())).toEqual({
  majorLatin: 'Aptos Display',
  minorLatin: 'Aptos',
});
expect(readThemeFonts(themeXml({ prefix: 'd', declarationsOnFontScheme: true }))).toEqual({
  majorLatin: 'Aptos Display',
  minorLatin: 'Aptos',
});
```

Require `undefined` for zero/repeated roots; wrong-namespace root; zero/repeated direct `themeElements`, `fontScheme`, `majorFont`, `minorFont`, or `latin`; descendant-only impostors; repeated unqualified `typeface`; empty/whitespace typeface; and element children under Latin. Require malformed XML to throw `LosslessXmlError` before `readThemeFonts()` is called. Assert every successful parse/read leaves `xml.changed === false`.

- [ ] **Step 2: Write RED tests for patching and isolation**

Require major-only, minor-only, and both updates; custom prefixes; missing unqualified `typeface` repair; preservation of a wrong-namespace `x:typeface`; XML escaping for `A&B <Display> "One"`; same-value no-op; detached input; and exact preservation of `panose`, `pitchFamily`, `charset`, `ea`, `cs`, script `font`, colors, format scheme, extensions, comments, and whitespace.

Use these isolation assertions:

```ts
const xml = LosslessXmlDocument.parse(source);
replaceThemeFonts(xml, { minorLatin: 'Noto Sans' });
expect(xml.serialize()).toContain('<a:latin typeface="Noto Sans" panose="020B"/>');
expect(xml.serialize()).toContain('<a:font script="Hans" typeface="等线"/>');
expect(xml.serialize()).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>');

const same = LosslessXmlDocument.parse(source);
replaceThemeFonts(same, { majorLatin: 'Aptos Display' });
expect(same.changed).toBe(false);
expect(same.serialize()).toBe(source);
```

Require invalid font values, accessor/symbol/unknown keys, arrays, custom-prototype objects, empty updates, unsafe structure, duplicate attributes, and missing Latin elements to throw without patches.

- [ ] **Step 3: Run RED**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/codecs/src/theme-fonts.internal.test.ts --reporter=dot
```

Expected: FAIL because `theme-fonts.internal.ts` does not exist.

- [ ] **Step 4: Implement the focused codec helper**

Use these public shapes and constants:

```ts
import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';

const DRAWINGML_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

export interface ThemeFontSnapshot {
  readonly majorLatin: string;
  readonly minorLatin: string;
}

export interface ThemeFontUpdate {
  readonly majorLatin?: string;
  readonly minorLatin?: string;
}
```

Implement direct-child navigation with in-scope namespace resolution from each element through its ancestors. A valid root chain has exactly one expected child at every level. `readThemeFonts()` requires one unqualified `typeface` attribute on both Latin elements, no element children, and two valid font strings.

Normalize `ThemeFontUpdate` through `Reflect.ownKeys()` and `Object.getOwnPropertyDescriptor()`. Allow only ordinary or null-prototype objects, only `majorLatin`/`minorLatin` own data properties, and at least one defined field. Reject values unless `typeof value === 'string'`, `/\S/u.test(value)`, and the existing XML 1.0 character contract passes.

For replacement, validate both targets before adding patches. Replace an existing attribute with `xml.replaceAttribute()`. Repair a missing unqualified attribute by inserting this escaped fragment immediately before the Latin start-tag close marker:

```ts
` typeface="${escapeXmlAttribute(value)}"`
```

Do not remove or rewrite other attributes. Do not save or serialize inside the helper.

- [ ] **Step 5: Wire ThemeModel and the presentation-direct relationship**

Re-export the two public value types at module scope:

```ts
export type { ThemeFontSnapshot, ThemeFontUpdate } from './theme-fonts.internal.js';
```

Then change the `ThemeModel` members to:

```ts

get fonts(): ThemeFontSnapshot | undefined {
  return readThemeFonts(this.codec.parse(this.partUri));
}

setFonts(value: ThemeFontUpdate): void {
  this.codec.pkg.transaction(() => {
    const xml = this.codec.parse(this.partUri);
    replaceThemeFonts(xml, value);
    if (xml.changed) this.codec.save(this.partUri, xml);
  });
}
```

Add a strict primary resolver:

```ts
get presentationTheme(): ThemeModel | undefined {
  const relationships = this.pkg.relationships(this.presentationPartUri)
    .filter(({ type }) => type === `${REL}theme`);
  if (relationships.length !== 1) return undefined;
  const [relationship] = relationships;
  if (relationship?.targetMode === 'External' || !relationship?.resolvedTarget) return undefined;
  const part = this.pkg.getPart(relationship.resolvedTarget);
  if (part?.contentType !== THEME_CONTENT_TYPE) return undefined;
  return this.modelForTheme(relationship.resolvedTarget);
}
```

Reuse a single `THEME_CONTENT_TYPE` constant in ownership, `themes`, and `createTheme()`.

- [ ] **Step 6: Extend codec integration coverage**

Change `featureFixture()` theme root from the placeholder namespace `xmlns:a="a"` to the real DrawingML namespace, then add a direct `rId3` presentation theme relationship. Assert:

```ts
expect(codec.presentationTheme).toBe(theme);
expect(codec.presentationTheme).toBe(codec.presentationTheme);
theme.setFonts({ minorLatin: 'Noto Sans' });
expect(theme.fonts).toEqual({ majorLatin: 'Aptos Display', minorLatin: 'Noto Sans' });
```

Add separate packages for no direct relationship, master-only relationship, duplicate relationship, external relationship, dangling target, wrong content type, alternate theme URI, and an unrelated extra theme part. Only the unique valid direct relationship may resolve.

- [ ] **Step 7: Run focused, package, and type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/codecs/src/theme-fonts.internal.test.ts \
  packages/codecs/src/codecs.test.ts --reporter=dot
node node_modules/vitest/vitest.mjs run packages/codecs/src --reporter=dot
node node_modules/typescript/bin/tsc -b packages/codecs --pretty false
git diff --check
```

- [ ] **Step 8: Review, commit, push, and verify**

Review expanded-name lookup, direct-child cardinality, missing-attribute insertion, input descriptor safety, unknown content preservation, no-op journal behavior, stable identity, and transaction rollback. Then:

```sh
git add -- packages/codecs/src/theme-fonts.internal.ts \
  packages/codecs/src/theme-fonts.internal.test.ts \
  packages/codecs/src/master.ts packages/codecs/src/codecs.test.ts
git diff --cached --check
git commit -m "feat: add presentation theme font codec"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 2: Expose the live PptxGenJS-style document theme property

**Files:**
- Create: `packages/sdk/src/presentation-theme.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Produces `PresentationTheme { headFontFace: string; bodyFontFace: string }`.
- Produces `PresentationThemeOptions { headFontFace?: string; bodyFontFace?: string }`.
- Produces `normalizePresentationTheme(value): PresentationTheme`.
- Produces live `PptxDocument.theme: PresentationTheme | undefined` with whole-replacement setter semantics.

- [ ] **Step 1: Write the failing live-property lifecycle test**

Open a canonical created package and assert:

```ts
const document = PptxDocument.create();
const themeModel = document.masterLayoutTheme.presentationTheme;
expect(document.theme).toEqual({
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
});
expect(document.theme).not.toBe(document.theme);
expect(document.masterLayoutTheme.presentationTheme).toBe(themeModel);

document.theme = {
  headFontFace: 'Noto Sans Display',
  bodyFontFace: 'Noto Sans',
};
expect(document.theme).toEqual({
  headFontFace: 'Noto Sans Display',
  bodyFontFace: 'Noto Sans',
});
```

Cover head-only → custom/`Calibri`, body-only → `Calibri Light`/custom, `{}` → both fallbacks, same-value exact no-op, XML escaping, invalid zero mutation, outer rollback, write/reopen, unrelated theme content preservation, missing/duplicate/external/dangling/wrong-content-type relationship setter rejection, and getter `undefined` for those unsafe states.

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts --reporter=dot
```

Expected: FAIL because `PptxDocument.theme` does not exist.

- [ ] **Step 3: Implement the SDK theme value normalizer**

Create:

```ts
export interface PresentationTheme {
  readonly headFontFace: string;
  readonly bodyFontFace: string;
}

export interface PresentationThemeOptions {
  readonly headFontFace?: string;
  readonly bodyFontFace?: string;
}

const DEFAULT_HEAD_FONT = 'Calibri Light';
const DEFAULT_BODY_FONT = 'Calibri';

export function normalizePresentationTheme(value: unknown): PresentationTheme {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Presentation theme must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Presentation theme must be an ordinary object');
  }
  const allowed = new Set(['headFontFace', 'bodyFontFace']);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`Unsupported presentation theme property: ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Presentation theme ${key} must be a data property`);
    }
  }
  const read = (key: 'headFontFace' | 'bodyFontFace', fallback: string): string => {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    return descriptor?.value === undefined
      ? fallback
      : normalizePresentationFontName(descriptor.value, key);
  };
  return {
    headFontFace: read('headFontFace', DEFAULT_HEAD_FONT),
    bodyFontFace: read('bodyFontFace', DEFAULT_BODY_FONT),
  };
}

function normalizePresentationFontName(value: unknown, key: string): string {
  if (typeof value !== 'string' || !/\S/u.test(value)) {
    throw new TypeError(`Presentation theme ${key} must be a non-whitespace string`);
  }
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
    throw new TypeError(`Presentation theme ${key} contains invalid XML characters`);
  }
  return value;
}
```

Keep this SDK normalizer private. The five-line font predicate intentionally mirrors the codec boundary so the SDK can reject before package construction without making an internal codec normalizer part of the public aggregate API.

- [ ] **Step 4: Add the live getter and setter**

Export the two public types from `index.ts`, then add:

```ts
get theme(): PresentationTheme | undefined {
  const fonts = this.#masterLayoutTheme.presentationTheme?.fonts;
  return fonts === undefined
    ? undefined
    : { headFontFace: fonts.majorLatin, bodyFontFace: fonts.minorLatin };
}

set theme(value: PresentationThemeOptions) {
  const normalized = normalizePresentationTheme(value);
  const theme = this.#masterLayoutTheme.presentationTheme;
  if (!theme) throw new Error('Presentation does not have one editable direct theme');
  theme.setFonts({
    majorLatin: normalized.headFontFace,
    minorLatin: normalized.bodyFontFace,
  });
}
```

Normalization must finish before resolving or mutating the package.

- [ ] **Step 5: Run focused, SDK, declaration, and full type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/codecs/src/theme-fonts.internal.test.ts \
  packages/codecs/src/codecs.test.ts \
  packages/sdk/src/index.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/sdk packages/pptx --pretty false
pnpm typecheck
git diff --check
```

- [ ] **Step 6: Review, commit, push, and verify**

Review detached snapshots, whole-replacement fallbacks, stable codec identity, pre-mutation normalization, exact no-op, rollback, unsafe relationship handling, and public declaration inference. Then:

```sh
git add -- packages/sdk/src/presentation-theme.ts \
  packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: expose presentation theme fonts"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 3: Add explicit native theme-font creation

**Files:**
- Modify: `packages/sdk/src/create.ts`
- Modify: `packages/sdk/src/index.ts`
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Extends `CreatePresentationOptions` with `readonly theme?: PresentationThemeOptions`.
- Routes only an explicitly present, non-`undefined` theme object through the live whole-replacement property.
- Keeps zero-input and runtime-`undefined` creation byte-identical.

- [ ] **Step 1: Write failing creation tests**

Require:

```ts
expect(PptxDocument.create().theme).toEqual({
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
});
expect(PptxDocument.create({ theme: {} }).theme).toEqual({
  headFontFace: 'Calibri Light',
  bodyFontFace: 'Calibri',
});
expect(PptxDocument.create({
  theme: { headFontFace: 'Noto Sans Display' },
}).theme).toEqual({
  headFontFace: 'Noto Sans Display',
  bodyFontFace: 'Calibri',
});
```

Cover body-only, both, runtime `undefined`, frozen input, null-prototype input, caller mutation after create, custom XML escaping, invalid types, empty/whitespace, accessor/symbol/unknown keys, arrays, custom prototype, and zero-input package-byte regression.

- [ ] **Step 2: Run RED**

```sh
node node_modules/vitest/vitest.mjs run packages/sdk/src/index.test.ts --reporter=dot
```

Expected: FAIL because `CreatePresentationOptions.theme` is absent and `PptxDocument.create()` ignores it.

- [ ] **Step 3: Extend the creation type and route explicit input**

In `create.ts`:

```ts
import type { PresentationThemeOptions } from './presentation-theme.js';

export interface CreatePresentationOptions {
  readonly author?: string;
  readonly company?: string;
  readonly createdAt?: string;
  readonly format?: PresentationFormat;
  readonly lastModifiedBy?: string;
  readonly modifiedAt?: string;
  readonly revision?: string;
  readonly rtlMode?: boolean;
  readonly slideSize?: BuiltInSlideSize | CustomSlideSize;
  readonly subject?: string;
  readonly theme?: PresentationThemeOptions;
  readonly title?: string;
}
```

In `PptxDocument.create()`, normalize before package construction and apply only when supplied:

```ts
const theme = options.theme === undefined
  ? undefined
  : normalizePresentationTheme(options.theme);
const document = new PptxDocument(createPresentationPackage(options));
if (theme !== undefined) document.theme = theme;
```

Place this before metadata setters so invalid theme input cannot mutate even the locally created package. Do not pass theme into `createPresentationPackage()` or rewrite `THEME_XML`; the canonical template remains the one source for omitted creation.

- [ ] **Step 4: Run focused, create regression, and package type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/sdk/src/index.test.ts \
  packages/sdk/src/browser.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/sdk packages/pptx --pretty false
git diff --check
```

- [ ] **Step 5: Review, commit, push, and verify**

Review explicit-vs-omitted presence, PptxGenJS fallback mapping, canonical default bytes, validation timing, detached inputs, create transaction behavior, and browser-safe imports. Then:

```sh
git add -- packages/sdk/src/create.ts packages/sdk/src/index.ts packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "feat: create presentation theme fonts"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 4: Compare public PptxGenJS theme output

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Proves public PptxGenJS 4.0.1 default, empty, head-only, body-only, and full custom theme output imports through `document.theme`.
- Proves native explicit creation matches final supported typeface semantics.

- [ ] **Step 1: Extend the public PptxGenJS test interface**

Add only:

```ts
theme: {
  readonly headFontFace?: string;
  readonly bodyFontFace?: string;
} | undefined;
```

Do not add `_theme`, `_slides`, or any private member.

- [ ] **Step 2: Add the conformance matrix**

For default, `{}`, `{ headFontFace: 'Aptos Display' }`, `{ bodyFontFace: 'Aptos' }`, and both `Noto Sans` fields, create a fresh PptxGenJS instance, optionally assign `theme`, add one slide, and import with `importPptxGenJS()`.

Require exact snapshots:

```ts
expect(importedDefault.theme).toEqual({
  headFontFace: 'Calibri Light',
  bodyFontFace: 'Calibri',
});
expect(importedHeadOnly.theme).toEqual({
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Calibri',
});
expect(importedCustom.theme).toEqual({
  headFontFace: 'Noto Sans Display',
  bodyFontFace: 'Noto Sans',
});
```

Build matching native explicit documents, compare the two Latin typeface snapshots, mutate one imported theme through `ThemeModel.setFonts()`, then write/reopen and require preservation. Assert adapter import performs no mutation before the explicit edit.

- [ ] **Step 3: Run focused adapter and type gates**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/pptxgenjs-adapter/src/index.test.ts --reporter=dot
node node_modules/typescript/bin/tsc -b packages/pptxgenjs-adapter packages/pptx --pretty false
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

Review public-API-only usage, runtime version assertion, fallback evidence, native/imported snapshot comparison, no private dependency, and reopen preservation. Then:

```sh
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare presentation theme fonts with pptxgenjs"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 5: Prove packed Node, browser, and declaration surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Proves packed native create/read/replace/partial-edit/write/reopen in Node and browser bundles.
- Proves public `PresentationTheme`, `PresentationThemeOptions`, `ThemeFontSnapshot`, and `ThemeFontUpdate` declarations.
- Adds `presentationThemeFonts: true` to the smoke JSON.

- [ ] **Step 1: Extend Node packed runtime smoke**

Create with both custom fonts, capture `createdTheme`, assign a head-only theme and require the body fallback, then patch the live model's minor Latin, write/reopen, and require:

```ts
const themed = PptxDocument.create({
  theme: { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' },
});
const createdTheme = themed.theme;
themed.theme = { headFontFace: 'Noto Sans Display' };
const replacedTheme = themed.theme;
themed.masterLayoutTheme.presentationTheme.setFonts({ minorLatin: 'Noto Sans' });
const reopenedTheme = (await PptxDocument.open(await themed.write())).theme;
```

Set `presentationThemeFonts` true only when all three snapshots have the exact expected values.

- [ ] **Step 2: Extend browser runtime smoke**

Repeat create, whole replacement, model patch, `writeBlob()`, reopen, and detached snapshot checks through the browser bundle. Do not import Node modules into the generated browser script.

- [ ] **Step 3: Extend declaration compile smoke**

Import the four public types and add:

```ts
const themeOptions: PresentationThemeOptions = {
  headFontFace: 'Aptos Display',
};
const themedDocument = PptxDocument.create({ theme: themeOptions });
const themeSnapshot: PresentationTheme | undefined = themedDocument.theme;
const fontSnapshot: ThemeFontSnapshot | undefined =
  themedDocument.masterLayoutTheme.presentationTheme?.fonts;
const fontUpdate: ThemeFontUpdate = { minorLatin: 'Aptos' };
themedDocument.masterLayoutTheme.presentationTheme?.setFonts(fontUpdate);
```

- [ ] **Step 4: Run the real tarball smoke and full typecheck**

```sh
pnpm typecheck
pnpm --filter @jiayunxie/pptx build
pptx_theme_pack_dir=$(mktemp -d /tmp/pptx-theme-fonts-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_theme_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_theme_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Expected JSON contains `"presentationThemeFonts":true`, `"types":true`, and `"cli":"0.1.0"`.

- [ ] **Step 5: Review, commit, push, and verify**

Review that smoke imports the actual tarball, exercises Node/browser/types, tests both API layers, and has no workspace-source dependency. Then:

```sh
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed presentation theme fonts"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 6: Document theme-font creation, editing, and parity boundaries

**Files:**
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Marks PptxGenJS presentation theme fonts supported.
- Documents native default, explicit partial fallback, strict input, relationship ownership, whole vs partial editing, and adapter behavior.

- [ ] **Step 1: Update the parity matrix**

Add this row beside the other presentation-level properties:

```md
| presentation `pptx.theme.headFontFace/bodyFontFace` | `CreatePresentationOptions.theme` / `document.theme` / `ThemeModel.setFonts()` | 已支持 native strict 创建、presentation-direct theme 读取/整组编辑、具体 theme partial 编辑；zero-input native 保留 Aptos，explicit partial 使用 PptxGenJS Calibri fallback |
```

Replace the pending-theme sentence after company with a paragraph covering public PptxGenJS default/partial/custom output, native intentional zero-input divergence, strict invalid handling, presentation relationship lookup, non-Latin font preservation, and adapter final-state import. Keep custom properties, sections, masters, and placeholders listed as pending.

- [ ] **Step 2: Update package README**

Add one concise API contract paragraph after presentation metadata and before table details. Include one create example and one `ThemeModel.setFonts()` edit example. Explicitly state that assignment to `document.theme` is whole replacement while `setFonts()` is partial.

- [ ] **Step 3: Run docs and type gates**

```sh
rg -n 'theme|headFontFace|bodyFontFace|ThemeModel.setFonts' \
  docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
pnpm typecheck
git diff --check
```

- [ ] **Step 4: Review, commit, push, and verify**

Review that docs do not overclaim theme colors, non-Latin editing, multiple-theme selection, master definition, or byte parity with PptxGenJS's entire theme template. Then:

```sh
git add -- docs/compatibility/pptxgenjs-baseline.md packages/pptx/README.md
git diff --cached --check
git commit -m "docs: document presentation theme fonts"
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected divergence: `0 0`.

---

### Task 7: Run full release and real-deck QA

**Files:**
- No repository changes expected.
- Generate artifacts only under `/tmp/pptx-theme-fonts-qa-20260730`.

**Interfaces:**
- Proves full repository compatibility, performance, package semantics, validation, mutation isolation, render stability, and no hidden relationship fallback.

- [ ] **Step 1: Run complete repository gates**

```sh
pnpm typecheck
pnpm test
pnpm test:performance
pnpm --filter @jiayunxie/pptx build
pptx_theme_qa_pack_dir=$(mktemp -d /tmp/pptx-theme-fonts-qa-pack.XXXXXX)
pnpm --filter @jiayunxie/pptx pack --pack-destination "$pptx_theme_qa_pack_dir"
node scripts/smoke-npm-package.mjs \
  "$pptx_theme_qa_pack_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require all tests, the 5-second performance budget, packed Node/browser/types smoke, and CLI smoke to pass.

- [ ] **Step 2: Generate representative decks**

Using built package public APIs, generate:

```text
native-default.pptx
native-custom.pptx
native-partial.pptx
native-edited.pptx
native-reopened.pptx
native-second-write.pptx
pptxgenjs-default.pptx
pptxgenjs-custom.pptx
```

Each deck has one slide with identical heading/body text whose run fonts are `+mj-lt` and `+mn-lt`, so rendering resolves through the edited theme instead of an explicit face name. Use a deterministic locally available custom heading/body pair for the native/PptxGenJS comparison. Keep generator source in `/tmp` only.

- [ ] **Step 3: Validate packages and exact theme state**

Run `pptx-inspect package validate --profile powerpoint-2010` on every deck and require zero errors/warnings. Run `pptx-inspect part read` on the presentation relationships and resolved theme part; assert exact major/minor typefaces and preservation of color/format/script content.

- [ ] **Step 4: Verify mutation isolation and save stability**

Run package diff custom→edited and require only the resolved theme part to change. Run edited→reopened and edited→second-write and require zero part changes. Confirm no presentation, master, layout, notes, slide, relationship, content-type, metadata, or media part changes.

- [ ] **Step 5: Render and inspect every representative slide**

Use the bundled LibreOffice, `pdfinfo`, and `pdftoppm` on native custom/edited/reopened and PptxGenJS custom. Require identical page counts and sizes. Configure a deterministic available font if necessary, compare PNG SHA-256 where content is equivalent, run Presentations `render_slides.py` and `slides_test.py`, and inspect every rendered slide individually for fallback, wrapping, clipping, and overflow.

- [ ] **Step 6: Final repository and remote audit**

```sh
git diff --check
git diff --cached --check
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git status --short --branch
git rev-list --left-right --count origin/main...HEAD
```

Expected: only user-owned `?? .pnpm-store/`, branch `main...origin/main`, and divergence `0 0`. Do not create an empty QA commit.
