import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative, resolve, sep } from 'node:path';
import test from 'node:test';

import ts from 'typescript';

import {
  extractPptxGenJSPublicSurface,
  resolvePptxGenJSPackage,
} from './pptxgenjs-surface-declarations.mjs';

const FIXTURE = `
export default Deck;
declare class Deck {
  readonly mode: Deck.Mode;
  addSlide(options?: Deck.AddSlideOptions): Deck.Slide;
  /** @deprecated use the options overload */
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

const PERMUTED_FIXTURE = `
declare namespace Deck {
  export class Slide {
    addText(text: string, options?: AddSlideOptions): Slide;
    hidden: boolean;
  }
  export interface AddSlideOptions extends BaseOptions {
    nested?: { next?: AddSlideOptions; label?: string };
    align?: Align;
  }
  export interface BaseOptions { readonly x?: number }
  export type Align = 'right' | 'left';
  export enum Mode { standard = 'standard', wide = 'wide' }
}
declare class Deck {
  addSlide(masterName?: string): Deck.Slide;
  readonly mode: Deck.Mode;
  addSlide(options?: Deck.AddSlideOptions): Deck.Slide;
}
export default Deck;
`;

function extract(sourceText = FIXTURE) {
  return extractPptxGenJSPublicSurface({
    sourceText,
    fileName: 'types/index.d.ts',
    typescript: ts,
  });
}

function assertDeepFrozen(value, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  assert.equal(Object.isFrozen(value), true);
  for (const nested of Object.values(value)) assertDeepFrozen(nested, seen);
}

test('extracts stable reachable public atoms with overloads and inheritance', () => {
  const result = extract();
  const ids = result.atoms.map(({ id }) => id);

  assert.deepEqual(result.roots, { presentation: 'Deck', slide: 'Slide' });
  assert.deepEqual(ids, [...ids].sort());
  assert.deepEqual(ids, [
    'class:Deck#addSlide',
    'class:Deck@property:mode',
    'inline:interface:AddSlideOptions@property:nested@property:nested.label',
    'inline:interface:AddSlideOptions@property:nested@property:nested.next',
    'interface:AddSlideOptions@property:align',
    'interface:AddSlideOptions@property:nested',
    'interface:AddSlideOptions@property:x',
    'method:Slide#addText',
    'property:Slide#hidden',
    'union:Align#left',
    'union:Align#right',
    'union:Mode#standard',
    'union:Mode#wide',
  ]);

  const overload = result.atoms.find(({ id }) => id === 'class:Deck#addSlide');
  assert.deepEqual(overload.signatures, [
    'addSlide(masterName?: string): Deck.Slide',
    'addSlide(options?: Deck.AddSlideOptions): Deck.Slide',
  ]);
  assert.equal(overload.deprecated, false);
  assert.deepEqual(overload.deprecatedSignatures, [
    'addSlide(masterName?: string): Deck.Slide',
  ]);

  const inherited = result.atoms.find(
    ({ id }) => id === 'interface:AddSlideOptions@property:x',
  );
  assert.equal(inherited.declaredIn, 'BaseOptions');
  assert.equal(inherited.optional, true);
  assert.equal(inherited.readonly, true);
  assert.equal(ids.includes('interface:BaseOptions@property:x'), false);
  assert.deepEqual(result.diagnostics, []);
  assertDeepFrozen(result);
});

test('keeps atom IDs stable when declarations and union members are reordered', () => {
  assert.deepEqual(
    extract(PERMUTED_FIXTURE).atoms.map(({ id }) => id),
    extract(FIXTURE).atoms.map(({ id }) => id),
  );
});

test('terminates recursive references and ignores ambient container types', () => {
  const result = extract(`
    export default Deck;
    declare class Deck {
      create(options: Deck.Options): Promise<Deck.Slide[]>;
    }
    declare namespace Deck {
      export interface Options {
        children?: ReadonlyArray<Options>;
        callback?: Function;
      }
      export class Slide { hidden: boolean }
    }
  `);
  assert.deepEqual(result.atoms.map(({ id }) => id), [
    'class:Deck#create',
    'interface:Options@property:callback',
    'interface:Options@property:children',
    'property:Slide#hidden',
  ]);
});

test('emits primitive, tuple, and template-literal union alternatives', () => {
  const result = extract(`
    export default Deck;
    declare class Deck { setPosition(value: Deck.Coord, margin: Deck.Margin): Deck.Slide }
    declare namespace Deck {
      export type Coord = number | \`\${number}%\`;
      export type Margin = number | [number, number, number, number];
      export class Slide { hidden: boolean }
    }
  `);
  assert.deepEqual(result.atoms.map(({ id }) => id), [
    'class:Deck#setPosition',
    'property:Slide#hidden',
    'union:Coord#${number}%',
    'union:Coord#number',
    'union:Margin#[number,number,number,number]',
    'union:Margin#number',
  ]);
});

test('emits direct option union tokens under their owning property atom', () => {
  const result = extract(`
    export default Deck;
    declare class Deck { configure(options: Deck.Options): Deck.Slide }
    declare namespace Deck {
      export interface Options {
        mode?: 'fast' | 'safe';
        value?: string | false;
      }
      export class Slide { hidden: boolean }
    }
  `);
  assert.deepEqual(result.atoms.map(({ id }) => id), [
    'class:Deck#configure',
    'interface:Options@property:mode',
    'interface:Options@property:value',
    'property:Slide#hidden',
    'union:interface:Options@property:mode#fast',
    'union:interface:Options@property:mode#safe',
    'union:interface:Options@property:value#false',
    'union:interface:Options@property:value#string',
  ]);
});

test('keeps method and nested union tokens separated by their declaration path', () => {
  const result = extract(`
    export default Deck;
    declare class Deck {
      configure(
        mode: 'fast' | 'safe',
        fallback: 'fast' | 'other',
        options?: Deck.Options,
      ): Deck.Slide;
    }
    declare namespace Deck {
      export interface Options {
        nested?: { mode?: 'fast' | 'safe'; fallback?: 'fast' | 'other' };
      }
      export class Slide { hidden: boolean }
    }
  `);
  const ids = result.atoms.map(({ id }) => id);
  assert.equal(ids.includes('union:class:Deck#configure@path:mode#fast'), true);
  assert.equal(ids.includes('union:class:Deck#configure@path:fallback#fast'), true);
  assert.equal(ids.includes('union:class:Deck#configure@path:mode#safe'), true);
  assert.equal(ids.includes('union:class:Deck#configure@path:fallback#other'), true);
  assert.equal(
    ids.includes('union:interface:Options@property:nested@path:nested.mode#fast'),
    true,
  );
  assert.equal(
    ids.includes('union:interface:Options@property:nested@path:nested.fallback#fast'),
    true,
  );
});

test('rejects malformed, missing-root, unresolved-local, and conflicting declarations', () => {
  assert.throws(
    () => extract('export default Deck; declare class Deck {'),
    /parse diagnostics/u,
  );
  assert.throws(
    () => extract('declare class Deck {}'),
    /default export/u,
  );
  assert.throws(
    () => extract(`
      export default Deck;
      declare class Deck { create(options: Deck.Missing): Deck.Slide }
      declare namespace Deck { export class Slide {} }
    `),
    /unresolved local type Deck\.Missing/u,
  );
  assert.throws(
    () => extract(`
      export default Deck;
      declare class Deck { value: string; value(): string }
      declare namespace Deck { export class Slide {} }
    `),
    /conflicting atom/u,
  );
});

test('does not turn comments or private members into public atoms', () => {
  const result = extract(`
    export default Deck;
    declare class Deck {
      // secret: string
      private hiddenState: string;
      public title: string;
    }
    declare namespace Deck {
      export class Slide { protected internal: string; hidden: boolean }
    }
  `);
  assert.deepEqual(result.atoms.map(({ id }) => id), [
    'class:Deck@property:title',
    'property:Slide#hidden',
  ]);
});

test('resolves the locked real package without depending on its install layout', async () => {
  const packageInfo = await resolvePptxGenJSPackage(
    resolve('packages/pptxgenjs-adapter/package.json'),
  );
  assert.equal(packageInfo.version, '4.0.1');
  assert.match(packageInfo.declarationSha256, /^[a-f0-9]{64}$/u);
  assert.equal(packageInfo.entryPath.startsWith(`${packageInfo.root}${sep}`), true);
  assert.equal(
    packageInfo.declarationPath.startsWith(`${packageInfo.root}${sep}`),
    true,
  );
  assert.equal(relative(packageInfo.root, packageInfo.declarationPath), join('types', 'index.d.ts'));
});

test('locks the real 4.0.1 declaration hash and reachable atom inventory', async () => {
  const packageInfo = await resolvePptxGenJSPackage(
    resolve('packages/pptxgenjs-adapter/package.json'),
  );
  const sourceText = await readFile(packageInfo.declarationPath, 'utf8');
  const first = extractPptxGenJSPublicSurface({
    sourceText,
    fileName: 'types/index.d.ts',
    typescript: ts,
  });
  const second = extractPptxGenJSPublicSurface({
    sourceText,
    fileName: 'types/index.d.ts',
    typescript: ts,
  });

  assert.equal(
    packageInfo.declarationSha256,
    '0726d015dbcb55ccfa75546cb2fd43fe13a0dfeb783d08572f1c62f59193bbe5',
  );
  assert.equal(first.atoms.length, 1_774);
  assert.deepEqual(first, second);
  const ids = new Set(first.atoms.map(({ id }) => id));
  for (const id of [
    'class:PptxGenJS@property:theme',
    'method:Slide#addMedia',
    'union:Coord#${number}%',
    'union:interface:BorderProps@property:type#none',
  ]) {
    assert.equal(ids.has(id), true, id);
  }
  assert.equal(
    first.atoms.find(({ id }) => id === 'union:SchemeColor#tx1')?.catalogKey,
    'text1',
  );
  assert.equal(
    first.atoms.find(({ id }) => id === 'union:ChartType#bar3D')?.catalogKey,
    'bar3d',
  );
});

test('rejects a resolved PptxGenJS package with a different version', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-surface-resolver-'));
  try {
    const adapterDirectory = join(directory, 'adapter');
    const packageDirectory = join(directory, 'node_modules', 'pptxgenjs');
    await mkdir(adapterDirectory, { recursive: true });
    await mkdir(join(packageDirectory, 'dist'), { recursive: true });
    await mkdir(join(packageDirectory, 'types'), { recursive: true });
    await writeFile(
      join(adapterDirectory, 'package.json'),
      JSON.stringify({ name: 'fixture-adapter', type: 'module' }),
    );
    await writeFile(
      join(packageDirectory, 'package.json'),
      JSON.stringify({
        name: 'pptxgenjs',
        version: '4.1.0',
        main: 'dist/index.cjs',
        types: 'types',
      }),
    );
    await writeFile(join(packageDirectory, 'dist', 'index.cjs'), 'module.exports = class Deck {};');
    await writeFile(join(packageDirectory, 'types', 'index.d.ts'), 'export default class Deck {}');

    await assert.rejects(
      resolvePptxGenJSPackage(join(adapterDirectory, 'package.json')),
      /expected pptxgenjs@4\.0\.1 but resolved 4\.1\.0/u,
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
