import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { resolve, sep } from 'node:path';
import test from 'node:test';

import ts from 'typescript';

import {
  extractPptxGenJSPublicSurface,
  resolvePptxGenJSPackage,
} from './pptxgenjs-surface-declarations.mjs';
import {
  hashRuntimeProbe,
  probePptxGenJSRuntime,
} from './pptxgenjs-runtime-probe.mjs';

let inputPromise;

async function loadInputs() {
  inputPromise ??= (async () => {
    const packageInfo = await resolvePptxGenJSPackage(
      resolve('packages/pptxgenjs-adapter/package.json'),
    );
    const sourceText = await readFile(packageInfo.declarationPath, 'utf8');
    const surface = extractPptxGenJSPublicSurface({
      sourceText,
      fileName: 'types/index.d.ts',
      typescript: ts,
    });
    return { packageInfo, surface };
  })();
  return inputPromise;
}

function assertDeepFrozen(value, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  assert.equal(Object.isFrozen(value), true);
  for (const nested of Object.values(value)) assertDeepFrozen(nested, seen);
}

function walk(value, visit, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) {
    visit(value);
    return;
  }
  seen.add(value);
  for (const [key, nested] of Object.entries(value)) {
    visit(key);
    walk(nested, visit, seen);
  }
}

test('probes the locked public runtime and declared catalogs deterministically', async () => {
  const input = await loadInputs();
  const first = await probePptxGenJSRuntime(input);
  const second = await probePptxGenJSRuntime(input);

  assert.deepEqual(first, second);
  assert.equal(first.packageVersion, '4.0.1');
  assert.equal(first.declarationSha256, input.packageInfo.declarationSha256);
  assert.equal(
    first.runtimeEntrySha256,
    '873d182a8e2e1c0b5e522ef146117936b96b9b2024667bd4c1de59e2b031d27a',
  );
  assert.deepEqual(first.classMembers, {
    methods: [
      'addSection',
      'addSlide',
      'defineLayout',
      'defineSlideMaster',
      'stream',
      'tableToSlides',
      'write',
      'writeFile',
    ],
    properties: [
      'AlignH',
      'AlignV',
      'ChartType',
      'OutputType',
      'PlaceholderType',
      'SchemeColor',
      'ShapeType',
      'author',
      'company',
      'layout',
      'presLayout',
      'revision',
      'rtlMode',
      'subject',
      'theme',
      'title',
      'version',
    ],
  });
  assert.deepEqual(first.slideMembers, {
    methods: [
      'addChart',
      'addImage',
      'addMedia',
      'addNotes',
      'addShape',
      'addTable',
      'addText',
    ],
    properties: [
      'background',
      'bkgd',
      'color',
      'hidden',
      'newAutoPagedSlides',
      'slideNumber',
    ],
  });
  assert.deepEqual(Object.values(first.catalogs.OutputType), [
    'arraybuffer',
    'base64',
    'binarystring',
    'blob',
    'nodebuffer',
    'uint8array',
  ]);
  assert.equal(first.catalogs.PlaceholderType, null);
  assert.deepEqual(first.runtimeMismatches, [
    {
      kind: 'catalog-extra',
      owner: 'ShapeType',
      name: 'custGeom',
      declarationName: 'ShapeType',
      actual: 'custGeom',
      expected: null,
    },
    {
      kind: 'missing-property',
      owner: 'PptxGenJS',
      name: 'PlaceholderType',
    },
  ]);
  assert.deepEqual(first.minimalCalls, {
    addSection: 'undefined',
    addSlide: 'Slide',
    addNotesReturnsSlide: true,
  });
  assert.equal(hashRuntimeProbe(first), hashRuntimeProbe(second));
  assert.equal(
    hashRuntimeProbe(first),
    'fe342796785d4b14e88757a0a28b56cb1cd76457de83e67b241a6a3d3bb06b64',
  );
  walk(first, (value) => {
    if (typeof value !== 'string') return;
    assert.equal(value.startsWith('_'), false, value);
    assert.equal(value.includes(`${resolve('.')}${sep}`), false, value);
  });
  assertDeepFrozen(first);
});

test('rejects a declared presentation method missing from the runtime', async () => {
  const input = await loadInputs();
  const packageRequire = createRequire(input.packageInfo.packageJsonPath);
  const Runtime = packageRequire(input.packageInfo.entryPath);
  class MissingWrite {
    constructor() {
      const instance = new Runtime();
      return new Proxy(instance, {
        get(target, property, receiver) {
          if (property === 'write') return undefined;
          return Reflect.get(target, property, receiver);
        },
        has(target, property) {
          return property === 'write' ? false : Reflect.has(target, property);
        },
      });
    }
  }

  await assert.rejects(
    probePptxGenJSRuntime({
      ...input,
      loadRuntime: () => MissingWrite,
    }),
    /missing declared method PptxGenJS\.write/u,
  );
});

test('rejects a runtime catalog that differs from its declared enum', async () => {
  const input = await loadInputs();
  const packageRequire = createRequire(input.packageInfo.packageJsonPath);
  const Runtime = packageRequire(input.packageInfo.entryPath);
  class BrokenOutputCatalog {
    constructor() {
      const instance = new Runtime();
      return new Proxy(instance, {
        get(target, property, receiver) {
          if (property === 'OutputType') {
            return { ...Reflect.get(target, property, receiver), extra: 'extra' };
          }
          return Reflect.get(target, property, receiver);
        },
      });
    }
  }

  await assert.rejects(
    probePptxGenJSRuntime({
      ...input,
      loadRuntime: () => BrokenOutputCatalog,
    }),
    /catalog OutputType differs from declaration OutputType/u,
  );
});

test('rejects disappearance of either locked upstream runtime mismatch', async () => {
  const input = await loadInputs();
  const packageRequire = createRequire(input.packageInfo.packageJsonPath);
  const Runtime = packageRequire(input.packageInfo.entryPath);
  const placeholderCatalog = Object.fromEntries(
    input.surface.atoms
      .filter(({ kind, owner }) => kind === 'union-member' && owner === 'PLACEHOLDER_TYPES')
      .map(({ catalogKey, name }) => [catalogKey, name]),
  );
  class FixedPlaceholderCatalog {
    constructor() {
      const instance = new Runtime();
      return new Proxy(instance, {
        get(target, property, receiver) {
          return property === 'PlaceholderType'
            ? placeholderCatalog
            : Reflect.get(target, property, receiver);
        },
        has(target, property) {
          return property === 'PlaceholderType' || Reflect.has(target, property);
        },
      });
    }
  }
  class FixedShapeCatalog {
    constructor() {
      const instance = new Runtime();
      return new Proxy(instance, {
        get(target, property, receiver) {
          if (property !== 'ShapeType') return Reflect.get(target, property, receiver);
          const { custGeom: omitted, ...declared } = Reflect.get(target, property, receiver);
          assert.equal(omitted, 'custGeom');
          return declared;
        },
      });
    }
  }

  await assert.rejects(
    probePptxGenJSRuntime({ ...input, loadRuntime: () => FixedPlaceholderCatalog }),
    /expected runtime mismatch missing-property:PptxGenJS\.PlaceholderType was not observed/u,
  );
  await assert.rejects(
    probePptxGenJSRuntime({ ...input, loadRuntime: () => FixedShapeCatalog }),
    /expected runtime mismatch catalog-extra:ShapeType\.custGeom=custGeom was not observed/u,
  );
});

test('rejects package version and entry-root drift before loading runtime code', async () => {
  const input = await loadInputs();
  await assert.rejects(
    probePptxGenJSRuntime({
      ...input,
      packageInfo: { ...input.packageInfo, version: '4.0.0' },
    }),
    /expected package version 4\.0\.1/u,
  );
  await assert.rejects(
    probePptxGenJSRuntime({
      ...input,
      packageInfo: {
        ...input.packageInfo,
        entryPath: resolve(input.packageInfo.root, '..', 'outside.cjs'),
      },
    }),
    /runtime entry is outside the package root/u,
  );
});
