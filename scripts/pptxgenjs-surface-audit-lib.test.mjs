import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';

import { buildPptxGenJSAudit } from './pptxgenjs-surface-audit-lib.mjs';
import { PPTXGENJS_SURFACE_MANIFEST } from './pptxgenjs-surface-manifest.mjs';

const IDS = Object.freeze([
  'class:Deck#write',
  'method:Slide#addText',
  'union:OutputType#blob',
  'interface:ImageProps@property:altText',
  'property:Slide#bkgd',
  'interface:ShapeProps@property:shadow',
]);

function atom(id) {
  return Object.freeze({
    id,
    kind: id.startsWith('union:') ? 'union-member' : 'property',
    owner: id.split(/[:#@]/u)[1] ?? 'Deck',
    name: id.split(/[:#@]/u).at(-1) ?? id,
    declaredIn: 'Fixture',
    optional: false,
    readonly: false,
    deprecated: false,
    signatures: [],
    deprecatedSignatures: [],
    catalogKey: '',
    typeText: 'unknown',
  });
}

function surface(ids = IDS.slice(0, 3)) {
  return Object.freeze({
    schemaVersion: 1,
    atoms: Object.freeze(ids.map(atom)),
    roots: Object.freeze({ presentation: 'Deck', slide: 'Slide' }),
    diagnostics: Object.freeze([]),
  });
}

const runtimeProbe = Object.freeze({
  schemaVersion: 1,
  packageVersion: '4.0.1',
  declarationSha256: 'a'.repeat(64),
  runtimeEntrySha256: 'b'.repeat(64),
  runtimeMismatches: Object.freeze([]),
});

function emptyEvidence() {
  return {
    code: [],
    tests: [],
    package: [],
    ooxml: [],
    clients: [],
  };
}

function manifest(entries = [], extensions = []) {
  return {
    schemaVersion: 1,
    packageVersion: '4.0.1',
    entries,
    extensions,
  };
}

function entry(id, status, overrides = {}) {
  return {
    id,
    status,
    native: [],
    evidence: emptyEvidence(),
    note: `${id} fixture`,
    ...overrides,
  };
}

function validEvidence() {
  return {
    code: [{ path: 'src/write.ts', pattern: 'export function write', commit: 'abc1234' }],
    tests: [{ path: 'tests/write.test.ts', title: 'writes blob', commit: 'abc1234' }],
    package: [{ path: 'scripts/smoke.mjs', pattern: 'packed write', commit: 'abc1234' }],
    ooxml: [{ path: 'docs/ooxml.md', pattern: 'ppt/slides/slide1.xml', commit: 'abc1234' }],
    clients: [{ path: 'docs/client.md', pattern: 'PowerPoint pass', commit: 'abc1234' }],
  };
}

const validControl = Object.freeze({
  path: 'controls/control.mjs',
  pattern: 'PptxGenJS control',
  commit: 'abc1234',
});

async function withRepository(run) {
  const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-audit-lib-'));
  try {
    for (const path of ['src', 'tests', 'scripts', 'docs', 'controls']) {
      await mkdir(join(directory, path), { recursive: true });
    }
    await writeFile(join(directory, 'src', 'write.ts'), 'export function write() {}\n');
    await writeFile(join(directory, 'tests', 'write.test.ts'), "test('writes blob', () => {});\n");
    await writeFile(join(directory, 'scripts', 'smoke.mjs'), 'packed write\n');
    await writeFile(join(directory, 'docs', 'ooxml.md'), 'ppt/slides/slide1.xml\n');
    await writeFile(join(directory, 'docs', 'client.md'), 'PowerPoint pass\n');
    await writeFile(join(directory, 'controls', 'control.mjs'), 'PptxGenJS control\n');
    return await run(directory);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

function assertDeepFrozen(value, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  assert.equal(Object.isFrozen(value), true);
  for (const nested of Object.values(value)) assertDeepFrozen(nested, seen);
}

test('exports an immutable evidence-backed initial manifest batch', () => {
  assert.equal(PPTXGENJS_SURFACE_MANIFEST.schemaVersion, 1);
  assert.equal(PPTXGENJS_SURFACE_MANIFEST.packageVersion, '4.0.1');
  assert.equal(PPTXGENJS_SURFACE_MANIFEST.entries.length, 554);
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries.map(({ status }) => status).sort(),
    [
      ...Array(4).fill('defect-excluded'),
      ...Array(415).fill('supported'),
      ...Array(60).fill('deliberate-difference'),
      ...Array(75).fill('deprecated-alias'),
    ].sort(),
  );
  const lineFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    /^(?:union:)?interface:(?:ShapeLineProps@property:(?:alpha|beginArrowType|color|dashType|endArrowType|lineDash|lineHead|lineTail|pt|size|transparency|type|width)|(?:ShapeProps|TextPropsOptions)@property:(?:line|lineDash|lineHead|lineSize|lineTail))(?:#.+)?$/u
      .test(id));
  assert.equal(lineFamilyEntries.length, 105);
  assert.deepEqual(
    lineFamilyEntries.map(({ status }) => status).sort(),
    [
      ...Array(31).fill('deliberate-difference'),
      ...Array(74).fill('deprecated-alias'),
    ].sort(),
  );
  const lineFamilyById = new Map(lineFamilyEntries.map((entry) => [entry.id, entry]));
  for (const entry of lineFamilyEntries.filter(
    ({ status }) => status === 'deprecated-alias',
  )) {
    assert.equal(lineFamilyById.get(entry.canonical)?.status, 'deliberate-difference');
  }
  assert.equal(
    lineFamilyById.get('union:interface:TextPropsOptions@property:lineHead#triangle')
      ?.canonical,
    'union:interface:ShapeLineProps@property:beginArrowType#triangle',
  );
  assert.equal(
    lineFamilyById.get('interface:ShapeProps@property:lineSize')?.canonical,
    'interface:ShapeLineProps@property:width',
  );
  const fillFamilyEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    /^(?:union:)?interface:(?:ShapeFillProps@property:(?:alpha|color|transparency|type)|(?:ShapeProps|TextPropsOptions)@property:fill)(?:#.+)?$/u.test(id));
  assert.equal(fillFamilyEntries.length, 8);
  assert.deepEqual(
    fillFamilyEntries.map(({ status }) => status).sort(),
    [
      ...Array(7).fill('deliberate-difference'),
      'deprecated-alias',
    ].sort(),
  );
  const fillFamilyById = new Map(fillFamilyEntries.map((entry) => [entry.id, entry]));
  assert.equal(
    fillFamilyById.get('interface:ShapeFillProps@property:alpha')?.canonical,
    'interface:ShapeFillProps@property:transparency',
  );
  const tableFillEntries = PPTXGENJS_SURFACE_MANIFEST.entries.filter(({ id }) =>
    /^interface:(?:TableCellProps|TableProps|TableToSlidesProps)@property:fill$/u.test(id));
  assert.deepEqual(
    tableFillEntries.map(({ id, status }) => ({ id, status })),
    [
      {
        id: 'interface:TableCellProps@property:fill',
        status: 'deliberate-difference',
      },
      {
        id: 'interface:TableProps@property:fill',
        status: 'deliberate-difference',
      },
      {
        id: 'interface:TableToSlidesProps@property:fill',
        status: 'defect-excluded',
      },
    ],
  );
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => id.endsWith('#folderCorner'))
      .map(({ id, status }) => ({ id, status })),
    [
      { id: 'union:ShapeType#folderCorner', status: 'defect-excluded' },
      { id: 'union:SHAPE_NAME#folderCorner', status: 'defect-excluded' },
    ],
  );
  assert.deepEqual(
    PPTXGENJS_SURFACE_MANIFEST.entries
      .filter(({ id }) => /^union:PLACEHOLDER_TYPES?#(?:pic|tbl)$/u.test(id))
      .map(({ id, status }) => ({ id, status })),
    [
      { id: 'union:PLACEHOLDER_TYPE#pic', status: 'deliberate-difference' },
      { id: 'union:PLACEHOLDER_TYPE#tbl', status: 'deliberate-difference' },
      { id: 'union:PLACEHOLDER_TYPES#pic', status: 'deliberate-difference' },
      { id: 'union:PLACEHOLDER_TYPES#tbl', status: 'deliberate-difference' },
    ],
  );
  assert.deepEqual(PPTXGENJS_SURFACE_MANIFEST.extensions, []);
  assertDeepFrozen(PPTXGENJS_SURFACE_MANIFEST);
});

test('defaults missing declaration entries to unverified without double-counting', async () => {
  await withRepository(async (repositoryRoot) => {
    const commits = [];
    const report = await buildPptxGenJSAudit({
      surface: surface(),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', {
          native: ['Deck.write'],
          evidence: validEvidence(),
          serialization: true,
          client: true,
        }),
      ]),
      repositoryRoot,
      gitCommitExists: async (commit) => {
        commits.push(commit);
        return commit === 'abc1234';
      },
    });

    assert.deepEqual(report.counts, {
      supported: 1,
      'deliberate-difference': 0,
      'deprecated-alias': 0,
      'defect-excluded': 0,
      unsupported: 0,
      unverified: 2,
      stale: 0,
    });
    assert.equal(report.declarationTotal, 3);
    assert.equal(report.complete, false);
    assert.deepEqual(report.incompleteIds, [IDS[1], IDS[2]]);
    assert.deepEqual(report.diagnostics, []);
    assert.deepEqual(commits, ['abc1234']);
    assertDeepFrozen(report);
  });
});

test('accepts every explicit non-stale status while open statuses still fail completion', async () => {
  await withRepository(async (repositoryRoot) => {
    const evidence = validEvidence();
    const report = await buildPptxGenJSAudit({
      surface: surface(IDS),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', { native: ['Deck.write'], evidence }),
        entry(IDS[1], 'deliberate-difference', {
          native: ['Slide.addText'], evidence, control: validControl,
        }),
        entry(IDS[2], 'deprecated-alias', {
          native: ['Deck.write'], evidence, canonical: IDS[0], control: validControl,
        }),
        entry(IDS[3], 'defect-excluded', {
          evidence, control: validControl,
        }),
        entry(IDS[4], 'unsupported', { control: validControl }),
        entry(IDS[5], 'unverified'),
      ]),
      repositoryRoot,
      gitCommitExists: async () => true,
    });

    assert.deepEqual(report.counts, {
      supported: 1,
      'deliberate-difference': 1,
      'deprecated-alias': 1,
      'defect-excluded': 1,
      unsupported: 1,
      unverified: 1,
      stale: 0,
    });
    assert.deepEqual(report.diagnostics, []);
    assert.deepEqual(report.incompleteIds, [IDS[4], IDS[5]].sort());
    assert.equal(report.complete, false);
  });
});

test('reports status-specific missing evidence with stable diagnostics', async () => {
  await withRepository(async (repositoryRoot) => {
    const report = await buildPptxGenJSAudit({
      surface: surface(IDS.slice(0, 5)),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', { serialization: true, client: true }),
        entry(IDS[1], 'deliberate-difference'),
        entry(IDS[2], 'deprecated-alias', { canonical: 'class:Deck#missing' }),
        entry(IDS[3], 'defect-excluded'),
        entry(IDS[4], 'unsupported', { note: '' }),
      ]),
      repositoryRoot,
      gitCommitExists: async () => true,
    });
    const codes = report.diagnostics.map(({ code }) => code);
    for (const code of [
      'invalid-canonical',
      'missing-client-evidence',
      'missing-code-evidence',
      'missing-control',
      'missing-native',
      'missing-note',
      'missing-ooxml-evidence',
      'missing-package-evidence',
      'missing-test-evidence',
    ]) {
      assert.equal(codes.includes(code), true, code);
    }
    assert.deepEqual(
      report.diagnostics,
      [...report.diagnostics].sort((left, right) => (
        left.id.localeCompare(right.id)
        || left.code.localeCompare(right.code)
        || left.message.localeCompare(right.message)
      )),
    );
    assert.equal(report.complete, false);
  });
});

test('rejects duplicate IDs, illegal status, unknown fields, accessors, and class data', async () => {
  await withRepository(async (repositoryRoot) => {
    const build = (candidate) => buildPptxGenJSAudit({
      surface: surface(),
      runtimeProbe,
      manifest: candidate,
      repositoryRoot,
      gitCommitExists: async () => true,
    });
    await assert.rejects(
      build(manifest([entry(IDS[0], 'unverified'), entry(IDS[0], 'unverified')])),
      /duplicate manifest entry/u,
    );
    await assert.rejects(
      build(manifest([entry(IDS[0], 'stale')])),
      /invalid status stale/u,
    );
    await assert.rejects(
      build({ ...manifest(), schemaVersion: 2 }),
      /manifest schemaVersion must be 1/u,
    );
    await assert.rejects(
      build({ ...manifest(), packageVersion: '4.1.0' }),
      /manifest packageVersion must be 4\.0\.1/u,
    );
    await assert.rejects(
      build(manifest([{ ...entry(IDS[0], 'unverified'), surprise: true }])),
      /unknown key surprise/u,
    );
    const accessorEntry = entry(IDS[0], 'unverified');
    Object.defineProperty(accessorEntry, 'note', { enumerable: true, get: () => 'unsafe' });
    await assert.rejects(build(manifest([accessorEntry])), /note must be a data property/u);
    const accessorLink = { path: 'src/write.ts', pattern: 'write' };
    Object.defineProperty(accessorLink, 'pattern', { enumerable: true, get: () => 'unsafe' });
    await assert.rejects(
      build(manifest([entry(IDS[0], 'unverified', {
        evidence: { ...emptyEvidence(), code: [accessorLink] },
      })])),
      /pattern must be a data property/u,
    );

    class ManifestData {
      schemaVersion = 1;
      packageVersion = '4.0.1';
      entries = [];
      extensions = [];
    }
    await assert.rejects(build(new ManifestData()), /manifest must be a plain data object/u);
  });
});

test('diagnoses invalid paths, missing literals, missing files, and commit objects', async () => {
  await withRepository(async (repositoryRoot) => {
    const evidence = validEvidence();
    evidence.code = [
      { path: 'src/missing.ts', pattern: 'missing', commit: 'deadbee' },
      { path: 'src/write.ts', pattern: 'absent literal' },
      { path: resolve(repositoryRoot, 'src', 'write.ts'), pattern: 'export function write' },
    ];
    evidence.tests = [{ path: 'tests/write.test.ts', title: 'absent title' }];
    evidence.package = [{ path: '../outside.mjs', pattern: 'outside' }];
    const report = await buildPptxGenJSAudit({
      surface: surface([IDS[0]]),
      runtimeProbe,
      manifest: manifest([
        entry(IDS[0], 'supported', {
          native: ['Deck.write'], evidence, serialization: true, client: true,
        }),
      ]),
      repositoryRoot,
      gitCommitExists: async () => false,
    });
    const codes = new Set(report.diagnostics.map(({ code }) => code));
    for (const code of [
      'evidence-commit-missing',
      'evidence-file-missing',
      'evidence-path-invalid',
      'evidence-pattern-missing',
      'evidence-title-missing',
    ]) {
      assert.equal(codes.has(code), true, code);
    }
    assert.deepEqual(report.incompleteIds, [IDS[0]]);
    assert.equal(report.complete, false);
  });
});

test('verifies real Git commit objects with the default checker', async () => {
  const evidence = emptyEvidence();
  evidence.code = [{
    path: 'scripts/pptxgenjs-surface-audit-lib.mjs',
    pattern: 'buildPptxGenJSAudit',
    commit: '0b6094e',
  }];
  evidence.tests = [{
    path: 'scripts/pptxgenjs-surface-audit-lib.test.mjs',
    title: 'verifies real Git commit objects',
    commit: '0b6094e',
  }];
  evidence.package = [{
    path: 'package.json',
    pattern: 'typecheck',
    commit: '0b6094e',
  }];
  const report = await buildPptxGenJSAudit({
    surface: surface([IDS[0]]),
    runtimeProbe,
    manifest: manifest([
      entry(IDS[0], 'supported', { native: ['Deck.write'], evidence }),
    ]),
    repositoryRoot: resolve('.'),
  });
  assert.deepEqual(report.diagnostics, []);
  assert.equal(report.complete, true);
});

test('keeps stale entries and extensions outside the declaration denominator', async () => {
  await withRepository(async (repositoryRoot) => {
    const report = await buildPptxGenJSAudit({
      surface: surface(IDS.slice(0, 2)),
      runtimeProbe,
      manifest: manifest(
        [entry('class:Deck#removed', 'unverified')],
        [{
          id: 'extension:PptxDocument#strictWrite',
          native: ['PptxDocument.write'],
          evidence: emptyEvidence(),
          note: 'strict native write mode',
        }],
      ),
      repositoryRoot,
      gitCommitExists: async () => true,
    });
    assert.equal(report.declarationTotal, 2);
    assert.equal(report.counts.unverified, 2);
    assert.equal(report.counts.stale, 1);
    assert.equal(report.extensions.length, 1);
    assert.deepEqual(report.incompleteIds, [
      'class:Deck#removed',
      IDS[0],
      IDS[1],
    ].sort());

    await assert.rejects(
      buildPptxGenJSAudit({
        surface: surface([IDS[0]]),
        runtimeProbe,
        manifest: manifest([entry('extension:collision', 'unverified')], [{
          id: 'extension:collision',
          native: ['Deck.write'],
          evidence: emptyEvidence(),
          note: 'collision',
        }]),
        repositoryRoot,
        gitCommitExists: async () => true,
      }),
      /extension ID collides with manifest entry/u,
    );
  });
});
