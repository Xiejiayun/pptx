import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';

import {
  groupAuditAtom,
  renderAuditJson,
  renderAuditMarkdown,
  runAuditCli,
} from './pptxgenjs-surface-audit.mjs';

const EMPTY_EVIDENCE = Object.freeze({
  code: Object.freeze([]),
  tests: Object.freeze([]),
  package: Object.freeze([]),
  ooxml: Object.freeze([]),
  clients: Object.freeze([]),
});

function fixtureAtom(id, owner, name, status = 'supported') {
  return {
    id,
    kind: 'property',
    owner,
    name,
    declaredIn: owner,
    optional: false,
    readonly: false,
    deprecated: false,
    signatures: [],
    deprecatedSignatures: [],
    catalogKey: '',
    typeText: 'string',
    status,
    native: status === 'supported' ? [`Native.${name}`] : [],
    evidence: EMPTY_EVIDENCE,
    note: `fixture | ${name}`,
    serialization: false,
    client: false,
  };
}

function fixtureReport(overrides = {}) {
  return {
    schemaVersion: 1,
    packageVersion: '4.0.1',
    declarationSha256: 'a'.repeat(64),
    runtimeEntrySha256: 'b'.repeat(64),
    runtimeProbeSha256: 'c'.repeat(64),
    declarationTotal: 3,
    counts: {
      supported: 3,
      'deliberate-difference': 0,
      'deprecated-alias': 0,
      'defect-excluded': 0,
      unsupported: 0,
      unverified: 0,
      stale: 0,
    },
    complete: true,
    incompleteIds: [],
    diagnostics: [],
    atoms: [
      fixtureAtom('class:Deck#write', 'Deck', 'write'),
      fixtureAtom('interface:TextProps@property:fontSize', 'TextProps', 'fontSize'),
      fixtureAtom('interface:ChartProps@property:title', 'ChartProps', 'title'),
    ],
    staleEntries: [],
    extensions: [],
    runtimeMismatches: [],
    ...overrides,
  };
}

function captureStream() {
  let value = '';
  return {
    stream: { write: (chunk) => { value += String(chunk); } },
    read: () => value,
  };
}

async function withDirectory(run) {
  const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-audit-cli-'));
  try {
    return await run(directory);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

test('renders canonical JSON and grouped Markdown deterministically', () => {
  const report = fixtureReport();
  const firstJson = renderAuditJson(report);
  const secondJson = renderAuditJson(report);
  const firstMarkdown = renderAuditMarkdown(report);
  const secondMarkdown = renderAuditMarkdown(report);

  assert.equal(firstJson, secondJson);
  assert.equal(firstMarkdown, secondMarkdown);
  assert.equal(firstJson.endsWith('\n'), true);
  assert.equal(firstMarkdown.endsWith('\n'), true);
  const parsed = JSON.parse(firstJson);
  assert.deepEqual(parsed.counts, report.counts);
  assert.deepEqual(Object.keys(parsed), [...Object.keys(parsed)].sort());
  assert.equal(firstJson.includes('/Users/'), false);
  assert.equal(firstJson.includes('timestamp'), false);
  assert.equal(firstMarkdown.indexOf('## Text') >= 0, true);
  assert.equal(firstMarkdown.indexOf('## Text') < firstMarkdown.indexOf('## Chart'), true);
  assert.match(firstMarkdown, /fixture \\| fontSize/u);
});

test('assigns every reviewer group through fixed pure rules', () => {
  assert.equal(groupAuditAtom(fixtureAtom('class:Deck#write', 'Deck', 'write')), 'output-runtime');
  assert.equal(groupAuditAtom(fixtureAtom('class:Deck#addSlide', 'Deck', 'addSlide')), 'slide-lifecycle');
  assert.equal(groupAuditAtom(fixtureAtom('interface:TextProps@property:font', 'TextProps', 'font')), 'text');
  assert.equal(groupAuditAtom(fixtureAtom('interface:ShapeProps@property:fill', 'ShapeProps', 'fill')), 'shape');
  assert.equal(groupAuditAtom(fixtureAtom('interface:ShapeProps@property:hyperlink', 'ShapeProps', 'hyperlink')), 'shape');
  assert.equal(groupAuditAtom(fixtureAtom('interface:ImageProps@property:path', 'ImageProps', 'path')), 'image');
  assert.equal(groupAuditAtom(fixtureAtom('interface:MediaProps@property:path', 'MediaProps', 'path')), 'media');
  assert.equal(groupAuditAtom(fixtureAtom('interface:ChartProps@property:title', 'ChartProps', 'title')), 'chart');
  assert.equal(groupAuditAtom(fixtureAtom('interface:TableProps@property:rows', 'TableProps', 'rows')), 'table');
  assert.equal(groupAuditAtom(fixtureAtom('interface:SlideMasterProps@property:title', 'SlideMasterProps', 'title')), 'master-layout');
  assert.equal(groupAuditAtom(fixtureAtom('method:Slide#addText', 'Slide', 'addText')), 'text');
  assert.equal(groupAuditAtom(fixtureAtom('class:Deck#tableToSlides', 'Deck', 'tableToSlides')), 'table');
  assert.equal(groupAuditAtom(fixtureAtom('interface:Unknown@property:value', 'Unknown', 'value')), 'other');
});

test('writes byte-identical artifacts and checks a complete report without mutation', async () => {
  await withDirectory(async (outputDirectory) => {
    const stdout = captureStream();
    const stderr = captureStream();
    const options = {
      outputDirectory,
      createReport: async () => fixtureReport(),
      stdout: stdout.stream,
      stderr: stderr.stream,
    };
    assert.equal(await runAuditCli({ ...options, argv: ['--write'] }), 0);
    const jsonPath = join(outputDirectory, 'pptxgenjs-surface-audit.json');
    const markdownPath = join(outputDirectory, 'pptxgenjs-surface-audit.md');
    const first = [await readFile(jsonPath, 'utf8'), await readFile(markdownPath, 'utf8')];
    assert.equal(await runAuditCli({ ...options, argv: ['--write'] }), 0);
    const second = [await readFile(jsonPath, 'utf8'), await readFile(markdownPath, 'utf8')];
    assert.deepEqual(first, second);
    assert.equal(await runAuditCli({ ...options, argv: [] }), 0);
    const third = [await readFile(jsonPath, 'utf8'), await readFile(markdownPath, 'utf8')];
    assert.deepEqual(third, second);
    assert.equal(stderr.read(), '');
    assert.match(stdout.read(), /complete=true/u);
  });
});

test('fails check mode on artifact drift or incomplete statuses', async () => {
  await withDirectory(async (outputDirectory) => {
    const stderr = captureStream();
    const completeOptions = {
      outputDirectory,
      createReport: async () => fixtureReport(),
      stdout: captureStream().stream,
      stderr: stderr.stream,
    };
    await runAuditCli({ ...completeOptions, argv: ['--write'] });
    const jsonPath = join(outputDirectory, 'pptxgenjs-surface-audit.json');
    await writeFile(jsonPath, '{}\n');
    assert.equal(await runAuditCli({ ...completeOptions, argv: ['--check'] }), 1);
    assert.match(stderr.read(), /artifact drift/u);

    const incomplete = fixtureReport({
      complete: false,
      counts: {
        ...fixtureReport().counts,
        supported: 2,
        unverified: 1,
      },
      incompleteIds: ['interface:ChartProps@property:title'],
      atoms: [
        fixtureAtom('class:Deck#write', 'Deck', 'write'),
        fixtureAtom('interface:TextProps@property:fontSize', 'TextProps', 'fontSize'),
        fixtureAtom('interface:ChartProps@property:title', 'ChartProps', 'title', 'unverified'),
      ],
    });
    const incompleteOptions = { ...completeOptions, createReport: async () => incomplete };
    assert.equal(await runAuditCli({ ...incompleteOptions, argv: ['--write'] }), 0);
    assert.equal(await runAuditCli({ ...incompleteOptions, argv: ['--check'] }), 1);
  });
});

test('rejects invalid flags before touching existing artifacts', async () => {
  await withDirectory(async (outputDirectory) => {
    const jsonPath = join(outputDirectory, 'pptxgenjs-surface-audit.json');
    const markdownPath = join(outputDirectory, 'pptxgenjs-surface-audit.md');
    await writeFile(jsonPath, 'json sentinel');
    await writeFile(markdownPath, 'markdown sentinel');
    const options = {
      outputDirectory,
      createReport: async () => fixtureReport(),
      stdout: captureStream().stream,
      stderr: captureStream().stream,
    };
    await assert.rejects(
      runAuditCli({ ...options, argv: ['--write', '--check'] }),
      /choose only one/u,
    );
    await assert.rejects(
      runAuditCli({ ...options, argv: ['--unknown'] }),
      /unknown argument/u,
    );
    assert.equal(await readFile(jsonPath, 'utf8'), 'json sentinel');
    assert.equal(await readFile(markdownPath, 'utf8'), 'markdown sentinel');
  });
});
