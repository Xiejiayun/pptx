import { open, mkdir, readFile, rename, rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import ts from 'typescript';

import { buildPptxGenJSAudit } from './pptxgenjs-surface-audit-lib.mjs';
import {
  extractPptxGenJSPublicSurface,
  resolvePptxGenJSPackage,
} from './pptxgenjs-surface-declarations.mjs';
import {
  hashRuntimeProbe,
  probePptxGenJSRuntime,
} from './pptxgenjs-runtime-probe.mjs';
import { PPTXGENJS_SURFACE_MANIFEST } from './pptxgenjs-surface-manifest.mjs';

const SCRIPT_PATH = fileURLToPath(import.meta.url);
const DEFAULT_REPOSITORY_ROOT = resolve(dirname(SCRIPT_PATH), '..');
const GROUPS = Object.freeze([
  ['presentation', 'Presentation'],
  ['slide-lifecycle', 'Slide lifecycle'],
  ['text', 'Text'],
  ['shape', 'Shape'],
  ['image', 'Image'],
  ['media', 'Media'],
  ['chart', 'Chart'],
  ['table', 'Table'],
  ['master-layout', 'Master and layout'],
  ['output-runtime', 'Output and runtime'],
  ['other', 'Other'],
]);
const STATUS_ORDER = Object.freeze([
  'supported',
  'deliberate-difference',
  'deprecated-alias',
  'defect-excluded',
  'unsupported',
  'unverified',
  'stale',
]);
let temporarySequence = 0;

function compareText(left, right) {
  return left < right ? -1 : left > right ? 1 : 0;
}

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (value === null || typeof value !== 'object') return value;
  const result = {};
  for (const key of Object.keys(value).sort(compareText)) {
    result[key] = canonicalize(value[key]);
  }
  return result;
}

function escapeMarkdown(value) {
  return String(value)
    .replace(/\\/gu, '\\\\')
    .replace(/\|/gu, '\\|')
    .replace(/\r?\n/gu, '<br>');
}

function evidenceSummary(atom) {
  const values = [];
  for (const category of ['code', 'tests', 'package', 'ooxml', 'clients']) {
    for (const link of atom.evidence?.[category] ?? []) {
      values.push(`${category}:${link.path}`);
    }
  }
  if (atom.control) values.push(`control:${atom.control.path}`);
  return values.length > 0 ? values.join('<br>') : '—';
}

export function groupAuditAtom(atom) {
  const owner = String(atom.owner ?? '').toLowerCase();
  const name = String(atom.name ?? '').toLowerCase();
  const id = String(atom.id ?? '').toLowerCase();
  const text = `${owner} ${name} ${id}`;

  if (/output|write|stream|version/u.test(text)) return 'output-runtime';
  if (
    id.startsWith('property:slide#')
    || id.startsWith('method:slide#') && /addnotes/u.test(id)
    || /addslide|addsection|hidden|slidenumber|newautopagedslides|background/u.test(text)
  ) {
    return 'slide-lifecycle';
  }
  if (/master|layout|placeholder|theme/u.test(text)) return 'master-layout';
  if (/media/u.test(text)) return 'media';
  if (/image/u.test(text)) return 'image';
  if (/chart|axis|legend|datalabel|gridline/u.test(text)) return 'chart';
  if (/table|cell|row/u.test(text)) return 'table';
  if (/text|font|paragraph|hyperlink/u.test(owner)) return 'text';
  if (/shape|fill|line|shadow/u.test(owner)) return 'shape';
  if (/text|font|paragraph|hyperlink/u.test(text)) return 'text';
  if (/shape|fill|line|shadow/u.test(text)) return 'shape';
  if (id.startsWith('class:')) return 'presentation';
  return 'other';
}

export function renderAuditJson(report) {
  return `${JSON.stringify(canonicalize(report), null, 2)}\n`;
}

export function renderAuditMarkdown(report) {
  const lines = [
    '# PptxGenJS 4.0.1 public-surface audit',
    '',
    `- Complete: \`${String(report.complete)}\``,
    `- Declaration atoms: \`${report.declarationTotal}\``,
    `- Declaration SHA-256: \`${report.declarationSha256}\``,
    `- Runtime entry SHA-256: \`${report.runtimeEntrySha256}\``,
    `- Runtime probe SHA-256: \`${report.runtimeProbeSha256}\``,
    '',
    '## Status',
    '',
    '| Status | Count |',
    '| --- | ---: |',
  ];
  for (const status of STATUS_ORDER) lines.push(`| ${status} | ${report.counts[status]} |`);

  lines.push('', '## Runtime declaration differences', '');
  if (report.runtimeMismatches.length === 0) {
    lines.push('None.');
  } else {
    for (const mismatch of report.runtimeMismatches) {
      lines.push(`- \`${escapeMarkdown(mismatch.kind)}:${escapeMarkdown(mismatch.owner)}.${escapeMarkdown(mismatch.name)}\``);
    }
  }

  lines.push('', '## Diagnostics', '');
  if (report.diagnostics.length === 0) {
    lines.push('None.');
  } else {
    for (const diagnostic of report.diagnostics) {
      lines.push(`- \`${escapeMarkdown(diagnostic.id)}\` / \`${escapeMarkdown(diagnostic.code)}\`: ${escapeMarkdown(diagnostic.message)}`);
    }
  }

  lines.push('', '## Incomplete IDs', '');
  if (report.incompleteIds.length === 0) {
    lines.push('None.');
  } else {
    for (const id of report.incompleteIds) lines.push(`- \`${escapeMarkdown(id)}\``);
  }

  const atomsByGroup = new Map(GROUPS.map(([group]) => [group, []]));
  for (const atom of report.atoms) atomsByGroup.get(groupAuditAtom(atom)).push(atom);
  for (const stale of report.staleEntries ?? []) atomsByGroup.get('other').push(stale);
  for (const [group, title] of GROUPS) {
    const atoms = atomsByGroup.get(group);
    if (atoms.length === 0) continue;
    atoms.sort((left, right) => compareText(left.id, right.id));
    lines.push(
      '',
      `## ${title}`,
      '',
      '| Atom | Status | Native | Evidence | Note |',
      '| --- | --- | --- | --- | --- |',
    );
    for (const atom of atoms) {
      lines.push(`| \`${escapeMarkdown(atom.id)}\` | ${escapeMarkdown(atom.status)} | ${escapeMarkdown(atom.native?.join('<br>') || '—')} | ${escapeMarkdown(evidenceSummary(atom))} | ${escapeMarkdown(atom.note || '—')} |`);
    }
  }
  return `${lines.join('\n')}\n`;
}

export async function createAuditReport({
  repositoryRoot = DEFAULT_REPOSITORY_ROOT,
} = {}) {
  const root = resolve(repositoryRoot);
  const packageInfo = await resolvePptxGenJSPackage(
    join(root, 'packages', 'pptxgenjs-adapter', 'package.json'),
  );
  const sourceText = await readFile(packageInfo.declarationPath, 'utf8');
  const surface = extractPptxGenJSPublicSurface({
    sourceText,
    fileName: 'types/index.d.ts',
    typescript: ts,
  });
  const runtimeProbe = await probePptxGenJSRuntime({ packageInfo, surface });
  const report = await buildPptxGenJSAudit({
    surface,
    runtimeProbe,
    manifest: PPTXGENJS_SURFACE_MANIFEST,
    repositoryRoot: root,
  });
  return Object.freeze({
    ...report,
    runtimeProbeSha256: hashRuntimeProbe(runtimeProbe),
  });
}

function parseMode(argv) {
  let mode = 'check';
  let selected = false;
  for (const argument of argv) {
    if (argument !== '--write' && argument !== '--check') {
      throw new TypeError(`unknown argument ${argument}`);
    }
    if (selected) throw new TypeError('choose only one of --write or --check');
    selected = true;
    mode = argument.slice(2);
  }
  return mode;
}

async function stageTemporaryFile(path, content) {
  const handle = await open(path, 'wx');
  try {
    await handle.writeFile(content, 'utf8');
    await handle.sync();
  } finally {
    await handle.close();
  }
}

async function writeArtifacts(outputDirectory, json, markdown) {
  await mkdir(outputDirectory, { recursive: true });
  const jsonPath = join(outputDirectory, 'pptxgenjs-surface-audit.json');
  const markdownPath = join(outputDirectory, 'pptxgenjs-surface-audit.md');
  const suffix = `${process.pid}.${temporarySequence += 1}.tmp`;
  const temporaryJson = `${jsonPath}.${suffix}`;
  const temporaryMarkdown = `${markdownPath}.${suffix}`;
  try {
    await stageTemporaryFile(temporaryJson, json);
    await stageTemporaryFile(temporaryMarkdown, markdown);
    await rename(temporaryJson, jsonPath);
    await rename(temporaryMarkdown, markdownPath);
  } finally {
    await rm(temporaryJson, { force: true });
    await rm(temporaryMarkdown, { force: true });
  }
}

async function artifactsMatch(outputDirectory, json, markdown) {
  try {
    const [existingJson, existingMarkdown] = await Promise.all([
      readFile(join(outputDirectory, 'pptxgenjs-surface-audit.json'), 'utf8'),
      readFile(join(outputDirectory, 'pptxgenjs-surface-audit.md'), 'utf8'),
    ]);
    return existingJson === json && existingMarkdown === markdown;
  } catch {
    return false;
  }
}

export async function runAuditCli({
  argv = [],
  repositoryRoot = DEFAULT_REPOSITORY_ROOT,
  outputDirectory = join(repositoryRoot, 'docs', 'compatibility'),
  createReport = () => createAuditReport({ repositoryRoot }),
  stdout = process.stdout,
  stderr = process.stderr,
} = {}) {
  const mode = parseMode(argv);
  const report = await createReport();
  const json = renderAuditJson(report);
  const markdown = renderAuditMarkdown(report);

  if (mode === 'write') {
    await writeArtifacts(resolve(outputDirectory), json, markdown);
    stdout.write(`PptxGenJS surface audit written: complete=${String(report.complete)} atoms=${report.declarationTotal}\n`);
    return 0;
  }

  if (!(await artifactsMatch(resolve(outputDirectory), json, markdown))) {
    stderr.write('PptxGenJS surface audit artifact drift detected; run with --write.\n');
    return 1;
  }
  if (!report.complete) {
    stderr.write(`PptxGenJS surface audit incomplete: unsupported=${report.counts.unsupported} unverified=${report.counts.unverified} stale=${report.counts.stale}.\n`);
    return 1;
  }
  stdout.write(`PptxGenJS surface audit verified: complete=true atoms=${report.declarationTotal}\n`);
  return 0;
}

if (process.argv[1] && resolve(process.argv[1]) === SCRIPT_PATH) {
  runAuditCli({ argv: process.argv.slice(2) })
    .then((exitCode) => { process.exitCode = exitCode; })
    .catch((error) => {
      process.stderr.write(`${error.stack ?? error.message}\n`);
      process.exitCode = 1;
    });
}
