#!/usr/bin/env node
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  access,
  mkdir,
  readFile,
  readdir,
  rename,
  stat,
  writeFile,
} from 'node:fs/promises';
import { constants as fsConstants } from 'node:fs';
import { homedir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');

function elapsed(start) {
  return Math.round((performance.now() - start) * 1000) / 1000;
}

async function executable(candidate) {
  try {
    await access(candidate, fsConstants.X_OK);
    return candidate;
  } catch {
    return undefined;
  }
}

async function latestPresentationSkill(root) {
  const entries = await readdir(root, { withFileTypes: true });
  const versions = entries
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort((a, b) => b.localeCompare(a, undefined, { numeric: true }));
  if (versions.length === 0) throw new Error(`No presentation skill found under ${root}`);
  return path.join(root, versions[0], 'skills', 'presentations');
}

export async function resolveRuntime(options = {}) {
  const home = options.home ?? homedir();
  const dependencies = options.dependencies
    ?? process.env.CODEX_WORKSPACE_DEPENDENCIES
    ?? path.join(home, '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies');
  const node = await executable(path.join(dependencies, 'node', 'bin', 'node'));
  const python = await executable(path.join(dependencies, 'python', 'bin', 'python3.12'))
    ?? await executable(path.join(dependencies, 'python', 'bin', 'python3'));
  const skill = options.presentationSkill
    ?? process.env.PRESENTATIONS_SKILL_DIR
    ?? await latestPresentationSkill(path.join(
      home,
      '.codex',
      'plugins',
      'cache',
      'openai-primary-runtime',
      'presentations',
    ));
  if (!node) throw new Error(`Bundled Node runtime is missing under ${dependencies}`);
  if (!python) throw new Error(`Bundled Python runtime is missing under ${dependencies}`);
  const tools = path.join(skill, 'container_tools');
  for (const file of ['render_slides.py', 'slides_test.py', 'create_montage.py']) {
    await access(path.join(tools, file));
  }
  return { dependencies, node, python, presentationSkill: skill, tools };
}

export function parseValidateJson(text, maxWarnings = 0) {
  try {
    const parsed = JSON.parse(text.trim());
    const errorCount = Number(parsed?.data?.errorCount ?? 0);
    const warningCount = Number(parsed?.data?.warningCount ?? 0);
    return {
      ok: parsed?.ok === true
        && parsed?.data?.valid === true
        && errorCount === 0
        && warningCount <= maxWarnings,
      valid: parsed?.data?.valid === true,
      errorCount,
      warningCount,
      diagnostics: parsed?.data?.diagnostics ?? [],
    };
  } catch (error) {
    return { ok: false, errorCount: 1, warningCount: 0, parseError: String(error) };
  }
}

export function parseOverflow(text) {
  const passed = text.includes('Test passed. No overflow detected.');
  const failingSlides = [...text.matchAll(/(?:Slide|slide)[^0-9]*(\d+)/g)]
    .map((match) => Number(match[1]))
    .filter((value, index, values) => values.indexOf(value) === index);
  return { ok: passed && !text.includes('ERROR:'), failingSlides };
}

export async function validateRenderedSlides(renderDir, expectedSlides) {
  const entries = (await readdir(renderDir))
    .filter((entry) => /^slide-\d+\.png$/u.test(entry));
  if (entries.length !== expectedSlides) {
    throw new Error(`Expected ${expectedSlides} rendered slides, found ${entries.length}`);
  }
  for (const entry of entries) {
    const metadata = await stat(path.join(renderDir, entry));
    if (metadata.size === 0) throw new Error(`Rendered slide is empty: ${entry}`);
  }
  return entries.sort((a, b) => Number(a.match(/\d+/u)[0]) - Number(b.match(/\d+/u)[0]));
}

function runCommand(command, args, timeoutMs) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: repoRoot, env: process.env, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill('SIGKILL');
    }, timeoutMs);
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.once('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.once('close', (code, signal) => {
      clearTimeout(timer);
      resolve({ code, signal, timedOut, stdout, stderr });
    });
  });
}

async function stage(name, task) {
  const start = performance.now();
  try {
    const details = await task();
    return { name, ok: details.ok !== false, ms: elapsed(start), ...details };
  } catch (error) {
    return { name, ok: false, ms: elapsed(start), error: String(error) };
  }
}

function requireSuccessfulProcess(result, label) {
  if (result.timedOut) throw new Error(`${label} timed out`);
  if (result.code !== 0) {
    throw new Error(`${label} exited ${result.code}: ${result.stderr.trim() || result.stdout.trim()}`);
  }
}

export async function runQa(options) {
  const started = performance.now();
  const input = path.resolve(options.input);
  const outDir = path.resolve(options.outDir);
  const expectedSlides = options.expectedSlides;
  const profile = options.profile ?? 'powerpoint-2010';
  const slaMs = options.slaMs ?? 45_000;
  const timeoutMs = options.timeoutMs ?? 60_000;
  const maxWarnings = options.maxWarnings ?? 0;
  const runtime = await resolveRuntime(options.runtime);
  await access(input);
  await mkdir(path.dirname(outDir), { recursive: true });
  await mkdir(outDir);

  const renderDir = path.join(outDir, 'slides');
  const montage = path.join(outDir, 'montage.png');
  const cli = path.join(repoRoot, 'packages', 'pptx', 'dist', 'cli.js');
  const sdk = path.join(repoRoot, 'packages', 'pptx', 'dist', 'index.js');
  const bytes = await readFile(input);
  const contentHash = createHash('sha256').update(bytes).digest('hex');

  const reopenPromise = stage('reopen', async () => {
    const { PptxDocument } = await import(pathToFileURL(sdk).href);
    const document = await PptxDocument.open(bytes);
    const slideCount = document.slides.length;
    const titles = document.slides.map((slide) => slide.title.text);
    return {
      ok: slideCount > 0 && (expectedSlides === undefined || slideCount === expectedSlides),
      slideCount,
      titles,
    };
  });
  const validatePromise = stage('validate', async () => {
    const result = await runCommand(runtime.node, [
      cli, '--json', 'package', 'validate', input, '--profile', profile,
    ], timeoutMs);
    requireSuccessfulProcess(result, 'package validation');
    return { ...parseValidateJson(result.stdout, maxWarnings), stderr: result.stderr.trim() };
  });
  const renderPromise = stage('render', async () => {
    const result = await runCommand(runtime.python, [
      path.join(runtime.tools, 'render_slides.py'), input, '--output_dir', renderDir,
    ], timeoutMs);
    requireSuccessfulProcess(result, 'slide render');
    const reopened = await reopenPromise;
    if (!reopened.ok) throw new Error('Cannot validate render count because reopen failed');
    const files = await validateRenderedSlides(renderDir, reopened.slideCount);
    return { ok: true, count: files.length, files, stderr: result.stderr.trim() };
  });
  const overflowPromise = stage('overflow', async () => {
    const result = await runCommand(runtime.python, [
      path.join(runtime.tools, 'slides_test.py'), input,
    ], timeoutMs);
    requireSuccessfulProcess(result, 'overflow check');
    return { ...parseOverflow(result.stdout), stdout: result.stdout.trim(), stderr: result.stderr.trim() };
  });
  const montagePromise = stage('montage', async () => {
    const rendered = await renderPromise;
    if (!rendered.ok) throw new Error('Cannot create montage because rendering failed');
    const result = await runCommand(runtime.python, [
      path.join(runtime.tools, 'create_montage.py'),
      '--input_dir', renderDir,
      '--output_file', montage,
      '--fail_on_image_error',
    ], timeoutMs);
    requireSuccessfulProcess(result, 'montage');
    const metadata = await stat(montage);
    if (metadata.size === 0) throw new Error('Montage is empty');
    return { ok: true, bytes: metadata.size, path: montage, stderr: result.stderr.trim() };
  });

  const [reopen, validate, render, overflow, montageStage] = await Promise.all([
    reopenPromise, validatePromise, renderPromise, overflowPromise, montagePromise,
  ]);
  const stages = { reopen, validate, render, overflow, montage: montageStage };
  const failures = Object.values(stages)
    .filter((candidate) => !candidate.ok)
    .map((candidate) => candidate.name);
  const totalMs = elapsed(started);
  if (totalMs > slaMs) failures.push('sla');
  const timedOut = Object.values(stages).some((candidate) => candidate.error?.includes('timed out'));
  const result = {
    ok: failures.length === 0,
    input,
    contentHash,
    runtime,
    profile,
    slideCount: reopen.slideCount ?? 0,
    slaMs,
    timeoutMs,
    maxWarnings,
    elapsedMs: totalMs,
    criticalPathMs: Math.max(reopen.ms, validate.ms, render.ms, overflow.ms, montageStage.ms),
    stages,
    outputs: { outDir, renderDir, montage },
    failures,
    timedOut,
  };
  const temporary = path.join(outDir, '.qa-result.json.tmp');
  const final = path.join(outDir, 'qa-result.json');
  await writeFile(temporary, `${JSON.stringify(result, null, 2)}\n`);
  await rename(temporary, final);
  return result;
}

function parseArgs(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index];
    if (!value.startsWith('--') && !options.input) options.input = value;
    else if (value === '--out-dir') options.outDir = argv[++index];
    else if (value === '--expected-slides') options.expectedSlides = Number(argv[++index]);
    else if (value === '--profile') options.profile = argv[++index];
    else if (value === '--sla-ms') options.slaMs = Number(argv[++index]);
    else if (value === '--timeout-ms') options.timeoutMs = Number(argv[++index]);
    else if (value === '--max-warnings') options.maxWarnings = Number(argv[++index]);
    else if (value !== '--json') throw new Error(`Unknown argument: ${value}`);
  }
  if (!options.input) throw new Error('Usage: ppt-fast-qa.mjs <deck.pptx> --out-dir <new-directory> [options]');
  if (!options.outDir) throw new Error('--out-dir is required and must name a new directory');
  return options;
}

export async function main(argv = process.argv.slice(2)) {
  try {
    const result = await runQa(parseArgs(argv));
    process.stdout.write(`${JSON.stringify(result)}\n`);
    process.exitCode = result.ok ? 0 : result.timedOut || result.failures.includes('sla') ? 3 : 1;
  } catch (error) {
    process.stdout.write(`${JSON.stringify({ ok: false, failures: ['runtime'], error: String(error) })}\n`);
    process.exitCode = 2;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
