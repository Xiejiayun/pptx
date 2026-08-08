import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, rename, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import {
  beginAcceptanceRun,
  buildAcceptanceRun,
  commitFinalVerdict,
  finalizeAcceptanceRun,
} from './ppt-fast-accept.mjs';

const sha256 = async (file) => createHash('sha256').update(await readFile(file)).digest('hex');

async function newRun(options = {}) {
  const parent = await mkdtemp(path.join(tmpdir(), 'ppt-fast-accept-'));
  const runDir = path.join(parent, 'run');
  await beginAcceptanceRun({
    runDir,
    query: 'Generate a visually rich presentation on Amazon rainforest biodiversity.',
    expectedSlides: 2,
    slaMs: options.slaMs ?? 180_000,
    startedNs: options.startedNs,
  });
  await writeFile(path.join(runDir, 'content.json'), '{"title":"Amazon","slides":[]}\n');
  return runDir;
}

async function fakeCreate(_content, output, deckSpecOutput) {
  await writeFile(deckSpecOutput, '{"schemaVersion":1}\n');
  await writeFile(output, Buffer.from('fake-pptx'));
  return { slideCount: 2 };
}

async function fakeQa(options) {
  await mkdir(options.outDir);
  const renderDir = path.join(options.outDir, 'slides');
  const montage = path.join(options.outDir, 'montage.png');
  await mkdir(renderDir);
  await writeFile(path.join(renderDir, 'slide-1.png'), Buffer.from('slide-1'));
  await writeFile(path.join(renderDir, 'slide-2.png'), Buffer.from('slide-2'));
  await writeFile(montage, Buffer.from('montage'));
  const result = {
    ok: true,
    input: options.input,
    contentHash: await sha256(options.input),
    slideCount: 2,
    elapsedMs: 12,
    stages: {
      reopen: { ok: true, slideCount: 2 },
      validate: { ok: true, warningCount: 0 },
      render: { ok: true, count: 2 },
      overflow: { ok: true },
      montage: { ok: true },
    },
    outputs: { outDir: options.outDir, renderDir, montage },
    failures: [],
  };
  await writeFile(path.join(options.outDir, 'qa-result.json'), `${JSON.stringify(result)}\n`);
  return result;
}

async function fakeQaSlaOnly(options) {
  const result = await fakeQa(options);
  result.ok = false;
  result.failures = ['sla'];
  await writeFile(path.join(options.outDir, 'qa-result.json'), `${JSON.stringify(result)}\n`);
  return result;
}

async function fakeQaMissingValidation(options) {
  const result = await fakeQa(options);
  delete result.stages.validate;
  await writeFile(path.join(options.outDir, 'qa-result.json'), `${JSON.stringify(result)}\n`);
  return result;
}

async function builtRun(options = {}) {
  const runDir = await newRun(options);
  await buildAcceptanceRun({ runDir }, { createFastPresentation: fakeCreate, runQa: fakeQa });
  const pptxSha256 = await sha256(path.join(runDir, 'deck.pptx'));
  return { runDir, pptxSha256 };
}

async function reviewBinding(runDir) {
  const manifestPath = path.join(runDir, 'review-manifest.json');
  const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
  return {
    reviewManifestSha256: await sha256(manifestPath),
    qaResultSha256: manifest.qaResultSha256,
    montageSha256: manifest.montageSha256,
    renderedSlideSha256s: manifest.renderedSlideSha256s,
  };
}

async function passingReview(runDir, pptxSha256, overrides = {}) {
  return {
    verdict: 'pass',
    inspectedSlides: [1, 2],
    inspectedMontage: true,
    querySatisfied: true,
    contentConsistent: true,
    sourcesPresent: true,
    issues: [],
    pptxSha256,
    ...await reviewBinding(runDir),
    ...overrides,
  };
}

function runCli(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [path.resolve('scripts/ppt-fast-accept.mjs'), ...args], {
      cwd: path.resolve('.'),
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.once('error', reject);
    child.once('close', (code, signal) => resolve({ code, signal, stdout, stderr }));
  });
}

test('acceptance begin creates an exclusive run directory and starts before generation', async () => {
  const runDir = await newRun();
  const state = JSON.parse(await readFile(path.join(runDir, 'run-state.json'), 'utf8'));
  assert.equal(state.status, 'active');
  assert.equal(state.expectedSlides, 2);
  assert.match(await readFile(path.join(runDir, 'query.txt'), 'utf8'), /Amazon rainforest biodiversity/);
  await assert.rejects(beginAcceptanceRun({ runDir, query: 'duplicate', expectedSlides: 2 }), /EEXIST/);
});

test('acceptance finalizes only after QA artifacts and complete visual review are hash-bound', async () => {
  const { runDir, pptxSha256 } = await builtRun();
  const template = JSON.parse(await readFile(path.join(runDir, 'review-template.json'), 'utf8'));
  assert.equal(template.verdict, 'pending');
  assert.equal(template.pptxSha256, pptxSha256);
  assert.deepEqual(template.inspectedSlides, []);
  const review = await passingReview(runDir, pptxSha256);
  const { verdict } = await finalizeAcceptanceRun({ runDir, review });
  assert.equal(verdict.verdict, 'pass');
  assert.equal(verdict.qualityPass, true);
  assert.equal(verdict.slaPass, true);
  assert.equal(verdict.artifacts.renderedSlides.length, 2);
  assert.equal(JSON.parse(await readFile(path.join(runDir, 'run-state.json'), 'utf8')).status, 'passed');
});

test('acceptance records failure for stale hashes or incomplete visual coverage', async () => {
  const first = await builtRun();
  const firstReview = await passingReview(first.runDir, '0'.repeat(64));
  const hashMismatch = await finalizeAcceptanceRun({
    runDir: first.runDir,
    review: firstReview,
  });
  assert.equal(hashMismatch.verdict.verdict, 'fail');
  assert.ok(hashMismatch.verdict.failures.includes('review_pptx_hash_mismatch'));

  const second = await builtRun();
  const secondReview = await passingReview(second.runDir, second.pptxSha256, { inspectedSlides: [1] });
  const incomplete = await finalizeAcceptanceRun({
    runDir: second.runDir,
    review: secondReview,
  });
  assert.equal(incomplete.verdict.verdict, 'fail');
  assert.ok(incomplete.verdict.failures.includes('visual_review_incomplete'));

  const third = await builtRun();
  await writeFile(path.join(third.runDir, 'content.json'), '{"title":"mutated after QA"}\n');
  const thirdReview = await passingReview(third.runDir, third.pptxSha256);
  const mutated = await finalizeAcceptanceRun({
    runDir: third.runDir,
    review: thirdReview,
  });
  assert.equal(mutated.verdict.verdict, 'fail');
  assert.ok(mutated.verdict.failures.includes('content_hash_mismatch'));
});

test('acceptance keeps quality and SLA results separate', async () => {
  const startedNs = process.hrtime.bigint() - 200_000_000n;
  const { runDir, pptxSha256 } = await builtRun({ slaMs: 50, startedNs });
  const review = await passingReview(runDir, pptxSha256);
  const { verdict } = await finalizeAcceptanceRun({ runDir, review });
  assert.equal(verdict.verdict, 'fail');
  assert.equal(verdict.qualityPass, true);
  assert.equal(verdict.slaPass, false);
  assert.equal(verdict.finalizationReserveMs, 2_000);
  assert.equal(verdict.verdictPublishReserveMs, 1_000);
  assert.ok(verdict.failures.includes('sla'));
});

test('acceptance does not treat a QA-only sub-SLA miss as a quality failure', async () => {
  const runDir = await newRun();
  await buildAcceptanceRun({ runDir }, {
    createFastPresentation: fakeCreate,
    runQa: fakeQaSlaOnly,
  });
  const state = JSON.parse(await readFile(path.join(runDir, 'run-state.json'), 'utf8'));
  assert.equal(state.status, 'qa-complete');
  assert.equal(state.attempts[0].qaOk, true);
  assert.equal(state.attempts[0].qaBudgetPass, false);
  const pptxSha256 = await sha256(path.join(runDir, 'deck.pptx'));
  const review = await passingReview(runDir, pptxSha256);
  const { verdict } = await finalizeAcceptanceRun({ runDir, review });
  assert.equal(verdict.qualityPass, true);
  assert.equal(verdict.verdict, 'pass');
});

test('acceptance rejects an incomplete QA result even when its failures array is empty', async () => {
  const runDir = await newRun();
  await assert.rejects(
    buildAcceptanceRun({ runDir }, {
      createFastPresentation: fakeCreate,
      runQa: fakeQaMissingValidation,
    }),
    /Acceptance QA failed: validate/,
  );
  const state = JSON.parse(await readFile(path.join(runDir, 'run-state.json'), 'utf8'));
  assert.equal(state.status, 'repairable');
});

test('acceptance retries verdict commit until the recorded elapsed bound covers the actual write', async () => {
  const parent = await mkdtemp(path.join(tmpdir(), 'ppt-fast-verdict-'));
  const finalVerdict = path.join(parent, 'final-verdict.json');
  let writes = 0;
  const delayedWrite = async (file, value) => {
    writes += 1;
    if (writes === 1) {
      await assert.rejects(readFile(finalVerdict), /ENOENT/);
      await new Promise((resolve) => setTimeout(resolve, 10));
    }
    await writeFile(file, `${JSON.stringify(value)}\n`);
  };
  const verdict = await commitFinalVerdict(
    { finalVerdict },
    { qualityPass: true, slaMs: 10_000, failures: [] },
    process.hrtime.bigint(),
    { atomicWriteJson: delayedWrite, reserveMs: 1, publishReserveMs: 1 },
  );
  assert.ok(verdict.verdictCommitAttempt >= 2);
  assert.equal(verdict.verdict, 'pass');
  assert.ok(verdict.elapsedMs >= verdict.observedElapsedMs);
  assert.deepEqual(JSON.parse(await readFile(finalVerdict, 'utf8')), verdict);
});

test('acceptance measures publication privately and exposes the canonical verdict only once', async () => {
  const parent = await mkdtemp(path.join(tmpdir(), 'ppt-fast-publish-'));
  const finalVerdict = path.join(parent, 'final-verdict.json');
  let canonicalPublishes = 0;
  const delayedPublish = async (source, target) => {
    if (target === finalVerdict) canonicalPublishes += 1;
    await new Promise((resolve) => setTimeout(resolve, 10));
    await rename(source, target);
  };
  const verdict = await commitFinalVerdict(
    { finalVerdict },
    { qualityPass: true, slaMs: 10_000, failures: [] },
    process.hrtime.bigint(),
    { rename: delayedPublish },
  );
  assert.equal(verdict.verdict, 'pass');
  assert.equal(canonicalPublishes, 1);
  assert.ok(verdict.verdictPublishReserveMs >= 120);
  assert.deepEqual(JSON.parse(await readFile(finalVerdict, 'utf8')), verdict);
});

test('acceptance rejects content changed while the exact parsed bytes are compiling', async () => {
  const runDir = await newRun();
  const mutatingCreate = async (...args) => {
    const result = await fakeCreate(...args);
    await writeFile(path.join(runDir, 'content.json'), '{"title":"changed during compile","slides":[]}\n');
    return result;
  };
  await assert.rejects(
    buildAcceptanceRun({ runDir }, { createFastPresentation: mutatingCreate, runQa: fakeQa }),
    /content changed during compilation/,
  );
  const state = JSON.parse(await readFile(path.join(runDir, 'run-state.json'), 'utf8'));
  assert.equal(state.status, 'repairable');
});

test('acceptance writes a terminal verdict after the second failed build', async () => {
  const runDir = await newRun();
  const failCreate = async () => { throw new Error('intentional build failure'); };
  await assert.rejects(
    buildAcceptanceRun({ runDir }, { createFastPresentation: failCreate, runQa: fakeQa }),
    /intentional build failure/,
  );
  await assert.rejects(
    buildAcceptanceRun({ runDir }, { createFastPresentation: failCreate, runQa: fakeQa }),
    /intentional build failure/,
  );
  const state = JSON.parse(await readFile(path.join(runDir, 'run-state.json'), 'utf8'));
  const verdict = JSON.parse(await readFile(path.join(runDir, 'final-verdict.json'), 'utf8'));
  assert.equal(state.status, 'failed');
  assert.equal(verdict.verdict, 'fail');
  assert.deepEqual(verdict.failures, ['build_failed']);
});

test('acceptance serializes concurrent builds with the shared run lock', async () => {
  const runDir = await newRun();
  const slowCreate = async (...args) => {
    await new Promise((resolve) => setTimeout(resolve, 50));
    return fakeCreate(...args);
  };
  const outcomes = await Promise.allSettled([
    buildAcceptanceRun({ runDir }, { createFastPresentation: slowCreate, runQa: fakeQa }),
    buildAcceptanceRun({ runDir }, { createFastPresentation: slowCreate, runQa: fakeQa }),
  ]);
  assert.equal(outcomes.filter((outcome) => outcome.status === 'fulfilled').length, 1);
  assert.equal(outcomes.filter((outcome) => outcome.status === 'rejected').length, 1);
});

test('acceptance serializes repair and finalization with the shared run lock', async () => {
  const { runDir, pptxSha256 } = await builtRun();
  const review = await passingReview(runDir, pptxSha256);
  const slowCreate = async (...args) => {
    await new Promise((resolve) => setTimeout(resolve, 50));
    return fakeCreate(...args);
  };
  const outcomes = await Promise.allSettled([
    buildAcceptanceRun({ runDir, forceRepair: true }, {
      createFastPresentation: slowCreate,
      runQa: fakeQa,
    }),
    finalizeAcceptanceRun({ runDir, review }),
  ]);
  assert.equal(outcomes.filter((outcome) => outcome.status === 'fulfilled').length, 1);
  assert.equal(outcomes.filter((outcome) => outcome.status === 'rejected').length, 1);
});

test('acceptance requires review hashes for QA, montage, and every rendered slide', async (t) => {
  const cases = [
    ['qaResultSha256', 'review_qa_hash_mismatch', '0'.repeat(64)],
    ['montageSha256', 'review_montage_hash_mismatch', '0'.repeat(64)],
    ['renderedSlideSha256s', 'review_render_hash_mismatch', ['0'.repeat(64), '0'.repeat(64)]],
  ];
  for (const [field, failure, value] of cases) {
    await t.test(field, async () => {
      const { runDir, pptxSha256 } = await builtRun();
      const review = await passingReview(runDir, pptxSha256, { [field]: value });
      const { verdict } = await finalizeAcceptanceRun({ runDir, review });
      assert.equal(verdict.verdict, 'fail');
      assert.ok(verdict.failures.includes(failure));
    });
  }
});

test('acceptance detects a review manifest changed after QA', async () => {
  const { runDir, pptxSha256 } = await builtRun();
  const manifestPath = path.join(runDir, 'review-manifest.json');
  const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
  manifest.montageSha256 = '0'.repeat(64);
  await writeFile(manifestPath, `${JSON.stringify(manifest)}\n`);
  const review = await passingReview(runDir, pptxSha256);
  const { verdict } = await finalizeAcceptanceRun({ runDir, review });
  assert.equal(verdict.verdict, 'fail');
  assert.ok(verdict.failures.includes('review_manifest_hash_mismatch'));
  assert.ok(verdict.failures.includes('review_manifest_artifact_mismatch'));
});

test('acceptance CLI exits nonzero and reports ok false for a failed final verdict', async () => {
  const { runDir, pptxSha256 } = await builtRun();
  const review = await passingReview(runDir, pptxSha256, { inspectedSlides: [1] });
  const reviewFile = path.join(runDir, 'review-input.json');
  await writeFile(reviewFile, `${JSON.stringify(review)}\n`);
  const result = await runCli(['finalize', '--run-dir', runDir, '--review-file', reviewFile]);
  assert.equal(result.code, 3, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.ok, false);
  assert.equal(output.verdict.verdict, 'fail');
});

test('acceptance commits only one final verdict under concurrent finalization', async () => {
  const { runDir, pptxSha256 } = await builtRun();
  const review = await passingReview(runDir, pptxSha256);
  const outcomes = await Promise.allSettled([
    finalizeAcceptanceRun({ runDir, review }),
    finalizeAcceptanceRun({ runDir, review }),
  ]);
  assert.equal(outcomes.filter((outcome) => outcome.status === 'fulfilled').length, 1);
  assert.equal(outcomes.filter((outcome) => outcome.status === 'rejected').length, 1);
  assert.equal(JSON.parse(await readFile(path.join(runDir, 'final-verdict.json'), 'utf8')).verdict, 'pass');
});
