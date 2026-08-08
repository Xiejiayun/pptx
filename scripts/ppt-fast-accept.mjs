#!/usr/bin/env node
import { createHash, randomUUID } from 'node:crypto';
import {
  access,
  mkdir,
  open,
  readFile,
  rename,
  stat,
  unlink,
} from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { createFastPresentation } from './ppt-fast-create.mjs';
import { runQa } from './ppt-fast-qa.mjs';

const DEFAULT_SLA_MS = 180_000;
const MAX_BUILD_ATTEMPTS = 2;
const FINALIZATION_RESERVE_MS = 1_000;
const VERDICT_PUBLISH_RESERVE_MS = 1_000;
const MAX_VERDICT_COMMIT_ATTEMPTS = 3;

function elapsedMs(startNs, endNs = process.hrtime.bigint()) {
  return Math.round(Number(endNs - BigInt(startNs)) / 1_000) / 1_000;
}

async function atomicWriteJson(file, value) {
  const temporary = `${file}.${process.pid}.${randomUUID()}.tmp`;
  const handle = await open(temporary, 'wx');
  try {
    await handle.writeFile(`${JSON.stringify(value, null, 2)}\n`, 'utf8');
    await handle.sync();
  } finally {
    await handle.close();
  }
  await rename(temporary, file);
}

async function sha256(file) {
  const bytes = await readFile(file);
  return createHash('sha256').update(bytes).digest('hex');
}

function sha256Bytes(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

async function readJson(file) {
  return JSON.parse(await readFile(file, 'utf8'));
}

export async function commitFinalVerdict(paths, baseVerdict, startedMonotonicNs, dependencies = {}) {
  const writeJson = dependencies.atomicWriteJson ?? atomicWriteJson;
  const publish = dependencies.rename ?? rename;
  const reserveFloorMs = dependencies.reserveMs ?? FINALIZATION_RESERVE_MS;
  let publishReserveMs = dependencies.publishReserveMs;
  if (publishReserveMs === undefined) {
    const probeSource = `${paths.finalVerdict}.${process.pid}.${randomUUID()}.publish-probe`;
    const probeTarget = `${probeSource}.committed`;
    await atomicWriteJson(probeSource, { probe: true });
    const publishStartedNs = process.hrtime.bigint();
    await publish(probeSource, probeTarget);
    const publishMs = elapsedMs(publishStartedNs);
    await unlink(probeTarget).catch(() => {});
    publishReserveMs = Math.max(
      VERDICT_PUBLISH_RESERVE_MS,
      Math.ceil(publishMs * 2 + 100),
    );
  }
  const qualityFailures = baseVerdict.failures.filter((failure) => failure !== 'sla');
  let observedElapsedMs = elapsedMs(startedMonotonicNs);
  let reserveMs = reserveFloorMs;
  for (let attempt = 1; attempt <= MAX_VERDICT_COMMIT_ATTEMPTS; attempt += 1) {
    const elapsedUpperBoundMs = observedElapsedMs + reserveMs + publishReserveMs;
    const slaPass = elapsedUpperBoundMs < baseVerdict.slaMs;
    const failures = [...qualityFailures, ...(slaPass ? [] : ['sla'])];
    const candidatePath = `${paths.finalVerdict}.${process.pid}.${randomUUID()}.candidate`;
    const candidate = {
      ...baseVerdict,
      verdict: baseVerdict.qualityPass && slaPass ? 'pass' : 'fail',
      slaPass,
      committedAt: new Date().toISOString(),
      elapsedMs: elapsedUpperBoundMs,
      observedElapsedMs,
      finalizationReserveMs: reserveMs + publishReserveMs,
      verdictPublishReserveMs: publishReserveMs,
      verdictCommitAttempt: attempt,
      failures,
    };
    const writeStartedNs = process.hrtime.bigint();
    await writeJson(candidatePath, candidate);
    const preparedElapsedMs = elapsedMs(startedMonotonicNs);
    const writeMs = elapsedMs(writeStartedNs);
    if (preparedElapsedMs + publishReserveMs <= elapsedUpperBoundMs) {
      await publish(candidatePath, paths.finalVerdict);
      return candidate;
    }
    await unlink(candidatePath).catch(() => {});
    observedElapsedMs = preparedElapsedMs;
    reserveMs = Math.max(reserveFloorMs, Math.ceil(writeMs * 2 + 100));
  }
  const elapsed = Math.max(elapsedMs(startedMonotonicNs), baseVerdict.slaMs);
  const terminal = {
    ...baseVerdict,
    verdict: 'fail',
    slaPass: false,
    committedAt: new Date().toISOString(),
    elapsedMs: elapsed,
    observedElapsedMs: elapsed,
    finalizationReserveMs: 0,
    verdictCommitAttempt: MAX_VERDICT_COMMIT_ATTEMPTS + 1,
    timingFailure: 'unable_to_bound_verdict_commit',
    failures: [...qualityFailures, 'sla'],
  };
  await atomicWriteJson(paths.finalVerdict, terminal);
  return terminal;
}

function acceptancePaths(runDir) {
  const root = path.resolve(runDir);
  return {
    root,
    query: path.join(root, 'query.txt'),
    state: path.join(root, 'run-state.json'),
    content: path.join(root, 'content.json'),
    deckSpec: path.join(root, 'deck-spec.json'),
    deck: path.join(root, 'deck.pptx'),
    compileResult: path.join(root, 'compile-result.json'),
    reviewManifest: path.join(root, 'review-manifest.json'),
    reviewTemplate: path.join(root, 'review-template.json'),
    visualReview: path.join(root, 'visual-review.json'),
    finalVerdict: path.join(root, 'final-verdict.json'),
    lock: path.join(root, '.ppt-fast-accept.lock'),
  };
}

async function withRunLock(paths, task) {
  const lock = await open(paths.lock, 'wx');
  try {
    return await task();
  } finally {
    await lock.close();
    await unlink(paths.lock).catch(() => {});
  }
}

export async function beginAcceptanceRun(options) {
  const startedNs = options.startedNs ?? process.hrtime.bigint();
  const startedAt = new Date().toISOString();
  const paths = acceptancePaths(options.runDir);
  const query = String(options.query ?? '').trim();
  if (!query) throw new Error('Acceptance query must not be empty');
  if (!Number.isInteger(options.expectedSlides) || options.expectedSlides < 1) {
    throw new Error('Acceptance expectedSlides must be a positive integer');
  }
  if (!Number.isFinite(options.slaMs ?? DEFAULT_SLA_MS) || (options.slaMs ?? DEFAULT_SLA_MS) <= 0) {
    throw new Error('Acceptance slaMs must be positive');
  }
  await mkdir(paths.root);
  const queryHandle = await open(paths.query, 'wx');
  try {
    await queryHandle.writeFile(`${query}\n`, 'utf8');
    await queryHandle.sync();
  } finally {
    await queryHandle.close();
  }
  const state = {
    schemaVersion: 1,
    status: 'active',
    startedAt,
    startedMonotonicNs: startedNs.toString(),
    slaMs: options.slaMs ?? DEFAULT_SLA_MS,
    expectedSlides: options.expectedSlides,
    maxWarnings: options.maxWarnings ?? 0,
    buildAttempts: 0,
    querySha256: await sha256(paths.query),
    attempts: [],
  };
  await atomicWriteJson(paths.state, state);
  return { paths, state };
}

function requireActiveState(state) {
  if (!['active', 'repairable', 'qa-complete'].includes(state.status)) {
    throw new Error(`Acceptance run is not active: ${state.status}`);
  }
}

function collectQaQualityFailures(qaResult, expectedSlides, maxWarnings) {
  const failures = (Array.isArray(qaResult.failures) ? qaResult.failures : [])
    .filter((failure) => failure !== 'sla');
  if (qaResult.stages?.reopen?.ok !== true || qaResult.slideCount !== expectedSlides) failures.push('reopen');
  if (qaResult.stages?.validate?.ok !== true
    || qaResult.stages.validate.warningCount > maxWarnings) failures.push('validate');
  if (qaResult.stages?.render?.ok !== true
    || qaResult.stages.render.count !== expectedSlides) failures.push('render');
  if (qaResult.stages?.overflow?.ok !== true) failures.push('overflow');
  if (qaResult.stages?.montage?.ok !== true) failures.push('montage');
  return [...new Set(failures)];
}

async function buildLocked(options, dependencies, paths) {
  const state = await readJson(paths.state);
  requireActiveState(state);
  if (state.status === 'qa-complete' && options.forceRepair !== true) {
    throw new Error('Acceptance QA already completed; use forceRepair only for a concrete visual defect');
  }
  const attempt = state.buildAttempts + 1;
  if (attempt > MAX_BUILD_ATTEMPTS) throw new Error('Acceptance run allows at most one repair');
  await access(paths.content);
  const create = dependencies.createFastPresentation ?? createFastPresentation;
  const qa = dependencies.runQa ?? runQa;
  const compileStartedNs = process.hrtime.bigint();
  const attemptRecord = {
    attempt,
    generationMs: elapsedMs(state.startedMonotonicNs, compileStartedNs),
    compileStartedAt: new Date().toISOString(),
  };
  try {
    const contentBytes = await readFile(paths.content);
    const contentSha256 = sha256Bytes(contentBytes);
    const content = JSON.parse(contentBytes.toString('utf8'));
    const compileResult = await create(content, paths.deck, paths.deckSpec);
    if (await sha256(paths.content) !== contentSha256) {
      throw new Error('Acceptance content changed during compilation');
    }
    const compileEndedNs = process.hrtime.bigint();
    attemptRecord.compileMs = elapsedMs(compileStartedNs, compileEndedNs);
    attemptRecord.contentSha256 = contentSha256;
    attemptRecord.deckSpecSha256 = await sha256(paths.deckSpec);
    attemptRecord.pptxSha256 = await sha256(paths.deck);
    await atomicWriteJson(paths.compileResult, {
      schemaVersion: 1,
      attempt,
      ok: true,
      slideCount: compileResult.slideCount,
      elapsedMs: attemptRecord.compileMs,
      contentSha256: attemptRecord.contentSha256,
      deckSpecSha256: attemptRecord.deckSpecSha256,
      pptxSha256: attemptRecord.pptxSha256,
    });
    const qaStartedNs = process.hrtime.bigint();
    const qaDir = path.join(paths.root, `qa-${attempt}`);
    const qaResult = await qa({
      input: paths.deck,
      outDir: qaDir,
      expectedSlides: state.expectedSlides,
      maxWarnings: state.maxWarnings,
      slaMs: options.qaSlaMs ?? 45_000,
      timeoutMs: options.qaTimeoutMs ?? 60_000,
    });
    attemptRecord.qaMs = elapsedMs(qaStartedNs);
    attemptRecord.qaResult = path.relative(paths.root, path.join(qaDir, 'qa-result.json'));
    const qaFailures = Array.isArray(qaResult.failures) ? qaResult.failures : [];
    const qaQualityFailures = collectQaQualityFailures(
      qaResult,
      state.expectedSlides,
      state.maxWarnings,
    );
    attemptRecord.qaOk = qaQualityFailures.length === 0;
    attemptRecord.qaBudgetPass = !qaFailures.includes('sla');
    if (qaQualityFailures.length > 0) {
      throw new Error(`Acceptance QA failed: ${qaQualityFailures.join(',')}`);
    }
    const qaResultPath = path.join(qaDir, 'qa-result.json');
    const montagePath = path.join(qaDir, 'montage.png');
    const renderedSlideSha256s = await Promise.all(Array.from(
      { length: state.expectedSlides },
      (_, index) => sha256(path.join(qaDir, 'slides', `slide-${index + 1}.png`)),
    ));
    const reviewManifest = {
      schemaVersion: 1,
      attempt,
      expectedSlides: state.expectedSlides,
      pptxSha256: attemptRecord.pptxSha256,
      qaResultSha256: await sha256(qaResultPath),
      montageSha256: await sha256(montagePath),
      renderedSlideSha256s,
    };
    await atomicWriteJson(paths.reviewManifest, reviewManifest);
    const reviewManifestSha256 = await sha256(paths.reviewManifest);
    await atomicWriteJson(paths.reviewTemplate, {
      schemaVersion: 1,
      reviewer: 'Codex',
      verdict: 'pending',
      inspectedSlides: [],
      inspectedMontage: false,
      querySatisfied: false,
      contentConsistent: false,
      sourcesPresent: false,
      issues: ['Review every full-size slide and replace this placeholder.'],
      notes: '',
      pptxSha256: reviewManifest.pptxSha256,
      reviewManifestSha256,
      qaResultSha256: reviewManifest.qaResultSha256,
      montageSha256: reviewManifest.montageSha256,
      renderedSlideSha256s: reviewManifest.renderedSlideSha256s,
    });
    state.status = 'qa-complete';
    state.latestAttempt = attempt;
    state.latestQaDir = path.relative(paths.root, qaDir);
    state.qaCompletedMonotonicNs = process.hrtime.bigint().toString();
    state.reviewManifestSha256 = reviewManifestSha256;
    state.buildAttempts = attempt;
    state.attempts.push(attemptRecord);
    delete state.lastFailure;
    await atomicWriteJson(paths.state, state);
    return { paths, state, qaResult };
  } catch (error) {
    attemptRecord.ok = false;
    attemptRecord.error = String(error);
    state.buildAttempts = attempt;
    state.status = attempt < MAX_BUILD_ATTEMPTS ? 'repairable' : 'failed';
    state.lastFailure = String(error);
    state.attempts.push(attemptRecord);
    await atomicWriteJson(paths.state, state);
    if (attempt >= MAX_BUILD_ATTEMPTS) {
      const observedElapsedMs = elapsedMs(state.startedMonotonicNs);
      const terminal = {
        schemaVersion: 1,
        verdict: 'fail',
        qualityPass: false,
        slaPass: observedElapsedMs + FINALIZATION_RESERVE_MS < state.slaMs,
        startedAt: state.startedAt,
        committedAt: new Date().toISOString(),
        elapsedMs: observedElapsedMs + FINALIZATION_RESERVE_MS,
        slaMs: state.slaMs,
        failures: ['build_failed'],
        error: String(error),
      };
      await atomicWriteJson(paths.finalVerdict, terminal);
      state.completedAt = terminal.committedAt;
      state.elapsedMs = terminal.elapsedMs;
      await atomicWriteJson(paths.state, state);
    }
    throw error;
  }
}

export async function buildAcceptanceRun(options, dependencies = {}) {
  const paths = acceptancePaths(options.runDir);
  return withRunLock(paths, () => buildLocked(options, dependencies, paths));
}

function exactSlideCoverage(value, expectedSlides) {
  if (!Array.isArray(value)) return false;
  const actual = [...new Set(value.map(Number))].sort((a, b) => a - b);
  return actual.length === expectedSlides
    && actual.every((slide, index) => slide === index + 1);
}

async function artifact(pathValue, root) {
  const metadata = await stat(pathValue);
  if (!metadata.isFile() || metadata.size === 0) throw new Error(`Missing artifact: ${pathValue}`);
  return { path: path.relative(root, pathValue), sha256: await sha256(pathValue), bytes: metadata.size };
}

async function finalizeLocked(options, paths) {
  const state = await readJson(paths.state);
  if (state.status !== 'qa-complete') throw new Error(`Acceptance QA is not complete: ${state.status}`);
  try {
    await access(paths.finalVerdict);
    throw new Error('Final verdict already exists');
  } catch (error) {
    if (error?.code !== 'ENOENT') throw error;
  }
  const qaResultPath = path.join(paths.root, state.latestQaDir, 'qa-result.json');
  const expectedQaDir = path.join(paths.root, state.latestQaDir);
  const expectedMontage = path.join(expectedQaDir, 'montage.png');
  const expectedRenderDir = path.join(expectedQaDir, 'slides');
  const qaResult = await readJson(qaResultPath);
  const compileResult = await readJson(paths.compileResult);
  const reviewManifest = await readJson(paths.reviewManifest);
  const review = options.review ?? await readJson(options.reviewFile);
  const artifacts = {
    query: await artifact(paths.query, paths.root),
    content: await artifact(paths.content, paths.root),
    deckSpec: await artifact(paths.deckSpec, paths.root),
    pptx: await artifact(paths.deck, paths.root),
    compileResult: await artifact(paths.compileResult, paths.root),
    reviewManifest: await artifact(paths.reviewManifest, paths.root),
    qaResult: await artifact(qaResultPath, paths.root),
    montage: await artifact(expectedMontage, paths.root),
    renderedSlides: await Promise.all(Array.from(
      { length: state.expectedSlides },
      (_, index) => artifact(path.join(expectedRenderDir, `slide-${index + 1}.png`), paths.root),
    )),
  };
  const failures = [];
  const latestAttempt = state.attempts.find((candidate) => candidate.attempt === state.latestAttempt);
  if (state.querySha256 !== artifacts.query.sha256) failures.push('query_hash_mismatch');
  if (latestAttempt?.contentSha256 !== artifacts.content.sha256) failures.push('content_hash_mismatch');
  if (latestAttempt?.deckSpecSha256 !== artifacts.deckSpec.sha256) failures.push('deck_spec_hash_mismatch');
  if (latestAttempt?.pptxSha256 !== artifacts.pptx.sha256) failures.push('compiled_pptx_hash_mismatch');
  if (compileResult.contentSha256 !== artifacts.content.sha256
    || compileResult.deckSpecSha256 !== artifacts.deckSpec.sha256
    || compileResult.pptxSha256 !== artifacts.pptx.sha256) failures.push('compile_result_hash_mismatch');
  if (state.reviewManifestSha256 !== artifacts.reviewManifest.sha256) failures.push('review_manifest_hash_mismatch');
  if (reviewManifest.attempt !== state.latestAttempt
    || reviewManifest.expectedSlides !== state.expectedSlides
    || reviewManifest.pptxSha256 !== artifacts.pptx.sha256
    || reviewManifest.qaResultSha256 !== artifacts.qaResult.sha256
    || reviewManifest.montageSha256 !== artifacts.montage.sha256
    || !Array.isArray(reviewManifest.renderedSlideSha256s)
    || reviewManifest.renderedSlideSha256s.length !== artifacts.renderedSlides.length
    || reviewManifest.renderedSlideSha256s.some(
      (hash, index) => hash !== artifacts.renderedSlides[index].sha256,
    )) failures.push('review_manifest_artifact_mismatch');
  if (path.resolve(qaResult.outputs?.outDir ?? '') !== expectedQaDir) failures.push('qa_output_dir_mismatch');
  if (path.resolve(qaResult.outputs?.montage ?? '') !== expectedMontage) failures.push('qa_montage_path_mismatch');
  if (path.resolve(qaResult.outputs?.renderDir ?? '') !== expectedRenderDir) failures.push('qa_render_dir_mismatch');
  const qaQualityFailures = collectQaQualityFailures(
    qaResult,
    state.expectedSlides,
    state.maxWarnings,
  );
  if (qaQualityFailures.length > 0) failures.push('qa_failed');
  if (path.resolve(qaResult.input ?? '') !== paths.deck) failures.push('qa_input_mismatch');
  if (qaResult.slideCount !== state.expectedSlides) failures.push('slide_count_mismatch');
  if (qaResult.contentHash !== artifacts.pptx.sha256) failures.push('qa_pptx_hash_mismatch');
  if (qaResult.stages?.validate?.warningCount > state.maxWarnings) failures.push('qa_warning_threshold');
  if (qaResult.stages?.overflow?.ok !== true) failures.push('overflow');
  if (qaResult.stages?.render?.count !== state.expectedSlides) failures.push('render_count_mismatch');
  if (review.verdict !== 'pass') failures.push('visual_review_failed');
  if (!exactSlideCoverage(review.inspectedSlides, state.expectedSlides)) failures.push('visual_review_incomplete');
  if (review.inspectedMontage !== true) failures.push('montage_not_reviewed');
  if (!Array.isArray(review.issues) || review.issues.length > 0) failures.push('unresolved_visual_issues');
  if (review.querySatisfied !== true) failures.push('query_not_satisfied');
  if (review.contentConsistent !== true) failures.push('content_inconsistent');
  if (review.sourcesPresent !== true) failures.push('sources_missing');
  if (review.pptxSha256 !== artifacts.pptx.sha256) failures.push('review_pptx_hash_mismatch');
  if (review.reviewManifestSha256 !== artifacts.reviewManifest.sha256) {
    failures.push('review_manifest_hash_mismatch');
  }
  if (review.qaResultSha256 !== artifacts.qaResult.sha256) failures.push('review_qa_hash_mismatch');
  if (review.montageSha256 !== artifacts.montage.sha256) failures.push('review_montage_hash_mismatch');
  if (!Array.isArray(review.renderedSlideSha256s)
    || review.renderedSlideSha256s.length !== artifacts.renderedSlides.length
    || review.renderedSlideSha256s.some(
      (hash, index) => hash !== artifacts.renderedSlides[index].sha256,
    )) failures.push('review_render_hash_mismatch');
  const normalizedReview = {
    schemaVersion: 1,
    reviewer: review.reviewer ?? 'Codex',
    reviewedAt: review.reviewedAt ?? new Date().toISOString(),
    verdict: review.verdict,
    inspectedSlides: [...new Set(review.inspectedSlides ?? [])].sort((a, b) => a - b),
    inspectedMontage: review.inspectedMontage === true,
    querySatisfied: review.querySatisfied === true,
    contentConsistent: review.contentConsistent === true,
    sourcesPresent: review.sourcesPresent === true,
    issues: review.issues ?? [],
    notes: review.notes ?? '',
    pptxSha256: review.pptxSha256,
    reviewManifestSha256: review.reviewManifestSha256,
    qaResultSha256: review.qaResultSha256,
    montageSha256: review.montageSha256,
    renderedSlideSha256s: review.renderedSlideSha256s,
  };
  await atomicWriteJson(paths.visualReview, normalizedReview);
  artifacts.visualReview = await artifact(paths.visualReview, paths.root);
  const qualityPass = failures.length === 0;
  const verdict = await commitFinalVerdict(paths, {
    schemaVersion: 1,
    qualityPass,
    startedAt: state.startedAt,
    slaMs: state.slaMs,
    stages: {
      generationMs: state.attempts[0]?.generationMs,
      attempts: state.attempts,
      qaMs: qaResult.elapsedMs,
      visualReviewMs: state.qaCompletedMonotonicNs
        ? elapsedMs(state.qaCompletedMonotonicNs)
        : undefined,
    },
    artifacts,
    failures,
  }, state.startedMonotonicNs);
  state.status = verdict.verdict === 'pass' ? 'passed' : 'failed';
  state.completedAt = verdict.committedAt;
  state.elapsedMs = verdict.elapsedMs;
  await atomicWriteJson(paths.state, state);
  return { paths, verdict };
}

export async function finalizeAcceptanceRun(options) {
  const paths = acceptancePaths(options.runDir);
  return withRunLock(paths, () => finalizeLocked(options, paths));
}

function parseArgs(argv) {
  const result = { command: argv[0] };
  for (let index = 1; index < argv.length; index += 1) {
    const value = argv[index];
    if (value === '--run-dir') result.runDir = argv[++index];
    else if (value === '--query') result.query = argv[++index];
    else if (value === '--query-file') result.queryFile = argv[++index];
    else if (value === '--review-file') result.reviewFile = argv[++index];
    else if (value === '--expected-slides') result.expectedSlides = Number(argv[++index]);
    else if (value === '--sla-ms') result.slaMs = Number(argv[++index]);
    else if (value === '--max-warnings') result.maxWarnings = Number(argv[++index]);
    else if (value === '--force-repair') result.forceRepair = true;
    else throw new Error(`Unknown argument: ${value}`);
  }
  if (!result.runDir) throw new Error('--run-dir is required');
  return result;
}

export async function main(argv = process.argv.slice(2), invokedNs = process.hrtime.bigint()) {
  try {
    const options = parseArgs(argv);
    let result;
    if (options.command === 'begin') {
      const query = options.queryFile ? await readFile(path.resolve(options.queryFile), 'utf8') : options.query;
      result = await beginAcceptanceRun({ ...options, query, startedNs: invokedNs });
    } else if (options.command === 'build') {
      result = await buildAcceptanceRun(options);
    } else if (options.command === 'finalize') {
      if (!options.reviewFile) throw new Error('--review-file is required for finalize');
      result = await finalizeAcceptanceRun({ ...options, reviewFile: path.resolve(options.reviewFile) });
    } else {
      throw new Error('Usage: ppt-fast-accept.mjs <begin|build|finalize> --run-dir <new-directory> [options]');
    }
    const accepted = result.verdict === undefined || result.verdict.verdict === 'pass';
    process.stdout.write(`${JSON.stringify({ ok: accepted, verdict: result.verdict, runDir: result.paths.root })}\n`);
    if (!accepted) process.exitCode = 3;
  } catch (error) {
    process.stdout.write(`${JSON.stringify({ ok: false, error: String(error) })}\n`);
    process.exitCode = 1;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
