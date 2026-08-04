import { execFile } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { isAbsolute, relative, resolve } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const EXPECTED_SCHEMA_VERSION = 1;
const EXPECTED_PACKAGE_VERSION = '4.0.1';
const EVIDENCE_CATEGORIES = Object.freeze([
  'code',
  'tests',
  'package',
  'ooxml',
  'clients',
]);
const MANIFEST_STATUSES = new Set([
  'supported',
  'deliberate-difference',
  'deprecated-alias',
  'defect-excluded',
  'unsupported',
  'unverified',
]);
const ENTRY_KEYS = new Set([
  'id',
  'status',
  'native',
  'evidence',
  'note',
  'serialization',
  'client',
  'canonical',
  'control',
]);

function compareText(left, right) {
  return left < right ? -1 : left > right ? 1 : 0;
}

function deepFreeze(value, seen = new Set()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return value;
  seen.add(value);
  for (const nested of Object.values(value)) deepFreeze(nested, seen);
  return Object.freeze(value);
}

function readDataObject(value, allowedKeys, context) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be a plain data object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be a plain data object`);
  }
  if (Object.getOwnPropertySymbols(value).length > 0) {
    throw new TypeError(`${context} must not contain symbol keys`);
  }
  const result = {};
  for (const key of Object.getOwnPropertyNames(value)) {
    if (!allowedKeys.has(key)) throw new TypeError(`${context} has unknown key ${key}`);
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context}.${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function nonEmptyString(value, context) {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new TypeError(`${context} must be a non-empty string`);
  }
  return value;
}

function normalizeStringArray(value, context) {
  if (!Array.isArray(value)) throw new TypeError(`${context} must be an array`);
  const values = value.map((entry, index) => nonEmptyString(entry, `${context}[${index}]`));
  if (new Set(values).size !== values.length) throw new TypeError(`${context} has duplicates`);
  return values.sort(compareText);
}

function normalizeEvidenceLink(value, category, context) {
  const field = category === 'tests' ? 'title' : 'pattern';
  const data = readDataObject(value, new Set(['path', field, 'commit']), context);
  const link = {
    path: nonEmptyString(data.path, `${context}.path`),
    [field]: nonEmptyString(data[field], `${context}.${field}`),
  };
  if (data.commit !== undefined) {
    const commit = nonEmptyString(data.commit, `${context}.commit`);
    if (!/^[0-9a-f]{7,40}$/iu.test(commit)) {
      throw new TypeError(`${context}.commit must be a 7-40 digit hexadecimal commit ID`);
    }
    link.commit = commit.toLowerCase();
  }
  return link;
}

function normalizeEvidence(value, context) {
  const data = readDataObject(value, new Set(EVIDENCE_CATEGORIES), context);
  const evidence = {};
  for (const category of EVIDENCE_CATEGORIES) {
    if (!Array.isArray(data[category])) {
      throw new TypeError(`${context}.${category} must be an array`);
    }
    evidence[category] = data[category].map((link, index) => normalizeEvidenceLink(
      link,
      category,
      `${context}.${category}[${index}]`,
    ));
  }
  return evidence;
}

function normalizeControl(value, context) {
  if (value === undefined) return undefined;
  return normalizeEvidenceLink(value, 'control', context);
}

function normalizeEntry(value, index) {
  const context = `manifest.entries[${index}]`;
  const data = readDataObject(value, ENTRY_KEYS, context);
  const id = nonEmptyString(data.id, `${context}.id`);
  const status = nonEmptyString(data.status, `${context}.status`);
  if (!MANIFEST_STATUSES.has(status)) throw new TypeError(`${context} has invalid status ${status}`);
  if (data.serialization !== undefined && typeof data.serialization !== 'boolean') {
    throw new TypeError(`${context}.serialization must be a boolean`);
  }
  if (data.client !== undefined && typeof data.client !== 'boolean') {
    throw new TypeError(`${context}.client must be a boolean`);
  }
  return {
    id,
    status,
    native: normalizeStringArray(data.native, `${context}.native`),
    evidence: normalizeEvidence(data.evidence, `${context}.evidence`),
    note: typeof data.note === 'string' ? data.note : '',
    serialization: data.serialization ?? false,
    client: data.client ?? false,
    canonical: data.canonical === undefined
      ? undefined
      : nonEmptyString(data.canonical, `${context}.canonical`),
    control: normalizeControl(data.control, `${context}.control`),
  };
}

function normalizeExtension(value, index) {
  const context = `manifest.extensions[${index}]`;
  const data = readDataObject(
    value,
    new Set(['id', 'native', 'evidence', 'note']),
    context,
  );
  const id = nonEmptyString(data.id, `${context}.id`);
  if (!id.startsWith('extension:')) {
    throw new TypeError(`${context}.id must start with extension:`);
  }
  const native = normalizeStringArray(data.native, `${context}.native`);
  if (native.length === 0) throw new TypeError(`${context}.native must not be empty`);
  return {
    id,
    native,
    evidence: normalizeEvidence(data.evidence, `${context}.evidence`),
    note: nonEmptyString(data.note, `${context}.note`),
  };
}

function normalizeManifest(value) {
  const data = readDataObject(
    value,
    new Set(['schemaVersion', 'packageVersion', 'entries', 'extensions']),
    'manifest',
  );
  if (data.schemaVersion !== EXPECTED_SCHEMA_VERSION) {
    throw new TypeError(`manifest schemaVersion must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (data.packageVersion !== EXPECTED_PACKAGE_VERSION) {
    throw new TypeError(`manifest packageVersion must be ${EXPECTED_PACKAGE_VERSION}`);
  }
  if (!Array.isArray(data.entries)) throw new TypeError('manifest.entries must be an array');
  if (!Array.isArray(data.extensions)) throw new TypeError('manifest.extensions must be an array');

  const entries = data.entries.map(normalizeEntry);
  const entryIds = new Set();
  for (const entry of entries) {
    if (entryIds.has(entry.id)) throw new TypeError(`duplicate manifest entry ${entry.id}`);
    entryIds.add(entry.id);
  }
  const extensions = data.extensions.map(normalizeExtension);
  const extensionIds = new Set();
  for (const extension of extensions) {
    if (extensionIds.has(extension.id)) {
      throw new TypeError(`duplicate manifest extension ${extension.id}`);
    }
    extensionIds.add(extension.id);
  }
  return { entries, extensions };
}

function addDiagnostic(diagnostics, id, code, message) {
  diagnostics.push({ id, code, message });
}

function requireEvidence(entry, category, code, diagnostics) {
  if (entry.evidence[category].length === 0) {
    addDiagnostic(
      diagnostics,
      entry.id,
      code,
      `${entry.status} requires ${category} evidence`,
    );
  }
}

function validateStatusEvidence(entry, declarationIds, entryById, diagnostics) {
  if (entry.status === 'unverified') return;
  if (entry.note.trim().length === 0) {
    addDiagnostic(diagnostics, entry.id, 'missing-note', `${entry.status} requires a note`);
  }
  if (
    entry.status === 'supported'
    || entry.status === 'deliberate-difference'
    || entry.status === 'deprecated-alias'
  ) {
    if (entry.native.length === 0) {
      addDiagnostic(diagnostics, entry.id, 'missing-native', `${entry.status} requires native mapping`);
    }
    requireEvidence(entry, 'code', 'missing-code-evidence', diagnostics);
    requireEvidence(entry, 'tests', 'missing-test-evidence', diagnostics);
    requireEvidence(entry, 'package', 'missing-package-evidence', diagnostics);
    if (entry.serialization) {
      requireEvidence(entry, 'ooxml', 'missing-ooxml-evidence', diagnostics);
    }
    if (entry.client) {
      requireEvidence(entry, 'clients', 'missing-client-evidence', diagnostics);
    }
  }
  if (entry.status === 'defect-excluded') {
    requireEvidence(entry, 'tests', 'missing-test-evidence', diagnostics);
  }
  if (
    entry.status === 'deliberate-difference'
    || entry.status === 'deprecated-alias'
    || entry.status === 'defect-excluded'
    || entry.status === 'unsupported'
  ) {
    if (!entry.control) {
      addDiagnostic(diagnostics, entry.id, 'missing-control', `${entry.status} requires a control`);
    }
  }
  if (entry.status === 'deprecated-alias') {
    if (
      !entry.canonical
      || entry.canonical === entry.id
      || !declarationIds.has(entry.canonical)
    ) {
      addDiagnostic(
        diagnostics,
        entry.id,
        'invalid-canonical',
        'deprecated-alias requires another declaration atom as canonical target',
      );
    } else {
      const canonicalStatus = entryById.get(entry.canonical)?.status ?? 'unverified';
      if (!['supported', 'deliberate-difference'].includes(canonicalStatus)) {
        addDiagnostic(
          diagnostics,
          entry.id,
          'canonical-unclosed',
          `canonical target ${entry.canonical} is ${canonicalStatus}`,
        );
      }
    }
  }
}

function isInside(parent, child) {
  const childRelative = relative(parent, child);
  return childRelative === ''
    || (!childRelative.startsWith('..') && !isAbsolute(childRelative));
}

function isValidEvidencePath(path, repositoryRoot) {
  if (isAbsolute(path) || path.includes('\\')) return false;
  return isInside(repositoryRoot, resolve(repositoryRoot, path));
}

async function defaultGitCommitExists(repositoryRoot, commit) {
  try {
    await execFileAsync(
      'git',
      ['-C', repositoryRoot, 'cat-file', '-e', `${commit}^{commit}`],
      { encoding: 'utf8' },
    );
    return true;
  } catch {
    return false;
  }
}

async function verifyEvidenceLink({
  link,
  category,
  atomId,
  repositoryRoot,
  diagnostics,
  fileCache,
  commitCache,
  gitCommitExists,
}) {
  if (!isValidEvidencePath(link.path, repositoryRoot)) {
    addDiagnostic(
      diagnostics,
      atomId,
      'evidence-path-invalid',
      `${category} evidence path must stay repository-relative`,
    );
  } else {
    const absolutePath = resolve(repositoryRoot, link.path);
    let file = fileCache.get(absolutePath);
    if (!file) {
      file = readFile(absolutePath, 'utf8')
        .then((content) => ({ content }))
        .catch((error) => ({ error }));
      fileCache.set(absolutePath, file);
    }
    const result = await file;
    if (result.error) {
      addDiagnostic(
        diagnostics,
        atomId,
        result.error.code === 'ENOENT' ? 'evidence-file-missing' : 'evidence-file-unreadable',
        `${category} evidence file ${link.path} cannot be read`,
      );
    } else {
      const field = category === 'tests' ? 'title' : 'pattern';
      if (!result.content.includes(link[field])) {
        addDiagnostic(
          diagnostics,
          atomId,
          category === 'tests' ? 'evidence-title-missing' : 'evidence-pattern-missing',
          `${category} evidence literal is absent from ${link.path}`,
        );
      }
    }
  }

  if (link.commit) {
    let exists = commitCache.get(link.commit);
    if (!exists) {
      exists = Promise.resolve(gitCommitExists(link.commit)).catch(() => false);
      commitCache.set(link.commit, exists);
    }
    if (!(await exists)) {
      addDiagnostic(
        diagnostics,
        atomId,
        'evidence-commit-missing',
        `${category} evidence commit ${link.commit} is unavailable`,
      );
    }
  }
}

async function verifyEntryEvidence(entry, context) {
  for (const category of EVIDENCE_CATEGORIES) {
    for (const link of entry.evidence[category]) {
      await verifyEvidenceLink({ ...context, link, category, atomId: entry.id });
    }
  }
  if (entry.control) {
    await verifyEvidenceLink({
      ...context,
      link: entry.control,
      category: 'control',
      atomId: entry.id,
    });
  }
}

function reportLink(link, repositoryRoot) {
  return isValidEvidencePath(link.path, repositoryRoot)
    ? { ...link }
    : { ...link, path: '<invalid>' };
}

function reportEvidence(entry, repositoryRoot) {
  const evidence = {};
  for (const category of EVIDENCE_CATEGORIES) {
    evidence[category] = entry.evidence[category].map((link) => reportLink(link, repositoryRoot));
  }
  return evidence;
}

function reportEntry(entry, repositoryRoot) {
  return {
    ...entry,
    evidence: reportEvidence(entry, repositoryRoot),
    control: entry.control ? reportLink(entry.control, repositoryRoot) : undefined,
  };
}

function validateAuditInputs(surface, runtimeProbe) {
  if (surface === null || typeof surface !== 'object' || !Array.isArray(surface.atoms)) {
    throw new TypeError('surface must contain declaration atoms');
  }
  if (surface.schemaVersion !== EXPECTED_SCHEMA_VERSION) {
    throw new TypeError(`surface schemaVersion must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (Array.isArray(surface.diagnostics) && surface.diagnostics.length > 0) {
    throw new TypeError('surface must not contain declaration diagnostics');
  }
  if (runtimeProbe === null || typeof runtimeProbe !== 'object') {
    throw new TypeError('runtimeProbe must be an object');
  }
  if (runtimeProbe.schemaVersion !== EXPECTED_SCHEMA_VERSION) {
    throw new TypeError(`runtimeProbe schemaVersion must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (runtimeProbe.packageVersion !== EXPECTED_PACKAGE_VERSION) {
    throw new TypeError(`runtimeProbe packageVersion must be ${EXPECTED_PACKAGE_VERSION}`);
  }
  for (const hash of ['declarationSha256', 'runtimeEntrySha256']) {
    if (typeof runtimeProbe[hash] !== 'string' || !/^[0-9a-f]{64}$/u.test(runtimeProbe[hash])) {
      throw new TypeError(`runtimeProbe ${hash} must be a lowercase SHA-256`);
    }
  }
  const declarationIds = new Set();
  for (const atom of surface.atoms) {
    if (atom === null || typeof atom !== 'object' || typeof atom.id !== 'string') {
      throw new TypeError('surface atoms must have string IDs');
    }
    if (declarationIds.has(atom.id)) throw new TypeError(`duplicate declaration atom ${atom.id}`);
    declarationIds.add(atom.id);
  }
  return declarationIds;
}

export async function buildPptxGenJSAudit({
  surface,
  runtimeProbe,
  manifest,
  repositoryRoot,
  gitCommitExists,
}) {
  if (typeof repositoryRoot !== 'string' || repositoryRoot.length === 0) {
    throw new TypeError('repositoryRoot must be a non-empty string');
  }
  const root = resolve(repositoryRoot);
  const declarationIds = validateAuditInputs(surface, runtimeProbe);
  const normalized = normalizeManifest(manifest);
  const entryById = new Map(normalized.entries.map((entry) => [entry.id, entry]));

  for (const extension of normalized.extensions) {
    if (declarationIds.has(extension.id)) {
      throw new TypeError(`extension ID collides with declaration ${extension.id}`);
    }
    if (entryById.has(extension.id)) {
      throw new TypeError(`extension ID collides with manifest entry ${extension.id}`);
    }
  }

  const diagnostics = [];
  for (const entry of normalized.entries) {
    validateStatusEvidence(entry, declarationIds, entryById, diagnostics);
  }

  const commitChecker = gitCommitExists === undefined
    ? (commit) => defaultGitCommitExists(root, commit)
    : gitCommitExists;
  if (typeof commitChecker !== 'function') throw new TypeError('gitCommitExists must be a function');
  const evidenceContext = {
    repositoryRoot: root,
    diagnostics,
    fileCache: new Map(),
    commitCache: new Map(),
    gitCommitExists: commitChecker,
  };
  for (const entry of normalized.entries) {
    await verifyEntryEvidence(entry, evidenceContext);
  }
  for (const extension of normalized.extensions) {
    await verifyEntryEvidence(extension, evidenceContext);
  }

  const counts = {
    supported: 0,
    'deliberate-difference': 0,
    'deprecated-alias': 0,
    'defect-excluded': 0,
    unsupported: 0,
    unverified: 0,
    stale: 0,
  };
  const atomResults = surface.atoms.map((atom) => {
    const entry = entryById.get(atom.id);
    const status = entry?.status ?? 'unverified';
    counts[status] += 1;
    return {
      ...atom,
      status,
      native: entry?.native ?? [],
      evidence: entry ? reportEvidence(entry, root) : emptyReportEvidence(),
      note: entry?.note ?? '',
      serialization: entry?.serialization ?? false,
      client: entry?.client ?? false,
      canonical: entry?.canonical,
      control: entry?.control ? reportLink(entry.control, root) : undefined,
    };
  });
  const staleEntries = normalized.entries
    .filter((entry) => !declarationIds.has(entry.id))
    .map((entry) => ({
      ...reportEntry(entry, root),
      manifestStatus: entry.status,
      status: 'stale',
    }));
  counts.stale = staleEntries.length;

  diagnostics.sort((left, right) => (
    compareText(left.id, right.id)
    || compareText(left.code, right.code)
    || compareText(left.message, right.message)
  ));
  const incompleteIds = new Set([
    ...atomResults
      .filter(({ status }) => status === 'unsupported' || status === 'unverified')
      .map(({ id }) => id),
    ...staleEntries.map(({ id }) => id),
    ...diagnostics.map(({ id }) => id),
  ]);
  const sortedIncompleteIds = [...incompleteIds].sort(compareText);
  const complete = diagnostics.length === 0
    && counts.unsupported === 0
    && counts.unverified === 0
    && counts.stale === 0;

  return deepFreeze({
    schemaVersion: EXPECTED_SCHEMA_VERSION,
    packageVersion: runtimeProbe.packageVersion,
    declarationSha256: runtimeProbe.declarationSha256,
    runtimeEntrySha256: runtimeProbe.runtimeEntrySha256,
    declarationTotal: surface.atoms.length,
    counts,
    complete,
    incompleteIds: sortedIncompleteIds,
    diagnostics,
    atoms: atomResults,
    staleEntries,
    extensions: normalized.extensions.map((extension) => reportEntry(extension, root)),
    runtimeMismatches: runtimeProbe.runtimeMismatches ?? [],
  });
}

function emptyReportEvidence() {
  return {
    code: [],
    tests: [],
    package: [],
    ooxml: [],
    clients: [],
  };
}
