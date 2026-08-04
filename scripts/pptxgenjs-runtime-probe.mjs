import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { isAbsolute, relative } from 'node:path';

const EXPECTED_PACKAGE_VERSION = '4.0.1';
const CATALOGS = Object.freeze([
  ['AlignH', 'AlignH'],
  ['AlignV', 'AlignV'],
  ['ChartType', 'ChartType'],
  ['OutputType', 'OutputType'],
  ['PlaceholderType', 'PLACEHOLDER_TYPES'],
  ['SchemeColor', 'SchemeColor'],
  ['ShapeType', 'ShapeType'],
]);
const EXPECTED_RUNTIME_MISMATCHES = Object.freeze([
  'catalog-extra:ShapeType.custGeom=custGeom',
  'missing-property:PptxGenJS.PlaceholderType',
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

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (value === null || typeof value !== 'object') return value;
  const result = {};
  for (const key of Object.keys(value).sort(compareText)) {
    result[key] = canonicalize(value[key]);
  }
  return result;
}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

function isInside(parent, child) {
  const childRelative = relative(parent, child);
  return childRelative === ''
    || (!childRelative.startsWith('..') && !isAbsolute(childRelative));
}

function declaredMembers(surface, prefix) {
  const methods = new Set();
  const properties = new Set();
  for (const atom of surface.atoms) {
    if (prefix === 'class') {
      if (atom.id.startsWith(`class:${surface.roots.presentation}#`)) {
        methods.add(atom.name);
      } else if (atom.id.startsWith(`class:${surface.roots.presentation}@property:`)) {
        properties.add(atom.name);
      }
    } else if (atom.id.startsWith('method:Slide#')) {
      methods.add(atom.name);
    } else if (atom.id.startsWith('property:Slide#')) {
      properties.add(atom.name);
    }
  }
  return {
    methods: [...methods].sort(compareText),
    properties: [...properties].sort(compareText),
  };
}

function collectMemberMismatches(instance, owner, members) {
  if (instance === null || (typeof instance !== 'object' && typeof instance !== 'function')) {
    throw new TypeError(`${owner} runtime value must be an object`);
  }
  const mismatches = [];
  for (const name of members.methods) {
    if (!(name in instance) || typeof instance[name] !== 'function') {
      mismatches.push({ kind: 'missing-method', owner, name });
    }
  }
  for (const name of members.properties) {
    if (!(name in instance)) mismatches.push({ kind: 'missing-property', owner, name });
  }
  return mismatches;
}

function mismatchId({ kind, owner, name, actual }) {
  return kind.startsWith('catalog-') && actual !== null && actual !== undefined
    ? `${kind}:${owner}.${name}=${String(actual)}`
    : `${kind}:${owner}.${name}`;
}

function verifyExpectedMismatches(mismatches) {
  const actual = new Set(mismatches.map(mismatchId));
  const expected = new Set(EXPECTED_RUNTIME_MISMATCHES);
  for (const mismatch of mismatches) {
    const id = mismatchId(mismatch);
    if (expected.has(id)) continue;
    if (mismatch.kind.startsWith('catalog-')) {
      throw new Error(
        `catalog ${mismatch.owner} differs from declaration ${mismatch.declarationName}`,
      );
    }
    const memberKind = mismatch.kind === 'missing-method' ? 'method' : 'property';
    throw new Error(`missing declared ${memberKind} ${mismatch.owner}.${mismatch.name}`);
  }
  for (const id of expected) {
    if (!actual.has(id)) throw new Error(`expected runtime mismatch ${id} was not observed`);
  }
}

function declaredCatalog(surface, declarationName) {
  const entries = surface.atoms
    .filter((atom) => atom.kind === 'union-member' && atom.owner === declarationName)
    .map((atom) => [atom.catalogKey || atom.name, atom.name])
    .sort(([left], [right]) => compareText(left, right));
  if (entries.length === 0) {
    throw new Error(`declaration catalog ${declarationName} has no reachable members`);
  }
  const result = {};
  for (const [key, value] of entries) {
    if (Object.hasOwn(result, key) && result[key] !== value) {
      throw new Error(`declaration catalog ${declarationName} has conflicting key ${key}`);
    }
    result[key] = value;
  }
  return result;
}

function runtimeCatalog(value, propertyName) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`runtime catalog ${propertyName} must be an object`);
  }
  const result = {};
  for (const key of Object.keys(value).sort(compareText)) {
    const entry = value[key];
    if (typeof entry !== 'string') {
      throw new TypeError(`runtime catalog ${propertyName}.${key} must be a string`);
    }
    result[key] = entry;
  }
  return result;
}

function collectCatalogMismatches(actual, expected, propertyName, declarationName) {
  const mismatches = [];
  for (const key of Object.keys(expected)) {
    if (!Object.hasOwn(actual, key)) {
      mismatches.push({
        kind: 'catalog-missing',
        owner: propertyName,
        name: key,
        declarationName,
        actual: null,
        expected: expected[key],
      });
    } else if (actual[key] !== expected[key]) {
      mismatches.push({
        kind: 'catalog-value',
        owner: propertyName,
        name: key,
        declarationName,
        actual: actual[key],
        expected: expected[key],
      });
    }
  }
  for (const key of Object.keys(actual)) {
    if (!Object.hasOwn(expected, key)) {
      mismatches.push({
        kind: 'catalog-extra',
        owner: propertyName,
        name: key,
        declarationName,
        actual: actual[key],
        expected: null,
      });
    }
  }
  return mismatches;
}

function defaultRuntimeLoader(packageInfo) {
  const packageRequire = createRequire(packageInfo.packageJsonPath);
  return packageRequire(packageInfo.entryPath);
}

export function hashRuntimeProbe(probe) {
  if (probe === null || typeof probe !== 'object' || Array.isArray(probe)) {
    throw new TypeError('runtime probe must be an object');
  }
  return sha256(JSON.stringify(canonicalize(probe)));
}

export async function probePptxGenJSRuntime({
  packageInfo,
  surface,
  loadRuntime = defaultRuntimeLoader,
}) {
  if (packageInfo === null || typeof packageInfo !== 'object') {
    throw new TypeError('packageInfo must be an object');
  }
  if (surface === null || typeof surface !== 'object' || !Array.isArray(surface.atoms)) {
    throw new TypeError('surface must contain declaration atoms');
  }
  if (typeof loadRuntime !== 'function') throw new TypeError('loadRuntime must be a function');
  if (packageInfo.version !== EXPECTED_PACKAGE_VERSION) {
    throw new Error(`expected package version ${EXPECTED_PACKAGE_VERSION}`);
  }
  for (const key of ['root', 'packageJsonPath', 'entryPath', 'declarationSha256']) {
    if (typeof packageInfo[key] !== 'string' || packageInfo[key].length === 0) {
      throw new TypeError(`packageInfo.${key} must be a non-empty string`);
    }
  }
  if (!isInside(packageInfo.root, packageInfo.entryPath)) {
    throw new Error('runtime entry is outside the package root');
  }
  if (!isInside(packageInfo.root, packageInfo.packageJsonPath)) {
    throw new Error('package metadata is outside the package root');
  }

  const runtimeEntrySha256 = sha256(await readFile(packageInfo.entryPath));
  const loadedRuntime = await loadRuntime(packageInfo);
  const Runtime = typeof loadedRuntime === 'function'
    ? loadedRuntime
    : loadedRuntime?.default;
  if (typeof Runtime !== 'function') {
    throw new TypeError('pptxgenjs runtime must export a constructor');
  }

  const presentation = new Runtime();
  if (presentation.version !== packageInfo.version) {
    throw new Error(
      `runtime version ${String(presentation.version)} differs from package ${packageInfo.version}`,
    );
  }
  const classMembers = declaredMembers(surface, 'class');
  const slideMembers = declaredMembers(surface, 'slide');
  const runtimeMismatches = collectMemberMismatches(
    presentation,
    surface.roots.presentation,
    classMembers,
  );

  const catalogs = {};
  for (const [propertyName, declarationName] of CATALOGS) {
    if (!(propertyName in presentation)) {
      catalogs[propertyName] = null;
      continue;
    }
    const actual = runtimeCatalog(presentation[propertyName], propertyName);
    const expected = declaredCatalog(surface, declarationName);
    runtimeMismatches.push(...collectCatalogMismatches(
      actual,
      expected,
      propertyName,
      declarationName,
    ));
    catalogs[propertyName] = actual;
  }

  const sectionResult = presentation.addSection({ title: 'Audit' });
  const slide = presentation.addSlide();
  runtimeMismatches.push(...collectMemberMismatches(slide, 'Slide', slideMembers));
  runtimeMismatches.sort((left, right) => compareText(mismatchId(left), mismatchId(right)));
  verifyExpectedMismatches(runtimeMismatches);
  const notesResult = slide.addNotes('Audit note');

  return deepFreeze({
    schemaVersion: 1,
    packageVersion: packageInfo.version,
    declarationSha256: packageInfo.declarationSha256,
    runtimeEntrySha256,
    classMembers,
    slideMembers,
    catalogs,
    runtimeMismatches,
    minimalCalls: {
      addSection: typeof sectionResult,
      addSlide: surface.roots.slide,
      addNotesReturnsSlide: notesResult === slide,
    },
  });
}
