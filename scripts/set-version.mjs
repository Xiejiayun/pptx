import { readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const repositoryRoot = resolve(
  process.env.PPTX_RELEASE_ROOT ?? resolve(dirname(fileURLToPath(import.meta.url)), '..'),
);
const manifestPaths = [
  'package.json',
  'packages/cli/package.json',
  'packages/codecs/package.json',
  'packages/lossless-xml/package.json',
  'packages/model/package.json',
  'packages/opc/package.json',
  'packages/pptx/package.json',
  'packages/pptxgenjs-adapter/package.json',
  'packages/sdk/package.json',
  'packages/testkit/package.json',
  'packages/validator/package.json',
  'plugins/advanced-charts/package.json',
  'plugins/animations/package.json',
  'plugins/smartart/package.json',
  'plugins/transitions/package.json',
];
const versionSourcePath = 'packages/sdk/src/version.ts';
const publicManifestPath = 'packages/pptx/package.json';
const versionReferences = [
  ['packages/sdk/src/index.test.ts', (version) => [`expect(current).toBe('${version}');`]],
  ['packages/pptx/src/index.test.ts', (version) => [`expect(current).toBe('${version}');`]],
  ['scripts/playwright-browser-smoke.js', (version) => [
    `constant: '${version}',`,
    `created: '${version}',`,
    `reopened: '${version}',`,
  ]],
  ['scripts/pptxgenjs-surface-manifest.mjs', (version) => [
    `pattern: "export const PPTX_VERSION = '${version}' as const;"`,
  ]],
  ['docs/compatibility/pptxgenjs-surface-audit.json', (version) => [
    `"pattern": "export const PPTX_VERSION = '${version}' as const;"`,
  ]],
];
const semverPattern = /^\d+\.\d+\.\d+$/u;

const manifests = await Promise.all(manifestPaths.map(async (path) => ({
  path,
  contents: JSON.parse(await readFile(resolve(repositoryRoot, path), 'utf8')),
})));

if (process.argv[2] === '--check') {
  const source = await readFile(resolve(repositoryRoot, versionSourcePath), 'utf8');
  const sourceVersion = source.match(/PPTX_VERSION = '([^']+)'/u)?.[1];
  const expected = manifests[0].contents.version;
  const drift = manifests.filter(({ contents }) => contents.version !== expected);
  const scopeDrift = manifests.filter(({ path, contents }) => (
    path === publicManifestPath ? contents.private === true : contents.private !== true
  ));
  const publicManifest = manifests.find(({ path }) => path === publicManifestPath)?.contents;
  const publicConfigDrift = publicManifest?.name !== '@jiayunxie/pptx'
    || publicManifest?.publishConfig?.access !== 'public'
    || publicManifest?.publishConfig?.tag !== 'next'
    || publicManifest?.publishConfig?.provenance !== true;
  const referenceDrift = [];
  for (const [path, expectedReferences] of versionReferences) {
    const source = await readFile(resolve(repositoryRoot, path), 'utf8');
    if (expectedReferences(expected).some((reference) => !source.includes(reference))) referenceDrift.push(path);
  }
  if (!semverPattern.test(expected) || sourceVersion !== expected || drift.length > 0 || scopeDrift.length > 0 || publicConfigDrift || referenceDrift.length > 0) {
    throw new Error(`Release drift: expected ${expected}; source ${sourceVersion}; version manifests ${drift.map(({ path }) => path).join(', ')}; publish scope ${scopeDrift.map(({ path }) => path).join(', ')}; public config ${publicConfigDrift}; references ${referenceDrift.join(', ')}`);
  }
  process.stdout.write(`${expected}\n`);
  process.exit(0);
}

const version = process.argv[2];
if (!version || !semverPattern.test(version)) {
  throw new Error('Usage: node scripts/set-version.mjs <major.minor.patch>');
}

const previousVersion = manifests[0].contents.version;
const referenceUpdates = await Promise.all(versionReferences.map(async ([path, expectedReferences]) => {
  let source = await readFile(resolve(repositoryRoot, path), 'utf8');
  for (const previousReference of expectedReferences(previousVersion)) {
    if (!source.includes(previousReference)) throw new Error(`Version reference drift: ${path}`);
  }
  const nextReferences = expectedReferences(version);
  expectedReferences(previousVersion).forEach((previousReference, index) => {
    source = source.replaceAll(previousReference, nextReferences[index]);
  });
  return { path, source };
}));

await Promise.all(manifests.map(async ({ path, contents }) => {
  contents.version = version;
  await writeFile(resolve(repositoryRoot, path), `${JSON.stringify(contents, null, 2)}\n`);
}));
await Promise.all(referenceUpdates.map(({ path, source }) => (
  writeFile(resolve(repositoryRoot, path), source)
)));
await writeFile(
  resolve(repositoryRoot, versionSourcePath),
  `export const PPTX_VERSION = '${version}' as const;\n\nexport type PptxVersion = typeof PPTX_VERSION;\n`,
);
process.stdout.write(`Updated ${manifestPaths.length} manifests and ${versionSourcePath} to ${version}\n`);
