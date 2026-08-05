import assert from 'node:assert/strict';
import { cp, mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const fixturePaths = [
  'package.json',
  'packages/cli/package.json',
  'packages/codecs/package.json',
  'packages/lossless-xml/package.json',
  'packages/model/package.json',
  'packages/opc/package.json',
  'packages/pptx/package.json',
  'packages/pptx/src/index.test.ts',
  'packages/pptxgenjs-adapter/package.json',
  'packages/sdk/package.json',
  'packages/sdk/src/index.test.ts',
  'packages/sdk/src/version.ts',
  'packages/testkit/package.json',
  'packages/validator/package.json',
  'plugins/advanced-charts/package.json',
  'plugins/animations/package.json',
  'plugins/smartart/package.json',
  'plugins/transitions/package.json',
  'scripts/playwright-browser-smoke.js',
  'scripts/pptxgenjs-surface-manifest.mjs',
  'docs/compatibility/pptxgenjs-surface-audit.json',
];

function run(root, ...args) {
  return spawnSync(process.execPath, [join(repositoryRoot, 'scripts/set-version.mjs'), ...args], {
    cwd: root,
    encoding: 'utf8',
    env: { ...process.env, PPTX_RELEASE_ROOT: root },
  });
}

test('checks release scope and updates every executable version reference idempotently', async () => {
  const root = await mkdtemp(join(tmpdir(), 'pptx-version-'));
  try {
    for (const path of fixturePaths) {
      await mkdir(dirname(join(root, path)), { recursive: true });
      await cp(join(repositoryRoot, path), join(root, path));
    }

    assert.equal(run(root, '--check').status, 0);
    const cliPath = join(root, 'packages/cli/package.json');
    const cliManifest = JSON.parse(await readFile(cliPath, 'utf8'));
    delete cliManifest.private;
    await writeFile(cliPath, `${JSON.stringify(cliManifest, null, 2)}\n`);
    const scopeFailure = run(root, '--check');
    assert.notEqual(scopeFailure.status, 0);
    assert.match(scopeFailure.stderr, /publish scope/u);
    cliManifest.private = true;
    await writeFile(cliPath, `${JSON.stringify(cliManifest, null, 2)}\n`);

    assert.notEqual(run(root, '1.2').status, 0);
    assert.equal(run(root, '9.8.7').status, 0);
    assert.equal(run(root, '--check').status, 0);
    const first = await Promise.all(fixturePaths.map((path) => readFile(join(root, path), 'utf8')));
    assert.equal(run(root, '9.8.7').status, 0);
    const second = await Promise.all(fixturePaths.map((path) => readFile(join(root, path), 'utf8')));
    assert.deepEqual(second, first);
    assert.match(await readFile(join(root, 'packages/sdk/src/version.ts'), 'utf8'), /PPTX_VERSION = '9\.8\.7'/u);
    assert.match(await readFile(join(root, 'scripts/playwright-browser-smoke.js'), 'utf8'), /constant: '9\.8\.7'/u);
    assert.equal(JSON.parse(await readFile(join(root, 'packages/pptx/package.json'), 'utf8')).name, '@jiayunxie/pptx');
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
