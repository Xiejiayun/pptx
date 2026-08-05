import { readFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { PPTX_VERSION } from './version.js';

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('PPTX_VERSION', () => {
  it('matches every workspace package manifest', async () => {
    const versions = await Promise.all([
      'package.json',
      'packages/cli/package.json',
      'packages/codecs/package.json',
      'packages/lossless-xml/package.json',
      'packages/model/package.json',
      'packages/opc/package.json',
      'packages/sdk/package.json',
      'packages/pptx/package.json',
      'packages/pptxgenjs-adapter/package.json',
      'packages/testkit/package.json',
      'packages/validator/package.json',
      'plugins/advanced-charts/package.json',
      'plugins/animations/package.json',
      'plugins/smartart/package.json',
      'plugins/transitions/package.json',
    ].map(async (path) => JSON.parse(
      await readFile(resolve(repositoryRoot, path), 'utf8'),
    ).version));

    expect(PPTX_VERSION).toMatch(/^\d+\.\d+\.\d+$/u);
    expect(new Set(versions)).toEqual(new Set([PPTX_VERSION]));
  });
});
