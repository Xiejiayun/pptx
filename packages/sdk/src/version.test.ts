import { readFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { PPTX_VERSION } from './version.js';

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('PPTX_VERSION', () => {
  it('matches every public package manifest', async () => {
    const versions = await Promise.all([
      'package.json',
      'packages/sdk/package.json',
      'packages/pptx/package.json',
    ].map(async (path) => JSON.parse(
      await readFile(resolve(repositoryRoot, path), 'utf8'),
    ).version));

    expect(PPTX_VERSION).toMatch(/^\d+\.\d+\.\d+$/u);
    expect(versions).toEqual([PPTX_VERSION, PPTX_VERSION, PPTX_VERSION]);
  });
});
