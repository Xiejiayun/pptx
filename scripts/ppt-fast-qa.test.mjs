import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import {
  parseOverflow,
  parseValidateJson,
  validateRenderedSlides,
} from './ppt-fast-qa.mjs';

test('validate parser treats valid:false as failure even when the process exits zero', () => {
  const result = parseValidateJson(JSON.stringify({
    ok: true,
    data: { valid: false, errorCount: 1, warningCount: 0 },
  }));
  assert.equal(result.ok, false);
  assert.equal(result.errorCount, 1);
});

test('validate parser enforces the configured warning threshold', () => {
  const text = JSON.stringify({
    ok: true,
    data: { valid: true, errorCount: 0, warningCount: 1 },
  });
  assert.equal(parseValidateJson(text, 0).ok, false);
  assert.equal(parseValidateJson(text, 1).ok, true);
});

test('overflow parser treats ERROR output as failure even when the process exits zero', () => {
  const result = parseOverflow('ERROR: Slide 7: Text overflow detected.');
  assert.equal(result.ok, false);
  assert.deepEqual(result.failingSlides, [7]);
  assert.equal(parseOverflow('Test passed. No overflow detected.').ok, true);
});

test('render validation rejects missing slides and an empty montage', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'ppt-fast-qa-test-'));
  const renderDir = path.join(directory, 'rendered');
  await mkdir(renderDir);
  await writeFile(path.join(renderDir, 'slide-1.png'), Buffer.from([137, 80, 78, 71]));
  await assert.rejects(validateRenderedSlides(renderDir, 2), /Expected 2 rendered slides/);
});
