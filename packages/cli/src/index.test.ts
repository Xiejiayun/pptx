import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { PPTX_VERSION, PptxDocument } from '@pptx/sdk';
import { createMinimalPptx } from '@pptx/testkit';
import { runCli } from './index.js';

async function run(args: string[]): Promise<{ code: number; stdout: string; stderr: string }> {
  let stdout = '';
  let stderr = '';
  const code = await runCli(args, {
    stdout: (text) => { stdout += text; },
    stderr: (text) => { stderr += text; },
  });
  return { code, stdout, stderr };
}

describe('pptx-inspect CLI', () => {
  it('returns stable JSON for doctor and package inspection', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-cli-'));
    const file = join(directory, 'fixture.pptx');
    await writeFile(file, await createMinimalPptx('CLI fixture'));
    const doctor = await run(['--json', 'doctor']);
    expect(JSON.parse(doctor.stdout)).toMatchObject({
      ok: true,
      command: 'doctor',
      data: { version: PPTX_VERSION, auth: { required: false } },
    });
    const version = await run(['--version']);
    expect(version).toEqual({ code: 0, stdout: `${PPTX_VERSION}\n`, stderr: '' });
    const inspection = await run(['--json', 'package', 'inspect', file]);
    expect(JSON.parse(inspection.stdout)).toMatchObject({ ok: true, command: 'package.inspect', data: { partCount: 5 } });
  });

  it('supports dry-run and explicit-output title edits', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-cli-'));
    const input = join(directory, 'input.pptx');
    const output = join(directory, 'output.pptx');
    await writeFile(input, await createMinimalPptx('Before'));
    const dryRun = await run(['--json', 'slides', 'set-title', input, '--slide', '1', '--text', 'After', '--dry-run']);
    expect(JSON.parse(dryRun.stdout)).toMatchObject({ ok: true, data: { dryRun: true, before: 'Before', after: 'After' } });
    const written = await run(['--json', 'slides', 'set-title', input, '--slide', '1', '--text', 'After', '--out', output]);
    expect(written.code).toBe(0);
    expect((await PptxDocument.open(new Uint8Array(await readFile(output)))).slides[0]?.title.text).toBe('After');
  });

  it('returns machine-readable errors', async () => {
    const result = await run(['--json', 'slides', 'list', '/missing/file.pptx']);
    expect(result.code).toBe(1);
    expect(JSON.parse(result.stdout)).toMatchObject({ ok: false, error: { code: 'CLI_ERROR' } });
  });
});
