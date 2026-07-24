import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createMinimalPptx, renderWithLibreOffice } from '../packages/testkit/dist/index.js';
import { PptxDocument } from '../packages/sdk/dist/index.js';

const directory = await mkdtemp(join(tmpdir(), 'pptx-libreoffice-'));
const input = join(directory, 'input.pptx');
const output = join(directory, 'output.pptx');
await writeFile(input, await createMinimalPptx('LibreOffice CI'));
const document = await PptxDocument.open(input);
document.slides[0].title.text = 'LibreOffice round-trip';
await document.writeFile(output);
const pdf = await renderWithLibreOffice(output, { outputDirectory: directory });
process.stdout.write(`${JSON.stringify({ ok: true, input, output, pdf })}\n`);

