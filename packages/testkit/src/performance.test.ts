import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';

describe('performance budgets', () => {
  it.skipIf(process.env.RUN_PERF !== '1')('opens a 1,000-part package within the smoke budget', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
    for (let index = 0; index < 1_000; index += 1) zip.file(`ppt/fuzz/part${index}.xml`, `<x index="${index}"/>`);
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    const started = performance.now();
    const pkg = await OpcPackage.open(bytes);
    const elapsed = performance.now() - started;
    expect(pkg.parts).toHaveLength(1_001);
    expect(elapsed).toBeLessThan(5_000);
  }, 15_000);
});

