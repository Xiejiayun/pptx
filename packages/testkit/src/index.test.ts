import { describe, expect, it } from 'vitest';
import { PptxDocument } from '@pptx/sdk';
import { assertMutationIsolation, createMinimalPptx, diffPackages, inspectPackage } from './index.js';

describe('testkit', () => {
  it('inspects package payloads and enforces mutation isolation', async () => {
    const before = await createMinimalPptx('Before');
    const document = await PptxDocument.open(before);
    document.slides[0]!.title.text = 'After';
    const after = await document.write();
    const inspection = await inspectPackage(after);
    expect(inspection.partCount).toBeGreaterThan(3);
    const diff = await diffPackages(before, after);
    expect(diff.changed.map(({ after: part }) => part.uri)).toEqual(['/ppt/slides/slide1.xml']);
    expect(() => assertMutationIsolation(diff, ['/ppt/slides/slide1.xml'])).not.toThrow();
    expect(() => assertMutationIsolation(diff, [])).toThrow(/Unexpected/);
  });
});

