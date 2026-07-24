import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { createMinimalPptx } from './index.js';
import { OpcPackage } from '@pptx/opc';

describe('deterministic fuzz smoke tests', () => {
  it('round-trips randomized unknown XML without changing bytes', () => {
    const random = lcg(0x5eed1234);
    for (let iteration = 0; iteration < 250; iteration += 1) {
      const attributes = Array.from({ length: 1 + Math.floor(random() * 8) }, (_, index) =>
        ` x:a${index}="${token(random, 3 + Math.floor(random() * 12))}"`,
      ).join('');
      const source = `<?xml version="1.0"?><p:root xmlns:p="p" xmlns:x="urn:fuzz"${attributes}>\n  <x:opaque>${token(
        random,
        20,
      )}</x:opaque><p:t>value-${iteration}</p:t>\n</p:root>`;
      expect(LosslessXmlDocument.parse(source).serialize()).toBe(source);
    }
  });

  it('survives repeated package open/write cycles without mutations', async () => {
    let bytes = await createMinimalPptx('Fuzz package');
    for (let iteration = 0; iteration < 20; iteration += 1) {
      const pkg = await OpcPackage.open(bytes);
      const output = await pkg.write();
      expect(output).toEqual(bytes);
      bytes = output;
    }
  });
});

function lcg(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(1_664_525, state) + 1_013_904_223) >>> 0;
    return state / 0x1_0000_0000;
  };
}

function token(random: () => number, length: number): string {
  const alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-';
  return Array.from({ length }, () => alphabet[Math.floor(random() * alphabet.length)]).join('');
}

