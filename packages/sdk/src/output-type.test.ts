import { describe, expect, it } from 'vitest';
import { OUTPUT_TYPES, type OutputType } from './output-type.js';

describe('OUTPUT_TYPES', () => {
  it('publishes the complete frozen output type catalog', () => {
    expect(OUTPUT_TYPES).toEqual([
      'arraybuffer',
      'base64',
      'binarystring',
      'blob',
      'nodebuffer',
      'uint8array',
    ]);
    expect(new Set(OUTPUT_TYPES)).toHaveLength(6);
    expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);
    expect(() => {
      (OUTPUT_TYPES as unknown as string[]).push('uint8array');
    }).toThrow(TypeError);
    expect(OUTPUT_TYPES).toEqual([
      'arraybuffer',
      'base64',
      'binarystring',
      'blob',
      'nodebuffer',
      'uint8array',
    ]);
  });

  it('keeps the runtime catalog and OutputType synchronized', () => {
    const values: readonly OutputType[] = OUTPUT_TYPES;
    expect(values).toBe(OUTPUT_TYPES);
    if (false) {
      // @ts-expect-error STREAM is handled by the separate stream API
      const stream: OutputType = 'STREAM';
      // @ts-expect-error unknown output type is not supported
      const unknown: OutputType = 'buffer';
      void [stream, unknown];
    }
  });
});
