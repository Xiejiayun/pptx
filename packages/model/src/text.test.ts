import { describe, expect, it } from 'vitest';
import { TEXT_ALIGNMENTS, type TextAlignment } from './text.js';

describe('TEXT_ALIGNMENTS', () => {
  it('publishes the complete frozen horizontal alignment catalog', () => {
    expect(TEXT_ALIGNMENTS).toEqual(['left', 'center', 'right', 'justify']);
    expect(new Set(TEXT_ALIGNMENTS)).toHaveLength(4);
    expect(Object.isFrozen(TEXT_ALIGNMENTS)).toBe(true);
    expect(() => {
      (TEXT_ALIGNMENTS as unknown as string[]).push('left');
    }).toThrow(TypeError);
    expect(TEXT_ALIGNMENTS).toEqual(['left', 'center', 'right', 'justify']);
  });

  it('keeps the runtime catalog and TextAlignment union synchronized', () => {
    const values: readonly TextAlignment[] = TEXT_ALIGNMENTS;
    expect(values).toBe(TEXT_ALIGNMENTS);
    if (false) {
      // @ts-expect-error unknown horizontal alignment is not supported
      const invalid: TextAlignment = 'distributed';
      void invalid;
    }
  });
});
