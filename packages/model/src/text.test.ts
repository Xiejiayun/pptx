import { describe, expect, it } from 'vitest';
import {
  TEXT_ALIGNMENTS,
  TEXT_VERTICAL_ALIGNMENTS,
  type TextAlignment,
  type TextBoxVerticalAlignment,
} from './text.js';

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

describe('TEXT_VERTICAL_ALIGNMENTS', () => {
  it('publishes the complete frozen vertical alignment catalog', () => {
    expect(TEXT_VERTICAL_ALIGNMENTS).toEqual(['top', 'middle', 'bottom']);
    expect(new Set(TEXT_VERTICAL_ALIGNMENTS)).toHaveLength(3);
    expect(Object.isFrozen(TEXT_VERTICAL_ALIGNMENTS)).toBe(true);
    expect(() => {
      (TEXT_VERTICAL_ALIGNMENTS as unknown as string[]).push('top');
    }).toThrow(TypeError);
    expect(TEXT_VERTICAL_ALIGNMENTS).toEqual(['top', 'middle', 'bottom']);
  });

  it('keeps the runtime catalog and TextBoxVerticalAlignment synchronized', () => {
    const values: readonly TextBoxVerticalAlignment[] = TEXT_VERTICAL_ALIGNMENTS;
    expect(values).toBe(TEXT_VERTICAL_ALIGNMENTS);
    if (false) {
      // @ts-expect-error unknown vertical alignment is not supported
      const invalid: TextBoxVerticalAlignment = 'distributed';
      void invalid;
    }
  });
});
