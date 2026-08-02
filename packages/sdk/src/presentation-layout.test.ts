import { describe, expect, it } from 'vitest';
import { inches } from '@pptx/model';
import { presentationLayoutFromSlideSize } from './presentation-layout.js';

describe('presentationLayoutFromSlideSize', () => {
  it.each([
    [inches(10), inches(7.5), 'screen4x3'],
    [inches(10), inches(5.625), 'screen16x9'],
    [inches(10), inches(6.25), 'screen16x10'],
    [inches(13 + 1 / 3), inches(7.5), 'custom'],
    [inches(11.7), inches(8.3), 'custom'],
  ] as const)('maps %s × %s to %s', (width, height, name) => {
    expect(presentationLayoutFromSlideSize({ width, height })).toEqual({
      name,
      width,
      height,
    });
  });

  it('uses exact dimensions and returns detached snapshots', () => {
    const input = { width: inches(10), height: inches(5.625) };
    const first = presentationLayoutFromSlideSize(input);
    const second = presentationLayoutFromSlideSize(input);

    expect(first).toEqual(second);
    expect(first).not.toBe(second);
    expect(presentationLayoutFromSlideSize({
      width: input.width,
      height: (input.height + 1) as typeof input.height,
    })).toEqual({
      name: 'custom',
      width: input.width,
      height: input.height + 1,
    });
  });
});
