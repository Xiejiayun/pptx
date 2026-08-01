import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import {
  normalizeFirstSlideNumber,
  readFirstSlideNumber,
  replaceFirstSlideNumber,
} from './presentation-slide-number.internal.js';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

describe('presentation slide number state', () => {
  it('normalizes every signed Int32 boundary and rejects non-integers', () => {
    for (const value of [-2_147_483_648, -1, -0, 0, 1, 2_147_483_647]) {
      expect(normalizeFirstSlideNumber(value)).toBe(value === 0 ? 0 : value);
    }
    for (const value of [
      undefined,
      null,
      '1',
      true,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      1.5,
      -2_147_483_649,
      2_147_483_648,
    ]) {
      expect(() => normalizeFirstSlideNumber(value)).toThrow();
    }
  });

  it('reads absent, signed, and boundary direct values without mutation', () => {
    for (const [lexical, expected] of [
      [undefined, undefined],
      ['+1', 1],
      ['0', 0],
      ['-0', 0],
      ['-20', -20],
      ['-2147483648', -2_147_483_648],
      ['2147483647', 2_147_483_647],
    ] as const) {
      const source = presentationXml(lexical);
      const xml = LosslessXmlDocument.parse(source);
      expect(readFirstSlideNumber(xml)).toBe(expected);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('returns undefined for wrong namespace, duplicate roots, qualified lookalikes, and invalid lexical state', () => {
    const cases = [
      presentationXml('1').replace(P, 'urn:wrong'),
      `${presentationXml('1')}${presentationXml('2')}`,
      presentationXml(undefined).replace('<p:presentation ', '<p:presentation xmlns:x="urn:foreign" x:firstSlideNum="9" '),
      presentationXml(''),
      presentationXml(' 1'),
      presentationXml('1.5'),
      presentationXml('2147483648'),
      presentationXml('-2147483649'),
      presentationXml('1').replace('firstSlideNum="1"', 'firstSlideNum="1" firstSlideNum="2"'),
    ];
    for (const source of cases) {
      const xml = LosslessXmlDocument.parse(source);
      expect(readFirstSlideNumber(xml)).toBeUndefined();
      expect(xml.changed).toBe(false);
    }
  });

  it('creates, edits, repairs, clears, and preserves same semantic values exactly', () => {
    const absent = LosslessXmlDocument.parse(presentationXml());
    expect(replaceFirstSlideNumber(absent, 0)).toBe(true);
    expect(absent.serialize()).toContain(' firstSlideNum="0"');

    const sameSource = presentationXml('+1');
    const same = LosslessXmlDocument.parse(sameSource);
    expect(replaceFirstSlideNumber(same, 1)).toBe(false);
    expect(same.changed).toBe(false);
    expect(same.serialize()).toBe(sameSource);

    const invalid = LosslessXmlDocument.parse(presentationXml('bad'));
    expect(replaceFirstSlideNumber(invalid, -5)).toBe(true);
    expect(invalid.serialize()).toContain('firstSlideNum="-5"');

    const qualified = LosslessXmlDocument.parse(
      presentationXml('7').replace('<p:presentation ', '<p:presentation xmlns:x="urn:foreign" x:firstSlideNum="keep" '),
    );
    expect(replaceFirstSlideNumber(qualified, undefined)).toBe(true);
    expect(qualified.serialize()).not.toContain(' firstSlideNum="7"');
    expect(qualified.serialize()).toContain('x:firstSlideNum="keep"');

    const absentClear = LosslessXmlDocument.parse(presentationXml());
    expect(replaceFirstSlideNumber(absentClear, undefined)).toBe(false);
    expect(absentClear.changed).toBe(false);
  });

  it('rejects ambiguous roots and duplicate direct attributes before patching', () => {
    for (const source of [
      `${presentationXml('1')}${presentationXml('2')}`,
      presentationXml('1').replace('firstSlideNum="1"', 'firstSlideNum="1" firstSlideNum="2"'),
    ]) {
      const xml = LosslessXmlDocument.parse(source);
      expect(() => replaceFirstSlideNumber(xml, 3)).toThrow(/ambiguous/i);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });
});

function presentationXml(firstSlideNumber?: string): string {
  const attribute = firstSlideNumber === undefined
    ? ''
    : ` firstSlideNum="${firstSlideNumber}"`;
  return `<?xml version="1.0"?><p:presentation xmlns:p="${P}" data-opaque="keep"${attribute}><p:sldIdLst/><p:extLst><p:ext uri="opaque"/></p:extLst></p:presentation>`;
}
