import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, LosslessXmlError } from '@pptx/lossless-xml';
import {
  readThemeFonts,
  replaceThemeFonts,
  type ThemeFontUpdate,
} from './theme-fonts.internal.js';

const DRAWINGML_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function themeXml(
  major = '<a:latin typeface="Aptos Display"/>',
  minor = '<a:latin typeface="Aptos"/>',
): string {
  return `<?xml version="1.0"?><a:theme xmlns:a="${DRAWINGML_NAMESPACE}" name="Office"><a:themeElements><a:clrScheme name="Office"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:fontScheme name="Office"><a:majorFont>${major}<a:ea typeface=""/><a:cs typeface=""/><a:font script="Hans" typeface="等线 Light"/></a:majorFont><a:minorFont>${minor}<a:ea typeface=""/><a:cs typeface=""/><a:font script="Hans" typeface="等线"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst/></a:fmtScheme></a:themeElements><!--opaque--><a:extLst><a:ext uri="opaque"/></a:extLst></a:theme>`;
}

function read(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const value = readThemeFonts(xml);
  expect(xml.changed).toBe(false);
  return value;
}

function expectRejectedWithoutPatch(source: string, update: unknown): void {
  const xml = LosslessXmlDocument.parse(source);
  expect(() => replaceThemeFonts(xml, update as ThemeFontUpdate)).toThrow();
  expect(xml.changed).toBe(false);
  expect(xml.serialize()).toBe(source);
}

describe('theme font codec', () => {
  it('reads a strict detached snapshot through canonical and alternate namespace prefixes', () => {
    const source = themeXml();
    const first = read(source);
    const second = read(source);
    expect(first).toEqual({ majorLatin: 'Aptos Display', minorLatin: 'Aptos' });
    expect(second).toEqual(first);
    expect(second).not.toBe(first);

    const alternate = `<?xml version="1.0"?><r:theme xmlns:r="${DRAWINGML_NAMESPACE}"><r:themeElements><d:fontScheme xmlns:d="${DRAWINGML_NAMESPACE}"><d:majorFont><d:latin typeface="Noto Sans Display"/></d:majorFont><d:minorFont><d:latin typeface="Noto Sans"/></d:minorFont></d:fontScheme></r:themeElements></r:theme>`;
    expect(read(alternate)).toEqual({
      majorLatin: 'Noto Sans Display',
      minorLatin: 'Noto Sans',
    });
  });

  it('does not guess malformed, ambiguous, wrong-namespace, or descendant state', () => {
    const valid = themeXml();
    const cases = [
      '',
      `${valid}${valid}`,
      valid.replace(DRAWINGML_NAMESPACE, 'urn:wrong'),
      valid.replace('<a:themeElements>', '<a:themeElements/><a:themeElements>'),
      valid.replace('<a:fontScheme name="Office">', '<a:fontScheme name="Office"/><a:fontScheme name="Office">'),
      valid.replace('<a:majorFont>', '<a:majorFont/><a:majorFont>'),
      valid.replace('<a:minorFont>', '<a:minorFont/><a:minorFont>'),
      valid.replace('<a:themeElements>', '<a:themeElements><a:wrapper>').replace('</a:themeElements>', '</a:wrapper></a:themeElements>'),
      valid.replace('<a:latin typeface="Aptos Display"/>', ''),
      valid.replace('<a:latin typeface="Aptos Display"/>', '<a:latin typeface="A"/><a:latin typeface="B"/>'),
      valid.replace('<a:latin typeface="Aptos Display"/>', '<x:latin xmlns:x="urn:wrong" typeface="Wrong"/>'),
      valid.replace('typeface="Aptos Display"', 'typeface="A" typeface="B"'),
      valid.replace('typeface="Aptos Display"', 'typeface="   "'),
      valid.replace('<a:latin typeface="Aptos Display"/>', '<a:latin typeface="Aptos Display"><a:ext/></a:latin>'),
      valid.replace('<a:latin typeface="Aptos Display"/>', '<a:latin typeface="Aptos Display">text</a:latin>'),
    ];
    for (const source of cases) expect(read(source)).toBeUndefined();
    expect(() => LosslessXmlDocument.parse('<a:theme>')).toThrow(LosslessXmlError);
  });

  it('patches either or both direct Latin typefaces and preserves opaque theme state', () => {
    const source = themeXml(
      '<a:latin typeface="Aptos Display" panose="020B" pitchFamily="34"/>',
      '<a:latin typeface="Aptos" charset="00"/>',
    );
    const minor = LosslessXmlDocument.parse(source);
    replaceThemeFonts(minor, { minorLatin: 'Noto Sans' });
    const minorOutput = minor.serialize();
    expect(minorOutput).toBe(source.replace('typeface="Aptos" charset="00"', 'typeface="Noto Sans" charset="00"'));
    expect(minorOutput).toContain('<a:latin typeface="Aptos Display" panose="020B" pitchFamily="34"/>');
    expect(minorOutput).toContain('<a:font script="Hans" typeface="等线"/>');
    expect(minorOutput).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>');
    expect(minorOutput).toContain('<!--opaque--><a:extLst>');

    const both = LosslessXmlDocument.parse(source);
    replaceThemeFonts(both, {
      majorLatin: 'A&B <Display> "One"',
      minorLatin: 'A&B <Body> "Two"',
    });
    expect(both.serialize()).toContain('typeface="A&amp;B &lt;Display&gt; &quot;One&quot;"');
    expect(both.serialize()).toContain('typeface="A&amp;B &lt;Body&gt; &quot;Two&quot;"');
    expect(read(both.serialize())).toEqual({
      majorLatin: 'A&B <Display> "One"',
      minorLatin: 'A&B <Body> "Two"',
    });
  });

  it('repairs a missing unqualified typeface without deleting a namespaced impostor', () => {
    const source = themeXml(
      '<a:latin xmlns:x="urn:foreign" x:typeface="keep" panose="020B"/>',
    );
    expect(read(source)).toBeUndefined();
    const xml = LosslessXmlDocument.parse(source);
    replaceThemeFonts(xml, { majorLatin: 'Noto Sans Display' });
    expect(xml.serialize()).toContain(
      '<a:latin xmlns:x="urn:foreign" x:typeface="keep" panose="020B" typeface="Noto Sans Display"/>',
    );
    expect(read(xml.serialize())).toEqual({
      majorLatin: 'Noto Sans Display',
      minorLatin: 'Aptos',
    });
  });

  it('keeps same-value updates as exact no-ops and detaches update objects', () => {
    const source = themeXml();
    const same = LosslessXmlDocument.parse(source);
    replaceThemeFonts(same, { majorLatin: 'Aptos Display' });
    expect(same.changed).toBe(false);
    expect(same.serialize()).toBe(source);

    const update: { majorLatin: string } = { majorLatin: 'Noto Sans Display' };
    const xml = LosslessXmlDocument.parse(source);
    replaceThemeFonts(xml, update);
    update.majorLatin = 'Changed later';
    expect(read(xml.serialize())?.majorLatin).toBe('Noto Sans Display');
  });

  it('rejects unsafe update descriptors and values before adding any patch', () => {
    const source = themeXml();
    const accessor = Object.create(null) as Record<string, unknown>;
    let getterCalls = 0;
    Object.defineProperty(accessor, 'majorLatin', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'Never';
      },
    });
    const symbol = { majorLatin: 'Aptos Display', [Symbol('extra')]: true };
    const customPrototype = Object.create({ inherited: true }) as { majorLatin?: string };
    customPrototype.majorLatin = 'Aptos Display';
    const invalid = [
      undefined,
      null,
      [],
      'Aptos',
      {},
      { majorLatin: undefined },
      { unknown: 'Aptos' },
      { majorLatin: '' },
      { majorLatin: '   ' },
      { majorLatin: 1 },
      { majorLatin: 'bad\u0000font' },
      accessor,
      symbol,
      customPrototype,
    ];
    for (const update of invalid) expectRejectedWithoutPatch(source, update);
    expect(getterCalls).toBe(0);
  });

  it('rejects unsafe theme structure atomically even when one requested side is valid', () => {
    const source = themeXml().replace(
      '<a:latin typeface="Aptos"/>',
      '<a:latin typeface="A"/><a:latin typeface="B"/>',
    );
    expectRejectedWithoutPatch(source, { majorLatin: 'Noto Sans Display' });

    const invalidUntouchedSide = themeXml().replace('typeface="Aptos"', 'typeface="   "');
    expectRejectedWithoutPatch(invalidUntouchedSide, { majorLatin: 'Noto Sans Display' });
  });
});
