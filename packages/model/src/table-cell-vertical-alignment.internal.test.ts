import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableVerticalAlignment,
  replaceTableVerticalAlignment,
} from './table-cell-vertical-alignment.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';

function parseFrame(rows: string): {
  xml: LosslessXmlDocument;
  frame: XmlElement;
} {
  return parseSource(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a">' +
      '<a:graphic><a:graphicData><a:tbl>' + rows +
      '</a:tbl></a:graphicData></a:graphic>' +
      '</p:graphicFrame>',
  );
}

function parseSource(source: string): {
  xml: LosslessXmlDocument;
  frame: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(source);
  const frame = xml.roots[0];
  if (!frame) throw new Error('Fixture has no root frame');
  return { xml, frame };
}

describe('table-level vertical alignment', () => {
  it('reads only one safe uniform direct physical-cell value', () => {
    for (const [token, expected] of [
      ['t', 'top'],
      ['ctr', 'middle'],
      ['b', 'bottom'],
    ] as const) {
      const uniform = parseFrame(
        '<a:tr h="1">' +
          `<a:tc gridSpan="2"><a:tcPr anchor="${token}"/></a:tc>` +
          `<a:tc hMerge="1"><a:tcPr anchor="${token}"/></a:tc>` +
        '</a:tr>' +
        `<a:tr h="2"><a:tc vMerge="1"><a:tcPr anchor="${token}"/></a:tc></a:tr>`,
      );
      expect(readTableVerticalAlignment(uniform.xml, uniform.frame)).toBe(expected);
      expect(uniform.xml.changed).toBe(false);
    }

    const unsafeRows = [
      '<a:tr h="1"><a:tc><a:tcPr/></a:tc><a:tc><a:tcPr/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr anchor="t"/></a:tc>' +
        '<a:tc><a:tcPr anchor="b"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr anchor="ctr"/></a:tc>' +
        '<a:tc><a:tcPr/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr anchor="mid"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr anchor="ctr" anchor="b"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr anchor="ctr"/><a:tcPr anchor="ctr"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc/></a:tr>',
      '',
      '<a:tr h="1"/>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(readTableVerticalAlignment(fixture.xml, fixture.frame), rows)
        .toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }

    const opaque = parseFrame(
      '<x:keep xmlns:x="urn:test"><a:tr h="9"><a:tc><a:tcPr anchor="b"/></a:tc></a:tr></x:keep>' +
      '<a:tr h="1"><a:tc><a:tcPr anchor="ctr"/></a:tc></a:tr>',
    );
    expect(readTableVerticalAlignment(opaque.xml, opaque.frame)).toBe('middle');
  });

  it('returns undefined for ambiguous direct table paths', () => {
    const sources = [
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic/><a:graphic>' +
        '<a:graphicData><a:tbl><a:tr><a:tc><a:tcPr anchor="t"/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData/>' +
        '<a:graphicData><a:tbl><a:tr><a:tc><a:tcPr anchor="t"/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData>' +
        '<a:tbl><a:tr><a:tc><a:tcPr anchor="t"/></a:tc></a:tr></a:tbl>' +
        '<a:tbl><a:tr><a:tc><a:tcPr anchor="t"/></a:tc></a:tr></a:tbl>' +
        '</a:graphicData></a:graphic></p:graphicFrame>',
      '<p:sp xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData><a:tbl>' +
        '<a:tr><a:tc><a:tcPr anchor="t"/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:sp>',
    ];
    for (const source of sources) {
      const fixture = parseSource(source);
      expect(readTableVerticalAlignment(fixture.xml, fixture.frame), source)
        .toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('replaces and clears every physical cell while preserving unrelated XML', () => {
    const target = parseFrame(
      '<a:tr h="1">' +
        '<a:tc gridSpan="2"><a:tcPr keep="A" anchor="t"><a:solidFill/></a:tcPr></a:tc>' +
        '<a:tc hMerge="1"><a:tcPr keep="B"/></a:tc>' +
      '</a:tr>' +
      '<a:tr h="2"><a:tc vMerge="1"><a:tcPr keep="C" anchor="b"/></a:tc></a:tr>',
    );
    expect(replaceTableVerticalAlignment(
      target.xml,
      target.frame,
      'middle',
      PART_URI,
    )).toBe(true);
    expect(target.xml.serialize().match(/ anchor="ctr"/g)).toHaveLength(3);
    expect(target.xml.serialize()).toContain('gridSpan="2"');
    expect(target.xml.serialize()).toContain('hMerge="1"');
    expect(target.xml.serialize()).toContain('vMerge="1"');
    expect(target.xml.serialize()).toContain('keep="A"');
    expect(target.xml.serialize()).toContain('keep="B"');
    expect(target.xml.serialize()).toContain('keep="C"');
    expect(target.xml.serialize()).toContain('<a:solidFill/>');
    const afterMiddle = target.xml.serialize();
    const middle = parseSource(afterMiddle);
    expect(replaceTableVerticalAlignment(
      middle.xml,
      middle.frame,
      'middle',
      PART_URI,
    )).toBe(false);
    expect(middle.xml.serialize()).toBe(afterMiddle);
    expect(replaceTableVerticalAlignment(
      middle.xml,
      middle.frame,
      undefined,
      PART_URI,
    )).toBe(true);
    expect(middle.xml.serialize()).not.toContain(' anchor=');

    const invalid = parseFrame(
      '<a:tr h="1"><a:tc><a:tcPr keep="INVALID" anchor="mid"/></a:tc></a:tr>',
    );
    expect(replaceTableVerticalAlignment(
      invalid.xml,
      invalid.frame,
      'bottom',
      PART_URI,
    )).toBe(true);
    expect(invalid.xml.serialize()).toContain('keep="INVALID" anchor="b"');
    const normalizedInvalid = parseSource(invalid.xml.serialize());
    expect(replaceTableVerticalAlignment(
      normalizedInvalid.xml,
      normalizedInvalid.frame,
      undefined,
      PART_URI,
    )).toBe(true);
    expect(normalizedInvalid.xml.serialize()).toContain('keep="INVALID"');
    expect(normalizedInvalid.xml.serialize()).not.toContain(' anchor=');
  });

  it('rejects unsafe edits with a stable model error', () => {
    const unsafeRows = [
      '',
      '<a:tr h="1"/>',
      '<a:tr h="1"><a:tc/></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr/><a:tcPr/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr anchor="t" anchor="b"/></a:tc></a:tr>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(() => replaceTableVerticalAlignment(
        fixture.xml,
        fixture.frame,
        'middle',
        PART_URI,
      ), rows).toThrow(ModelParseError);
    }

    const ambiguous = parseSource(
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic/><a:graphic>' +
        '<a:graphicData><a:tbl><a:tr><a:tc><a:tcPr/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
    );
    expect(() => replaceTableVerticalAlignment(
      ambiguous.xml,
      ambiguous.frame,
      'top',
      PART_URI,
    )).toThrow(/one complete set of direct physical cells/);
  });
});
