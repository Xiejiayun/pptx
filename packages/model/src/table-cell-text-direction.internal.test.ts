import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableTextDirection,
  replaceTableTextDirection,
} from './table-cell-text-direction.internal.js';

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

describe('table-level text direction', () => {
  it('reads only one safe uniform explicit direct physical-cell value', () => {
    for (const direction of ['horz', 'vert', 'vert270', 'wordArtVert'] as const) {
      const uniform = parseFrame(
        '<a:tr h="1">' +
          `<a:tc gridSpan="2"><a:tcPr vert="${direction}"/></a:tc>` +
          `<a:tc hMerge="1"><a:tcPr vert="${direction}"/></a:tc>` +
        '</a:tr>' +
        `<a:tr h="2"><a:tc vMerge="1"><a:tcPr vert="${direction}"/></a:tc></a:tr>`,
      );
      expect(readTableTextDirection(uniform.xml, uniform.frame)).toBe(direction);
      expect(uniform.xml.changed).toBe(false);
    }

    const unsafeRows = [
      '<a:tr h="1"><a:tc><a:tcPr/></a:tc><a:tc><a:tcPr/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr vert="vert"/></a:tc>' +
        '<a:tc><a:tcPr vert="vert270"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr vert="vert"/></a:tc>' +
        '<a:tc><a:tcPr/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr vert="eaVert"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr vert="vert" vert="vert270"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr vert="vert"/><a:tcPr vert="vert"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc/></a:tr>',
      '',
      '<a:tr h="1"/>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(readTableTextDirection(fixture.xml, fixture.frame), rows).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }

    const opaque = parseFrame(
      '<x:keep xmlns:x="urn:test"><a:tr h="9"><a:tc>' +
        '<a:tcPr vert="wordArtVert"/></a:tc></a:tr></x:keep>' +
      '<a:tr h="1"><a:tc><a:tcPr vert="vert270"/></a:tc></a:tr>',
    );
    expect(readTableTextDirection(opaque.xml, opaque.frame)).toBe('vert270');
    expect(opaque.xml.changed).toBe(false);
  });

  it('returns undefined for ambiguous direct table paths', () => {
    const sources = [
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic/><a:graphic>' +
        '<a:graphicData><a:tbl><a:tr><a:tc><a:tcPr vert="vert"/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData/>' +
        '<a:graphicData><a:tbl><a:tr><a:tc><a:tcPr vert="vert"/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData>' +
        '<a:tbl><a:tr><a:tc><a:tcPr vert="vert"/></a:tc></a:tr></a:tbl>' +
        '<a:tbl><a:tr><a:tc><a:tcPr vert="vert"/></a:tc></a:tr></a:tbl>' +
        '</a:graphicData></a:graphic></p:graphicFrame>',
      '<p:sp xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData><a:tbl>' +
        '<a:tr><a:tc><a:tcPr vert="vert"/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:sp>',
    ];
    for (const source of sources) {
      const fixture = parseSource(source);
      expect(readTableTextDirection(fixture.xml, fixture.frame), source).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('replaces, writes explicit horizontal, and clears every physical cell', () => {
    const target = parseFrame(
      '<a:tr h="1">' +
        '<a:tc gridSpan="2"><a:tcPr keep="A" vert="vert"><a:solidFill/></a:tcPr></a:tc>' +
        '<a:tc hMerge="1"><a:tcPr keep="B"/></a:tc>' +
      '</a:tr>' +
      '<a:tr h="2"><a:tc vMerge="1"><a:tcPr keep="C" vert="eaVert"/></a:tc></a:tr>',
    );
    expect(replaceTableTextDirection(
      target.xml,
      target.frame,
      'wordArtVert',
      PART_URI,
    )).toBe(true);
    const stackedSource = target.xml.serialize();
    expect(stackedSource.match(/ vert="wordArtVert"/g)).toHaveLength(3);
    expect(stackedSource).toContain('gridSpan="2"');
    expect(stackedSource).toContain('hMerge="1"');
    expect(stackedSource).toContain('vMerge="1"');
    expect(stackedSource).toContain('keep="A"');
    expect(stackedSource).toContain('keep="B"');
    expect(stackedSource).toContain('keep="C"');
    expect(stackedSource).toContain('<a:solidFill/>');

    const stacked = parseSource(stackedSource);
    expect(replaceTableTextDirection(
      stacked.xml,
      stacked.frame,
      'wordArtVert',
      PART_URI,
    )).toBe(false);
    expect(stacked.xml.serialize()).toBe(stackedSource);
    expect(replaceTableTextDirection(
      stacked.xml,
      stacked.frame,
      'horz',
      PART_URI,
    )).toBe(true);
    const horizontalSource = stacked.xml.serialize();
    expect(horizontalSource.match(/ vert="horz"/g)).toHaveLength(3);

    const horizontal = parseSource(horizontalSource);
    expect(replaceTableTextDirection(
      horizontal.xml,
      horizontal.frame,
      undefined,
      PART_URI,
    )).toBe(true);
    const clearedSource = horizontal.xml.serialize();
    expect(clearedSource).not.toContain(' vert=');
    expect(clearedSource).toContain('keep="A"');
    expect(clearedSource).toContain('keep="B"');
    expect(clearedSource).toContain('keep="C"');
    expect(clearedSource).toContain('<a:solidFill/>');

    const cleared = parseSource(clearedSource);
    expect(replaceTableTextDirection(
      cleared.xml,
      cleared.frame,
      undefined,
      PART_URI,
    )).toBe(false);
    expect(cleared.xml.serialize()).toBe(clearedSource);
  });

  it('rejects unsafe edits with a stable model error', () => {
    const unsafeRows = [
      '',
      '<a:tr h="1"/>',
      '<a:tr h="1"><a:tc/></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr/><a:tcPr/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr vert="vert" vert="vert270"/></a:tc></a:tr>',
      '<a:tr h="1"><a:tc><a:tcPr/></a:tc><a:tc><a:tcPr/></a:tc>' +
        '<a:tc><a:tcPr vert="vert" vert="vert270"/></a:tc></a:tr>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(() => replaceTableTextDirection(
        fixture.xml,
        fixture.frame,
        'vert270',
        PART_URI,
      ), rows).toThrow(ModelParseError);
    }

    const ambiguous = parseSource(
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic/><a:graphic>' +
        '<a:graphicData><a:tbl><a:tr><a:tc><a:tcPr/></a:tc></a:tr>' +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
    );
    expect(() => replaceTableTextDirection(
      ambiguous.xml,
      ambiguous.frame,
      'vert',
      PART_URI,
    )).toThrow(/one complete set of direct physical cells/);

    const invalid = parseFrame(
      '<a:tr h="1"><a:tc><a:tcPr keep="INVALID" vert="eaVert"/></a:tc></a:tr>',
    );
    expect(replaceTableTextDirection(
      invalid.xml,
      invalid.frame,
      'vert270',
      PART_URI,
    )).toBe(true);
    expect(invalid.xml.serialize()).toContain('keep="INVALID" vert="vert270"');
  });
});
