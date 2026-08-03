import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableHorizontalAlignment,
  replaceTableHorizontalAlignment,
} from './table-cell-horizontal-alignment.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';

function cell(paragraph: string, attributes = ''): string {
  return `<a:tc${attributes}><a:txBody><a:bodyPr/>${paragraph}</a:txBody>` +
    '<a:tcPr anchor="ctr" vert="vert270" keep="CELL"><a:solidFill/></a:tcPr></a:tc>';
}

function paragraph(alignment?: string, extra = ''): string {
  const properties = alignment === undefined
    ? ''
    : `<a:pPr algn="${alignment}" keep="PPR"/>`;
  return `<a:p keep="P">${properties}<a:r><a:rPr lang="en-US"/>` +
    `<a:t>${extra || alignment || 'Absent'}</a:t></a:r><a:endParaRPr/></a:p>`;
}

function parseFrame(rows: string): {
  xml: LosslessXmlDocument;
  frame: XmlElement;
} {
  return parseSource(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData>' +
      '<a:tbl><a:tblGrid><a:gridCol w="1"/></a:tblGrid>' + rows +
      '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
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

describe('table-level horizontal alignment', () => {
  it('reads only one safe uniform direct single-paragraph value', () => {
    for (const [token, value] of [
      ['l', 'left'],
      ['ctr', 'center'],
      ['r', 'right'],
      ['just', 'justify'],
    ] as const) {
      const fixture = parseFrame(
        `<a:tr h="1">${cell(paragraph(token), ' gridSpan="2"')}` +
          `${cell(paragraph(token), ' hMerge="1"')}</a:tr>` +
        `<a:tr h="2">${cell(paragraph(token), ' vMerge="1"')}</a:tr>`,
      );
      expect(readTableHorizontalAlignment(fixture.xml, fixture.frame)).toBe(value);
      expect(fixture.xml.changed).toBe(false);
    }

    const unsafeRows = [
      `<a:tr>${cell(paragraph())}${cell(paragraph())}</a:tr>`,
      `<a:tr>${cell(paragraph('l'))}${cell(paragraph('r'))}</a:tr>`,
      `<a:tr>${cell(paragraph('l'))}${cell(paragraph())}</a:tr>`,
      `<a:tr>${cell(paragraph('dist'))}</a:tr>`,
      `<a:tr>${cell('<a:p><a:pPr algn="l" algn="r"/><a:endParaRPr/></a:p>')}</a:tr>`,
      `<a:tr>${cell('<a:p><a:pPr algn="l"/><a:pPr algn="l"/></a:p>')}</a:tr>`,
      `<a:tr>${cell('<a:p><a:pPr algn="l"/></a:p><a:p><a:pPr algn="l"/></a:p>')}</a:tr>`,
      '<a:tr><a:tc><a:tcPr/></a:tc></a:tr>',
      '<a:tr><a:tc><a:txBody/><a:txBody/><a:tcPr/></a:tc></a:tr>',
      '',
      '<a:tr/>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(readTableHorizontalAlignment(fixture.xml, fixture.frame), rows).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }

    const opaque = parseFrame(
      '<x:keep xmlns:x="urn:test"><a:tr><a:tc><a:txBody>' +
        '<a:p><a:pPr algn="just"/></a:p></a:txBody></a:tc></a:tr></x:keep>' +
      `<a:tr>${cell(paragraph('ctr'))}</a:tr>`,
    );
    expect(readTableHorizontalAlignment(opaque.xml, opaque.frame)).toBe('center');
    expect(opaque.xml.changed).toBe(false);
  });

  it('returns undefined for ambiguous direct table paths', () => {
    const table = `<a:tbl><a:tr>${cell(paragraph('ctr'))}</a:tr></a:tbl>`;
    const sources = [
      `<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic/><a:graphic>` +
        `<a:graphicData>${table}</a:graphicData></a:graphic></p:graphicFrame>`,
      `<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData/>` +
        `<a:graphicData>${table}</a:graphicData></a:graphic></p:graphicFrame>`,
      `<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData>` +
        `${table}${table}</a:graphicData></a:graphic></p:graphicFrame>`,
      `<p:sp xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData>${table}` +
        '</a:graphicData></a:graphic></p:sp>',
    ];
    for (const source of sources) {
      const fixture = parseSource(source);
      expect(readTableHorizontalAlignment(fixture.xml, fixture.frame), source).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('replaces, writes explicit left, inserts properties, and clears every cell', () => {
    const target = parseFrame(
      '<a:tr h="1">' +
        cell(paragraph('ctr', 'First'), ' gridSpan="2"') +
        cell(paragraph(undefined, 'Second'), ' hMerge="1"') +
      '</a:tr>' +
      '<a:tr h="2">' +
        cell(paragraph('dist', 'Third'), ' vMerge="1"') +
      '</a:tr>',
    );
    const original = target.xml.serialize();
    expect(replaceTableHorizontalAlignment(
      target.xml,
      target.frame,
      'justify',
      PART_URI,
    )).toBe(true);
    const justifiedSource = target.xml.serialize();
    expect(justifiedSource.match(/<a:pPr\b[^>]* algn="just"/g)).toHaveLength(3);
    expect(justifiedSource).toContain('<a:pPr algn="just"/><a:r>');
    for (const retained of [
      'gridSpan="2"',
      'hMerge="1"',
      'vMerge="1"',
      'anchor="ctr"',
      'vert="vert270"',
      'keep="CELL"',
      '<a:solidFill/>',
      '<a:rPr lang="en-US"/>',
      '<a:endParaRPr/>',
    ]) expect(justifiedSource).toContain(retained);
    expect(justifiedSource).not.toBe(original);

    const justified = parseSource(justifiedSource);
    expect(replaceTableHorizontalAlignment(
      justified.xml,
      justified.frame,
      'justify',
      PART_URI,
    )).toBe(false);
    expect(justified.xml.serialize()).toBe(justifiedSource);
    expect(replaceTableHorizontalAlignment(
      justified.xml,
      justified.frame,
      'left',
      PART_URI,
    )).toBe(true);
    const leftSource = justified.xml.serialize();
    expect(leftSource.match(/<a:pPr\b[^>]* algn="l"/g)).toHaveLength(3);

    const left = parseSource(leftSource);
    expect(replaceTableHorizontalAlignment(
      left.xml,
      left.frame,
      undefined,
      PART_URI,
    )).toBe(true);
    const clearedSource = left.xml.serialize();
    expect(clearedSource).not.toMatch(/<a:pPr\b[^>]* algn=/);
    const cleared = parseSource(clearedSource);
    expect(replaceTableHorizontalAlignment(
      cleared.xml,
      cleared.frame,
      undefined,
      PART_URI,
    )).toBe(false);
    expect(cleared.xml.serialize()).toBe(clearedSource);
  });

  it('rejects unsafe edits while allowing supported repair and clear', () => {
    const unsafeRows = [
      '',
      '<a:tr/>',
      '<a:tr><a:tc><a:tcPr/></a:tc></a:tr>',
      '<a:tr><a:tc><a:txBody/><a:txBody/><a:tcPr/></a:tc></a:tr>',
      `<a:tr>${cell('<a:p/><a:p/>')}</a:tr>`,
      `<a:tr>${cell('<a:p><a:pPr/><a:pPr/></a:p>')}</a:tr>`,
      `<a:tr>${cell('<a:p><a:pPr algn="l" algn="r"/></a:p>')}</a:tr>`,
      `<a:tr>${cell(paragraph('ctr'))}${cell('<a:p><a:pPr algn="l" algn="r"/></a:p>')}</a:tr>`,
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(() => replaceTableHorizontalAlignment(
        fixture.xml,
        fixture.frame,
        'right',
        PART_URI,
      ), rows).toThrow(ModelParseError);
    }

    const invalid = parseFrame(`<a:tr>${cell(paragraph('dist'))}</a:tr>`);
    expect(replaceTableHorizontalAlignment(
      invalid.xml,
      invalid.frame,
      'center',
      PART_URI,
    )).toBe(true);
    const repairedSource = invalid.xml.serialize();
    expect(repairedSource).toContain('algn="ctr"');
    const repaired = parseSource(repairedSource);
    expect(replaceTableHorizontalAlignment(
      repaired.xml,
      repaired.frame,
      undefined,
      PART_URI,
    )).toBe(true);
    expect(repaired.xml.serialize()).not.toMatch(/<a:pPr\b[^>]* algn=/);
  });
});
