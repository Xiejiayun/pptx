import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableBorders,
  replaceTableBorders,
} from './table-cell-borders.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';

function none(tag: string, width = '0'): string {
  return `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr">` +
    `<a:noFill/></a:${tag}>`;
}

function line(
  tag: string,
  color = '4472C4',
  width = '12700',
  dash?: 'solid' | 'sysDash',
): string {
  return `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    `${dash ? `<a:prstDash val="${dash}"/>` : ''}` +
    '<a:round/><a:headEnd type="none" w="med" len="med"/>' +
    `<a:tailEnd type="none" w="med" len="med"/></a:${tag}>`;
}

function schemeLine(
  tag: string,
  color = 'accent1',
  width = '25400',
  dash?: 'solid' | 'sysDash',
): string {
  return `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr">` +
    `<a:solidFill><a:schemeClr val="${color}"/></a:solidFill>` +
    `${dash ? `<a:prstDash val="${dash}"/>` : ''}` +
    '<a:round/><a:headEnd type="none" w="med" len="med"/>' +
    `<a:tailEnd type="none" w="med" len="med"/></a:${tag}>`;
}

function cell(
  borders = '',
  extraProperties = '',
  cellAttributes = '',
  propertiesAttributes = '',
): string {
  return `<a:tc${cellAttributes}><a:txBody><a:bodyPr lIns="11">` +
    '<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:bodyPr>' +
    '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"><a:ln w="12700">' +
    '<a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></a:rPr>' +
    '<a:t>Keep</a:t></a:r></a:p></a:txBody>' +
    `<a:tcPr anchor="ctr" vert="vert270" marL="50800" keep="yes"` +
    `${propertiesAttributes}>${borders}${extraProperties}` +
    '<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>' +
    '<a:extLst><a:ext uri="keep"/></a:extLst></a:tcPr></a:tc>';
}

function parseFrame(rows: string): {
  xml: LosslessXmlDocument;
  frame: XmlElement;
} {
  return parseSource(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a"><p:xfrm><a:off x="1" y="2"/>' +
      '<a:ext cx="3" cy="4"/></p:xfrm><a:graphic><a:graphicData>' +
      '<a:tbl><a:tblPr firstRow="1"><a:tableStyleId>style</a:tableStyleId></a:tblPr>' +
      '<a:tblGrid><a:gridCol w="10"/><a:gridCol w="20"/></a:tblGrid>' + rows +
      '</a:tbl></a:graphicData></a:graphic><p:extLst><p:ext uri="frame"/>' +
      '</p:extLst></p:graphicFrame>',
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

describe('table-level borders', () => {
  it('reads complete and partial uniform direct physical-cell vectors', () => {
    const complete = [
      line('lnL', '112233', '0'),
      none('lnR', '7'),
      line('lnT', '4472c4', '12700'),
      schemeLine('lnB', 'accent1', '25400', 'sysDash'),
    ].join('');
    const lexicalEquivalent = [
      line('lnL', '112233', '00'),
      none('lnR', '0007'),
      line('lnT', '4472C4', '012700'),
      schemeLine('lnB', 'accent1', '025400', 'sysDash'),
    ].join('');
    const fixture = parseFrame(
      '<a:tr h="30">' +
        cell(complete, '', ' gridSpan="2"') +
        cell(lexicalEquivalent, '', ' hMerge="1"') +
      '</a:tr>' +
      '<a:tr h="40">' +
        cell(complete, '', ' vMerge="1"') +
        cell(lexicalEquivalent) +
      '</a:tr>',
    );
    expect(readTableBorders(fixture.xml, fixture.frame)).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: '4472C4' },
        width: 1,
      },
      right: { kind: 'none' },
      bottom: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 2,
        style: 'dash',
      },
      left: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 0,
      },
    });
    expect(fixture.xml.changed).toBe(false);

    const partial = parseFrame(
      '<a:tr>' +
        cell(line('lnL', 'ABCDEF', '20116800') + none('lnT')) +
        cell(line('lnL', 'abcdef', '20116800') + none('lnT', '9')) +
      '</a:tr>',
    );
    expect(readTableBorders(partial.xml, partial.frame)).toEqual({
      top: { kind: 'none' },
      left: {
        kind: 'line',
        color: { kind: 'srgb', value: 'ABCDEF' },
        width: 1584,
      },
    });
    expect(partial.xml.changed).toBe(false);
  });

  it('returns undefined for absent, mixed, malformed, advanced, or unsafe state', () => {
    const top = line('lnT');
    const unsafeRows = [
      `<a:tr>${cell()}${cell()}</a:tr>`,
      `<a:tr>${cell(top)}${cell()}</a:tr>`,
      `<a:tr>${cell(top)}${cell(line('lnL'))}</a:tr>`,
      `<a:tr>${cell(top)}${cell(line('lnT', 'FF0000'))}</a:tr>`,
      `<a:tr>${cell(line('lnT'))}${cell(line('lnT', '4472C4', '12700', 'solid'))}</a:tr>`,
      `<a:tr>${cell(none('lnT'))}${cell(line('lnT'))}</a:tr>`,
      `<a:tr>${cell('<a:lnT w="invalid"><a:noFill/></a:lnT>')}` +
        `${cell('<a:lnT w="invalid"><a:noFill/></a:lnT>')}</a:tr>`,
      `<a:tr>${cell('<a:lnT w="12700"><a:gradFill/></a:lnT>')}` +
        `${cell('<a:lnT w="12700"><a:gradFill/></a:lnT>')}</a:tr>`,
      `<a:tr>${cell(`${top}${top}`)}${cell(top)}</a:tr>`,
      '<a:tr><a:tc><a:txBody/><a:tcPr/><a:tcPr/></a:tc></a:tr>',
      '<a:tr><a:tc><a:txBody/></a:tc></a:tr>',
      '',
      '<a:tr/>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(readTableBorders(fixture.xml, fixture.frame), rows).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('rejects ambiguous table paths and ignores diagonals, descendants, and wrong prefixes', () => {
    const table = `<a:tbl><a:tr>${cell(line('lnT'))}</a:tr></a:tbl>`;
    const ambiguousSources = [
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic/><a:graphic>' +
        `<a:graphicData>${table}</a:graphicData></a:graphic></p:graphicFrame>`,
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData/>' +
        `<a:graphicData>${table}</a:graphicData></a:graphic></p:graphicFrame>`,
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData>' +
        `${table}${table}</a:graphicData></a:graphic></p:graphicFrame>`,
      '<p:sp xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData>' +
        `${table}</a:graphicData></a:graphic></p:sp>`,
    ];
    for (const source of ambiguousSources) {
      const fixture = parseSource(source);
      expect(readTableBorders(fixture.xml, fixture.frame), source).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }

    const lookalikes = '<a:lnTlToBr w="12700"><a:solidFill>' +
      '<a:srgbClr val="FF0000"/></a:solidFill></a:lnTlToBr>' +
      '<x:lnT xmlns:x="urn:test" w="0"><x:noFill/></x:lnT>' +
      '<x:keep xmlns:x="urn:test"><a:lnT w="0"><a:noFill/></a:lnT></x:keep>';
    const fixture = parseFrame(
      `<a:tr>${cell(line('lnT'), lookalikes)}${cell(line('lnT'), lookalikes)}</a:tr>`,
    );
    expect(readTableBorders(fixture.xml, fixture.frame)).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: '4472C4' },
        width: 1,
      },
    });
    expect(fixture.xml.changed).toBe(false);
  });

  it('whole-replaces, clears, repairs, and preserves unrelated XML', () => {
    const diagonal = '<a:lnTlToBr w="12700"><a:solidFill>' +
      '<a:srgbClr val="FF0000"/></a:solidFill></a:lnTlToBr>';
    const target = parseFrame(
      '<a:tr h="30">' +
        cell(line('lnL') + none('lnR') + line('lnT') + none('lnB'), diagonal,
          ' gridSpan="2"', ' one="1"') +
        cell('<a:lnL w="12700"><a:gradFill/></a:lnL>', diagonal,
          ' hMerge="1"', ' two="2"') +
      '</a:tr>' +
      '<a:tr h="40">' +
        cell('', diagonal, ' vMerge="1"', ' three="3"') +
        cell(none('lnL') + none('lnR') + none('lnT') + none('lnB'), diagonal,
          '', ' four="4"') +
      '</a:tr>',
    );
    const partial = {
      top: {
        kind: 'line' as const,
        color: { kind: 'scheme' as const, value: 'accent2' as const },
        width: 1.5,
        style: 'dash' as const,
      },
      bottom: { kind: 'none' as const },
    };
    expect(replaceTableBorders(target.xml, target.frame, partial, PART_URI)).toBe(true);
    const partialSource = target.xml.serialize();
    const partialFixture = parseSource(partialSource);
    expect(readTableBorders(partialFixture.xml, partialFixture.frame)).toEqual(partial);
    expect(partialSource.match(/<a:lnT(?:\s|>)/g)).toHaveLength(4);
    expect(partialSource.match(/<a:lnB(?:\s|>)/g)).toHaveLength(4);
    expect([...partialSource.matchAll(/<a:lnL(?:\s|>)/g)]).toHaveLength(0);
    expect([...partialSource.matchAll(/<a:lnR(?:\s|>)/g)]).toHaveLength(0);
    expect(partialSource.match(/<a:lnTlToBr\b/g)).toHaveLength(4);
    expect(partialSource.match(/<a:prstDash val="sysDash"\/>/g)).toHaveLength(4);
    expect(partialSource.match(/<a:solidFill><a:srgbClr val="D9EAF7"\/>/g))
      .toHaveLength(4);
    for (const retained of [
      'gridSpan="2"',
      'hMerge="1"',
      'vMerge="1"',
      'anchor="ctr"',
      'vert="vert270"',
      'marL="50800"',
      'one="1"',
      'two="2"',
      'three="3"',
      'four="4"',
      '<a:bodyPr lIns="11">',
      '<a:pPr algn="ctr"/>',
      '<a:rPr lang="en-US">',
      '<a:tblGrid><a:gridCol w="10"/><a:gridCol w="20"/></a:tblGrid>',
      '<p:ext uri="frame"/>',
    ]) expect(partialSource).toContain(retained);
    expect(partialSource).toMatch(
      /<a:lnT[\s\S]*<a:lnB[\s\S]*<a:lnTlToBr[\s\S]*<a:solidFill/,
    );

    expect(replaceTableBorders(
      partialFixture.xml,
      partialFixture.frame,
      partial,
      PART_URI,
    )).toBe(false);
    expect(partialFixture.xml.serialize()).toBe(partialSource);

    const fullNone = {
      top: { kind: 'none' as const },
      right: { kind: 'none' as const },
      bottom: { kind: 'none' as const },
      left: { kind: 'none' as const },
    };
    expect(replaceTableBorders(
      partialFixture.xml,
      partialFixture.frame,
      fullNone,
      PART_URI,
    )).toBe(true);
    const noneFixture = parseSource(partialFixture.xml.serialize());
    expect(readTableBorders(noneFixture.xml, noneFixture.frame)).toEqual(fullNone);
    expect(noneFixture.xml.serialize().match(/<a:noFill\/>/g)).toHaveLength(16);

    expect(replaceTableBorders(
      noneFixture.xml,
      noneFixture.frame,
      undefined,
      PART_URI,
    )).toBe(true);
    const clearedSource = noneFixture.xml.serialize();
    const clearedFixture = parseSource(clearedSource);
    expect(readTableBorders(clearedFixture.xml, clearedFixture.frame)).toBeUndefined();
    expect([...clearedSource.matchAll(/<a:lnL(?:\s|>)/g)]).toHaveLength(0);
    expect([...clearedSource.matchAll(/<a:lnR(?:\s|>)/g)]).toHaveLength(0);
    expect([...clearedSource.matchAll(/<a:lnT(?:\s|>)/g)]).toHaveLength(0);
    expect([...clearedSource.matchAll(/<a:lnB(?:\s|>)/g)]).toHaveLength(0);
    expect(clearedSource.match(/<a:lnTlToBr\b/g)).toHaveLength(4);
    expect(replaceTableBorders(
      clearedFixture.xml,
      clearedFixture.frame,
      undefined,
      PART_URI,
    )).toBe(false);
    expect(clearedFixture.xml.serialize()).toBe(clearedSource);
  });

  it('rejects unsafe bulk edits and preserves wrong-prefix lookalikes', () => {
    const missingTable = parseFrame('');
    expect(() => replaceTableBorders(
      missingTable.xml,
      missingTable.frame,
      { top: { kind: 'none' } },
      PART_URI,
    )).toThrow(ModelParseError);

    const repeatedFinal = parseFrame(
      '<a:tr>' + cell(none('lnT')) + cell(`${none('lnL')}${none('lnL')}`) + '</a:tr>',
    );
    expect(() => replaceTableBorders(
      repeatedFinal.xml,
      repeatedFinal.frame,
      { top: { kind: 'none' } },
      PART_URI,
    )).toThrow('Table cell contains repeated direct lnL elements');

    const wrongPrefix = parseFrame(
      '<a:tr>' + cell('', '<x:lnT xmlns:x="urn:test" w="0"><x:noFill/></x:lnT>') +
        '</a:tr>',
    );
    expect(replaceTableBorders(
      wrongPrefix.xml,
      wrongPrefix.frame,
      { top: { kind: 'none' } },
      PART_URI,
    )).toBe(true);
    const reopenedWrongPrefix = parseSource(wrongPrefix.xml.serialize());
    expect(readTableBorders(reopenedWrongPrefix.xml, reopenedWrongPrefix.frame)).toEqual({
      top: { kind: 'none' },
    });
    expect(reopenedWrongPrefix.xml.serialize())
      .toContain('<x:lnT xmlns:x="urn:test" w="0"><x:noFill/></x:lnT>');
  });
});
