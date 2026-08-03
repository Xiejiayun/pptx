import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableFill,
  replaceTableFill,
} from './table-cell-fill.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';
const SCHEME_FILL = '<a:solidFill><a:schemeClr val="accent1">' +
  '<a:alpha val="75000"/></a:schemeClr></a:solidFill>';

function cell(
  fill = '',
  extraProperties = '',
  cellAttributes = '',
  propertiesAttributes = '',
): string {
  return `<a:tc${cellAttributes}><a:txBody><a:bodyPr lIns="11">` +
    '<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:bodyPr>' +
    '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"><a:solidFill>' +
    '<a:srgbClr val="123456"/></a:solidFill></a:rPr><a:t>Keep</a:t></a:r></a:p>' +
    `</a:txBody><a:tcPr anchor="ctr" vert="vert270" marL="50800"` +
    ` marR="25400" marT="12700" marB="38100" keep="yes"${propertiesAttributes}>` +
    '<a:lnL w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill>' +
    '<a:prstDash val="solid"/></a:lnL><a:lnR><a:noFill/></a:lnR>' +
    `<a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB>${fill}${extraProperties}` +
    '</a:tcPr></a:tc>';
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
      '</a:tbl></a:graphicData></a:graphic><p:extLst><p:ext uri="keep"/></p:extLst>' +
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

describe('table-level fill', () => {
  it('reads uniform supported direct physical-cell fills', () => {
    const scheme = parseFrame(
      '<a:tr h="30">' +
        cell(SCHEME_FILL, '', ' gridSpan="2"') +
        cell(SCHEME_FILL, '', ' hMerge="1"') +
      '</a:tr>' +
      '<a:tr h="40">' +
        cell(SCHEME_FILL, '', ' vMerge="1"') +
        cell(SCHEME_FILL) +
      '</a:tr>',
    );
    expect(readTableFill(scheme.xml, scheme.frame)).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(scheme.xml.changed).toBe(false);

    const none = parseFrame(
      `<a:tr>${cell('<a:noFill/>')}${cell('<a:noFill/>')}</a:tr>`,
    );
    expect(readTableFill(none.xml, none.frame)).toEqual({ kind: 'none' });
    expect(none.xml.changed).toBe(false);

    const canonicalColor = parseFrame(
      '<a:tr>' +
        cell('<a:solidFill><a:srgbClr val="d9eaf7"/></a:solidFill>') +
        cell('<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>') +
      '</a:tr>',
    );
    expect(readTableFill(canonicalColor.xml, canonicalColor.frame)).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'D9EAF7' },
    });

    const explicitZero = parseFrame(
      '<a:tr>' +
        cell('<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="100000"/>' +
          '</a:srgbClr></a:solidFill>') +
        cell('<a:solidFill><a:srgbClr val="00ff00"><a:alpha val="100000"/>' +
          '</a:srgbClr></a:solidFill>') +
      '</a:tr>',
    );
    expect(readTableFill(explicitZero.xml, explicitZero.frame)).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
      transparency: 0,
    });
  });

  it('returns undefined for absent, mixed, malformed, advanced, or unsafe state', () => {
    const unsafeRows = [
      `<a:tr>${cell()}${cell()}</a:tr>`,
      `<a:tr>${cell(SCHEME_FILL)}${cell()}</a:tr>`,
      `<a:tr>${cell('<a:noFill/>')}${cell(SCHEME_FILL)}</a:tr>`,
      '<a:tr>' +
        cell('<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>') +
        cell('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>') +
      '</a:tr>',
      '<a:tr>' +
        cell('<a:solidFill><a:srgbClr val="D9EAF7"><a:alpha val="75000"/>' +
          '</a:srgbClr></a:solidFill>') +
        cell('<a:solidFill><a:srgbClr val="D9EAF7"><a:alpha val="50000"/>' +
          '</a:srgbClr></a:solidFill>') +
      '</a:tr>',
      '<a:tr>' +
        cell('<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>') +
        cell('<a:solidFill><a:srgbClr val="D9EAF7"><a:alpha val="100000"/>' +
          '</a:srgbClr></a:solidFill>') +
      '</a:tr>',
      `<a:tr>${cell('<a:solidFill/>')}${cell('<a:solidFill/>')}</a:tr>`,
      `<a:tr>${cell('<a:gradFill/>')}${cell('<a:gradFill/>')}</a:tr>`,
      '<a:tr>' + cell('<a:noFill/><a:solidFill><a:srgbClr val="D9EAF7"/>' +
        '</a:solidFill>') + cell('<a:noFill/>') + '</a:tr>',
      '<a:tr><a:tc><a:txBody/><a:tcPr><a:noFill/></a:tcPr>' +
        '<a:tcPr><a:noFill/></a:tcPr></a:tc></a:tr>',
      '<a:tr><a:tc><a:txBody/></a:tc></a:tr>',
      '',
      '<a:tr/>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(readTableFill(fixture.xml, fixture.frame), rows).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('rejects ambiguous table paths and ignores wrong-prefix or descendant lookalikes', () => {
    const table = `<a:tbl><a:tr>${cell(SCHEME_FILL)}</a:tr></a:tbl>`;
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
      expect(readTableFill(fixture.xml, fixture.frame), source).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }

    const lookalikes = parseFrame(
      '<a:tr>' +
        cell(SCHEME_FILL, '<x:solidFill xmlns:x="urn:test"><x:srgbClr val="FF0000"/>' +
          '</x:solidFill>') +
        cell(SCHEME_FILL, '<x:noFill xmlns:x="urn:test"/>') +
      '</a:tr>',
    );
    expect(readTableFill(lookalikes.xml, lookalikes.frame)).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
  });

  it('whole-replaces, clears, repairs, and preserves unrelated XML', () => {
    const target = parseFrame(
      '<a:tr h="30">' +
        cell('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>', '<a:extLst/>',
          ' gridSpan="2"', ' one="1"') +
        cell('<a:solidFill/>', '', ' hMerge="1"', ' two="2"') +
      '</a:tr>' +
      '<a:tr h="40">' +
        cell('<a:gradFill><a:gsLst/></a:gradFill>', '', ' vMerge="1"', ' three="3"') +
        cell('', '<a:extLst/>', '', ' four="4"') +
      '</a:tr>',
    );
    const solid = {
      kind: 'solid' as const,
      color: { kind: 'scheme' as const, value: 'accent1' as const },
      transparency: 25,
    };
    expect(replaceTableFill(target.xml, target.frame, solid, PART_URI)).toBe(true);
    const solidSource = target.xml.serialize();
    expect(solidSource.match(/<a:schemeClr val="accent1">/g)).toHaveLength(4);
    expect(solidSource.match(/<a:alpha val="75000"\/>/g)).toHaveLength(4);
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
      '<a:lnL w="12700">',
      '<a:tblGrid><a:gridCol w="10"/><a:gridCol w="20"/></a:tblGrid>',
      '<p:ext uri="keep"/>',
    ]) expect(solidSource).toContain(retained);
    expect(solidSource.indexOf(SCHEME_FILL)).toBeLessThan(solidSource.indexOf('<a:extLst/>'));

    const uniform = parseSource(solidSource);
    expect(readTableFill(uniform.xml, uniform.frame)).toEqual(solid);
    expect(replaceTableFill(uniform.xml, uniform.frame, solid, PART_URI)).toBe(false);
    expect(uniform.xml.serialize()).toBe(solidSource);

    expect(replaceTableFill(
      uniform.xml,
      uniform.frame,
      { kind: 'none' },
      PART_URI,
    )).toBe(true);
    const none = parseSource(uniform.xml.serialize());
    expect(readTableFill(none.xml, none.frame)).toEqual({ kind: 'none' });
    expect(none.xml.serialize().match(/<a:noFill\/>/g)).toHaveLength(16);

    expect(replaceTableFill(none.xml, none.frame, {
      kind: 'solid',
      color: { kind: 'srgb', value: 'D9EAF7' },
      transparency: 0,
    }, PART_URI)).toBe(true);
    const explicitZero = parseSource(none.xml.serialize());
    expect(readTableFill(explicitZero.xml, explicitZero.frame)).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'D9EAF7' },
      transparency: 0,
    });
    expect(explicitZero.xml.serialize().match(/<a:alpha val="100000"\/>/g)).toHaveLength(4);

    expect(replaceTableFill(explicitZero.xml, explicitZero.frame, undefined, PART_URI)).toBe(true);
    const cleared = parseSource(explicitZero.xml.serialize());
    const clearedSource = cleared.xml.serialize();
    expect(readTableFill(cleared.xml, cleared.frame)).toBeUndefined();
    expect(replaceTableFill(cleared.xml, cleared.frame, undefined, PART_URI)).toBe(false);
    expect(cleared.xml.serialize()).toBe(clearedSource);
  });

  it('rejects unsafe bulk edits while allowing one safe direct choice to normalize', () => {
    const missingTable = parseFrame('');
    expect(() => replaceTableFill(
      missingTable.xml,
      missingTable.frame,
      { kind: 'none' },
      PART_URI,
    )).toThrow(ModelParseError);

    const repeatedFinal = parseFrame(
      '<a:tr>' + cell('<a:noFill/>') +
        cell('<a:noFill/><a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>') +
      '</a:tr>',
    );
    expect(() => replaceTableFill(
      repeatedFinal.xml,
      repeatedFinal.frame,
      { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
      PART_URI,
    )).toThrow('Table cell contains multiple direct fill choices');

    const wrongPrefix = parseFrame(
      '<a:tr>' + cell('', '<x:solidFill xmlns:x="urn:test"/>') + '</a:tr>',
    );
    expect(replaceTableFill(
      wrongPrefix.xml,
      wrongPrefix.frame,
      { kind: 'none' },
      PART_URI,
    )).toBe(true);
    const reopenedWrongPrefix = parseSource(wrongPrefix.xml.serialize());
    expect(readTableFill(reopenedWrongPrefix.xml, reopenedWrongPrefix.frame))
      .toEqual({ kind: 'none' });
    expect(reopenedWrongPrefix.xml.serialize())
      .toContain('<x:solidFill xmlns:x="urn:test"/>');
  });
});
