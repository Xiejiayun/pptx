import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTableMargins,
  replaceTableMargins,
} from './table-cell-margins.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';

function cell(attributes = '', extraProperties = '', cellAttributes = ''): string {
  return `<a:tc${cellAttributes}><a:txBody><a:bodyPr lIns="11"/>` +
    '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"/><a:t>Keep</a:t></a:r></a:p>' +
    `</a:txBody><a:tcPr anchor="ctr" vert="vert270"${attributes}>` +
    `<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>${extraProperties}` +
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

describe('table-level margins', () => {
  it('reads complete and partial uniform direct physical-cell vectors', () => {
    const complete = parseFrame(
      '<a:tr h="30">' +
        cell(' marL="50800" marR="25400" marT="12700" marB="38100"', '', ' gridSpan="2"') +
        cell(' marL="050800" marR="0025400" marT="0012700" marB="038100"', '', ' hMerge="1"') +
      '</a:tr>' +
      '<a:tr h="40">' +
        cell(' marL="50800" marR="25400" marT="12700" marB="38100"', '', ' vMerge="1"') +
        cell(' marL="50800" marR="25400" marT="12700" marB="38100"') +
      '</a:tr>',
    );
    expect(readTableMargins(complete.xml, complete.frame)).toEqual({
      top: 1,
      right: 2,
      bottom: 3,
      left: 4,
    });
    expect(complete.xml.changed).toBe(false);

    const partial = parseFrame(
      '<a:tr>' +
        cell(' marL="-25400" marT="0"') +
        cell(' marL="-025400" marT="00"') +
      '</a:tr>',
    );
    expect(readTableMargins(partial.xml, partial.frame)).toEqual({ top: 0, left: -2 });
    expect(partial.xml.changed).toBe(false);
  });

  it('returns undefined for absent, mixed, malformed, repeated, or incomplete state', () => {
    const unsafeRows = [
      `<a:tr>${cell()}${cell()}</a:tr>`,
      `<a:tr>${cell(' marT="12700"')}${cell()}</a:tr>`,
      `<a:tr>${cell(' marT="12700"')}${cell(' marL="12700"')}</a:tr>`,
      `<a:tr>${cell(' marT="12700"')}${cell(' marT="25400"')}</a:tr>`,
      `<a:tr>${cell(' marT="1.0"')}${cell(' marT="1.0"')}</a:tr>`,
      `<a:tr>${cell(' marT="2147483648"')}${cell(' marT="2147483648"')}</a:tr>`,
      `<a:tr>${cell(' marT="12700" marT="12700"')}${cell(' marT="12700"')}</a:tr>`,
      '<a:tr><a:tc><a:txBody/><a:tcPr marT="12700"/><a:tcPr marT="12700"/></a:tc></a:tr>',
      '<a:tr><a:tc><a:txBody/></a:tc></a:tr>',
      '',
      '<a:tr/>',
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(readTableMargins(fixture.xml, fixture.frame), rows).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }
  });

  it('rejects ambiguous direct table paths and ignores qualified or descendant lookalikes', () => {
    const table = `<a:tbl><a:tr>${cell(' marT="12700"')}</a:tr></a:tbl>`;
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
      expect(readTableMargins(fixture.xml, fixture.frame), source).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }

    const lookalikes = parseFrame(
      '<a:tr>' +
        cell(' xmlns:x="urn:test" x:marT="999" marT="12700"', '<x:keep marT="999"/>') +
        cell(' xmlns:x="urn:test" x:marT="888" marT="12700"', '<x:keep marT="888"/>') +
      '</a:tr>',
    );
    expect(readTableMargins(lookalikes.xml, lookalikes.frame)).toEqual({ top: 1 });
  });

  it('whole-replaces, clears, repairs, and preserves unrelated XML across all cells', () => {
    const target = parseFrame(
      '<a:tr h="30">' +
        cell(' marL="12700" marR="25400" marT="38100" marB="50800" keep="one"', '<a:extLst/>', ' gridSpan="2"') +
        cell(' marL="bad" marR="25400" marT="38100" marB="50800" keep="two"', '', ' hMerge="1"') +
      '</a:tr>' +
      '<a:tr h="40">' +
        cell(' marL="12700" keep="three"', '', ' vMerge="1"') +
        cell(' keep="four"') +
      '</a:tr>',
    );
    expect(replaceTableMargins(target.xml, target.frame, {
      top: 6,
      right: 6,
      bottom: 6,
      left: 6,
    }, PART_URI)).toBe(true);
    const uniformSource = target.xml.serialize();
    expect(uniformSource.match(/marL="76200"/g)).toHaveLength(4);
    expect(uniformSource.match(/marR="76200"/g)).toHaveLength(4);
    expect(uniformSource.match(/marT="76200"/g)).toHaveLength(4);
    expect(uniformSource.match(/marB="76200"/g)).toHaveLength(4);
    for (const retained of [
      'gridSpan="2"',
      'hMerge="1"',
      'vMerge="1"',
      'anchor="ctr"',
      'vert="vert270"',
      'keep="one"',
      'keep="two"',
      'keep="three"',
      'keep="four"',
      '<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>',
      '<a:bodyPr lIns="11"/>',
      '<a:pPr algn="ctr"/>',
      '<a:rPr lang="en-US"/>',
      '<a:tblGrid><a:gridCol w="10"/><a:gridCol w="20"/></a:tblGrid>',
      '<p:ext uri="keep"/>',
    ]) expect(uniformSource).toContain(retained);

    const uniform = parseSource(uniformSource);
    expect(readTableMargins(uniform.xml, uniform.frame)).toEqual({
      top: 6,
      right: 6,
      bottom: 6,
      left: 6,
    });
    expect(replaceTableMargins(uniform.xml, uniform.frame, {
      top: 6,
      right: 6,
      bottom: 6,
      left: 6,
    }, PART_URI)).toBe(false);
    expect(uniform.xml.serialize()).toBe(uniformSource);

    expect(replaceTableMargins(
      uniform.xml,
      uniform.frame,
      { top: 2, left: 4 },
      PART_URI,
    )).toBe(true);
    const partialSource = uniform.xml.serialize();
    expect(partialSource.match(/marT="25400"/g)).toHaveLength(4);
    expect(partialSource.match(/marL="50800"/g)).toHaveLength(4);
    expect(partialSource).not.toMatch(/\smarR=/);
    expect(partialSource).not.toMatch(/\smarB=/);

    const partial = parseSource(partialSource);
    expect(readTableMargins(partial.xml, partial.frame)).toEqual({ top: 2, left: 4 });
    expect(replaceTableMargins(partial.xml, partial.frame, undefined, PART_URI)).toBe(true);
    const clearedSource = partial.xml.serialize();
    expect(clearedSource).not.toMatch(/\smar[LTRB]=/);
    const cleared = parseSource(clearedSource);
    expect(readTableMargins(cleared.xml, cleared.frame)).toBeUndefined();
    expect(replaceTableMargins(cleared.xml, cleared.frame, undefined, PART_URI)).toBe(false);
    expect(cleared.xml.serialize()).toBe(clearedSource);
  });

  it('rejects unsafe edits, including repeated owned state on a late cell', () => {
    const unsafeRows = [
      '',
      '<a:tr/>',
      '<a:tr><a:tc><a:txBody/></a:tc></a:tr>',
      '<a:tr><a:tc><a:txBody/><a:tcPr/><a:tcPr/></a:tc></a:tr>',
      `<a:tr>${cell(' marT="12700"')}${cell(' marT="12700" marT="25400"')}</a:tr>`,
    ];
    for (const rows of unsafeRows) {
      const fixture = parseFrame(rows);
      expect(() => replaceTableMargins(
        fixture.xml,
        fixture.frame,
        { top: 2 },
        PART_URI,
      ), rows).toThrow(ModelParseError);
    }
  });
});
