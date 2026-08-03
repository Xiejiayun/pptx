import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  clearTableMergeRegionAt,
  normalizeTableMergeRegionInput,
  readTableMergeState,
  replaceTableMergeRegion,
} from './table-cell-merge.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';

function parseFrame(
  rows: string,
  columnCount = 4,
): { xml: LosslessXmlDocument; frame: XmlElement } {
  return parseSource(
    '<p:graphicFrame xmlns:p="p" xmlns:a="a" xmlns:x="urn:test">' +
      '<a:graphic><a:graphicData><a:tbl>' +
      '<a:tblGrid>' +
      Array.from({ length: columnCount }, (_, index) =>
        `<a:gridCol w="${index + 1}"/>`).join('') +
      '</a:tblGrid>' +
      rows +
      '</a:tbl></a:graphicData></a:graphic>' +
      '</p:graphicFrame>',
  );
}

function parseSource(
  source: string,
): { xml: LosslessXmlDocument; frame: XmlElement } {
  const xml = LosslessXmlDocument.parse(source);
  const frame = xml.roots[0];
  if (!frame) throw new Error('Fixture has no root frame');
  return { xml, frame };
}

function cell(attributes = '', body = '<a:tcPr/>'): string {
  return `<a:tc${attributes}>${body}</a:tc>`;
}

function row(cells: string): string {
  return `<a:tr h="1">${cells}</a:tr>`;
}

function plainRows(rowCount: number, columnCount: number): string {
  return Array.from({ length: rowCount }, (_, rowIndex) => row(
    Array.from({ length: columnCount }, (_, columnIndex) => cell(
      ` keep='${rowIndex},${columnIndex}'`,
      '<a:txBody><a:p><a:r><a:t>' +
        `${rowIndex},${columnIndex}` +
        '</a:t></a:r></a:p></a:txBody>' +
        `<a:tcPr keep="STYLE-${rowIndex},${columnIndex}"/>`,
    )).join(''),
  )).join('');
}

describe('table cell merge topology', () => {
  it('reads horizontal, vertical, rectangular, offset, and lexical states into frozen snapshots', () => {
    const fixture = parseFrame(
      row(
        cell(' gridSpan="02" keep="H"') +
        cell(' hMerge="true"') +
        cell(' x:gridSpan="9" x:hMerge="1" keep="IMPOSTOR"') +
        cell(' rowSpan="02" keep="V"'),
      ) +
      row(
        cell(' keep="PLAIN"') +
        cell(' rowSpan="2" gridSpan="2" keep="RECT"') +
        cell(' rowSpan="02" hMerge="1"') +
        cell(' vMerge="true"'),
      ) +
      row(
        cell(' keep="PLAIN-2"') +
        cell(' gridSpan="02" vMerge="1"') +
        cell(' vMerge="true" hMerge="1"') +
        cell(' hMerge="0" vMerge="false" keep="FALSE"'),
      ),
    );

    const state = readTableMergeState(fixture.frame);
    expect(state?.regions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 1, colspan: 2 },
      { rowIndex: 0, columnIndex: 3, rowspan: 2, colspan: 1 },
      { rowIndex: 1, columnIndex: 1, rowspan: 2, colspan: 2 },
    ]);
    expect(state?.cells[0]?.[0]).toEqual({
      rowIndex: 0,
      columnIndex: 0,
      rowspan: 1,
      colspan: 2,
      isAnchor: true,
    });
    expect(state?.cells[2]?.[2]).toEqual({
      rowIndex: 1,
      columnIndex: 1,
      rowspan: 2,
      colspan: 2,
      isAnchor: false,
    });
    expect(state?.cells[0]?.[2]).toBeUndefined();
    expect(state?.cells[2]?.[3]).toBeUndefined();
    expect(Object.isFrozen(state)).toBe(true);
    expect(Object.isFrozen(state?.regions)).toBe(true);
    expect(Object.isFrozen(state?.regions[0])).toBe(true);
    expect(Object.isFrozen(state?.cells)).toBe(true);
    expect(Object.isFrozen(state?.cells[2])).toBe(true);
    expect(Object.isFrozen(state?.cells[2]?.[2])).toBe(true);
    expect(fixture.xml.changed).toBe(false);
  });

  it('distinguishes recognized unmerged tables from malformed or ambiguous topology', () => {
    const unmerged = parseFrame(plainRows(2, 4));
    expect(readTableMergeState(unmerged.frame)).toEqual({
      regions: [],
      cells: [
        [undefined, undefined, undefined, undefined],
        [undefined, undefined, undefined, undefined],
      ],
    });

    const invalidRows = [
      '',
      row(cell() + cell() + cell()),
      row(cell() + cell() + cell() + cell()) + row(cell() + cell() + cell()),
      row(cell(' gridSpan="0"') + cell() + cell() + cell()),
      row(cell(' gridSpan="-1"') + cell() + cell() + cell()),
      row(cell(' rowSpan="1.5"') + cell() + cell() + cell()),
      row(cell(' hMerge="yes"') + cell() + cell() + cell()),
      row(cell(' gridSpan="2" gridSpan="2"') + cell(' hMerge="1"') + cell() + cell()),
      row(cell(' gridSpan="2"') + cell() + cell() + cell()),
      row(cell(' hMerge="1"') + cell() + cell() + cell()),
      row(cell(' rowSpan="2"') + cell() + cell() + cell()),
      row(cell(' rowSpan="2"') + cell() + cell() + cell()) +
        row(cell() + cell() + cell() + cell()),
      row(cell(' gridSpan="3"') + cell(' hMerge="1"') + cell(' hMerge="1"') + cell(' rowSpan="2"')) +
        row(cell() + cell(' gridSpan="3" vMerge="1"') + cell(' vMerge="1" hMerge="1"') + cell(' vMerge="1"')),
    ];
    for (const rows of invalidRows) {
      const fixture = parseFrame(rows);
      expect(readTableMergeState(fixture.frame), rows).toBeUndefined();
      expect(fixture.xml.changed).toBe(false);
    }

    const invalidSources = [
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"/>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic/><a:graphic/></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData/><a:graphicData/></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData><a:tbl/><a:tbl/></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData><a:tbl><a:tblGrid/><a:tr h="1"><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
      '<p:graphicFrame xmlns:p="p" xmlns:a="a"><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol/><a:gridCol/></a:tblGrid><a:tblGrid><a:gridCol/></a:tblGrid><a:tr h="1"><a:tc/><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>',
    ];
    for (const source of invalidSources) {
      const fixture = parseSource(source);
      expect(readTableMergeState(fixture.frame), source).toBeUndefined();
    }
  });

  it('normalizes strict zero-based physical regions', () => {
    expect(normalizeTableMergeRegionInput(1, 2, 3, 4)).toEqual({
      rowIndex: 1,
      columnIndex: 2,
      rowspan: 3,
      colspan: 4,
    });
    const invalid = [
      [-1, 0, 1, 2],
      [0, -1, 1, 2],
      [0.5, 0, 1, 2],
      [0, 0, 0, 2],
      [0, 0, 1, 0],
      [0, 0, 1.5, 2],
      [0, 0, 1, 2.5],
      [0, 0, 1, 1],
      [0, 0, Number.MAX_SAFE_INTEGER + 1, 2],
      ['0', 0, 1, 2],
      [0, 0, Number.NaN, 2],
      [0, 0, Number.POSITIVE_INFINITY, 2],
    ] as const;
    for (const values of invalid) {
      expect(() => normalizeTableMergeRegionInput(
        values[0],
        values[1],
        values[2],
        values[3],
      ), String(values)).toThrow();
    }
  });

  it('adds canonical non-overlapping merges and preserves all non-merge cell state', () => {
    const fixture = parseFrame(plainRows(3, 4));
    expect(replaceTableMergeRegion(
      fixture.xml,
      fixture.frame,
      normalizeTableMergeRegionInput(0, 0, 2, 2),
      PART_URI,
    )).toBe(true);
    const merged = fixture.xml.serialize();
    expect(merged).toContain("<a:tc keep='0,0' rowSpan=\"2\" gridSpan=\"2\">");
    expect(merged).toContain("<a:tc keep='0,1' rowSpan=\"2\" hMerge=\"1\">");
    expect(merged).toContain("<a:tc keep='1,0' gridSpan=\"2\" vMerge=\"1\">");
    expect(merged).toContain("<a:tc keep='1,1' vMerge=\"1\" hMerge=\"1\">");
    expect(merged).toContain('<a:t>1,1</a:t>');
    expect(merged).toContain('<a:tcPr keep="STYLE-1,1"/>');
    expect(merged).toContain("<a:tc keep='2,3'>");

    const reparsed = parseSource(merged);
    expect(replaceTableMergeRegion(
      reparsed.xml,
      reparsed.frame,
      normalizeTableMergeRegionInput(0, 0, 2, 2),
      PART_URI,
    )).toBe(false);
    expect(reparsed.xml.changed).toBe(false);

    expect(replaceTableMergeRegion(
      reparsed.xml,
      reparsed.frame,
      normalizeTableMergeRegionInput(2, 2, 1, 2),
      PART_URI,
    )).toBe(true);
    const withSecond = parseSource(reparsed.xml.serialize());
    expect(readTableMergeState(withSecond.frame)?.regions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
      { rowIndex: 2, columnIndex: 2, rowspan: 1, colspan: 2 },
    ]);
  });

  it('rejects overlap, bounds, and unsupported topology before mutation', () => {
    const merged = parseFrame(
      row(cell(' rowSpan="2" gridSpan="2"') + cell(' rowSpan="2" hMerge="1"') + cell() + cell()) +
      row(cell(' gridSpan="2" vMerge="1"') + cell(' vMerge="1" hMerge="1"') + cell() + cell()) +
      row(cell() + cell() + cell() + cell()),
    );
    for (const region of [
      normalizeTableMergeRegionInput(0, 1, 1, 2),
      normalizeTableMergeRegionInput(1, 0, 2, 1),
      normalizeTableMergeRegionInput(0, 0, 3, 2),
      normalizeTableMergeRegionInput(2, 3, 1, 2),
      normalizeTableMergeRegionInput(2, 0, 2, 1),
    ]) {
      expect(() => replaceTableMergeRegion(
        merged.xml,
        merged.frame,
        region,
        PART_URI,
      )).toThrow();
      expect(merged.xml.changed).toBe(false);
    }

    const malformed = parseFrame(
      row(cell(' hMerge="1"') + cell() + cell() + cell()),
    );
    expect(() => replaceTableMergeRegion(
      malformed.xml,
      malformed.frame,
      normalizeTableMergeRegionInput(0, 0, 1, 2),
      PART_URI,
    )).toThrow(ModelParseError);
    expect(() => clearTableMergeRegionAt(
      malformed.xml,
      malformed.frame,
      0,
      0,
      PART_URI,
    )).toThrow(`Table merge state is not safely editable: ${PART_URI}`);
    expect(malformed.xml.changed).toBe(false);
  });

  it('unmerges from any physical member without changing hidden cell state', () => {
    const fixture = parseFrame(
      row(
        cell(' rowSpan="02" gridSpan="02" keep="ANCHOR"', '<a:txBody><a:p><a:r><a:t>Anchor</a:t></a:r></a:p></a:txBody><a:tcPr keep="A"/>') +
        cell(' rowSpan="02" hMerge="true" keep="TOP"', '<a:txBody><a:p><a:r><a:t>Hidden top</a:t></a:r></a:p></a:txBody><a:tcPr keep="B"/>') +
        cell(' keep="02"') +
        cell(' keep="03"'),
      ) +
      row(
        cell(' gridSpan="02" vMerge="true" keep="LEFT"', '<a:txBody><a:p><a:r><a:t>Hidden left</a:t></a:r></a:p></a:txBody><a:tcPr keep="C"/>') +
        cell(' vMerge="1" hMerge="1" keep="INNER"', '<a:txBody><a:p><a:r><a:t>Hidden inner</a:t></a:r></a:p></a:txBody><a:tcPr keep="D"/>') +
        cell(' keep="12"') +
        cell(' keep="13"'),
      ),
    );
    expect(clearTableMergeRegionAt(
      fixture.xml,
      fixture.frame,
      1,
      1,
      PART_URI,
    )).toBe(true);
    const unmerged = fixture.xml.serialize();
    for (const token of ['gridSpan=', 'rowSpan=', 'hMerge=', 'vMerge=']) {
      expect(unmerged).not.toContain(token);
    }
    for (const retained of [
      'keep="ANCHOR"',
      'keep="TOP"',
      'keep="LEFT"',
      'keep="INNER"',
      '<a:t>Hidden top</a:t>',
      '<a:t>Hidden left</a:t>',
      '<a:t>Hidden inner</a:t>',
      '<a:tcPr keep="D"/>',
    ]) expect(unmerged).toContain(retained);

    const reparsed = parseSource(unmerged);
    expect(clearTableMergeRegionAt(
      reparsed.xml,
      reparsed.frame,
      1,
      1,
      PART_URI,
    )).toBe(false);
    expect(reparsed.xml.changed).toBe(false);
    expect(() => clearTableMergeRegionAt(
      reparsed.xml,
      reparsed.frame,
      2,
      0,
      PART_URI,
    )).toThrow(RangeError);
    expect(reparsed.xml.changed).toBe(false);
  });
});
