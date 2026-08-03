import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import { requireEditablePlainTableCellText } from './table-cell-rich-text.internal.js';
import { renderEmptyTableCellFragment } from './table-create.internal.js';
import { replaceTableCellMergeAttributes } from './table-cell-merge.internal.js';
import {
  normalizeTableColumnInsertInput,
  normalizeTableDeleteInput,
  normalizeTableRowInsertInput,
  requireEditableTableStructure,
} from './table-structure-edit.internal.js';

const PART_URI = '/ppt/slides/slide1.xml';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

function parseFrame(source: string): {
  xml: LosslessXmlDocument;
  frame: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(source);
  const frame = xml.roots[0];
  if (!frame) throw new Error('Fixture has no frame');
  return { xml, frame };
}

function cell(attributes = '', text = ''): string {
  return `<a:tc${attributes}><a:txBody><a:bodyPr/><a:lstStyle/>` +
    `<a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody>` +
    '<a:tcPr/></a:tc>';
}

function validFrame(overrides: {
  readonly transform?: string;
  readonly grid?: string;
  readonly rows?: string;
  readonly tableExtra?: string;
  readonly frameExtra?: string;
} = {}): string {
  const transform = overrides.transform ??
    '<p:xfrm keep="X"><a:off x="1" y="2"/><a:ext cx="999" cy="777"/></p:xfrm>';
  const grid = overrides.grid ??
    '<a:tblGrid keep="GRID"><a:gridCol w="100"/><a:gridCol w="200"/>' +
    '<a:gridCol w="300"/><x:gridCol w="999"/></a:tblGrid>';
  const rows = overrides.rows ??
    '<a:tr h="40" keep="ROW-0">' +
      cell(' rowSpan="2" gridSpan="2" keep="A"', 'A') +
      cell(' rowSpan="2" hMerge="1" keep="B"', 'B') +
      cell(' keep="C"', 'C') +
      '<x:tc/>' +
    '</a:tr>' +
    '<a:tr h="60" keep="ROW-1">' +
      cell(' gridSpan="2" vMerge="1" keep="D"', 'D') +
      cell(' vMerge="1" hMerge="1" keep="E"', 'E') +
      cell(' keep="F"', 'F') +
    '</a:tr>' +
    '<x:tr h="999"><x:tc/></x:tr>';
  return `<p:graphicFrame xmlns:p="${PRESENTATION_NAMESPACE}" ` +
    `xmlns:a="${DRAWING_NAMESPACE}" xmlns:x="urn:foreign">` +
    `${transform}${overrides.frameExtra ?? ''}` +
    '<a:graphic><a:graphicData>' +
    `<a:tbl keep="TABLE"><a:tblPr/>${grid}${rows}` +
    `${overrides.tableExtra ?? ''}</a:tbl>` +
    '</a:graphicData></a:graphic></p:graphicFrame>';
}

describe('table structure edit inputs', () => {
  it('normalizes frozen row and column inserts with defaults, scalars, vectors, and rounding', () => {
    expect(normalizeTableRowInsertInput(0, undefined)).toEqual({
      rowIndex: 0,
      count: 1,
    });
    expect(normalizeTableRowInsertInput(3, { count: 2, rowHeights: 0 })).toEqual({
      rowIndex: 3,
      count: 2,
      rowHeights: [0, 0],
    });
    const row = normalizeTableRowInsertInput(
      1,
      Object.assign(Object.create(null), {
        count: 2,
        rowHeights: [10.4, 20.6],
      }),
    );
    expect(row).toEqual({ rowIndex: 1, count: 2, rowHeights: [10, 21] });
    expect(Object.isFrozen(row)).toBe(true);
    expect(Object.isFrozen(row.rowHeights)).toBe(true);

    const column = normalizeTableColumnInsertInput(2, {
      count: 2,
      columnWidths: [100.4, 200.6],
    });
    expect(column).toEqual({
      columnIndex: 2,
      count: 2,
      columnWidths: [100, 201],
    });
    expect(Object.isFrozen(column)).toBe(true);
    expect(Object.isFrozen(column.columnWidths)).toBe(true);
    expect(normalizeTableColumnInsertInput(3, { columnWidths: 50 })).toEqual({
      columnIndex: 3,
      count: 1,
      columnWidths: [50],
    });
  });

  it('normalizes delete defaults and rejects unsafe indexes, counts, options, and dimensions', () => {
    expect(normalizeTableDeleteInput(2, undefined, 'row')).toEqual({
      index: 2,
      count: 1,
    });
    expect(normalizeTableDeleteInput(1, 3, 'column')).toEqual({
      index: 1,
      count: 3,
    });

    const invalidIndexes: readonly unknown[] = [
      -1,
      0.5,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.MAX_SAFE_INTEGER + 1,
      '0',
    ];
    for (const value of invalidIndexes) {
      expect(() => normalizeTableRowInsertInput(value, undefined)).toThrow();
      expect(() => normalizeTableColumnInsertInput(value, undefined)).toThrow();
      expect(() => normalizeTableDeleteInput(value, 1, 'row')).toThrow();
    }

    const invalidCounts: readonly unknown[] = [
      0,
      -1,
      1.5,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.MAX_SAFE_INTEGER + 1,
      '1',
    ];
    for (const count of invalidCounts) {
      expect(() => normalizeTableRowInsertInput(0, { count })).toThrow();
      expect(() => normalizeTableColumnInsertInput(0, { count })).toThrow();
      expect(() => normalizeTableDeleteInput(0, count, 'column')).toThrow();
    }
    expect(() => normalizeTableRowInsertInput(0, {
      count: 1_000_000,
      rowHeights: 1,
    })).toThrow(/physical table cell limit/);
    expect(() => normalizeTableColumnInsertInput(0, {
      count: 1_000_000,
      columnWidths: 1,
    })).toThrow(/physical table cell limit/);

    const inherited = Object.create({ count: 2 }) as Record<string, unknown>;
    const accessor = Object.defineProperty({}, 'count', { get: () => 2 });
    const symbol = Object.assign({}, { [Symbol('count')]: 2 });
    for (const options of [
      null,
      [],
      inherited,
      accessor,
      symbol,
      { unknown: true },
    ]) {
      expect(() => normalizeTableRowInsertInput(0, options)).toThrow();
      expect(() => normalizeTableColumnInsertInput(0, options)).toThrow();
    }

    const sparse = Array(2) as number[];
    sparse[0] = 10;
    for (const rowHeights of [-1, Number.NaN, Number.POSITIVE_INFINITY, sparse, [10]]) {
      expect(() => normalizeTableRowInsertInput(0, {
        count: 2,
        rowHeights,
      })).toThrow();
    }
    for (const columnWidths of [0, -1, 0.4, Number.NaN, sparse, [10]]) {
      expect(() => normalizeTableColumnInsertInput(0, {
        count: 2,
        columnWidths,
      })).toThrow();
    }
  });
});

describe('editable table structure boundary', () => {
  it('reads exact direct structure, sizes, transform, namespaces, and merge state', () => {
    const fixture = parseFrame(validFrame());
    const structure = requireEditableTableStructure(fixture.frame, PART_URI);

    expect(structure.table.localName).toBe('tbl');
    expect(structure.grid.localName).toBe('tblGrid');
    expect(structure.gridColumns.map(({ localName }) => localName)).toEqual([
      'gridCol',
      'gridCol',
      'gridCol',
    ]);
    expect(structure.rows).toHaveLength(2);
    expect(structure.cells.map((row) => row.length)).toEqual([3, 3]);
    expect(structure.columnWidths).toEqual([100, 200, 300]);
    expect(structure.rowHeights).toEqual([40, 60]);
    expect(structure.width).toBe(999);
    expect(structure.height).toBe(777);
    expect(structure.widthAttribute.name).toBe('cx');
    expect(structure.heightAttribute.name).toBe('cy');
    expect(structure.mergeState.regions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
    ]);
    expect(Object.isFrozen(structure)).toBe(true);
    expect(Object.isFrozen(structure.gridColumns)).toBe(true);
    expect(Object.isFrozen(structure.cells)).toBe(true);
    expect(Object.isFrozen(structure.cells[0])).toBe(true);
    expect(Object.isFrozen(structure.columnWidths)).toBe(true);
    expect(fixture.xml.changed).toBe(false);
  });

  it('rejects missing, repeated, jagged, unsafe, and malformed owned structure', () => {
    const invalid = [
      validFrame({ transform: '' }),
      validFrame({ frameExtra: '<p:xfrm><a:ext cx="1" cy="1"/></p:xfrm>' }),
      validFrame({ transform: '<p:xfrm/>' }),
      validFrame({ transform: '<p:xfrm><a:ext cx="1" cy="1"/><a:ext cx="1" cy="1"/></p:xfrm>' }),
      validFrame({ transform: '<p:xfrm><a:ext cx="-1" cy="1"/></p:xfrm>' }),
      validFrame({ transform: '<p:xfrm><a:ext cx="1" cy="1.5"/></p:xfrm>' }),
      validFrame({ grid: '<a:tblGrid/>' }),
      validFrame({ grid: '<a:tblGrid><a:gridCol/></a:tblGrid>' }),
      validFrame({ grid: '<a:tblGrid><a:gridCol w="0"/></a:tblGrid>' }),
      validFrame({ grid: '<a:tblGrid><a:gridCol w="1" w="2"/></a:tblGrid>' }),
      validFrame({
        grid: '<a:tblGrid><a:gridCol w="9007199254740991"/><a:gridCol w="1"/></a:tblGrid>',
        rows: '<a:tr h="1">' + cell() + cell() + '</a:tr>',
      }),
      validFrame({ grid: '<a:tblGrid><a:gridCol w="1"/></a:tblGrid>', rows: '' }),
      validFrame({
        grid: '<a:tblGrid><a:gridCol w="1"/><a:gridCol w="2"/></a:tblGrid>',
        rows: '<a:tr h="1">' + cell() + '</a:tr>',
      }),
      validFrame({
        grid: '<a:tblGrid><a:gridCol w="1"/></a:tblGrid>',
        rows: '<a:tr>' + cell() + '</a:tr>',
      }),
      validFrame({
        grid: '<a:tblGrid><a:gridCol w="1"/></a:tblGrid>',
        rows: '<a:tr h="-1">' + cell() + '</a:tr>',
      }),
      validFrame({
        grid: '<a:tblGrid><a:gridCol w="1"/><a:gridCol w="2"/></a:tblGrid>',
        rows: '<a:tr h="1">' + cell(' gridSpan="2"') + cell() + '</a:tr>',
      }),
      validFrame({ tableExtra: '<a:tblGrid><a:gridCol w="1"/></a:tblGrid>' }),
    ];

    for (const source of invalid) {
      const fixture = parseFrame(source);
      expect(
        () => requireEditableTableStructure(fixture.frame, PART_URI),
        source,
      ).toThrowError(new ModelParseError(
        'Table structure is not safely editable',
        PART_URI,
      ));
      expect(fixture.xml.changed).toBe(false);
    }

    const wrongRoot = parseFrame(
      `<x:graphicFrame xmlns:x="urn:foreign" xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    );
    expect(() => requireEditableTableStructure(wrongRoot.frame, PART_URI))
      .toThrow(/Table structure is not safely editable/);
  });
});

describe('canonical structural table cells', () => {
  it('renders editable empty cells for every merge role', () => {
    const plain = renderEmptyTableCellFragment();
    expect(plain).toMatch(/^<a:tc>/);
    expect(plain.match(/<a:txBody>/g)).toHaveLength(1);
    expect(plain.match(/<a:bodyPr\/>/g)).toHaveLength(1);
    expect(plain.match(/<a:lstStyle\/>/g)).toHaveLength(1);
    expect(plain.match(/<a:p>/g)).toHaveLength(1);
    expect(plain.match(/<a:r>/g)).toHaveLength(1);
    expect(plain.match(/<a:t\b[^>]*><\/a:t>/g)).toHaveLength(1);
    expect(plain.match(/<a:tcPr\b/g)).toHaveLength(1);
    expect(plain).not.toMatch(/hlink|rowSpan|gridSpan|vMerge|hMerge/);
    const editable = parseFrame(
      plain.replace('<a:tc', `<a:tc xmlns:a="${DRAWING_NAMESPACE}"`),
    );
    const text = requireEditablePlainTableCellText(
      editable.xml,
      editable.frame,
      PART_URI,
    );
    editable.xml.replaceText(text, 'Filled');
    expect(editable.xml.serialize()).toContain('>Filled</a:t>');

    expect(renderEmptyTableCellFragment({ rowSpan: 2, gridSpan: 3 }))
      .toMatch(/^<a:tc rowSpan="2" gridSpan="3"><a:txBody>/);
    expect(renderEmptyTableCellFragment({ rowSpan: 2, horizontal: true }))
      .toMatch(/^<a:tc rowSpan="2" hMerge="1"><a:txBody>/);
    expect(renderEmptyTableCellFragment({ gridSpan: 3, vertical: true }))
      .toMatch(/^<a:tc gridSpan="3" vMerge="1"><a:txBody>/);
    expect(renderEmptyTableCellFragment({ vertical: true, horizontal: true }))
      .toMatch(/^<a:tc vMerge="1" hMerge="1"><a:txBody>/);
  });

  it('preserves semantic no-ops and rewrites only unqualified merge attributes', () => {
    const equivalent = parseFrame(
      `<a:tc xmlns:a="${DRAWING_NAMESPACE}" xmlns:x="urn:foreign" ` +
      `rowSpan='02' gridSpan="03" vMerge="true" hMerge='1' ` +
      `x:rowSpan="9" keep='YES'> \n<a:txBody keep="BODY"/><a:tcPr/> </a:tc>`,
    );
    expect(replaceTableCellMergeAttributes(equivalent.xml, equivalent.frame, {
      rowSpan: 2,
      gridSpan: 3,
      vertical: true,
      horizontal: true,
    })).toBe(false);
    expect(equivalent.xml.changed).toBe(false);

    expect(replaceTableCellMergeAttributes(equivalent.xml, equivalent.frame, {
      gridSpan: 4,
      horizontal: true,
    })).toBe(true);
    const updated = equivalent.xml.serialize();
    expect(updated).toContain(`x:rowSpan="9" keep='YES' gridSpan="4" hMerge="1"`);
    expect(updated).not.toMatch(/\srowSpan=['"]02|\svMerge=/);
    expect(updated).toContain(' \n<a:txBody keep="BODY"/><a:tcPr/> ');
  });
});
