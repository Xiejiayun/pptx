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
  deleteTableRows,
  insertTableRows,
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
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

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

function plainRows(heights: readonly number[], columnCount: number): string {
  return heights.map((height, rowIndex) =>
    `<a:tr h="${height}" keep="ROW-${rowIndex}">` +
    Array.from({ length: columnCount }, (_, columnIndex) =>
      cell(` keep="CELL-${rowIndex}-${columnIndex}"`, `${rowIndex},${columnIndex}`))
      .join('') +
    '</a:tr>').join('');
}

function plainFrame(
  heights: readonly number[],
  columnCount = 3,
  tableExtra = '',
): string {
  return validFrame({
    transform: '<p:xfrm><a:ext cx="600" cy="777"/></p:xfrm>',
    grid: '<a:tblGrid>' + Array.from(
      { length: columnCount },
      (_, index) => `<a:gridCol w="${(index + 1) * 100}"/>`,
    ).join('') + '</a:tblGrid>',
    rows: plainRows(heights, columnCount),
    tableExtra,
  });
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

describe('lossless table row insertion', () => {
  it('inserts prepend, middle, append, scalar/vector/default heights and updates cy safely', () => {
    const middle = parseFrame(plainFrame([40, 60], 2));
    const before = requireEditableTableStructure(middle.frame, PART_URI);
    const firstRow = middle.xml.original(before.rows[0]!);
    const secondRow = middle.xml.original(before.rows[1]!);
    insertTableRows(
      middle.xml,
      middle.frame,
      normalizeTableRowInsertInput(1, { count: 2, rowHeights: [50, 0] }),
      PART_URI,
    );
    const middleSource = middle.xml.serialize();
    expect(middleSource).toContain(firstRow);
    expect(middleSource).toContain(secondRow);
    const middleReopened = parseFrame(middleSource);
    const middleState = requireEditableTableStructure(middleReopened.frame, PART_URI);
    expect(middleState.rowHeights).toEqual([40, 50, 0, 60]);
    expect(middleState.cells.map((row) => row.length)).toEqual([2, 2, 2, 2]);
    expect(middleState.height).toBe(777);
    for (const rowIndex of [1, 2]) {
      for (const insertedCell of middleState.cells[rowIndex]!) {
        expect(requireEditablePlainTableCellText(
          middleReopened.xml,
          insertedCell,
          PART_URI,
        )).toBeDefined();
      }
    }

    const append = parseFrame(plainFrame([40, 60], 2));
    insertTableRows(
      append.xml,
      append.frame,
      normalizeTableRowInsertInput(2, undefined),
      PART_URI,
    );
    const appendState = requireEditableTableStructure(
      parseFrame(append.xml.serialize()).frame,
      PART_URI,
    );
    expect(appendState.rowHeights).toEqual([40, 60, 60]);
    expect(appendState.height).toBe(160);

    const copiedMiddle = parseFrame(plainFrame([40, 60], 2));
    insertTableRows(
      copiedMiddle.xml,
      copiedMiddle.frame,
      normalizeTableRowInsertInput(1, undefined),
      PART_URI,
    );
    expect(requireEditableTableStructure(
      parseFrame(copiedMiddle.xml.serialize()).frame,
      PART_URI,
    ).rowHeights).toEqual([40, 60, 60]);

    const prepend = parseFrame(plainFrame([40, 60], 2));
    insertTableRows(
      prepend.xml,
      prepend.frame,
      normalizeTableRowInsertInput(0, { count: 2, rowHeights: 25 }),
      PART_URI,
    );
    const prependState = requireEditableTableStructure(
      parseFrame(prepend.xml.serialize()).frame,
      PART_URI,
    );
    expect(prependState.rowHeights).toEqual([25, 25, 40, 60]);
    expect(prependState.height).toBe(150);

    const alternateSource = plainFrame([10], 1)
      .replaceAll('a:', 'd:')
      .replace('xmlns:a=', 'xmlns:d=');
    const alternate = parseFrame(alternateSource);
    insertTableRows(
      alternate.xml,
      alternate.frame,
      normalizeTableRowInsertInput(1, { rowHeights: 10 }),
      PART_URI,
    );
    const alternateUpdated = alternate.xml.serialize();
    expect(alternateUpdated).toContain(
      `<a:tr xmlns:a="${DRAWING_NAMESPACE}" h="10">`,
    );
    expect(requireEditableTableStructure(
      parseFrame(alternateUpdated).frame,
      PART_URI,
    ).rowHeights).toEqual([10, 10]);
  });

  it('expands only strict-inside merge regions and renders editable continuations', () => {
    const inside = parseFrame(validFrame());
    insertTableRows(
      inside.xml,
      inside.frame,
      normalizeTableRowInsertInput(1, { count: 2, rowHeights: 50 }),
      PART_URI,
    );
    const reopened = parseFrame(inside.xml.serialize());
    const state = requireEditableTableStructure(reopened.frame, PART_URI);
    expect(state.rowHeights).toEqual([40, 50, 50, 60]);
    expect(state.mergeState.regions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 4, colspan: 2 },
    ]);
    expect(reopened.xml.original(state.cells[0]![0]!)).toContain('rowSpan="4"');
    expect(reopened.xml.original(state.cells[1]![0]!))
      .toMatch(/^<a:tc gridSpan="2" vMerge="1"><a:txBody>/);
    expect(reopened.xml.original(state.cells[1]![1]!))
      .toMatch(/^<a:tc vMerge="1" hMerge="1"><a:txBody>/);
    const insertedText = requireEditablePlainTableCellText(
      reopened.xml,
      state.cells[1]![1]!,
      PART_URI,
    );
    reopened.xml.replaceText(insertedText, 'Hidden but editable');
    expect(reopened.xml.serialize()).toContain('>Hidden but editable</a:t>');

    const before = parseFrame(validFrame());
    const originalBefore = requireEditableTableStructure(before.frame, PART_URI);
    const originalAnchor = before.xml.original(originalBefore.cells[0]![0]!);
    insertTableRows(
      before.xml,
      before.frame,
      normalizeTableRowInsertInput(0, { rowHeights: 5 }),
      PART_URI,
    );
    const beforeSource = before.xml.serialize();
    expect(beforeSource).toContain(originalAnchor);
    expect(requireEditableTableStructure(parseFrame(beforeSource).frame, PART_URI)
      .mergeState.regions).toEqual([
        { rowIndex: 1, columnIndex: 0, rowspan: 2, colspan: 2 },
      ]);

    const after = parseFrame(validFrame());
    insertTableRows(
      after.xml,
      after.frame,
      normalizeTableRowInsertInput(2, { rowHeights: 5 }),
      PART_URI,
    );
    expect(requireEditableTableStructure(parseFrame(after.xml.serialize()).frame, PART_URI)
      .mergeState.regions).toEqual([
        { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
      ]);

    const multiple = parseFrame(validFrame({
      transform: '<p:xfrm><a:ext cx="300" cy="30"/></p:xfrm>',
      grid: '<a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/>' +
        '<a:gridCol w="100"/></a:tblGrid>',
      rows: '<a:tr h="10">' + cell(' rowSpan="3"', 'L') + cell('', 'M0') +
        cell(' rowSpan="3"', 'R') + '</a:tr>' +
        '<a:tr h="10">' + cell(' vMerge="1"', 'L1') + cell('', 'M1') +
        cell(' vMerge="1"', 'R1') + '</a:tr>' +
        '<a:tr h="10">' + cell(' vMerge="1"', 'L2') + cell('', 'M2') +
        cell(' vMerge="1"', 'R2') + '</a:tr>',
    }));
    insertTableRows(
      multiple.xml,
      multiple.frame,
      normalizeTableRowInsertInput(1, { rowHeights: 10 }),
      PART_URI,
    );
    expect(requireEditableTableStructure(parseFrame(multiple.xml.serialize()).frame, PART_URI)
      .mergeState.regions).toEqual([
        { rowIndex: 0, columnIndex: 0, rowspan: 4, colspan: 1 },
        { rowIndex: 0, columnIndex: 2, rowspan: 4, colspan: 1 },
      ]);
  });

  it('rejects invalid ranges and physical-cell overflow before scheduling XML changes', () => {
    const fixture = parseFrame(plainFrame([40, 60], 3));
    for (const input of [
      normalizeTableRowInsertInput(3, undefined),
      Object.freeze({ rowIndex: 0, count: 333_333 }),
    ]) {
      expect(() => insertTableRows(
        fixture.xml,
        fixture.frame,
        input,
        PART_URI,
      )).toThrow();
      expect(fixture.xml.changed).toBe(false);
    }

    const overflow = parseFrame(validFrame({
      transform: '<p:xfrm><a:ext cx="1" cy="1"/></p:xfrm>',
      grid: '<a:tblGrid><a:gridCol w="1"/></a:tblGrid>',
      rows: '<a:tr h="9007199254740991">' + cell() +
        '</a:tr><a:tr h="1">' + cell() + '</a:tr>',
    }));
    expect(() => insertTableRows(
      overflow.xml,
      overflow.frame,
      normalizeTableRowInsertInput(2, undefined),
      PART_URI,
    )).toThrow(/sum must fit a safe integer/);
    expect(overflow.xml.changed).toBe(false);
  });
});

describe('lossless table row deletion', () => {
  it('deletes exact row ranges, preserves interstitial/survivor bytes, updates cy, and collects IDs', () => {
    const linkedRows = [0, 1, 2, 3].map((rowIndex) => {
      const relationship = rowIndex === 0
        ? '<a:rPr><a:hlinkClick r:id="rIdShared"/></a:rPr>'
        : rowIndex === 1
          ? '<a:rPr><a:hlinkClick r:id="rIdUnique"/></a:rPr>'
          : rowIndex === 2
            ? '<a:rPr><a:hlinkClick r:id="rIdShared" x:id="rIdForeign"/></a:rPr>'
            : '';
      return `<a:tr h="${(rowIndex + 1) * 10}" keep="ROW-${rowIndex}">` +
        `<a:tc keep="CELL-${rowIndex}"><a:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p><a:r>${relationship}<a:t>${rowIndex}</a:t></a:r></a:p>` +
        '</a:txBody><a:tcPr/></a:tc></a:tr>';
    }).join('<x:between keep="YES"/>');
    const source = validFrame({
      transform: '<p:xfrm><a:ext cx="100" cy="100"/></p:xfrm>',
      grid: '<a:tblGrid><a:gridCol w="100"/></a:tblGrid>',
      rows: linkedRows,
    }).replace(
      'xmlns:x="urn:foreign"',
      `xmlns:x="urn:foreign" xmlns:r="${RELATIONSHIP_NAMESPACE}"`,
    );
    const fixture = parseFrame(source);
    const original = requireEditableTableStructure(fixture.frame, PART_URI);
    const first = fixture.xml.original(original.rows[0]!);
    const last = fixture.xml.original(original.rows[3]!);
    const removed = deleteTableRows(
      fixture.xml,
      fixture.frame,
      normalizeTableDeleteInput(1, 2, 'row'),
      PART_URI,
    );
    expect([...removed].sort()).toEqual(['rIdShared', 'rIdUnique']);
    const updated = fixture.xml.serialize();
    expect(updated).toContain(first);
    expect(updated).toContain(last);
    expect(updated.match(/<x:between keep="YES"\/>/g)).toHaveLength(3);
    expect(updated).not.toContain('rIdUnique');
    expect(updated).toContain('rIdShared');
    expect(updated).not.toContain('rIdForeign');
    const state = requireEditableTableStructure(parseFrame(updated).frame, PART_URI);
    expect(state.rowHeights).toEqual([10, 40]);
    expect(state.height).toBe(50);

    const automatic = parseFrame(plainFrame([10, 0, 20], 1));
    deleteTableRows(
      automatic.xml,
      automatic.frame,
      normalizeTableDeleteInput(0, 1, 'row'),
      PART_URI,
    );
    const automaticState = requireEditableTableStructure(
      parseFrame(automatic.xml.serialize()).frame,
      PART_URI,
    );
    expect(automaticState.rowHeights).toEqual([0, 20]);
    expect(automaticState.height).toBe(777);
  });

  it('shrinks merges, promotes hidden anchors, degrades to one dimension, and dissolves 1 x 1', () => {
    const promoted = parseFrame(validFrame());
    deleteTableRows(
      promoted.xml,
      promoted.frame,
      normalizeTableDeleteInput(0, 1, 'row'),
      PART_URI,
    );
    const promotedReopened = parseFrame(promoted.xml.serialize());
    const promotedState = requireEditableTableStructure(promotedReopened.frame, PART_URI);
    expect(promotedState.mergeState.regions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 1, colspan: 2 },
    ]);
    expect(promotedReopened.xml.text(promotedState.cells[0]![0]!)).toBe('D');
    expect(promotedReopened.xml.original(promotedState.cells[0]![0]!))
      .toContain('gridSpan="2"');
    expect(promotedReopened.xml.original(promotedState.cells[0]![0]!))
      .not.toMatch(/rowSpan|vMerge/);

    const vertical = parseFrame(validFrame({
      transform: '<p:xfrm><a:ext cx="200" cy="20"/></p:xfrm>',
      grid: '<a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>',
      rows: '<a:tr h="10">' + cell(' rowSpan="2"', 'Top') + cell('', 'P') +
        '</a:tr><a:tr h="10">' + cell(' vMerge="1" keep="HIDDEN"', 'Hidden') +
        cell('', 'Q') + '</a:tr>',
    }));
    deleteTableRows(
      vertical.xml,
      vertical.frame,
      normalizeTableDeleteInput(0, 1, 'row'),
      PART_URI,
    );
    const verticalReopened = parseFrame(vertical.xml.serialize());
    const verticalState = requireEditableTableStructure(verticalReopened.frame, PART_URI);
    expect(verticalState.mergeState.regions).toEqual([]);
    expect(verticalReopened.xml.text(verticalState.cells[0]![0]!)).toBe('Hidden');
    expect(verticalReopened.xml.original(verticalState.cells[0]![0]!))
      .toContain('keep="HIDDEN"');
    expect(verticalReopened.xml.original(verticalState.cells[0]![0]!))
      .not.toMatch(/rowSpan|gridSpan|vMerge|hMerge/);

    const removed = parseFrame(validFrame({
      transform: '<p:xfrm><a:ext cx="100" cy="40"/></p:xfrm>',
      grid: '<a:tblGrid><a:gridCol w="100"/></a:tblGrid>',
      rows: '<a:tr h="10">' + cell('', 'Before') + '</a:tr>' +
        '<a:tr h="10">' + cell(' rowSpan="2"', 'Merged') + '</a:tr>' +
        '<a:tr h="10">' + cell(' vMerge="1"', 'Hidden') + '</a:tr>' +
        '<a:tr h="10">' + cell('', 'After') + '</a:tr>',
    }));
    deleteTableRows(
      removed.xml,
      removed.frame,
      normalizeTableDeleteInput(1, 2, 'row'),
      PART_URI,
    );
    const removedState = requireEditableTableStructure(
      parseFrame(removed.xml.serialize()).frame,
      PART_URI,
    );
    expect(removedState.mergeState.regions).toEqual([]);
    expect(removedState.rowHeights).toEqual([10, 10]);
    expect(removed.xml.serialize()).not.toContain('Merged');
    expect(removed.xml.serialize()).not.toContain('Hidden');
  });

  it('rejects out-of-range and last-row deletion without scheduling XML changes', () => {
    const fixture = parseFrame(plainFrame([40, 60], 2));
    for (const input of [
      normalizeTableDeleteInput(2, 1, 'row'),
      normalizeTableDeleteInput(1, 2, 'row'),
      normalizeTableDeleteInput(0, 2, 'row'),
    ]) {
      expect(() => deleteTableRows(
        fixture.xml,
        fixture.frame,
        input,
        PART_URI,
      )).toThrow();
      expect(fixture.xml.changed).toBe(false);
    }
  });
});
