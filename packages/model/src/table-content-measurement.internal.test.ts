import { describe, expect, it } from 'vitest';
import {
  normalizeTableDefinition,
  type NormalizedTableCell,
} from './table-create.internal.js';
import {
  materializeMeasuredTableRows,
  measureTableCellContent,
  measureTableAutoPageRows,
  type MeasuredTableCellContent,
  type MeasuredTableRowLayout,
} from './table-content-measurement.internal.js';

const EMU_PER_POINT = 12_700;
const EMU_PER_INCH = 914_400;
const BASE_ADVANCE = Math.round(12 * EMU_PER_POINT / 2.3);
const BASE_LINE_HEIGHT = Math.round(12 * 1.67 * EMU_PER_INCH / 100);

function cell(
  value: unknown,
  tableOptions: Readonly<Record<string, unknown>> = {},
): NormalizedTableCell {
  const input = Array.isArray(value) ? { text: value } : value;
  return normalizeTableDefinition([[input]], tableOptions).rows[0]![0]!;
}

function lineTexts(measured: Readonly<MeasuredTableCellContent>): readonly string[] {
  return measured.lines.map(({ slices }) => slices.map(({ text }) => text).join(''));
}

function zeroMarginCell(text: unknown): NormalizedTableCell {
  return cell(text, { margin: 0 });
}

function syntheticContent(height: number): Readonly<MeasuredTableCellContent> {
  return {
    paragraphs: [],
    lines: [{
      paragraphIndex: 0,
      slices: [],
      height,
      startsParagraph: true,
      endsParagraph: true,
    }],
    topMargin: 0,
    rightMargin: 0,
    bottomMargin: 0,
    leftMargin: 0,
  };
}

function syntheticRow(
  sourceRowIndex: number,
  height: number,
  cells: readonly (Readonly<MeasuredTableCellContent> | undefined)[],
): Readonly<MeasuredTableRowLayout> {
  return {
    sourceRowIndex,
    bands: [],
    contentHeight: 0,
    height,
    fragmentable: false,
    cells,
  };
}

describe('deterministic table-cell content measurement', () => {
  it('wraps at the exact ASCII boundary without losing whitespace or long-word text', () => {
    const width = 5 * BASE_ADVANCE;
    expect(lineTexts(measureTableCellContent(zeroMarginCell('Alpha'), width)))
      .toEqual(['Alpha']);
    expect(lineTexts(measureTableCellContent(zeroMarginCell('Alpha '), width)))
      .toEqual(['Alpha', ' ']);
    expect(lineTexts(measureTableCellContent(zeroMarginCell('Alphabet'), width)))
      .toEqual(['Alpha', 'bet']);

    const positive = measureTableCellContent(zeroMarginCell('AAAAAA'), width, 1);
    const negative = measureTableCellContent(zeroMarginCell('AAAAAA'), width, -1);
    expect(positive.lines).toHaveLength(1);
    expect(negative.lines.length).toBeGreaterThan(1);
  });

  it.each([
    ['ASCII punctuation', '!!', '!', 0.6],
    ['ASCII digits', '11', '1', 1],
    ['Latin', 'AA', 'A', 1],
    ['Greek', 'ΩΩ', 'Ω', 1],
    ['Cyrillic', 'ЖЖ', 'Ж', 1],
    ['CJK', '漢漢', '漢', 2.3],
    ['surrogate emoji', '😀😀', '😀', 2.3],
    ['combining marks', 'a\u0301a\u0301', 'a\u0301', 1],
    ['variation selectors', '✈️✈️', '✈️', 2.3],
    ['emoji modifiers', '👍🏽👍🏽', '👍🏽', 2.3],
    ['ZWJ families', '👨‍👩‍👧‍👦👨‍👩‍👧‍👦', '👨‍👩‍👧‍👦', 2.3],
    ['whitespace', '  ', ' ', 0.5],
  ])('keeps %s clusters intact', (_name, text, cluster, units) => {
    const width = Math.round(12 * EMU_PER_POINT * units / 2.3);
    const measured = measureTableCellContent(zeroMarginCell(text), width);
    expect(lineTexts(measured)).toEqual([cluster, cluster]);
    expect(measured.lines.flatMap(({ slices }) => slices.map(({ text: slice }) => slice)))
      .toEqual([cluster, cluster]);
  });

  it('advances tabs to declared stops or the next four-space grid', () => {
    const text = [{
      runs: [{ text: 'A\tB' }],
      tabStops: [{ position: 0.1 }],
    }];
    const declared = measureTableCellContent(zeroMarginCell(text), 170_000);
    const fallback = measureTableCellContent(zeroMarginCell([{
      runs: [{ text: 'A\tB' }],
    }]), 170_000);
    expect(lineTexts(declared)).toEqual(['A\tB']);
    expect(lineTexts(fallback)).toEqual(['A\t', 'B']);
  });

  it('preserves empty runs, empty paragraphs, paragraph boundaries, and soft breaks', () => {
    const source = cell([
      {
        runs: [
          { text: '' },
          { text: 'A' },
          { text: 'B', softBreakBefore: true },
        ],
      },
      { runs: [] },
      {
        runs: [
          { text: 'C', breakLine: true },
          { text: 'D' },
        ],
      },
    ], { margin: 0 });
    const measured = measureTableCellContent(source, 20 * BASE_ADVANCE);

    expect(measured.paragraphs).toHaveLength(4);
    expect(lineTexts(measured)).toEqual(['A', 'B', '', 'C', 'D']);
    expect(measured.lines[0]).toMatchObject({ startsParagraph: true, endsParagraph: false });
    expect(measured.lines[1]).toMatchObject({ startsParagraph: false, endsParagraph: true });
    expect(measured.lines[1]!.slices[0]).toMatchObject({
      text: 'B',
      retainsSoftBreak: true,
      startsAtRunStart: true,
      endsAtRunEnd: true,
    });
    expect(measured.lines[2]).toMatchObject({
      slices: [],
      startsParagraph: true,
      endsParagraph: true,
    });

    for (const [paragraphIndex, paragraph] of measured.paragraphs.entries()) {
      for (const [runIndex, run] of paragraph.runs.entries()) {
        const reproduced = measured.lines
          .flatMap(({ slices }) => slices)
          .filter((slice) =>
            slice.paragraphIndex === paragraphIndex && slice.runIndex === runIndex)
          .map(({ text }) => text)
          .join('');
        expect(reproduced).toBe(run.text);
      }
    }
  });

  it('uses run, cell, and default font sizes for natural line height', () => {
    const mixed = measureTableCellContent(cell([{
      runs: [
        { text: 'A', style: { fontSize: 10 } },
        { text: 'B', style: { fontSize: 20 } },
      ],
    }], { margin: 0 }), 20 * BASE_ADVANCE);
    const fallback = measureTableCellContent(cell('A', {
      margin: 0,
      fontSize: 18,
    }), 20 * BASE_ADVANCE);
    const defaulted = measureTableCellContent(zeroMarginCell('A'), 20 * BASE_ADVANCE);

    expect(mixed.lines[0]!.height)
      .toBe(Math.round(20 * 1.67 * EMU_PER_INCH / 100));
    expect(fallback.lines[0]!.height)
      .toBe(Math.round(18 * 1.67 * EMU_PER_INCH / 100));
    expect(defaulted.lines[0]!.height).toBe(BASE_LINE_HEIGHT);
  });

  it('applies character spacing and cell-over-table weight precedence', () => {
    const width = 2 * BASE_ADVANCE;
    const positiveSpacing = measureTableCellContent(cell([{
      runs: [{ text: 'AA', style: { characterSpacing: 1 } }],
    }], { margin: 0 }), width);
    const negativeSpacing = measureTableCellContent(cell([{
      runs: [{ text: 'AA', style: { characterSpacing: -1 } }],
    }], { margin: 0 }), width);
    expect(positiveSpacing.lines).toHaveLength(2);
    expect(negativeSpacing.lines).toHaveLength(1);

    const overridden = cell({
      text: 'AAAAAA',
      options: { autoPageCharWeight: 1, autoPageLineWeight: -1 },
    }, { autoPage: true, rowHeights: [1], margin: 0 });
    const measured = measureTableCellContent(overridden, 5 * BASE_ADVANCE, -1, 1);
    expect(measured.lines).toHaveLength(1);
    expect(measured.lines[0]!.height)
      .toBe(Math.round(12 * 0.67 * EMU_PER_INCH / 100));
  });

  it('resolves exact, multiple, before, and after paragraph spacing', () => {
    const exact = measureTableCellContent(cell([{
      runs: [{ text: 'A' }],
      spacing: {
        before: 2,
        after: 3,
        line: { kind: 'exact', points: 20 },
      },
    }], { margin: 0 }), 20 * BASE_ADVANCE, undefined, 1);
    expect(exact.lines[0]!.height).toBe(25 * EMU_PER_POINT);

    const multiple = measureTableCellContent(cell([{
      runs: [{ text: 'AAAAAA' }],
      spacing: {
        before: 2,
        after: 3,
        line: { kind: 'multiple', factor: 2 },
      },
    }], { margin: 0 }), 3 * BASE_ADVANCE);
    expect(multiple.lines.map(({ height }) => height)).toEqual([
      (2 * BASE_LINE_HEIGHT) + (2 * EMU_PER_POINT),
      (2 * BASE_LINE_HEIGHT) + (3 * EMU_PER_POINT),
    ]);

    const inherited = measureTableCellContent(cell('A', {
      margin: 0,
      spacing: { line: { kind: 'exact', points: 9 } },
    }), 20 * BASE_ADVANCE);
    expect(inherited.lines[0]!.height).toBe(9 * EMU_PER_POINT);
  });

  it('accounts for paragraph margins, indent, bullet indent, and colspan width', () => {
    const plain = measureTableCellContent(zeroMarginCell([{
      runs: [{ text: 'AAAAA' }],
    }]), 5 * BASE_ADVANCE);
    const left = measureTableCellContent(zeroMarginCell([{
      runs: [{ text: 'AAAAA' }],
      marginLeft: 1,
    }]), 5 * BASE_ADVANCE);
    const right = measureTableCellContent(zeroMarginCell([{
      runs: [{ text: 'AAAAA' }],
      marginRight: 1,
    }]), 5 * BASE_ADVANCE);
    const indented = measureTableCellContent(zeroMarginCell([{
      runs: [{ text: 'AAAAA' }],
      indent: 1,
    }]), 5 * BASE_ADVANCE);
    const bulleted = measureTableCellContent(zeroMarginCell([{
      runs: [{ text: 'AAAAA' }],
      bullet: { kind: 'bullet', indent: 1 },
    }]), 5 * BASE_ADVANCE);
    expect(plain.lines).toHaveLength(1);
    expect(left.lines.length).toBeGreaterThan(1);
    expect(right.lines.length).toBeGreaterThan(1);
    expect(indented.lines.length).toBeGreaterThan(1);
    expect(bulleted.lines.length).toBeGreaterThan(1);

    const colspanCell = cell({
      text: 'AAAAAA',
      options: { colspan: 2, margin: 0 },
    });
    expect(colspanCell.colspan).toBe(2);
    expect(measureTableCellContent(colspanCell, 3 * BASE_ADVANCE).lines).toHaveLength(2);
    expect(measureTableCellContent(colspanCell, 6 * BASE_ADVANCE).lines).toHaveLength(1);
  });

  it('uses canonical or custom cell margins and ignores viewer-only direction and fit', () => {
    const canonical = measureTableCellContent(cell('A'), 1_000_000);
    expect(canonical).toMatchObject({
      topMargin: Math.round(3.6 * EMU_PER_POINT),
      rightMargin: Math.round(7.2 * EMU_PER_POINT),
      bottomMargin: Math.round(3.6 * EMU_PER_POINT),
      leftMargin: Math.round(7.2 * EMU_PER_POINT),
    });

    const directedSource = cell({
      text: 'A',
      options: {
        margin: [1, 2, 3, 4],
        textDirection: 'vert270',
        fit: 'shrink',
      },
    });
    const custom = measureTableCellContent(directedSource, 1_000_000);
    expect(custom).toMatchObject({
      topMargin: 1 * EMU_PER_POINT,
      rightMargin: 2 * EMU_PER_POINT,
      bottomMargin: 3 * EMU_PER_POINT,
      leftMargin: 4 * EMU_PER_POINT,
    });
    expect(lineTexts(custom)).toEqual(lineTexts(canonical));
    expect(directedSource).toMatchObject({
      textDirection: 'vert270',
      textFit: 'shrink',
    });
  });

  it('rejects unusable inline widths and safe-integer overflow without mutation', () => {
    const source = cell('A');
    const snapshot = JSON.stringify(source);
    expect(() => measureTableCellContent(source, 0)).toThrow(/width/i);
    expect(() => measureTableCellContent(source, 100)).toThrow(/inline width/i);
    expect(() => measureTableCellContent(zeroMarginCell([{
      runs: [{ text: 'A' }],
      marginLeft: 1,
      indent: -1,
    }]), EMU_PER_POINT)).toThrow(/inline width/i);
    expect(JSON.stringify(source)).toBe(snapshot);

    const minimumMargin = -2_147_483_648 / EMU_PER_POINT;
    const overflow = cell({
      text: 'A',
      options: {
        margin: {
          top: 0,
          right: minimumMargin,
          bottom: 0,
          left: minimumMargin,
        },
      },
    });
    const overflowSnapshot = JSON.stringify(overflow);
    expect(() => measureTableCellContent(overflow, Number.MAX_SAFE_INTEGER))
      .toThrow(/safe integer/i);
    expect(JSON.stringify(overflow)).toBe(overflowSnapshot);
  });

  it('returns detached deeply frozen paragraphs, lines, and slices', () => {
    const source = cell([{
      runs: [{ text: 'AB', style: { fontSize: 10, bold: true } }],
    }], { margin: 0 });
    const measured = measureTableCellContent(source, BASE_ADVANCE);

    expect(Object.isFrozen(measured)).toBe(true);
    expect(Object.isFrozen(measured.paragraphs)).toBe(true);
    expect(Object.isFrozen(measured.paragraphs[0])).toBe(true);
    expect(Object.isFrozen(measured.paragraphs[0]!.runs)).toBe(true);
    expect(Object.isFrozen(measured.paragraphs[0]!.runs[0])).toBe(true);
    expect(Object.isFrozen(measured.paragraphs[0]!.runs[0]!.style)).toBe(true);
    expect(Object.isFrozen(measured.lines)).toBe(true);
    expect(Object.isFrozen(measured.lines[0])).toBe(true);
    expect(Object.isFrozen(measured.lines[0]!.slices)).toBe(true);
    expect(Object.isFrozen(measured.lines[0]!.slices[0])).toBe(true);
    expect(measured.paragraphs).not.toBe(source.richText);
  });
});

describe('measured table row bands and materialization', () => {
  it('builds ordinal max bands with effective margins and positive minimums', () => {
    const line20 = Math.round(20 * 1.67 * EMU_PER_INCH / 100);
    const definition = normalizeTableDefinition([[
      {
        text: 'AAAA',
        options: { margin: [1, 0, 2, 0] },
      },
      {
        text: [{ runs: [{ text: 'A', style: { fontSize: 20 } }] }],
        options: { margin: [3, 0, 1, 0] },
      },
    ]], {
      autoPage: true,
      columnWidths: [2 * BASE_ADVANCE, 2 * BASE_ADVANCE],
      rowHeights: [0],
    });
    const measured = measureTableAutoPageRows(definition);

    expect(measured[0]!.bands).toEqual([line20, BASE_LINE_HEIGHT]);
    expect(measured[0]!.contentHeight).toBe(
      (3 * EMU_PER_POINT) + line20 + BASE_LINE_HEIGHT + (2 * EMU_PER_POINT),
    );
    expect(measured[0]!.height).toBe(measured[0]!.contentHeight);
    expect(measured[0]!.fragmentable).toBe(true);
    expect(measured[0]!.cells.every((entry) => entry !== undefined)).toBe(true);
    expect(Object.isFrozen(measured)).toBe(true);
    expect(Object.isFrozen(measured[0])).toBe(true);
    expect(Object.isFrozen(measured[0]!.bands)).toBe(true);
    expect(Object.isFrozen(measured[0]!.cells)).toBe(true);

    const minimum = measured[0]!.contentHeight + 123;
    const minimumDefinition = normalizeTableDefinition([['A', 'B']], {
      autoPage: true,
      autoPageCharWeight: 0,
      columnWidths: [4 * BASE_ADVANCE, 4 * BASE_ADVANCE],
      rowHeights: [minimum],
    });
    expect(measureTableAutoPageRows(minimumDefinition)[0]!.height).toBe(minimum);
  });

  it('measures empty anchors and exact colspan width but excludes continuations', () => {
    const empty = normalizeTableDefinition([['', 'A']], {
      autoPage: true,
      margin: 0,
      columnWidths: [BASE_ADVANCE, BASE_ADVANCE],
      rowHeights: [0],
    });
    expect(measureTableAutoPageRows(empty)[0]!.bands).toEqual([BASE_LINE_HEIGHT]);

    const colspan = normalizeTableDefinition([
      [{ text: 'AAAAAA', options: { colspan: 2, margin: 0 } }],
      ['A', 'B'],
    ], {
      autoPage: true,
      columnWidths: [3 * BASE_ADVANCE, 3 * BASE_ADVANCE],
      rowHeights: [0, 0],
    });
    const measured = measureTableAutoPageRows(colspan);
    expect(measured[0]!.bands).toEqual([BASE_LINE_HEIGHT]);
    expect(measured[0]!.cells[0]).toBeDefined();
    expect(measured[0]!.cells[1]).toBeUndefined();
  });

  it('returns fixed definitions by identity and freezes measured geometry containers', () => {
    const fixed = normalizeTableDefinition([['A'], ['B']], {
      autoPage: true,
      rowHeights: [100, 200],
    });
    expect(materializeMeasuredTableRows(fixed, measureTableAutoPageRows(fixed)))
      .toBe(fixed);

    const automatic = normalizeTableDefinition([['A'], ['B']], {
      autoPage: true,
      rowHeights: [0, 0],
    });
    const measured = measureTableAutoPageRows(automatic);
    const materialized = materializeMeasuredTableRows(automatic, measured);
    expect(materialized).not.toBe(automatic);
    expect(materialized.rows).not.toBe(automatic.rows);
    expect(materialized.rows[0]).not.toBe(automatic.rows[0]);
    expect(materialized.rows[0]![0]).toBe(automatic.rows[0]![0]);
    expect(materialized.rowHeights.every((height) => height > 0)).toBe(true);
    expect(materialized.autoRowHeight).toBe(false);
    expect(materialized.height).toBe(
      materialized.rowHeights.reduce((sum, height) => sum + height, 0),
    );
    expect(materialized.autoPage?.measureContent).toBe(false);
    expect(Object.isFrozen(materialized)).toBe(true);
    expect(Object.isFrozen(materialized.rows)).toBe(true);
    expect(materialized.rows.every(Object.isFrozen)).toBe(true);
    expect(Object.isFrozen(materialized.rowHeights)).toBe(true);
    expect(Object.isFrozen(materialized.columnWidths)).toBe(true);
    expect(automatic.rowHeights).toEqual([0, 0]);
    expect(automatic.autoRowHeight).toBe(true);
    expect(automatic.autoPage?.measureContent).toBe(true);
  });

  it('distributes a two-row rowspan deficit from the first covered row', () => {
    const definition = normalizeTableDefinition([
      [{ text: 'Span', options: { rowspan: 2 } }, 'A'],
      ['B'],
      ['C', 'D'],
    ], {
      autoPage: true,
      rowHeights: [0, 0, 0],
    });
    const measured = [
      syntheticRow(0, 10, [syntheticContent(100), syntheticContent(1)]),
      syntheticRow(1, 10, [undefined, syntheticContent(1)]),
      syntheticRow(2, 11, [syntheticContent(1), syntheticContent(1)]),
    ];
    const materialized = materializeMeasuredTableRows(definition, measured);
    expect(materialized.rowHeights).toEqual([50, 50, 11]);
    expect(materialized.height).toBe(111);
  });

  it('orders nested rowspan lower bounds by span length without breaking shorter spans', () => {
    const definition = normalizeTableDefinition([
      [
        { text: 'Two', options: { rowspan: 2 } },
        { text: 'Three', options: { rowspan: 3 } },
        'A',
      ],
      ['B'],
      ['C', 'D'],
    ], {
      autoPage: true,
      rowHeights: [0, 0, 0],
    });
    const measured = [
      syntheticRow(0, 10, [
        syntheticContent(80),
        syntheticContent(111),
        syntheticContent(1),
      ]),
      syntheticRow(1, 10, [undefined, undefined, syntheticContent(1)]),
      syntheticRow(2, 11, [syntheticContent(1), undefined, syntheticContent(1)]),
    ];
    const materialized = materializeMeasuredTableRows(definition, measured);
    expect(materialized.rowHeights).toEqual([47, 47, 17]);
    expect(materialized.rowHeights[0]! + materialized.rowHeights[1]!).toBeGreaterThanOrEqual(80);
    expect(materialized.height).toBe(111);
  });

  it('gives continuation-only rows a one-EMU automatic minimum and rejects overflow', () => {
    const merged = normalizeTableDefinition([
      [{ text: 'Span', options: { rowspan: 2, colspan: 2, margin: 0 } }],
      [],
    ], {
      autoPage: true,
      rowHeights: [0, 0],
    });
    const measured = measureTableAutoPageRows(merged);
    expect(measured[1]).toMatchObject({
      bands: [],
      contentHeight: 0,
      height: 1,
      fragmentable: false,
      cells: [undefined, undefined],
    });

    const ordinary = normalizeTableDefinition([['A'], ['B']], {
      autoPage: true,
      rowHeights: [0, 0],
    });
    const overflowing = [
      syntheticRow(0, Number.MAX_SAFE_INTEGER, [syntheticContent(1)]),
      syntheticRow(1, 1, [syntheticContent(1)]),
    ];
    expect(() => materializeMeasuredTableRows(ordinary, overflowing))
      .toThrow(/safe integer/i);
  });
});
