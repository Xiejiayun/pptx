import { describe, expect, it } from 'vitest';
import {
  normalizeTableDefinition,
  type NormalizedTableCell,
} from './table-create.internal.js';
import {
  measureTableCellContent,
  type MeasuredTableCellContent,
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
