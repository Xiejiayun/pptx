import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { normalizeChartDefinition } from './chart-definition.internal.js';
import {
  buildChartWorkbook,
  chartWorkbookMatches,
  planChartWorkbook,
  readChartWorkbookCells,
} from './chart-workbook.internal.js';

const WORKSHEET_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const STYLES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

describe('chart workbook planning', () => {
  it('shares equal categorical columns and emits exact formula ranges', () => {
    const definition = normalizeChartDefinition({ groups: [{
      type: 'bar',
      series: [
        { name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] },
        { name: 'Cost', categories: ['Q1', 'Q2'], values: [7, 12] },
      ],
    }] });
    const plan = planChartWorkbook(definition);

    expect(plan.formulas).toEqual([
      {
        groupIndex: 0,
        seriesIndex: 0,
        name: 'Sheet1!$B$1',
        categories: ['Sheet1!$A$2:$A$3'],
        values: 'Sheet1!$B$2:$B$3',
      },
      {
        groupIndex: 0,
        seriesIndex: 1,
        name: 'Sheet1!$C$1',
        categories: ['Sheet1!$A$2:$A$3'],
        values: 'Sheet1!$C$2:$C$3',
      },
    ]);
    expect(Object.isFrozen(plan)).toBe(true);
    expect(Object.isFrozen(plan.formulas)).toBe(true);
    expect(Object.isFrozen(plan.formulas[0]?.categories)).toBe(true);
  });

  it('plans independent category blocks, multi-level categories, and combinations in order', () => {
    const definition = normalizeChartDefinition({ groups: [
      {
        type: 'bar',
        series: [
          {
            name: 'Revenue',
            categories: [['FY25', 'FY25'], ['Q1', 'Q2']],
            values: [10, 20],
          },
          { name: 'Cost', categories: ['Jan', 'Feb'], values: [7, 12] },
        ],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
      },
    ] });
    const plan = planChartWorkbook(definition);

    expect(plan.formulas).toEqual([
      {
        groupIndex: 0,
        seriesIndex: 0,
        name: 'Sheet1!$D$1',
        categories: ['Sheet1!$A$2:$A$3', 'Sheet1!$B$2:$B$3'],
        values: 'Sheet1!$D$2:$D$3',
      },
      {
        groupIndex: 0,
        seriesIndex: 1,
        name: 'Sheet1!$E$1',
        categories: ['Sheet1!$C$2:$C$3'],
        values: 'Sheet1!$E$2:$E$3',
      },
      {
        groupIndex: 1,
        seriesIndex: 0,
        name: 'Sheet1!$G$1',
        categories: ['Sheet1!$F$2:$F$3'],
        values: 'Sheet1!$G$2:$G$3',
      },
    ]);
  });

  it('plans separate X/Y/size vectors and Excel columns beyond Z', () => {
    const scatter = normalizeChartDefinition({ groups: [{
      type: 'scatter',
      series: Array.from({ length: 14 }, (_value, index) => ({
        name: `Series ${index + 1}`,
        xValues: [index, index + 1],
        values: [index + 2, index + 3],
      })),
    }] });
    const scatterPlan = planChartWorkbook(scatter);
    expect(scatterPlan.formulas[0]).toEqual({
      groupIndex: 0,
      seriesIndex: 0,
      name: 'Sheet1!$B$1',
      xValues: 'Sheet1!$A$2:$A$3',
      values: 'Sheet1!$B$2:$B$3',
    });
    expect(scatterPlan.formulas[13]).toEqual({
      groupIndex: 0,
      seriesIndex: 13,
      name: 'Sheet1!$AB$1',
      xValues: 'Sheet1!$AA$2:$AA$3',
      values: 'Sheet1!$AB$2:$AB$3',
    });

    const bubble = normalizeChartDefinition({ groups: [{
      type: 'bubble',
      series: [{ name: 'Bubbles', xValues: [1, 2], values: [3, 4], sizes: [5, 6] }],
    }] });
    expect(planChartWorkbook(bubble).formulas).toEqual([{
      groupIndex: 0,
      seriesIndex: 0,
      name: 'Sheet1!$B$1',
      xValues: 'Sheet1!$A$2:$A$3',
      values: 'Sheet1!$B$2:$B$3',
      sizes: 'Sheet1!$C$2:$C$3',
    }]);
  });

  it('escapes inline strings and preserves numeric zero in worksheet XML', () => {
    const definition = normalizeChartDefinition({ groups: [{
      type: 'bar',
      series: [{ name: 'A&B <total>', categories: [' Q1 ', 'Q2'], values: [0, -0] }],
    }] });
    const xml = planChartWorkbook(definition).worksheetXml;

    expect(xml).toContain('A&amp;B &lt;total&gt;');
    expect(xml).toContain('<t xml:space="preserve"> Q1 </t>');
    expect(xml.match(/<v>0<\/v>/g)).toHaveLength(2);
    expect(xml).not.toContain('<v>-0</v>');
  });
});

describe('chart workbook generation and strict readback', () => {
  const definition = normalizeChartDefinition({ groups: [{
    type: 'bar',
    series: [
      { name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] },
      { name: 'Cost', categories: ['Q1', 'Q2'], values: [7, 12] },
    ],
  }] });

  it('builds byte-identical minimal XLSX packages with exact typed cells', async () => {
    const first = await buildChartWorkbook(definition);
    const second = await buildChartWorkbook(definition);

    expect(second).toEqual(first);
    expect(await readChartWorkbookCells(first)).toEqual([
      ['', 'Revenue', 'Cost'],
      ['Q1', 10, 7],
      ['Q2', 20, 12],
    ]);
    expect(await chartWorkbookMatches(first, definition)).toBe(true);

    const workbook = await OpcPackage.open(first);
    expect(workbook.hasPart('/xl/workbook.xml')).toBe(true);
    expect(workbook.hasPart('/xl/worksheets/sheet1.xml')).toBe(true);
    expect(workbook.hasPart('/xl/styles.xml')).toBe(true);
    expect(workbook.hasPart('/xl/sharedStrings.xml')).toBe(false);
    expect(workbook.parts.some(({ uri }) => uri.startsWith('/docProps/'))).toBe(false);
    expect(workbook.relationships('/')[0]?.resolvedTarget).toBe('/xl/workbook.xml');
    expect(workbook.relationships('/xl/workbook.xml')).toMatchObject([
      { id: 'rId1', type: WORKSHEET_RELATIONSHIP, resolvedTarget: '/xl/worksheets/sheet1.xml' },
      { id: 'rId2', type: STYLES_RELATIONSHIP, resolvedTarget: '/xl/styles.xml' },
    ]);
    const sheetXml = new TextDecoder().decode(
      workbook.requirePart('/xl/worksheets/sheet1.xml').bytes,
    );
    expect(sheetXml).toContain('t="inlineStr"');
    expect(sheetXml).not.toContain('sharedStrings');
  });

  it('matches multi-level, scatter, bubble, and combination cell matrices', async () => {
    const definitions = [
      normalizeChartDefinition({ groups: [{
        type: 'bar',
        series: [{
          name: 'Revenue',
          categories: [['FY25', 'FY25'], ['Q1', 'Q2']],
          values: [10, 20],
        }],
      }] }),
      normalizeChartDefinition({ groups: [{
        type: 'scatter',
        series: [
          { name: 'First', xValues: [1, 2], values: [3, 4] },
          { name: 'Second', xValues: [5, 6], values: [7, 8] },
        ],
      }] }),
      normalizeChartDefinition({ groups: [{
        type: 'bubble',
        series: [{ name: 'Bubbles', xValues: [1, 2], values: [3, 4], sizes: [5, 6] }],
      }] }),
      normalizeChartDefinition({ groups: [
        {
          type: 'bar',
          series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
        },
        {
          type: 'line',
          axis: 'secondary',
          series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
        },
      ] }),
    ];
    const expected = [
      [['', '', 'Revenue'], ['FY25', 'Q1', 10], ['FY25', 'Q2', 20]],
      [['', 'First', '', 'Second'], [1, 3, 5, 7], [2, 4, 6, 8]],
      [['', 'Bubbles', ''], [1, 3, 5], [2, 4, 6]],
      [['', 'Revenue', '', 'Trend'], ['Q1', 10, 'Q1', 11], ['Q2', 20, 'Q2', 21]],
    ];

    for (const [index, candidate] of definitions.entries()) {
      const bytes = await buildChartWorkbook(candidate!);
      expect(await readChartWorkbookCells(bytes)).toEqual(expected[index]);
      expect(await chartWorkbookMatches(bytes, candidate!)).toBe(true);
    }
  });

  it('returns false for different data and malformed workbook grammar', async () => {
    const bytes = await buildChartWorkbook(definition);
    const changed = normalizeChartDefinition({ groups: [{
      type: 'bar',
      series: [
        { name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 99] },
        { name: 'Cost', categories: ['Q1', 'Q2'], values: [7, 12] },
      ],
    }] });
    expect(await chartWorkbookMatches(bytes, changed)).toBe(false);

    const malformedXml = [
      (xml: string) => xml.replace('<c r="B2"><v>10</v></c>', ''),
      (xml: string) => xml.replace(
        '<c r="B2"><v>10</v></c>',
        '<c r="B2"><f>1+1</f><v>10</v></c>',
      ),
      (xml: string) => xml.replace('</sheetData>', '</sheetData><mergeCells count="1"/>'),
    ];
    for (const mutate of malformedXml) {
      const malformed = await replaceWorksheet(bytes, mutate);
      await expect(readChartWorkbookCells(malformed)).rejects.toThrow();
      expect(await chartWorkbookMatches(malformed, definition)).toBe(false);
    }

    const sharedStrings = await OpcPackage.open(bytes);
    sharedStrings.setPart(
      '/xl/sharedStrings.xml',
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
    );
    const sharedStringBytes = await sharedStrings.write();
    await expect(readChartWorkbookCells(sharedStringBytes)).rejects.toThrow(/shared strings/);
  });
});

async function replaceWorksheet(
  bytes: Uint8Array,
  mutate: (xml: string) => string,
): Promise<Uint8Array> {
  const workbook = await OpcPackage.open(bytes);
  const part = workbook.requirePart('/xl/worksheets/sheet1.xml');
  workbook.setPart('/xl/worksheets/sheet1.xml', mutate(new TextDecoder().decode(part.bytes)));
  return workbook.write();
}
