import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { normalizeChartDefinition } from './chart-definition.internal.js';
import { renderChartPart } from './chart-render.internal.js';
import { readChartState } from './chart-state.internal.js';
import { planChartWorkbook } from './chart-workbook.internal.js';

const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const CHART_URI = '/ppt/charts/chart1.xml';
const WORKBOOK_URI = '/ppt/embeddings/workbook1.xlsx';

describe('strict chart semantic state', () => {
  it('reads a canonical categorical chart and preserves package bytes and journal state', () => {
    const pkg = chartPackage(chartXml(
      groupXml('barChart', categoricalSeriesXml('Revenue', ['Q1', 'Q2'], [10, 20]), [10, 20]),
      axisXml([10, 20]),
    ));
    const before = packageSnapshot(pkg);

    const state = readChartState(pkg, CHART_URI);

    expect(state).toEqual({
      status: 'recognized',
      workbookPartUri: WORKBOOK_URI,
      definition: {
        groups: [{
          type: 'bar',
          axis: 'primary',
          series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
        }],
      },
    });
    expect(Object.isFrozen(state)).toBe(true);
    expect(Object.isFrozen(state.definition?.groups[0]?.series[0]?.values)).toBe(true);
    expect(packageSnapshot(pkg)).toEqual(before);
  });

  it('reads axis-free, scatter, bubble, and combination groups in semantic order', () => {
    const pie = readChartState(chartPackage(chartXml(
      groupXml('pieChart', categoricalSeriesXml('Share', ['A', 'B'], [60, 40])),
    )), CHART_URI);
    expect(pie.definition?.groups).toEqual([{
      type: 'pie',
      series: [{ name: 'Share', categories: ['A', 'B'], values: [60, 40] }],
    }]);

    const scatter = readChartState(chartPackage(chartXml(
      groupXml('scatterChart', pointSeriesXml('Samples', [1, 2], [3, 4]), [10, 20]),
      axisXml([10, 20], 'valAx', 'valAx'),
    )), CHART_URI);
    expect(scatter.definition?.groups[0]).toEqual({
      type: 'scatter',
      axis: 'primary',
      series: [{ name: 'Samples', values: [3, 4], xValues: [1, 2] }],
    });

    const bubble = readChartState(chartPackage(chartXml(
      groupXml('bubbleChart', pointSeriesXml('Bubbles', [1, 2], [3, 4], [5, 6]), [10, 20]),
      axisXml([10, 20], 'valAx', 'valAx'),
    )), CHART_URI);
    expect(bubble.definition?.groups[0]).toEqual({
      type: 'bubble',
      axis: 'primary',
      series: [{ name: 'Bubbles', values: [3, 4], xValues: [1, 2], sizes: [5, 6] }],
    });

    const combination = readChartState(chartPackage(chartXml(
      groupXml('barChart', categoricalSeriesXml('Revenue', ['Q1', 'Q2'], [10, 20]), [10, 20])
        + groupXml('lineChart', categoricalSeriesXml('Trend', ['Q1', 'Q2'], [11, 21]), [30, 40]),
      axisXml([10, 20]) + axisXml([30, 40]),
    )), CHART_URI);
    expect(combination.definition?.groups).toEqual([
      {
        type: 'bar',
        axis: 'primary',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
      },
    ]);
  });

  it('strictly reads renderer-produced scatter and bubble vectors', () => {
    const inputs = [
      normalizeChartDefinition({ groups: [{
        type: 'scatter',
        series: [
          { name: 'First', xValues: [0, 2], values: [1, 3] },
          { name: 'Second', xValues: [10, 20], values: [11, 21] },
        ],
      }] }),
      normalizeChartDefinition({ groups: [{
        type: 'bubble',
        series: [{ name: 'Bubbles', xValues: [1, 2], values: [3, 4], sizes: [5, 6] }],
      }] }),
    ];
    for (const definition of inputs) {
      const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');
      expect(readChartState(chartPackage(xml), CHART_URI)).toEqual({
        status: 'recognized',
        definition: {
          groups: [{
            ...definition.groups[0],
            axis: 'primary',
          }],
        },
        workbookPartUri: WORKBOOK_URI,
      });
    }
  });

  it('orders cache points by canonical idx and reads numeric and multi-level categories', () => {
    const series = '<c:ser><c:idx val="0"/><c:order val="0"/>'
      + stringReference('tx', 'Revenue', 'Sheet1!$B$1')
      + '<c:cat><c:numRef><c:f>Sheet1!$A$2:$A$3</c:f>'
      + '<c:numCache><c:ptCount val="2"/><c:pt idx="1"><c:v>2</c:v></c:pt>'
      + '<c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:cat>'
      + numericReference('val', [20, 10], 'Sheet1!$B$2:$B$3', [1, 0])
      + '</c:ser>';
    const state = readChartState(chartPackage(chartXml(
      groupXml('barChart', series, [10, 20]),
      axisXml([10, 20]),
    )), CHART_URI);
    expect(state.definition?.groups[0]?.series[0]).toEqual({
      name: 'Revenue', categories: [1, 2], values: [10, 20],
    });

    const multiLevel = '<c:ser><c:idx val="0"/><c:order val="0"/>'
      + stringReference('tx', 'Revenue', 'Sheet1!$C$1')
      + '<c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$3</c:f>'
      + '<c:multiLvlStrCache><c:ptCount val="2"/>'
      + `<c:lvl>${points(['FY25', 'FY25'])}</c:lvl>`
      + `<c:lvl>${points(['Q1', 'Q2'])}</c:lvl>`
      + '</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>'
      + numericReference('val', [10, 20], 'Sheet1!$C$2:$C$3')
      + '</c:ser>';
    expect(readChartState(chartPackage(chartXml(
      groupXml('barChart', multiLevel, [10, 20]),
      axisXml([10, 20]),
    )), CHART_URI).definition?.groups[0]?.series[0]?.categories).toEqual([
      ['FY25', 'FY25'],
      ['Q1', 'Q2'],
    ]);
  });

  it('distinguishes cache-only and modern charts', () => {
    const cacheOnly = chartPackage(chartXml(
      groupXml('pieChart', categoricalSeriesXml('Share', ['A'], [100])),
    ), { externalData: false });
    expect(readChartState(cacheOnly, CHART_URI)).toMatchObject({
      status: 'cache-only',
      definition: { groups: [{ type: 'pie' }] },
    });

    const modern = chartPackage(
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">'
        + '<cx:chart/></cx:chartSpace>',
      { externalData: false },
    );
    expect(readChartState(modern, CHART_URI)).toEqual({
      status: 'modern',
      reason: 'Chart uses the Office 2016 modern chart namespace',
    });
  });

  it.each([
    ['wrong namespace', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A'], [1]))).replaceAll(CHART_NS, 'urn:wrong'), 'unsupported'],
    ['multiple roots', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A'], [1]))) + '<c:chartSpace xmlns:c="' + CHART_NS + '"/>', 'ambiguous'],
    ['multiple charts', `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:r="${REL_NS}"><c:chart/><c:chart/></c:chartSpace>`, 'ambiguous'],
    ['multiple plot areas', `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:r="${REL_NS}"><c:chart><c:plotArea/><c:plotArea/></c:chart></c:chartSpace>`, 'ambiguous'],
    ['unknown chart group', chartXml('<c:stockChart/>'), 'unsupported'],
    ['duplicate series idx', chartXml(
      groupXml('pieChart', categoricalSeriesXml('A', ['A'], [1]) + categoricalSeriesXml('B', ['B'], [2]).replace('idx val="0"', 'idx val="0"').replace('order val="0"', 'order val="1"')),
    ),
      'ambiguous'],
    ['duplicate point idx', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A', 'B'], [1, 2]).replace('idx="1"', 'idx="0"'))), 'ambiguous'],
    ['negative point idx', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A'], [1]).replace('idx="0"', 'idx="-1"'))), 'unsupported'],
    ['decimal point idx', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A'], [1]).replace('idx="0"', 'idx="0.5"'))), 'unsupported'],
    ['point count mismatch', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A'], [1]).replace('ptCount val="1"', 'ptCount val="2"'))), 'unsupported'],
    ['non-finite number', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A'], [1]).replace('<c:v>1</c:v>', '<c:v>NaN</c:v>'))), 'unsupported'],
    ['malformed formula', chartXml(groupXml('pieChart', categoricalSeriesXml('S', ['A'], [1]).replace('Sheet1!$B$2:$B$2', 'OFFSET(A1,0,0)'))), 'unsupported'],
    ['duplicate axis id', chartXml(
      groupXml('barChart', categoricalSeriesXml('S', ['A'], [1]), [10, 20]),
      axisXml([10, 20]).replace('axId val="20"', 'axId val="10"')),
      'ambiguous'],
    ['dangling axis id', chartXml(
      groupXml('barChart', categoricalSeriesXml('S', ['A'], [1]), [10, 99]),
      axisXml([10, 20])),
      'unsupported'],
  ])('returns %s state for %s', (_label, xml, status) => {
    expect(readChartState(chartPackage(xml), CHART_URI)).toMatchObject({ status });
  });

  it('rejects repeated or invalid workbook links and ignores unrelated style and extension content', () => {
    const base = chartXml(
      groupXml('barChart', categoricalSeriesXml('S', ['A'], [1]), [10, 20]),
      axisXml([10, 20]) + '<c:extLst><c:ext uri="urn:keep"><x:opaque xmlns:x="urn:test"/></c:ext></c:extLst>',
    );
    expect(readChartState(chartPackage(base.replace(
      '</c:chartSpace>',
      '<c:externalData r:id="rId1"/></c:chartSpace>',
    )), CHART_URI)).toMatchObject({ status: 'ambiguous' });

    expect(readChartState(chartPackage(base, { relationshipType: 'urn:wrong' }), CHART_URI))
      .toMatchObject({ status: 'unsupported' });
    expect(readChartState(chartPackage(base, { externalRelationship: true }), CHART_URI))
      .toMatchObject({ status: 'unsupported' });
    expect(readChartState(chartPackage(base), CHART_URI)).toMatchObject({ status: 'recognized' });
  });
});

function chartPackage(
  xml: string,
  options: {
    readonly externalData?: boolean;
    readonly relationshipType?: string;
    readonly externalRelationship?: boolean;
  } = {},
): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.transaction(() => {
    pkg.setPart(CHART_URI, options.externalData === false
      ? xml.replace('<c:externalData r:id="rId1"/>', '')
      : xml, CHART_CONTENT_TYPE);
    pkg.setPart(WORKBOOK_URI, new Uint8Array([80, 75, 3, 4]), XLSX_CONTENT_TYPE);
    pkg.addRelationship(CHART_URI, {
      id: 'rId1',
      type: options.relationshipType ?? PACKAGE_REL,
      target: '../embeddings/workbook1.xlsx',
      ...(options.externalRelationship ? { targetMode: 'External' as const } : {}),
    });
  });
  return pkg;
}

function chartXml(groups: string, axes = ''): string {
  return `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:r="${REL_NS}">`
    + `<c:chart><c:plotArea>${groups}${axes}</c:plotArea></c:chart>`
    + '<c:externalData r:id="rId1"/><c:extLst><c:ext uri="urn:keep"/></c:extLst>'
    + '</c:chartSpace>';
}

function groupXml(localName: string, series: string, axisIds: readonly number[] = []): string {
  return `<c:${localName}>${series}`
    + axisIds.map((id) => `<c:axId val="${id}"/>`).join('')
    + `</c:${localName}>`;
}

function categoricalSeriesXml(name: string, categories: readonly string[], values: readonly number[]): string {
  return '<c:ser><c:idx val="0"/><c:order val="0"/>'
    + stringReference('tx', name, 'Sheet1!$B$1')
    + stringReference('cat', categories, `Sheet1!$A$2:$A$${categories.length + 1}`)
    + numericReference('val', values, `Sheet1!$B$2:$B$${values.length + 1}`)
    + '</c:ser>';
}

function pointSeriesXml(
  name: string,
  xValues: readonly number[],
  values: readonly number[],
  sizes?: readonly number[],
): string {
  return '<c:ser><c:idx val="0"/><c:order val="0"/>'
    + stringReference('tx', name, 'Sheet1!$A$1')
    + numericReference('xVal', xValues, `Sheet1!$A$2:$A$${xValues.length + 1}`)
    + numericReference('yVal', values, `Sheet1!$B$2:$B$${values.length + 1}`)
    + (sizes ? numericReference('bubbleSize', sizes, `Sheet1!$C$2:$C$${sizes.length + 1}`) : '')
    + '</c:ser>';
}

function stringReference(
  container: 'tx' | 'cat',
  value: string | readonly string[],
  formula: string,
): string {
  const values = typeof value === 'string' ? [value] : value;
  return `<c:${container}><c:strRef><c:f>${formula}</c:f><c:strCache>`
    + `<c:ptCount val="${values.length}"/>${points(values)}`
    + `</c:strCache></c:strRef></c:${container}>`;
}

function numericReference(
  container: 'val' | 'xVal' | 'yVal' | 'bubbleSize',
  values: readonly number[],
  formula: string,
  indexes = values.map((_value, index) => index),
): string {
  return `<c:${container}><c:numRef><c:f>${formula}</c:f><c:numCache>`
    + `<c:ptCount val="${values.length}"/>`
    + values.map((value, index) => `<c:pt idx="${indexes[index]}"><c:v>${value}</c:v></c:pt>`).join('')
    + `</c:numCache></c:numRef></c:${container}>`;
}

function points(values: readonly string[]): string {
  return values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('');
}

function axisXml(ids: readonly [number, number], first = 'catAx', second = 'valAx'): string {
  return `<c:${first}><c:axId val="${ids[0]}"/><c:crossAx val="${ids[1]}"/></c:${first}>`
    + `<c:${second}><c:axId val="${ids[1]}"/><c:crossAx val="${ids[0]}"/></c:${second}>`;
}

function packageSnapshot(pkg: OpcPackage): unknown {
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes, relationships }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
      relationships,
    })),
    graph: pkg.graph,
    mutations: [...pkg.mutations],
  };
}
