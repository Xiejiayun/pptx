import { describe, expect, it } from 'vitest';
import { degrees, inches } from './units.js';
import { normalizeChartDefinition } from './chart-definition.internal.js';
import { planChartWorkbook } from './chart-workbook.internal.js';
import {
  normalizeAddChartOptions,
  renderChartGraphicFrame,
  renderChartPart,
} from './chart-render.internal.js';
import type { ChartType } from './chart.js';

const CATEGORICAL_TYPES: readonly [ChartType, string, number][] = [
  ['area', 'areaChart', 2],
  ['bar', 'barChart', 2],
  ['bar3D', 'bar3DChart', 3],
  ['doughnut', 'doughnutChart', 0],
  ['line', 'lineChart', 2],
  ['pie', 'pieChart', 0],
  ['radar', 'radarChart', 2],
];

describe('categorical chart rendering', () => {
  it.each(CATEGORICAL_TYPES)('renders %s as one canonical %s group', (type, localName, axisCount) => {
    const definition = normalizeChartDefinition({ groups: [{
      type,
      series: [{ name: 'Revenue & Cost', categories: ['Q1', 'Q2'], values: [10, 20] }],
    }] });
    const plan = planChartWorkbook(definition);
    const xml = renderChartPart(definition, plan.formulas, 'rId7');

    expect(xml).toContain(`<c:${localName}>`);
    expect(xml).toContain('<c:idx val="0"/><c:order val="0"/>');
    expect(xml).toContain('<c:f>Sheet1!$B$1</c:f>');
    expect(xml).toContain('<c:v>Revenue &amp; Cost</c:v>');
    expect(xml).toContain('<c:f>Sheet1!$A$2:$A$3</c:f>');
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>');
    expect(xml).toContain('<c:ptCount val="2"/>');
    expect(xml).toContain(
      '<c:externalData r:id="rId7"><c:autoUpdate val="0"/></c:externalData>',
    );
    expect((xml.match(/<c:axId /g) ?? []).length).toBe(axisCount * 2);
    expect((xml.match(/<c:crossAx /g) ?? []).length).toBe(axisCount);
  });

  it('renders type-specific defaults, numeric categories, and multi-level caches', () => {
    const bar = definitionXml('bar');
    expect(bar).toContain('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>');
    expect(definitionXml('area')).toContain('<c:areaChart><c:grouping val="standard"/>');
    expect(definitionXml('line')).toContain('<c:lineChart><c:grouping val="standard"/>');
    expect(definitionXml('radar')).toContain('<c:radarChart><c:radarStyle val="standard"/>');
    expect(definitionXml('pie')).toContain('<c:pieChart><c:varyColors val="1"/>');
    expect(definitionXml('doughnut')).toContain(
      '<c:doughnutChart><c:varyColors val="1"/>',
    );
    expect(definitionXml('doughnut')).toContain('<c:holeSize val="50"/>');

    const numeric = normalizeChartDefinition({ groups: [{
      type: 'bar',
      series: [{ name: 'Numeric', categories: [1, 2], values: [3, 4] }],
    }] });
    expect(renderChartPart(numeric, planChartWorkbook(numeric).formulas, 'rId1'))
      .toContain('<c:cat><c:numRef>');

    const multi = normalizeChartDefinition({ groups: [{
      type: 'bar',
      series: [{
        name: 'Revenue',
        categories: [['FY25', 'FY25'], ['Q1', 'Q2']],
        values: [10, 20],
      }],
    }] });
    const multiXml = renderChartPart(multi, planChartWorkbook(multi).formulas, 'rId1');
    expect(multiXml).toContain('<c:multiLvlStrRef>');
    expect(multiXml).toContain('<c:f>Sheet1!$A$2:$B$3</c:f>');
    expect(multiXml.match(/<c:lvl>/g)).toHaveLength(2);
  });

  it('renders multiple series with stable idx/order and matching formula ownership', () => {
    const definition = normalizeChartDefinition({ groups: [{
      type: 'bar',
      series: [
        { name: 'Revenue', categories: ['Q1'], values: [10] },
        { name: 'Cost', categories: ['Q1'], values: [7] },
      ],
    }] });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');
    expect(xml).toContain('<c:idx val="0"/><c:order val="0"/>');
    expect(xml).toContain('<c:idx val="1"/><c:order val="1"/>');
    expect(xml).toContain('<c:f>Sheet1!$C$1</c:f>');
    expect(xml).toContain('<c:f>Sheet1!$C$2:$C$2</c:f>');
  });
});

describe('scatter and bubble chart rendering', () => {
  it('renders independent scatter X/Y formulas, caches, zeroes, and value axes', () => {
    const definition = normalizeChartDefinition({ groups: [{
      type: 'scatter',
      series: [
        { name: 'First', xValues: [0, 2, 4], values: [1, 3, 5] },
        { name: 'Second', xValues: [10, 20, 30], values: [6, 7, 8] },
      ],
    }] });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');

    expect(xml).toContain('<c:scatterChart><c:scatterStyle val="lineMarker"/>');
    expect(xml).toContain('<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>');
    expect(xml).toContain('<c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>');
    expect(xml).toContain('<c:xVal><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f>');
    expect(xml).toContain('<c:yVal><c:numRef><c:f>Sheet1!$D$2:$D$4</c:f>');
    expect(xml).toContain('<c:pt idx="0"><c:v>0</c:v></c:pt>');
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(xml).not.toContain('<c:catAx>');
  });

  it('renders bubble X/Y/size formulas and positive numeric caches', () => {
    const definition = normalizeChartDefinition({ groups: [{
      type: 'bubble',
      series: [{
        name: 'Bubbles',
        xValues: [1, 2, 3],
        values: [4, 5, 6],
        sizes: [7, 8, 9],
      }],
    }] });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');

    expect(xml).toContain('<c:bubbleChart><c:varyColors val="0"/>');
    expect(xml).toContain('<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>');
    expect(xml).toContain('<c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>');
    expect(xml).toContain('<c:bubbleSize><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f>');
    expect(xml).toContain('<c:bubble3D val="0"/>');
    expect(xml).toContain('<c:sizeRepresents val="area"/>');
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
  });
});

describe('combination chart rendering', () => {
  it('shares primary axes and allocates a deterministic secondary pair in group order', () => {
    const definition = normalizeChartDefinition({ groups: [
      {
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
      },
    ] });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');

    expect(xml.indexOf('<c:barChart>')).toBeLessThan(xml.indexOf('<c:lineChart>'));
    expect(xml).toMatch(
      /<c:barChart>[\s\S]*?<c:axId val="10000001"\/><c:axId val="10000002"\/><\/c:barChart>/,
    );
    expect(xml).toMatch(
      /<c:lineChart>[\s\S]*?<c:axId val="10000003"\/><c:axId val="10000004"\/><\/c:lineChart>/,
    );
    expect(xml).toContain('<c:idx val="0"/><c:order val="0"/>');
    expect(xml).toContain('<c:idx val="1"/><c:order val="1"/>');
    expect(xml).toContain('<c:f>Sheet1!$B$1</c:f>');
    expect(xml).toContain('<c:f>Sheet1!$D$1</c:f>');
    expect(xml.match(/<c:crossAx /g)).toHaveLength(4);
    expect(xml).toContain('<c:axPos val="t"/>');
    expect(xml).toContain('<c:axPos val="r"/>');
  });

  it('reuses one primary pair across primary-only and same-type groups', () => {
    const definition = normalizeChartDefinition({ groups: [
      {
        type: 'area',
        series: [{ name: 'Actual', categories: ['Q1'], values: [10] }],
      },
      {
        type: 'line',
        series: [{ name: 'Plan', categories: ['Q1'], values: [11] }],
      },
      {
        type: 'line',
        series: [{ name: 'Forecast', categories: ['Q1'], values: [12] }],
      },
    ] });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');
    expect(xml.match(/<c:axId val="10000001"\/>/g)).toHaveLength(4);
    expect(xml.match(/<c:axId val="10000002"\/>/g)).toHaveLength(4);
    expect(xml).not.toContain('10000003');
    expect(xml).toContain('<c:idx val="2"/><c:order val="2"/>');
  });

  it('renders scatter-only combinations with two value-axis pairs', () => {
    const definition = normalizeChartDefinition({ groups: [
      { type: 'scatter', series: [{ name: 'A', xValues: [1], values: [2] }] },
      {
        type: 'scatter',
        axis: 'secondary',
        series: [{ name: 'B', xValues: [3], values: [4] }],
      },
    ] });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');
    expect(xml.match(/<c:scatterChart>/g)).toHaveLength(2);
    expect(xml.match(/<c:valAx>/g)).toHaveLength(4);
    expect(xml).not.toContain('<c:catAx>');
  });
});

describe('chart option rendering', () => {
  it('renders root, title, legend, area, axes, labels, table, palette, and series options', () => {
    const definition = normalizeChartDefinition({
      groups: [{
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
        options: {
          direction: 'bar',
          grouping: 'stacked',
          gapWidth: 0,
          overlap: -25,
          varyColors: false,
          dataLabels: {
            showValue: true,
            showCategoryName: false,
            showSeriesName: true,
            position: 'insideEnd',
            numberFormat: '0.0%',
            showLeaderLines: false,
            face: 'Aptos',
            size: 9,
            bold: true,
            italic: false,
            color: { kind: 'scheme', value: 'tx1' },
          },
          series: [{
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: '112233' },
              transparency: 25,
            },
            line: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent1' },
              width: 2,
              dash: 'dash',
            },
          }],
        },
      }],
      options: {
        language: 'zh-CN',
        style: 48,
        roundedCorners: true,
        displayBlanksAs: 'zero',
        title: {
          visible: true,
          text: 'Revenue & Cost',
          overlay: true,
          rotation: -30,
          position: { x: 0.1, y: 0.2 },
          face: 'Aptos Display',
          size: 18,
          bold: true,
          italic: false,
          color: { kind: 'srgb', value: '445566' },
        } as never,
        legend: {
          visible: true,
          position: 'topRight',
          overlay: false,
          face: 'Aptos',
          size: 10,
          color: { kind: 'scheme', value: 'tx2' },
        },
        chartArea: {
          fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFFFFF' } },
          line: { kind: 'none' },
        },
        plotArea: {
          fill: { kind: 'none' },
          line: { kind: 'line', color: { kind: 'srgb', value: '999999' }, width: 1 },
        },
        categoryAxis: {
          visible: false,
          position: 'top',
          title: { text: 'Quarter', rotation: 0, size: 11 },
          minimum: 0,
          maximum: 10,
          majorUnit: 2,
          minorUnit: 1,
          numberFormat: '0',
          orientation: 'maxMin',
          labelPosition: 'high',
          labelRotation: -45,
          line: { kind: 'none' },
          majorGridLine: {
            kind: 'line',
            color: { kind: 'srgb', value: 'CCCCCC' },
            width: 0.5,
            dash: 'sysDot',
          },
          minorGridLine: { kind: 'none' },
          majorTickMark: 'inside',
          minorTickMark: 'cross',
          face: 'Aptos Narrow',
          size: 8,
        },
        valueAxis: {
          position: 'right',
          minimum: -5,
          maximum: 100,
          logarithmicBase: 10,
          numberFormat: '#,##0.00',
          majorTickMark: 'outside',
          minorTickMark: 'none',
        },
        dataTable: {
          visible: true,
          showHorizontalBorder: false,
          showVerticalBorder: true,
          showOutline: false,
          showLegendKeys: true,
          numberFormat: '#,##0.00',
          size: 9,
        },
        colors: [{ kind: 'srgb', value: '4472C4' }],
      },
    });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');

    expect(xml).toContain('<c:lang val="zh-CN"/>');
    expect(xml).toContain('<c:roundedCorners val="1"/><c:style val="48"/>');
    expect(xml).toContain('<a:t>Revenue &amp; Cost</a:t>');
    expect(xml).toContain(
      '<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/>'
      + '<c:x val="0.1"/><c:y val="0.2"/></c:manualLayout></c:layout>',
    );
    expect(xml).toContain('<c:legendPos val="tr"/>');
    expect(xml).toContain('<c:barDir val="bar"/><c:grouping val="stacked"/>');
    expect(xml).toContain('<c:gapWidth val="0"/><c:overlap val="-25"/>');
    expect(xml).toContain('<a:alpha val="75000"/>');
    expect(xml).toContain('<a:prstDash val="dash"/>');
    expect(xml).toContain('<c:dLblPos val="inEnd"/>');
    expect(xml).toContain('<c:showVal val="1"/>');
    expect(xml).toContain('<c:showSerName val="1"/>');
    expect(xml).toContain('<c:delete val="1"/><c:axPos val="t"/>');
    expect(xml).toContain('<c:logBase val="10"/>');
    expect(xml).toContain('<c:orientation val="maxMin"/>');
    expect(xml).toContain('<c:majorUnit val="2"/><c:minorUnit val="1"/>');
    expect(xml).toContain('<c:showHorzBorder val="0"/><c:showVertBorder val="1"/>');
    expect(xml).toContain('<c:showOutline val="0"/><c:showKeys val="1"/>');
    expect(xml).toContain('<c:numCache><c:formatCode>#,##0.00</c:formatCode>');
    expect(xml).not.toMatch(/<c:dTable>[\s\S]*?<c:numFmt[\s\S]*?<\/c:dTable>/);
    expect(xml).toContain('<c:dispBlanksAs val="zero"/>');
  });

  it.each([
    ['area', { grouping: 'percentStacked' }, '<c:grouping val="percentStacked"/>'],
    ['bar3D', { direction: 'column', grouping: 'stacked', gapDepth: 25 }, '<c:gapDepth val="25"/>'],
    ['bubble', { scale: 80, showNegativeBubbles: true, sizeRepresents: 'width' }, '<c:sizeRepresents val="width"/>'],
    ['doughnut', { firstSliceAngle: 45, holeSize: 60 }, '<c:holeSize val="60"/>'],
    ['line', { grouping: 'stacked', smooth: true, marker: { shape: 'diamond', size: 8 } }, '<c:symbol val="diamond"/>'],
    ['pie', { firstSliceAngle: 90 }, '<c:firstSliceAng val="90"/>'],
    ['radar', { style: 'filled' }, '<c:radarStyle val="filled"/>'],
    ['scatter', { style: 'smoothMarker', smooth: true, marker: { shape: 'triangle', size: 7 } }, '<c:scatterStyle val="smoothMarker"/>'],
  ] as const)('renders %s type-specific options', (type, options, expected) => {
    const series = type === 'scatter'
      ? [{ name: 'S', xValues: [1], values: [2] }]
      : type === 'bubble'
        ? [{ name: 'S', xValues: [1], values: [2], sizes: [3] }]
        : [{ name: 'S', categories: ['A'], values: [2] }];
    const definition = normalizeChartDefinition({
      groups: [{ type, series, options }],
    } as never);
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');
    expect(xml).toContain(expected);
  });

  it('renders explicit 3D view options for bar3D charts', () => {
    const definition = normalizeChartDefinition({
      groups: [{
        type: 'bar3D',
        series: [{ name: 'S', categories: ['A'], values: [1] }],
      }],
      options: { rotationX: -30, rotationY: 360, rightAngleAxes: false, perspective: 0 },
    });
    const xml = renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');
    expect(xml).toContain(
      '<c:view3D><c:rotX val="-30"/><c:rotY val="360"/>'
      + '<c:rAngAx val="0"/><c:perspective val="0"/></c:view3D>',
    );
  });
});

describe('chart graphic frame rendering', () => {
  it('normalizes, freezes, escapes, and renders exact frame placement', () => {
    const options = normalizeAddChartOptions({
      name: 'Revenue & Cost',
      altText: 'Quarterly <chart>',
      x: inches(2),
      y: inches(3),
      width: inches(7),
      height: inches(4),
      rotation: degrees(1),
      flipHorizontal: true,
      flipVertical: true,
    });
    const frame = renderChartGraphicFrame(5, 'rId9', options);

    expect(Object.isFrozen(options)).toBe(true);
    expect(frame).toContain(
      '<p:cNvPr id="5" name="Revenue &amp; Cost" descr="Quarterly &lt;chart&gt;"/>',
    );
    expect(frame).toContain('<p:xfrm rot="60000" flipH="1" flipV="1">');
    expect(frame).toContain(`<a:off x="${inches(2)}" y="${inches(3)}"/>`);
    expect(frame).toContain(`<a:ext cx="${inches(7)}" cy="${inches(4)}"/>`);
    expect(frame).toContain('<c:chart r:id="rId9"/>');
  });

  it('uses one-inch placement, a six-by-four-inch size, and rejects unsafe options', () => {
    expect(normalizeAddChartOptions()).toMatchObject({
      name: 'Chart',
      x: inches(1),
      y: inches(1),
      width: inches(6),
      height: inches(4),
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });

    const accessor = {};
    Object.defineProperty(accessor, 'name', { get: () => 'Unsafe', enumerable: true });
    for (const value of [
      { width: 0 },
      { height: Number.NaN },
      { rotation: 21_600_001 },
      { flipHorizontal: 1 },
      { altText: 'Bad\u0000text' },
      { unknown: true },
      accessor,
    ]) {
      expect(() => normalizeAddChartOptions(value as never)).toThrow();
    }
  });
});

function definitionXml(type: ChartType): string {
  const definition = normalizeChartDefinition({ groups: [{
    type,
    series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
  }] });
  return renderChartPart(definition, planChartWorkbook(definition).formulas, 'rId1');
}
