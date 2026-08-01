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
