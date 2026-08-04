import { describe, expect, it } from 'vitest';
import { CHART_TYPES, type ChartDefinitionInput } from './chart.js';
import { normalizeChartDefinition } from './chart-definition.internal.js';

describe('chart definition normalization', () => {
  it('publishes the exact frozen chart catalog', () => {
    expect(CHART_TYPES).toEqual([
      'area',
      'bar',
      'bar3D',
      'bubble',
      'doughnut',
      'line',
      'pie',
      'radar',
      'scatter',
    ]);
    expect(Object.isFrozen(CHART_TYPES)).toBe(true);
  });

  it('deeply detaches and freezes categorical, multi-level, scatter, and bubble data', () => {
    const input: ChartDefinitionInput = {
      groups: [
        {
          type: 'bar',
          series: [{ name: 'Revenue', categories: ['Q1', 2], values: [10, -0] }],
        },
        {
          type: 'line',
          axis: 'secondary',
          series: [{
            name: 'Trend',
            categories: [['FY25', 'FY25'], ['Q1', 'Q2']],
            values: [11, 21],
          }],
        },
      ],
    };
    const normalized = normalizeChartDefinition(input);

    expect(normalized).toEqual({
      groups: [
        {
          type: 'bar',
          series: [{ name: 'Revenue', categories: ['Q1', 2], values: [10, 0] }],
        },
        {
          type: 'line',
          axis: 'secondary',
          series: [{
            name: 'Trend',
            categories: [['FY25', 'FY25'], ['Q1', 'Q2']],
            values: [11, 21],
          }],
        },
      ],
      options: {},
    });
    expect(normalized.groups[0]?.series[0]?.values[1]).toBe(0);
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(Object.isFrozen(normalized.groups)).toBe(true);
    expect(Object.isFrozen(normalized.groups[0])).toBe(true);
    expect(Object.isFrozen(normalized.groups[0]?.series)).toBe(true);
    expect(Object.isFrozen(normalized.groups[0]?.series[0])).toBe(true);
    expect(Object.isFrozen(normalized.groups[0]?.series[0]?.categories)).toBe(true);
    expect(Object.isFrozen(normalized.groups[0]?.series[0]?.values)).toBe(true);
    expect(Object.isFrozen(normalized.groups[1]?.series[0]?.categories?.[0])).toBe(true);

    (input.groups[0]!.series[0]!.values as number[])[0] = 999;
    expect(normalized.groups[0]?.series[0]?.values[0]).toBe(10);

    const scatter = normalizeChartDefinition({
      groups: [{
        type: 'scatter',
        series: [{ name: 'Samples', xValues: [1, 2], values: [3, 4] }],
      }],
    });
    expect(scatter.groups[0]?.series[0]).toEqual({
      name: 'Samples',
      values: [3, 4],
      xValues: [1, 2],
    });

    const bubble = normalizeChartDefinition({
      groups: [{
        type: 'bubble',
        series: [{ name: 'Samples', xValues: [1], values: [2], sizes: [3] }],
      }],
    });
    expect(bubble.groups[0]?.series[0]).toEqual({
      name: 'Samples',
      values: [2],
      xValues: [1],
      sizes: [3],
    });
  });

  it('accepts null-prototype data objects without retaining them', () => {
    const series = Object.assign(Object.create(null), {
      name: 'Revenue',
      categories: ['Q1'],
      values: [10],
    });
    const group = Object.assign(Object.create(null), { type: 'bar', series: [series] });
    const input = Object.assign(Object.create(null), { groups: [group] });

    expect(normalizeChartDefinition(input)).toEqual({
      groups: [{
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1'], values: [10] }],
      }],
      options: {},
    });
  });

  it('normalizes, detaches, and freezes chart, axis, label, series, and type options', () => {
    const input: ChartDefinitionInput = {
      groups: [{
        type: 'bar',
        series: [categoricalSeries() as never],
        options: {
          direction: 'bar',
          grouping: 'stacked',
          gapWidth: 0,
          overlap: -25,
          varyColors: false,
          dataLabels: {
            showValue: true,
            showCategoryName: false,
            position: 'insideEnd',
            numberFormat: '#,##0.00',
            color: { kind: 'scheme', value: 'tx1' },
          },
          series: [{
            fill: { kind: 'solid', color: { kind: 'srgb', value: '#112233' }, transparency: 25 },
            line: { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 0 },
          }],
        },
      }],
      options: {
        language: 'zh-CN',
        style: 48,
        roundedCorners: false,
        displayBlanksAs: 'zero',
        title: { visible: true, text: 'Revenue', overlay: false, rotation: 0, size: 18 },
        legend: { visible: true, position: 'topRight', overlay: true, bold: false },
        chartArea: { fill: { kind: 'none' }, line: { kind: 'none' } },
        plotArea: { fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFFFFF' } } },
        categoryAxis: {
          visible: true,
          position: 'bottom',
          minimum: 0,
          maximum: 10,
          majorUnit: 2,
          minorUnit: 1,
          labelRotation: -45,
          line: { kind: 'line', color: { kind: 'srgb', value: '000000' }, dash: 'dash' },
        },
        valueAxis: { logarithmicBase: 10, numberFormat: '0%', orientation: 'maxMin' },
        dataTable: { visible: true, showHorizontalBorder: false, size: 9 },
        colors: [{ kind: 'srgb', value: '4472C4' }, { kind: 'scheme', value: 'accent2' }],
        rightAngleAxes: true,
        rotationX: -30,
        rotationY: 360,
        perspective: 0,
      },
    };
    const normalized = normalizeChartDefinition(input);

    expect(normalized.options).toMatchObject({
      language: 'zh-CN',
      style: 48,
      displayBlanksAs: 'zero',
      title: { text: 'Revenue', size: 18, rotation: 0 },
      categoryAxis: { minimum: 0, maximum: 10, labelRotation: -45 },
      colors: [{ kind: 'srgb', value: '4472C4' }, { kind: 'scheme', value: 'accent2' }],
      rotationY: 360,
      perspective: 0,
    });
    expect(normalized.groups[0]?.options).toMatchObject({
      direction: 'bar', grouping: 'stacked', gapWidth: 0, overlap: -25,
    });
    expect(normalized.groups[0]?.options?.series?.[0]?.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 25,
    });
    expect(Object.isFrozen(normalized.options)).toBe(true);
    expect(Object.isFrozen(normalized.options.categoryAxis)).toBe(true);
    expect(Object.isFrozen(normalized.options.colors)).toBe(true);
    expect(Object.isFrozen(normalized.groups[0]?.options?.series?.[0]?.fill)).toBe(true);

    (input.options!.colors as { kind: 'srgb'; value: string }[])[0]!.value = 'FFFFFF';
    expect(normalized.options.colors?.[0]).toEqual({ kind: 'srgb', value: '4472C4' });
  });

  it('rejects unsafe chart option values before mutation planning', () => {
    const invalidOptions = [
      { style: 0 },
      { colors: [] },
      { title: { rotation: 91 } },
      { legend: { position: 'center' } },
      { categoryAxis: { minimum: 2, maximum: 1 } },
      { categoryAxis: { kind: 'category', baseTimeUnit: 'days' } },
      { categoryAxis: { labelFrequency: 0 } },
      { categoryAxis: { crossesAt: Number.POSITIVE_INFINITY } },
      { valueAxis: { logarithmicBase: 1 } },
      { valueAxis: { displayUnit: 'lakhs' } },
      { valueAxis: { displayUnitLabel: true } },
      { dataTable: { visible: 1 } },
      { rotationX: Number.NaN },
      { perspective: 241 },
      { unknown: true },
    ];
    for (const options of invalidOptions) {
      expect(() => normalizeChartDefinition({
        groups: [{ type: 'bar', series: [categoricalSeries() as never] }],
        options,
      } as never)).toThrow();
    }

    for (const options of [
      { gapWidth: 501 },
      { overlap: -101 },
      { firstSliceAngle: 1 },
      { marker: { shape: 'circle' } },
      { dataLabels: { position: 'top' } },
      { series: [{ marker: { shape: 'circle' } }] },
    ]) {
      expect(() => normalizeChartDefinition({
        groups: [{ type: 'bar', series: [categoricalSeries() as never], options }],
      } as never)).toThrow();
    }

    expect(() => normalizeChartDefinition({
      groups: [{
        type: 'bar3D',
        series: [categoricalSeries() as never],
        options: { overlap: 10 },
      }],
    } as never)).toThrow(/unsupported property overlap/);

    expect(() => normalizeChartDefinition({
      groups: [{ type: 'bar', series: [categoricalSeries() as never] }],
      options: { seriesAxis: { visible: false } },
    } as never)).toThrow(/require a bar3D chart/);

    const accessor = {};
    Object.defineProperty(accessor, 'style', { get: () => 1, enumerable: true });
    expect(() => normalizeChartDefinition({
      groups: [{ type: 'bar', series: [categoricalSeries() as never] }],
      options: accessor,
    } as never)).toThrow(/data property/);
  });

  it('normalizes strict category date value and series axis domains', () => {
    const normalized = normalizeChartDefinition({
      groups: [{ type: 'bar3D', series: [categoricalSeries() as never] }],
      options: {
        categoryAxis: {
          crossesAt: 0,
          baseTimeUnit: 'days',
          majorTimeUnit: 'months',
          minorTimeUnit: 'years',
          labelFrequency: 4,
          multiLevelLabels: true,
        },
        valueAxis: {
          crossesAt: 'autoZero',
          displayUnit: 'hundredMillions',
          displayUnitLabel: true,
        },
        seriesAxis: {
          visible: false,
          labelPosition: 'low',
          labelFrequency: 3,
          majorUnit: 2,
          minorUnit: 1,
          numberFormat: '0.0',
        },
      },
    });

    expect(normalized.options).toMatchObject({
      categoryAxis: {
        kind: 'date',
        crossesAt: 0,
        baseTimeUnit: 'days',
        majorTimeUnit: 'months',
        minorTimeUnit: 'years',
        labelFrequency: 4,
        multiLevelLabels: true,
      },
      valueAxis: {
        crossesAt: 'autoZero',
        displayUnit: 'hundredMillions',
        displayUnitLabel: true,
      },
      seriesAxis: {
        visible: false,
        labelPosition: 'low',
        labelFrequency: 3,
        majorUnit: 2,
        minorUnit: 1,
        numberFormat: '0.0',
      },
    });
    expect(Object.isFrozen(normalized.options.categoryAxis)).toBe(true);
    expect(Object.isFrozen(normalized.options.valueAxis)).toBe(true);
    expect(Object.isFrozen(normalized.options.seriesAxis)).toBe(true);
  });

  it('accepts compatible primary-only, primary-secondary, same-type, and scatter combinations', () => {
    for (const groups of [
      [
        { type: 'bar', series: [categoricalSeries()] },
        { type: 'line', series: [categoricalSeries()] },
      ],
      [
        { type: 'area', series: [categoricalSeries()] },
        { type: 'line', axis: 'secondary', series: [categoricalSeries()] },
      ],
      [
        { type: 'bar', series: [categoricalSeries()] },
        { type: 'bar', axis: 'secondary', series: [categoricalSeries()] },
      ],
      [
        { type: 'scatter', series: [pointSeries()] },
        { type: 'scatter', axis: 'secondary', series: [pointSeries()] },
      ],
    ]) {
      expect(normalizeChartDefinition({ groups } as never).groups).toHaveLength(2);
    }
  });

  it.each([
    ['empty groups', { groups: [] }],
    ['empty series', { groups: [{ type: 'bar', series: [] }] }],
    ['empty name', categorical({ name: '' })],
    ['missing categories', categorical({ categories: undefined }, true)],
    ['empty categories', categorical({ categories: [] })],
    ['unequal categorical vectors', categorical({ values: [1] })],
    ['unequal multi-level categories', categorical({ categories: [['FY25', 'FY25'], ['Q1']] })],
    ['non-finite values', categorical({ values: [1, Number.NaN] })],
    ['categorical x values', categorical({ xValues: [1, 2] })],
    ['scatter categories', pointDefinition('scatter', { categories: ['A', 'B'] })],
    ['scatter missing x values', pointDefinition('scatter', { xValues: undefined }, true)],
    ['unequal scatter vectors', pointDefinition('scatter', { xValues: [1] })],
    ['scatter sizes', pointDefinition('scatter', { sizes: [1, 2] })],
    ['bubble missing sizes', pointDefinition('bubble', { sizes: undefined }, true)],
    ['non-positive bubble size', pointDefinition('bubble', { sizes: [1, 0] })],
    ['multiple pie series', axisFreeDefinition('pie', 2)],
    ['multiple doughnut series', axisFreeDefinition('doughnut', 2)],
    ['secondary-only definition', {
      groups: [{ type: 'line', axis: 'secondary', series: [categoricalSeries()] }],
    }],
    ['secondary group before primary group', {
      groups: [
        { type: 'line', axis: 'secondary', series: [categoricalSeries()] },
        { type: 'bar', series: [categoricalSeries()] },
      ],
    }],
    ['unsafe axis-free combination', {
      groups: [
        { type: 'bar', series: [categoricalSeries()] },
        { type: 'pie', series: [categoricalSeries()] },
      ],
    }],
    ['unsafe scatter combination', {
      groups: [
        { type: 'bar', series: [categoricalSeries()] },
        { type: 'scatter', series: [pointSeries()] },
      ],
    }],
  ])('rejects %s', (_label, input) => {
    expect(() => normalizeChartDefinition(input as ChartDefinitionInput)).toThrow();
  });

  it('rejects unknown keys, symbols, accessors, sparse arrays, and exotic prototypes', () => {
    expect(() => normalizeChartDefinition({ groups: [], extra: true } as never)).toThrow(/unsupported property/);
    expect(() => normalizeChartDefinition({
      groups: [{ type: 'bar', series: [categoricalSeries()], [Symbol('extra')]: true }],
    } as never)).toThrow(/unsupported property/);

    const accessor = { type: 'bar', series: [categoricalSeries()] };
    Object.defineProperty(accessor, 'axis', { get: () => 'primary', enumerable: true });
    expect(() => normalizeChartDefinition({ groups: [accessor] } as never)).toThrow(/data property/);

    const sparse = new Array(2);
    sparse[1] = 10;
    expect(() => normalizeChartDefinition(categorical({ values: sparse }))).toThrow(/dense array/);

    class Definition {
      groups = [{ type: 'bar', series: [categoricalSeries()] }];
    }
    expect(() => normalizeChartDefinition(new Definition() as never)).toThrow(/ordinary object/);
  });

  it('rejects invalid type, axis, string, and array property descriptors', () => {
    expect(() => normalizeChartDefinition({
      groups: [{ type: 'stock', series: [categoricalSeries()] }],
    } as never)).toThrow(/type/);
    expect(() => normalizeChartDefinition({
      groups: [{ type: 'bar', axis: 'tertiary', series: [categoricalSeries()] }],
    } as never)).toThrow(/axis/);
    expect(() => normalizeChartDefinition(categorical({ name: 'Bad\u0000name' }))).toThrow(/XML/);

    const values = [1, 2];
    Object.defineProperty(values, '0', { get: () => 1, enumerable: true });
    expect(() => normalizeChartDefinition(categorical({ values }))).toThrow(/data property/);
  });
});

function categoricalSeries(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return { name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20], ...overrides };
}

function categorical(
  overrides: Record<string, unknown>,
  removeUndefined = false,
): ChartDefinitionInput {
  const series = categoricalSeries(overrides);
  if (removeUndefined) {
    for (const [key, value] of Object.entries(series)) if (value === undefined) delete series[key];
  }
  return { groups: [{ type: 'bar', series: [series as never] }] };
}

function pointSeries(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return { name: 'Samples', xValues: [1, 2], values: [3, 4], ...overrides };
}

function pointDefinition(
  type: 'scatter' | 'bubble',
  overrides: Record<string, unknown>,
  removeUndefined = false,
): ChartDefinitionInput {
  const series = pointSeries(type === 'bubble' ? { sizes: [5, 6], ...overrides } : overrides);
  if (removeUndefined) {
    for (const [key, value] of Object.entries(series)) if (value === undefined) delete series[key];
  }
  return { groups: [{ type, series: [series as never] }] };
}

function axisFreeDefinition(type: 'pie' | 'doughnut', count: number): ChartDefinitionInput {
  return { groups: [{ type, series: Array.from({ length: count }, () => categoricalSeries() as never) }] };
}
