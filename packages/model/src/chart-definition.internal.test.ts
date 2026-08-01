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

    expect(normalizeChartDefinition(input)).toEqual({ groups: [{
      type: 'bar',
      series: [{ name: 'Revenue', categories: ['Q1'], values: [10] }],
    }] });
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
