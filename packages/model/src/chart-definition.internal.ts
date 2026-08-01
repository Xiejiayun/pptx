import {
  CHART_TYPES,
  type ChartCategories,
  type ChartDefinition,
  type ChartDefinitionInput,
  type ChartGroup,
  type ChartSeries,
  type ChartType,
} from './chart.js';
import {
  normalizeChartGroupOptions,
  normalizeChartOptions,
} from './chart-options.internal.js';

const ROOT_KEYS = new Set(['groups', 'options']);
const GROUP_KEYS = new Set(['type', 'series', 'axis', 'options']);
const SERIES_KEYS = new Set(['name', 'categories', 'values', 'xValues', 'sizes']);
const CHART_TYPE_SET = new Set<string>(CHART_TYPES);
const CATEGORICAL_TYPES = new Set<ChartType>([
  'area',
  'bar',
  'bar3D',
  'doughnut',
  'line',
  'pie',
  'radar',
]);
const COMBINABLE_CATEGORICAL_TYPES = new Set<ChartType>(['area', 'bar', 'line', 'radar']);

export function normalizeChartDefinition(
  value: ChartDefinitionInput,
): Readonly<ChartDefinition> {
  const root = readObject(value, ROOT_KEYS, new Set(['groups']), 'Chart definition');
  const groups = readArray(root.groups, 'Chart groups').map((group, groupIndex) =>
    normalizeGroup(group, groupIndex));
  if (groups.length === 0) {
    throw new RangeError('Chart definition requires at least one group');
  }
  validateGroupCompatibility(groups);
  return Object.freeze({
    groups: Object.freeze(groups),
    options: normalizeChartOptions(root.options),
  });
}

function normalizeGroup(value: unknown, groupIndex: number): Readonly<ChartGroup> {
  const context = `Chart group ${groupIndex}`;
  const group = readObject(value, GROUP_KEYS, new Set(['type', 'series']), context);
  const type = normalizeChartType(group.type, `${context} type`);
  const series = readArray(group.series, `${context} series`).map((entry, seriesIndex) =>
    normalizeSeries(entry, type, groupIndex, seriesIndex));
  if (series.length === 0) throw new RangeError(`${context} requires at least one series`);
  if ((type === 'pie' || type === 'doughnut') && series.length !== 1) {
    throw new RangeError(`${context} ${type} charts require exactly one series`);
  }
  const axis = Object.hasOwn(group, 'axis')
    ? normalizeAxis(group.axis, `${context} axis`)
    : undefined;
  if ((type === 'pie' || type === 'doughnut') && axis !== undefined) {
    throw new TypeError(`${context} ${type} charts do not use an axis assignment`);
  }
  const options = normalizeChartGroupOptions(type, group.options, series.length);
  return Object.freeze({
    type,
    series: Object.freeze(series),
    ...(axis === undefined ? {} : { axis }),
    ...(options === undefined ? {} : { options }),
  }) as Readonly<ChartGroup>;
}

function normalizeSeries(
  value: unknown,
  type: ChartType,
  groupIndex: number,
  seriesIndex: number,
): Readonly<ChartSeries> {
  const context = `Chart group ${groupIndex} series ${seriesIndex}`;
  const series = readObject(value, SERIES_KEYS, new Set(['name', 'values']), context);
  const name = normalizeXmlString(series.name, `${context} name`, true);
  const values = normalizeNumberArray(series.values, `${context} values`);
  if (values.length === 0) throw new RangeError(`${context} values must not be empty`);
  const categories = Object.hasOwn(series, 'categories')
    ? normalizeCategories(series.categories, `${context} categories`)
    : undefined;
  const xValues = Object.hasOwn(series, 'xValues')
    ? normalizeNumberArray(series.xValues, `${context} xValues`)
    : undefined;
  const sizes = Object.hasOwn(series, 'sizes')
    ? normalizeNumberArray(series.sizes, `${context} sizes`)
    : undefined;

  if (CATEGORICAL_TYPES.has(type)) {
    if (!categories) throw new TypeError(`${context} categories are required for ${type}`);
    if (xValues !== undefined || sizes !== undefined) {
      throw new TypeError(`${context} xValues and sizes are not supported for ${type}`);
    }
    const categoryCount = categoryPointCount(categories);
    if (categoryCount !== values.length) {
      throw new RangeError(`${context} categories and values must have equal lengths`);
    }
  } else {
    if (categories !== undefined) {
      throw new TypeError(`${context} categories are not supported for ${type}`);
    }
    if (!xValues) throw new TypeError(`${context} xValues are required for ${type}`);
    if (xValues.length === 0 || xValues.length !== values.length) {
      throw new RangeError(`${context} xValues and values must have equal non-zero lengths`);
    }
    if (type === 'scatter' && sizes !== undefined) {
      throw new TypeError(`${context} sizes are not supported for scatter`);
    }
    if (type === 'bubble') {
      if (!sizes) throw new TypeError(`${context} sizes are required for bubble`);
      if (sizes.length !== values.length) {
        throw new RangeError(`${context} sizes and values must have equal lengths`);
      }
      if (sizes.some((size) => size <= 0)) {
        throw new RangeError(`${context} sizes must be positive`);
      }
    }
  }

  return Object.freeze({
    name,
    ...(categories === undefined ? {} : { categories }),
    values,
    ...(xValues === undefined ? {} : { xValues }),
    ...(sizes === undefined ? {} : { sizes }),
  });
}

function validateGroupCompatibility(groups: readonly Readonly<ChartGroup>[]): void {
  if ((groups[0]!.axis ?? 'primary') !== 'primary') {
    throw new RangeError('Chart definition first group must use the primary axis');
  }
  if (groups.length === 1) return;
  if (groups.every(({ type }) => type === 'scatter')) return;
  if (groups.some(({ type }) => !COMBINABLE_CATEGORICAL_TYPES.has(type))) {
    throw new TypeError('Chart group types cannot be combined safely');
  }
}

function normalizeChartType(value: unknown, context: string): ChartType {
  if (typeof value !== 'string' || !CHART_TYPE_SET.has(value)) {
    throw new TypeError(`${context} is unsupported`);
  }
  return value as ChartType;
}

function normalizeAxis(value: unknown, context: string): 'primary' | 'secondary' {
  if (value !== 'primary' && value !== 'secondary') {
    throw new TypeError(`${context} must be primary or secondary`);
  }
  return value;
}

function normalizeCategories(value: unknown, context: string): ChartCategories {
  const categories = readArray(value, context);
  if (categories.length === 0) throw new RangeError(`${context} must not be empty`);
  const nested = categories.every(Array.isArray);
  if (nested) {
    const levels = categories.map((level, levelIndex) => {
      const values = readArray(level, `${context} level ${levelIndex}`);
      if (values.length === 0) throw new RangeError(`${context} levels must not be empty`);
      return Object.freeze(values.map((entry, pointIndex) =>
        normalizeXmlString(entry, `${context} level ${levelIndex} item ${pointIndex}`, false)));
    });
    const length = levels[0]!.length;
    if (levels.some((level) => level.length !== length)) {
      throw new RangeError(`${context} levels must have equal lengths`);
    }
    return Object.freeze(levels);
  }
  if (categories.some(Array.isArray)) {
    throw new TypeError(`${context} must be a flat category array or an array of string levels`);
  }
  return Object.freeze(categories.map((entry, index) => {
    if (typeof entry === 'number') return normalizeFiniteNumber(entry, `${context} item ${index}`);
    return normalizeXmlString(entry, `${context} item ${index}`, false);
  }));
}

function categoryPointCount(categories: ChartCategories): number {
  const first = categories[0];
  return Array.isArray(first) ? first.length : categories.length;
}

function normalizeNumberArray(value: unknown, context: string): readonly number[] {
  return Object.freeze(readArray(value, context).map((entry, index) =>
    normalizeFiniteNumber(entry, `${context} item ${index}`)));
}

function normalizeFiniteNumber(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  return Object.is(value, -0) ? 0 : value;
}

function normalizeXmlString(value: unknown, context: string, nonEmpty: boolean): string {
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
  if (nonEmpty && value.length === 0) throw new RangeError(`${context} must not be empty`);
  if (containsInvalidXmlCharacter(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value;
}

function readObject(
  value: unknown,
  allowed: ReadonlySet<string>,
  required: ReadonlySet<string>,
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string') {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    if (!allowed.has(key)) throw new TypeError(`${context} contains unsupported property ${key}`);
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  for (const key of required) {
    if (!Object.hasOwn(result, key)) throw new TypeError(`${context} is missing property ${key}`);
  }
  return result;
}

function readArray(value: unknown, context: string): readonly unknown[] {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} must be an ordinary array`);
  }
  const keys = Reflect.ownKeys(value);
  if (keys.length !== value.length + 1 || !keys.includes('length')) {
    throw new TypeError(`${context} must be a dense array without extra properties`);
  }
  const result: unknown[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const key = String(index);
    if (!keys.includes(key)) throw new TypeError(`${context} must be a dense array`);
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} item ${index} must be a data property`);
    }
    result.push(descriptor.value);
  }
  return result;
}

function containsInvalidXmlCharacter(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 0x9 || code === 0xa || code === 0xd) continue;
    if (code < 0x20 || code === 0xfffe || code === 0xffff) return true;
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next < 0xdc00 || next > 0xdfff) return true;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return true;
    }
  }
  return false;
}
