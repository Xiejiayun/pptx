import {
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

export type NormalizedTableRowHeightInput =
  | { readonly kind: 'scalar'; readonly value: number }
  | { readonly kind: 'array'; readonly values: readonly number[] };

interface ParsedTableRows {
  readonly rows: readonly XmlElement[];
  readonly heightAttributes: readonly XmlAttribute[];
  readonly heights: readonly number[];
}

interface ParsedTableTransform {
  readonly heightAttribute: XmlAttribute;
  readonly height: number;
}

interface MaterializedRowHeights {
  readonly values: readonly number[];
  readonly explicitTotal: number | undefined;
}

export function normalizeTableRowHeightInput(
  value: unknown,
): NormalizedTableRowHeightInput {
  if (!Array.isArray(value)) {
    return {
      kind: 'scalar',
      value: normalizeNonNegativeDimension(value, 'Table row heights'),
    };
  }
  const values = readDenseArray(value, 'Table row heights')
    .map((item, index) =>
      normalizeNonNegativeDimension(
        item,
        'Table row heights ' + String(index),
      ));
  return { kind: 'array', values };
}

export function readTableRowHeights(
  _xml: LosslessXmlDocument,
  frame: XmlElement,
): readonly number[] | undefined {
  const parsed = parseTableRows(frame);
  return parsed ? [...parsed.heights] : undefined;
}

export function replaceTableRowHeights(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  input: NormalizedTableRowHeightInput,
  partUri: string,
): boolean {
  const parsed = parseTableRows(frame);
  if (!parsed) {
    throw new ModelParseError(
      'Table must contain one complete set of direct rows',
      partUri,
    );
  }
  const target = materializeRowHeights(input, parsed.rows.length);
  const transform = parseTableTransform(frame);
  if (!transform) {
    throw new ModelParseError(
      'Table must contain one editable direct transform height',
      partUri,
    );
  }

  let changed = false;
  for (let index = 0; index < parsed.rows.length; index += 1) {
    if (parsed.heights[index] !== target.values[index]) {
      xml.replaceAttribute(
        parsed.heightAttributes[index]!,
        String(target.values[index]),
      );
      changed = true;
    }
  }
  if (
    target.explicitTotal !== undefined &&
    transform.height !== target.explicitTotal
  ) {
    xml.replaceAttribute(
      transform.heightAttribute,
      String(target.explicitTotal),
    );
    changed = true;
  }
  return changed;
}

function parseTableRows(frame: XmlElement): ParsedTableRows | undefined {
  if (frame.localName !== 'graphicFrame') return undefined;
  const graphic = exactDirectChild(frame, 'graphic');
  const graphicData = graphic
    ? exactDirectChild(graphic, 'graphicData')
    : undefined;
  const table = graphicData
    ? exactDirectChild(graphicData, 'tbl')
    : undefined;
  if (!table) return undefined;

  const rows = directChildren(table, 'tr');
  if (rows.length === 0) return undefined;
  const heightAttributes: XmlAttribute[] = [];
  const heights: number[] = [];
  for (const row of rows) {
    const heightAttribute = exactUnqualifiedAttribute(row, 'h');
    const height = heightAttribute
      ? parseNonNegativeSafeInteger(heightAttribute.value)
      : undefined;
    if (!heightAttribute || height === undefined) return undefined;
    heightAttributes.push(heightAttribute);
    heights.push(height);
  }
  return { rows, heightAttributes, heights };
}

function parseTableTransform(
  frame: XmlElement,
): ParsedTableTransform | undefined {
  const transform = exactDirectChild(frame, 'xfrm');
  const extent = transform ? exactDirectChild(transform, 'ext') : undefined;
  const heightAttribute = extent
    ? exactUnqualifiedAttribute(extent, 'cy')
    : undefined;
  const height = heightAttribute
    ? parseNonNegativeSafeInteger(heightAttribute.value)
    : undefined;
  return heightAttribute && height !== undefined
    ? { heightAttribute, height }
    : undefined;
}

function materializeRowHeights(
  input: NormalizedTableRowHeightInput,
  count: number,
): MaterializedRowHeights {
  const values = input.kind === 'scalar'
    ? Array.from({ length: count }, () => input.value)
    : input.values;
  if (values.length !== count) {
    throw new TypeError(
      'Table row heights must match the table row count',
    );
  }
  return {
    values,
    explicitTotal: values.some((value) => value === 0)
      ? undefined
      : sumDimensions(values, 'Table row heights'),
  };
}

function readDenseArray(
  value: readonly unknown[],
  context: string,
): readonly unknown[] {
  if (value.length === 0) {
    throw new TypeError(context + ' must be a non-empty array');
  }
  const allowed = new Set([
    'length',
    ...Array.from({ length: value.length }, (_, index) => String(index)),
  ]);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(
        context + ' contains unsupported property ' + String(key),
      );
    }
  }
  return Array.from({ length: value.length }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(
        context + ' must be dense and contain only data items',
      );
    }
    return descriptor.value;
  });
}

function normalizeNonNegativeDimension(
  value: unknown,
  context: string,
): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(context + ' must be finite');
  }
  if (value < 0) {
    throw new RangeError(context + ' must be non-negative');
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(
      context + ' must round to a safe integer EMU value',
    );
  }
  return rounded === 0 ? 0 : rounded;
}

function sumDimensions(
  dimensions: readonly number[],
  context: string,
): number {
  return dimensions.reduce((sum, dimension) => {
    if (dimension > Number.MAX_SAFE_INTEGER - sum) {
      throw new RangeError(
        context + ' sum must fit a safe integer EMU value',
      );
    }
    return sum + dimension;
  }, 0);
}

function exactDirectChild(
  element: XmlElement,
  localName: string,
): XmlElement | undefined {
  const matches = directChildren(element, localName);
  return matches.length === 1 ? matches[0] : undefined;
}

function directChildren(
  element: XmlElement,
  localName: string,
): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement =>
      child.type === 'element' && child.localName === localName,
  );
}

function exactUnqualifiedAttribute(
  element: XmlElement,
  name: string,
): XmlAttribute | undefined {
  const matches = element.attributes.filter(
    (attribute) => attribute.name === name,
  );
  return matches.length === 1 ? matches[0] : undefined;
}

function parseNonNegativeSafeInteger(value: string): number | undefined {
  if (!/^\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}
