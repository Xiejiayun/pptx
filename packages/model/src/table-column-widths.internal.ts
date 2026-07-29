import {
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

export type NormalizedTableColumnWidthInput =
  | { readonly kind: 'scalar'; readonly value: number }
  | {
      readonly kind: 'array';
      readonly values: readonly number[];
      readonly sum: number;
    };

interface ParsedTableGrid {
  readonly columns: readonly XmlElement[];
  readonly widthAttributes: readonly XmlAttribute[];
  readonly widths: readonly number[];
}

interface ParsedTableTransform {
  readonly widthAttribute: XmlAttribute;
  readonly width: number;
}

export function normalizeTableColumnWidthInput(
  value: unknown,
): NormalizedTableColumnWidthInput {
  if (!Array.isArray(value)) {
    return {
      kind: 'scalar',
      value: normalizePositiveDimension(value, 'Table column widths'),
    };
  }
  const values = readDenseArray(value, 'Table column widths')
    .map((item, index) =>
      normalizePositiveDimension(
        item,
        'Table column widths ' + String(index),
      ));
  return {
    kind: 'array',
    values,
    sum: sumDimensions(values, 'Table column widths'),
  };
}

export function readTableColumnWidths(
  _xml: LosslessXmlDocument,
  frame: XmlElement,
): readonly number[] | undefined {
  const grid = parseTableGrid(frame);
  return grid ? [...grid.widths] : undefined;
}

export function replaceTableColumnWidths(
  xml: LosslessXmlDocument,
  frame: XmlElement,
  input: NormalizedTableColumnWidthInput,
  partUri: string,
): boolean {
  const grid = parseTableGrid(frame);
  if (!grid) {
    throw new ModelParseError(
      'Table must contain one complete direct column grid',
      partUri,
    );
  }
  const target = materializeWidths(input, grid.columns.length);
  const targetSum = input.kind === 'array'
    ? input.sum
    : sumDimensions(target, 'Table column widths');
  const transform = parseTableTransform(frame);
  if (!transform) {
    throw new ModelParseError(
      'Table must contain one editable direct transform width',
      partUri,
    );
  }

  let changed = false;
  for (let index = 0; index < grid.columns.length; index += 1) {
    if (grid.widths[index] !== target[index]) {
      xml.replaceAttribute(
        grid.widthAttributes[index]!,
        String(target[index]),
      );
      changed = true;
    }
  }
  if (transform.width !== targetSum) {
    xml.replaceAttribute(transform.widthAttribute, String(targetSum));
    changed = true;
  }
  return changed;
}

function parseTableGrid(frame: XmlElement): ParsedTableGrid | undefined {
  if (frame.localName !== 'graphicFrame') return undefined;
  const graphic = exactDirectChild(frame, 'graphic');
  const graphicData = graphic
    ? exactDirectChild(graphic, 'graphicData')
    : undefined;
  const table = graphicData
    ? exactDirectChild(graphicData, 'tbl')
    : undefined;
  const grid = table ? exactDirectChild(table, 'tblGrid') : undefined;
  if (!grid) return undefined;

  const columns = directChildren(grid, 'gridCol');
  if (columns.length === 0) return undefined;
  const widthAttributes: XmlAttribute[] = [];
  const widths: number[] = [];
  for (const column of columns) {
    const widthAttribute = exactUnqualifiedAttribute(column, 'w');
    const width = widthAttribute
      ? parseUnsignedSafeInteger(widthAttribute.value, true)
      : undefined;
    if (!widthAttribute || width === undefined) return undefined;
    widthAttributes.push(widthAttribute);
    widths.push(width);
  }
  return { columns, widthAttributes, widths };
}

function parseTableTransform(
  frame: XmlElement,
): ParsedTableTransform | undefined {
  const transform = exactDirectChild(frame, 'xfrm');
  const extent = transform ? exactDirectChild(transform, 'ext') : undefined;
  const widthAttribute = extent
    ? exactUnqualifiedAttribute(extent, 'cx')
    : undefined;
  const width = widthAttribute
    ? parseUnsignedSafeInteger(widthAttribute.value, false)
    : undefined;
  return widthAttribute && width !== undefined
    ? { widthAttribute, width }
    : undefined;
}

function materializeWidths(
  input: NormalizedTableColumnWidthInput,
  count: number,
): readonly number[] {
  if (input.kind === 'scalar') {
    return Array.from({ length: count }, () => input.value);
  }
  if (input.values.length !== count) {
    throw new TypeError(
      'Table column widths must match the table column count',
    );
  }
  return input.values;
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

function normalizePositiveDimension(
  value: unknown,
  context: string,
): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(context + ' must be finite');
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError(
      context + ' must round to a safe integer EMU value',
    );
  }
  if (rounded <= 0) {
    throw new RangeError(context + ' must be greater than zero');
  }
  return rounded;
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

function parseUnsignedSafeInteger(
  value: string,
  positive: boolean,
): number | undefined {
  if (!/^\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) return undefined;
  return positive ? (parsed > 0 ? parsed : undefined) : parsed;
}
