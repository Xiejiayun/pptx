import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  requireTextBodyProperties,
  updateTextBodyAttribute,
} from './text-body-properties.internal.js';
import type { TextBoxTextDirection } from './text.js';

const DIRECTIONS = new Set<string>([
  'eaVert',
  'horz',
  'mongolianVert',
  'vert',
  'vert270',
  'wordArtVert',
  'wordArtVertRtl',
]);

export function normalizeTextBoxTextDirection(
  value: unknown,
  context: string,
): TextBoxTextDirection {
  if (typeof value !== 'string' || !DIRECTIONS.has(value)) {
    throw new TypeError(
      `${context} must be eaVert, horz, mongolianVert, vert, vert270, wordArtVert, or wordArtVertRtl`,
    );
  }
  return value as TextBoxTextDirection;
}

export function renderTextBoxTextDirectionAttribute(value: TextBoxTextDirection): string {
  return ` vert="${value}"`;
}

export function readTextBoxTextDirection(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): TextBoxTextDirection | undefined {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const token = xml.attribute(bodyProperties, 'vert')?.value;
  return token !== undefined && DIRECTIONS.has(token)
    ? token as TextBoxTextDirection
    : undefined;
}

export function replaceTextBoxTextDirection(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: TextBoxTextDirection | undefined,
  partUri: string,
): void {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const updated = updateTextBodyAttribute(xml.original(bodyProperties), 'vert', value);
  xml.replaceElement(bodyProperties, updated);
}
