import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  requireTextBodyProperties,
  updateTextBodyAttribute,
} from './text-body-properties.internal.js';

const TO_OOXML = new Map<boolean, string>([
  [true, 'square'],
  [false, 'none'],
]);

const FROM_OOXML = new Map<string, boolean>([
  ['square', true],
  ['none', false],
]);

export function normalizeTextBoxWrap(value: unknown, context: string): boolean {
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}

export function renderTextBoxWrapAttribute(value: boolean): string {
  return ` wrap="${TO_OOXML.get(value)!}"`;
}

export function readTextBoxWrap(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): boolean | undefined {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const token = xml.attribute(bodyProperties, 'wrap')?.value;
  return token === undefined ? undefined : FROM_OOXML.get(token);
}

export function replaceTextBoxWrap(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: boolean | undefined,
  partUri: string,
): void {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const updated = updateTextBodyAttribute(
    xml.original(bodyProperties),
    'wrap',
    value === undefined ? undefined : TO_OOXML.get(value)!,
  );
  xml.replaceElement(bodyProperties, updated);
}
