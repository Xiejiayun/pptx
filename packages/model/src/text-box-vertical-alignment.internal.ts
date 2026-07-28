import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  requireTextBodyProperties,
  updateTextBodyAttribute,
} from './text-body-properties.internal.js';
import type { TextBoxVerticalAlignment } from './text.js';

const TO_OOXML: Readonly<Record<TextBoxVerticalAlignment, string>> = {
  top: 't',
  middle: 'ctr',
  bottom: 'b',
};

const FROM_OOXML = new Map<string, TextBoxVerticalAlignment>([
  ['t', 'top'],
  ['ctr', 'middle'],
  ['b', 'bottom'],
]);

export function normalizeTextBoxVerticalAlignment(
  value: unknown,
  context: string,
): TextBoxVerticalAlignment {
  if (value !== 'top' && value !== 'middle' && value !== 'bottom') {
    throw new TypeError(`${context} must be top, middle, or bottom`);
  }
  return value;
}

export function renderTextBoxVerticalAlignmentAttribute(
  value: TextBoxVerticalAlignment,
): string {
  return ` anchor="${TO_OOXML[value]}"`;
}

export function readTextBoxVerticalAlignment(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): TextBoxVerticalAlignment | undefined {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const token = xml.attribute(bodyProperties, 'anchor')?.value;
  return token === undefined ? undefined : FROM_OOXML.get(token);
}

export function replaceTextBoxVerticalAlignment(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: TextBoxVerticalAlignment | undefined,
  partUri: string,
): void {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const updated = updateTextBodyAttribute(
    xml.original(bodyProperties),
    'anchor',
    value === undefined ? undefined : TO_OOXML[value],
  );
  xml.replaceElement(bodyProperties, updated);
}
