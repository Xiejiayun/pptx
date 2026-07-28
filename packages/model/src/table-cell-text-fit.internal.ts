import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTextBoxFit,
  replaceTextBoxFit,
} from './text-box-fit.internal.js';
import type { TextBoxFit } from './text.js';

export function readTableCellTextFit(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  partUri: string,
): TextBoxFit | undefined {
  if (!hasUniqueTextBodyProperties(cell)) return undefined;
  return readTextBoxFit(xml, cell, partUri);
}

export function replaceTableCellTextFit(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  value: TextBoxFit | undefined,
  partUri: string,
): boolean {
  if (!hasUniqueTextBodyProperties(cell)) {
    throw new ModelParseError(
      'Table cell must contain one direct text body with one body properties element',
      partUri,
    );
  }
  replaceTextBoxFit(xml, cell, value, partUri);
  return xml.changed;
}

function hasUniqueTextBodyProperties(cell: XmlElement): boolean {
  const textBodies = directChildren(cell, 'txBody');
  return textBodies.length === 1 && directChildren(textBodies[0]!, 'bodyPr').length === 1;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
