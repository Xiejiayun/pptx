import type { LosslessXmlDocument, XmlElement } from '@pptx/lossless-xml';
import {
  readTextRunHyperlink,
  type NormalizedHyperlink,
  type ShapeHyperlinkReadContext,
} from './shape-hyperlink.internal.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

export function readTableCellHyperlink(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): NormalizedHyperlink | undefined {
  if (cell.localName !== 'tc' || namespaceUri(cell) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const body = singleDrawingChild(cell, 'txBody');
  if (!body) return undefined;
  const paragraph = singleDrawingChild(body, 'p');
  if (!paragraph) return undefined;

  const paragraphChildren = directChildren(paragraph);
  if (paragraphChildren.some((child) =>
    namespaceUri(child) !== DRAWING_NAMESPACE
    || !['pPr', 'r', 'endParaRPr'].includes(child.localName))) {
    return undefined;
  }
  let paragraphIndex = paragraphChildren[0]?.localName === 'pPr' ? 1 : 0;
  const run = paragraphChildren[paragraphIndex];
  if (run?.localName !== 'r') return undefined;
  paragraphIndex += 1;
  if (paragraphChildren[paragraphIndex]?.localName === 'endParaRPr') {
    paragraphIndex += 1;
  }
  if (paragraphIndex !== paragraphChildren.length) return undefined;

  const runChildren = directChildren(run);
  if (
    runChildren.length !== 2
    || runChildren[0]!.localName !== 'rPr'
    || runChildren[1]!.localName !== 't'
    || runChildren.some((child) => namespaceUri(child) !== DRAWING_NAMESPACE)
  ) {
    return undefined;
  }
  return readTextRunHyperlink(runChildren[0]!, context);
}

function singleDrawingChild(
  element: XmlElement,
  localName: string,
): XmlElement | undefined {
  const matches = directChildren(element).filter((child) =>
    child.localName === localName);
  if (matches.length !== 1 || namespaceUri(matches[0]!) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  return matches[0];
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function namespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declaration = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const matches = current.attributes.filter(({ name }) => name === declaration);
    if (matches.length > 1) return undefined;
    if (matches[0]) return matches[0].value;
  }
  return undefined;
}
