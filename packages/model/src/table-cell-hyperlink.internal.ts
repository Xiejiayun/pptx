import type { LosslessXmlDocument, XmlElement } from '@pptx/lossless-xml';
import {
  readTextRunHyperlink,
  readTextRunHyperlinkBinding,
  type NormalizedHyperlink,
  type ShapeHyperlinkReadContext,
} from './shape-hyperlink.internal.js';
import { ModelParseError } from './errors.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

export function readTableCellHyperlink(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): NormalizedHyperlink | undefined {
  const properties = readTableCellRunProperties(cell);
  return properties ? readTextRunHyperlink(properties, context) : undefined;
}

export interface EditableTableCellHyperlinkState {
  readonly properties: XmlElement;
  readonly hyperlink?: NormalizedHyperlink;
  readonly relationshipId?: string;
}

export function requireEditableTableCellHyperlinkState(
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
  partUri: string,
): EditableTableCellHyperlinkState {
  const properties = readTableCellRunProperties(cell);
  if (!properties) {
    throw new ModelParseError('Table cell hyperlink state is not safely editable', partUri);
  }
  const clickCandidates = directChildren(properties).filter(
    ({ localName }) => localName === 'hlinkClick',
  );
  if (clickCandidates.length === 0) return Object.freeze({ properties });
  const binding = clickCandidates.length === 1
    ? readTextRunHyperlinkBinding(properties, context)
    : undefined;
  if (!binding) {
    throw new ModelParseError('Table cell hyperlink state is not safely editable', partUri);
  }
  return Object.freeze({
    properties,
    hyperlink: binding.hyperlink,
    relationshipId: binding.relationshipId,
  });
}

function readTableCellRunProperties(cell: XmlElement): XmlElement | undefined {
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
  return runChildren[0]!;
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
