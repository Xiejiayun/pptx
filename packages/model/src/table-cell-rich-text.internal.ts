import type { LosslessXmlDocument, XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readRichTextState,
  type ReadRichTextState,
} from './rich-text.internal.js';
import type { ShapeHyperlinkReadContext } from './shape-hyperlink.internal.js';
import type { RichTextParagraph } from './text.js';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

export function readTableCellRichText(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): readonly RichTextParagraph[] {
  return readOwnedTableCellRichTextState(xml, cell, context)?.paragraphs ?? [];
}

export function readTableCellText(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): string {
  return readTableCellRichText(xml, cell, context)
    .map(({ runs }) => runs.map((run) =>
      `${run.softBreakBefore ? '\n' : ''}${run.text}`).join(''))
    .join('\n');
}

export function requireEditableTableCellRichTextState(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
  partUri: string,
): ReadRichTextState {
  const state = readOwnedTableCellRichTextState(xml, cell, context);
  const body = ownedTextBody(cell);
  if (!state || !body || ownedParagraphs(body)?.length === 0) {
    throw new ModelParseError(
      'Table cell rich text state is not safely editable',
      partUri,
    );
  }
  return state;
}

export function requireEditablePlainTableCellText(
  _xml: LosslessXmlDocument,
  cell: XmlElement,
  partUri: string,
): XmlElement {
  const body = ownedTextBody(cell);
  const paragraphs = body ? ownedParagraphs(body) : undefined;
  const paragraph = paragraphs?.length === 1 ? paragraphs[0] : undefined;
  const target = paragraph ? plainParagraphText(paragraph) : undefined;
  if (!target) {
    throw new ModelParseError(
      'Table cell text state is not safely editable',
      partUri,
    );
  }
  return target;
}

function readOwnedTableCellRichTextState(
  xml: LosslessXmlDocument,
  cell: XmlElement,
  context: ShapeHyperlinkReadContext,
): ReadRichTextState | undefined {
  const body = ownedTextBody(cell);
  if (!body || !ownedParagraphs(body)) return undefined;
  return readRichTextState(xml, cell, context);
}

function ownedTextBody(cell: XmlElement): XmlElement | undefined {
  if (cell.localName !== 'tc' || namespaceUri(cell) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const candidates = directChildren(cell).filter(({ localName }) =>
    localName === 'txBody');
  return candidates.length === 1
    && namespaceUri(candidates[0]!) === DRAWING_NAMESPACE
    ? candidates[0]
    : undefined;
}

function ownedParagraphs(body: XmlElement): readonly XmlElement[] | undefined {
  const candidates = directChildren(body).filter(({ localName }) => localName === 'p');
  return candidates.every((candidate) => namespaceUri(candidate) === DRAWING_NAMESPACE)
    ? candidates
    : undefined;
}

function plainParagraphText(paragraph: XmlElement): XmlElement | undefined {
  const children = directChildren(paragraph);
  if (children.some((child) => namespaceUri(child) !== DRAWING_NAMESPACE)) {
    return undefined;
  }
  let index = children[0]?.localName === 'pPr' ? 1 : 0;
  const run = children[index];
  if (run?.localName !== 'r') return undefined;
  index += 1;
  if (children[index]?.localName === 'endParaRPr') index += 1;
  if (index !== children.length) return undefined;

  const runChildren = directChildren(run);
  if (runChildren.some((child) => namespaceUri(child) !== DRAWING_NAMESPACE)) {
    return undefined;
  }
  let runIndex = runChildren[0]?.localName === 'rPr' ? 1 : 0;
  const text = runChildren[runIndex];
  if (text?.localName !== 't') return undefined;
  runIndex += 1;
  return runIndex === runChildren.length ? text : undefined;
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
