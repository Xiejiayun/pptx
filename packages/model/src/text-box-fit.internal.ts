import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import { requireTextBodyProperties } from './text-body-properties.internal.js';
import type { TextBoxFit } from './text.js';

const FROM_OOXML = new Map<string, TextBoxFit>([
  ['noAutofit', 'none'],
  ['normAutofit', 'shrink'],
  ['spAutoFit', 'resize'],
]);

const TO_OOXML: Readonly<Record<Exclude<TextBoxFit, 'none'>, string>> = {
  shrink: 'normAutofit',
  resize: 'spAutoFit',
};

const POST_FIT_CHILDREN = new Set(['scene3d', 'sp3d', 'extLst']);

export function normalizeTextBoxFit(value: unknown, context: string): TextBoxFit {
  if (value !== 'none' && value !== 'shrink' && value !== 'resize') {
    throw new TypeError(`${context} must be none, shrink, or resize`);
  }
  return value;
}

export function renderTextBoxFitChild(value: TextBoxFit): string {
  return value === 'none' ? '' : `<a:${TO_OOXML[value]}/>`;
}

export function readTextBoxFit(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  partUri: string,
): TextBoxFit | undefined {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const choices = fitChildren(bodyProperties);
  return choices.length === 1 ? FROM_OOXML.get(choices[0]!.localName) : undefined;
}

export function replaceTextBoxFit(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  value: TextBoxFit | undefined,
  partUri: string,
): void {
  const bodyProperties = requireTextBodyProperties(xml, shape, partUri);
  const focused = LosslessXmlDocument.parse(xml.original(bodyProperties));
  const root = focused.roots[0];
  if (!root || root.localName !== 'bodyPr') {
    throw new ModelParseError('Invalid text body properties template');
  }
  const choices = fitChildren(root);
  const targetName = value === undefined || value === 'none' ? undefined : TO_OOXML[value];
  if (targetName === undefined && choices.length === 0) return;
  if (targetName !== undefined && choices.length === 1 && choices[0]!.localName === targetName) return;

  if (targetName === undefined) {
    for (const choice of choices) focused.removeElement(choice);
  } else {
    const prefix = elementPrefix(root);
    const child = `<${prefix}${targetName}/>`;
    if (choices[0]) {
      focused.replaceElement(choices[0], child);
      for (const extra of choices.slice(1)) focused.removeElement(extra);
    } else if (root.selfClosing) {
      focused.appendChildXml(root, child);
    } else {
      const anchor = directChildren(root).find(
        (candidate) => elementPrefix(candidate) === prefix && POST_FIT_CHILDREN.has(candidate.localName),
      );
      if (anchor) focused.replace(anchor.start, anchor.start, child);
      else focused.appendChildXml(root, child);
    }
  }
  xml.replaceElement(bodyProperties, focused.serialize());
}

function fitChildren(element: XmlElement): XmlElement[] {
  const prefix = elementPrefix(element);
  return directChildren(element).filter(
    (child) => elementPrefix(child) === prefix && FROM_OOXML.has(child.localName),
  );
}

function elementPrefix(element: XmlElement): string {
  const separator = element.name.indexOf(':');
  return separator < 0 ? '' : element.name.slice(0, separator + 1);
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}
