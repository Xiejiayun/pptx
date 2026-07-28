import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { GradientCodec, type GradientFill } from '@pptx/codecs';
import type { Relationship } from '@pptx/opc';
import { ModelParseError } from './errors.js';
import type { PresentationModel } from './presentation.js';
import {
  normalizeParagraphBullet,
  normalizeParagraphLevel,
  normalizeParagraphSpacing,
  normalizeRichText,
  normalizeTextAlignment,
  readRichText,
  renderParagraphProperties,
  renderRichTextParagraphs,
  resolveParagraphSpacing,
  replaceRichText,
  type NormalizedParagraphBullet,
  type NormalizedParagraphSpacing,
  type NormalizedParagraphSpacingUpdate,
} from './rich-text.internal.js';
import { decodeShape, ShapeModel, type SemanticShape } from './shapes.js';
import type { ParagraphBullet, ParagraphSpacing, RichTextParagraph, TextAlignment } from './text.js';
import { inches, type Transform } from './units.js';

export interface AddTextOptions extends Partial<Transform> {
  readonly name?: string;
  readonly align?: TextAlignment;
  readonly bullet?: ParagraphBullet;
  readonly level?: number;
  readonly spacing?: ParagraphSpacing;
}

export class SlideTitleModel {
  constructor(private readonly slide: SlideModel) {}

  get text(): string {
    const { xml } = this.slide.parse();
    const shape = findTitleShape(xml);
    if (!shape) return '';
    return readPlainText(xml, shape);
  }

  set text(value: string) {
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml } = this.slide.parse();
      const shape = findTitleShape(xml);
      if (!shape) throw new ModelParseError('Slide does not contain a title shape', this.slide.partUri);
      const properties = xml.descendants(shape, 'cNvPr')[0];
      const id = Number.parseInt(properties ? xml.attribute(properties, 'id')?.value ?? '' : '', 10);
      if (Number.isFinite(id)) this.slide.setShapeText(id, value);
      else replacePlainText(xml, shape, value, this.slide.partUri, (updated) => this.slide.setXml(updated));
    });
  }
}

export class SlideModel {
  readonly title = new SlideTitleModel(this);
  readonly #shapeModels = new Map<number, SemanticShape>();
  #relationshipId: string;
  #slideId: number;

  constructor(
    readonly presentation: PresentationModel,
    readonly partUri: string,
    relationshipId: string,
    slideId: number,
  ) {
    this.#relationshipId = relationshipId;
    this.#slideId = slideId;
  }

  get relationshipId(): string {
    return this.#relationshipId;
  }

  get slideId(): number {
    return this.#slideId;
  }

  /** @internal */
  syncIdentity(relationshipId: string, slideId: number): void {
    this.#relationshipId = relationshipId;
    this.#slideId = slideId;
  }

  get background(): GradientFill | undefined {
    return new GradientCodec().getSlideBackground(this.presentation.opcPackage, this.partUri);
  }

  set background(value: GradientFill) {
    new GradientCodec().setSlideBackground(this.presentation.opcPackage, this.partUri, value);
  }

  get relationships(): readonly Relationship[] {
    return this.presentation.opcPackage.relationships(this.partUri);
  }

  get shapes(): readonly SemanticShape[] {
    const { xml } = this.parse();
    const candidates = xml
      .elements()
      .filter(({ localName }) => ['sp', 'pic', 'graphicFrame', 'grpSp'].includes(localName));
    const shapes: SemanticShape[] = [];
    for (const element of candidates) {
      const decoded = decodeShape(this, xml, element);
      if (!decoded) continue;
      const existing = this.#shapeModels.get(decoded.id);
      if (existing && existing.kind === decoded.kind && existing.constructor === decoded.constructor) {
        shapes.push(existing);
      } else {
        this.#shapeModels.set(decoded.id, decoded);
        shapes.push(decoded);
      }
    }
    return shapes;
  }

  get opaqueExtensionCount(): number {
    const { xml } = this.parse();
    return xml.elements('extLst').length + xml.elements('AlternateContent').length;
  }

  parse(): { xml: LosslessXmlDocument } {
    const part = this.presentation.opcPackage.requirePart(this.partUri);
    try {
      return { xml: LosslessXmlDocument.parse(part.bytes) };
    } catch (error) {
      throw new ModelParseError(error instanceof Error ? error.message : String(error), this.partUri);
    }
  }

  resolveShape(id: number): { xml: LosslessXmlDocument; element: XmlElement } {
    const { xml } = this.parse();
    const element = xml
      .elements()
      .filter(({ localName }) => ['sp', 'pic', 'graphicFrame', 'grpSp'].includes(localName))
      .find((candidate) => {
        const properties = xml.descendants(candidate, 'cNvPr')[0];
        return Number.parseInt(properties ? xml.attribute(properties, 'id')?.value ?? '' : '', 10) === id;
      });
    if (!element) throw new ModelParseError(`Shape ${id} was not found`, this.partUri);
    return { xml, element };
  }

  setShapeText(id: number, value: string): void {
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      replacePlainText(xml, element, value, this.partUri, (updated) => this.setXml(updated));
    });
  }

  getShapeText(id: number): string {
    const { xml, element } = this.resolveShape(id);
    return readPlainText(xml, element);
  }

  setShapeRichText(id: number, value: readonly RichTextParagraph[]): void {
    this.presentation.opcPackage.transaction(() => {
      const paragraphs = normalizeRichText(value);
      const { xml, element } = this.resolveShape(id);
      replaceRichText(xml, element, paragraphs, this.partUri, (updated) => this.setXml(updated));
    });
  }

  getShapeRichText(id: number): readonly RichTextParagraph[] {
    const { xml, element } = this.resolveShape(id);
    return readRichText(xml, element);
  }

  setShapeTransform(id: number, changes: Partial<Transform>): void {
    const { xml, element } = this.resolveShape(id);
    const xfrm = xml.descendants(element, 'xfrm')[0];
    const off = xfrm ? xml.descendants(xfrm, 'off')[0] : undefined;
    const ext = xfrm ? xml.descendants(xfrm, 'ext')[0] : undefined;
    if (!xfrm || !off || !ext) throw new ModelParseError(`Shape ${id} has no editable transform`, this.partUri);
    if (changes.x !== undefined) setAttribute(xml, off, 'x', changes.x);
    if (changes.y !== undefined) setAttribute(xml, off, 'y', changes.y);
    if (changes.width !== undefined) setAttribute(xml, ext, 'cx', changes.width);
    if (changes.height !== undefined) setAttribute(xml, ext, 'cy', changes.height);
    if (changes.rotation !== undefined) setAttribute(xml, xfrm, 'rot', changes.rotation);
    if (changes.flipHorizontal !== undefined) setAttribute(xml, xfrm, 'flipH', changes.flipHorizontal ? 1 : 0);
    if (changes.flipVertical !== undefined) setAttribute(xml, xfrm, 'flipV', changes.flipVertical ? 1 : 0);
    this.setXml(xml.serialize());
  }

  addText(value: string, options: AddTextOptions = {}): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const normalized = validateTextInput(value, options);
      const bullet = normalized.bullet === false ? undefined : normalized.bullet;
      const spacing = resolveParagraphSpacing(normalized.spacing);
      const paragraphs = normalized.value
        .split('\n')
        .map((line) => textParagraphXml(line, 'a:', options.align, bullet, spacing, normalized.level))
        .join('');
      return this.addTextShape(paragraphs, options);
    });
  }

  addRichText(value: readonly RichTextParagraph[], options: AddTextOptions = {}): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const paragraphs = normalizeRichText(value);
      const defaults = validateAddTextOptions(options);
      return this.addTextShape(
        renderRichTextParagraphs(paragraphs, {
          ...(options.align ? { defaultAlign: options.align } : {}),
          ...(defaults.bullet !== undefined ? { defaultBullet: defaults.bullet } : {}),
          ...(defaults.level !== undefined ? { defaultLevel: defaults.level } : {}),
          ...(defaults.spacing !== undefined ? { defaultSpacing: defaults.spacing } : {}),
        }),
        options,
      );
    });
  }

  setXml(xml: string): void {
    this.presentation.setXmlPart(this.partUri, xml);
  }

  private addTextShape(paragraphs: string, options: AddTextOptions): ShapeModel {
    const { xml } = this.parse();
    const shapeTree = xml
      .elements('spTree')
      .find(({ parent }) => parent?.localName === 'cSld');
    if (!shapeTree) throw new ModelParseError('Slide does not contain a shape tree', this.partUri);
    const nextId = allocateShapeId(xml);
    const shapeXml = textShapeXml(nextId, paragraphs, options);
    const extensionList = shapeTree.children.find(
      (child): child is XmlElement => child.type === 'element' && child.localName === 'extLst',
    );
    if (extensionList) xml.replace(extensionList.start, extensionList.start, shapeXml);
    else xml.appendChildXml(shapeTree, shapeXml);
    this.setXml(xml.serialize());
    const shape = this.shapes.find((candidate) => candidate.id === nextId);
    if (!(shape instanceof ShapeModel) || shape.kind !== 'text') {
      throw new ModelParseError(`Created text shape ${nextId} could not be resolved`, this.partUri);
    }
    return shape;
  }
}

export function findTitleShape(xml: LosslessXmlDocument): XmlElement | undefined {
  const shapes = xml.elements('sp');
  return (
    shapes.find((shape) =>
      xml.descendants(shape, 'ph').some((placeholder) => {
        const type = xml.attribute(placeholder, 'type')?.value;
        return type === 'title' || type === 'ctrTitle';
      }),
    ) ?? shapes.find((shape) => xml.descendants(shape, 't').length > 0)
  );
}

function replacePlainText(
  xml: LosslessXmlDocument,
  element: XmlElement,
  value: string,
  partUri: string,
  save: (xml: string) => void,
): void {
  const normalized = validatePlainText(value);
  const textBody = directChildren(element, 'txBody')[0];
  if (!textBody) throw new ModelParseError('Shape does not contain a text body', partUri);
  const paragraphs = directChildren(textBody, 'p');
  const template = paragraphs[0];
  if (!template) throw new ModelParseError('Shape does not contain a text paragraph', partUri);
  const replacement = normalized
    .split('\n')
    .map((line) => renderParagraphTemplate(xml, template, line))
    .join('');
  xml.replaceElement(template, replacement);
  for (const extra of paragraphs.slice(1)) xml.removeElement(extra);
  save(xml.serialize());
}

function setAttribute(xml: LosslessXmlDocument, element: XmlElement, name: string, value: number): void {
  const attribute = xml.attribute(element, name);
  if (attribute) {
    xml.replaceAttribute(attribute, String(value));
    return;
  }
  const insertionPoint = element.startTagEnd - (element.selfClosing ? 2 : 1);
  xml.replace(insertionPoint, insertionPoint, ` ${name}="${value}"`);
}

interface NormalizedTextInput {
  readonly value: string;
  readonly bullet: NormalizedParagraphBullet | false | undefined;
  readonly level: number | undefined;
  readonly spacing: NormalizedParagraphSpacingUpdate | undefined;
}

function validateTextInput(value: string, options: AddTextOptions): NormalizedTextInput {
  const normalized = validatePlainText(value);
  const defaults = validateAddTextOptions(options);
  return {
    value: normalized,
    bullet: defaults.bullet,
    level: defaults.level,
    spacing: defaults.spacing,
  };
}

interface NormalizedAddTextOptions {
  readonly bullet?: NormalizedParagraphBullet | false;
  readonly level?: number;
  readonly spacing?: NormalizedParagraphSpacingUpdate;
}

function validateAddTextOptions(options: AddTextOptions): NormalizedAddTextOptions {
  if (options.name !== undefined && typeof options.name !== 'string') {
    throw new TypeError('Text shape name must be a string');
  }
  if (options.align !== undefined) normalizeTextAlignment(options.align, 'Text alignment');
  if (options.name !== undefined && containsInvalidXmlCharacter(options.name)) {
    throw new TypeError('Text shape name contains invalid XML characters');
  }
  for (const [name, candidate] of [
    ['x', options.x],
    ['y', options.y],
    ['width', options.width],
    ['height', options.height],
    ['rotation', options.rotation],
  ] as const) {
    if (candidate !== undefined && !Number.isFinite(candidate)) {
      throw new TypeError(`Text shape ${name} must be finite`);
    }
  }
  for (const [name, candidate] of [
    ['flipHorizontal', options.flipHorizontal],
    ['flipVertical', options.flipVertical],
  ] as const) {
    if (candidate !== undefined && typeof candidate !== 'boolean') {
      throw new TypeError(`Text shape ${name} must be a boolean`);
    }
  }
  if (options.width !== undefined && Math.round(options.width) <= 0) {
    throw new RangeError('Text shape width must be greater than zero');
  }
  if (options.height !== undefined && Math.round(options.height) <= 0) {
    throw new RangeError('Text shape height must be greater than zero');
  }
  const bullet = options.bullet === undefined
    ? undefined
    : normalizeParagraphBullet(options.bullet, 'Text bullet');
  const level = options.level === undefined
    ? undefined
    : normalizeParagraphLevel(options.level, 'Text level');
  const spacing = options.spacing === undefined
    ? undefined
    : normalizeParagraphSpacing(options.spacing, 'Text spacing');
  return {
    ...(bullet !== undefined ? { bullet } : {}),
    ...(level !== undefined ? { level } : {}),
    ...(spacing !== undefined ? { spacing } : {}),
  };
}

function validatePlainText(value: string): string {
  if (typeof value !== 'string') throw new TypeError('Text shape value must be a string');
  if (containsInvalidXmlCharacter(value)) throw new TypeError('Text shape value contains invalid XML characters');
  return value.replace(/\r\n?/g, '\n');
}

function containsInvalidXmlCharacter(value: string): boolean {
  return /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value);
}

function allocateShapeId(xml: LosslessXmlDocument): number {
  return xml.elements('cNvPr').reduce((maximum, element) => {
    const value = Number.parseInt(xml.attribute(element, 'id')?.value ?? '', 10);
    return Number.isFinite(value) ? Math.max(maximum, value) : maximum;
  }, 1) + 1;
}

function textShapeXml(id: number, paragraphs: string, options: AddTextOptions): string {
  const x = Math.round(options.x ?? 0);
  const y = Math.round(options.y ?? 0);
  const width = Math.round(options.width ?? inches(1));
  const height = Math.round(options.height ?? inches(1));
  const rotation = Math.round(options.rotation ?? 0);
  const transformAttributes = [
    rotation === 0 ? '' : ` rot="${rotation}"`,
    options.flipHorizontal ? ' flipH="1"' : '',
    options.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const name = escapeXmlAttribute(options.name ?? `Text ${id}`);
  return `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${transformAttributes}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

function textParagraphXml(
  value: string,
  prefix = 'a:',
  align?: TextAlignment,
  bullet?: NormalizedParagraphBullet,
  spacing?: NormalizedParagraphSpacing,
  level?: number,
): string {
  const properties = renderParagraphProperties(undefined, prefix, align, bullet, spacing, level);
  const endProperties = `<${prefix}endParaRPr lang="en-US" dirty="0"/>`;
  if (value.length === 0) return `<${prefix}p>${properties}${endProperties}</${prefix}p>`;
  return `<${prefix}p>${properties}${defaultTextRunXml(value, prefix)}${endProperties}</${prefix}p>`;
}

function defaultTextRunXml(value: string, prefix = 'a:'): string {
  return `<${prefix}r><${prefix}rPr lang="en-US" dirty="0"><${prefix}solidFill><${prefix}schemeClr val="tx1"/></${prefix}solidFill><${prefix}latin typeface="+mn-lt"/></${prefix}rPr><${prefix}t xml:space="preserve">${escapeXmlText(value)}</${prefix}t></${prefix}r>`;
}

function readPlainText(xml: LosslessXmlDocument, element: XmlElement): string {
  const textBody = directChildren(element, 'txBody')[0];
  if (!textBody) return '';
  return directChildren(textBody, 'p').map((paragraph) => readParagraphText(xml, paragraph)).join('\n');
}

function readParagraphText(xml: LosslessXmlDocument, paragraph: XmlElement): string {
  let value = '';
  const visit = (element: XmlElement): void => {
    for (const child of element.children) {
      if (child.type !== 'element') continue;
      if (child.localName === 't') value += xml.text(child);
      else if (child.localName === 'br') value += '\n';
      else visit(child);
    }
  };
  visit(paragraph);
  return value;
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function renderParagraphTemplate(
  source: LosslessXmlDocument,
  template: XmlElement,
  value: string,
): string {
  const prefix = qualifiedPrefix(template.name);
  if (template.selfClosing) return textParagraphXml(value, prefix);
  const paragraph = LosslessXmlDocument.parse(source.original(template));
  const root = paragraph.roots[0];
  if (!root) return textParagraphXml(value, prefix);
  for (const lineBreak of paragraph.descendants(root, 'br')) paragraph.removeElement(lineBreak);
  const textNodes = paragraph.descendants(root, 't');
  const first = textNodes[0];
  if (!first) {
    if (value.length === 0) return paragraph.serialize();
    const endProperties = directChildren(root, 'endParaRPr')[0];
    const insertionPoint = endProperties?.start ?? root.endTagStart;
    paragraph.replace(insertionPoint, insertionPoint, defaultTextRunXml(value, qualifiedPrefix(root.name)));
    return paragraph.serialize();
  }
  replaceTextElement(paragraph, first, value);
  for (const extra of textNodes.slice(1)) replaceTextElement(paragraph, extra, '');
  return paragraph.serialize();
}

function replaceTextElement(xml: LosslessXmlDocument, element: XmlElement, value: string): void {
  const space = xml.attribute(element, 'xml:space');
  if (element.selfClosing) {
    const original = xml.original(element);
    const marker = original.lastIndexOf('/>');
    if (marker < 0) throw new ModelParseError(`Invalid self-closing text element ${element.name}`);
    let opening = original.slice(0, marker).replace(/\s+$/, '');
    if (space) {
      const relativeStart = space.valueStart - element.start;
      const relativeEnd = space.valueEnd - element.start;
      opening = `${opening.slice(0, relativeStart)}preserve${opening.slice(relativeEnd)}`;
    } else {
      opening += ' xml:space="preserve"';
    }
    xml.replaceElement(element, `${opening}>${escapeXmlText(value)}</${element.name}>`);
    return;
  }
  if (space) xml.replaceAttribute(space, 'preserve');
  else {
    const insertionPoint = element.startTagEnd - 1;
    xml.replace(insertionPoint, insertionPoint, ' xml:space="preserve"');
  }
  xml.replaceText(element, value);
}

function qualifiedPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : `${name.slice(0, separator)}:`;
}
