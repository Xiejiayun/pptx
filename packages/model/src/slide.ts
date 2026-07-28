import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { GradientCodec, type GradientFill } from '@pptx/codecs';
import type { Relationship } from '@pptx/opc';
import type { PresentationModel } from './presentation.js';
import { decodeShape, ShapeModel, type SemanticShape } from './shapes.js';
import { inches, type Transform } from './units.js';

export interface AddTextOptions extends Partial<Transform> {
  readonly name?: string;
}

export class ModelParseError extends Error {
  constructor(message: string, readonly partUri?: string) {
    super(partUri ? `${message}: ${partUri}` : message);
    this.name = 'ModelParseError';
  }
}

export class SlideTitleModel {
  constructor(private readonly slide: SlideModel) {}

  get text(): string {
    const { xml } = this.slide.parse();
    const shape = findTitleShape(xml);
    if (!shape) return '';
    return xml.descendants(shape, 't').map((node) => xml.text(node)).join('');
  }

  set text(value: string) {
    const { xml } = this.slide.parse();
    const shape = findTitleShape(xml);
    if (!shape) throw new ModelParseError('Slide does not contain a title shape', this.slide.partUri);
    const properties = xml.descendants(shape, 'cNvPr')[0];
    const id = Number.parseInt(properties ? xml.attribute(properties, 'id')?.value ?? '' : '', 10);
    if (Number.isFinite(id)) this.slide.setShapeText(id, value);
    else replaceTextRuns(xml, shape, value, this.slide.partUri, (updated) => this.slide.setXml(updated));
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
    const { xml, element } = this.resolveShape(id);
    replaceTextRuns(xml, element, value, this.partUri, (updated) => this.setXml(updated));
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
      validateTextInput(value, options);
      const { xml } = this.parse();
      const shapeTree = xml
        .elements('spTree')
        .find(({ parent }) => parent?.localName === 'cSld');
      if (!shapeTree) throw new ModelParseError('Slide does not contain a shape tree', this.partUri);
      const nextId = allocateShapeId(xml);
      const shapeXml = textShapeXml(nextId, value, options);
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
    });
  }

  setXml(xml: string): void {
    this.presentation.setXmlPart(this.partUri, xml);
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

function replaceTextRuns(
  xml: LosslessXmlDocument,
  element: XmlElement,
  value: string,
  partUri: string,
  save: (xml: string) => void,
): void {
  validatePlainText(value);
  const runs = xml.descendants(element, 't');
  const first = runs[0];
  if (!first) throw new ModelParseError('Shape does not contain text', partUri);
  xml.replaceText(first, value);
  for (const extra of runs.slice(1)) xml.replaceText(extra, '');
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

function validateTextInput(value: string, options: AddTextOptions): void {
  validatePlainText(value);
  if (options.name !== undefined && typeof options.name !== 'string') {
    throw new TypeError('Text shape name must be a string');
  }
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
}

function validatePlainText(value: string): void {
  if (typeof value !== 'string') throw new TypeError('Text shape value must be a string');
  if (/\r|\n/.test(value)) throw new TypeError('Basic text shapes do not support line breaks yet');
  if (containsInvalidXmlCharacter(value)) throw new TypeError('Text shape value contains invalid XML characters');
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

function textShapeXml(id: number, value: string, options: AddTextOptions): string {
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
  const text = escapeXmlText(value);
  return `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${transformAttributes}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr indent="0" marL="0"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" dirty="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:rPr><a:t xml:space="preserve">${text}</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`;
}
