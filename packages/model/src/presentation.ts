import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  joinPartUri,
  OpcPackage,
  PackageError,
  partUriDirname,
  relativeRelationshipTarget,
} from '@pptx/opc';
import {
  detectPresentationFormat,
  presentationFormatProfile,
  UnsupportedPresentationFormatError,
  type PresentationFormat,
  type PresentationFormatProfile,
} from './format.js';
import { SlideModel } from './slide.js';
import {
  cloneSlideDependencies,
  garbageCollectOwnedDependencies,
  ownedSlideDependencyRoots,
} from './dependency.internal.js';
import {
  readPresentationTitle,
  replacePresentationTitle,
} from './presentation-title.internal.js';
import {
  readPresentationAuthor,
  replacePresentationAuthor,
} from './presentation-author.internal.js';
import {
  readPresentationSubject,
  replacePresentationSubject,
} from './presentation-subject.internal.js';
import {
  readPresentationCompany,
  replacePresentationCompany,
} from './presentation-company.internal.js';
import {
  readPresentationCreatedAt,
  replacePresentationCreatedAt,
} from './presentation-created-at.internal.js';
import {
  readPresentationModifiedAt,
  replacePresentationModifiedAt,
} from './presentation-modified-at.internal.js';
import {
  readPresentationLastModifiedBy,
  replacePresentationLastModifiedBy,
} from './presentation-last-modified-by.internal.js';
import {
  readPresentationRevision,
  replacePresentationRevision,
} from './presentation-revision.internal.js';
import type { Emu, SlideSize } from './units.js';

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SLIDE_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const SLIDE_LAYOUT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const PRESENTATION_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const MIN_SLIDE_SIZE = 914_400;
const MAX_SLIDE_SIZE = 51_206_400;

export class PresentationModel {
  readonly presentationPartUri: string;
  readonly format: PresentationFormat;
  readonly formatProfile: PresentationFormatProfile;
  readonly #slideModels = new Map<string, SlideModel>();

  constructor(readonly opcPackage: OpcPackage) {
    const rootRelationship = opcPackage.relationships('/').find(({ type }) => type.endsWith('/officeDocument'));
    this.presentationPartUri = rootRelationship?.resolvedTarget ?? '/ppt/presentation.xml';
    if (!opcPackage.hasPart(this.presentationPartUri)) {
      throw new PackageError('Presentation part was not found', this.presentationPartUri);
    }
    const contentType = opcPackage.requirePart(this.presentationPartUri).contentType;
    const format = detectPresentationFormat(contentType);
    if (!format) throw new UnsupportedPresentationFormatError(contentType, this.presentationPartUri);
    this.format = format;
    this.formatProfile = presentationFormatProfile(format);
  }

  get slides(): readonly SlideModel[] {
    const { xml } = this.parsePresentation();
    const relationships = this.opcPackage.relationships(this.presentationPartUri);
    const ordered = xml
      .elements('sldId')
      .map((element) => ({
        relationshipId: xml.attribute(element, 'r:id')?.value,
        slideId: Number.parseInt(xml.attribute(element, 'id')?.value ?? '', 10),
      }))
      .filter(
        (entry): entry is { relationshipId: string; slideId: number } =>
          Boolean(entry.relationshipId) && Number.isFinite(entry.slideId),
      );
    const slides: SlideModel[] = [];
    for (const entry of ordered) {
      const relationship = relationships.find(({ id }) => id === entry.relationshipId);
      if (relationship?.type.endsWith('/slide') && relationship.resolvedTarget) {
        slides.push(this.slideModel(relationship.resolvedTarget, relationship.id, entry.slideId));
      }
    }
    return slides;
  }

  get slideSize(): SlideSize {
    const { xml } = this.parsePresentation();
    const root = presentationRoot(xml, this.presentationPartUri);
    const size = directChild(root, 'sldSz');
    if (!size) throw new PackageError('Presentation slide size is missing', this.presentationPartUri);
    return {
      width: readSlideSizeCoordinate(xml, size, 'cx', 'width', this.presentationPartUri),
      height: readSlideSizeCoordinate(xml, size, 'cy', 'height', this.presentationPartUri),
    };
  }

  get rtlMode(): boolean | undefined {
    const { xml } = this.parsePresentation();
    const root = presentationRoot(xml, this.presentationPartUri);
    const value = xml.attribute(root, 'rtl')?.value;
    if (value === undefined) return undefined;
    if (['1', 'true', 'on'].includes(value)) return true;
    if (['0', 'false', 'off'].includes(value)) return false;
    return undefined;
  }

  get author(): string | undefined {
    return readPresentationAuthor(this.opcPackage);
  }

  set author(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationAuthor(this.opcPackage, value);
    });
  }

  get company(): string | undefined {
    return readPresentationCompany(this.opcPackage);
  }

  set company(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationCompany(this.opcPackage, value);
    });
  }

  get createdAt(): string | undefined {
    return readPresentationCreatedAt(this.opcPackage);
  }

  set createdAt(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationCreatedAt(this.opcPackage, value);
    });
  }

  get modifiedAt(): string | undefined {
    return readPresentationModifiedAt(this.opcPackage);
  }

  set modifiedAt(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationModifiedAt(this.opcPackage, value);
    });
  }

  get lastModifiedBy(): string | undefined {
    return readPresentationLastModifiedBy(this.opcPackage);
  }

  set lastModifiedBy(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationLastModifiedBy(this.opcPackage, value);
    });
  }

  get revision(): string | undefined {
    return readPresentationRevision(this.opcPackage);
  }

  set revision(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationRevision(this.opcPackage, value);
    });
  }

  get subject(): string | undefined {
    return readPresentationSubject(this.opcPackage);
  }

  set subject(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationSubject(this.opcPackage, value);
    });
  }

  get title(): string | undefined {
    return readPresentationTitle(this.opcPackage);
  }

  set title(value: string | undefined) {
    this.opcPackage.transaction(() => {
      replacePresentationTitle(this.opcPackage, value);
    });
  }

  set rtlMode(value: boolean | undefined) {
    this.opcPackage.transaction(() => {
      if (value !== undefined && typeof value !== 'boolean') {
        throw new TypeError('Presentation RTL mode must be a boolean or undefined');
      }
      const { xml } = this.parsePresentation();
      const root = presentationRoot(xml, this.presentationPartUri);
      updatePresentationAttribute(xml, root, 'rtl', value === undefined ? undefined : value ? '1' : '0');
      this.setXmlPart(this.presentationPartUri, xml.serialize());
    });
  }

  set slideSize(value: SlideSize) {
    this.opcPackage.transaction(() => {
      const normalized = normalizeSlideSize(value);
      const { xml } = this.parsePresentation();
      const root = presentationRoot(xml, this.presentationPartUri);
      const notesSize = directChild(root, 'notesSz');
      if (!notesSize) throw new PackageError('Presentation notes size is missing', this.presentationPartUri);
      const size = directChild(root, 'sldSz');
      if (size) {
        setNumericAttribute(xml, size, 'cx', normalized.width);
        setNumericAttribute(xml, size, 'cy', normalized.height);
      } else {
        xml.replace(
          notesSize.start,
          notesSize.start,
          `<p:sldSz xmlns:p="${PRESENTATION_NAMESPACE}" cx="${normalized.width}" cy="${normalized.height}"/>`,
        );
      }
      this.setXmlPart(this.presentationPartUri, xml.serialize());
    });
  }

  addSlide(): SlideModel {
    return this.opcPackage.transaction(() => {
      const slideUri = this.opcPackage.allocatePartUri(
        joinPartUri(partUriDirname(this.presentationPartUri), 'slides'),
        'slide',
        '.xml',
      );
      this.opcPackage.setPart(slideUri, blankSlideXml(), SLIDE_CONTENT_TYPE);
      const template = this.slides[0];
      const layoutPartUri =
        template?.relationships.find(({ type }) => type === SLIDE_LAYOUT_RELATIONSHIP)?.resolvedTarget ??
        this.defaultLayoutPartUri();
      if (layoutPartUri) {
        this.opcPackage.addRelationship(slideUri, {
          type: SLIDE_LAYOUT_RELATIONSHIP,
          target: relativeRelationshipTarget(slideUri, layoutPartUri),
        });
      }
      return this.attachSlide(slideUri);
    });
  }

  duplicateSlide(index: number): SlideModel {
    return this.opcPackage.transaction(() => {
      const source = this.requireSlide(index);
      const slideUri = this.opcPackage.allocatePartUri(
        joinPartUri(partUriDirname(this.presentationPartUri), 'slides'),
        'slide',
        '.xml',
      );
      const sourcePart = this.opcPackage.requirePart(source.partUri);
      this.opcPackage.setPart(slideUri, sourcePart.bytes, sourcePart.contentType);
      cloneSlideDependencies(this.opcPackage, source.partUri, slideUri);
      return this.attachSlide(slideUri);
    });
  }

  deleteSlide(index: number): void {
    this.opcPackage.transaction(() => {
      const slide = this.requireSlide(index);
      const { xml } = this.parsePresentation();
      const entry = this.slideIdElements(xml).find(
        (element) => xml.attribute(element, 'r:id')?.value === slide.relationshipId,
      );
      if (!entry) throw new PackageError(`Slide entry ${slide.relationshipId} is missing`, this.presentationPartUri);
      const ownedDependencies = ownedSlideDependencyRoots(this.opcPackage, slide.partUri);
      xml.removeElement(entry);
      this.setXmlPart(this.presentationPartUri, xml.serialize());
      this.opcPackage.removeRelationship(this.presentationPartUri, slide.relationshipId);
      this.opcPackage.deletePart(slide.partUri);
      garbageCollectOwnedDependencies(this.opcPackage, ownedDependencies);
    });
  }

  moveSlide(fromIndex: number, toIndex: number): void {
    this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      const list = xml.elements('sldIdLst')[0];
      if (!list) throw new PackageError('Presentation has no slide id list', this.presentationPartUri);
      const elements = this.slideIdElements(xml);
      if (!elements[fromIndex]) throw new RangeError(`Slide index ${fromIndex} is out of range`);
      const boundedTarget = Math.max(0, Math.min(toIndex, elements.length - 1));
      const [moved] = elements.splice(fromIndex, 1);
      elements.splice(boundedTarget, 0, moved!);
      const known = new Set(this.slideIdElements(xml));
      const opaqueChildren = list.children
        .filter((child): child is XmlElement => child.type === 'element' && !known.has(child))
        .map((child) => xml.original(child));
      xml.replace(
        list.startTagEnd,
        list.endTagStart,
        [...elements.map((element) => xml.original(element)), ...opaqueChildren].join(''),
      );
      this.setXmlPart(this.presentationPartUri, xml.serialize());
    });
  }

  setXmlPart(partUri: string, xml: string): void {
    const part = this.opcPackage.requirePart(partUri);
    this.opcPackage.setPart(partUri, xml, part.contentType);
  }

  protected parsePresentation(): { xml: LosslessXmlDocument } {
    return { xml: LosslessXmlDocument.parse(this.opcPackage.requirePart(this.presentationPartUri).bytes) };
  }

  private attachSlide(slideUri: string): SlideModel {
    const relativeTarget = relativeRelationshipTarget(this.presentationPartUri, slideUri);
    const relationship = this.opcPackage.addRelationship(this.presentationPartUri, {
      type: SLIDE_RELATIONSHIP,
      target: relativeTarget,
    });
    const { xml } = this.parsePresentation();
    const root = xml.elements('presentation')[0];
    if (!root) throw new PackageError('Invalid presentation XML', this.presentationPartUri);
    let list = xml.elements('sldIdLst')[0];
    const slideId = Math.max(255, ...this.slides.map(({ slideId: id }) => id)) + 1;
    const entry = `<p:sldId id="${slideId}" r:id="${relationship.id}"/>`;
    if (list) {
      xml.appendChildXml(list, entry);
    } else {
      xml.appendChildXml(root, `<p:sldIdLst>${entry}</p:sldIdLst>`);
      list = undefined;
    }
    this.setXmlPart(this.presentationPartUri, xml.serialize());
    return this.slideModel(slideUri, relationship.id, slideId);
  }

  private defaultLayoutPartUri(): string | undefined {
    for (const master of this.opcPackage
      .relationships(this.presentationPartUri)
      .filter(({ type, resolvedTarget }) => type.endsWith('/slideMaster') && resolvedTarget)) {
      const layout = this.opcPackage
        .relationships(master.resolvedTarget!)
        .find(({ type, resolvedTarget }) => type === SLIDE_LAYOUT_RELATIONSHIP && resolvedTarget);
      if (layout?.resolvedTarget) return layout.resolvedTarget;
    }
    return undefined;
  }

  private slideModel(partUri: string, relationshipId: string, slideId: number): SlideModel {
    const existing = this.#slideModels.get(partUri);
    if (existing) {
      existing.syncIdentity(relationshipId, slideId);
      return existing;
    }
    const created = new SlideModel(this, partUri, relationshipId, slideId);
    this.#slideModels.set(partUri, created);
    return created;
  }

  private slideIdElements(xml: LosslessXmlDocument): XmlElement[] {
    return xml.elements('sldId');
  }

  private requireSlide(index: number): SlideModel {
    const slide = this.slides[index];
    if (!slide) throw new RangeError(`Slide index ${index} is out of range`);
    return slide;
  }
}

function blankSlideXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}

function presentationRoot(xml: LosslessXmlDocument, partUri: string): XmlElement {
  const root = xml.elements('presentation').find(({ parent }) => !parent);
  if (!root) throw new PackageError('Presentation root is missing', partUri);
  return root;
}

function directChild(parent: XmlElement, localName: string): XmlElement | undefined {
  return parent.children.find(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function readSlideSizeCoordinate(
  xml: LosslessXmlDocument,
  size: XmlElement,
  attributeName: 'cx' | 'cy',
  dimensionName: 'width' | 'height',
  partUri: string,
): Emu {
  const raw = xml.attribute(size, attributeName)?.value ?? '';
  const value = Number(raw);
  if (!Number.isInteger(value) || value < MIN_SLIDE_SIZE || value > MAX_SLIDE_SIZE) {
    throw new PackageError(`Presentation slide ${dimensionName} is invalid`, partUri);
  }
  return value as Emu;
}

function normalizeSlideSize(value: SlideSize): SlideSize {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Slide size must contain a width and height');
  }
  return {
    width: normalizeSlideSizeCoordinate(value.width, 'width'),
    height: normalizeSlideSizeCoordinate(value.height, 'height'),
  };
}

function normalizeSlideSizeCoordinate(value: unknown, name: 'width' | 'height'): Emu {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Slide ${name} must be a finite number`);
  }
  const rounded = Math.round(value);
  if (rounded < MIN_SLIDE_SIZE || rounded > MAX_SLIDE_SIZE) {
    throw new RangeError(`Slide ${name} must be between 1 and 56 inches`);
  }
  return rounded as Emu;
}

function setNumericAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: 'cx' | 'cy',
  value: number,
): void {
  const attribute = xml.attribute(element, name);
  if (attribute) {
    xml.replaceAttribute(attribute, String(value));
    return;
  }
  const insertionPoint = element.startTagEnd - (element.selfClosing ? 2 : 1);
  xml.replace(insertionPoint, insertionPoint, ` ${name}="${value}"`);
}

function updatePresentationAttribute(
  xml: LosslessXmlDocument,
  root: XmlElement,
  name: string,
  value: string | undefined,
): void {
  const attribute = xml.attribute(root, name);
  if (value !== undefined) {
    if (attribute) xml.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? xml.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      xml.replace(insertionPoint, insertionPoint, ` ${name}="${value}"`);
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(xml.source[start - 1] ?? '')) start -= 1;
    xml.replace(start, attribute.end, '');
  }
}
