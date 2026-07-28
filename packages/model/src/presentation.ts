import { posix } from 'node:path';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { OpcPackage, PackageError, relationshipPartUri } from '@pptx/opc';
import {
  detectPresentationFormat,
  presentationFormatProfile,
  UnsupportedPresentationFormatError,
  type PresentationFormat,
  type PresentationFormatProfile,
} from './format.js';
import { SlideModel } from './slide.js';

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SLIDE_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const SLIDE_LAYOUT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';

export class PresentationModel {
  readonly presentationPartUri: string;
  readonly format: PresentationFormat;
  readonly formatProfile: PresentationFormatProfile;

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
        slides.push(new SlideModel(this, relationship.resolvedTarget, relationship.id, entry.slideId));
      }
    }
    return slides;
  }

  addSlide(): SlideModel {
    const slideUri = this.opcPackage.allocatePartUri(posix.dirname(this.presentationPartUri) + '/slides', 'slide', '.xml');
    this.opcPackage.setPart(slideUri, blankSlideXml(), SLIDE_CONTENT_TYPE);
    const template = this.slides[0];
    if (template) {
      for (const relationship of template.relationships.filter(({ type }) => type === SLIDE_LAYOUT_RELATIONSHIP)) {
        this.opcPackage.addRelationship(slideUri, {
          type: relationship.type,
          target: relationship.target,
          targetMode: relationship.targetMode,
        });
      }
    }
    return this.attachSlide(slideUri);
  }

  duplicateSlide(index: number): SlideModel {
    const source = this.requireSlide(index);
    const slideUri = this.opcPackage.allocatePartUri(posix.dirname(this.presentationPartUri) + '/slides', 'slide', '.xml');
    const sourcePart = this.opcPackage.requirePart(source.partUri);
    this.opcPackage.setPart(slideUri, sourcePart.bytes, sourcePart.contentType);
    const sourceRelationshipsUri = relationshipPartUri(source.partUri);
    const relationshipsPart = this.opcPackage.getPart(sourceRelationshipsUri);
    if (relationshipsPart) {
      this.opcPackage.setPart(
        relationshipPartUri(slideUri),
        relationshipsPart.bytes,
        relationshipsPart.contentType,
      );
    }
    return this.attachSlide(slideUri);
  }

  deleteSlide(index: number): void {
    const slide = this.requireSlide(index);
    const { xml } = this.parsePresentation();
    const entry = this.slideIdElements(xml).find(
      (element) => xml.attribute(element, 'r:id')?.value === slide.relationshipId,
    );
    if (!entry) throw new PackageError(`Slide entry ${slide.relationshipId} is missing`, this.presentationPartUri);
    xml.removeElement(entry);
    this.setXmlPart(this.presentationPartUri, xml.serialize());
    this.opcPackage.removeRelationship(this.presentationPartUri, slide.relationshipId);
    this.opcPackage.deletePart(slide.partUri);
  }

  moveSlide(fromIndex: number, toIndex: number): void {
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
  }

  setXmlPart(partUri: string, xml: string): void {
    const part = this.opcPackage.requirePart(partUri);
    this.opcPackage.setPart(partUri, xml, part.contentType);
  }

  protected parsePresentation(): { xml: LosslessXmlDocument } {
    return { xml: LosslessXmlDocument.parse(this.opcPackage.requirePart(this.presentationPartUri).bytes) };
  }

  private attachSlide(slideUri: string): SlideModel {
    const relativeTarget = posix.relative(posix.dirname(this.presentationPartUri), slideUri);
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
    return new SlideModel(this, slideUri, relationship.id, slideId);
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
