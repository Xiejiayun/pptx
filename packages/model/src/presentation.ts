import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { readSlideNumber, replaceSlideNumber } from '@pptx/codecs';
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
import { ModelParseError } from './errors.js';
import { removeDrawingHyperlinkReferences } from './shape-hyperlink.internal.js';
import { SlideModel } from './slide.js';
import {
  cloneSlideDependencies,
  garbageCollectMediaDependencies,
  garbageCollectOwnedDependencies,
  mediaSlideDependencyTargets,
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
import {
  assignPresentationSlideToSection,
  copyPresentationSlideSection,
  createPresentationSectionId,
  deletePresentationSection,
  insertPresentationSection,
  movePresentationSection,
  normalizeAddPresentationSectionOptions,
  normalizeAddPresentationSlideOptions,
  normalizePresentationSectionId,
  normalizePresentationSectionTitle,
  readPresentationSections,
  removePresentationSlideFromSections,
  renamePresentationSection,
  sortPresentationSectionSlides,
} from './presentation-sections.internal.js';
import type { Emu, SlideSize } from './units.js';
import {
  normalizeFirstSlideNumber,
  readFirstSlideNumber,
  replaceFirstSlideNumber,
  synchronizeSlideNumberCaches,
} from './presentation-slide-number.internal.js';
import type { RichTextColor } from './text.js';
import { resolveSlideLayoutPartUri } from './presentation-layout.internal.js';
import { materializeLayoutPlaceholders } from './placeholder.internal.js';

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const SLIDE_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const SLIDE_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const SLIDE_LAYOUT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const PRESENTATION_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const MIN_SLIDE_SIZE = 914_400;
const MAX_SLIDE_SIZE = 51_206_400;
const DEFAULT_SECTION_TITLE = /^Default-[1-9]\d*$/;
const DRAWING_HYPERLINK_OWNER_CONTENT_TYPES = new Set([
  SLIDE_CONTENT_TYPE,
  SLIDE_LAYOUT_CONTENT_TYPE,
  SLIDE_MASTER_CONTENT_TYPE,
]);

export interface PresentationSection {
  readonly id: string;
  readonly title: string;
  readonly slideIds: readonly number[];
}

export interface AddSectionOptions {
  readonly title: string;
  readonly order?: number;
}

export interface AddSlideOptions {
  readonly masterName?: string;
  readonly sectionTitle?: string;
}

export interface PreparedSlideInsertionAfter {
  readonly sourcePartUri: string;
  readonly sourceSlideId: number;
  readonly layoutPartUri: string;
  readonly materializeSlideNumber: boolean;
}

export class PresentationModel {
  readonly presentationPartUri: string;
  readonly format: PresentationFormat;
  readonly formatProfile: PresentationFormatProfile;
  readonly #slideModels = new Map<string, SlideModel>();
  readonly #slideDefaultColors = new Map<string, Readonly<RichTextColor>>();

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
    const ordered = this
      .slideIdElements(xml)
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

  get sections(): readonly PresentationSection[] | undefined {
    const { xml } = this.parsePresentation();
    return readPresentationSections(xml, new Set(this.slides.map(({ slideId }) => slideId)));
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

  get firstSlideNumber(): number | undefined {
    return readFirstSlideNumber(this.parsePresentation().xml);
  }

  set firstSlideNumber(value: number | undefined) {
    const normalized = value === undefined ? undefined : normalizeFirstSlideNumber(value);
    this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      if (!replaceFirstSlideNumber(xml, normalized)) return;
      this.setXmlPart(this.presentationPartUri, xml.serialize());
      synchronizeSlideNumberCaches(this);
    });
  }

  /** @internal */
  effectiveSlideNumber(slide: SlideModel): number {
    const index = this.slides.indexOf(slide);
    if (index < 0) {
      throw new Error('Slide is not part of the current presentation');
    }
    const effective = (this.firstSlideNumber ?? 1) + index;
    if (!Number.isSafeInteger(effective)) {
      throw new RangeError('Effective slide number exceeds the JavaScript safe integer range');
    }
    return effective;
  }

  /** @internal */
  getSlideDefaultColor(partUri: string): Readonly<RichTextColor> | undefined {
    return this.#slideDefaultColors.get(partUri);
  }

  /** @internal */
  setSlideDefaultColor(
    partUri: string,
    value: Readonly<RichTextColor> | undefined,
  ): void {
    const current = this.#slideDefaultColors.get(partUri);
    if (richTextColorsEqual(current, value)) return;
    if (value === undefined) this.#slideDefaultColors.delete(partUri);
    else this.#slideDefaultColors.set(partUri, value);
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

  addSection(options: AddSectionOptions): PresentationSection {
    const normalized = normalizeAddPresentationSectionOptions(options);
    return this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
      const sections = this.requireEditableSections(xml, slideIds);
      const section = insertPresentationSection(
        xml,
        slideIds,
        normalized.title,
        normalized.order ?? sections.length,
        createPresentationSectionId(),
      );
      if (xml.changed) this.setXmlPart(this.presentationPartUri, xml.serialize());
      return { ...section, slideIds: [...section.slideIds] };
    });
  }

  renameSection(sectionId: string, title: string): void {
    const normalizedId = normalizePresentationSectionId(sectionId);
    const normalizedTitle = normalizePresentationSectionTitle(title);
    this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
      if (renamePresentationSection(xml, slideIds, normalizedId, normalizedTitle)) {
        this.setXmlPart(this.presentationPartUri, xml.serialize());
      }
    });
  }

  moveSection(sectionId: string, toIndex: number): void {
    const normalizedId = normalizePresentationSectionId(sectionId);
    this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
      if (movePresentationSection(xml, slideIds, normalizedId, toIndex)) {
        this.setXmlPart(this.presentationPartUri, xml.serialize());
      }
    });
  }

  deleteSection(sectionId: string): void {
    const normalizedId = normalizePresentationSectionId(sectionId);
    this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
      if (deletePresentationSection(xml, slideIds, normalizedId)) {
        this.setXmlPart(this.presentationPartUri, xml.serialize());
      }
    });
  }

  assignSlideToSection(slideIndex: number, sectionId: string | undefined): void {
    const normalizedId = sectionId === undefined
      ? undefined
      : normalizePresentationSectionId(sectionId);
    this.opcPackage.transaction(() => {
      const slide = this.requireSlide(slideIndex);
      const { xml } = this.parsePresentation();
      const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
      if (assignPresentationSlideToSection(xml, slideIds, slide.slideId, normalizedId)) {
        this.setXmlPart(this.presentationPartUri, xml.serialize());
      }
    });
  }

  addSlide(options: AddSlideOptions = {}): SlideModel {
    const normalized = normalizeAddPresentationSlideOptions(options);
    const inheritedLayoutPartUri = normalized.masterName === undefined
      ? this.inheritedLayoutPartUri()
      : undefined;
    const layoutPartUri = resolveSlideLayoutPartUri(
      this.opcPackage,
      this.presentationPartUri,
      normalized.masterName,
      inheritedLayoutPartUri,
    );
    const slide = this.opcPackage.transaction(() => {
      const initial = this.parsePresentation().xml;
      const initialSlideIds = new Set(this.slides.map(({ slideId }) => slideId));
      const sections = this.requireEditableSections(initial, initialSlideIds);
      let targetSectionId: string | undefined;
      if (normalized.sectionTitle !== undefined) {
        const target = sections.find(({ title }) => title === normalized.sectionTitle);
        if (!target) {
          throw new RangeError(`Presentation section ${normalized.sectionTitle} was not found`);
        }
        targetSectionId = target.id;
      } else if (sections.length > 0) {
        const last = sections.at(-1)!;
        if (DEFAULT_SECTION_TITLE.test(last.title)) {
          targetSectionId = last.id;
        } else {
          const titles = new Set(sections.map(({ title }) => title));
          let number = 1;
          while (titles.has(`Default-${number}`)) number += 1;
          const section = insertPresentationSection(
            initial,
            initialSlideIds,
            `Default-${number}`,
            sections.length,
            createPresentationSectionId(),
          );
          targetSectionId = section.id;
          this.setXmlPart(this.presentationPartUri, initial.serialize());
        }
      }

      const slideUri = this.opcPackage.allocatePartUri(
        joinPartUri(partUriDirname(this.presentationPartUri), 'slides'),
        'slide',
        '.xml',
      );
      this.opcPackage.setPart(slideUri, blankSlideXml(), SLIDE_CONTENT_TYPE);
      const layoutSlideNumber = layoutPartUri === undefined
        ? undefined
        : readSlideNumber(this.opcPackage, layoutPartUri, 'layout');
      if (layoutPartUri) {
        this.opcPackage.addRelationship(slideUri, {
          type: SLIDE_LAYOUT_RELATIONSHIP,
          target: relativeRelationshipTarget(slideUri, layoutPartUri),
        });
        materializeLayoutPlaceholders(
          this.opcPackage,
          layoutPartUri,
          slideUri,
          normalized.masterName !== undefined,
          layoutSlideNumber !== undefined,
        );
      }
      const slide = this.attachSlide(slideUri);
      if (layoutSlideNumber !== undefined) {
        replaceSlideNumber(
          this.opcPackage,
          slideUri,
          'slide',
          layoutSlideNumber,
          String(this.effectiveSlideNumber(slide)),
        );
      }
      if (targetSectionId !== undefined) {
        const { xml } = this.parsePresentation();
        const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
        if (assignPresentationSlideToSection(xml, slideIds, slide.slideId, targetSectionId)) {
          this.setXmlPart(this.presentationPartUri, xml.serialize());
        }
      }
      return slide;
    });
    this.#slideDefaultColors.delete(slide.partUri);
    return slide;
  }

  /** @internal */
  prepareSlideInsertionAfter(source: SlideModel): PreparedSlideInsertionAfter {
    this.requireAttachedSlideIndex(source, 'Source slide');
    const { xml } = this.parsePresentation();
    this.requireEditableSections(
      xml,
      new Set(this.slides.map(({ slideId }) => slideId)),
    );
    const layoutPartUri = this.requireDirectSlideLayoutPartUri(source);
    const materializeSlideNumber =
      readSlideNumber(this.opcPackage, layoutPartUri, 'layout') !== undefined;
    return Object.freeze({
      sourcePartUri: source.partUri,
      sourceSlideId: source.slideId,
      layoutPartUri,
      materializeSlideNumber,
    });
  }

  /** @internal */
  insertPreparedBlankSlideAfter(
    after: SlideModel,
    prepared: Readonly<PreparedSlideInsertionAfter>,
  ): SlideModel {
    const inserted = this.opcPackage.transaction(() => {
      const source = this.requirePreparedInsertionSource(prepared);
      const insertionIndex = this.requireAttachedSlideIndex(after, 'Insertion slide') + 1;
      const initial = this.parsePresentation().xml;
      this.requireEditableSections(
        initial,
        new Set(this.slides.map(({ slideId }) => slideId)),
      );
      const layoutSlideNumber = readSlideNumber(
        this.opcPackage,
        prepared.layoutPartUri,
        'layout',
      );
      if ((layoutSlideNumber !== undefined) !== prepared.materializeSlideNumber) {
        throw new ModelParseError(
          'Prepared slide layout number state has changed',
          prepared.layoutPartUri,
        );
      }

      const slideUri = this.opcPackage.allocatePartUri(
        joinPartUri(partUriDirname(this.presentationPartUri), 'slides'),
        'slide',
        '.xml',
      );
      this.opcPackage.setPart(slideUri, blankSlideXml(), SLIDE_CONTENT_TYPE);
      this.opcPackage.addRelationship(slideUri, {
        type: SLIDE_LAYOUT_RELATIONSHIP,
        target: relativeRelationshipTarget(slideUri, prepared.layoutPartUri),
      });
      materializeLayoutPlaceholders(
        this.opcPackage,
        prepared.layoutPartUri,
        slideUri,
        false,
        prepared.materializeSlideNumber,
      );
      const slide = this.attachSlide(slideUri, insertionIndex);
      if (layoutSlideNumber !== undefined) {
        replaceSlideNumber(
          this.opcPackage,
          slideUri,
          'slide',
          layoutSlideNumber,
          String(this.effectiveSlideNumber(slide)),
        );
      }

      const sectionXml = this.parsePresentation().xml;
      const sectionSlideIds = new Set(this.slides.map(({ slideId }) => slideId));
      if (copyPresentationSlideSection(
        sectionXml,
        sectionSlideIds,
        source.slideId,
        slide.slideId,
      )) {
        this.setXmlPart(this.presentationPartUri, sectionXml.serialize());
      }
      const sortedXml = this.parsePresentation().xml;
      const orderedSlideIds = this.slides.map(({ slideId }) => slideId);
      if (sortPresentationSectionSlides(
        sortedXml,
        new Set(orderedSlideIds),
        orderedSlideIds,
      )) {
        this.setXmlPart(this.presentationPartUri, sortedXml.serialize());
      }
      synchronizeSlideNumberCaches(this);
      return slide;
    });
    this.#slideDefaultColors.delete(inserted.partUri);
    return inserted;
  }

  /** @internal */
  discardDetachedSlideModel(partUri: string): void {
    const model = this.#slideModels.get(partUri);
    if (!model || this.slides.some((slide) => slide === model)) return;
    this.#slideModels.delete(partUri);
    this.#slideDefaultColors.delete(partUri);
  }

  duplicateSlide(index: number): SlideModel {
    let allocatedPartUri: string | undefined;
    let previousModel: SlideModel | undefined;
    let previousDefaultColor: Readonly<RichTextColor> | undefined;
    let sourceDefaultColor: Readonly<RichTextColor> | undefined;
    try {
      const duplicate = this.opcPackage.transaction(() => {
        const source = this.requireSlide(index);
        sourceDefaultColor = this.#slideDefaultColors.get(source.partUri);
        const before = this.parsePresentation().xml;
        this.requireEditableSections(
          before,
          new Set(this.slides.map(({ slideId }) => slideId)),
        );
        const slideUri = this.opcPackage.allocatePartUri(
          joinPartUri(partUriDirname(this.presentationPartUri), 'slides'),
          'slide',
          '.xml',
        );
        allocatedPartUri = slideUri;
        previousModel = this.#slideModels.get(slideUri);
        previousDefaultColor = this.#slideDefaultColors.get(slideUri);
        const sourcePart = this.opcPackage.requirePart(source.partUri);
        this.opcPackage.setPart(slideUri, sourcePart.bytes, sourcePart.contentType);
        cloneSlideDependencies(this.opcPackage, source.partUri, slideUri);
        const duplicate = this.attachSlide(slideUri);
        const { xml } = this.parsePresentation();
        const slideIds = new Set(this.slides.map(({ slideId }) => slideId));
        if (copyPresentationSlideSection(xml, slideIds, source.slideId, duplicate.slideId)) {
          this.setXmlPart(this.presentationPartUri, xml.serialize());
        }
        synchronizeSlideNumberCaches(this);
        return duplicate;
      });
      if (sourceDefaultColor === undefined) {
        this.#slideDefaultColors.delete(duplicate.partUri);
      } else {
        this.#slideDefaultColors.set(duplicate.partUri, sourceDefaultColor);
      }
      return duplicate;
    } catch (error) {
      if (allocatedPartUri) {
        if (previousModel) this.#slideModels.set(allocatedPartUri, previousModel);
        else this.#slideModels.delete(allocatedPartUri);
        if (previousDefaultColor === undefined) {
          this.#slideDefaultColors.delete(allocatedPartUri);
        } else {
          this.#slideDefaultColors.set(allocatedPartUri, previousDefaultColor);
        }
      }
      throw error;
    }
  }

  deleteSlide(index: number): void {
    const slide = this.requireSlide(index);
    this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      const entry = this.slideIdElements(xml).find(
        (element) => xml.attribute(element, 'r:id')?.value === slide.relationshipId,
      );
      if (!entry) throw new PackageError(`Slide entry ${slide.relationshipId} is missing`, this.presentationPartUri);
      const ownedDependencies = ownedSlideDependencyRoots(this.opcPackage, slide.partUri);
      const mediaDependencies = mediaSlideDependencyTargets(this.opcPackage, slide.partUri);
      const incomingBySource = new Map<string, Set<string>>();
      const incoming = this.opcPackage.graph.find(({ uri }) => uri === slide.partUri)?.incoming ?? [];
      for (const { sourceUri, relationship } of incoming) {
        if (
          sourceUri === slide.partUri
          || sourceUri === this.presentationPartUri
          || relationship.type !== SLIDE_RELATIONSHIP
          || relationship.targetMode !== 'Internal'
        ) continue;
        const sourcePart = this.opcPackage.getPart(sourceUri);
        if (
          !sourcePart
          || !DRAWING_HYPERLINK_OWNER_CONTENT_TYPES.has(sourcePart.contentType)
        ) continue;
        const relationshipIds = incomingBySource.get(sourceUri) ?? new Set<string>();
        relationshipIds.add(relationship.id);
        incomingBySource.set(sourceUri, relationshipIds);
      }
      for (const [sourceUri, relationshipIds] of incomingBySource) {
        const sourcePart = this.opcPackage.requirePart(sourceUri);
        const sourceXml = LosslessXmlDocument.parse(sourcePart.bytes);
        if (removeDrawingHyperlinkReferences(sourceXml, relationshipIds)) {
          this.setXmlPart(sourceUri, sourceXml.serialize());
        }
      }
      removePresentationSlideFromSections(
        xml,
        new Set(this.slides.map(({ slideId }) => slideId)),
        slide.slideId,
      );
      xml.removeElement(entry);
      this.setXmlPart(this.presentationPartUri, xml.serialize());
      this.opcPackage.removeRelationship(this.presentationPartUri, slide.relationshipId);
      this.opcPackage.deletePart(slide.partUri);
      garbageCollectOwnedDependencies(this.opcPackage, ownedDependencies);
      garbageCollectMediaDependencies(this.opcPackage, mediaDependencies);
      synchronizeSlideNumberCaches(this);
    });
    this.#slideDefaultColors.delete(slide.partUri);
  }

  moveSlide(fromIndex: number, toIndex: number): void {
    this.opcPackage.transaction(() => {
      const { xml } = this.parsePresentation();
      const validSlideIds = new Set(this.slides.map(({ slideId }) => slideId));
      this.requireEditableSections(xml, validSlideIds);
      const list = this.slideIdList(xml);
      if (!list) throw new PackageError('Presentation has no slide id list', this.presentationPartUri);
      const elements = this.slideIdElements(xml);
      if (!elements[fromIndex]) throw new RangeError(`Slide index ${fromIndex} is out of range`);
      const boundedTarget = Math.max(0, Math.min(toIndex, elements.length - 1));
      const [moved] = elements.splice(fromIndex, 1);
      elements.splice(boundedTarget, 0, moved!);
      const known = new Set(elements);
      const opaqueChildren = list.children
        .filter((child): child is XmlElement => child.type === 'element' && !known.has(child))
        .map((child) => xml.original(child));
      xml.replace(
        list.startTagEnd,
        list.endTagStart,
        [...elements.map((element) => xml.original(element)), ...opaqueChildren].join(''),
      );
      const orderedSlideIds = elements.map((element) => {
        const id = Number(xml.attribute(element, 'id')?.value);
        if (!Number.isSafeInteger(id) || !validSlideIds.has(id)) {
          throw new PackageError('Presentation slide ID is invalid', this.presentationPartUri);
        }
        return id;
      });
      sortPresentationSectionSlides(xml, validSlideIds, orderedSlideIds);
      this.setXmlPart(this.presentationPartUri, xml.serialize());
      synchronizeSlideNumberCaches(this);
    });
  }

  setXmlPart(partUri: string, xml: string): void {
    const part = this.opcPackage.requirePart(partUri);
    this.opcPackage.setPart(partUri, xml, part.contentType);
  }

  protected parsePresentation(): { xml: LosslessXmlDocument } {
    return { xml: LosslessXmlDocument.parse(this.opcPackage.requirePart(this.presentationPartUri).bytes) };
  }

  private attachSlide(slideUri: string, insertionIndex = this.slides.length): SlideModel {
    const currentSlides = this.slides;
    if (
      !Number.isSafeInteger(insertionIndex)
      || insertionIndex < 0
      || insertionIndex > currentSlides.length
    ) {
      throw new RangeError('Slide insertion index is out of range');
    }
    const relativeTarget = relativeRelationshipTarget(this.presentationPartUri, slideUri);
    const relationship = this.opcPackage.addRelationship(this.presentationPartUri, {
      type: SLIDE_RELATIONSHIP,
      target: relativeTarget,
    });
    const { xml } = this.parsePresentation();
    const root = xml.elements('presentation')[0];
    if (!root) throw new PackageError('Invalid presentation XML', this.presentationPartUri);
    let list = this.slideIdList(xml);
    const slideId = Math.max(255, ...currentSlides.map(({ slideId: id }) => id)) + 1;
    if (!Number.isSafeInteger(slideId)) {
      throw new RangeError('Presentation slide IDs are exhausted');
    }
    const entry = `<p:sldId id="${slideId}" r:id="${relationship.id}"/>`;
    if (list) {
      const ordered = this.slideIdElements(xml);
      if (insertionIndex === ordered.length) {
        xml.appendChildXml(list, entry);
      } else {
        const following = ordered[insertionIndex];
        if (!following) {
          throw new ModelParseError(
            'Presentation slide order is not safely editable',
            this.presentationPartUri,
          );
        }
        xml.replace(following.start, following.start, entry);
      }
    } else {
      if (insertionIndex !== 0) {
        throw new ModelParseError(
          'Presentation slide order is not safely editable',
          this.presentationPartUri,
        );
      }
      xml.appendChildXml(root, `<p:sldIdLst>${entry}</p:sldIdLst>`);
      list = undefined;
    }
    this.setXmlPart(this.presentationPartUri, xml.serialize());
    return this.slideModel(slideUri, relationship.id, slideId);
  }

  private requireAttachedSlideIndex(slide: SlideModel, context: string): number {
    if (slide.presentation !== this) {
      throw new Error(`${context} belongs to a different presentation`);
    }
    const index = this.slides.indexOf(slide);
    if (index < 0) throw new Error(`${context} is not attached to the current presentation`);
    return index;
  }

  private requireDirectSlideLayoutPartUri(source: SlideModel): string {
    const relationships = source.relationships.filter(
      ({ type }) => type === SLIDE_LAYOUT_RELATIONSHIP,
    );
    if (relationships.length !== 1) {
      throw new ModelParseError(
        'Source slide layout relationship is missing or ambiguous',
        source.partUri,
      );
    }
    const relationship = relationships[0]!;
    if (relationship.targetMode !== 'Internal' || !relationship.resolvedTarget) {
      throw new ModelParseError(
        'Source slide layout relationship must be internal',
        source.partUri,
      );
    }
    const target = this.opcPackage.getPart(relationship.resolvedTarget);
    if (target?.contentType !== SLIDE_LAYOUT_CONTENT_TYPE) {
      throw new ModelParseError(
        'Source slide layout relationship has an invalid target',
        source.partUri,
      );
    }
    return target.uri;
  }

  private requirePreparedInsertionSource(
    prepared: Readonly<PreparedSlideInsertionAfter>,
  ): SlideModel {
    const source = this.slides.find((slide) =>
      slide.partUri === prepared.sourcePartUri
      && slide.slideId === prepared.sourceSlideId);
    if (!source) {
      throw new Error('Prepared source slide is no longer attached to the current presentation');
    }
    const layoutPartUri = this.requireDirectSlideLayoutPartUri(source);
    if (layoutPartUri !== prepared.layoutPartUri) {
      throw new ModelParseError(
        'Prepared source slide layout has changed',
        source.partUri,
      );
    }
    return source;
  }

  private inheritedLayoutPartUri(): string | undefined {
    const template = this.slides[0];
    if (!template) return undefined;
    const layouts = template.relationships.filter((relationship) => {
      if (
        relationship.type !== SLIDE_LAYOUT_RELATIONSHIP
        || relationship.targetMode !== 'Internal'
        || !relationship.resolvedTarget
      ) return false;
      return this.opcPackage.getPart(relationship.resolvedTarget)?.contentType ===
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
    });
    if (layouts.length > 1) {
      throw new ModelParseError('Source slide layout is ambiguous', template.partUri);
    }
    return layouts[0]?.resolvedTarget;
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
    const list = this.slideIdList(xml);
    if (!list) return [];
    return list.children.filter(
      (child): child is XmlElement => child.type === 'element' && child.localName === 'sldId',
    );
  }

  private slideIdList(xml: LosslessXmlDocument): XmlElement | undefined {
    const root = xml.elements('presentation').find(({ parent }) => !parent);
    return root ? directChild(root, 'sldIdLst') : undefined;
  }

  private requireEditableSections(
    xml: LosslessXmlDocument,
    slideIds: ReadonlySet<number>,
  ): readonly PresentationSection[] {
    if (!xml.elements('presentation').some(({ parent }) => !parent)) {
      throw new PackageError('Invalid presentation XML', this.presentationPartUri);
    }
    const sections = readPresentationSections(xml, slideIds);
    if (!sections) {
      throw new ModelParseError(
        'Presentation sections are not safely editable',
        this.presentationPartUri,
      );
    }
    return sections;
  }

  private requireSlide(index: number): SlideModel {
    const slide = this.slides[index];
    if (!slide) throw new RangeError(`Slide index ${index} is out of range`);
    return slide;
  }
}

function richTextColorsEqual(
  left: Readonly<RichTextColor> | undefined,
  right: Readonly<RichTextColor> | undefined,
): boolean {
  return left === right
    || (
      left !== undefined
      && right !== undefined
      && left.kind === right.kind
      && left.value === right.value
    );
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
