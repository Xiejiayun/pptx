import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import {
  MediaCodec,
  normalizeSlideNumberOptions,
  readSlideNumber,
  replaceSlideNumber,
  type AddMediaOptions,
  type MediaPlaybackSettings,
  type MediaKind,
  type MediaSource,
  type ReplaceMediaSourceOptions,
  type ReplaceMediaPosterOptions,
  type SlideNumber,
  type SlideNumberOptions,
} from '@pptx/codecs';
import {
  relativeRelationshipTarget,
  type Relationship,
  type RelationshipInput,
} from '@pptx/opc';
import type {
  AddChartOptions,
  ChartDefinition,
  ChartGroupInput,
  ChartSeriesInput,
  ChartState,
  ChartType,
} from './chart.js';
import { normalizeChartDefinition } from './chart-definition.internal.js';
import { replaceChartDefinition } from './chart-edit.internal.js';
import {
  normalizeAddChartOptions,
  renderChartGraphicFrame,
  renderChartPart,
} from './chart-render.internal.js';
import {
  buildChartWorkbook,
  planChartWorkbook,
} from './chart-workbook.internal.js';
import { ModelParseError } from './errors.js';
import { garbageCollectOwnedDependencies } from './dependency.internal.js';
import type { Hyperlink } from './hyperlink.js';
import type {
  AddImageOptions,
  AddSvgImageOptions,
  ImageSourceRectangle,
} from './image.js';
import {
  normalizeEmbeddedRasterImage,
  renderEmbeddedRasterImageXml,
} from './image-create.internal.js';
import {
  normalizeEmbeddedSvgImage,
  renderEmbeddedSvgImageXml,
} from './svg-image-create.internal.js';
import {
  normalizeImageSourceRectangle,
  readImageSourceRectangle,
  replaceImageSourceRectangle,
} from './image-source-rectangle.internal.js';
import type { PresentationModel } from './presentation.js';
import {
  normalizeParagraphBullet,
  normalizeParagraphIndent,
  normalizeParagraphLevel,
  normalizeParagraphMarginLeft,
  normalizeParagraphMarginRight,
  normalizeParagraphRtl,
  normalizeParagraphSpacing,
  normalizeParagraphTabStops,
  normalizeRichText,
  normalizeRichTextColor,
  normalizeTextAlignment,
  normalizeTextLanguage,
  readRichText,
  renderParagraphProperties,
  renderRichTextParagraphs,
  resolveParagraphSpacing,
  replaceRichText,
  type NormalizedParagraphBullet,
  type NormalizedParagraphSpacing,
  type NormalizedParagraphSpacingUpdate,
  type NormalizedParagraphTabStop,
} from './rich-text.internal.js';
import {
  ChartModel,
  decodeShape,
  ImageModel,
  MediaModel,
  ShapeModel,
  TableModel,
  type SemanticShape,
  type TableCellBorderInput,
  type TableCellFill,
  type TableCellTextDirection,
} from './shapes.js';
import {
  normalizeTableDefinition,
  renderTableGraphicFrame,
} from './table-create.internal.js';
import {
  readSlideHidden,
  replaceSlideHidden,
} from './slide-visibility.internal.js';
import {
  normalizeSlideNotes,
  readSlideNotes,
  replaceSlideNotes,
} from './slide-notes.internal.js';
import {
  readSlideBackground,
  replaceSlideBackground,
} from './slide-background.internal.js';
import type { SlideBackground } from './slide-background.js';
import {
  normalizeCustomShape,
  normalizePresetShape,
  readPresetShapeType,
  renderCustomShapeXml,
  renderPresetShapeXml,
  replacePresetShapeType,
} from './preset-shape.internal.js';
import type {
  AddCustomShapeOptions,
  CustomGeometry,
} from './custom-geometry.js';
import {
  normalizeCustomGeometry,
  readCustomGeometry,
  replaceCustomGeometry,
} from './custom-geometry.internal.js';
import type {
  AddShapeOptions,
  PresetShapeType,
  ShapeAdjustment,
  ShapeArrows,
  ShapeFill,
  ShapeLine,
  ShapeShadow,
} from './preset-shape.js';
import {
  normalizeShapeAdjustments,
  readShapeAdjustments,
  replaceShapeAdjustments,
} from './shape-adjustments.internal.js';
import {
  normalizeShapeArrows,
  readShapeArrows,
  replaceShapeArrows,
} from './shape-arrows.internal.js';
import {
  readShapeFill,
  replaceShapeFill,
} from './shape-fill.internal.js';
import { normalizeSimpleFill } from './simple-fill.internal.js';
import {
  readShapeLine,
  replaceShapeLine,
} from './shape-line.internal.js';
import { normalizeSimpleLine } from './simple-line.internal.js';
import {
  readShapeShadow,
  replaceShapeShadow,
} from './shape-shadow.internal.js';
import { normalizeShapeShadow } from './simple-shadow.internal.js';
import {
  HYPERLINK_RELATIONSHIP_TYPE,
  normalizeHyperlink,
  readShapeHyperlink,
  relationshipReferenceCount,
  replaceShapeHyperlinkElement,
  requireShapeHyperlinkRelationshipId,
  shapeHyperlinksEqual,
  SLIDE_RELATIONSHIP_TYPE,
  type NormalizedHyperlink,
} from './shape-hyperlink.internal.js';
import {
  normalizeTextBoxMargins,
  readTextBoxMargins,
  renderTextBoxMarginAttributes,
  replaceTextBoxMargins,
} from './text-box-margins.internal.js';
import {
  normalizeTextBoxFit,
  readTextBoxFit,
  renderTextBoxFitChild,
  replaceTextBoxFit,
} from './text-box-fit.internal.js';
import {
  normalizeTextBoxTextDirection,
  readTextBoxTextDirection,
  renderTextBoxTextDirectionAttribute,
  replaceTextBoxTextDirection,
} from './text-box-text-direction.internal.js';
import {
  normalizeTextBoxVerticalAlignment,
  readTextBoxVerticalAlignment,
  renderTextBoxVerticalAlignmentAttribute,
  replaceTextBoxVerticalAlignment,
} from './text-box-vertical-alignment.internal.js';
import {
  normalizeTextBoxWrap,
  readTextBoxWrap,
  renderTextBoxWrapAttribute,
  replaceTextBoxWrap,
} from './text-box-wrapping.internal.js';
import type {
  ParagraphBullet,
  ParagraphSpacing,
  ParagraphTabStop,
  RichTextColor,
  RichTextParagraph,
  TextBoxFit,
  TextBoxMarginInput,
  TextBoxMargins,
  TextBoxTextDirection,
  TextBoxVerticalAlignment,
  TextAlignment,
} from './text.js';
import { inches, type Transform } from './units.js';

const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const CHART_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const PACKAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export interface AddTextOptions extends Partial<Transform> {
  readonly name?: string;
  readonly align?: TextAlignment;
  readonly bullet?: ParagraphBullet;
  readonly fit?: TextBoxFit;
  readonly lang?: string;
  readonly level?: number;
  readonly margin?: TextBoxMarginInput;
  readonly paragraphIndent?: number;
  readonly paragraphMarginLeft?: number;
  readonly paragraphMarginRight?: number;
  readonly rtlMode?: boolean;
  readonly spacing?: ParagraphSpacing;
  readonly tabStops?: readonly ParagraphTabStop[];
  readonly valign?: TextBoxVerticalAlignment;
  readonly vert?: TextBoxTextDirection;
  readonly wrap?: boolean;
}

export interface AddTableOptions {
  readonly align?: TextAlignment;
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly rowHeights?: number | readonly number[];
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}

export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly fit?: TextBoxFit;
  readonly margin?: TextBoxMarginInput;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}

export interface AddTableCell {
  readonly text: string;
  readonly options?: AddTableCellOptions;
}

export type AddTableCellInput = string | AddTableCell;

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

  get background(): SlideBackground | undefined {
    return readSlideBackground(this.presentation.opcPackage, this.partUri);
  }

  set background(value: SlideBackground | undefined) {
    replaceSlideBackground(this.presentation.opcPackage, this.partUri, value);
  }

  get color(): Readonly<RichTextColor> | undefined {
    return this.presentation.getSlideDefaultColor(this.partUri);
  }

  set color(value: RichTextColor | undefined) {
    const normalized = value === undefined
      ? undefined
      : normalizeRichTextColor(value, 'Slide default text color');
    this.presentation.setSlideDefaultColor(this.partUri, normalized);
  }

  get slideNumber(): Readonly<SlideNumber> | undefined {
    return readSlideNumber(this.presentation.opcPackage, this.partUri, 'slide');
  }

  set slideNumber(value: SlideNumberOptions | undefined) {
    const normalized = value === undefined
      ? undefined
      : normalizeSlideNumberOptions(value);
    replaceSlideNumber(
      this.presentation.opcPackage,
      this.partUri,
      'slide',
      normalized,
      String(this.presentation.effectiveSlideNumber(this)),
    );
  }

  get hidden(): boolean | undefined {
    const { xml } = this.parse();
    return readSlideHidden(xml);
  }

  set hidden(value: boolean) {
    if (typeof value !== 'boolean') {
      throw new TypeError('Slide hidden state must be a boolean');
    }
    this.presentation.opcPackage.transaction(() => {
      const { xml } = this.parse();
      if (!replaceSlideHidden(xml, value)) return;
      this.setXml(xml.serialize());
    });
  }

  get notes(): string | undefined {
    return readSlideNotes(
      this.presentation.opcPackage,
      this.presentation.presentationPartUri,
      this.partUri,
    );
  }

  set notes(value: string | undefined) {
    replaceSlideNotes(
      this.presentation.opcPackage,
      this.presentation.presentationPartUri,
      this.partUri,
      value,
    );
  }

  addNotes(value: string): this {
    const normalized = normalizeSlideNotes(value, 'Slide notes');
    replaceSlideNotes(
      this.presentation.opcPackage,
      this.presentation.presentationPartUri,
      this.partUri,
      normalized,
    );
    return this;
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

  get media(): readonly MediaModel[] {
    return this.shapes.filter((shape): shape is MediaModel => shape instanceof MediaModel);
  }

  async addAudio(source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const descriptor = await new MediaCodec(this.presentation.opcPackage)
      .addAudio(this.partUri, source, options);
    return this.requireMedia(descriptor.shapeId);
  }

  async addVideo(source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const descriptor = await new MediaCodec(this.presentation.opcPackage)
      .addVideo(this.partUri, source, options);
    return this.requireMedia(descriptor.shapeId);
  }

  setMediaName(shapeId: number, value: string): void {
    new MediaCodec(this.presentation.opcPackage).setName(this.partUri, shapeId, value);
  }

  setMediaAltText(shapeId: number, value: string | undefined): void {
    new MediaCodec(this.presentation.opcPackage).setAltText(this.partUri, shapeId, value);
  }

  setMediaSettings(
    shapeId: number,
    value: MediaPlaybackSettings | undefined,
  ): void {
    new MediaCodec(this.presentation.opcPackage).setSettings(this.partUri, shapeId, value);
  }

  async replaceMediaSource(
    shapeId: number,
    kind: MediaKind,
    source: MediaSource,
    options: ReplaceMediaSourceOptions = {},
  ): Promise<void> {
    await new MediaCodec(this.presentation.opcPackage)
      .replaceSource(this.partUri, shapeId, kind, source, options);
  }

  async replaceMediaPoster(
    shapeId: number,
    source?: MediaSource,
    options: ReplaceMediaPosterOptions = {},
  ): Promise<void> {
    await new MediaCodec(this.presentation.opcPackage)
      .replacePoster(this.partUri, shapeId, source, options);
  }

  deleteMedia(shapeId: number): void {
    new MediaCodec(this.presentation.opcPackage).delete(this.partUri, shapeId);
  }

  get opaqueExtensionCount(): number {
    const { xml } = this.parse();
    return xml.elements('extLst').length + xml.elements('AlternateContent').length;
  }

  private requireMedia(shapeId: number): MediaModel {
    const model = this.shapes.find((shape) => shape.id === shapeId);
    if (!(model instanceof MediaModel)) {
      throw new ModelParseError(`Media shape ${shapeId} was not found`, this.partUri);
    }
    return model;
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

  getShapePresetType(id: number): PresetShapeType | undefined {
    const { xml, element } = this.resolveShape(id);
    return readPresetShapeType(xml, element);
  }

  getShapeCustomGeometry(id: number): CustomGeometry | undefined {
    const { xml, element } = this.resolveShape(id);
    return readCustomGeometry(xml, element);
  }

  setShapeCustomGeometry(id: number, value: CustomGeometry): void {
    const geometry = normalizeCustomGeometry(value, 'Custom geometry');
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceCustomGeometry(xml, element, geometry, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  setShapePresetType(id: number, value: PresetShapeType): void {
    const type = normalizePresetShape(value, undefined).type;
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replacePresetShapeType(xml, element, type, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  getShapeAdjustments(id: number): readonly ShapeAdjustment[] | undefined {
    const { xml, element } = this.resolveShape(id);
    return readShapeAdjustments(xml, element);
  }

  setShapeAdjustments(id: number, value: readonly ShapeAdjustment[]): void {
    const adjustments = normalizeShapeAdjustments(value, 'Shape adjustments');
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceShapeAdjustments(xml, element, adjustments, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  getShapeFill(id: number): ShapeFill | undefined {
    const { xml, element } = this.resolveShape(id);
    return readShapeFill(xml, element);
  }

  setShapeFill(id: number, value: ShapeFill | undefined): void {
    const fill = value === undefined
      ? undefined
      : normalizeSimpleFill(value, 'Shape fill');
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceShapeFill(xml, element, fill, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  getShapeLine(id: number): ShapeLine | undefined {
    const { xml, element } = this.resolveShape(id);
    return readShapeLine(xml, element);
  }

  setShapeLine(id: number, value: ShapeLine | undefined): void {
    const line = value === undefined
      ? undefined
      : normalizeSimpleLine(value, 'Shape line');
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceShapeLine(xml, element, line, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  getShapeArrows(id: number): ShapeArrows | undefined {
    const { xml, element } = this.resolveShape(id);
    return readShapeArrows(xml, element);
  }

  setShapeArrows(id: number, value: ShapeArrows | undefined): void {
    const arrows = value === undefined
      ? undefined
      : normalizeShapeArrows(value, 'Shape arrows');
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceShapeArrows(xml, element, arrows, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  getShapeShadow(id: number): ShapeShadow | undefined {
    const { xml, element } = this.resolveShape(id);
    return readShapeShadow(xml, element);
  }

  setShapeShadow(id: number, value: ShapeShadow | undefined): void {
    const shadow = value === undefined
      ? undefined
      : normalizeShapeShadow(value, 'Shape shadow');
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceShapeShadow(xml, element, shadow, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  getShapeHyperlink(id: number): Hyperlink | undefined {
    const { xml, element } = this.resolveShape(id);
    return readShapeHyperlink(xml, element, {
      relationships: this.relationships,
      slidePartUris: this.presentation.slides.map(({ partUri }) => partUri),
    });
  }

  setShapeHyperlink(id: number, value: Hyperlink | undefined): void {
    const hyperlink = value === undefined
      ? undefined
      : normalizeHyperlink(value, 'Shape hyperlink');
    let targetSlide: SlideModel | undefined;
    if (hyperlink?.slide !== undefined) {
      targetSlide = this.presentation.slides[hyperlink.slide - 1];
      if (!targetSlide) {
        throw new RangeError(`Shape hyperlink slide ${hyperlink.slide} is out of range`);
      }
    }

    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      const relationshipId = requireShapeHyperlinkRelationshipId(element, this.partUri);
      const current = readShapeHyperlink(xml, element, {
        relationships: this.relationships,
        slidePartUris: this.presentation.slides.map(({ partUri }) => partUri),
      });
      if (relationshipId !== undefined && current === undefined) {
        throw new ModelParseError('Shape hyperlink state is not safely editable', this.partUri);
      }
      if (shapeHyperlinksEqual(current, hyperlink)) return;

      if (hyperlink === undefined) {
        replaceShapeHyperlinkElement(xml, element, undefined, undefined, this.partUri);
        const updated = xml.serialize();
        this.setXml(updated);
        if (
          relationshipId !== undefined
          && relationshipReferenceCount(
            LosslessXmlDocument.parse(updated),
            relationshipId,
          ) === 0
        ) {
          this.presentation.opcPackage.removeRelationship(this.partUri, relationshipId);
        }
        return;
      }

      if (
        current !== undefined
        && relationshipId !== undefined
        && shapeHyperlinkTargetsEqual(current, hyperlink)
      ) {
        replaceShapeHyperlinkElement(
          xml,
          element,
          hyperlink,
          relationshipId,
          this.partUri,
        );
        this.setXml(xml.serialize());
        return;
      }

      const relationshipInput: RelationshipInput = hyperlink.url !== undefined
        ? {
            type: HYPERLINK_RELATIONSHIP_TYPE,
            target: hyperlink.url,
            targetMode: 'External',
          }
        : {
            type: SLIDE_RELATIONSHIP_TYPE,
            target: relativeRelationshipTarget(this.partUri, targetSlide!.partUri),
            targetMode: 'Internal',
          };
      const canUpdateRelationship = current !== undefined
        && relationshipId !== undefined
        && relationshipReferenceCount(xml, relationshipId) === 1;
      const nextRelationshipId = canUpdateRelationship
        ? this.presentation.opcPackage.updateRelationship(
            this.partUri,
            relationshipId,
            relationshipInput,
          ).id
        : this.presentation.opcPackage.addRelationship(
            this.partUri,
            relationshipInput,
          ).id;
      replaceShapeHyperlinkElement(
        xml,
        element,
        hyperlink,
        nextRelationshipId,
        this.partUri,
      );
      const updated = xml.serialize();
      this.setXml(updated);
      if (
        relationshipId !== undefined
        && relationshipId !== nextRelationshipId
        && relationshipReferenceCount(
          LosslessXmlDocument.parse(updated),
          relationshipId,
        ) === 0
      ) {
        this.presentation.opcPackage.removeRelationship(this.partUri, relationshipId);
      }
    });
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

  getShapeTextMargins(id: number): TextBoxMargins | undefined {
    const { xml, element } = this.resolveShape(id);
    return readTextBoxMargins(xml, element, this.partUri);
  }

  setShapeTextMargins(id: number, value: TextBoxMarginInput | undefined): void {
    this.presentation.opcPackage.transaction(() => {
      const margins = normalizeTextBoxMargins(value, 'Text margins');
      const { xml, element } = this.resolveShape(id);
      replaceTextBoxMargins(xml, element, margins, this.partUri);
      this.setXml(xml.serialize());
    });
  }

  getShapeTextVerticalAlignment(id: number): TextBoxVerticalAlignment | undefined {
    const { xml, element } = this.resolveShape(id);
    return readTextBoxVerticalAlignment(xml, element, this.partUri);
  }

  setShapeTextVerticalAlignment(id: number, value: TextBoxVerticalAlignment | undefined): void {
    this.presentation.opcPackage.transaction(() => {
      const alignment = value === undefined
        ? undefined
        : normalizeTextBoxVerticalAlignment(value, 'Text vertical alignment');
      const { xml, element } = this.resolveShape(id);
      replaceTextBoxVerticalAlignment(xml, element, alignment, this.partUri);
      this.setXml(xml.serialize());
    });
  }

  getShapeTextWrap(id: number): boolean | undefined {
    const { xml, element } = this.resolveShape(id);
    return readTextBoxWrap(xml, element, this.partUri);
  }

  setShapeTextWrap(id: number, value: boolean | undefined): void {
    this.presentation.opcPackage.transaction(() => {
      const textWrap = value === undefined ? undefined : normalizeTextBoxWrap(value, 'Text wrap');
      const { xml, element } = this.resolveShape(id);
      replaceTextBoxWrap(xml, element, textWrap, this.partUri);
      this.setXml(xml.serialize());
    });
  }

  getShapeTextDirection(id: number): TextBoxTextDirection | undefined {
    const { xml, element } = this.resolveShape(id);
    return readTextBoxTextDirection(xml, element, this.partUri);
  }

  setShapeTextDirection(id: number, value: TextBoxTextDirection | undefined): void {
    this.presentation.opcPackage.transaction(() => {
      const direction = value === undefined
        ? undefined
        : normalizeTextBoxTextDirection(value, 'Text direction');
      const { xml, element } = this.resolveShape(id);
      replaceTextBoxTextDirection(xml, element, direction, this.partUri);
      this.setXml(xml.serialize());
    });
  }

  getShapeTextFit(id: number): TextBoxFit | undefined {
    const { xml, element } = this.resolveShape(id);
    return readTextBoxFit(xml, element, this.partUri);
  }

  setShapeTextFit(id: number, value: TextBoxFit | undefined): void {
    this.presentation.opcPackage.transaction(() => {
      const fit = value === undefined ? undefined : normalizeTextBoxFit(value, 'Text fit');
      const { xml, element } = this.resolveShape(id);
      replaceTextBoxFit(xml, element, fit, this.partUri);
      this.setXml(xml.serialize());
    });
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

  getImageSourceRectangle(id: number): Readonly<ImageSourceRectangle> | undefined {
    const { xml, element } = this.resolveShape(id);
    return readImageSourceRectangle(xml, element);
  }

  setImageSourceRectangle(id: number, value: ImageSourceRectangle | undefined): void {
    const normalized = value === undefined
      ? undefined
      : normalizeImageSourceRectangle(value, 'Image source rectangle');
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceImageSourceRectangle(xml, element, normalized, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
  }

  addImage(bytes: Uint8Array, options: AddImageOptions): ImageModel {
    const definition = normalizeEmbeddedRasterImage(bytes, options);
    return this.presentation.opcPackage.transaction(() => {
      const pkg = this.presentation.opcPackage;
      const mediaPartUri = pkg.allocatePartUri(
        '/ppt/media',
        'image',
        definition.extension,
      );
      pkg.setPart(mediaPartUri, definition.bytes, definition.contentType);
      const relationship = pkg.addRelationship(this.partUri, {
        type: IMAGE_RELATIONSHIP_TYPE,
        target: relativeRelationshipTarget(this.partUri, mediaPartUri),
        targetMode: 'Internal',
      });
      const { xml } = this.parse();
      const shapeTree = requirePresetShapeTree(xml, this.partUri);
      const nextId = allocatePresetShapeId(xml, shapeTree, this.partUri);
      const imageCount = this.shapes.filter(({ kind }) => kind === 'image').length;
      const pictureXml = renderEmbeddedRasterImageXml(
        nextId,
        definition,
        relationship.id,
        `Image ${imageCount}`,
      );
      const extensionList = directElementChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, pictureXml);
      else xml.appendChildXml(shapeTree, pictureXml);
      this.setXml(xml.serialize());
      const image = this.shapes.find(({ id }) => id === nextId);
      if (!(image instanceof ImageModel) || image.kind !== 'image') {
        throw new ModelParseError(`Created image ${nextId} could not be resolved`, this.partUri);
      }
      return image;
    });
  }

  addSvgImage(
    svgBytes: Uint8Array,
    fallbackPngBytes: Uint8Array,
    options: AddSvgImageOptions = {},
  ): ImageModel {
    const definition = normalizeEmbeddedSvgImage(svgBytes, fallbackPngBytes, options);
    return this.presentation.opcPackage.transaction(() => {
      const pkg = this.presentation.opcPackage;
      const fallbackPartUri = pkg.allocatePartUri('/ppt/media', 'image', '.png');
      const svgPartUri = pkg.allocatePartUri('/ppt/media', 'image', '.svg');
      pkg.setPart(fallbackPartUri, definition.fallbackPngBytes, 'image/png');
      pkg.setPart(svgPartUri, definition.svgBytes, 'image/svg+xml');
      const fallbackRelationship = pkg.addRelationship(this.partUri, {
        type: IMAGE_RELATIONSHIP_TYPE,
        target: relativeRelationshipTarget(this.partUri, fallbackPartUri),
        targetMode: 'Internal',
      });
      const svgRelationship = pkg.addRelationship(this.partUri, {
        type: IMAGE_RELATIONSHIP_TYPE,
        target: relativeRelationshipTarget(this.partUri, svgPartUri),
        targetMode: 'Internal',
      });
      const { xml } = this.parse();
      const shapeTree = requirePresetShapeTree(xml, this.partUri);
      const nextId = allocatePresetShapeId(xml, shapeTree, this.partUri);
      const imageCount = this.shapes.filter(({ kind }) => kind === 'image').length;
      const pictureXml = renderEmbeddedSvgImageXml(
        nextId,
        definition,
        fallbackRelationship.id,
        svgRelationship.id,
        `Image ${imageCount}`,
      );
      const extensionList = directElementChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, pictureXml);
      else xml.appendChildXml(shapeTree, pictureXml);
      this.setXml(xml.serialize());
      const image = this.shapes.find(({ id }) => id === nextId);
      if (!(image instanceof ImageModel) || image.kind !== 'image') {
        throw new ModelParseError(`Created image ${nextId} could not be resolved`, this.partUri);
      }
      return image;
    });
  }

  addShape(
    type: PresetShapeType,
    options: AddShapeOptions = {},
  ): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const normalized = normalizePresetShape(type, options);
      let targetSlide: SlideModel | undefined;
      if (normalized.hyperlink?.slide !== undefined) {
        targetSlide = this.presentation.slides[normalized.hyperlink.slide - 1];
        if (!targetSlide) {
          throw new RangeError(
            `Preset shape hyperlink slide ${normalized.hyperlink.slide} is out of range`,
          );
        }
      }
      const { xml } = this.parse();
      const shapeTree = requirePresetShapeTree(xml, this.partUri);
      const nextId = allocatePresetShapeId(xml, shapeTree, this.partUri);
      let hyperlinkRelationshipId: string | undefined;
      if (normalized.hyperlink?.url !== undefined) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: HYPERLINK_RELATIONSHIP_TYPE,
          target: normalized.hyperlink.url,
          targetMode: 'External',
        }).id;
      } else if (targetSlide) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: SLIDE_RELATIONSHIP_TYPE,
          target: relativeRelationshipTarget(this.partUri, targetSlide.partUri),
          targetMode: 'Internal',
        }).id;
      }
      const shapeXml = renderPresetShapeXml(nextId, normalized, hyperlinkRelationshipId);
      const extensionList = directChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, shapeXml);
      else xml.appendChildXml(shapeTree, shapeXml);
      this.setXml(xml.serialize());
      const shape = this.shapes.find((candidate) => candidate.id === nextId);
      if (!(shape instanceof ShapeModel) || shape.kind !== 'shape') {
        throw new ModelParseError(`Created preset shape ${nextId} could not be resolved`, this.partUri);
      }
      return shape;
    });
  }

  addCustomShape(
    geometry: CustomGeometry,
    options: AddCustomShapeOptions = {},
  ): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const normalized = normalizeCustomShape(geometry, options);
      let targetSlide: SlideModel | undefined;
      if (normalized.hyperlink?.slide !== undefined) {
        targetSlide = this.presentation.slides[normalized.hyperlink.slide - 1];
        if (!targetSlide) {
          throw new RangeError(
            `Custom shape hyperlink slide ${normalized.hyperlink.slide} is out of range`,
          );
        }
      }
      const { xml } = this.parse();
      const shapeTree = requirePresetShapeTree(xml, this.partUri);
      const nextId = allocatePresetShapeId(xml, shapeTree, this.partUri);
      let hyperlinkRelationshipId: string | undefined;
      if (normalized.hyperlink?.url !== undefined) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: HYPERLINK_RELATIONSHIP_TYPE,
          target: normalized.hyperlink.url,
          targetMode: 'External',
        }).id;
      } else if (targetSlide) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: SLIDE_RELATIONSHIP_TYPE,
          target: relativeRelationshipTarget(this.partUri, targetSlide.partUri),
          targetMode: 'Internal',
        }).id;
      }
      const shapeXml = renderCustomShapeXml(nextId, normalized, hyperlinkRelationshipId);
      const extensionList = directChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, shapeXml);
      else xml.appendChildXml(shapeTree, shapeXml);
      this.setXml(xml.serialize());
      const shape = this.shapes.find((candidate) => candidate.id === nextId);
      if (!(shape instanceof ShapeModel) || shape.kind !== 'shape') {
        throw new ModelParseError(`Created custom shape ${nextId} could not be resolved`, this.partUri);
      }
      return shape;
    });
  }

  addTable(
    rows: readonly (readonly AddTableCellInput[])[],
    options: AddTableOptions = {},
  ): TableModel {
    return this.presentation.opcPackage.transaction(() => {
      const definition = normalizeTableDefinition(rows, options);
      const { xml } = this.parse();
      const shapeTree = requireTableShapeTree(xml, this.partUri);
      const nextId = allocateShapeId(xml);
      const tableXml = renderTableGraphicFrame(nextId, definition);
      const extensionList = directChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, tableXml);
      else xml.appendChildXml(shapeTree, tableXml);
      this.setXml(xml.serialize());
      const table = this.shapes.find((candidate) => candidate.id === nextId);
      if (!(table instanceof TableModel) || table.kind !== 'table') {
        throw new ModelParseError(`Created table ${nextId} could not be resolved`, this.partUri);
      }
      return table;
    });
  }

  addChart(
    type: ChartType,
    series: readonly ChartSeriesInput[],
    options?: AddChartOptions,
  ): Promise<ChartModel>;
  addChart(
    groups: readonly ChartGroupInput[],
    options?: AddChartOptions,
  ): Promise<ChartModel>;
  async addChart(
    typeOrGroups: ChartType | readonly ChartGroupInput[],
    seriesOrOptions?: readonly ChartSeriesInput[] | AddChartOptions,
    options: AddChartOptions = {},
  ): Promise<ChartModel> {
    const definition = Array.isArray(typeOrGroups)
      ? normalizeChartDefinition({ groups: typeOrGroups })
      : normalizeChartDefinition({
          groups: [{
            type: typeOrGroups as ChartType,
            series: seriesOrOptions as readonly ChartSeriesInput[],
          }],
        });
    const normalizedOptions = normalizeAddChartOptions(
      Array.isArray(typeOrGroups) ? seriesOrOptions as AddChartOptions | undefined : options,
    );
    const workbookBytes = await buildChartWorkbook(definition);
    const plan = planChartWorkbook(definition);
    return this.presentation.opcPackage.transaction(() => {
      const pkg = this.presentation.opcPackage;
      const chartPartUri = pkg.allocatePartUri('/ppt/charts', 'chart', '.xml');
      const workbookPartUri = pkg.allocatePartUri(
        '/ppt/embeddings',
        'Microsoft_Excel_Worksheet',
        '.xlsx',
      );
      pkg.setPart(workbookPartUri, workbookBytes, WORKBOOK_CONTENT_TYPE);
      const workbookRelationshipId = pkg.allocateRelationshipId(chartPartUri);
      pkg.setPart(
        chartPartUri,
        renderChartPart(definition, plan.formulas, workbookRelationshipId),
        CHART_CONTENT_TYPE,
      );
      pkg.addRelationship(chartPartUri, {
        id: workbookRelationshipId,
        type: PACKAGE_RELATIONSHIP_TYPE,
        target: relativeRelationshipTarget(chartPartUri, workbookPartUri),
        targetMode: 'Internal',
      });
      const chartRelationship = pkg.addRelationship(this.partUri, {
        type: CHART_RELATIONSHIP_TYPE,
        target: relativeRelationshipTarget(this.partUri, chartPartUri),
        targetMode: 'Internal',
      });
      const { xml } = this.parse();
      const shapeTree = requirePresetShapeTree(xml, this.partUri);
      const nextId = allocatePresetShapeId(xml, shapeTree, this.partUri);
      const frame = renderChartGraphicFrame(nextId, chartRelationship.id, normalizedOptions);
      const extensionList = directElementChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, frame);
      else xml.appendChildXml(shapeTree, frame);
      this.setXml(xml.serialize());
      const chart = this.shapes.find((candidate) => candidate.id === nextId);
      if (!(chart instanceof ChartModel) || chart.kind !== 'chart') {
        throw new ModelParseError(`Created chart ${nextId} could not be resolved`, this.partUri);
      }
      return chart;
    });
  }

  replaceChartDefinition(
    shapeId: number,
    current: Readonly<ChartState>,
    next: Readonly<ChartDefinition>,
    workbookBytes: Uint8Array | undefined,
  ): void {
    replaceChartDefinition(
      this.presentation.opcPackage,
      this,
      shapeId,
      current,
      next,
      workbookBytes,
    );
  }

  deleteChart(shapeId: number): void {
    this.presentation.opcPackage.transaction(() => {
      const pkg = this.presentation.opcPackage;
      const { xml, element } = this.resolveShape(shapeId);
      if (element.localName !== 'graphicFrame') {
        throw new ModelParseError(`Chart shape ${shapeId} was not found`, this.partUri);
      }
      const references = xml.descendants(element, 'chart');
      if (references.length !== 1) {
        throw new ModelParseError(
          `Chart shape ${shapeId} must contain exactly one chart reference`,
          this.partUri,
        );
      }
      const ids = references[0]!.attributes.filter((attribute) =>
        attribute.name.slice(attribute.name.lastIndexOf(':') + 1) === 'id'
        && (
          attribute.name === 'r:id'
          || relationshipAttributeNamespace(references[0]!, attribute.name)
            === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        ));
      if (ids.length !== 1) {
        throw new ModelParseError(`Chart shape ${shapeId} has an invalid relationship id`, this.partUri);
      }
      const relationshipId = ids[0]!.value;
      const relationships = this.relationships.filter(({ id }) => id === relationshipId);
      const relationship = relationships.length === 1 ? relationships[0] : undefined;
      if (
        !relationship
        || (
          relationship.type !== CHART_RELATIONSHIP_TYPE
          && !relationship.type.endsWith('/chartEx')
        )
        || relationship.targetMode !== 'Internal'
        || !relationship.resolvedTarget
      ) {
        throw new ModelParseError(
          `Chart shape ${shapeId} has an invalid chart relationship`,
          this.partUri,
        );
      }
      const chartPartUri = relationship.resolvedTarget;
      const removeRelationship = chartRelationshipReferenceCount(xml, relationshipId) === 1;
      xml.removeElement(element);
      this.setXml(xml.serialize());
      if (removeRelationship) pkg.removeRelationship(this.partUri, relationshipId);
      garbageCollectOwnedDependencies(pkg, [chartPartUri]);
    });
  }

  addText(value: string, options: AddTextOptions = {}): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const normalized = validateTextInput(value, options);
      const bullet = normalized.bullet === false ? undefined : normalized.bullet;
      const spacing = resolveParagraphSpacing(normalized.spacing);
      const paragraphs = normalized.value
        .split('\n')
        .map((line) => textParagraphXml(
          line,
          'a:',
          options.align,
          normalized.rtl,
          bullet,
          spacing,
          normalized.level,
          normalized.tabStops,
          normalized.language,
          normalized.marginLeft,
          normalized.marginRight,
          normalized.indent,
        ))
        .join('');
      return this.addTextShape(
        paragraphs,
        options,
        normalized.margin,
        normalized.verticalAlignment,
        normalized.textDirection,
        normalized.textFit,
        normalized.textWrap,
      );
    });
  }

  addRichText(value: readonly RichTextParagraph[], options: AddTextOptions = {}): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const paragraphs = normalizeRichText(value);
      const defaults = validateAddTextOptions(options);
      return this.addTextShape(
        renderRichTextParagraphs(paragraphs, {
          ...(defaults.language !== undefined ? { defaultLanguage: defaults.language } : {}),
          ...(options.align ? { defaultAlign: options.align } : {}),
          ...(defaults.rtl !== undefined ? { defaultRtl: defaults.rtl } : {}),
          ...(defaults.bullet !== undefined ? { defaultBullet: defaults.bullet } : {}),
          ...(defaults.indent !== undefined ? { defaultIndent: defaults.indent } : {}),
          ...(defaults.level !== undefined ? { defaultLevel: defaults.level } : {}),
          ...(defaults.marginLeft !== undefined ? { defaultMarginLeft: defaults.marginLeft } : {}),
          ...(defaults.marginRight !== undefined ? { defaultMarginRight: defaults.marginRight } : {}),
          ...(defaults.spacing !== undefined ? { defaultSpacing: defaults.spacing } : {}),
          ...(defaults.tabStops !== undefined ? { defaultTabStops: defaults.tabStops } : {}),
        }),
        options,
        defaults.margin,
        defaults.verticalAlignment,
        defaults.textDirection,
        defaults.textFit,
        defaults.textWrap,
      );
    });
  }

  setXml(xml: string): void {
    this.presentation.setXmlPart(this.partUri, xml);
  }

  private addTextShape(
    paragraphs: string,
    options: AddTextOptions,
    margins: TextBoxMargins | undefined,
    verticalAlignment: TextBoxVerticalAlignment,
    textDirection: TextBoxTextDirection | undefined,
    textFit: TextBoxFit | undefined,
    textWrap: boolean,
  ): ShapeModel {
    const { xml } = this.parse();
    const shapeTree = xml
      .elements('spTree')
      .find(({ parent }) => parent?.localName === 'cSld');
    if (!shapeTree) throw new ModelParseError('Slide does not contain a shape tree', this.partUri);
    const nextId = allocateShapeId(xml);
    const shapeXml = textShapeXml(
      nextId,
      paragraphs,
      options,
      margins,
      verticalAlignment,
      textDirection,
      textFit,
      textWrap,
    );
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

function shapeHyperlinkTargetsEqual(
  left: NormalizedHyperlink,
  right: NormalizedHyperlink,
): boolean {
  if (left.url !== undefined || right.url !== undefined) {
    return left.url !== undefined
      && right.url !== undefined
      && left.url === right.url;
  }
  return left.slide === right.slide;
}

export function findTitleShape(xml: LosslessXmlDocument): XmlElement | undefined {
  const shapes = xml.elements('sp');
  return (
    shapes.find((shape) =>
      xml.descendants(shape, 'ph').some((placeholder) => {
        const type = xml.attribute(placeholder, 'type')?.value;
        return type === 'title' || type === 'ctrTitle';
      }),
    ) ?? shapes.find((shape) =>
      xml.descendants(shape, 't').length > 0
      && !xml.descendants(shape, 'ph').some(
        (placeholder) => xml.attribute(placeholder, 'type')?.value === 'sldNum',
      ),
    )
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
  readonly indent: number | undefined;
  readonly language: string | undefined;
  readonly level: number | undefined;
  readonly margin: TextBoxMargins | undefined;
  readonly marginLeft: number | undefined;
  readonly marginRight: number | undefined;
  readonly rtl: boolean | undefined;
  readonly spacing: NormalizedParagraphSpacingUpdate | undefined;
  readonly tabStops: readonly NormalizedParagraphTabStop[] | undefined;
  readonly verticalAlignment: TextBoxVerticalAlignment;
  readonly textDirection: TextBoxTextDirection | undefined;
  readonly textFit: TextBoxFit | undefined;
  readonly textWrap: boolean;
}

function validateTextInput(value: string, options: AddTextOptions): NormalizedTextInput {
  const normalized = validatePlainText(value);
  const defaults = validateAddTextOptions(options);
  if (defaults.bullet && defaults.marginLeft !== undefined) {
    throw new TypeError('Paragraph left margin cannot be combined with an active bullet');
  }
  if (defaults.bullet && defaults.indent !== undefined) {
    throw new TypeError('Paragraph indent cannot be combined with an active bullet');
  }
  return {
    value: normalized,
    bullet: defaults.bullet,
    indent: defaults.indent,
    language: defaults.language,
    level: defaults.level,
    margin: defaults.margin,
    marginLeft: defaults.marginLeft,
    marginRight: defaults.marginRight,
    rtl: defaults.rtl,
    spacing: defaults.spacing,
    tabStops: defaults.tabStops,
    verticalAlignment: defaults.verticalAlignment,
    textDirection: defaults.textDirection,
    textFit: defaults.textFit,
    textWrap: defaults.textWrap,
  };
}

interface NormalizedAddTextOptions {
  readonly bullet?: NormalizedParagraphBullet | false;
  readonly indent?: number;
  readonly language?: string;
  readonly level?: number;
  readonly margin?: TextBoxMargins;
  readonly marginLeft?: number;
  readonly marginRight?: number;
  readonly rtl?: boolean;
  readonly spacing?: NormalizedParagraphSpacingUpdate;
  readonly tabStops?: readonly NormalizedParagraphTabStop[];
  readonly verticalAlignment: TextBoxVerticalAlignment;
  readonly textDirection?: TextBoxTextDirection;
  readonly textFit?: TextBoxFit;
  readonly textWrap: boolean;
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
  const language = options.lang === undefined
    ? undefined
    : normalizeTextLanguage(options.lang, 'Text language');
  const indent = options.paragraphIndent === undefined
    ? undefined
    : normalizeParagraphIndent(options.paragraphIndent, 'Paragraph indent');
  const margin = options.margin === undefined
    ? undefined
    : normalizeTextBoxMargins(options.margin, 'Text margin');
  const marginLeft = options.paragraphMarginLeft === undefined
    ? undefined
    : normalizeParagraphMarginLeft(options.paragraphMarginLeft, 'Paragraph left margin');
  const marginRight = options.paragraphMarginRight === undefined
    ? undefined
    : normalizeParagraphMarginRight(options.paragraphMarginRight, 'Paragraph right margin');
  const rtl = options.rtlMode === undefined
    ? undefined
    : normalizeParagraphRtl(options.rtlMode, 'Text RTL mode');
  const spacing = options.spacing === undefined
    ? undefined
    : normalizeParagraphSpacing(options.spacing, 'Text spacing');
  const tabStops = options.tabStops === undefined
    ? undefined
    : normalizeParagraphTabStops(options.tabStops, 'Text tabStops');
  const verticalAlignment = options.valign === undefined
    ? 'middle'
    : normalizeTextBoxVerticalAlignment(options.valign, 'Text vertical alignment');
  const textDirection = options.vert === undefined
    ? undefined
    : normalizeTextBoxTextDirection(options.vert, 'Text direction');
  const textFit = options.fit === undefined
    ? undefined
    : normalizeTextBoxFit(options.fit, 'Text fit');
  const textWrap = options.wrap === undefined
    ? true
    : normalizeTextBoxWrap(options.wrap, 'Text wrap');
  return {
    ...(bullet !== undefined ? { bullet } : {}),
    ...(indent !== undefined ? { indent } : {}),
    ...(language !== undefined ? { language } : {}),
    ...(level !== undefined ? { level } : {}),
    ...(margin !== undefined ? { margin } : {}),
    ...(marginLeft !== undefined ? { marginLeft } : {}),
    ...(marginRight !== undefined ? { marginRight } : {}),
    ...(rtl !== undefined ? { rtl } : {}),
    ...(spacing !== undefined ? { spacing } : {}),
    ...(tabStops !== undefined ? { tabStops } : {}),
    verticalAlignment,
    ...(textDirection !== undefined ? { textDirection } : {}),
    ...(textFit !== undefined ? { textFit } : {}),
    textWrap,
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

function allocatePresetShapeId(
  xml: LosslessXmlDocument,
  shapeTree: XmlElement,
  partUri: string,
): number {
  const identifiers = new Set<number>();
  let maximum = 0;
  for (const properties of xml.descendants(shapeTree, 'cNvPr')) {
    if (namespaceUri(properties) !== PRESENTATION_NAMESPACE) {
      throw new ModelParseError('Slide shape tree contains an unsafe non-visual property', partUri);
    }
    const attributes = properties.attributes.filter(
      ({ name, localName }) => localName === 'id' && !name.startsWith('xmlns:'),
    );
    const attribute = attributes[0];
    if (
      attributes.length !== 1
      || !attribute
      || attribute.name !== 'id'
      || !/^\d+$/.test(attribute.value)
    ) {
      throw new ModelParseError('Slide shape tree contains an invalid shape id', partUri);
    }
    const id = Number(attribute.value);
    if (!Number.isSafeInteger(id) || id > 4_294_967_295) {
      throw new ModelParseError('Slide shape tree contains an unsafe shape id', partUri);
    }
    if (identifiers.has(id)) {
      throw new ModelParseError(`Slide shape tree contains duplicate shape id ${id}`, partUri);
    }
    identifiers.add(id);
    maximum = Math.max(maximum, id);
  }
  if (maximum >= 4_294_967_295) {
    throw new ModelParseError('Slide shape ids are exhausted', partUri);
  }
  return maximum + 1;
}

function requirePresetShapeTree(
  xml: LosslessXmlDocument,
  partUri: string,
): XmlElement {
  const root = xml.roots.length === 1 ? xml.roots[0] : undefined;
  if (!root || root.localName !== 'sld' || namespaceUri(root) !== PRESENTATION_NAMESPACE) {
    throw new ModelParseError('Slide does not have a safe presentation root', partUri);
  }
  const commonSlideData = directElementChildren(root, 'cSld');
  if (
    commonSlideData.length !== 1
    || namespaceUri(commonSlideData[0]!) !== PRESENTATION_NAMESPACE
  ) {
    throw new ModelParseError('Slide must contain exactly one direct common slide data element', partUri);
  }
  const shapeTrees = directElementChildren(commonSlideData[0]!, 'spTree');
  if (shapeTrees.length !== 1 || namespaceUri(shapeTrees[0]!) !== PRESENTATION_NAMESPACE) {
    throw new ModelParseError('Slide must contain exactly one direct shape tree', partUri);
  }
  const extensionLists = directElementChildren(shapeTrees[0]!, 'extLst');
  if (
    extensionLists.length > 1
    || (extensionLists[0] && namespaceUri(extensionLists[0]) !== PRESENTATION_NAMESPACE)
  ) {
    throw new ModelParseError('Slide shape tree contains an unsafe extension list', partUri);
  }
  return shapeTrees[0]!;
}

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';

function namespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const declarations = current.attributes.filter(({ name }) => name === declarationName);
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function relationshipAttributeNamespace(
  element: XmlElement,
  attributeName: string,
): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(attributeName));
}

function chartRelationshipReferenceCount(
  xml: LosslessXmlDocument,
  relationshipId: string,
): number {
  return xml.elements().reduce((count, element) => count + element.attributes.filter((attribute) =>
    attribute.value === relationshipId
    && attribute.name.slice(attribute.name.lastIndexOf(':') + 1) === 'id'
    && (
      attribute.name === 'r:id'
      || relationshipAttributeNamespace(element, attribute.name)
        === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    )).length, 0);
}

function directElementChildren(
  element: XmlElement,
  localName: string,
): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function requireTableShapeTree(
  xml: LosslessXmlDocument,
  partUri: string,
): XmlElement {
  const candidates = xml.elements('spTree').filter((tree) =>
    tree.parent?.localName === 'cSld'
    && tree.parent.parent?.localName === 'sld');
  if (candidates.length !== 1) {
    throw new ModelParseError('Slide must contain exactly one direct shape tree', partUri);
  }
  const extensionLists = directChildren(candidates[0]!, 'extLst');
  if (extensionLists.length > 1) {
    throw new ModelParseError('Slide shape tree contains repeated extension lists', partUri);
  }
  return candidates[0]!;
}

function textShapeXml(
  id: number,
  paragraphs: string,
  options: AddTextOptions,
  margins: TextBoxMargins | undefined,
  verticalAlignment: TextBoxVerticalAlignment,
  textDirection: TextBoxTextDirection | undefined,
  textFit: TextBoxFit | undefined,
  textWrap: boolean,
): string {
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
  const wrapAttribute = renderTextBoxWrapAttribute(textWrap);
  const marginAttributes = renderTextBoxMarginAttributes(margins);
  const verticalAlignmentAttribute = renderTextBoxVerticalAlignmentAttribute(verticalAlignment);
  const textDirectionAttribute = textDirection === undefined
    ? ''
    : renderTextBoxTextDirectionAttribute(textDirection);
  const fitChild = textFit === undefined ? '' : renderTextBoxFitChild(textFit);
  const bodyProperties = fitChild === ''
    ? `<a:bodyPr${wrapAttribute}${marginAttributes} rtlCol="0"${verticalAlignmentAttribute}${textDirectionAttribute}/>`
    : `<a:bodyPr${wrapAttribute}${marginAttributes} rtlCol="0"${verticalAlignmentAttribute}${textDirectionAttribute}>${fitChild}</a:bodyPr>`;
  return `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${transformAttributes}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody>${bodyProperties}<a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

function textParagraphXml(
  value: string,
  prefix = 'a:',
  align?: TextAlignment,
  rtl?: boolean,
  bullet?: NormalizedParagraphBullet,
  spacing?: NormalizedParagraphSpacing,
  level?: number,
  tabStops?: readonly NormalizedParagraphTabStop[],
  language?: string,
  marginLeft?: number,
  marginRight?: number,
  indent?: number,
): string {
  const properties = renderParagraphProperties(
    undefined,
    prefix,
    align,
    rtl,
    bullet,
    spacing,
    level,
    tabStops,
    marginLeft,
    marginRight,
    indent,
  );
  const languageValue = escapeXmlAttribute(language ?? 'en-US');
  const endProperties = `<${prefix}endParaRPr lang="${languageValue}" dirty="0"/>`;
  if (value.length === 0) return `<${prefix}p>${properties}${endProperties}</${prefix}p>`;
  return `<${prefix}p>${properties}${defaultTextRunXml(value, prefix, language)}${endProperties}</${prefix}p>`;
}

function defaultTextRunXml(value: string, prefix = 'a:', language?: string): string {
  const languageValue = escapeXmlAttribute(language ?? 'en-US');
  const alternateLanguage = language === undefined ? '' : ' altLang="en-US"';
  return `<${prefix}r><${prefix}rPr lang="${languageValue}"${alternateLanguage} dirty="0"><${prefix}solidFill><${prefix}schemeClr val="tx1"/></${prefix}solidFill><${prefix}latin typeface="+mn-lt"/></${prefix}rPr><${prefix}t xml:space="preserve">${escapeXmlText(value)}</${prefix}t></${prefix}r>`;
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
