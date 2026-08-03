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
import {
  commitPreparedChart,
  prepareChartCreation,
} from './chart-create.internal.js';
import { replaceChartDefinition } from './chart-edit.internal.js';
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
import type {
  AddPlaceholderOptions,
  PlaceholderIdentity,
  PlaceholderSelector,
} from './placeholder.js';
import {
  normalizePlaceholderIdentity,
  readShapePlaceholder,
  resolvePlaceholderOwner,
  type ResolvedPlaceholderOwner,
} from './placeholder.internal.js';
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
  readRichTextState,
  renderColorChoice,
  renderParagraphProperties,
  renderRichTextParagraphs,
  resolveParagraphSpacing,
  replaceRichText,
  richTextParagraphsEqual,
  type NormalizedParagraphBullet,
  type NormalizedParagraphSpacing,
  type NormalizedParagraphSpacingUpdate,
  type NormalizedParagraphTabStop,
  type RichTextRunHyperlinkRelationshipIds,
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
  distributeTableDimension,
  normalizeTableDefinition,
  renderTableGraphicFrame,
  type NormalizedTableDefinition,
  type TableCellHyperlinkRelationshipIds,
  type TableCellRichTextRunHyperlinkRelationshipIds,
} from './table-create.internal.js';
import { planTableAutoPages } from './table-auto-page.internal.js';
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
  type BackgroundOwnerKind,
} from './slide-background.internal.js';
import type { SlideBackground } from './slide-background.js';
import {
  normalizeCustomShape,
  normalizePresetShape,
  normalizePresetShapeType,
  readPresetShapeType,
  renderCustomShapeXml,
  renderPresetShapeGeometry,
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
  renderShapeArrows,
  replaceShapeArrows,
  type NormalizedShapeArrows,
} from './shape-arrows.internal.js';
import {
  readShapeFill,
  replaceShapeFill,
} from './shape-fill.internal.js';
import { normalizeSimpleFill, renderSimpleFill } from './simple-fill.internal.js';
import {
  readShapeLine,
  replaceShapeLine,
} from './shape-line.internal.js';
import {
  normalizeSimpleLine,
  renderSimpleLine,
  type NormalizedSimpleLine,
} from './simple-line.internal.js';
import {
  readShapeShadow,
  replaceShapeShadow,
} from './shape-shadow.internal.js';
import {
  normalizeShapeShadow,
  renderSimpleShadow,
  type NormalizedShapeShadow,
} from './simple-shadow.internal.js';
import {
  drawingHyperlinkRelationshipIds,
  HYPERLINK_RELATIONSHIP_TYPE,
  normalizeHyperlink,
  readShapeHyperlink,
  relationshipReferenceCount,
  renderShapeHyperlink,
  replaceShapeHyperlinkElement,
  replaceTextRunHyperlinkElement,
  requireShapeHyperlinkRelationshipId,
  shapeHyperlinksEqual,
  SLIDE_RELATIONSHIP_TYPE,
  type NormalizedHyperlink,
} from './shape-hyperlink.internal.js';
import { requireEditableTableCellHyperlinkState } from './table-cell-hyperlink.internal.js';
import { requireEditableTableCellRichTextState } from './table-cell-rich-text.internal.js';
import { readDirectTablePhysicalCellMatrix } from './table-physical-cells.internal.js';
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
import {
  readTextShapeIsTextBox,
  replaceTextShapeIsTextBox,
} from './text-shape-is-text-box.internal.js';
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
import {
  EMU_PER_INCH,
  inches,
  points,
  type Emu,
  type Transform,
} from './units.js';

const IMAGE_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const CHART_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface AddTextOptions extends Partial<Transform> {
  readonly name?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly align?: TextAlignment;
  readonly arrows?: ShapeArrows;
  readonly bullet?: ParagraphBullet;
  readonly fill?: ShapeFill;
  readonly hyperlink?: Hyperlink;
  readonly isTextBox?: boolean;
  readonly line?: ShapeLine;
  readonly shape?: PresetShapeType;
  readonly rectRadius?: Emu;
  readonly shadow?: ShapeShadow;
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
  readonly autoPage?: boolean;
  readonly autoPageRepeatHeader?: boolean;
  readonly autoPageHeaderRows?: number;
  readonly autoPageSlideStartY?: number;
  readonly slideMargin?: TableAutoPageMarginInput;
  readonly align?: TextAlignment;
  readonly bold?: boolean;
  readonly color?: RichTextColor;
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly name?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly columnWidths?: number | readonly number[];
  readonly rowHeights?: number | readonly number[];
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
  readonly spacing?: ParagraphSpacing;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}

export type TableAutoPageMarginInput =
  | number
  | readonly [number, number, number, number];

export interface AddTableCellOptions {
  readonly align?: TextAlignment;
  readonly bold?: boolean;
  readonly border?: TableCellBorderInput;
  readonly color?: RichTextColor;
  readonly colspan?: number;
  readonly fill?: TableCellFill;
  readonly fit?: TextBoxFit;
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly hyperlink?: Hyperlink;
  readonly margin?: TextBoxMarginInput;
  readonly rowspan?: number;
  readonly spacing?: ParagraphSpacing;
  readonly textDirection?: TableCellTextDirection;
  readonly valign?: TextBoxVerticalAlignment;
}

export interface AddTableCell {
  readonly text: string | readonly RichTextParagraph[];
  readonly options?: AddTableCellOptions;
}

export type AddTableCellInput = string | AddTableCell;

interface PreparedRichTextRunHyperlink {
  readonly paragraphIndex: number;
  readonly runIndex: number;
  readonly relationship: RelationshipInput;
}

interface PreparedTableCellHyperlink {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly paragraphIndex?: number;
  readonly runIndex?: number;
  readonly target:
    | { readonly kind: 'external'; readonly url: string }
    | { readonly kind: 'slide'; readonly partUri: string };
}

interface CreatedTableCellHyperlinkRelationships {
  readonly defaultRelationshipIds: TableCellHyperlinkRelationshipIds;
  readonly runRelationshipIds: TableCellRichTextRunHyperlinkRelationshipIds;
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
  readonly #staleShapeModels = new WeakSet<SemanticShape>();
  #newAutoPagedSlidePartUris: readonly string[] = Object.freeze([]);
  #relationshipId: string;
  #slideId: number;

  constructor(
    readonly presentation: PresentationModel,
    readonly partUri: string,
    relationshipId: string,
    slideId: number,
    private readonly backgroundOwnerKind: BackgroundOwnerKind = 'slide',
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
    return readSlideBackground(
      this.presentation.opcPackage,
      this.partUri,
      this.backgroundOwnerKind,
    );
  }

  set background(value: SlideBackground | undefined) {
    replaceSlideBackground(
      this.presentation.opcPackage,
      this.partUri,
      value,
      this.backgroundOwnerKind,
    );
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

  get newAutoPagedSlides(): readonly SlideModel[] {
    const attached = new Map(this.presentation.slides.map((slide) => [slide.partUri, slide]));
    return Object.freeze(this.#newAutoPagedSlidePartUris.flatMap((partUri) => {
      const slide = attached.get(partUri);
      return slide === undefined ? [] : [slide];
    }));
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

  get placeholders(): readonly SemanticShape[] {
    return this.shapes.filter(({ placeholder }) => placeholder !== undefined);
  }

  get media(): readonly MediaModel[] {
    return this.shapes.filter((shape): shape is MediaModel => shape instanceof MediaModel);
  }

  async addAudio(source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const descriptor = await this.mediaCodec()
      .addAudio(this.partUri, source, options);
    if (descriptor.placeholder) this.invalidateShapeModel(descriptor.shapeId);
    return this.requireMedia(descriptor.shapeId);
  }

  async addVideo(source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const descriptor = await this.mediaCodec()
      .addVideo(this.partUri, source, options);
    if (descriptor.placeholder) this.invalidateShapeModel(descriptor.shapeId);
    return this.requireMedia(descriptor.shapeId);
  }

  private mediaCodec(): MediaCodec {
    return new MediaCodec(
      this.presentation.opcPackage,
      (slidePartUri, selector) => {
        const owner = resolvePlaceholderOwner(
          this.presentation.opcPackage,
          slidePartUri,
          selector,
          'media',
        );
        return {
          identity: owner.identity,
          name: owner.name,
          shapeId: owner.shapeId,
          x: owner.transform.x,
          y: owner.transform.y,
          width: owner.transform.width,
          height: owner.transform.height,
          rotation: owner.transform.rotation,
          flipHorizontal: owner.transform.flipHorizontal,
          flipVertical: owner.transform.flipVertical,
          start: owner.slideElement.start,
          end: owner.slideElement.end,
        };
      },
    );
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

  resolveShape(
    id: number,
    handle?: SemanticShape,
  ): { xml: LosslessXmlDocument; element: XmlElement } {
    if (handle && this.#staleShapeModels.has(handle)) {
      throw new ModelParseError(`Shape ${id} handle is stale`, this.partUri);
    }
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

  getShapeIsTextBox(id: number): boolean | undefined {
    const { xml, element } = this.resolveShape(id);
    return readTextShapeIsTextBox(xml, element);
  }

  setShapeIsTextBox(id: number, value: boolean): void {
    if (typeof value !== 'boolean') {
      throw new TypeError('Shape isTextBox must be a boolean');
    }
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      if (replaceTextShapeIsTextBox(xml, element, value, this.partUri)) {
        this.setXml(xml.serialize());
      }
    });
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
    const type = normalizePresetShapeType(value, 'Shape preset type');
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

  setTableCellHyperlink(
    id: number,
    rowIndex: number,
    columnIndex: number,
    value: Hyperlink | undefined,
  ): void {
    const hyperlink = value === undefined
      ? undefined
      : normalizeHyperlink(value, 'Table cell hyperlink');
    let targetSlide: SlideModel | undefined;
    if (hyperlink?.slide !== undefined) {
      targetSlide = this.presentation.slides[hyperlink.slide - 1];
      if (!targetSlide) {
        throw new RangeError(
          `Table cell ${rowIndex},${columnIndex} hyperlink slide ` +
          `${hyperlink.slide} is out of range`,
        );
      }
    }

    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      const matrix = readDirectTablePhysicalCellMatrix(element);
      if (!matrix) {
        throw new ModelParseError('Table cell hyperlink state is not safely editable', this.partUri);
      }
      const cell = matrix[rowIndex]?.[columnIndex];
      if (!cell) {
        throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      }
      const current = requireEditableTableCellHyperlinkState(
        cell,
        {
          relationships: this.relationships,
          slidePartUris: this.presentation.slides.map(({ partUri }) => partUri),
        },
        this.partUri,
      );
      if (shapeHyperlinksEqual(current.hyperlink, hyperlink)) return;

      if (hyperlink === undefined) {
        replaceTextRunHyperlinkElement(
          xml,
          current.properties,
          undefined,
          undefined,
          this.partUri,
        );
        const updated = xml.serialize();
        this.setXml(updated);
        if (
          current.relationshipId !== undefined
          && relationshipReferenceCount(
            LosslessXmlDocument.parse(updated),
            current.relationshipId,
          ) === 0
        ) {
          this.presentation.opcPackage.removeRelationship(
            this.partUri,
            current.relationshipId,
          );
        }
        return;
      }

      if (
        current.hyperlink !== undefined
        && current.relationshipId !== undefined
        && shapeHyperlinkTargetsEqual(current.hyperlink, hyperlink)
      ) {
        replaceTextRunHyperlinkElement(
          xml,
          current.properties,
          hyperlink,
          current.relationshipId,
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
      const canUpdateRelationship = current.relationshipId !== undefined
        && relationshipReferenceCount(xml, current.relationshipId) === 1;
      const nextRelationshipId = canUpdateRelationship
        ? this.presentation.opcPackage.updateRelationship(
            this.partUri,
            current.relationshipId!,
            relationshipInput,
          ).id
        : this.presentation.opcPackage.addRelationship(
            this.partUri,
            relationshipInput,
          ).id;
      replaceTextRunHyperlinkElement(
        xml,
        current.properties,
        hyperlink,
        nextRelationshipId,
        this.partUri,
      );
      const updated = xml.serialize();
      this.setXml(updated);
      if (
        current.relationshipId !== undefined
        && current.relationshipId !== nextRelationshipId
        && relationshipReferenceCount(
          LosslessXmlDocument.parse(updated),
          current.relationshipId,
        ) === 0
      ) {
        this.presentation.opcPackage.removeRelationship(
          this.partUri,
          current.relationshipId,
        );
      }
    });
  }

  setTableCellRichText(
    id: number,
    rowIndex: number,
    columnIndex: number,
    value: readonly RichTextParagraph[],
  ): void {
    const paragraphs = normalizeRichText(value);
    const preparedRunHyperlinks = this.prepareRichTextRunHyperlinks(paragraphs);
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      const matrix = readDirectTablePhysicalCellMatrix(element);
      if (!matrix) {
        throw new ModelParseError(
          'Table cell rich text state is not safely editable',
          this.partUri,
        );
      }
      const cell = matrix[rowIndex]?.[columnIndex];
      if (!cell) {
        throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      }
      const readContext = {
        relationships: this.relationships,
        slidePartUris: this.presentation.slides.map(({ partUri }) => partUri),
      };
      const current = requireEditableTableCellRichTextState(
        xml,
        cell,
        readContext,
        this.partUri,
      );
      const normalizedCurrent = current.paragraphs.length === 0
        ? current.paragraphs
        : normalizeRichText(current.paragraphs);
      if (richTextParagraphsEqual(normalizedCurrent, paragraphs)) return;

      const textBody = cell.children.find(
        (child): child is XmlElement =>
          child.type === 'element' && child.localName === 'txBody',
      )!;
      const previousRelationshipIds = drawingHyperlinkRelationshipIds(textBody);
      const preparedByPosition = new Map(preparedRunHyperlinks.map((prepared) => [
        `${prepared.paragraphIndex}:${prepared.runIndex}`,
        prepared,
      ] as const));
      const relationshipIds = paragraphs.map(({ runs }) =>
        runs.map(() => undefined as string | undefined));

      for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
        for (const [runIndex, run] of paragraph.runs.entries()) {
          const hyperlink = run.style?.hyperlink;
          if (hyperlink === undefined || hyperlink === false) continue;
          const previous = current.runHyperlinkBindings[paragraphIndex]?.[runIndex];
          if (previous && shapeHyperlinkTargetsEqual(previous.hyperlink, hyperlink)) {
            relationshipIds[paragraphIndex]![runIndex] = previous.relationshipId;
            continue;
          }
          const prepared = preparedByPosition.get(`${paragraphIndex}:${runIndex}`);
          if (!prepared) throw new Error('Prepared table-cell rich text hyperlink was not found');
          const relationshipId = previous
            && relationshipReferenceCount(xml, previous.relationshipId) === 1
            ? this.presentation.opcPackage.updateRelationship(
                this.partUri,
                previous.relationshipId,
                prepared.relationship,
              ).id
            : this.presentation.opcPackage.addRelationship(
                this.partUri,
                prepared.relationship,
              ).id;
          relationshipIds[paragraphIndex]![runIndex] = relationshipId;
        }
      }

      const updated = replaceRichText(
        xml,
        cell,
        paragraphs,
        this.partUri,
        relationshipIds,
      );
      this.setXml(updated);
      const updatedXml = LosslessXmlDocument.parse(updated);
      for (const relationshipId of previousRelationshipIds) {
        if (
          relationshipReferenceCount(updatedXml, relationshipId) === 0
          && this.relationships.some(({ id: candidate }) => candidate === relationshipId)
        ) {
          this.presentation.opcPackage.removeRelationship(this.partUri, relationshipId);
        }
      }
    });
  }

  setShapeRichText(id: number, value: readonly RichTextParagraph[]): void {
    const paragraphs = normalizeRichText(value);
    const preparedRunHyperlinks = this.prepareRichTextRunHyperlinks(paragraphs);
    this.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolveShape(id);
      const readContext = {
        relationships: this.relationships,
        slidePartUris: this.presentation.slides.map(({ partUri }) => partUri),
      };
      const current = readRichTextState(xml, element, readContext);
      const normalizedCurrent = current.paragraphs.length === 0
        ? current.paragraphs
        : normalizeRichText(current.paragraphs);
      if (richTextParagraphsEqual(normalizedCurrent, paragraphs)) return;

      const textBody = element.children.find(
        (child): child is XmlElement => child.type === 'element' && child.localName === 'txBody',
      );
      const previousRelationshipIds = textBody
        ? drawingHyperlinkRelationshipIds(textBody)
        : new Set<string>();
      const preparedByPosition = new Map(preparedRunHyperlinks.map((prepared) => [
        `${prepared.paragraphIndex}:${prepared.runIndex}`,
        prepared,
      ] as const));
      const relationshipIds = paragraphs.map(({ runs }) =>
        runs.map(() => undefined as string | undefined));

      for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
        for (const [runIndex, run] of paragraph.runs.entries()) {
          const hyperlink = run.style?.hyperlink;
          if (hyperlink === undefined || hyperlink === false) continue;
          const previous = current.runHyperlinkBindings[paragraphIndex]?.[runIndex];
          if (previous && shapeHyperlinkTargetsEqual(previous.hyperlink, hyperlink)) {
            relationshipIds[paragraphIndex]![runIndex] = previous.relationshipId;
            continue;
          }
          const prepared = preparedByPosition.get(`${paragraphIndex}:${runIndex}`);
          if (!prepared) throw new Error('Prepared rich text run hyperlink was not found');
          const relationshipId = previous
            && relationshipReferenceCount(xml, previous.relationshipId) === 1
            ? this.presentation.opcPackage.updateRelationship(
                this.partUri,
                previous.relationshipId,
                prepared.relationship,
              ).id
            : this.presentation.opcPackage.addRelationship(
                this.partUri,
                prepared.relationship,
              ).id;
          relationshipIds[paragraphIndex]![runIndex] = relationshipId;
        }
      }

      const updated = replaceRichText(
        xml,
        element,
        paragraphs,
        this.partUri,
        relationshipIds,
      );
      this.setXml(updated);
      const updatedXml = LosslessXmlDocument.parse(updated);
      for (const relationshipId of previousRelationshipIds) {
        if (
          relationshipReferenceCount(updatedXml, relationshipId) === 0
          && this.relationships.some(({ id: candidate }) => candidate === relationshipId)
        ) {
          this.presentation.opcPackage.removeRelationship(this.partUri, relationshipId);
        }
      }
    });
  }

  getShapeRichText(id: number): readonly RichTextParagraph[] {
    const { xml, element } = this.resolveShape(id);
    return readRichText(xml, element, {
      relationships: this.relationships,
      slidePartUris: this.presentation.slides.map(({ partUri }) => partUri),
    });
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
      const owner = definition.placeholder === undefined
        ? undefined
        : resolvePlaceholderOwner(
            pkg,
            this.partUri,
            definition.placeholder,
            'image',
          );
      const rendered = owner === undefined
        ? definition
        : Object.freeze({
            ...definition,
            name: owner.name,
            ...owner.transform,
          });
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
      const nextId = owner?.shapeId ?? allocatePresetShapeId(xml, shapeTree, this.partUri);
      const imageCount = this.shapes.filter(({ kind }) => kind === 'image').length;
      const pictureXml = renderEmbeddedRasterImageXml(
        nextId,
        rendered,
        relationship.id,
        `Image ${imageCount}`,
        owner?.identity,
      );
      if (owner) xml.replace(owner.slideElement.start, owner.slideElement.end, pictureXml);
      else {
        const extensionList = directElementChildren(shapeTree, 'extLst')[0];
        if (extensionList) xml.replace(extensionList.start, extensionList.start, pictureXml);
        else xml.appendChildXml(shapeTree, pictureXml);
      }
      this.setXml(xml.serialize());
      if (owner) this.invalidateShapeModel(nextId);
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
      const owner = definition.placeholder === undefined
        ? undefined
        : resolvePlaceholderOwner(
            pkg,
            this.partUri,
            definition.placeholder,
            'image',
          );
      const rendered = owner === undefined
        ? definition
        : Object.freeze({
            ...definition,
            name: owner.name,
            ...owner.transform,
          });
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
      const nextId = owner?.shapeId ?? allocatePresetShapeId(xml, shapeTree, this.partUri);
      const imageCount = this.shapes.filter(({ kind }) => kind === 'image').length;
      const pictureXml = renderEmbeddedSvgImageXml(
        nextId,
        rendered,
        fallbackRelationship.id,
        svgRelationship.id,
        `Image ${imageCount}`,
        owner?.identity,
      );
      if (owner) xml.replace(owner.slideElement.start, owner.slideElement.end, pictureXml);
      else {
        const extensionList = directElementChildren(shapeTree, 'extLst')[0];
        if (extensionList) xml.replace(extensionList.start, extensionList.start, pictureXml);
        else xml.appendChildXml(shapeTree, pictureXml);
      }
      this.setXml(xml.serialize());
      if (owner) this.invalidateShapeModel(nextId);
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
      const owner = normalized.placeholder === undefined
        ? undefined
        : resolvePlaceholderOwner(
            this.presentation.opcPackage,
            this.partUri,
            normalized.placeholder,
            'text-shape',
          );
      const rendered = owner === undefined
        ? normalized
        : Object.freeze({
            ...normalized,
            name: owner.name,
            ...owner.transform,
          });
      let targetSlide: SlideModel | undefined;
      if (rendered.hyperlink?.slide !== undefined) {
        targetSlide = this.presentation.slides[rendered.hyperlink.slide - 1];
        if (!targetSlide) {
          throw new RangeError(
            `Preset shape hyperlink slide ${rendered.hyperlink.slide} is out of range`,
          );
        }
      }
      const { xml } = this.parse();
      const shapeTree = requirePresetShapeTree(xml, this.partUri);
      const nextId = owner?.shapeId ?? allocatePresetShapeId(xml, shapeTree, this.partUri);
      let hyperlinkRelationshipId: string | undefined;
      if (rendered.hyperlink?.url !== undefined) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: HYPERLINK_RELATIONSHIP_TYPE,
          target: rendered.hyperlink.url,
          targetMode: 'External',
        }).id;
      } else if (targetSlide) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: SLIDE_RELATIONSHIP_TYPE,
          target: relativeRelationshipTarget(this.partUri, targetSlide.partUri),
          targetMode: 'Internal',
        }).id;
      }
      const shapeXml = renderPresetShapeXml(
        nextId,
        rendered,
        hyperlinkRelationshipId,
        owner?.identity,
      );
      if (owner) xml.replace(owner.slideElement.start, owner.slideElement.end, shapeXml);
      else {
        const extensionList = directChildren(shapeTree, 'extLst')[0];
        if (extensionList) xml.replace(extensionList.start, extensionList.start, shapeXml);
        else xml.appendChildXml(shapeTree, shapeXml);
      }
      this.setXml(xml.serialize());
      if (owner) this.invalidateShapeModel(nextId);
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
      const owner = normalized.placeholder === undefined
        ? undefined
        : resolvePlaceholderOwner(
            this.presentation.opcPackage,
            this.partUri,
            normalized.placeholder,
            'text-shape',
          );
      const rendered = owner === undefined
        ? normalized
        : Object.freeze({
            ...normalized,
            name: owner.name,
            ...owner.transform,
          });
      let targetSlide: SlideModel | undefined;
      if (rendered.hyperlink?.slide !== undefined) {
        targetSlide = this.presentation.slides[rendered.hyperlink.slide - 1];
        if (!targetSlide) {
          throw new RangeError(
            `Custom shape hyperlink slide ${rendered.hyperlink.slide} is out of range`,
          );
        }
      }
      const { xml } = this.parse();
      const shapeTree = requirePresetShapeTree(xml, this.partUri);
      const nextId = owner?.shapeId ?? allocatePresetShapeId(xml, shapeTree, this.partUri);
      let hyperlinkRelationshipId: string | undefined;
      if (rendered.hyperlink?.url !== undefined) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: HYPERLINK_RELATIONSHIP_TYPE,
          target: rendered.hyperlink.url,
          targetMode: 'External',
        }).id;
      } else if (targetSlide) {
        hyperlinkRelationshipId = this.presentation.opcPackage.addRelationship(this.partUri, {
          type: SLIDE_RELATIONSHIP_TYPE,
          target: relativeRelationshipTarget(this.partUri, targetSlide.partUri),
          targetMode: 'Internal',
        }).id;
      }
      const shapeXml = renderCustomShapeXml(
        nextId,
        rendered,
        hyperlinkRelationshipId,
        owner?.identity,
      );
      if (owner) xml.replace(owner.slideElement.start, owner.slideElement.end, shapeXml);
      else {
        const extensionList = directChildren(shapeTree, 'extLst')[0];
        if (extensionList) xml.replace(extensionList.start, extensionList.start, shapeXml);
        else xml.appendChildXml(shapeTree, shapeXml);
      }
      this.setXml(xml.serialize());
      if (owner) this.invalidateShapeModel(nextId);
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
    const definition = normalizeTableDefinition(rows, options);
    const pageDefinitions = definition.autoPage === undefined
      ? Object.freeze([definition])
      : planTableAutoPages(definition, this.presentation.slideSize);
    const insertionPlans = pageDefinitions.slice(1).map(() =>
      this.presentation.prepareSlideInsertionAfter(this));
    const preparedHyperlinks = pageDefinitions.map((page) =>
      this.prepareTableCellHyperlinks(page));
    const generatedPartUris: string[] = [];
    try {
      const table = this.presentation.opcPackage.transaction(() => {
        const sourceTable = this.commitNormalizedTable(
          pageDefinitions[0]!,
          preparedHyperlinks[0]!,
        );
        let after: SlideModel = this;
        for (let index = 1; index < pageDefinitions.length; index += 1) {
          const generated = this.presentation.insertPreparedBlankSlideAfter(
            after,
            insertionPlans[index - 1]!,
          );
          generatedPartUris.push(generated.partUri);
          generated.commitNormalizedTable(
            pageDefinitions[index]!,
            preparedHyperlinks[index]!,
          );
          after = generated;
        }
        return sourceTable;
      });
      this.#newAutoPagedSlidePartUris = Object.freeze([...generatedPartUris]);
      return table;
    } catch (error) {
      for (const partUri of generatedPartUris) {
        this.presentation.discardDetachedSlideModel(partUri);
      }
      throw error;
    }
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
    const groups = Array.isArray(typeOrGroups)
      ? typeOrGroups
      : [{
          type: typeOrGroups as ChartType,
          series: seriesOrOptions as readonly ChartSeriesInput[],
        }];
    const prepared = await prepareChartCreation(
      groups,
      Array.isArray(typeOrGroups) ? seriesOrOptions as AddChartOptions | undefined : options,
    );
    return this.presentation.opcPackage.transaction(() => commitPreparedChart(this, prepared));
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
      const defaultColor = this.color;
      const normalized = validateTextInput(value, options);
      const owner = options.placeholder === undefined
        ? undefined
        : resolvePlaceholderOwner(
            this.presentation.opcPackage,
            this.partUri,
            options.placeholder,
            'text-shape',
          );
      const hyperlinkRelationshipId = this.createTextHyperlinkRelationship(
        normalized.hyperlink,
      );
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
          defaultColor,
          normalized.hyperlink,
          hyperlinkRelationshipId,
        ))
        .join('');
      return this.addTextShape(
        paragraphs,
        owner ? placeholderTextOptions(owner) : options,
        normalized.shape,
        normalized.rectRadius,
        owner?.isTextBox ?? normalized.isTextBox,
        normalized.fill,
        normalized.line,
        normalized.arrows,
        normalized.shadow,
        normalized.hyperlink,
        hyperlinkRelationshipId,
        false,
        normalized.margin,
        normalized.verticalAlignment,
        normalized.textDirection,
        normalized.textFit,
        normalized.textWrap,
        owner?.identity,
        owner,
      );
    });
  }

  addPlaceholder(
    value: string | readonly RichTextParagraph[],
    options: AddPlaceholderOptions,
  ): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const defaultColor = this.color;
      const plain = typeof value === 'string' ? validateTextInput(value, options) : undefined;
      const rich = typeof value === 'string' ? undefined : normalizeRichText(value);
      const defaults = typeof value === 'string' ? undefined : validateAddTextOptions(options);
      if (typeof options.name !== 'string' || options.name.length === 0) {
        throw new TypeError('Placeholder name must be a non-empty string');
      }
      const shapes = this.shapes;
      if (shapes.some(({ name }) => name === options.name)) {
        throw new RangeError(`Placeholder name ${options.name} is already in use`);
      }
      const placeholders = shapes.flatMap(({ placeholder }) =>
        placeholder === undefined ? [] : [placeholder]);
      const identity = normalizePlaceholderIdentity({
        type: options.type,
        index: options.index ?? 100 + placeholders.length,
      });
      if (placeholders.some((candidate) =>
        candidate.type === identity.type && candidate.index === identity.index)) {
        throw new RangeError(
          `Placeholder identity ${identity.type}:${identity.index} is already in use`,
        );
      }
      const hyperlink = plain ? plain.hyperlink : defaults!.hyperlink;
      const preparedRunHyperlinks = rich === undefined
        ? []
        : this.prepareRichTextRunHyperlinks(rich);
      const hyperlinkRelationshipId = this.createTextHyperlinkRelationship(hyperlink);
      const runHyperlinkRelationshipIds = rich === undefined
        ? undefined
        : this.createRichTextRunHyperlinkRelationships(rich, preparedRunHyperlinks);
      if (plain) {
        const bullet = plain.bullet === false ? undefined : plain.bullet;
        const spacing = resolveParagraphSpacing(plain.spacing);
        const paragraphs = plain.value
          .split('\n')
          .map((line) => textParagraphXml(
            line,
            'a:',
            options.align,
            plain.rtl,
            bullet,
            spacing,
            plain.level,
            plain.tabStops,
            plain.language,
            plain.marginLeft,
            plain.marginRight,
            plain.indent,
            defaultColor,
            plain.hyperlink,
            hyperlinkRelationshipId,
          ))
          .join('');
        return this.addTextShape(
          paragraphs,
          options,
          plain.shape,
          plain.rectRadius,
          plain.isTextBox,
          plain.fill,
          plain.line,
          plain.arrows,
          plain.shadow,
          plain.hyperlink,
          hyperlinkRelationshipId,
          false,
          plain.margin,
          plain.verticalAlignment,
          plain.textDirection,
          plain.textFit,
          plain.textWrap,
          identity,
        );
      }
      return this.addTextShape(
        renderRichTextParagraphs(rich!, {
          ...(defaultColor !== undefined ? { defaultColor } : {}),
          ...(defaults!.language !== undefined
            ? { defaultLanguage: defaults!.language }
            : {}),
          ...(options.align ? { defaultAlign: options.align } : {}),
          ...(defaults!.rtl !== undefined ? { defaultRtl: defaults!.rtl } : {}),
          ...(defaults!.bullet !== undefined ? { defaultBullet: defaults!.bullet } : {}),
          ...(defaults!.indent !== undefined ? { defaultIndent: defaults!.indent } : {}),
          ...(defaults!.level !== undefined ? { defaultLevel: defaults!.level } : {}),
          ...(defaults!.marginLeft !== undefined
            ? { defaultMarginLeft: defaults!.marginLeft }
            : {}),
          ...(defaults!.marginRight !== undefined
            ? { defaultMarginRight: defaults!.marginRight }
            : {}),
          ...(defaults!.spacing !== undefined ? { defaultSpacing: defaults!.spacing } : {}),
          ...(defaults!.tabStops !== undefined ? { defaultTabStops: defaults!.tabStops } : {}),
          ...(defaults!.hyperlink !== undefined
            ? {
                defaultHyperlink: defaults!.hyperlink,
                hyperlinkRelationshipId: hyperlinkRelationshipId!,
              }
            : {}),
          ...(runHyperlinkRelationshipIds === undefined
            ? {}
            : { runHyperlinkRelationshipIds }),
        }),
        options,
        defaults!.shape,
        defaults!.rectRadius,
        defaults!.isTextBox,
        defaults!.fill,
        defaults!.line,
        defaults!.arrows,
        defaults!.shadow,
        defaults!.hyperlink,
        hyperlinkRelationshipId,
        preparedRunHyperlinks.length > 0,
        defaults!.margin,
        defaults!.verticalAlignment,
        defaults!.textDirection,
        defaults!.textFit,
        defaults!.textWrap,
        identity,
      );
    });
  }

  addRichText(value: readonly RichTextParagraph[], options: AddTextOptions = {}): ShapeModel {
    return this.presentation.opcPackage.transaction(() => {
      const defaultColor = this.color;
      const paragraphs = normalizeRichText(value);
      const defaults = validateAddTextOptions(options);
      const owner = options.placeholder === undefined
        ? undefined
        : resolvePlaceholderOwner(
            this.presentation.opcPackage,
            this.partUri,
            options.placeholder,
            'text-shape',
          );
      const preparedRunHyperlinks = this.prepareRichTextRunHyperlinks(paragraphs);
      const hyperlinkRelationshipId = this.createTextHyperlinkRelationship(defaults.hyperlink);
      const runHyperlinkRelationshipIds = this.createRichTextRunHyperlinkRelationships(
        paragraphs,
        preparedRunHyperlinks,
      );
      return this.addTextShape(
        renderRichTextParagraphs(paragraphs, {
          ...(defaultColor !== undefined ? { defaultColor } : {}),
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
          ...(defaults.hyperlink !== undefined
            ? {
                defaultHyperlink: defaults.hyperlink,
                hyperlinkRelationshipId: hyperlinkRelationshipId!,
              }
            : {}),
          runHyperlinkRelationshipIds,
        }),
        owner ? placeholderTextOptions(owner) : options,
        defaults.shape,
        defaults.rectRadius,
        owner?.isTextBox ?? defaults.isTextBox,
        defaults.fill,
        defaults.line,
        defaults.arrows,
        defaults.shadow,
        defaults.hyperlink,
        hyperlinkRelationshipId,
        preparedRunHyperlinks.length > 0,
        defaults.margin,
        defaults.verticalAlignment,
        defaults.textDirection,
        defaults.textFit,
        defaults.textWrap,
        owner?.identity,
        owner,
      );
    });
  }

  setXml(xml: string): void {
    this.presentation.setXmlPart(this.partUri, xml);
  }

  getShapePlaceholder(id: number): Readonly<PlaceholderIdentity> | undefined {
    const { xml, element } = this.resolveShape(id);
    return readShapePlaceholder(xml, element);
  }

  private addTextShape(
    paragraphs: string,
    options: AddTextOptions,
    presetType: PresetShapeType,
    rectRadius: Emu | undefined,
    isTextBox: boolean,
    fill: ShapeFill,
    line: NormalizedSimpleLine,
    arrows: NormalizedShapeArrows | undefined,
    shadow: NormalizedShapeShadow | undefined,
    hyperlink: NormalizedHyperlink | undefined,
    hyperlinkRelationshipId: string | undefined,
    hasRunHyperlinks: boolean,
    margins: TextBoxMargins | undefined,
    verticalAlignment: TextBoxVerticalAlignment,
    textDirection: TextBoxTextDirection | undefined,
    textFit: TextBoxFit | undefined,
    textWrap: boolean,
    placeholder?: Readonly<PlaceholderIdentity>,
    owner?: ResolvedPlaceholderOwner,
  ): ShapeModel {
    const { xml } = this.parse();
    const shapeTree = requirePresetShapeTree(
      xml,
      this.partUri,
      'Slide does not contain a shape tree',
    );
    const nextId = owner?.shapeId ?? allocateShapeId(xml);
    const shapeXml = textShapeXml(
      nextId,
      paragraphs,
      options,
      presetType,
      rectRadius,
      isTextBox,
      fill,
      line,
      arrows,
      shadow,
      hyperlink,
      hyperlinkRelationshipId,
      hasRunHyperlinks,
      margins,
      verticalAlignment,
      textDirection,
      textFit,
      textWrap,
      placeholder,
    );
    if (owner) xml.replace(owner.slideElement.start, owner.slideElement.end, shapeXml);
    else {
      const extensionList = shapeTree.children.find(
        (child): child is XmlElement => child.type === 'element' && child.localName === 'extLst',
      );
      if (extensionList) xml.replace(extensionList.start, extensionList.start, shapeXml);
      else xml.appendChildXml(shapeTree, shapeXml);
    }
    this.setXml(xml.serialize());
    if (owner) this.invalidateShapeModel(nextId);
    const shape = this.shapes.find((candidate) => candidate.id === nextId);
    if (!(shape instanceof ShapeModel) || shape.kind !== 'text') {
      throw new ModelParseError(`Created text shape ${nextId} could not be resolved`, this.partUri);
    }
    return shape;
  }

  private createTextHyperlinkRelationship(
    hyperlink: NormalizedHyperlink | undefined,
  ): string | undefined {
    if (hyperlink === undefined) return undefined;
    if (hyperlink.url !== undefined) {
      return this.presentation.opcPackage.addRelationship(this.partUri, {
        type: HYPERLINK_RELATIONSHIP_TYPE,
        target: hyperlink.url,
        targetMode: 'External',
      }).id;
    }
    const target = this.presentation.slides[hyperlink.slide - 1];
    if (!target) {
      throw new RangeError(`Text shape hyperlink slide ${hyperlink.slide} is out of range`);
    }
    return this.presentation.opcPackage.addRelationship(this.partUri, {
      type: SLIDE_RELATIONSHIP_TYPE,
      target: relativeRelationshipTarget(this.partUri, target.partUri),
      targetMode: 'Internal',
    }).id;
  }

  private commitNormalizedTable(
    definition: Readonly<NormalizedTableDefinition>,
    preparedHyperlinks: readonly PreparedTableCellHyperlink[],
  ): TableModel {
    const owner = definition.placeholder === undefined
      ? undefined
      : resolvePlaceholderOwner(
          this.presentation.opcPackage,
          this.partUri,
          definition.placeholder,
          'table',
        );
    const rendered = owner === undefined
      ? definition
      : {
          ...definition,
          name: owner.name,
          x: owner.transform.x,
          y: owner.transform.y,
          width: owner.transform.width,
          height: owner.transform.height,
          autoRowHeight: false,
          columnWidths: scaleTableDimensions(
            definition.columnWidths,
            owner.transform.width,
            'Table placeholder width',
          ),
          rowHeights: scaleTableDimensions(
            definition.rowHeights,
            owner.transform.height,
            'Table placeholder height',
          ),
        };
    const { xml } = this.parse();
    const shapeTree = requireTableShapeTree(xml, this.partUri);
    const nextId = owner?.shapeId ?? allocateShapeId(xml);
    const hyperlinkRelationships = this.createTableCellHyperlinkRelationships(
      definition,
      preparedHyperlinks,
    );
    const tableXml = renderTableGraphicFrame(
      nextId,
      rendered,
      owner?.identity,
      owner?.transform,
      hyperlinkRelationships.defaultRelationshipIds,
      hyperlinkRelationships.runRelationshipIds,
    );
    if (owner) xml.replace(owner.slideElement.start, owner.slideElement.end, tableXml);
    else {
      const extensionList = directChildren(shapeTree, 'extLst')[0];
      if (extensionList) xml.replace(extensionList.start, extensionList.start, tableXml);
      else xml.appendChildXml(shapeTree, tableXml);
    }
    this.setXml(xml.serialize());
    if (owner) this.invalidateShapeModel(nextId);
    const table = this.shapes.find((candidate) => candidate.id === nextId);
    if (!(table instanceof TableModel) || table.kind !== 'table') {
      throw new ModelParseError(`Created table ${nextId} could not be resolved`, this.partUri);
    }
    return table;
  }

  private prepareTableCellHyperlinks(
    definition: Readonly<NormalizedTableDefinition>,
  ): readonly PreparedTableCellHyperlink[] {
    const prepared: PreparedTableCellHyperlink[] = [];
    for (const [rowIndex, row] of definition.rows.entries()) {
      for (const [columnIndex, cell] of row.entries()) {
        if (cell.continuation !== undefined) continue;
        const append = (
          hyperlink: NormalizedHyperlink,
          paragraphIndex?: number,
          runIndex?: number,
        ): void => {
          let target: PreparedTableCellHyperlink['target'];
          if (hyperlink.url !== undefined) {
            target = { kind: 'external', url: hyperlink.url };
          } else {
            const targetSlide = this.presentation.slides[hyperlink.slide - 1];
            if (!targetSlide) {
              throw new RangeError(
                `Table cell ${rowIndex},${columnIndex} hyperlink slide ` +
                `${hyperlink.slide} is out of range`,
              );
            }
            target = { kind: 'slide', partUri: targetSlide.partUri };
          }
          prepared.push({
            rowIndex,
            columnIndex,
            ...(paragraphIndex === undefined ? {} : { paragraphIndex }),
            ...(runIndex === undefined ? {} : { runIndex }),
            target,
          });
        };

        const inheritsCellHyperlink = cell.richText === undefined
          || cell.richText.some(({ runs }) =>
            runs.some(({ style }) => style?.hyperlink === undefined));
        if (cell.hyperlink !== undefined && inheritsCellHyperlink) {
          append(cell.hyperlink);
        }
        for (const [paragraphIndex, paragraph] of (cell.richText ?? []).entries()) {
          for (const [runIndex, run] of paragraph.runs.entries()) {
            const hyperlink = run.style?.hyperlink;
            if (hyperlink !== undefined && hyperlink !== false) {
              append(hyperlink, paragraphIndex, runIndex);
            }
          }
        }
      }
    }
    return prepared;
  }

  private createTableCellHyperlinkRelationships(
    definition: Readonly<NormalizedTableDefinition>,
    prepared: readonly PreparedTableCellHyperlink[],
  ): CreatedTableCellHyperlinkRelationships {
    const defaultRelationshipIds = definition.rows.map((row) =>
      row.map(() => undefined as string | undefined));
    const runRelationshipIds = definition.rows.map((row) => row.map((cell) =>
      cell.continuation === undefined
        ? (cell.richText ?? [{ runs: [{ text: cell.text }] }]).map(({ runs }) =>
            runs.map(() => undefined as string | undefined))
        : []));
    for (const {
      rowIndex,
      columnIndex,
      paragraphIndex,
      runIndex,
      target,
    } of prepared) {
      const relationship: RelationshipInput = target.kind === 'external'
        ? {
            type: HYPERLINK_RELATIONSHIP_TYPE,
            target: target.url,
            targetMode: 'External',
          }
        : {
            type: SLIDE_RELATIONSHIP_TYPE,
            target: relativeRelationshipTarget(this.partUri, target.partUri),
            targetMode: 'Internal',
          };
      const id = this.presentation.opcPackage.addRelationship(
        this.partUri,
        relationship,
      ).id;
      if (paragraphIndex === undefined || runIndex === undefined) {
        defaultRelationshipIds[rowIndex]![columnIndex] = id;
      } else {
        runRelationshipIds[rowIndex]![columnIndex]![paragraphIndex]![runIndex] = id;
      }
    }
    return { defaultRelationshipIds, runRelationshipIds };
  }

  private prepareRichTextRunHyperlinks(
    paragraphs: ReturnType<typeof normalizeRichText>,
  ): readonly PreparedRichTextRunHyperlink[] {
    const prepared: PreparedRichTextRunHyperlink[] = [];
    for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
      for (const [runIndex, run] of paragraph.runs.entries()) {
        const hyperlink = run.style?.hyperlink;
        if (hyperlink === undefined || hyperlink === false) continue;
        if (hyperlink.url !== undefined) {
          prepared.push({
            paragraphIndex,
            runIndex,
            relationship: {
              type: HYPERLINK_RELATIONSHIP_TYPE,
              target: hyperlink.url,
              targetMode: 'External',
            },
          });
          continue;
        }
        const target = this.presentation.slides[hyperlink.slide - 1];
        if (!target) {
          throw new RangeError(
            `Rich text run hyperlink slide ${hyperlink.slide} is out of range`,
          );
        }
        prepared.push({
          paragraphIndex,
          runIndex,
          relationship: {
            type: SLIDE_RELATIONSHIP_TYPE,
            target: relativeRelationshipTarget(this.partUri, target.partUri),
            targetMode: 'Internal',
          },
        });
      }
    }
    return prepared;
  }

  private createRichTextRunHyperlinkRelationships(
    paragraphs: ReturnType<typeof normalizeRichText>,
    prepared: readonly PreparedRichTextRunHyperlink[],
  ): RichTextRunHyperlinkRelationshipIds {
    const relationshipIds = paragraphs.map(({ runs }) =>
      runs.map(() => undefined as string | undefined));
    for (const { paragraphIndex, runIndex, relationship } of prepared) {
      relationshipIds[paragraphIndex]![runIndex] =
        this.presentation.opcPackage.addRelationship(this.partUri, relationship).id;
    }
    return relationshipIds;
  }

  /** @internal */
  invalidateShapeModel(id: number): void {
    const existing = this.#shapeModels.get(id);
    if (existing) this.#staleShapeModels.add(existing);
    this.#shapeModels.delete(id);
  }
}

function placeholderTextOptions(owner: ResolvedPlaceholderOwner): AddTextOptions {
  return {
    name: owner.name,
    x: owner.transform.x,
    y: owner.transform.y,
    width: owner.transform.width,
    height: owner.transform.height,
    rotation: owner.transform.rotation,
    flipHorizontal: owner.transform.flipHorizontal,
    flipVertical: owner.transform.flipVertical,
  };
}

function scaleTableDimensions(
  values: readonly number[],
  total: number,
  context: string,
): readonly number[] {
  if (!Number.isSafeInteger(total) || total < values.length) {
    throw new RangeError(`${context} must provide at least one EMU per item`);
  }
  const sourceTotal = values.reduce((sum, value) => sum + value, 0);
  if (sourceTotal === 0) return distributeTableDimension(total, values.length);
  let source = 0;
  let target = 0;
  return values.map((value, index) => {
    source += value;
    const boundary = index === values.length - 1
      ? total
      : Math.round(source * total / sourceTotal);
    const scaled = boundary - target;
    if (scaled <= 0) throw new RangeError(`${context} cannot preserve all item dimensions`);
    target = boundary;
    return scaled;
  });
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
  readonly arrows: NormalizedShapeArrows | undefined;
  readonly shadow: NormalizedShapeShadow | undefined;
  readonly bullet: NormalizedParagraphBullet | false | undefined;
  readonly fill: ShapeFill;
  readonly hyperlink: NormalizedHyperlink | undefined;
  readonly line: NormalizedSimpleLine;
  readonly shape: PresetShapeType;
  readonly rectRadius: Emu | undefined;
  readonly isTextBox: boolean;
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
    arrows: defaults.arrows,
    shadow: defaults.shadow,
    bullet: defaults.bullet,
    fill: defaults.fill,
    hyperlink: defaults.hyperlink,
    line: defaults.line,
    shape: defaults.shape,
    rectRadius: defaults.rectRadius,
    isTextBox: defaults.isTextBox,
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
  readonly arrows?: NormalizedShapeArrows;
  readonly shadow?: NormalizedShapeShadow;
  readonly bullet?: NormalizedParagraphBullet | false;
  readonly fill: ShapeFill;
  readonly hyperlink?: NormalizedHyperlink;
  readonly line: NormalizedSimpleLine;
  readonly shape: PresetShapeType;
  readonly rectRadius: Emu | undefined;
  readonly isTextBox: boolean;
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
  const arrows = normalizeShapeArrows(options.arrows, 'Text shape arrows');
  const hyperlink = options.hyperlink === undefined
    ? undefined
    : normalizeHyperlink(options.hyperlink, 'Text shape hyperlink');
  const shadow = options.shadow === undefined
    ? undefined
    : normalizeShapeShadow(options.shadow, 'Text shape shadow');
  const fill = normalizeSimpleFill(options.fill, 'Text shape fill') ?? { kind: 'none' };
  const line = normalizeSimpleLine(options.line, 'Text shape line') ?? { kind: 'none' };
  const shape = normalizeTextShapeType(options);
  const rectRadius = normalizeTextRectRadius(options, shape);
  const isTextBox = normalizeTextShapeIsTextBox(options);
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
    ...(arrows !== undefined ? { arrows } : {}),
    ...(shadow !== undefined ? { shadow } : {}),
    ...(bullet !== undefined ? { bullet } : {}),
    fill,
    ...(hyperlink !== undefined ? { hyperlink } : {}),
    line,
    shape,
    rectRadius,
    isTextBox,
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

function normalizeTextShapeType(options: AddTextOptions): PresetShapeType {
  const descriptor = Object.getOwnPropertyDescriptor(options, 'shape');
  if (!descriptor) return 'rect';
  if (!Object.hasOwn(descriptor, 'value')) {
    throw new TypeError('Text shape geometry must be a data property');
  }
  if (descriptor.value === undefined) return 'rect';
  return normalizePresetShapeType(descriptor.value, 'Text shape geometry');
}

function normalizeTextRectRadius(
  options: AddTextOptions,
  shape: PresetShapeType,
): Emu | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(options, 'rectRadius');
  if (!descriptor) return undefined;
  if (!Object.hasOwn(descriptor, 'value')) {
    throw new TypeError('Text rectangle radius must be a data property');
  }
  const value = descriptor.value;
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Text rectangle radius must be finite');
  }
  if (value < 0 || value > EMU_PER_INCH) {
    throw new RangeError('Text rectangle radius must be between 0 and 914400 EMU');
  }
  if (shape !== 'roundRect') {
    throw new TypeError('Text rectangle radius requires roundRect geometry');
  }
  const rounded = Math.round(value);
  if (!Number.isSafeInteger(rounded)) {
    throw new RangeError('Text rectangle radius must round to a safe EMU integer');
  }
  return (Object.is(rounded, -0) ? 0 : rounded) as Emu;
}

function normalizeTextShapeIsTextBox(options: AddTextOptions): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(options, 'isTextBox');
  if (!descriptor) return false;
  if (!Object.hasOwn(descriptor, 'value')) {
    throw new TypeError('Text shape isTextBox must be a data property');
  }
  const value = descriptor.value;
  if (value === undefined) return false;
  if (typeof value !== 'boolean') {
    throw new TypeError('Text shape isTextBox must be a boolean');
  }
  return value;
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

export function allocatePresetShapeId(
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

export function requirePresetShapeTree(
  xml: LosslessXmlDocument,
  partUri: string,
  missingTreeMessage = 'Slide must contain exactly one direct shape tree',
): XmlElement {
  const root = xml.roots.length === 1 ? xml.roots[0] : undefined;
  if (
    !root
    || !['sld', 'sldLayout', 'sldMaster'].includes(root.localName)
    || namespaceUri(root) !== PRESENTATION_NAMESPACE
  ) {
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
    throw new ModelParseError(missingTreeMessage, partUri);
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

export function directElementChildren(
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
  shape: PresetShapeType,
  rectRadius: Emu | undefined,
  isTextBox: boolean,
  fill: ShapeFill,
  line: NormalizedSimpleLine,
  arrows: NormalizedShapeArrows | undefined,
  shadow: NormalizedShapeShadow | undefined,
  hyperlink: NormalizedHyperlink | undefined,
  hyperlinkRelationshipId: string | undefined,
  hasRunHyperlinks: boolean,
  margins: TextBoxMargins | undefined,
  verticalAlignment: TextBoxVerticalAlignment,
  textDirection: TextBoxTextDirection | undefined,
  textFit: TextBoxFit | undefined,
  textWrap: boolean,
  placeholder?: Readonly<PlaceholderIdentity>,
): string {
  if ((hyperlink === undefined) !== (hyperlinkRelationshipId === undefined)) {
    throw new TypeError('Text shape hyperlink and relationship ID must be supplied together');
  }
  const x = Math.round(options.x ?? 0);
  const y = Math.round(options.y ?? 0);
  const width = Math.round(options.width ?? inches(1));
  const height = Math.round(options.height ?? inches(1));
  const adjustments = rectRadius === undefined
    ? undefined
    : normalizeShapeAdjustments([{
        name: 'adj',
        value: Math.round(rectRadius * 100_000 / Math.min(width, height)),
      }], 'Text rectangle radius adjustments');
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
  const applicationProperties = placeholder === undefined
    ? '<p:nvPr/>'
    : `<p:nvPr><p:ph type="${placeholder.type}" idx="${placeholder.index}"/></p:nvPr>`;
  const lineContents = renderSimpleLine(line, 'a:') + renderShapeArrows(arrows, 'a:');
  const lineXml = line.kind === 'none'
    ? `<a:ln>${lineContents}</a:ln>`
    : `<a:ln w="${points(line.width)}">${lineContents}</a:ln>`;
  const effectXml = shadow === undefined
    ? ''
    : `<a:effectLst>${renderSimpleShadow(shadow, 'a:')}</a:effectLst>`;
  const relationshipNamespace = hyperlink === undefined && !hasRunHyperlinks
    ? ''
    : ` xmlns:r="${RELATIONSHIP_NAMESPACE}"`;
  const hyperlinkXml = hyperlink === undefined
    ? ''
    : renderShapeHyperlink(
        hyperlink,
        hyperlinkRelationshipId!,
        { drawing: 'a', relationship: 'r' },
      );
  const nonVisualProperties = hyperlinkXml === ''
    ? `<p:cNvPr id="${id}" name="${name}"/>`
    : `<p:cNvPr id="${id}" name="${name}">${hyperlinkXml}</p:cNvPr>`;
  return `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"${relationshipNamespace}><p:nvSpPr>${nonVisualProperties}<p:cNvSpPr${isTextBox ? ' txBox="1"' : ''}/>${applicationProperties}</p:nvSpPr><p:spPr><a:xfrm${transformAttributes}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>${renderPresetShapeGeometry(shape, 'a:', adjustments)}${renderSimpleFill(fill, 'a:')}${lineXml}${effectXml}</p:spPr><p:txBody>${bodyProperties}<a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
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
  defaultColor?: Readonly<RichTextColor>,
  hyperlink?: NormalizedHyperlink,
  hyperlinkRelationshipId?: string,
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
  return `<${prefix}p>${properties}${defaultTextRunXml(
    value,
    prefix,
    language,
    defaultColor,
    hyperlink,
    hyperlinkRelationshipId,
  )}${endProperties}</${prefix}p>`;
}

function defaultTextRunXml(
  value: string,
  prefix = 'a:',
  language?: string,
  defaultColor: Readonly<RichTextColor> = { kind: 'scheme', value: 'tx1' },
  hyperlink?: NormalizedHyperlink,
  hyperlinkRelationshipId?: string,
): string {
  if ((hyperlink === undefined) !== (hyperlinkRelationshipId === undefined)) {
    throw new TypeError('Text run hyperlink and relationship ID must be supplied together');
  }
  const languageValue = escapeXmlAttribute(language ?? 'en-US');
  const alternateLanguage = language === undefined ? '' : ' altLang="en-US"';
  const underline = hyperlink === undefined ? '' : ' u="sng"';
  const hyperlinkXml = hyperlink === undefined
    ? ''
    : renderShapeHyperlink(
        hyperlink,
        hyperlinkRelationshipId!,
        { drawing: prefix.endsWith(':') ? prefix.slice(0, -1) : prefix, relationship: 'r' },
      );
  return `<${prefix}r><${prefix}rPr lang="${languageValue}"${alternateLanguage}${underline} dirty="0"><${prefix}solidFill>${renderColorChoice(defaultColor, prefix)}</${prefix}solidFill><${prefix}latin typeface="+mn-lt"/>${hyperlinkXml}</${prefix}rPr><${prefix}t xml:space="preserve">${escapeXmlText(value)}</${prefix}t></${prefix}r>`;
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
