import { LosslessXmlDocument, type XmlAttribute, type XmlElement } from '@pptx/lossless-xml';
import {
  GradientCodec,
  readMediaState,
  type GradientFill,
  type MediaKind,
  type MediaPlaybackSettings,
  type MediaSource,
  type MediaState,
  type ReplaceMediaSourceOptions,
  type ReplaceMediaPosterOptions,
} from '@pptx/codecs';
import {
  partUriBasename,
  partUriDirname,
  partUriExtension,
  relativeRelationshipTarget,
  type OpcPackage,
  type Relationship,
} from '@pptx/opc';
import type {
  ChartDefinition,
  ChartDefinitionInput,
  ChartGroupInput,
  ChartSeries,
  ChartSeriesInput,
  ChartState,
} from './chart.js';
import { normalizeChartDefinition } from './chart-definition.internal.js';
import { chartDiagnostics } from './chart-diagnostics.internal.js';
import {
  chartDefinitionDataEqual,
  chartDefinitionsEqual,
} from './chart-edit.internal.js';
import { readChartState } from './chart-state.internal.js';
import {
  buildChartWorkbook,
  chartWorkbookMatches,
} from './chart-workbook.internal.js';
import { cloneOwnedPartForMutation } from './dependency.internal.js';
import { evaluateCustomGeometry as evaluateGeometry } from './custom-geometry-evaluator.js';
import type { CustomGeometry, EvaluatedCustomGeometry } from './custom-geometry.js';
import type { Hyperlink } from './hyperlink.js';
import type { ImageSourceRectangle } from './image.js';
import {
  hasSvgImageExtensionCandidate,
  readSvgImageState,
  relationshipReferenceCount,
  type SvgImageState,
} from './svg-image-state.internal.js';
import type { SlideModel } from './slide.js';
import {
  normalizeTableCellBorders,
  readTableCellBorders,
  replaceTableCellBorders,
} from './table-cell-borders.internal.js';
import {
  normalizeTableCellFill,
  readTableCellFill,
  replaceTableCellFill,
} from './table-cell-fill.internal.js';
import {
  readTableCellHorizontalAlignment,
  replaceTableCellHorizontalAlignment,
} from './table-cell-horizontal-alignment.internal.js';
import {
  readTableCellMargins,
  replaceTableCellMargins,
} from './table-cell-margins.internal.js';
import {
  readTableCellTextFit,
  replaceTableCellTextFit,
} from './table-cell-text-fit.internal.js';
import {
  normalizeTableCellTextDirection,
  readTableCellTextDirection,
  replaceTableCellTextDirection,
} from './table-cell-text-direction.internal.js';
import {
  normalizeTableColumnWidthInput,
  readTableColumnWidths,
  replaceTableColumnWidths,
} from './table-column-widths.internal.js';
import {
  normalizeTableRowHeightInput,
  readTableRowHeights,
  replaceTableRowHeights,
} from './table-row-heights.internal.js';
import {
  readTableCellVerticalAlignment,
  replaceTableCellVerticalAlignment,
} from './table-cell-vertical-alignment.internal.js';
import { normalizeTextAlignment } from './rich-text.internal.js';
import { normalizeTextBoxFit } from './text-box-fit.internal.js';
import { normalizeTextBoxMargins } from './text-box-margins.internal.js';
import { normalizeTextBoxVerticalAlignment } from './text-box-vertical-alignment.internal.js';
import type {
  PresetShapeType,
  ShapeAdjustment,
  ShapeArrows,
  ShapeFill,
  ShapeLine,
  ShapeShadow,
} from './preset-shape.js';
import type {
  RichTextParagraph,
  RichTextColor,
  TextBoxFit,
  TextBoxMarginInput,
  TextBoxMargins,
  TextBoxTextDirection,
  TextBoxVerticalAlignment,
  TextAlignment,
} from './text.js';
import { type Emu, type OoxmlAngle, type Transform } from './units.js';

export type ShapeKind =
  | 'shape'
  | 'text'
  | 'image'
  | 'audio'
  | 'video'
  | 'table'
  | 'chart'
  | 'graphic-frame'
  | 'group'
  | 'unknown';

export type TableCellTextDirection = 'horz' | 'vert' | 'vert270' | 'wordArtVert';

export type TableCellBorderStyle = 'solid' | 'dash';

export type TableCellBorder =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly width: number;
      readonly style?: TableCellBorderStyle;
    };

export interface TableCellBorders {
  readonly top?: TableCellBorder;
  readonly right?: TableCellBorder;
  readonly bottom?: TableCellBorder;
  readonly left?: TableCellBorder;
}

export type TableCellBorderInput =
  | TableCellBorder
  | readonly [
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
      TableCellBorder | undefined,
    ]
  | TableCellBorders;

export type TableCellFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export interface TableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly horizontalAlignment?: TextAlignment;
  readonly margins?: TextBoxMargins;
  readonly textDirection?: TableCellTextDirection;
  readonly textFit?: TextBoxFit;
  readonly verticalAlignment?: TextBoxVerticalAlignment;
}

export interface TableRow {
  readonly cells: readonly TableCell[];
}

export abstract class BaseShapeModel {
  constructor(
    protected readonly slide: SlideModel,
    readonly id: number,
    private readonly initialName: string,
    readonly kind: ShapeKind,
  ) {}

  get name(): string {
    const { xml, element } = this.resolve();
    const properties = xml.descendants(element, 'cNvPr')[0];
    return properties ? xml.attribute(properties, 'name')?.value ?? this.initialName : this.initialName;
  }

  get transform(): Transform {
    const { xml, element } = this.resolve();
    const xfrm = xml.descendants(element, 'xfrm')[0];
    const off = xfrm ? xml.descendants(xfrm, 'off')[0] : undefined;
    const ext = xfrm ? xml.descendants(xfrm, 'ext')[0] : undefined;
    return {
      x: numberAttribute(xml, off, 'x') as Emu,
      y: numberAttribute(xml, off, 'y') as Emu,
      width: numberAttribute(xml, ext, 'cx') as Emu,
      height: numberAttribute(xml, ext, 'cy') as Emu,
      rotation: numberAttribute(xml, xfrm, 'rot') as OoxmlAngle,
      flipHorizontal: xml.attribute(xfrm ?? element, 'flipH')?.value === '1',
      flipVertical: xml.attribute(xfrm ?? element, 'flipV')?.value === '1',
    };
  }

  setTransform(changes: Partial<Transform>): void {
    this.slide.setShapeTransform(this.id, changes);
  }

  protected resolve(): { xml: LosslessXmlDocument; element: XmlElement } {
    return this.slide.resolveShape(this.id);
  }
}

export class ShapeModel extends BaseShapeModel {
  get customGeometry(): CustomGeometry | undefined {
    return this.slide.getShapeCustomGeometry(this.id);
  }

  set customGeometry(value: CustomGeometry) {
    this.slide.setShapeCustomGeometry(this.id, value);
  }

  evaluateCustomGeometry(): EvaluatedCustomGeometry | undefined {
    const geometry = this.customGeometry;
    if (!geometry) return undefined;
    const { width, height } = this.transform;
    return evaluateGeometry(geometry, { width, height });
  }

  get hyperlink(): Hyperlink | undefined {
    return this.slide.getShapeHyperlink(this.id);
  }

  set hyperlink(value: Hyperlink | undefined) {
    this.slide.setShapeHyperlink(this.id, value);
  }

  get presetType(): PresetShapeType | undefined {
    return this.slide.getShapePresetType(this.id);
  }

  set presetType(value: PresetShapeType) {
    this.slide.setShapePresetType(this.id, value);
  }

  get adjustments(): readonly ShapeAdjustment[] | undefined {
    return this.slide.getShapeAdjustments(this.id);
  }

  set adjustments(value: readonly ShapeAdjustment[]) {
    this.slide.setShapeAdjustments(this.id, value);
  }

  get text(): string {
    return this.slide.getShapeText(this.id);
  }

  set text(value: string) {
    this.slide.setShapeText(this.id, value);
  }

  get richText(): readonly RichTextParagraph[] {
    return this.slide.getShapeRichText(this.id);
  }

  set richText(value: readonly RichTextParagraph[]) {
    this.slide.setShapeRichText(this.id, value);
  }

  get textMargins(): TextBoxMargins | undefined {
    return this.slide.getShapeTextMargins(this.id);
  }

  set textMargins(value: TextBoxMarginInput | undefined) {
    this.slide.setShapeTextMargins(this.id, value);
  }

  get verticalAlignment(): TextBoxVerticalAlignment | undefined {
    return this.slide.getShapeTextVerticalAlignment(this.id);
  }

  set verticalAlignment(value: TextBoxVerticalAlignment | undefined) {
    this.slide.setShapeTextVerticalAlignment(this.id, value);
  }

  get textWrap(): boolean | undefined {
    return this.slide.getShapeTextWrap(this.id);
  }

  set textWrap(value: boolean | undefined) {
    this.slide.setShapeTextWrap(this.id, value);
  }

  get textDirection(): TextBoxTextDirection | undefined {
    return this.slide.getShapeTextDirection(this.id);
  }

  set textDirection(value: TextBoxTextDirection | undefined) {
    this.slide.setShapeTextDirection(this.id, value);
  }

  get textFit(): TextBoxFit | undefined {
    return this.slide.getShapeTextFit(this.id);
  }

  set textFit(value: TextBoxFit | undefined) {
    this.slide.setShapeTextFit(this.id, value);
  }

  get fill(): ShapeFill | undefined {
    return this.slide.getShapeFill(this.id);
  }

  set fill(value: ShapeFill | undefined) {
    this.slide.setShapeFill(this.id, value);
  }

  get line(): ShapeLine | undefined {
    return this.slide.getShapeLine(this.id);
  }

  set line(value: ShapeLine | undefined) {
    this.slide.setShapeLine(this.id, value);
  }

  get arrows(): ShapeArrows | undefined {
    return this.slide.getShapeArrows(this.id);
  }

  set arrows(value: ShapeArrows | undefined) {
    this.slide.setShapeArrows(this.id, value);
  }

  get shadow(): ShapeShadow | undefined {
    return this.slide.getShapeShadow(this.id);
  }

  set shadow(value: ShapeShadow | undefined) {
    this.slide.setShapeShadow(this.id, value);
  }

  get gradientFill(): GradientFill | undefined {
    return new GradientCodec().getShapeFill(this.slide.presentation.opcPackage, this.slide.partUri, this.id);
  }

  set gradientFill(value: GradientFill) {
    new GradientCodec().setShapeFill(this.slide.presentation.opcPackage, this.slide.partUri, this.id, value);
  }
}

export class MediaModel extends BaseShapeModel {
  declare readonly kind: MediaKind;

  override get name(): string {
    return super.name;
  }

  override set name(value: string) {
    this.slide.setMediaName(this.id, value);
  }

  get shapeId(): number {
    return this.id;
  }

  get slidePartUri(): string {
    return this.slide.partUri;
  }

  get altText(): string | undefined {
    return this.state().altText;
  }

  set altText(value: string | undefined) {
    this.slide.setMediaAltText(this.id, value);
  }

  get mediaPartUri(): string | undefined {
    return this.state().mediaPartUri;
  }

  get externalUrl(): string | undefined {
    return this.state().externalUrl;
  }

  get posterPartUri(): string | undefined {
    return this.state().posterPartUri;
  }

  get settings(): Readonly<MediaPlaybackSettings> {
    return this.state().settings;
  }

  set settings(value: MediaPlaybackSettings | undefined) {
    this.slide.setMediaSettings(this.id, value);
  }

  async replaceSource(
    source: MediaSource,
    options: ReplaceMediaSourceOptions = {},
  ): Promise<this> {
    await this.slide.replaceMediaSource(this.id, this.kind, source, options);
    return this;
  }

  async replacePoster(
    source?: MediaSource,
    options: ReplaceMediaPosterOptions = {},
  ): Promise<this> {
    await this.slide.replaceMediaPoster(this.id, source, options);
    return this;
  }

  remove(): void {
    this.slide.deleteMedia(this.id);
  }

  private state(): Readonly<MediaState> {
    const { xml, element } = this.resolve();
    const state = readMediaState(
      this.slide.presentation.opcPackage,
      this.slide.partUri,
      xml,
      element,
    );
    if (!state || state.kind !== this.kind) {
      throw new Error(`Media shape ${this.id} changed semantic kind on ${this.slide.partUri}`);
    }
    return state;
  }
}

export class ImageModel extends BaseShapeModel {
  get sourceRectangle(): Readonly<ImageSourceRectangle> | undefined {
    return this.slide.getImageSourceRectangle(this.id);
  }

  set sourceRectangle(value: ImageSourceRectangle | undefined) {
    this.slide.setImageSourceRectangle(this.id, value);
  }

  get sourcePartUri(): string | undefined {
    const { xml, element } = this.resolve();
    const svgState = readSvgImageState(
      xml,
      element,
      this.slide.relationships,
      this.slide.presentation.opcPackage,
    );
    if (svgState) return svgState.fallbackPartUri;
    const blip = xml.descendants(element, 'blip')[0];
    const id = blip ? xml.attribute(blip, 'r:embed')?.value ?? xml.attribute(blip, 'r:link')?.value : undefined;
    return id ? this.relationship(id)?.resolvedTarget : undefined;
  }

  get isSvg(): boolean {
    return this.svgState() !== undefined;
  }

  get fallbackPartUri(): string | undefined {
    return this.svgState()?.fallbackPartUri;
  }

  get svgPartUri(): string | undefined {
    return this.svgState()?.svgPartUri;
  }

  get externalUrl(): string | undefined {
    const { xml, element } = this.resolve();
    const blip = xml.descendants(element, 'blip')[0];
    const id = blip ? xml.attribute(blip, 'r:link')?.value : undefined;
    const relationship = id ? this.relationship(id) : undefined;
    return relationship?.targetMode === 'External' ? relationship.target : undefined;
  }

  replaceData(bytes: Uint8Array, contentType?: string): void {
    const pkg = this.slide.presentation.opcPackage;
    pkg.transaction(() => {
      const { xml, element } = this.resolve();
      if (hasSvgImageExtensionCandidate(xml, element)) {
        throw new Error('Use replaceSvgData() for SVG images');
      }
      const blip = xml.descendants(element, 'blip')[0];
      const reference = blip
        ? xml.attribute(blip, 'r:embed') ?? xml.attribute(blip, 'r:link')
        : undefined;
      const relationship = reference ? this.relationship(reference.value) : undefined;
      const target = relationship?.resolvedTarget;
      if (!reference || !relationship || !target || relationship.targetMode === 'External') {
        throw new Error(`Image ${this.id} is external or has no embedded part`);
      }
      const current = pkg.requirePart(target);
      const nextContentType = contentType ?? current.contentType;
      if (!isSharedTarget(pkg, xml, relationship)) {
        pkg.setPart(target, bytes, nextContentType);
        return;
      }
      const cloneUri = allocateSiblingPartUri(pkg, target);
      pkg.setPart(cloneUri, bytes, nextContentType);
      retargetShapeRelationship(this.slide, xml, reference, relationship, cloneUri);
    });
  }

  replaceSvgData(svgBytes: Uint8Array, fallbackPngBytes: Uint8Array): void {
    const detachedSvgBytes = normalizeReplacementBytes(svgBytes, 'SVG');
    const detachedFallbackPngBytes = normalizeReplacementBytes(
      fallbackPngBytes,
      'SVG fallback PNG',
    );
    const pkg = this.slide.presentation.opcPackage;
    pkg.transaction(() => {
      const { xml, element } = this.resolve();
      const state = readSvgImageState(
        xml,
        element,
        this.slide.relationships,
        pkg,
      );
      if (!state) {
        throw new Error(`Image ${this.id} is not a safely editable SVG image`);
      }
      replaceSvgImagePart(
        this.slide,
        xml,
        state.fallbackReference,
        state.fallbackRelationship,
        state.fallbackPartUri,
        detachedFallbackPngBytes,
        'image/png',
        '.png',
      );
      replaceSvgImagePart(
        this.slide,
        xml,
        state.svgReference,
        state.svgRelationship,
        state.svgPartUri,
        detachedSvgBytes,
        'image/svg+xml',
        '.svg',
      );
    });
  }

  private svgState(): SvgImageState | undefined {
    const { xml, element } = this.resolve();
    return readSvgImageState(
      xml,
      element,
      this.slide.relationships,
      this.slide.presentation.opcPackage,
    );
  }

  private relationship(id: string): Relationship | undefined {
    return this.slide.relationships.find((relationship) => relationship.id === id);
  }
}

export class TableModel extends BaseShapeModel {
  get rows(): readonly TableRow[] {
    const { xml, element } = this.resolve();
    return xml.descendants(element, 'tr').map((row) => ({
      cells: xml.descendants(row, 'tc').map((cell) => {
        const borders = readTableCellBorders(xml, cell);
        const fill = readTableCellFill(xml, cell);
        const horizontalAlignment = readTableCellHorizontalAlignment(xml, cell);
        const margins = readTableCellMargins(xml, cell);
        const textDirection = readTableCellTextDirection(xml, cell);
        const textFit = readTableCellTextFit(xml, cell, this.slide.partUri);
        const verticalAlignment = readTableCellVerticalAlignment(xml, cell);
        return {
          text: xml.descendants(cell, 't').map((node) => xml.text(node)).join(''),
          ...(borders !== undefined ? { borders } : {}),
          ...(fill !== undefined ? { fill } : {}),
          ...(horizontalAlignment !== undefined ? { horizontalAlignment } : {}),
          ...(margins !== undefined ? { margins } : {}),
          ...(textDirection !== undefined ? { textDirection } : {}),
          ...(textFit !== undefined ? { textFit } : {}),
          ...(verticalAlignment !== undefined ? { verticalAlignment } : {}),
        };
      }),
    }));
  }

  get columnWidths(): readonly number[] | undefined {
    const { xml, element } = this.resolve();
    return readTableColumnWidths(xml, element);
  }

  setColumnWidths(value: number | readonly number[]): void {
    const input = normalizeTableColumnWidthInput(value);
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      if (replaceTableColumnWidths(xml, element, input, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  get rowHeights(): readonly number[] | undefined {
    const { xml, element } = this.resolve();
    return readTableRowHeights(xml, element);
  }

  setRowHeights(value: number | readonly number[]): void {
    const input = normalizeTableRowHeightInput(value);
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      if (replaceTableRowHeights(xml, element, input, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  setCellText(rowIndex: number, columnIndex: number, value: string): void {
    const { xml, element } = this.resolve();
    const row = xml.descendants(element, 'tr')[rowIndex];
    const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
    const text = cell ? xml.descendants(cell, 't')[0] : undefined;
    if (!text) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
    xml.replaceText(text, value);
    this.slide.setXml(xml.serialize());
  }

  setCellBorders(
    rowIndex: number,
    columnIndex: number,
    value: TableCellBorderInput | undefined,
  ): void {
    const borders = normalizeTableCellBorders(value, 'Table cell borders');
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      const row = xml.descendants(element, 'tr')[rowIndex];
      const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
      if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      if (replaceTableCellBorders(xml, cell, borders, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  setCellFill(
    rowIndex: number,
    columnIndex: number,
    value: TableCellFill | undefined,
  ): void {
    const fill = normalizeTableCellFill(value, 'Table cell fill');
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      const row = xml.descendants(element, 'tr')[rowIndex];
      const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
      if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      if (replaceTableCellFill(xml, cell, fill, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  setCellMargins(
    rowIndex: number,
    columnIndex: number,
    value: TextBoxMarginInput | undefined,
  ): void {
    const margins = normalizeTextBoxMargins(value, 'Table cell margins');
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      const row = xml.descendants(element, 'tr')[rowIndex];
      const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
      if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      if (replaceTableCellMargins(xml, cell, margins, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  setCellTextDirection(
    rowIndex: number,
    columnIndex: number,
    value: TableCellTextDirection | undefined,
  ): void {
    const direction = value === undefined
      ? undefined
      : normalizeTableCellTextDirection(value, 'Table cell text direction');
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      const row = xml.descendants(element, 'tr')[rowIndex];
      const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
      if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      if (replaceTableCellTextDirection(xml, cell, direction, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  setCellTextFit(
    rowIndex: number,
    columnIndex: number,
    value: TextBoxFit | undefined,
  ): void {
    const fit = value === undefined
      ? undefined
      : normalizeTextBoxFit(value, 'Table cell text fit');
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      const row = xml.descendants(element, 'tr')[rowIndex];
      const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
      if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      if (replaceTableCellTextFit(xml, cell, fit, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  setCellVerticalAlignment(
    rowIndex: number,
    columnIndex: number,
    value: TextBoxVerticalAlignment | undefined,
  ): void {
    const alignment = value === undefined
      ? undefined
      : normalizeTextBoxVerticalAlignment(value, 'Table cell vertical alignment');
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      const row = xml.descendants(element, 'tr')[rowIndex];
      const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
      if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      if (replaceTableCellVerticalAlignment(xml, cell, alignment, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }

  setCellHorizontalAlignment(
    rowIndex: number,
    columnIndex: number,
    value: TextAlignment | undefined,
  ): void {
    const alignment = value === undefined
      ? undefined
      : normalizeTextAlignment(value, 'Table cell horizontal alignment');
    this.slide.presentation.opcPackage.transaction(() => {
      const { xml, element } = this.resolve();
      const row = xml.descendants(element, 'tr')[rowIndex];
      const cell = row ? xml.descendants(row, 'tc')[columnIndex] : undefined;
      if (!cell) throw new RangeError(`Table cell ${rowIndex},${columnIndex} was not found`);
      if (replaceTableCellHorizontalAlignment(xml, cell, alignment, this.slide.partUri)) {
        this.slide.setXml(xml.serialize());
      }
    });
  }
}

export class ChartModel extends BaseShapeModel {
  get altText(): string | undefined {
    const { xml, element } = this.resolve();
    const properties = xml.descendants(element, 'cNvPr')[0];
    return properties ? xml.attribute(properties, 'descr')?.value : undefined;
  }

  get chartPartUri(): string | undefined {
    const { xml, element } = this.resolve();
    const chart = xml.descendants(element, 'chart')[0];
    const id = chart ? xml.attribute(chart, 'r:id')?.value : undefined;
    return id ? this.slide.relationships.find((relationship) => relationship.id === id)?.resolvedTarget : undefined;
  }

  get definition(): Readonly<ChartDefinition> | undefined {
    const uri = this.chartPartUri;
    return uri
      ? readChartState(this.slide.presentation.opcPackage, uri).definition
      : undefined;
  }

  get series(): readonly Readonly<ChartSeries>[] {
    const definition = this.definition;
    return definition
      ? Object.freeze(definition.groups.flatMap(({ series }) => series))
      : EMPTY_CHART_SERIES;
  }

  get workbookPartUri(): string | undefined {
    const uri = this.chartPartUri;
    return uri
      ? readChartState(this.slide.presentation.opcPackage, uri).workbookPartUri
      : undefined;
  }

  get xml(): string {
    const uri = this.chartPartUri;
    return uri ? new TextDecoder().decode(this.slide.presentation.opcPackage.requirePart(uri).bytes) : '';
  }

  async replaceDefinition(value: ChartDefinitionInput): Promise<this> {
    const next = normalizeChartDefinition(value);
    const current = this.editableState();
    const workbookSynchronized = current.status === 'recognized'
      && current.workbookPartUri !== undefined
      && await chartWorkbookMatches(
        this.slide.presentation.opcPackage.requirePart(current.workbookPartUri).bytes,
        current.definition!,
        this.xml,
      );
    const definitionsEqual = current.status === 'recognized'
      && chartDefinitionsEqual(current.definition!, next);
    const dataEqual = chartDefinitionDataEqual(current.definition!, next);
    if (
      definitionsEqual
      && current.workbookPartUri
      && workbookSynchronized
    ) {
      return this;
    }
    const workbookBytes = current.status !== 'recognized'
      || !dataEqual
      || (definitionsEqual && !workbookSynchronized)
      ? await buildChartWorkbook(next)
      : undefined;
    this.slide.replaceChartDefinition(this.id, current, next, workbookBytes);
    return this;
  }

  async replaceSeries(value: readonly ChartSeriesInput[]): Promise<this> {
    const current = this.editableState();
    const definition = current.definition!;
    if (definition.groups.length !== 1) {
      throw new Error('replaceSeries() requires a chart with exactly one group');
    }
    const group = definition.groups[0]!;
    const replacementGroup = {
      type: group.type,
      series: value,
      ...(group.axis === undefined ? {} : { axis: group.axis }),
      ...(group.options === undefined ? {} : { options: group.options }),
    } as ChartGroupInput;
    return this.replaceDefinition({
      groups: [replacementGroup],
      options: definition.options,
    });
  }

  async diagnostics(): Promise<readonly import('./chart.js').ChartDiagnostic[]> {
    return (await chartDiagnostics(
      this.slide.presentation.opcPackage,
      this.slide.partUri,
    )).filter(({ objectId }) => objectId === String(this.id));
  }

  remove(): void {
    this.slide.deleteChart(this.id);
  }

  setXml(value: string): void {
    LosslessXmlDocument.parse(value);
    const pkg = this.slide.presentation.opcPackage;
    pkg.transaction(() => {
      const { xml, element } = this.resolve();
      const chart = xml.descendants(element, 'chart')[0];
      const reference = chart ? xml.attribute(chart, 'r:id') : undefined;
      const relationship = reference
        ? this.slide.relationships.find(({ id }) => id === reference.value)
        : undefined;
      const target = relationship?.resolvedTarget;
      if (!reference || !relationship || !target) throw new Error(`Chart ${this.id} has no chart part`);
      const current = pkg.requirePart(target);
      if (!isSharedTarget(pkg, xml, relationship)) {
        pkg.setPart(target, value, current.contentType);
        return;
      }
      const cloneUri = cloneOwnedPartForMutation(pkg, target);
      retargetShapeRelationship(this.slide, xml, reference, relationship, cloneUri);
      pkg.setPart(cloneUri, value, current.contentType);
    });
  }

  private editableState(): Readonly<ChartState> {
    const uri = this.chartPartUri;
    if (!uri) throw new Error(`Chart ${this.id} has no chart part`);
    const state = readChartState(this.slide.presentation.opcPackage, uri);
    if (
      (state.status !== 'recognized' && state.status !== 'cache-only')
      || !state.definition
    ) {
      throw new Error(`Chart semantic editing is unavailable for ${state.status} state`);
    }
    return state;
  }
}

export type SemanticShape =
  | ShapeModel
  | ImageModel
  | MediaModel
  | TableModel
  | ChartModel
  | BaseShapeModel;

export function decodeShape(slide: SlideModel, xml: LosslessXmlDocument, element: XmlElement): SemanticShape | undefined {
  const properties = xml.descendants(element, 'cNvPr')[0];
  if (!properties) return undefined;
  const id = Number.parseInt(xml.attribute(properties, 'id')?.value ?? '', 10);
  if (!Number.isFinite(id)) return undefined;
  const name = xml.attribute(properties, 'name')?.value ?? `Shape ${id}`;
  if (element.localName === 'pic') {
    const media = readMediaState(
      slide.presentation.opcPackage,
      slide.partUri,
      xml,
      element,
    );
    if (media) return new MediaModel(slide, id, name, media.kind);
    return new ImageModel(slide, id, name, 'image');
  }
  if (element.localName === 'graphicFrame') {
    if (xml.descendants(element, 'tbl').length > 0) return new TableModel(slide, id, name, 'table');
    if (xml.descendants(element, 'chart').length > 0) return new ChartModel(slide, id, name, 'chart');
    return new ShapeModel(slide, id, name, 'graphic-frame');
  }
  if (element.localName === 'grpSp') return new ShapeModel(slide, id, name, 'group');
  const hasText = xml.descendants(element, 'txBody').length > 0;
  return new ShapeModel(slide, id, name, hasText ? 'text' : 'shape');
}

function numberAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement | undefined,
  name: string,
): number {
  if (!element) return 0;
  const value = Number.parseInt(xml.attribute(element, name)?.value ?? '0', 10);
  return Number.isFinite(value) ? value : 0;
}

function isSharedTarget(
  pkg: OpcPackage,
  xml: LosslessXmlDocument,
  relationship: Relationship,
): boolean {
  const incoming = relationship.resolvedTarget
    ? pkg.graph.find(({ uri }) => uri === relationship.resolvedTarget)?.incoming.length ?? 0
    : 0;
  return incoming > 1 || imageRelationshipReferenceCount(xml, relationship.id) > 1;
}

const EMPTY_CHART_SERIES: readonly Readonly<ChartSeries>[] = Object.freeze([]);

function replaceSvgImagePart(
  slide: SlideModel,
  xml: LosslessXmlDocument,
  reference: XmlAttribute,
  relationship: Relationship,
  targetPartUri: string,
  bytes: Uint8Array,
  contentType: 'image/png' | 'image/svg+xml',
  extension: '.png' | '.svg',
): void {
  const pkg = slide.presentation.opcPackage;
  if (
    partUriExtension(targetPartUri) === extension
    && !isSharedTarget(pkg, xml, relationship)
  ) {
    pkg.setPart(targetPartUri, bytes, contentType);
    return;
  }
  const cloneUri = allocateSiblingPartUri(pkg, targetPartUri, extension);
  pkg.setPart(cloneUri, bytes, contentType);
  retargetShapeRelationship(slide, xml, reference, relationship, cloneUri);
}

function retargetShapeRelationship(
  slide: SlideModel,
  xml: LosslessXmlDocument,
  reference: XmlAttribute,
  relationship: Relationship,
  targetPartUri: string,
): void {
  const pkg = slide.presentation.opcPackage;
  const target = relativeRelationshipTarget(slide.partUri, targetPartUri);
  if (imageRelationshipReferenceCount(xml, relationship.id) > 1) {
    const cloneRelationship = pkg.addRelationship(slide.partUri, { type: relationship.type, target });
    xml.replaceAttribute(reference, cloneRelationship.id);
    slide.setXml(xml.serialize());
  } else {
    pkg.updateRelationship(slide.partUri, relationship.id, { target, targetMode: 'Internal' });
  }
}

function imageRelationshipReferenceCount(
  xml: LosslessXmlDocument,
  id: string,
): number {
  const namespaceAwareCount = relationshipReferenceCount(xml, id);
  if (namespaceAwareCount > 0) return namespaceAwareCount;
  return xml
    .elements()
    .flatMap(({ attributes }) => attributes)
    .filter(({ name, value }) => name.startsWith('r:') && value === id).length;
}

function allocateSiblingPartUri(
  pkg: OpcPackage,
  sourcePartUri: string,
  requestedExtension?: string,
): string {
  const sourceExtension = partUriExtension(sourcePartUri);
  const extension = requestedExtension ?? (sourceExtension || '.bin');
  const basename = partUriBasename(sourcePartUri, sourceExtension);
  const stem = basename.replace(/\d+$/, '') || 'part';
  return pkg.allocatePartUri(partUriDirname(sourcePartUri), stem, extension);
}

function normalizeReplacementBytes(value: unknown, name: string): Uint8Array {
  if (!(value instanceof Uint8Array)) {
    throw new TypeError(`${name} bytes must be a Uint8Array`);
  }
  if (value.length === 0) {
    throw new RangeError(`${name} bytes must not be empty`);
  }
  return new Uint8Array(value);
}
