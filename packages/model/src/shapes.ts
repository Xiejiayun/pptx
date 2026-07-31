import { LosslessXmlDocument, type XmlAttribute, type XmlElement } from '@pptx/lossless-xml';
import { GradientCodec, type GradientFill } from '@pptx/codecs';
import {
  partUriBasename,
  partUriDirname,
  partUriExtension,
  relativeRelationshipTarget,
  type OpcPackage,
  type Relationship,
} from '@pptx/opc';
import { cloneOwnedPartForMutation } from './dependency.internal.js';
import type { CustomGeometry } from './custom-geometry.js';
import type { Hyperlink } from './hyperlink.js';
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

export type ShapeKind = 'shape' | 'text' | 'image' | 'table' | 'chart' | 'graphic-frame' | 'group' | 'unknown';

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

export interface ChartSeries {
  readonly name: string;
  readonly categories: readonly string[];
  readonly values: readonly number[];
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

export class ImageModel extends BaseShapeModel {
  get sourcePartUri(): string | undefined {
    const { xml, element } = this.resolve();
    const blip = xml.descendants(element, 'blip')[0];
    const id = blip ? xml.attribute(blip, 'r:embed')?.value ?? xml.attribute(blip, 'r:link')?.value : undefined;
    return id ? this.relationship(id)?.resolvedTarget : undefined;
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
  get chartPartUri(): string | undefined {
    const { xml, element } = this.resolve();
    const chart = xml.descendants(element, 'chart')[0];
    const id = chart ? xml.attribute(chart, 'r:id')?.value : undefined;
    return id ? this.slide.relationships.find((relationship) => relationship.id === id)?.resolvedTarget : undefined;
  }

  get series(): readonly ChartSeries[] {
    const uri = this.chartPartUri;
    if (!uri) return [];
    const xml = LosslessXmlDocument.parse(this.slide.presentation.opcPackage.requirePart(uri).bytes);
    return xml.elements('ser').map((series) => {
      const nameContainer = xml.descendants(series, 'tx')[0];
      const categoryContainer = xml.descendants(series, 'cat')[0];
      const valueContainer = xml.descendants(series, 'val')[0];
      return {
        name: nameContainer ? lastValue(xml, nameContainer) : '',
        categories: categoryContainer ? pointValues(xml, categoryContainer) : [],
        values: valueContainer
          ? pointValues(xml, valueContainer).map((value) => Number(value)).filter(Number.isFinite)
          : [],
      };
    });
  }

  get xml(): string {
    const uri = this.chartPartUri;
    return uri ? new TextDecoder().decode(this.slide.presentation.opcPackage.requirePart(uri).bytes) : '';
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
}

export type SemanticShape = ShapeModel | ImageModel | TableModel | ChartModel | BaseShapeModel;

export function decodeShape(slide: SlideModel, xml: LosslessXmlDocument, element: XmlElement): SemanticShape | undefined {
  const properties = xml.descendants(element, 'cNvPr')[0];
  if (!properties) return undefined;
  const id = Number.parseInt(xml.attribute(properties, 'id')?.value ?? '', 10);
  if (!Number.isFinite(id)) return undefined;
  const name = xml.attribute(properties, 'name')?.value ?? `Shape ${id}`;
  if (element.localName === 'pic') return new ImageModel(slide, id, name, 'image');
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

function pointValues(xml: LosslessXmlDocument, element: XmlElement): string[] {
  const points = xml.descendants(element, 'pt');
  if (points.length > 0) {
    return points
      .sort(
        (left, right) =>
          Number(xml.attribute(left, 'idx')?.value ?? 0) - Number(xml.attribute(right, 'idx')?.value ?? 0),
      )
      .map((point) => lastValue(xml, point));
  }
  return xml.descendants(element, 'v').map((value) => xml.text(value));
}

function lastValue(xml: LosslessXmlDocument, element: XmlElement): string {
  const values = xml.descendants(element, 'v');
  const value = values.at(-1);
  return value ? xml.text(value) : '';
}

function isSharedTarget(
  pkg: OpcPackage,
  xml: LosslessXmlDocument,
  relationship: Relationship,
): boolean {
  const incoming = relationship.resolvedTarget
    ? pkg.graph.find(({ uri }) => uri === relationship.resolvedTarget)?.incoming.length ?? 0
    : 0;
  return incoming > 1 || relationshipReferenceCount(xml, relationship.id) > 1;
}

function relationshipReferenceCount(xml: LosslessXmlDocument, id: string): number {
  return xml
    .elements()
    .flatMap(({ attributes }) => attributes)
    .filter(({ name, value }) => name.startsWith('r:') && value === id).length;
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
  if (relationshipReferenceCount(xml, relationship.id) > 1) {
    const cloneRelationship = pkg.addRelationship(slide.partUri, { type: relationship.type, target });
    xml.replaceAttribute(reference, cloneRelationship.id);
    slide.setXml(xml.serialize());
  } else {
    pkg.updateRelationship(slide.partUri, relationship.id, { target, targetMode: 'Internal' });
  }
}

function allocateSiblingPartUri(pkg: OpcPackage, sourcePartUri: string): string {
  const extension = partUriExtension(sourcePartUri) || '.bin';
  const basename = partUriBasename(sourcePartUri, partUriExtension(sourcePartUri));
  const stem = basename.replace(/\d+$/, '') || 'part';
  return pkg.allocatePartUri(partUriDirname(sourcePartUri), stem, extension);
}
