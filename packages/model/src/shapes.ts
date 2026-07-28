import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { GradientCodec, type GradientFill } from '@pptx/codecs';
import type { Relationship } from '@pptx/opc';
import type { SlideModel } from './slide.js';
import { type Emu, type OoxmlAngle, type Transform } from './units.js';

export type ShapeKind = 'shape' | 'text' | 'image' | 'table' | 'chart' | 'graphic-frame' | 'group' | 'unknown';

export interface TableCell {
  readonly text: string;
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
  get text(): string {
    const { xml, element } = this.resolve();
    return xml.descendants(element, 't').map((node) => xml.text(node)).join('');
  }

  set text(value: string) {
    this.slide.setShapeText(this.id, value);
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
    const target = this.sourcePartUri;
    if (!target) throw new Error(`Image ${this.id} is external or has no embedded part`);
    const current = this.slide.presentation.opcPackage.requirePart(target);
    this.slide.presentation.opcPackage.setPart(target, bytes, contentType ?? current.contentType);
  }

  private relationship(id: string): Relationship | undefined {
    return this.slide.relationships.find((relationship) => relationship.id === id);
  }
}

export class TableModel extends BaseShapeModel {
  get rows(): readonly TableRow[] {
    const { xml, element } = this.resolve();
    return xml.descendants(element, 'tr').map((row) => ({
      cells: xml.descendants(row, 'tc').map((cell) => ({
        text: xml.descendants(cell, 't').map((node) => xml.text(node)).join(''),
      })),
    }));
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
    const uri = this.chartPartUri;
    if (!uri) throw new Error(`Chart ${this.id} has no chart part`);
    const current = this.slide.presentation.opcPackage.requirePart(uri);
    LosslessXmlDocument.parse(value);
    this.slide.presentation.opcPackage.setPart(uri, value, current.contentType);
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
