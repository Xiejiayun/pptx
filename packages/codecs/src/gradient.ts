import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';
import type { CodecDiagnostic } from './registry.js';

export type OoxmlColorSource =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scrgb'; readonly red: number; readonly green: number; readonly blue: number }
  | { readonly kind: 'scheme'; readonly value: string }
  | { readonly kind: 'system'; readonly value: string; readonly lastColor?: string }
  | { readonly kind: 'preset'; readonly value: string };

export interface OoxmlColorTransform {
  readonly kind: string;
  readonly value: number;
}

export interface OoxmlColor {
  readonly source: OoxmlColorSource;
  readonly alpha: number;
  readonly transforms: readonly OoxmlColorTransform[];
}

export interface GradientStop {
  readonly offset: number;
  readonly color: string | OoxmlColor;
  readonly alpha?: number;
}

export interface LinearGradientFill {
  readonly kind: 'linear-gradient';
  readonly angle: number;
  readonly scaled?: boolean;
  readonly rotateWithShape?: boolean;
  readonly flip?: 'none' | 'x' | 'y' | 'xy';
  readonly stops: readonly GradientStop[];
}

export interface PathGradientFill {
  readonly kind: 'path-gradient';
  readonly path: 'circle' | 'rect' | 'shape';
  readonly rotateWithShape?: boolean;
  readonly fillRectangle?: { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number };
  readonly stops: readonly GradientStop[];
}

export type GradientFill = LinearGradientFill | PathGradientFill;

export class GradientCodec {
  readonly id = 'builtin.gradient-transparency';
  readonly priority = 100;
  readonly ownership = { elements: ['a:gradFill', 'a:solidFill'] } as const;
  readonly capabilities = ['slide-background', 'shape-fill'] as const;

  decode(element: XmlElement, xml: LosslessXmlDocument): GradientFill {
    if (element.localName !== 'gradFill') throw new Error(`Expected gradFill, received ${element.name}`);
    const stops = xml.descendants(element, 'gs').map((stop) => {
      const colorElement = stop.children.find((child): child is XmlElement => child.type === 'element');
      const color = colorElement ? decodeColor(xml, colorElement) : defaultColor();
      return {
        offset: clamp(Number(xml.attribute(stop, 'pos')?.value ?? 0) / 100_000),
        color,
        alpha: color.alpha,
      };
    });
    const linear = xml.descendants(element, 'lin')[0];
    if (linear) {
      return {
        kind: 'linear-gradient',
        angle: Number(xml.attribute(linear, 'ang')?.value ?? 0) / 60_000,
        scaled: xml.attribute(linear, 'scaled')?.value !== '0',
        rotateWithShape: xml.attribute(element, 'rotWithShape')?.value !== '0',
        flip: normalizeFlip(xml.attribute(element, 'flip')?.value),
        stops,
      };
    }
    const path = xml.descendants(element, 'path')[0];
    const rectangle = path ? xml.descendants(path, 'fillToRect')[0] : undefined;
    return {
      kind: 'path-gradient',
      path: normalizePath(path ? xml.attribute(path, 'path')?.value : undefined),
      rotateWithShape: xml.attribute(element, 'rotWithShape')?.value !== '0',
      ...(rectangle
        ? {
            fillRectangle: {
              left: percentAttribute(xml, rectangle, 'l'),
              top: percentAttribute(xml, rectangle, 't'),
              right: percentAttribute(xml, rectangle, 'r'),
              bottom: percentAttribute(xml, rectangle, 'b'),
            },
          }
        : {}),
      stops,
    };
  }

  encode(fill: GradientFill): string {
    if (fill.stops.length < 2) throw new Error('A gradient requires at least two stops');
    const stops = [...fill.stops]
      .sort((left, right) => left.offset - right.offset)
      .map(
        (stop) =>
          `<a:gs pos="${Math.round(clamp(stop.offset) * 100_000)}">${encodeColor(stop.color, stop.alpha)}</a:gs>`,
      )
      .join('');
    const attributes = `${fill.rotateWithShape === false ? ' rotWithShape="0"' : ''}${
      fill.kind === 'linear-gradient' && fill.flip && fill.flip !== 'none' ? ` flip="${fill.flip}"` : ''
    }`;
    if (fill.kind === 'linear-gradient') {
      return `<a:gradFill${attributes}><a:gsLst>${stops}</a:gsLst><a:lin ang="${Math.round(
        fill.angle * 60_000,
      )}" scaled="${fill.scaled === false ? 0 : 1}"/></a:gradFill>`;
    }
    const rectangle = fill.fillRectangle
      ? `<a:fillToRect l="${toPercent(fill.fillRectangle.left)}" t="${toPercent(
          fill.fillRectangle.top,
        )}" r="${toPercent(fill.fillRectangle.right)}" b="${toPercent(fill.fillRectangle.bottom)}"/>`
      : '';
    return `<a:gradFill${attributes}><a:gsLst>${stops}</a:gsLst><a:path path="${fill.path}">${rectangle}</a:path></a:gradFill>`;
  }

  getSlideBackground(pkg: OpcPackage, slidePartUri: string): GradientFill | undefined {
    const xml = LosslessXmlDocument.parse(pkg.requirePart(slidePartUri).bytes);
    const background = xml.elements('bg')[0];
    const gradient = background ? xml.descendants(background, 'gradFill')[0] : undefined;
    return gradient ? this.decode(gradient, xml) : undefined;
  }

  setSlideBackground(pkg: OpcPackage, slidePartUri: string, fill: GradientFill): void {
    const part = pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const cSld = xml.elements('cSld')[0];
    if (!cSld) throw new Error(`Slide ${slidePartUri} has no cSld`);
    const encoded = this.encode(fill);
    const background = xml.elements('bg')[0];
    const properties = background ? xml.descendants(background, 'bgPr')[0] : undefined;
    if (!background) {
      xml.replace(cSld.startTagEnd, cSld.startTagEnd, `<p:bg><p:bgPr>${encoded}<a:effectLst/></p:bgPr></p:bg>`);
    } else if (!properties) {
      xml.appendChildXml(background, `<p:bgPr>${encoded}<a:effectLst/></p:bgPr>`);
    } else {
      const existing = properties.children.find(
        (child): child is XmlElement =>
          child.type === 'element' && ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(child.localName),
      );
      if (existing) xml.replaceElement(existing, encoded);
      else xml.replace(properties.startTagEnd, properties.startTagEnd, encoded);
    }
    pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
  }

  getShapeFill(pkg: OpcPackage, slidePartUri: string, shapeId: number): GradientFill | undefined {
    const xml = LosslessXmlDocument.parse(pkg.requirePart(slidePartUri).bytes);
    const shape = findShape(xml, shapeId);
    const properties = shape ? xml.descendants(shape, 'spPr')[0] : undefined;
    const gradient = properties ? xml.descendants(properties, 'gradFill')[0] : undefined;
    return gradient ? this.decode(gradient, xml) : undefined;
  }

  setShapeFill(pkg: OpcPackage, slidePartUri: string, shapeId: number, fill: GradientFill): void {
    const part = pkg.requirePart(slidePartUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    const shape = findShape(xml, shapeId);
    const properties = shape ? xml.descendants(shape, 'spPr')[0] : undefined;
    if (!properties) throw new Error(`Shape ${shapeId} has no shape properties`);
    const existing = properties.children.find(
      (child): child is XmlElement =>
        child.type === 'element' && ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(child.localName),
    );
    const encoded = this.encode(fill);
    if (existing) xml.replaceElement(existing, encoded);
    else xml.replace(properties.startTagEnd, properties.startTagEnd, encoded);
    pkg.setPart(slidePartUri, xml.serialize(), part.contentType);
  }

  diagnostics(fill: GradientFill, profile: string, partUri?: string): CodecDiagnostic[] {
    const diagnostics: CodecDiagnostic[] = [];
    if (fill.kind === 'path-gradient' && profile === 'google-slides-import') {
      diagnostics.push({
        severity: 'warning',
        code: 'GRADIENT_PATH_MAY_DEGRADE',
        message: 'Google Slides import may approximate path gradients',
        ...(partUri ? { partUri } : {}),
      });
    }
    if (fill.stops.some(({ alpha, color }) => (alpha ?? (typeof color === 'string' ? 1 : color.alpha)) < 1) && profile === 'keynote-current') {
      diagnostics.push({
        severity: 'info',
        code: 'GRADIENT_ALPHA_CLIENT_VARIANCE',
        message: 'Keynote may render transparent gradient stops with minor visual differences',
        ...(partUri ? { partUri } : {}),
      });
    }
    return diagnostics;
  }
}

export function decodeColor(xml: LosslessXmlDocument, element: XmlElement): OoxmlColor {
  let source: OoxmlColorSource;
  switch (element.localName) {
    case 'scrgbClr':
      source = {
        kind: 'scrgb',
        red: Number(xml.attribute(element, 'r')?.value ?? 0) / 100_000,
        green: Number(xml.attribute(element, 'g')?.value ?? 0) / 100_000,
        blue: Number(xml.attribute(element, 'b')?.value ?? 0) / 100_000,
      };
      break;
    case 'schemeClr':
      source = { kind: 'scheme', value: xml.attribute(element, 'val')?.value ?? 'accent1' };
      break;
    case 'sysClr': {
      const lastColor = xml.attribute(element, 'lastClr')?.value;
      source = {
        kind: 'system',
        value: xml.attribute(element, 'val')?.value ?? 'windowText',
        ...(lastColor ? { lastColor } : {}),
      };
      break;
    }
    case 'prstClr':
      source = { kind: 'preset', value: xml.attribute(element, 'val')?.value ?? 'black' };
      break;
    default:
      source = { kind: 'srgb', value: xml.attribute(element, 'val')?.value ?? '000000' };
  }
  const transforms = element.children
    .filter((child): child is XmlElement => child.type === 'element')
    .map((transform) => ({ kind: transform.localName, value: Number(xml.attribute(transform, 'val')?.value ?? 0) }));
  let alpha = 1;
  for (const transform of transforms) {
    if (transform.kind === 'alpha') alpha = transform.value / 100_000;
    else if (transform.kind === 'alphaMod' || transform.kind === 'alphaModFix') alpha *= transform.value / 100_000;
    else if (transform.kind === 'alphaOff') alpha += transform.value / 100_000;
  }
  return { source, transforms, alpha: clamp(alpha) };
}

export function encodeColor(color: string | OoxmlColor, alphaOverride?: number): string {
  const normalized = typeof color === 'string' ? fromHex(color) : color;
  const source = encodeColorSource(normalized.source);
  const alphaKinds = new Set(['alpha', 'alphaMod', 'alphaModFix', 'alphaOff']);
  const transforms = normalized.transforms
    .filter(({ kind }) => alphaOverride === undefined || !alphaKinds.has(kind))
    .map(({ kind, value }) => `<a:${kind} val="${Math.round(value)}"/>`)
    .join('');
  const alpha = clamp(alphaOverride ?? normalized.alpha);
  const hasAlphaTransform = normalized.transforms.some(({ kind }) => alphaKinds.has(kind));
  const alphaXml = alpha < 1 && (alphaOverride !== undefined || !hasAlphaTransform)
    ? `<a:alpha val="${Math.round(alpha * 100_000)}"/>`
    : '';
  return source.replace('/>', `>${transforms}${alphaXml}</${colorElementName(normalized.source)}>`);
}

function encodeColorSource(source: OoxmlColorSource): string {
  switch (source.kind) {
    case 'scrgb':
      return `<a:scrgbClr r="${toPercent(source.red)}" g="${toPercent(source.green)}" b="${toPercent(source.blue)}"/>`;
    case 'scheme':
      return `<a:schemeClr val="${escapeAttribute(source.value)}"/>`;
    case 'system':
      return `<a:sysClr val="${escapeAttribute(source.value)}"${
        source.lastColor ? ` lastClr="${escapeAttribute(source.lastColor)}"` : ''
      }/>`;
    case 'preset':
      return `<a:prstClr val="${escapeAttribute(source.value)}"/>`;
    default:
      return `<a:srgbClr val="${escapeAttribute(source.value.replace(/^#/, '').toUpperCase())}"/>`;
  }
}

function colorElementName(source: OoxmlColorSource): string {
  return { srgb: 'a:srgbClr', scrgb: 'a:scrgbClr', scheme: 'a:schemeClr', system: 'a:sysClr', preset: 'a:prstClr' }[
    source.kind
  ];
}

function fromHex(value: string): OoxmlColor {
  return { source: { kind: 'srgb', value: value.replace(/^#/, '').toUpperCase() }, alpha: 1, transforms: [] };
}

function defaultColor(): OoxmlColor {
  return fromHex('000000');
}

function clamp(value: number): number {
  return Math.max(0, Math.min(1, Number.isFinite(value) ? value : 0));
}

function normalizeFlip(value: string | undefined): NonNullable<LinearGradientFill['flip']> {
  return value === 'x' || value === 'y' || value === 'xy' ? value : 'none';
}

function normalizePath(value: string | undefined): PathGradientFill['path'] {
  return value === 'rect' || value === 'shape' ? value : 'circle';
}

function percentAttribute(xml: LosslessXmlDocument, element: XmlElement, name: string): number {
  return Number(xml.attribute(element, name)?.value ?? 0) / 100_000;
}

function toPercent(value: number): number {
  return Math.round(clamp(value) * 100_000);
}

function escapeAttribute(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function findShape(xml: LosslessXmlDocument, shapeId: number): XmlElement | undefined {
  return xml
    .elements()
    .filter(({ localName }) => ['sp', 'pic', 'graphicFrame', 'grpSp'].includes(localName))
    .find((shape) => {
      const properties = xml.descendants(shape, 'cNvPr')[0];
      return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === shapeId;
    });
}
