import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';
import type {
  SlideNumber,
  SlideNumberColor,
  SlideNumberMargins,
  SlideNumberOwnerKind,
  SlideNumberTextStyle,
} from './slide-number.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX_COORDINATE = 27_273_042_316_900;
const MIN_INT32 = -2_147_483_648;
const MAX_INT32 = 2_147_483_647;
const MAX_UINT32 = 4_294_967_295;
const EMU_PER_POINT = 12_700;
const INVALID = Symbol('invalid slide number value');
const OPTION_KEYS = new Set([
  'x',
  'y',
  'width',
  'height',
  'align',
  'rtl',
  'valign',
  'margin',
  'style',
]);
const STYLE_KEYS = new Set([
  'fontFamily',
  'fontSize',
  'lang',
  'bold',
  'italic',
  'color',
  'transparency',
]);
const COLOR_KEYS = new Set(['kind', 'value']);
const MARGIN_KEYS = new Set(['top', 'right', 'bottom', 'left']);
const SCHEME_COLORS = new Set([
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'bg1',
  'bg2',
  'dk1',
  'dk2',
  'folHlink',
  'hlink',
  'lt1',
  'lt2',
  'phClr',
  'tx1',
  'tx2',
]);
const ROOT_NAMES: Readonly<Record<SlideNumberOwnerKind, string>> = {
  slide: 'sld',
  layout: 'sldLayout',
  master: 'sldMaster',
};
const ALIGN_FROM_OOXML = new Map([
  ['l', 'left' as const],
  ['ctr', 'center' as const],
  ['r', 'right' as const],
  ['just', 'justify' as const],
]);
const VALIGN_FROM_OOXML = new Map([
  ['t', 'top' as const],
  ['ctr', 'middle' as const],
  ['b', 'bottom' as const],
]);

interface PartialSlideNumberStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

export function normalizeSlideNumberOptions(value: unknown): Readonly<SlideNumber> {
  const input = readDataObject(value, OPTION_KEYS, 'Slide number options');
  const margin = normalizeMargins(input.margin);
  const style = normalizeStyle(input.style);
  return deepFreeze({
    x: normalizeCoordinate(input.x, 0, 'Slide number x'),
    y: normalizeCoordinate(input.y, 0, 'Slide number y'),
    width: normalizeExtent(input.width, 800_000, 'Slide number width'),
    height: normalizeExtent(input.height, 300_000, 'Slide number height'),
    align: normalizeEnum(
      input.align,
      ['left', 'center', 'right', 'justify'] as const,
      'left',
      'Slide number align',
    ),
    rtl: normalizeBoolean(input.rtl, false, 'Slide number rtl'),
    ...(input.valign === undefined
      ? {}
      : {
          valign: normalizeEnum(
            input.valign,
            ['top', 'middle', 'bottom'] as const,
            undefined,
            'Slide number valign',
          ),
        }),
    ...(margin === undefined ? {} : { margin }),
    style,
  });
}

export function readSlideNumber(
  pkg: OpcPackage,
  ownerPartUri: string,
  ownerKind: SlideNumberOwnerKind,
): Readonly<SlideNumber> | undefined {
  const xml = LosslessXmlDocument.parse(pkg.requirePart(ownerPartUri).bytes);
  const root = uniqueRoot(xml, ROOT_NAMES[ownerKind], PRESENTATION_NAMESPACE);
  if (!root) return undefined;
  if (ownerKind === 'master' && !masterSlideNumberEnabled(root)) return undefined;

  const commonSlide = uniqueDirectChild(root, 'cSld', PRESENTATION_NAMESPACE);
  const shapeTree = commonSlide
    ? uniqueDirectChild(commonSlide, 'spTree', PRESENTATION_NAMESPACE)
    : undefined;
  if (!shapeTree) return undefined;

  const candidates = directChildren(shapeTree).filter(
    (child) => isElement(child, 'sp', PRESENTATION_NAMESPACE)
      && containsDirectSlideNumberPlaceholder(child),
  );
  if (candidates.length !== 1) return undefined;
  const shape = candidates[0]!;
  if (!validUniqueShapeId(xml, root, shape)) return undefined;

  const shapeProperties = uniqueDirectChild(shape, 'spPr', PRESENTATION_NAMESPACE);
  const transform = shapeProperties
    ? uniqueDirectChild(shapeProperties, 'xfrm', DRAWING_NAMESPACE)
    : undefined;
  const offset = transform
    ? uniqueDirectChild(transform, 'off', DRAWING_NAMESPACE)
    : undefined;
  const extent = transform
    ? uniqueDirectChild(transform, 'ext', DRAWING_NAMESPACE)
    : undefined;
  if (!offset || !extent) return undefined;
  const x = readCoordinate(offset, 'x');
  const y = readCoordinate(offset, 'y');
  const width = readExtent(extent, 'cx');
  const height = readExtent(extent, 'cy');
  if (x === INVALID || y === INVALID || width === INVALID || height === INVALID) {
    return undefined;
  }

  const textBody = uniqueDirectChild(shape, 'txBody', PRESENTATION_NAMESPACE);
  const bodyProperties = textBody
    ? uniqueDirectChild(textBody, 'bodyPr', DRAWING_NAMESPACE)
    : undefined;
  if (!textBody || !bodyProperties) return undefined;
  const listStyles = directChildren(textBody).filter(
    (child) => isElement(child, 'lstStyle', DRAWING_NAMESPACE),
  );
  if (listStyles.length > 1) return undefined;
  const paragraphs = directChildren(textBody).filter(
    (child) => isElement(child, 'p', DRAWING_NAMESPACE),
  );
  if (paragraphs.length !== 1) return undefined;
  const paragraph = paragraphs[0]!;
  const paragraphProperties = optionalUniqueDirectChild(
    paragraph,
    'pPr',
    DRAWING_NAMESPACE,
  );
  if (paragraphProperties === INVALID) return undefined;
  const field = readUniqueSlideNumberField(paragraph);
  if (!field) return undefined;

  const align = paragraphProperties
    ? readMappedAttribute(paragraphProperties, 'algn', ALIGN_FROM_OOXML, 'left')
    : 'left';
  const rtl = paragraphProperties
    ? readBooleanAttribute(paragraphProperties, 'rtl', false)
    : false;
  const valign = readOptionalMappedAttribute(
    bodyProperties,
    'anchor',
    VALIGN_FROM_OOXML,
  );
  const margin = readMargins(bodyProperties);
  if (align === INVALID || rtl === INVALID || valign === INVALID || margin === INVALID) {
    return undefined;
  }

  const defaultStyle = readListStyle(listStyles[0]);
  const fieldProperties = optionalUniqueDirectChild(
    field,
    'rPr',
    DRAWING_NAMESPACE,
  );
  if (fieldProperties === INVALID) return undefined;
  const fieldStyle = readStyleElement(fieldProperties);
  if (defaultStyle === INVALID || fieldStyle === INVALID) return undefined;
  const style = mergeStyles(defaultStyle, fieldStyle);

  return deepFreeze({
    x,
    y,
    width,
    height,
    align,
    rtl,
    ...(valign === undefined ? {} : { valign }),
    ...(margin === undefined ? {} : { margin }),
    style,
  });
}

function normalizeCoordinate(
  value: unknown,
  fallback: number,
  context: string,
): number {
  if (value === undefined) return fallback;
  const normalized = normalizeFiniteInteger(value, context);
  if (normalized < -MAX_COORDINATE || normalized > MAX_COORDINATE) {
    throw new RangeError(`${context} must fit the DrawingML coordinate range`);
  }
  return normalized;
}

function normalizeExtent(
  value: unknown,
  fallback: number,
  context: string,
): number {
  if (value === undefined) return fallback;
  const normalized = normalizeFiniteInteger(value, context);
  if (normalized <= 0 || normalized > MAX_COORDINATE) {
    throw new RangeError(`${context} must be positive and fit the DrawingML coordinate range`);
  }
  return normalized;
}

function normalizeFiniteInteger(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  return Math.round(value);
}

function normalizeStyle(value: unknown): SlideNumberTextStyle {
  const input = value === undefined
    ? Object.create(null) as Record<string, unknown>
    : readDataObject(value, STYLE_KEYS, 'Slide number style');
  const fontFamily = input.fontFamily === undefined
    ? undefined
    : normalizeXmlString(input.fontFamily, 'Slide number font family');
  const fontSize = input.fontSize === undefined
    ? undefined
    : normalizeFontSize(input.fontSize);
  const lang = input.lang === undefined
    ? 'en-US'
    : normalizeXmlString(input.lang, 'Slide number language');
  const bold = normalizeBoolean(input.bold, false, 'Slide number bold');
  const italic = normalizeBoolean(input.italic, false, 'Slide number italic');
  let color = input.color === undefined ? undefined : normalizeColor(input.color);
  const transparency = input.transparency === undefined
    ? undefined
    : normalizeTransparency(input.transparency);
  if (transparency !== undefined && color === undefined) {
    color = { kind: 'scheme', value: 'tx1' };
  }
  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    lang,
    bold,
    italic,
    ...(color === undefined ? {} : { color }),
    ...(transparency === undefined ? {} : { transparency }),
  };
}

function normalizeFontSize(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Slide number font size must be finite');
  }
  if (value < 1 || value > 4_000) {
    throw new RangeError('Slide number font size must be between 1 and 4000 points');
  }
  return Math.round(value * 100) / 100;
}

function normalizeTransparency(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Slide number transparency must be finite');
  }
  if (value < 0 || value > 100) {
    throw new RangeError('Slide number transparency must be between 0 and 100');
  }
  return Math.round(value * 1_000) / 1_000;
}

function normalizeColor(value: unknown): SlideNumberColor {
  const input = readDataObject(value, COLOR_KEYS, 'Slide number color');
  if (input.kind === 'srgb') {
    if (typeof input.value !== 'string' || !/^[\da-f]{6}$/i.test(input.value)) {
      throw new TypeError('Slide number sRGB color must contain six hex digits');
    }
    return { kind: 'srgb', value: input.value.toUpperCase() };
  }
  if (input.kind === 'scheme') {
    if (typeof input.value !== 'string' || !SCHEME_COLORS.has(input.value)) {
      throw new TypeError('Slide number scheme color is unsupported');
    }
    return { kind: 'scheme', value: input.value };
  }
  throw new TypeError('Slide number color kind must be srgb or scheme');
}

function normalizeMargins(value: unknown): SlideNumberMargins | undefined {
  if (value === undefined) return undefined;
  if (typeof value === 'number') {
    const side = normalizeMarginSide(value, 'Slide number margin');
    return { top: side, right: side, bottom: side, left: side };
  }
  if (Array.isArray(value)) return normalizeMarginTuple(value);
  const input = readDataObject(value, MARGIN_KEYS, 'Slide number margin');
  const margin: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    if (input[side] !== undefined) {
      margin[side] = normalizeMarginSide(input[side], `Slide number margin ${side}`);
    }
  }
  return Object.keys(margin).length === 0 ? undefined : margin;
}

function normalizeMarginTuple(value: unknown[]): SlideNumberMargins {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError('Slide number margin tuple must be an ordinary array');
  }
  const length = Object.getOwnPropertyDescriptor(value, 'length');
  if (!length || !Object.hasOwn(length, 'value') || length.value !== 4) {
    throw new RangeError('Slide number margin tuple must contain exactly four values');
  }
  const allowed = new Set(['0', '1', '2', '3', 'length']);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`Slide number margin tuple contains unsupported property ${String(key)}`);
    }
  }
  const sides = Array.from({ length: 4 }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor) {
      throw new TypeError('Slide number margin tuple must not contain sparse values');
    }
    if (!Object.hasOwn(descriptor, 'value')) {
      throw new TypeError('Slide number margin tuple must contain only data properties');
    }
    return normalizeMarginSide(
      descriptor.value,
      `Slide number margin ${['top', 'right', 'bottom', 'left'][index]}`,
    );
  });
  return { top: sides[0]!, right: sides[1]!, bottom: sides[2]!, left: sides[3]! };
}

function normalizeMarginSide(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const raw = Math.round(value * EMU_PER_POINT);
  if (raw < MIN_INT32 || raw > MAX_INT32) {
    throw new RangeError(`${context} must fit the OOXML signed Int32 coordinate range`);
  }
  return raw / EMU_PER_POINT;
}

function normalizeXmlString(value: unknown, context: string): string {
  if (typeof value !== 'string' || !/\S/u.test(value)) {
    throw new TypeError(`${context} must be a non-whitespace string`);
  }
  if (!isXmlSafe(value)) throw new TypeError(`${context} contains invalid XML characters`);
  return value;
}

function isXmlSafe(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if ((code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d)
      || code === 0xfffe || code === 0xffff) {
      return false;
    }
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next < 0xdc00 || next > 0xdfff) return false;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return false;
    }
  }
  return true;
}

function normalizeBoolean(
  value: unknown,
  fallback: boolean,
  context: string,
): boolean {
  if (value === undefined) return fallback;
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}

function normalizeEnum<const T extends string>(
  value: unknown,
  supported: readonly T[],
  fallback: T | undefined,
  context: string,
): T {
  if (value === undefined && fallback !== undefined) return fallback;
  if (typeof value !== 'string' || !supported.includes(value as T)) {
    throw new TypeError(`${context} must be ${supported.join(', ')}`);
  }
  return value as T;
}

function readDataObject(
  value: unknown,
  supported: ReadonlySet<string>,
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !supported.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function uniqueRoot(
  xml: LosslessXmlDocument,
  localName: string,
  namespace: string,
): XmlElement | undefined {
  const roots = xml.roots.filter((root) => isElement(root, localName, namespace));
  return roots.length === 1 && xml.roots.length === 1 ? roots[0] : undefined;
}

function containsDirectSlideNumberPlaceholder(shape: XmlElement): boolean {
  return directChildren(shape)
    .filter((child) => isElement(child, 'nvSpPr', PRESENTATION_NAMESPACE))
    .some((nonVisual) => directChildren(nonVisual)
      .filter((child) => isElement(child, 'nvPr', PRESENTATION_NAMESPACE))
      .some((application) => directChildren(application)
        .filter((child) => isElement(child, 'ph', PRESENTATION_NAMESPACE))
        .some((placeholder) => directAttribute(placeholder, 'type') === 'sldNum')));
}

function validUniqueShapeId(
  xml: LosslessXmlDocument,
  root: XmlElement,
  shape: XmlElement,
): boolean {
  const nonVisual = uniqueDirectChild(shape, 'nvSpPr', PRESENTATION_NAMESPACE);
  const properties = nonVisual
    ? uniqueDirectChild(nonVisual, 'cNvPr', PRESENTATION_NAMESPACE)
    : undefined;
  const application = nonVisual
    ? uniqueDirectChild(nonVisual, 'nvPr', PRESENTATION_NAMESPACE)
    : undefined;
  const placeholder = application
    ? uniqueDirectChild(application, 'ph', PRESENTATION_NAMESPACE)
    : undefined;
  if (!properties || !placeholder || directAttribute(placeholder, 'type') !== 'sldNum') {
    return false;
  }
  const id = readUnsignedInteger(properties, 'id', 1, MAX_UINT32);
  if (id === INVALID) return false;
  return descendantsAndSelf(root)
    .filter((element) => isElement(element, 'cNvPr', PRESENTATION_NAMESPACE))
    .filter((element) => readUnsignedInteger(element, 'id', 1, MAX_UINT32) === id)
    .length === 1
    && xml.roots.length === 1;
}

function readUniqueSlideNumberField(paragraph: XmlElement): XmlElement | undefined {
  const drawingChildren = directChildren(paragraph).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  const fields = drawingChildren.filter((child) => child.localName === 'fld');
  const slideNumberFields = fields.filter(
    (field) => directAttribute(field, 'type') === 'slidenum',
  );
  if (fields.length !== 1 || slideNumberFields.length !== 1) return undefined;
  if (drawingChildren.some((child) => !['pPr', 'fld', 'endParaRPr'].includes(child.localName))) {
    return undefined;
  }
  const field = slideNumberFields[0]!;
  const fieldChildren = directChildren(field).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  if (fieldChildren.some((child) => !['rPr', 't'].includes(child.localName))) {
    return undefined;
  }
  const text = fieldChildren.filter((child) => child.localName === 't');
  return text.length === 1 && directChildren(text[0]!).length === 0
    ? field
    : undefined;
}

function readListStyle(
  listStyle: XmlElement | undefined,
): PartialSlideNumberStyle | typeof INVALID {
  if (!listStyle) return {};
  const levels = directChildren(listStyle).filter(
    (child) => isElement(child, 'lvl1pPr', DRAWING_NAMESPACE),
  );
  if (levels.length > 1) return INVALID;
  if (levels.length === 0) return {};
  const defaults = directChildren(levels[0]!).filter(
    (child) => isElement(child, 'defRPr', DRAWING_NAMESPACE),
  );
  if (defaults.length > 1) return INVALID;
  return readStyleElement(defaults[0]);
}

function readStyleElement(
  properties: XmlElement | undefined,
): PartialSlideNumberStyle | typeof INVALID {
  if (!properties) return {};
  const size = readOptionalInteger(properties, 'sz', 100, 400_000);
  const lang = readOptionalXmlString(properties, 'lang');
  const bold = readOptionalBooleanAttribute(properties, 'b');
  const italic = readOptionalBooleanAttribute(properties, 'i');
  const fontFamily = readFontFamily(properties);
  const fill = readStyleColor(properties);
  if (size === INVALID || lang === INVALID || bold === INVALID
    || italic === INVALID || fontFamily === INVALID || fill === INVALID) {
    return INVALID;
  }
  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(size === undefined ? {} : { fontSize: size / 100 }),
    ...(lang === undefined ? {} : { lang }),
    ...(bold === undefined ? {} : { bold }),
    ...(italic === undefined ? {} : { italic }),
    ...(fill?.color === undefined ? {} : { color: fill.color }),
    ...(fill?.transparency === undefined ? {} : { transparency: fill.transparency }),
  };
}

function readFontFamily(properties: XmlElement): string | undefined | typeof INVALID {
  const families: string[] = [];
  for (const name of ['latin', 'ea', 'cs']) {
    const elements = directChildren(properties).filter(
      (child) => isElement(child, name, DRAWING_NAMESPACE),
    );
    if (elements.length > 1) return INVALID;
    if (elements[0]) {
      const typeface = directAttribute(elements[0], 'typeface');
      if (typeface === INVALID || typeface === undefined || !isValidXmlString(typeface)) {
        return INVALID;
      }
      families.push(typeface);
    }
  }
  const unique = [...new Set(families)];
  return unique.length > 1 ? INVALID : unique[0];
}

function readStyleColor(
  properties: XmlElement,
): { readonly color: SlideNumberColor; readonly transparency?: number }
  | undefined
  | typeof INVALID {
  const fillChoices = directChildren(properties).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE
      && ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(child.localName),
  );
  if (fillChoices.length === 0) return undefined;
  if (fillChoices.length !== 1 || fillChoices[0]?.localName !== 'solidFill') return INVALID;
  const colorElements = directChildren(fillChoices[0]).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  if (colorElements.length !== 1) return INVALID;
  const element = colorElements[0]!;
  const raw = directAttribute(element, 'val');
  let color: SlideNumberColor;
  if (element.localName === 'srgbClr') {
    if (typeof raw !== 'string' || !/^[\da-f]{6}$/i.test(raw)) return INVALID;
    color = { kind: 'srgb', value: raw.toUpperCase() };
  } else if (element.localName === 'schemeClr') {
    if (typeof raw !== 'string' || !SCHEME_COLORS.has(raw)) return INVALID;
    color = { kind: 'scheme', value: raw };
  } else {
    return INVALID;
  }
  const transforms = directChildren(element).filter(
    (child) => elementNamespaceUri(child) === DRAWING_NAMESPACE,
  );
  if (transforms.some((child) => child.localName !== 'alpha')) return INVALID;
  if (transforms.length > 1) return INVALID;
  if (transforms[0] && directChildren(transforms[0]).length > 0) return INVALID;
  const alpha = transforms[0]
    ? readUnsignedInteger(transforms[0], 'val', 0, 100_000)
    : undefined;
  if (alpha === INVALID) return INVALID;
  return {
    color,
    ...(alpha === undefined ? {} : { transparency: 100 - alpha / 1_000 }),
  };
}

function mergeStyles(
  defaults: PartialSlideNumberStyle,
  direct: PartialSlideNumberStyle,
): SlideNumberTextStyle {
  const color = direct.color ?? defaults.color;
  const transparency = direct.color === undefined
    ? defaults.transparency
    : direct.transparency;
  return {
    ...(direct.fontFamily === undefined && defaults.fontFamily === undefined
      ? {}
      : { fontFamily: direct.fontFamily ?? defaults.fontFamily }),
    ...(direct.fontSize === undefined && defaults.fontSize === undefined
      ? {}
      : { fontSize: direct.fontSize ?? defaults.fontSize }),
    lang: direct.lang ?? defaults.lang ?? 'en-US',
    bold: direct.bold ?? defaults.bold ?? false,
    italic: direct.italic ?? defaults.italic ?? false,
    ...(color === undefined ? {} : { color }),
    ...(transparency === undefined ? {} : { transparency }),
  };
}

function readMargins(
  element: XmlElement,
): SlideNumberMargins | undefined | typeof INVALID {
  const mapping = [
    ['top', 'tIns'],
    ['right', 'rIns'],
    ['bottom', 'bIns'],
    ['left', 'lIns'],
  ] as const;
  const result: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const [side, attribute] of mapping) {
    const raw = readOptionalSignedInteger(element, attribute, MIN_INT32, MAX_INT32);
    if (raw === INVALID) return INVALID;
    if (raw !== undefined) result[side] = raw / EMU_PER_POINT;
  }
  return Object.keys(result).length === 0 ? undefined : result;
}

function readCoordinate(element: XmlElement, name: string): number | typeof INVALID {
  return readSignedInteger(element, name, -MAX_COORDINATE, MAX_COORDINATE);
}

function readExtent(element: XmlElement, name: string): number | typeof INVALID {
  return readUnsignedInteger(element, name, 1, MAX_COORDINATE);
}

function readSignedInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | typeof INVALID {
  const value = directAttribute(element, name);
  if (typeof value !== 'string' || !/^-?\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readUnsignedInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | typeof INVALID {
  const value = directAttribute(element, name);
  if (typeof value !== 'string' || !/^\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readOptionalSignedInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  if (value === undefined) return undefined;
  if (!/^-?\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readOptionalInteger(
  element: XmlElement,
  name: string,
  minimum: number,
  maximum: number,
): number | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  if (value === undefined) return undefined;
  if (!/^\d+$/.test(value)) return INVALID;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : INVALID;
}

function readOptionalXmlString(
  element: XmlElement,
  name: string,
): string | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined || isValidXmlString(value) ? value : INVALID;
}

function isValidXmlString(value: string): boolean {
  return /\S/u.test(value) && isXmlSafe(value);
}

function readMappedAttribute<const T>(
  element: XmlElement,
  name: string,
  mapping: ReadonlyMap<string, T>,
  fallback: T,
): T | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? fallback : mapping.get(value) ?? INVALID;
}

function readOptionalMappedAttribute<const T>(
  element: XmlElement,
  name: string,
  mapping: ReadonlyMap<string, T>,
): T | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? undefined : mapping.get(value) ?? INVALID;
}

function readBooleanAttribute(
  element: XmlElement,
  name: string,
  fallback: boolean,
): boolean | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? fallback : parseBoolean(value);
}

function readOptionalBooleanAttribute(
  element: XmlElement,
  name: string,
): boolean | undefined | typeof INVALID {
  const value = directAttribute(element, name);
  if (value === INVALID) return INVALID;
  return value === undefined ? undefined : parseBoolean(value);
}

function parseBoolean(value: string): boolean | typeof INVALID {
  if (['1', 'true', 'on'].includes(value)) return true;
  if (['0', 'false', 'off'].includes(value)) return false;
  return INVALID;
}

function masterSlideNumberEnabled(root: XmlElement): boolean {
  const headers = directChildren(root).filter(
    (child) => isElement(child, 'hf', PRESENTATION_NAMESPACE),
  );
  if (headers.length > 1) return false;
  if (headers.length === 0) return true;
  const value = directAttribute(headers[0]!, 'sldNum');
  if (value === INVALID) return false;
  if (value === undefined) return true;
  return parseBoolean(value) === true;
}

function directAttribute(
  element: XmlElement,
  name: string,
): string | undefined | typeof INVALID {
  const matches = element.attributes.filter((attribute) => attribute.name === name);
  if (matches.length > 1) return INVALID;
  return matches[0]?.value;
}

function uniqueDirectChild(
  parent: XmlElement,
  localName: string,
  namespace: string,
): XmlElement | undefined {
  const matches = directChildren(parent).filter(
    (child) => isElement(child, localName, namespace),
  );
  return matches.length === 1 ? matches[0] : undefined;
}

function optionalUniqueDirectChild(
  parent: XmlElement,
  localName: string,
  namespace: string,
): XmlElement | undefined | typeof INVALID {
  const matches = directChildren(parent).filter(
    (child) => isElement(child, localName, namespace),
  );
  if (matches.length > 1) return INVALID;
  return matches[0];
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function descendantsAndSelf(element: XmlElement): XmlElement[] {
  const result = [element];
  for (const child of directChildren(element)) result.push(...descendantsAndSelf(child));
  return result;
}

function isElement(
  element: XmlElement,
  localName: string,
  namespace: string,
): boolean {
  return element.localName === localName && elementNamespaceUri(element) === namespace;
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  let current: XmlElement | undefined = element;
  while (current) {
    const declarations = current.attributes.filter(
      (attribute) => attribute.name === declarationName,
    );
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
    current = current.parent;
  }
  return undefined;
}

function deepFreeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) deepFreeze(child);
    Object.freeze(value);
  }
  return value;
}
