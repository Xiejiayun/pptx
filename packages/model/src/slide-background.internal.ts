import {
  GradientCodec,
  type GradientFill,
  type GradientStop,
  type OoxmlColor,
  type OoxmlColorSource,
} from '@pptx/codecs';
import {
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import type { OpcPackage, Relationship } from '@pptx/opc';
import type { RasterImageContentType } from './image.js';
import {
  normalizeSimpleFill,
  type SimpleFill,
} from './simple-fill.internal.js';
import type {
  SlideBackground,
  SlideBackgroundImage,
} from './slide-background.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const IMAGE_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/image`;

const BACKGROUND_KEYS = [
  'kind',
  'color',
  'transparency',
  'contentType',
  'bytes',
  'angle',
  'scaled',
  'rotateWithShape',
  'flip',
  'stops',
  'path',
  'fillRectangle',
] as const;
const SIMPLE_KEYS = ['kind', 'color', 'transparency'] as const;
const IMAGE_KEYS = ['kind', 'contentType', 'bytes'] as const;
const LINEAR_KEYS = [
  'kind',
  'angle',
  'scaled',
  'rotateWithShape',
  'flip',
  'stops',
] as const;
const PATH_KEYS = [
  'kind',
  'path',
  'rotateWithShape',
  'fillRectangle',
  'stops',
] as const;
const STOP_KEYS = ['offset', 'color', 'alpha'] as const;
const COLOR_KEYS = ['source', 'alpha', 'transforms'] as const;
const TRANSFORM_KEYS = ['kind', 'value'] as const;
const RECTANGLE_KEYS = ['left', 'top', 'right', 'bottom'] as const;
const COLOR_TRANSFORMS = new Set([
  'alpha',
  'alphaMod',
  'alphaModFix',
  'alphaOff',
  'blue',
  'blueMod',
  'blueOff',
  'comp',
  'gamma',
  'gray',
  'green',
  'greenMod',
  'greenOff',
  'hue',
  'hueMod',
  'hueOff',
  'inv',
  'invGamma',
  'lum',
  'lumMod',
  'lumOff',
  'red',
  'redMod',
  'redOff',
  'sat',
  'satMod',
  'satOff',
  'shade',
  'tint',
]);
const FILL_NAMES = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

export function normalizeSlideBackground(value: unknown): SlideBackground | undefined {
  if (value === undefined) return undefined;
  const input = readDataObject(value, BACKGROUND_KEYS, 'Slide background');
  switch (input.kind) {
    case 'none':
    case 'solid':
      return freezeSimpleFill(
        normalizeSimpleFill(input, 'Slide background')!,
      );
    case 'image':
      assertKeys(input, IMAGE_KEYS, 'Slide background image');
      return normalizeImage(input);
    case 'linear-gradient':
      assertKeys(input, LINEAR_KEYS, 'Slide linear gradient');
      return normalizeLinearGradient(input);
    case 'path-gradient':
      assertKeys(input, PATH_KEYS, 'Slide path gradient');
      return normalizePathGradient(input);
    default:
      throw new TypeError(
        'Slide background kind must be none, solid, image, linear-gradient, or path-gradient',
      );
  }
}

export function readSlideBackground(
  pkg: OpcPackage,
  slidePartUri: string,
): SlideBackground | undefined {
  const part = pkg.requirePart(slidePartUri);
  const xml = LosslessXmlDocument.parse(part.bytes);
  const roots = xml.roots.filter((root) =>
    isElement(root, 'sld', PRESENTATION_NAMESPACE));
  if (roots.length !== 1) return undefined;
  const slide = roots[0]!;
  const commonSlides = directChildren(slide).filter((element) =>
    isElement(element, 'cSld', PRESENTATION_NAMESPACE));
  if (commonSlides.length !== 1) return undefined;
  const backgrounds = directChildren(commonSlides[0]!).filter((element) =>
    isElement(element, 'bg', PRESENTATION_NAMESPACE));
  if (backgrounds.length === 0) return undefined;
  if (backgrounds.length !== 1) return undefined;
  const backgroundChildren = directChildren(backgrounds[0]!);
  const properties = backgroundChildren.filter((element) =>
    isElement(element, 'bgPr', PRESENTATION_NAMESPACE));
  const references = backgroundChildren.filter((element) =>
    isElement(element, 'bgRef', PRESENTATION_NAMESPACE));
  if (properties.length !== 1 || references.length > 0) return undefined;
  const choices = directChildren(properties[0]!).filter((element) =>
    elementNamespaceUri(element) === DRAWING_NAMESPACE
    && FILL_NAMES.has(element.localName));
  if (choices.length !== 1) return undefined;
  const choice = choices[0]!;

  if (choice.localName === 'noFill' || choice.localName === 'solidFill') {
    return readBackgroundSimpleFill(choice);
  }
  if (choice.localName === 'gradFill') {
    if (!isSafeGradient(choice)) return undefined;
    try {
      const decoded = new GradientCodec().decode(choice, xml);
      const rotateWithShape = booleanAttribute(choice, 'rotWithShape', true);
      const normalized = decoded.kind === 'linear-gradient'
        ? {
            ...decoded,
            rotateWithShape,
            scaled: booleanAttribute(
              directChildren(choice).find((element) =>
                isElement(element, 'lin', DRAWING_NAMESPACE))!,
              'scaled',
              true,
            ),
          }
        : { ...decoded, rotateWithShape };
      return normalizeSlideBackground(normalized);
    } catch {
      return undefined;
    }
  }
  if (choice.localName === 'blipFill') {
    return readImageBackground(pkg, slidePartUri, choice);
  }
  return undefined;
}

function normalizeImage(input: Record<string, unknown>): SlideBackgroundImage {
  if (
    input.contentType !== 'image/png'
    && input.contentType !== 'image/jpeg'
    && input.contentType !== 'image/gif'
  ) {
    throw new TypeError('Slide background image contentType is unsupported');
  }
  if (!(input.bytes instanceof Uint8Array)) {
    throw new TypeError('Slide background image bytes must be a Uint8Array');
  }
  if (input.bytes.length === 0) {
    throw new RangeError('Slide background image bytes must not be empty');
  }
  return Object.freeze({
    kind: 'image',
    contentType: input.contentType,
    bytes: new Uint8Array(input.bytes),
  });
}

function normalizeLinearGradient(input: Record<string, unknown>): GradientFill {
  const angle = finiteNumber(input.angle, 'Slide linear gradient angle');
  const scaled = optionalBoolean(input.scaled, 'Slide linear gradient scaled');
  const rotateWithShape = optionalBoolean(
    input.rotateWithShape,
    'Slide linear gradient rotateWithShape',
  );
  if (
    input.flip !== undefined
    && input.flip !== 'none'
    && input.flip !== 'x'
    && input.flip !== 'y'
    && input.flip !== 'xy'
  ) throw new TypeError('Slide linear gradient flip is unsupported');
  return Object.freeze({
    kind: 'linear-gradient',
    angle,
    ...(scaled === undefined ? {} : { scaled }),
    ...(rotateWithShape === undefined ? {} : { rotateWithShape }),
    ...(input.flip === undefined ? {} : { flip: input.flip }),
    stops: normalizeStops(input.stops),
  });
}

function normalizePathGradient(input: Record<string, unknown>): GradientFill {
  if (input.path !== 'circle' && input.path !== 'rect' && input.path !== 'shape') {
    throw new TypeError('Slide path gradient path is unsupported');
  }
  const rotateWithShape = optionalBoolean(
    input.rotateWithShape,
    'Slide path gradient rotateWithShape',
  );
  const fillRectangle = input.fillRectangle === undefined
    ? undefined
    : normalizeFillRectangle(input.fillRectangle);
  return Object.freeze({
    kind: 'path-gradient',
    path: input.path,
    ...(rotateWithShape === undefined ? {} : { rotateWithShape }),
    ...(fillRectangle === undefined ? {} : { fillRectangle }),
    stops: normalizeStops(input.stops),
  });
}

function normalizeStops(value: unknown): readonly GradientStop[] {
  if (!Array.isArray(value)) {
    throw new TypeError('Slide background gradient stops must be an array');
  }
  if (value.length < 2) {
    throw new RangeError('Slide background gradient requires at least two stops');
  }
  assertDenseArray(value, 'Slide background gradient stops');
  const stops = value.map((entry, index) => {
    const input = readDataObject(
      entry,
      STOP_KEYS,
      `Slide background gradient stop ${index}`,
    );
    const offset = rangedNumber(
      input.offset,
      `Slide background gradient stop ${index} offset`,
      0,
      1,
    );
    const alpha = input.alpha === undefined
      ? undefined
      : rangedNumber(
          input.alpha,
          `Slide background gradient stop ${index} alpha`,
          0,
          1,
        );
    const color = normalizeGradientColor(
      input.color,
      `Slide background gradient stop ${index} color`,
    );
    return Object.freeze({
      offset,
      color,
      ...(alpha === undefined ? {} : { alpha }),
    });
  });
  return Object.freeze(stops);
}

function normalizeGradientColor(value: unknown, context: string): string | OoxmlColor {
  if (typeof value === 'string') {
    if (!/^#?[\da-f]{6}$/i.test(value)) {
      throw new TypeError(`${context} must contain six hex digits`);
    }
    return value.replace(/^#/, '').toUpperCase();
  }
  const input = readDataObject(value, COLOR_KEYS, context);
  const source = normalizeColorSource(input.source, `${context} source`);
  const alpha = rangedNumber(input.alpha, `${context} alpha`, 0, 1);
  if (!Array.isArray(input.transforms)) {
    throw new TypeError(`${context} transforms must be an array`);
  }
  assertDenseArray(input.transforms, `${context} transforms`);
  const transforms = input.transforms.map((entry, index) => {
    const transform = readDataObject(
      entry,
      TRANSFORM_KEYS,
      `${context} transform ${index}`,
    );
    if (typeof transform.kind !== 'string' || !COLOR_TRANSFORMS.has(transform.kind)) {
      throw new TypeError(`${context} transform ${index} kind is unsupported`);
    }
    return Object.freeze({
      kind: transform.kind,
      value: finiteNumber(transform.value, `${context} transform ${index} value`),
    });
  });
  return Object.freeze({
    source,
    alpha,
    transforms: Object.freeze(transforms),
  });
}

function normalizeColorSource(value: unknown, context: string): OoxmlColorSource {
  const input = readDataObject(
    value,
    ['kind', 'value', 'lastColor', 'red', 'green', 'blue'],
    context,
  );
  switch (input.kind) {
    case 'srgb':
      assertKeys(input, ['kind', 'value'], context);
      return Object.freeze({
        kind: 'srgb',
        value: hexColor(input.value, `${context} value`),
      });
    case 'scrgb':
      assertKeys(input, ['kind', 'red', 'green', 'blue'], context);
      return Object.freeze({
        kind: 'scrgb',
        red: rangedNumber(input.red, `${context} red`, 0, 1),
        green: rangedNumber(input.green, `${context} green`, 0, 1),
        blue: rangedNumber(input.blue, `${context} blue`, 0, 1),
      });
    case 'scheme':
    case 'preset':
      assertKeys(input, ['kind', 'value'], context);
      return Object.freeze({
        kind: input.kind,
        value: nonEmptyString(input.value, `${context} value`),
      });
    case 'system': {
      assertKeys(input, ['kind', 'value', 'lastColor'], context);
      const lastColor = input.lastColor === undefined
        ? undefined
        : hexColor(input.lastColor, `${context} lastColor`);
      return Object.freeze({
        kind: 'system',
        value: nonEmptyString(input.value, `${context} value`),
        ...(lastColor === undefined ? {} : { lastColor }),
      });
    }
    default:
      throw new TypeError(`${context} kind is unsupported`);
  }
}

function normalizeFillRectangle(value: unknown) {
  const input = readDataObject(
    value,
    RECTANGLE_KEYS,
    'Slide path gradient fillRectangle',
  );
  return Object.freeze({
    left: finiteNumber(input.left, 'Slide path gradient fillRectangle left'),
    top: finiteNumber(input.top, 'Slide path gradient fillRectangle top'),
    right: finiteNumber(input.right, 'Slide path gradient fillRectangle right'),
    bottom: finiteNumber(input.bottom, 'Slide path gradient fillRectangle bottom'),
  });
}

function readImageBackground(
  pkg: OpcPackage,
  slidePartUri: string,
  choice: XmlElement,
): SlideBackgroundImage | undefined {
  const blips = directChildren(choice).filter((element) =>
    isElement(element, 'blip', DRAWING_NAMESPACE));
  if (blips.length !== 1) return undefined;
  const blip = blips[0]!;
  const embedAttributes = blip.attributes.filter((attribute) =>
    attribute.localName === 'embed');
  const linkAttributes = blip.attributes.filter((attribute) =>
    attribute.localName === 'link');
  if (
    embedAttributes.length !== 1
    || attributeNamespaceUri(blip, embedAttributes[0]!) !== RELATIONSHIP_NAMESPACE
    || linkAttributes.length > 0
  ) return undefined;
  const relationships = pkg.relationships(slidePartUri).filter((relationship) =>
    relationship.id === embedAttributes[0]!.value);
  if (relationships.length !== 1) return undefined;
  const relationship = relationships[0]!;
  if (!isInternalImageRelationship(pkg, relationship)) return undefined;
  const part = pkg.requirePart(relationship.resolvedTarget);
  const contentType = rasterContentType(part.contentType);
  if (!contentType) return undefined;
  return Object.freeze({
    kind: 'image',
    contentType,
    bytes: new Uint8Array(part.bytes),
  });
}

function isInternalImageRelationship(
  pkg: OpcPackage,
  relationship: Relationship,
): relationship is Relationship & { readonly resolvedTarget: string } {
  return relationship.type === IMAGE_RELATIONSHIP
    && relationship.targetMode === 'Internal'
    && typeof relationship.resolvedTarget === 'string'
    && pkg.hasPart(relationship.resolvedTarget);
}

function readBackgroundSimpleFill(choice: XmlElement): SimpleFill | undefined {
  if (nonNamespaceAttributes(choice).length > 0) return undefined;
  if (choice.localName === 'noFill') {
    return directChildren(choice).length === 0
      ? Object.freeze({ kind: 'none' })
      : undefined;
  }
  const colors = directChildren(choice);
  if (colors.length !== 1 || elementNamespaceUri(colors[0]!) !== DRAWING_NAMESPACE) {
    return undefined;
  }
  const color = colors[0]!;
  if (color.localName !== 'srgbClr' && color.localName !== 'schemeClr') {
    return undefined;
  }
  const attributes = nonNamespaceAttributes(color);
  if (attributes.length !== 1 || attributes[0]?.name !== 'val') return undefined;
  const transforms = directChildren(color);
  let transparency: number | undefined;
  if (transforms.length > 0) {
    if (transforms.length !== 1 || !isElement(transforms[0]!, 'alpha', DRAWING_NAMESPACE)) {
      return undefined;
    }
    const alpha = transforms[0]!;
    const alphaAttributes = nonNamespaceAttributes(alpha);
    if (
      directChildren(alpha).length > 0
      || alphaAttributes.length !== 1
      || alphaAttributes[0]?.name !== 'val'
      || !/^\d+$/.test(alphaAttributes[0].value)
    ) return undefined;
    const value = Number(alphaAttributes[0].value);
    if (!Number.isSafeInteger(value) || value < 0 || value > 100_000) return undefined;
    transparency = 100 - value / 1_000;
  }
  try {
    const fill = normalizeSimpleFill({
      kind: 'solid',
      color: {
        kind: color.localName === 'srgbClr' ? 'srgb' : 'scheme',
        value: attributes[0].value,
      },
      ...(transparency === undefined ? {} : { transparency }),
    }, 'Slide background');
    return fill ? freezeSimpleFill(fill) : undefined;
  } catch {
    return undefined;
  }
}

function rasterContentType(value: string): RasterImageContentType | undefined {
  return value === 'image/png' || value === 'image/jpeg' || value === 'image/gif'
    ? value
    : undefined;
}

function isSafeGradient(gradient: XmlElement): boolean {
  const attributes = nonNamespaceAttributes(gradient);
  if (
    attributes.some(({ name }) => name !== 'rotWithShape' && name !== 'flip')
    || unqualifiedAttributes(gradient, 'rotWithShape').length > 1
    || unqualifiedAttributes(gradient, 'flip').length > 1
  ) return false;
  const rotateWithShape = unqualifiedAttributes(gradient, 'rotWithShape')[0];
  const flip = unqualifiedAttributes(gradient, 'flip')[0];
  if (
    (rotateWithShape && !isBooleanToken(rotateWithShape.value))
    || (flip && !['none', 'x', 'y', 'xy'].includes(flip.value))
  ) return false;
  const children = directChildren(gradient);
  const stopLists = children.filter((element) =>
    isElement(element, 'gsLst', DRAWING_NAMESPACE));
  const linear = children.filter((element) =>
    isElement(element, 'lin', DRAWING_NAMESPACE));
  const paths = children.filter((element) =>
    isElement(element, 'path', DRAWING_NAMESPACE));
  if (
    stopLists.length !== 1
    || linear.length + paths.length !== 1
    || children.some((element) =>
      !['gsLst', 'lin', 'path'].includes(element.localName)
      || elementNamespaceUri(element) !== DRAWING_NAMESPACE)
  ) return false;
  const stops = directChildren(stopLists[0]!);
  if (
    stops.length < 2
    || stops.some((stop) => !isElement(stop, 'gs', DRAWING_NAMESPACE))
  ) return false;
  for (const stop of stops) {
    const position = unqualifiedAttributes(stop, 'pos');
    if (
      position.length !== 1
      || !/^\d+$/.test(position[0]!.value)
      || Number(position[0]!.value) > 100_000
    ) return false;
    const colors = directChildren(stop).filter((element) =>
      elementNamespaceUri(element) === DRAWING_NAMESPACE
      && ['srgbClr', 'scrgbClr', 'schemeClr', 'sysClr', 'prstClr'].includes(element.localName));
    if (
      colors.length !== 1
      || directChildren(stop).length !== 1
      || !isSafeColorElement(colors[0]!)
    ) return false;
  }
  if (linear.length === 1) {
    const linearAttributes = nonNamespaceAttributes(linear[0]!);
    if (linearAttributes.some(({ name }) => name !== 'ang' && name !== 'scaled')) {
      return false;
    }
    const angle = unqualifiedAttributes(linear[0]!, 'ang');
    const scaled = unqualifiedAttributes(linear[0]!, 'scaled');
    if (
      angle.length > 1
      || (angle[0] && !/^-?\d+$/.test(angle[0].value))
      || scaled.length > 1
      || (scaled[0] && !isBooleanToken(scaled[0].value))
      || directChildren(linear[0]!).length > 0
    ) return false;
  }
  if (paths.length === 1) {
    const pathAttributes = nonNamespaceAttributes(paths[0]!);
    if (pathAttributes.some(({ name }) => name !== 'path')) return false;
    const path = unqualifiedAttributes(paths[0]!, 'path');
    if (
      path.length !== 1
      || !['circle', 'rect', 'shape'].includes(path[0]!.value)
    ) return false;
    const rectangles = directChildren(paths[0]!).filter((element) =>
      isElement(element, 'fillToRect', DRAWING_NAMESPACE));
    if (rectangles.length > 1) return false;
  }
  return true;
}

function isSafeColorElement(color: XmlElement): boolean {
  const attributes = color.attributes.filter(({ name }) =>
    name !== 'xmlns' && !name.startsWith('xmlns:'));
  switch (color.localName) {
    case 'srgbClr':
      if (
        attributes.length !== 1
        || attributes[0]?.name !== 'val'
        || !/^[\da-f]{6}$/i.test(attributes[0].value)
      ) return false;
      break;
    case 'scrgbClr': {
      if (attributes.length !== 3) return false;
      for (const name of ['r', 'g', 'b']) {
        const values = unqualifiedAttributes(color, name);
        if (
          values.length !== 1
          || !/^\d+$/.test(values[0]!.value)
          || Number(values[0]!.value) > 100_000
        ) return false;
      }
      break;
    }
    case 'schemeClr':
    case 'prstClr': {
      if (
        attributes.length !== 1
        || attributes[0]?.name !== 'val'
        || attributes[0].value.length === 0
      ) return false;
      break;
    }
    case 'sysClr': {
      if (attributes.length < 1 || attributes.length > 2) return false;
      const values = unqualifiedAttributes(color, 'val');
      const lastColors = unqualifiedAttributes(color, 'lastClr');
      if (
        values.length !== 1
        || values[0]!.value.length === 0
        || lastColors.length > 1
        || (lastColors[0] !== undefined && !/^[\da-f]{6}$/i.test(lastColors[0].value))
      ) return false;
      break;
    }
  }
  for (const transform of directChildren(color)) {
    if (
      elementNamespaceUri(transform) !== DRAWING_NAMESPACE
      || !COLOR_TRANSFORMS.has(transform.localName)
      || directChildren(transform).length > 0
    ) return false;
    const values = unqualifiedAttributes(transform, 'val');
    const transformAttributes = transform.attributes.filter(({ name }) =>
      name !== 'xmlns' && !name.startsWith('xmlns:'));
    if (
      values.length !== 1
      || transformAttributes.length !== 1
      || !/^-?\d+$/.test(values[0]!.value)
    ) return false;
  }
  return true;
}

function freezeSimpleFill(fill: SimpleFill): SimpleFill {
  if (fill.kind === 'none') return Object.freeze({ kind: 'none' });
  return Object.freeze({
    kind: 'solid',
    color: Object.freeze({ ...fill.color }),
    ...(fill.transparency === undefined
      ? {}
      : { transparency: fill.transparency }),
  });
}

function readDataObject(
  value: unknown,
  supported: readonly string[],
  context: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const allowed = new Set(supported);
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
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

function assertKeys(
  value: Record<string, unknown>,
  supported: readonly string[],
  context: string,
): void {
  const allowed = new Set(supported);
  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${key}`);
    }
  }
}

function optionalBoolean(value: unknown, context: string): boolean | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}

function finiteNumber(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  return Object.is(value, -0) ? 0 : value;
}

function rangedNumber(
  value: unknown,
  context: string,
  minimum: number,
  maximum: number,
): number {
  const result = finiteNumber(value, context);
  if (result < minimum || result > maximum) {
    throw new RangeError(`${context} must be between ${minimum} and ${maximum}`);
  }
  return result;
}

function assertDenseArray(value: readonly unknown[], context: string): void {
  for (let index = 0; index < value.length; index += 1) {
    if (!Object.hasOwn(value, index)) {
      throw new TypeError(`${context} must not be sparse`);
    }
  }
}

function nonEmptyString(value: unknown, context: string): string {
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`${context} must be a non-empty string`);
  }
  return value;
}

function hexColor(value: unknown, context: string): string {
  if (typeof value !== 'string' || !/^#?[\da-f]{6}$/i.test(value)) {
    throw new TypeError(`${context} must contain six hex digits`);
  }
  return value.replace(/^#/, '').toUpperCase();
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}

function isElement(
  element: XmlElement,
  localName: string,
  namespace: string,
): boolean {
  return element.localName === localName
    && elementNamespaceUri(element) === namespace;
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function attributeNamespaceUri(
  element: XmlElement,
  attribute: XmlAttribute,
): string | undefined {
  const prefix = lexicalPrefix(attribute.name);
  if (prefix.length === 0) return undefined;
  return namespaceUriForPrefix(element, prefix);
}

function namespaceUriForPrefix(
  element: XmlElement,
  prefix: string,
): string | undefined {
  const declarationName = prefix.length === 0 ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const declaration = current.attributes.find(({ name }) => name === declarationName);
    if (declaration) return declaration.value.length === 0 ? undefined : declaration.value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function unqualifiedAttributes(
  element: XmlElement,
  name: string,
): XmlAttribute[] {
  return element.attributes.filter((attribute) =>
    attribute.name === name && attribute.localName === name);
}

function nonNamespaceAttributes(element: XmlElement): XmlAttribute[] {
  return element.attributes.filter(({ name }) =>
    name !== 'xmlns' && !name.startsWith('xmlns:'));
}

function isBooleanToken(value: string): boolean {
  return value === '0' || value === '1' || value === 'false' || value === 'true';
}

function booleanAttribute(
  element: XmlElement,
  name: string,
  defaultValue: boolean,
): boolean {
  const value = unqualifiedAttributes(element, name)[0]?.value;
  return value === undefined ? defaultValue : value === '1' || value === 'true';
}
