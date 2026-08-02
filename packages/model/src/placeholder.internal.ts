import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import type { OpcPackage } from '@pptx/opc';
import { ModelParseError } from './errors.js';
import {
  PLACEHOLDER_TYPES,
  type PlaceholderIdentity,
  type PlaceholderSelector,
  type PlaceholderType,
} from './placeholder.js';
import type { Transform } from './units.js';
import { readTextShapeIsTextBox } from './text-shape-is-text-box.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX_PLACEHOLDER_INDEX = 4_294_967_294;
const MAX_SHAPE_ID = 4_294_967_295;
const PLACEHOLDER_TYPE_SET = new Set<string>(PLACEHOLDER_TYPES);
const SLIDE_LAYOUT_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';

type PlaceholderState =
  | { readonly kind: 'none' | 'unsafe' }
  | { readonly kind: 'unsupported'; readonly type: string }
  | {
      readonly kind: 'supported';
      readonly identity: Readonly<PlaceholderIdentity>;
      readonly nonVisual: XmlElement;
    };

interface MaterializedPlaceholderDescriptor {
  readonly name: string;
  readonly shapeId: number;
  readonly identity: Readonly<PlaceholderIdentity>;
  readonly element: XmlElement;
  readonly transform: Transform;
  readonly isTextBox: boolean;
  readonly namespaceAttributes: string;
  readonly transformXml?: string;
  readonly bodyPropertiesXml?: string;
  readonly listStyleXml?: string;
}

export type PlaceholderDomain = 'text-shape' | 'image' | 'chart' | 'table' | 'media';

export interface ResolvedPlaceholderOwner {
  readonly identity: Readonly<PlaceholderIdentity>;
  readonly name: string;
  readonly shapeId: number;
  readonly transform: Transform;
  readonly slideElement: XmlElement;
  readonly layoutElement: XmlElement;
}

export function normalizePlaceholderIdentity(
  value: unknown,
  context = 'Placeholder identity',
): Readonly<PlaceholderIdentity> {
  const input = readDataObject(value, context, ['type', 'index']);
  if (!Object.hasOwn(input, 'type') || !Object.hasOwn(input, 'index')) {
    throw new TypeError(`${context} requires type and index`);
  }
  if (typeof input.type !== 'string' || !PLACEHOLDER_TYPE_SET.has(input.type)) {
    throw new TypeError(`${context} type must be ${PLACEHOLDER_TYPES.join(', ')}`);
  }
  if (typeof input.index !== 'number' || !Number.isSafeInteger(input.index)) {
    throw new TypeError(`${context} index must be an integer`);
  }
  if (input.index < 0 || input.index > MAX_PLACEHOLDER_INDEX) {
    throw new RangeError(`${context} index must be between 0 and ${MAX_PLACEHOLDER_INDEX}`);
  }
  return Object.freeze({
    type: input.type as PlaceholderType,
    index: input.index,
  });
}

export function normalizePlaceholderSelector(
  value: unknown,
): PlaceholderSelector {
  if (typeof value === 'string') {
    if (value.length === 0) throw new TypeError('Placeholder selector name must not be empty');
    if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) {
      throw new TypeError('Placeholder selector name contains invalid XML characters');
    }
    return value;
  }
  return normalizePlaceholderIdentity(value, 'Placeholder selector');
}

export function readShapePlaceholder(
  xml: LosslessXmlDocument,
  shape: XmlElement,
): Readonly<PlaceholderIdentity> | undefined {
  const state = readPlaceholderState(shape);
  return state.kind === 'supported' ? state.identity : undefined;
}

export function resolvePlaceholderOwner(
  pkg: OpcPackage,
  slidePartUri: string,
  selector: PlaceholderSelector,
  domain: PlaceholderDomain,
): ResolvedPlaceholderOwner {
  const normalizedSelector = normalizePlaceholderSelector(selector);
  const layoutRelationships = pkg.relationships(slidePartUri).filter(({ type }) =>
    type === SLIDE_LAYOUT_RELATIONSHIP);
  if (layoutRelationships.length !== 1) {
    throw new ModelParseError('Slide layout relationship is missing or ambiguous', slidePartUri);
  }
  const layoutRelationship = layoutRelationships[0]!;
  const layoutPartUri = layoutRelationship.targetMode === 'Internal'
    ? layoutRelationship.resolvedTarget
    : undefined;
  if (
    !layoutPartUri
    || pkg.getPart(layoutPartUri)?.contentType !== SLIDE_LAYOUT_CONTENT_TYPE
  ) {
    throw new ModelParseError('Slide layout relationship is invalid', slidePartUri);
  }
  const layoutXml = LosslessXmlDocument.parse(pkg.requirePart(layoutPartUri).bytes);
  const layoutTree = requireShapeTree(layoutXml, layoutPartUri, 'sldLayout');
  const layoutPlaceholders = readLayoutPlaceholderDescriptors(
    layoutXml,
    layoutTree,
    layoutPartUri,
    false,
  );
  const layoutMatches = layoutPlaceholders.filter((placeholder) =>
    typeof normalizedSelector === 'string'
      ? placeholder.name === normalizedSelector
      : identitiesEqual(placeholder.identity, normalizedSelector));
  if (layoutMatches.length === 0) {
    throw new RangeError('Placeholder selector was not found in the slide layout');
  }
  if (layoutMatches.length !== 1) {
    throw new ModelParseError('Placeholder selector is ambiguous in the slide layout', layoutPartUri);
  }
  const layout = layoutMatches[0]!;
  if (!domainAccepts(domain, layout.identity.type)) {
    throw new TypeError(
      `Placeholder type ${layout.identity.type} is not valid for ${domain} content`,
    );
  }

  const slideXml = LosslessXmlDocument.parse(pkg.requirePart(slidePartUri).bytes);
  const slideTree = requireShapeTree(slideXml, slidePartUri, 'sld');
  const slidePlaceholders = readLayoutPlaceholderDescriptors(
    slideXml,
    slideTree,
    slidePartUri,
    false,
  );
  const slideMatches = slidePlaceholders.filter((placeholder) =>
    placeholder.name === layout.name
    && identitiesEqual(placeholder.identity, layout.identity));
  if (slideMatches.length !== 1) {
    throw new ModelParseError('Slide placeholder owner is missing or ambiguous', slidePartUri);
  }
  const slide = slideMatches[0]!;
  if (!isEmptyPlaceholderOwner(slide.element)) {
    throw new ModelParseError('Slide placeholder owner is already filled or not empty', slidePartUri);
  }
  return {
    identity: layout.identity,
    name: layout.name,
    shapeId: slide.shapeId,
    transform: layout.transform,
    slideElement: slide.element,
    layoutElement: layout.element,
  };
}

export function materializeLayoutPlaceholders(
  pkg: OpcPackage,
  layoutPartUri: string,
  slidePartUri: string,
  rejectUnsupported = false,
  allowSlideNumber = false,
): void {
  const layoutXml = LosslessXmlDocument.parse(pkg.requirePart(layoutPartUri).bytes);
  const layoutTree = requireShapeTree(layoutXml, layoutPartUri, 'sldLayout');
  const descriptors = readLayoutPlaceholderDescriptors(
    layoutXml,
    layoutTree,
    layoutPartUri,
    rejectUnsupported,
    allowSlideNumber,
  );
  if (descriptors.length === 0) return;

  const slidePart = pkg.requirePart(slidePartUri);
  const slideXml = LosslessXmlDocument.parse(slidePart.bytes);
  const slideTree = requireShapeTree(slideXml, slidePartUri, 'sld');
  const ids = allocateShapeIds(slideTree, descriptors.length, slidePartUri);
  const rendered = descriptors.map((descriptor, index) =>
    renderMaterializedPlaceholder(descriptor, ids[index]!)).join('');
  const extensionLists = directChildren(slideTree, 'extLst', PRESENTATION_NAMESPACE);
  if (extensionLists.length > 1) {
    throw new ModelParseError('Slide shape tree contains repeated extension lists', slidePartUri);
  }
  if (extensionLists[0]) {
    slideXml.replace(extensionLists[0].start, extensionLists[0].start, rendered);
  } else {
    slideXml.appendChildXml(slideTree, rendered);
  }
  pkg.setPart(slidePartUri, slideXml.serialize(), slidePart.contentType);
}

function readLayoutPlaceholderDescriptors(
  xml: LosslessXmlDocument,
  shapeTree: XmlElement,
  partUri: string,
  rejectUnsupported: boolean,
  allowSlideNumber = false,
): readonly MaterializedPlaceholderDescriptor[] {
  const descriptors: MaterializedPlaceholderDescriptor[] = [];
  const names = new Set<string>();
  const identities = new Set<string>();
  for (const shape of shapeTree.children) {
    if (
      shape.type !== 'element'
      || !['sp', 'pic', 'graphicFrame', 'grpSp'].includes(shape.localName)
      || elementNamespaceUri(shape) !== PRESENTATION_NAMESPACE
    ) continue;
    const state = readPlaceholderState(shape);
    if (state.kind === 'unsafe') {
      throw new ModelParseError('Layout contains an unsafe placeholder identity', partUri);
    }
    if (
      state.kind === 'unsupported'
      && rejectUnsupported
      && !(allowSlideNumber && state.type === 'sldNum')
    ) {
      throw new ModelParseError('Layout contains an unsupported placeholder type', partUri);
    }
    if (state.kind !== 'supported') continue;
    const descriptor = readDescriptor(xml, shape, state, partUri);
    if (names.has(descriptor.name)) {
      throw new ModelParseError(`Layout contains duplicate placeholder name ${descriptor.name}`, partUri);
    }
    const identityKey = `${descriptor.identity.type}:${descriptor.identity.index}`;
    if (identities.has(identityKey)) {
      throw new ModelParseError(
        `Layout contains duplicate placeholder identity ${identityKey}`,
        partUri,
      );
    }
    names.add(descriptor.name);
    identities.add(identityKey);
    descriptors.push(descriptor);
  }
  return descriptors;
}

function readDescriptor(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  state: Extract<PlaceholderState, { readonly kind: 'supported' }>,
  partUri: string,
): MaterializedPlaceholderDescriptor {
  const properties = directChildren(state.nonVisual, 'cNvPr', PRESENTATION_NAMESPACE);
  if (properties.length !== 1) {
    throw new ModelParseError('Layout placeholder has ambiguous non-visual properties', partUri);
  }
  const name = strictAttribute(properties[0]!, 'name');
  const id = strictUnsignedAttribute(properties[0]!, 'id', MAX_SHAPE_ID);
  if (name === undefined || id === undefined) {
    throw new ModelParseError('Layout placeholder has an invalid name or shape id', partUri);
  }
  const isTextBox = shape.localName === 'sp'
    ? readTextShapeIsTextBox(xml, shape)
    : false;
  if (isTextBox === undefined) {
    throw new ModelParseError('Layout placeholder has an unsafe text box state', partUri);
  }
  const shapeProperties = directChildren(shape, 'spPr', PRESENTATION_NAMESPACE);
  if (shapeProperties.length > 1) {
    throw new ModelParseError('Layout placeholder has ambiguous shape properties', partUri);
  }
  const transforms = shapeProperties[0]
    ? directChildren(shapeProperties[0], 'xfrm', DRAWING_NAMESPACE)
    : [];
  if (transforms.length > 1) {
    throw new ModelParseError('Layout placeholder has ambiguous transform properties', partUri);
  }
  const textBodies = directChildren(shape, 'txBody', PRESENTATION_NAMESPACE);
  if (textBodies.length > 1) {
    throw new ModelParseError('Layout placeholder has ambiguous text bodies', partUri);
  }
  const bodyProperties = textBodies[0]
    ? directChildren(textBodies[0], 'bodyPr', DRAWING_NAMESPACE)
    : [];
  const listStyles = textBodies[0]
    ? directChildren(textBodies[0], 'lstStyle', DRAWING_NAMESPACE)
    : [];
  if (bodyProperties.length > 1 || listStyles.length > 1) {
    throw new ModelParseError('Layout placeholder has ambiguous text body properties', partUri);
  }
  return {
    name,
    shapeId: id,
    identity: state.identity,
    element: shape,
    transform: readTransform(shape, partUri),
    isTextBox,
    namespaceAttributes: inScopeNamespaceAttributes(shape),
    ...(transforms[0] ? { transformXml: xml.original(transforms[0]) } : {}),
    ...(bodyProperties[0] ? { bodyPropertiesXml: xml.original(bodyProperties[0]) } : {}),
    ...(listStyles[0] ? { listStyleXml: xml.original(listStyles[0]) } : {}),
  };
}

function readTransform(shape: XmlElement, partUri: string): Transform {
  const shapeProperties = directChildren(shape, 'spPr', PRESENTATION_NAMESPACE);
  const transforms = shapeProperties.length === 1
    ? directChildren(shapeProperties[0]!, 'xfrm', DRAWING_NAMESPACE)
    : [];
  if (transforms.length === 0) {
    return {
      x: 0 as Transform['x'],
      y: 0 as Transform['y'],
      width: 0 as Transform['width'],
      height: 0 as Transform['height'],
      rotation: 0 as Transform['rotation'],
      flipHorizontal: false,
      flipVertical: false,
    };
  }
  const transform = transforms[0]!;
  const offsets = directChildren(transform, 'off', DRAWING_NAMESPACE);
  const extents = directChildren(transform, 'ext', DRAWING_NAMESPACE);
  const x = offsets.length === 1 ? strictSignedAttribute(offsets[0]!, 'x') : undefined;
  const y = offsets.length === 1 ? strictSignedAttribute(offsets[0]!, 'y') : undefined;
  const width = extents.length === 1
    ? strictUnsignedAttribute(extents[0]!, 'cx', Number.MAX_SAFE_INTEGER)
    : undefined;
  const height = extents.length === 1
    ? strictUnsignedAttribute(extents[0]!, 'cy', Number.MAX_SAFE_INTEGER)
    : undefined;
  const rotation = strictSignedAttribute(transform, 'rot', 0);
  const flipHorizontal = strictBooleanAttribute(transform, 'flipH', false);
  const flipVertical = strictBooleanAttribute(transform, 'flipV', false);
  if (
    x === undefined
    || y === undefined
    || width === undefined
    || height === undefined
    || rotation === undefined
    || flipHorizontal === undefined
    || flipVertical === undefined
  ) {
    throw new ModelParseError('Placeholder has an invalid transform', partUri);
  }
  return {
    x: x as Transform['x'],
    y: y as Transform['y'],
    width: width as Transform['width'],
    height: height as Transform['height'],
    rotation: rotation as Transform['rotation'],
    flipHorizontal,
    flipVertical,
  };
}

function readPlaceholderState(shape: XmlElement): PlaceholderState {
  const nonVisualName = {
    sp: 'nvSpPr',
    pic: 'nvPicPr',
    graphicFrame: 'nvGraphicFramePr',
    grpSp: 'nvGrpSpPr',
  }[shape.localName];
  if (!nonVisualName) return { kind: 'none' };
  const nonVisual = directChildren(shape, nonVisualName, PRESENTATION_NAMESPACE);
  if (nonVisual.length !== 1) {
    return containsPresentationPlaceholder(shape) ? { kind: 'unsafe' } : { kind: 'none' };
  }
  const application = directChildren(nonVisual[0]!, 'nvPr', PRESENTATION_NAMESPACE);
  if (application.length !== 1) {
    return containsPresentationPlaceholder(nonVisual[0]!)
      ? { kind: 'unsafe' }
      : { kind: 'none' };
  }
  const placeholders = directChildren(application[0]!, 'ph', PRESENTATION_NAMESPACE);
  if (placeholders.length === 0) return { kind: 'none' };
  if (placeholders.length !== 1) return { kind: 'unsafe' };
  const placeholder = placeholders[0]!;
  const type = strictAttribute(placeholder, 'type');
  if (hasAmbiguousAttribute(placeholder, 'type')) return { kind: 'unsafe' };
  const resolvedType = type ?? 'body';
  if (!PLACEHOLDER_TYPE_SET.has(resolvedType)) {
    return { kind: 'unsupported', type: resolvedType };
  }
  const index = strictUnsignedAttribute(placeholder, 'idx', MAX_PLACEHOLDER_INDEX, 0);
  if (index === undefined || hasAmbiguousAttribute(placeholder, 'idx')) {
    return { kind: 'unsafe' };
  }
  return {
    kind: 'supported',
    identity: Object.freeze({ type: resolvedType as PlaceholderType, index }),
    nonVisual: nonVisual[0]!,
  };
}

function renderMaterializedPlaceholder(
  descriptor: MaterializedPlaceholderDescriptor,
  id: number,
): string {
  const name = escapeXmlAttribute(descriptor.name);
  const type = escapeXmlAttribute(descriptor.identity.type);
  const bodyProperties = descriptor.bodyPropertiesXml ?? '<a:bodyPr/>';
  const listStyle = descriptor.listStyleXml ?? '<a:lstStyle/>';
  const shapeProperties = descriptor.transformXml
    ? `<p:spPr>${descriptor.transformXml}</p:spPr>`
    : '<p:spPr/>';
  return `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}"`
    + `${descriptor.namespaceAttributes}><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/>`
    + `<p:cNvSpPr${descriptor.isTextBox ? ' txBox="1"' : ''}/><p:nvPr>`
    + `<p:ph type="${type}" idx="${descriptor.identity.index}"/></p:nvPr></p:nvSpPr>`
    + `${shapeProperties}<p:txBody>${bodyProperties}${listStyle}`
    + '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>';
}

function allocateShapeIds(
  shapeTree: XmlElement,
  count: number,
  partUri: string,
): readonly number[] {
  const used = new Set<number>();
  let maximum = 0;
  for (const element of descendants(shapeTree)) {
    if (element.localName !== 'cNvPr' || elementNamespaceUri(element) !== PRESENTATION_NAMESPACE) {
      continue;
    }
    const id = strictUnsignedAttribute(element, 'id', MAX_SHAPE_ID);
    if (id === undefined || used.has(id)) {
      throw new ModelParseError('Slide shape tree contains invalid or duplicate shape ids', partUri);
    }
    used.add(id);
    maximum = Math.max(maximum, id);
  }
  if (maximum + count > MAX_SHAPE_ID) {
    throw new ModelParseError('Slide shape ids are exhausted', partUri);
  }
  return Array.from({ length: count }, (_, index) => maximum + index + 1);
}

function requireShapeTree(
  xml: LosslessXmlDocument,
  partUri: string,
  rootName: 'sld' | 'sldLayout',
): XmlElement {
  const roots = xml.roots.filter((root) =>
    root.localName === rootName && elementNamespaceUri(root) === PRESENTATION_NAMESPACE);
  const commonSlides = roots.length === 1
    ? directChildren(roots[0]!, 'cSld', PRESENTATION_NAMESPACE)
    : [];
  const trees = commonSlides.length === 1
    ? directChildren(commonSlides[0]!, 'spTree', PRESENTATION_NAMESPACE)
    : [];
  if (trees.length !== 1) {
    throw new ModelParseError(`${rootName} does not contain a unique editable shape tree`, partUri);
  }
  return trees[0]!;
}

function strictAttribute(element: XmlElement, name: string): string | undefined {
  const attributes = semanticAttributes(element, name);
  return attributes.length === 1 && attributes[0]!.name === name
    ? attributes[0]!.value
    : undefined;
}

function strictUnsignedAttribute(
  element: XmlElement,
  name: string,
  maximum: number,
  defaultValue?: number,
): number | undefined {
  const attributes = semanticAttributes(element, name);
  if (attributes.length === 0) return defaultValue;
  if (attributes.length !== 1 || attributes[0]!.name !== name) return undefined;
  const value = attributes[0]!.value;
  if (!/^(0|[1-9]\d*)$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed <= maximum ? parsed : undefined;
}

function strictSignedAttribute(
  element: XmlElement,
  name: string,
  defaultValue?: number,
): number | undefined {
  const attributes = semanticAttributes(element, name);
  if (attributes.length === 0) return defaultValue;
  if (attributes.length !== 1 || attributes[0]!.name !== name) return undefined;
  if (!/^-?(0|[1-9]\d*)$/.test(attributes[0]!.value)) return undefined;
  const parsed = Number(attributes[0]!.value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function strictBooleanAttribute(
  element: XmlElement,
  name: string,
  defaultValue: boolean,
): boolean | undefined {
  const attributes = semanticAttributes(element, name);
  if (attributes.length === 0) return defaultValue;
  if (attributes.length !== 1 || attributes[0]!.name !== name) return undefined;
  if (attributes[0]!.value === '1' || attributes[0]!.value === 'true') return true;
  if (attributes[0]!.value === '0' || attributes[0]!.value === 'false') return false;
  return undefined;
}

function hasAmbiguousAttribute(element: XmlElement, name: string): boolean {
  const attributes = semanticAttributes(element, name);
  return attributes.length > 1 || (attributes[0] !== undefined && attributes[0].name !== name);
}

function semanticAttributes(element: XmlElement, name: string): XmlAttribute[] {
  return element.attributes.filter((attribute) =>
    attribute.localName === name && attribute.name !== 'xmlns' && !attribute.name.startsWith('xmlns:'));
}

function containsPresentationPlaceholder(element: XmlElement): boolean {
  return descendants(element).some((candidate) =>
    candidate.localName === 'ph' && elementNamespaceUri(candidate) === PRESENTATION_NAMESPACE);
}

function identitiesEqual(
  left: Readonly<PlaceholderIdentity>,
  right: Readonly<PlaceholderIdentity>,
): boolean {
  return left.type === right.type && left.index === right.index;
}

function domainAccepts(domain: PlaceholderDomain, type: PlaceholderType): boolean {
  return domain === 'text-shape'
    ? type === 'title' || type === 'body'
    : (domain === 'image' && type === 'pic')
      || (domain === 'chart' && type === 'chart')
      || (domain === 'table' && type === 'tbl')
      || (domain === 'media' && type === 'media');
}

function isEmptyPlaceholderOwner(shape: XmlElement): boolean {
  if (shape.localName !== 'sp' || elementNamespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return false;
  }
  const properties = directChildren(shape, 'spPr', PRESENTATION_NAMESPACE);
  if (properties.length !== 1) return false;
  return properties[0]!.children.every((child) =>
    child.type !== 'element'
    || (child.localName === 'xfrm' && elementNamespaceUri(child) === DRAWING_NAMESPACE));
}

function inScopeNamespaceAttributes(element: XmlElement): string {
  const declarations = new Map<string, string>();
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    for (const attribute of current.attributes) {
      if (
        (attribute.name === 'xmlns' || attribute.name.startsWith('xmlns:'))
        && !declarations.has(attribute.name)
      ) {
        declarations.set(attribute.name, attribute.value);
      }
    }
  }
  declarations.delete('xmlns:p');
  declarations.delete('xmlns:a');
  return [...declarations]
    .map(([name, value]) => ` ${name}="${escapeXmlAttribute(value)}"`)
    .join('');
}

function directChildren(
  element: XmlElement,
  localName: string,
  namespace: string,
): XmlElement[] {
  return element.children.filter((child): child is XmlElement =>
    child.type === 'element'
    && child.localName === localName
    && elementNamespaceUri(child) === namespace);
}

function descendants(element: XmlElement): XmlElement[] {
  const result: XmlElement[] = [];
  for (const child of element.children) {
    if (child.type !== 'element') continue;
    result.push(child, ...descendants(child));
  }
  return result;
}

function elementNamespaceUri(element: XmlElement): string | undefined {
  const prefix = lexicalPrefix(element.name);
  const declaration = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    const attributes = current.attributes.filter(({ name }) => name === declaration);
    if (attributes.length > 1) return undefined;
    if (attributes[0]) return attributes[0].value;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function readDataObject(
  value: unknown,
  context: string,
  supported: readonly string[],
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
