import {
  escapeXmlAttribute,
  type LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import type { Relationship } from '@pptx/opc';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const HYPERLINK_RELATIONSHIP_TYPE = `${RELATIONSHIP_NAMESPACE}/hyperlink`;
export const SLIDE_RELATIONSHIP_TYPE = `${RELATIONSHIP_NAMESPACE}/slide`;
export const INTERNAL_SLIDE_ACTION = 'ppaction://hlinksldjump';
const OPTION_KEYS = new Set(['url', 'slide', 'tooltip']);

export type NormalizedHyperlink = Readonly<
  | {
      readonly url: string;
      readonly slide?: never;
      readonly tooltip?: string;
    }
  | {
      readonly slide: number;
      readonly url?: never;
      readonly tooltip?: string;
    }
>;

export interface ShapeHyperlinkReadContext {
  readonly relationships: readonly Relationship[];
  readonly slidePartUris: readonly string[];
}

export interface HyperlinkPrefixes {
  readonly drawing: string;
  readonly relationship: string;
}

interface HyperlinkContainerState {
  readonly properties: XmlElement;
  readonly click: HyperlinkElementState | undefined;
  readonly hover: XmlElement | undefined;
  readonly extension: XmlElement | undefined;
}

interface HyperlinkElementState {
  readonly element: XmlElement;
  readonly relationshipId: XmlAttribute;
  readonly tooltip: XmlAttribute | undefined;
  readonly action: XmlAttribute | undefined;
}

interface LocalEdit {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

export function normalizeHyperlink(
  value: unknown,
  context: string,
): NormalizedHyperlink {
  const values = readOptions(value, context);
  const url = values.url;
  const slide = values.slide;
  if ((url === undefined) === (slide === undefined)) {
    throw new TypeError(`${context} must define exactly one of url or slide`);
  }
  const tooltip = normalizeOptionalXmlString(values.tooltip, `${context} tooltip`, true);

  if (url !== undefined) {
    const normalizedUrl = normalizeOptionalXmlString(url, `${context} url`, false);
    return Object.freeze({
      url: normalizedUrl!,
      ...(tooltip !== undefined ? { tooltip } : {}),
    });
  }

  if (typeof slide !== 'number' || !Number.isSafeInteger(slide) || slide <= 0) {
    throw new TypeError(`${context} slide must be a positive safe integer`);
  }
  return Object.freeze({
    slide,
    ...(tooltip !== undefined ? { tooltip } : {}),
  });
}

export function readShapeHyperlink(
  _xml: LosslessXmlDocument,
  shape: XmlElement,
  context: ShapeHyperlinkReadContext,
): NormalizedHyperlink | undefined {
  const container = inspectHyperlinkContainer(shape);
  const click = container?.click;
  if (!click) return undefined;
  const relationships = context.relationships.filter(
    ({ id }) => id === click.relationshipId.value,
  );
  if (relationships.length !== 1) return undefined;
  const relationship = relationships[0]!;
  const tooltip = click.tooltip?.value;

  if (!click.action) {
    if (
      relationship.type !== HYPERLINK_RELATIONSHIP_TYPE
      || relationship.targetMode !== 'External'
      || relationship.target.length === 0
    ) return undefined;
    return Object.freeze({
      url: relationship.target,
      ...(click.tooltip ? { tooltip: tooltip! } : {}),
    });
  }

  if (
    click.action.value !== INTERNAL_SLIDE_ACTION
    || relationship.type !== SLIDE_RELATIONSHIP_TYPE
    || relationship.targetMode !== 'Internal'
    || !relationship.resolvedTarget
  ) return undefined;
  const matches = context.slidePartUris
    .map((partUri, index) => ({ partUri, index }))
    .filter(({ partUri }) => partUri === relationship.resolvedTarget);
  if (matches.length !== 1) return undefined;
  return Object.freeze({
    slide: matches[0]!.index + 1,
    ...(click.tooltip ? { tooltip: tooltip! } : {}),
  });
}

export function requireShapeHyperlinkRelationshipId(
  shape: XmlElement,
  partUri: string,
): string | undefined {
  const container = inspectHyperlinkContainer(shape);
  if (!container) {
    throw new ModelParseError('Shape hyperlink container is not safely editable', partUri);
  }
  return container.click?.relationshipId.value;
}

export function renderShapeHyperlink(
  hyperlink: NormalizedHyperlink,
  relationshipId: string,
  prefixes: HyperlinkPrefixes,
): string {
  const drawingName = qualifiedName(prefixes.drawing, 'hlinkClick');
  const relationshipName = qualifiedName(prefixes.relationship, 'id');
  const tooltip = Object.hasOwn(hyperlink, 'tooltip')
    ? ` tooltip="${escapeXmlAttribute(hyperlink.tooltip!)}"`
    : '';
  const action = Object.hasOwn(hyperlink, 'slide')
    ? ` action="${INTERNAL_SLIDE_ACTION}"`
    : '';
  return `<${drawingName} ${relationshipName}="${escapeXmlAttribute(relationshipId)}"` +
    `${tooltip}${action}/>`;
}

export function replaceShapeHyperlinkElement(
  xml: LosslessXmlDocument,
  shape: XmlElement,
  hyperlink: NormalizedHyperlink | undefined,
  relationshipId: string | undefined,
  partUri: string,
): boolean {
  if ((hyperlink === undefined) !== (relationshipId === undefined)) {
    throw new TypeError('Shape hyperlink and relationship ID must be supplied together');
  }
  const container = inspectHyperlinkContainer(shape);
  if (!container) {
    throw new ModelParseError('Shape hyperlink container is not safely editable', partUri);
  }

  if (!container.click) {
    if (!hyperlink) return false;
    const rendered = renderInsertedHyperlink(
      container.properties,
      hyperlink,
      relationshipId!,
    );
    const anchor = container.hover ?? container.extension;
    if (anchor) xml.replace(anchor.start, anchor.start, rendered);
    else xml.appendChildXml(container.properties, rendered);
    return true;
  }

  if (!hyperlink) {
    xml.removeElement(container.click.element);
    return true;
  }

  const source = xml.original(container.click.element);
  const updated = patchExistingHyperlink(
    source,
    container.click,
    hyperlink,
    relationshipId!,
  );
  if (updated === source) return false;
  xml.replaceElement(container.click.element, updated);
  return true;
}

export function shapeHyperlinksEqual(
  left: NormalizedHyperlink | undefined,
  right: NormalizedHyperlink | undefined,
): boolean {
  if (!left || !right) return left === right;
  const leftIsUrl = Object.hasOwn(left, 'url');
  const rightIsUrl = Object.hasOwn(right, 'url');
  if (leftIsUrl !== rightIsUrl) return false;
  if (leftIsUrl) {
    if (left.url !== right.url) return false;
  } else if (left.slide !== right.slide) {
    return false;
  }
  const leftHasTooltip = Object.hasOwn(left, 'tooltip');
  const rightHasTooltip = Object.hasOwn(right, 'tooltip');
  return leftHasTooltip === rightHasTooltip
    && (!leftHasTooltip || left.tooltip === right.tooltip);
}

export function relationshipReferenceCount(
  xml: LosslessXmlDocument,
  relationshipId: string,
): number {
  let count = 0;
  for (const element of xml.elements()) {
    for (const attribute of element.attributes) {
      if (
        attribute.value === relationshipId
        && isRelationshipIdAttribute(element, attribute)
      ) count += 1;
    }
  }
  return count;
}

export function removeDrawingHyperlinkReferences(
  xml: LosslessXmlDocument,
  relationshipIds: ReadonlySet<string>,
): boolean {
  if (relationshipIds.size === 0) return false;
  const selected = xml.elements().filter((element) =>
    (element.localName === 'hlinkClick' || element.localName === 'hlinkHover')
    && namespaceUri(element) === DRAWING_NAMESPACE
    && element.attributes.some((attribute) =>
      relationshipIds.has(attribute.value)
      && isRelationshipIdAttribute(element, attribute)),
  );
  if (selected.length === 0) return false;
  const selectedSet = new Set(selected);
  for (const element of selected) {
    let parent = element.parent;
    let nested = false;
    while (parent) {
      if (selectedSet.has(parent)) {
        nested = true;
        break;
      }
      parent = parent.parent;
    }
    if (!nested) xml.removeElement(element);
  }
  return true;
}

function readOptions(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    if (descriptor.value !== undefined) result[key] = descriptor.value;
  }
  return result;
}

function normalizeOptionalXmlString(
  value: unknown,
  context: string,
  allowEmpty: boolean,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') throw new TypeError(`${context} must be a string`);
  if (!allowEmpty && value.length === 0) throw new TypeError(`${context} must not be empty`);
  if (containsInvalidXmlCharacter(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value;
}

function containsInvalidXmlCharacter(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const codePoint = value.codePointAt(index)!;
    const valid = codePoint === 0x9
      || codePoint === 0xa
      || codePoint === 0xd
      || (codePoint >= 0x20 && codePoint <= 0xd7ff)
      || (codePoint >= 0xe000 && codePoint <= 0xfffd)
      || (codePoint >= 0x10000 && codePoint <= 0x10ffff);
    if (!valid) return true;
    if (codePoint > 0xffff) index += 1;
  }
  return false;
}

function inspectHyperlinkContainer(
  shape: XmlElement,
): HyperlinkContainerState | undefined {
  if (shape.localName !== 'sp' || namespaceUri(shape) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const nonVisualContainers = directChildren(shape)
    .filter(({ localName }) => localName === 'nvSpPr');
  if (
    nonVisualContainers.length !== 1
    || namespaceUri(nonVisualContainers[0]!) !== PRESENTATION_NAMESPACE
  ) return undefined;
  const properties = directChildren(nonVisualContainers[0]!)
    .filter(({ localName }) => localName === 'cNvPr');
  if (
    properties.length !== 1
    || namespaceUri(properties[0]!) !== PRESENTATION_NAMESPACE
  ) return undefined;
  const root = properties[0]!;
  if (root.children.some((child) => child.type === 'text' && /\S/u.test(child.value))) {
    return undefined;
  }

  let click: HyperlinkElementState | undefined;
  let hover: XmlElement | undefined;
  let extension: XmlElement | undefined;
  let previousStage = -1;
  for (const child of directChildren(root)) {
    const stage = child.localName === 'hlinkClick'
      ? 0
      : child.localName === 'hlinkHover'
        ? 1
        : child.localName === 'extLst'
          ? 2
          : undefined;
    if (
      stage === undefined
      || namespaceUri(child) !== DRAWING_NAMESPACE
      || stage < previousStage
    ) return undefined;
    previousStage = stage;
    if (child.localName === 'hlinkClick') {
      if (click) return undefined;
      click = inspectHyperlinkElement(child);
      if (!click) return undefined;
    } else if (child.localName === 'hlinkHover') {
      if (hover) return undefined;
      hover = child;
    } else {
      if (extension) return undefined;
      extension = child;
    }
  }
  return { properties: root, click, hover, extension };
}

function inspectHyperlinkElement(
  element: XmlElement,
): HyperlinkElementState | undefined {
  if (namespaceUri(element) !== DRAWING_NAMESPACE) return undefined;
  if (element.children.some((child) => child.type === 'text' && /\S/u.test(child.value))) {
    return undefined;
  }
  let soundSeen = false;
  let extensionSeen = false;
  let previousChildStage = -1;
  for (const child of directChildren(element)) {
    const stage = child.localName === 'snd'
      ? 0
      : child.localName === 'extLst'
        ? 1
        : undefined;
    if (
      stage === undefined
      || namespaceUri(child) !== DRAWING_NAMESPACE
      || stage < previousChildStage
    ) return undefined;
    previousChildStage = stage;
    if (child.localName === 'snd') {
      if (soundSeen) return undefined;
      soundSeen = true;
    } else {
      if (extensionSeen) return undefined;
      extensionSeen = true;
    }
  }
  const relationshipIds: XmlAttribute[] = [];
  let tooltip: XmlAttribute | undefined;
  let action: XmlAttribute | undefined;
  for (const attribute of element.attributes) {
    if (isNamespaceDeclaration(attribute)) continue;
    if (attribute.localName === 'id') {
      if (!isRelationshipIdAttribute(element, attribute)) return undefined;
      relationshipIds.push(attribute);
      continue;
    }
    if (attribute.localName === 'tooltip') {
      if (attribute.name !== 'tooltip' || tooltip) return undefined;
      tooltip = attribute;
      continue;
    }
    if (attribute.localName === 'action') {
      if (attribute.name !== 'action' || action) return undefined;
      action = attribute;
    }
  }
  if (
    relationshipIds.length !== 1
    || relationshipIds[0]!.value.length === 0
    || (action && action.value !== INTERNAL_SLIDE_ACTION)
  ) return undefined;
  return {
    element,
    relationshipId: relationshipIds[0]!,
    tooltip,
    action,
  };
}

function renderInsertedHyperlink(
  properties: XmlElement,
  hyperlink: NormalizedHyperlink,
  relationshipId: string,
): string {
  const drawing = findEffectivePrefix(properties, DRAWING_NAMESPACE, true)
    ?? 'a';
  const relationship = findEffectivePrefix(
    properties,
    RELATIONSHIP_NAMESPACE,
    false,
    drawing,
  ) ?? (drawing === 'r' ? 'rel' : 'r');
  const prefixes = { drawing, relationship };
  const declarations = [
    namespaceUriForPrefix(properties, drawing) === DRAWING_NAMESPACE
      ? ''
      : drawing === ''
        ? ` xmlns="${DRAWING_NAMESPACE}"`
        : ` xmlns:${drawing}="${DRAWING_NAMESPACE}"`,
    namespaceUriForPrefix(properties, relationship) === RELATIONSHIP_NAMESPACE
      ? ''
      : ` xmlns:${relationship}="${RELATIONSHIP_NAMESPACE}"`,
  ].join('');
  const rendered = renderShapeHyperlink(hyperlink, relationshipId, prefixes);
  if (declarations.length === 0) return rendered;
  const marker = rendered.indexOf(' ');
  return rendered.slice(0, marker) + declarations + rendered.slice(marker);
}

function patchExistingHyperlink(
  source: string,
  state: HyperlinkElementState,
  hyperlink: NormalizedHyperlink,
  relationshipId: string,
): string {
  const offset = state.element.start;
  const edits: LocalEdit[] = [];
  if (state.relationshipId.value !== relationshipId) {
    edits.push({
      start: state.relationshipId.valueStart - offset,
      end: state.relationshipId.valueEnd - offset,
      replacement: escapeXmlAttribute(relationshipId),
    });
  }

  const targetHasTooltip = Object.hasOwn(hyperlink, 'tooltip');
  if (state.tooltip) {
    if (!targetHasTooltip) {
      edits.push(removeAttributeEdit(source, state.tooltip, offset));
    } else if (state.tooltip.value !== hyperlink.tooltip) {
      edits.push({
        start: state.tooltip.valueStart - offset,
        end: state.tooltip.valueEnd - offset,
        replacement: escapeXmlAttribute(hyperlink.tooltip!),
      });
    }
  }

  const targetIsInternal = Object.hasOwn(hyperlink, 'slide');
  if (state.action) {
    if (!targetIsInternal) edits.push(removeAttributeEdit(source, state.action, offset));
  }

  const additions = [
    !state.tooltip && targetHasTooltip
      ? ` tooltip="${escapeXmlAttribute(hyperlink.tooltip!)}"`
      : '',
    !state.action && targetIsInternal
      ? ` action="${INTERNAL_SLIDE_ACTION}"`
      : '',
  ].join('');
  if (additions.length > 0) {
    const insertion = startTagAttributeInsertion(source, state.element);
    edits.push({ start: insertion, end: insertion, replacement: additions });
  }
  return applyLocalEdits(source, edits);
}

function startTagAttributeInsertion(
  source: string,
  element: XmlElement,
): number {
  let index = element.startTagEnd - element.start - 1;
  while (index > 0 && /\s/u.test(source[index - 1] ?? '')) index -= 1;
  if (source[index - 1] === '/') index -= 1;
  return index;
}

function removeAttributeEdit(
  source: string,
  attribute: XmlAttribute,
  offset: number,
): LocalEdit {
  let start = attribute.start - offset;
  while (start > 0 && /[\t ]/u.test(source[start - 1] ?? '')) start -= 1;
  return {
    start,
    end: attribute.end - offset,
    replacement: '',
  };
}

function applyLocalEdits(
  source: string,
  edits: readonly LocalEdit[],
): string {
  if (edits.length === 0) return source;
  let output = source;
  const ordered = [...edits].sort(
    (left, right) => right.start - left.start || right.end - left.end,
  );
  let previousStart = source.length;
  for (const edit of ordered) {
    if (
      edit.start < 0
      || edit.end < edit.start
      || edit.end > source.length
      || edit.end > previousStart
    ) throw new Error('Overlapping local shape hyperlink edits');
    output = output.slice(0, edit.start) + edit.replacement + output.slice(edit.end);
    previousStart = edit.start;
  }
  return output;
}

function findEffectivePrefix(
  element: XmlElement,
  namespace: string,
  allowDefault: boolean,
  excluded?: string,
): string | undefined {
  const seen = new Set<string>();
  for (let current: XmlElement | undefined = element; current; current = current.parent) {
    for (const attribute of current.attributes) {
      const prefix = attribute.name === 'xmlns'
        ? ''
        : attribute.name.startsWith('xmlns:')
          ? attribute.name.slice('xmlns:'.length)
          : undefined;
      if (prefix === undefined || seen.has(prefix)) continue;
      seen.add(prefix);
      if (
        attribute.value === namespace
        && (allowDefault || prefix !== '')
        && prefix !== excluded
      ) return prefix;
    }
  }
  return undefined;
}

function isRelationshipIdAttribute(
  element: XmlElement,
  attribute: XmlAttribute,
): boolean {
  if (attribute.localName !== 'id') return false;
  const prefix = lexicalPrefix(attribute.name);
  return prefix !== ''
    && namespaceUriForPrefix(element, prefix) === RELATIONSHIP_NAMESPACE;
}

function isNamespaceDeclaration(attribute: XmlAttribute): boolean {
  return attribute.name === 'xmlns' || attribute.name.startsWith('xmlns:');
}

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

function qualifiedName(prefix: string, localName: string): string {
  return prefix === '' ? localName : `${prefix}:${localName}`;
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element',
  );
}
