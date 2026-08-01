import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const SECTION_NAMESPACE =
  'http://schemas.microsoft.com/office/powerpoint/2010/main';
const SECTION_EXTENSION_URI =
  '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const SECTION_ID = /^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/i;

export interface SectionSnapshotData {
  readonly id: string;
  readonly title: string;
  readonly slideIds: readonly number[];
}

export interface NormalizedAddPresentationSectionOptions {
  readonly title: string;
  readonly order?: number;
}

export interface NormalizedAddPresentationSlideOptions {
  readonly masterName?: string;
  readonly sectionTitle?: string;
}

interface SectionMemberState {
  readonly id: number;
  readonly element: XmlElement;
}

interface SectionState {
  readonly element: XmlElement;
  readonly idAttribute: XmlAttribute;
  readonly nameAttribute: XmlAttribute;
  readonly memberList: XmlElement;
  readonly members: readonly SectionMemberState[];
  readonly snapshot: SectionSnapshotData;
}

interface SectionDocumentState {
  readonly root: XmlElement;
  readonly extensionList?: XmlElement;
  readonly extension?: XmlElement;
  readonly sectionList?: XmlElement;
  readonly sections: readonly SectionState[];
}

export function readPresentationSections(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
): readonly SectionSnapshotData[] | undefined {
  const state = resolveSectionState(xml, validSlideIds);
  return state?.sections.map(({ snapshot }) => ({
    id: snapshot.id,
    title: snapshot.title,
    slideIds: [...snapshot.slideIds],
  }));
}

export function insertPresentationSection(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  title: string,
  order: number,
  id: string,
): SectionSnapshotData {
  const normalizedTitle = normalizePresentationSectionTitle(title);
  const normalizedId = normalizePresentationSectionId(id);
  const state = requireSectionState(xml, validSlideIds);
  const normalizedOrder = normalizePresentationSectionIndex(
    order,
    'Presentation section order',
    state.sections.length,
    true,
  );
  if (state.sections.some(({ snapshot }) => snapshot.id === normalizedId)) {
    throw new RangeError(`Presentation section ${normalizedId} already exists`);
  }

  const snapshot = { id: normalizedId, title: normalizedTitle, slideIds: [] } as const;
  if (state.sectionList) {
    const prefix = qualifier(state.sectionList.name);
    const sectionXml = renderSection(prefix, snapshot);
    insertOwnedChild(xml, state.sectionList, state.sections.map(({ element }) => element), normalizedOrder, sectionXml);
    return { ...snapshot, slideIds: [] };
  }

  const presentationPrefix = qualifier(state.root.name);
  const sectionXml = renderSection('p14:', snapshot);
  const listXml = `<p14:sectionLst xmlns:p14="${SECTION_NAMESPACE}">${sectionXml}</p14:sectionLst>`;
  const extensionXml = `<${presentationPrefix}ext uri="${SECTION_EXTENSION_URI}">${listXml}</${presentationPrefix}ext>`;
  if (state.extensionList) {
    xml.appendChildXml(state.extensionList, extensionXml);
  } else {
    xml.appendChildXml(
      state.root,
      `<${presentationPrefix}extLst>${extensionXml}</${presentationPrefix}extLst>`,
    );
  }
  return { ...snapshot, slideIds: [] };
}

export function renamePresentationSection(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  sectionId: string,
  title: string,
): boolean {
  const normalizedId = normalizePresentationSectionId(sectionId);
  const normalizedTitle = normalizePresentationSectionTitle(title);
  const section = requireSection(requireSectionState(xml, validSlideIds), normalizedId);
  if (section.snapshot.title === normalizedTitle) return false;
  xml.replaceAttribute(section.nameAttribute, normalizedTitle);
  return true;
}

export function movePresentationSection(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  sectionId: string,
  toIndex: number,
): boolean {
  const normalizedId = normalizePresentationSectionId(sectionId);
  const state = requireSectionState(xml, validSlideIds);
  const section = requireSection(state, normalizedId);
  const fromIndex = state.sections.indexOf(section);
  const normalizedIndex = normalizePresentationSectionIndex(
    toIndex,
    'Presentation section target index',
    state.sections.length,
    false,
  );
  if (fromIndex === normalizedIndex) return false;
  const reordered = [...state.sections];
  reordered.splice(fromIndex, 1);
  reordered.splice(normalizedIndex, 0, section);
  replaceOwnedSlots(xml, state.sections.map(({ element }) => element), reordered.map(({ element }) => element));
  return true;
}

export function deletePresentationSection(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  sectionId: string,
): boolean {
  const normalizedId = normalizePresentationSectionId(sectionId);
  const state = requireSectionState(xml, validSlideIds);
  const section = requireSection(state, normalizedId);
  if (state.sections.length > 1 || !state.extension || !state.sectionList || !state.extensionList) {
    xml.removeElement(section.element);
    return true;
  }

  if (!isCleanSectionExtension(xml, state)) {
    xml.removeElement(section.element);
    return true;
  }
  if (isCleanExtensionList(xml, state.extensionList, state.extension)) {
    xml.removeElement(state.extensionList);
  } else {
    xml.removeElement(state.extension);
  }
  return true;
}

export function assignPresentationSlideToSection(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  slideId: number,
  sectionId: string | undefined,
): boolean {
  requireSlideId(slideId, validSlideIds);
  const normalizedId = sectionId === undefined
    ? undefined
    : normalizePresentationSectionId(sectionId);
  const state = requireSectionState(xml, validSlideIds);
  const target = normalizedId === undefined ? undefined : requireSection(state, normalizedId);
  const current = state.sections.find(({ members }) =>
    members.some(({ id }) => id === slideId));
  if (current === target) return false;

  const member = current?.members.find(({ id }) => id === slideId);
  if (member) xml.removeElement(member.element);
  if (target) {
    const prefix = qualifier(target.memberList.name);
    insertOwnedChild(
      xml,
      target.memberList,
      target.members.map(({ element }) => element),
      target.members.length,
      `<${prefix}sldId id="${slideId}"/>`,
    );
  }
  return true;
}

export function removePresentationSlideFromSections(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  slideId: number,
): boolean {
  return assignPresentationSlideToSection(xml, validSlideIds, slideId, undefined);
}

export function copyPresentationSlideSection(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  sourceSlideId: number,
  targetSlideId: number,
): boolean {
  requireSlideId(sourceSlideId, validSlideIds);
  requireSlideId(targetSlideId, validSlideIds);
  const state = requireSectionState(xml, validSlideIds);
  const source = state.sections.find(({ members }) =>
    members.some(({ id }) => id === sourceSlideId));
  const target = state.sections.find(({ members }) =>
    members.some(({ id }) => id === targetSlideId));
  if (source === target) return false;
  const targetMember = target?.members.find(({ id }) => id === targetSlideId);
  if (targetMember) xml.removeElement(targetMember.element);
  if (source) {
    const prefix = qualifier(source.memberList.name);
    insertOwnedChild(
      xml,
      source.memberList,
      source.members.map(({ element }) => element),
      source.members.length,
      `<${prefix}sldId id="${targetSlideId}"/>`,
    );
  }
  return true;
}

export function sortPresentationSectionSlides(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
  orderedSlideIds: readonly number[],
): boolean {
  if (!Array.isArray(orderedSlideIds) || orderedSlideIds.length !== validSlideIds.size) {
    throw new TypeError('Ordered presentation slide IDs must contain every slide');
  }
  const order = new Map<number, number>();
  for (const [index, slideId] of orderedSlideIds.entries()) {
    requireSlideId(slideId, validSlideIds);
    if (order.has(slideId)) throw new TypeError('Ordered presentation slide IDs must be unique');
    order.set(slideId, index);
  }
  const state = requireSectionState(xml, validSlideIds);
  let changed = false;
  for (const section of state.sections) {
    const sorted = [...section.members].sort(
      (left, right) => order.get(left.id)! - order.get(right.id)!,
    );
    if (sorted.some((member, index) => member !== section.members[index])) {
      replaceOwnedSlots(
        xml,
        section.members.map(({ element }) => element),
        sorted.map(({ element }) => element),
      );
      changed = true;
    }
  }
  return changed;
}

export function normalizeAddPresentationSectionOptions(
  value: unknown,
): NormalizedAddPresentationSectionOptions {
  const data = readDataObject(value, 'Presentation section options', ['title', 'order']);
  const title = normalizePresentationSectionTitle(data.title);
  const order = data.order === undefined
    ? undefined
    : normalizeNonNegativeInteger(data.order, 'Presentation section order');
  return { title, ...(order !== undefined ? { order } : {}) };
}

export function normalizeAddPresentationSlideOptions(
  value: unknown,
): NormalizedAddPresentationSlideOptions {
  if (value === undefined) return {};
  const data = readDataObject(value, 'Add slide options', ['masterName', 'sectionTitle']);
  const masterName = data.masterName === undefined
    ? undefined
    : normalizeNonWhitespaceXmlString(data.masterName, 'Slide master name');
  const sectionTitle = data.sectionTitle === undefined
    ? undefined
    : normalizePresentationSectionTitle(data.sectionTitle);
  return {
    ...(masterName !== undefined ? { masterName } : {}),
    ...(sectionTitle !== undefined ? { sectionTitle } : {}),
  };
}

export function normalizePresentationSectionTitle(value: unknown): string {
  return normalizeNonWhitespaceXmlString(value, 'Presentation section title');
}

function normalizeNonWhitespaceXmlString(value: unknown, label: string): string {
  if (typeof value !== 'string' || !/\S/u.test(value)) {
    throw new TypeError(`${label} must be a non-whitespace string`);
  }
  if (!isValidXmlString(value)) {
    throw new TypeError(`${label} contains invalid XML characters`);
  }
  return value;
}

export function normalizePresentationSectionId(value: unknown): string {
  if (typeof value !== 'string' || !SECTION_ID.test(value)) {
    throw new TypeError('Presentation section ID must be a brace-wrapped GUID');
  }
  return value;
}

export function normalizePresentationSectionIndex(
  value: unknown,
  label: string,
  upperBound: number,
  allowEnd: boolean,
): number {
  const index = normalizeNonNegativeInteger(value, label);
  const maximum = allowEnd ? upperBound : upperBound - 1;
  if (index > maximum) throw new RangeError(`${label} is out of range`);
  return index;
}

export function createPresentationSectionId(): string {
  if (typeof globalThis.crypto?.randomUUID !== 'function') {
    throw new Error('Presentation section creation requires crypto.randomUUID()');
  }
  return `{${globalThis.crypto.randomUUID().toUpperCase()}}`;
}

function resolveSectionState(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
): SectionDocumentState | undefined {
  if (xml.roots.length !== 1) return undefined;
  const root = xml.roots[0];
  if (!root || root.localName !== 'presentation' || namespaceUri(root) !== PRESENTATION_NAMESPACE) {
    return undefined;
  }
  const extensionLists = directChildren(root).filter(
    (child) => child.localName === 'extLst' && namespaceUri(child) === PRESENTATION_NAMESPACE,
  );
  if (extensionLists.length > 1) return undefined;
  const extensionList = extensionLists[0];
  if (!extensionList) return { root, sections: [] };

  const extensions = directChildren(extensionList).filter(
    (child) => child.localName === 'ext' && namespaceUri(child) === PRESENTATION_NAMESPACE,
  );
  const owned: XmlElement[] = [];
  for (const extension of extensions) {
    const uriAttributes = extension.attributes.filter(({ name }) => name === 'uri');
    const hasSectionList = directChildren(extension).some(
      (child) => child.localName === 'sectionLst' && namespaceUri(child) === SECTION_NAMESPACE,
    );
    const matches = uriAttributes.some(
      ({ value }) => value.toUpperCase() === SECTION_EXTENSION_URI,
    );
    if (matches) {
      if (uriAttributes.length !== 1) return undefined;
      owned.push(extension);
    } else if (hasSectionList) {
      return undefined;
    }
  }
  if (owned.length > 1) return undefined;
  const extension = owned[0];
  if (!extension) return { root, extensionList, sections: [] };
  const sectionLists = directChildren(extension).filter(({ localName }) => localName === 'sectionLst');
  if (sectionLists.length !== 1 || namespaceUri(sectionLists[0]!) !== SECTION_NAMESPACE) {
    return undefined;
  }
  const sectionList = sectionLists[0]!;
  const sectionElements = directChildren(sectionList).filter(({ localName }) => localName === 'section');
  if (sectionElements.some((element) => namespaceUri(element) !== SECTION_NAMESPACE)) return undefined;

  const seenSections = new Set<string>();
  const seenSlides = new Set<number>();
  const sections: SectionState[] = [];
  for (const element of sectionElements) {
    const nameAttributes = element.attributes.filter(({ name }) => name === 'name');
    const idAttributes = element.attributes.filter(({ name }) => name === 'id');
    if (nameAttributes.length !== 1 || idAttributes.length !== 1) return undefined;
    const nameAttribute = nameAttributes[0]!;
    const idAttribute = idAttributes[0]!;
    if (!SECTION_ID.test(idAttribute.value) || !isValidXmlString(nameAttribute.value)) return undefined;
    if (seenSections.has(idAttribute.value)) return undefined;
    seenSections.add(idAttribute.value);

    const memberLists = directChildren(element).filter(({ localName }) => localName === 'sldIdLst');
    if (memberLists.length !== 1 || namespaceUri(memberLists[0]!) !== SECTION_NAMESPACE) {
      return undefined;
    }
    const memberList = memberLists[0]!;
    const memberElements = directChildren(memberList).filter(({ localName }) => localName === 'sldId');
    if (memberElements.some((member) => namespaceUri(member) !== SECTION_NAMESPACE)) return undefined;
    const members: SectionMemberState[] = [];
    for (const member of memberElements) {
      const idAttributes = member.attributes.filter(({ name }) => name === 'id');
      if (
        idAttributes.length !== 1
        || directChildren(member).length > 0
        || member.children.some((child) => child.type === 'text' && /\S/u.test(child.value))
      ) return undefined;
      const raw = idAttributes[0]!.value;
      if (!/^\d+$/.test(raw)) return undefined;
      const id = Number(raw);
      if (!Number.isSafeInteger(id) || !validSlideIds.has(id) || seenSlides.has(id)) {
        return undefined;
      }
      seenSlides.add(id);
      members.push({ id, element: member });
    }
    sections.push({
      element,
      idAttribute,
      nameAttribute,
      memberList,
      members,
      snapshot: {
        id: idAttribute.value,
        title: nameAttribute.value,
        slideIds: members.map(({ id }) => id),
      },
    });
  }
  return { root, extensionList, extension, sectionList, sections };
}

function requireSectionState(
  xml: LosslessXmlDocument,
  validSlideIds: ReadonlySet<number>,
): SectionDocumentState {
  const state = resolveSectionState(xml, validSlideIds);
  if (!state) throw new ModelParseError('Presentation sections are not safely editable');
  return state;
}

function requireSection(state: SectionDocumentState, id: string): SectionState {
  const section = state.sections.find(({ snapshot }) => snapshot.id === id);
  if (!section) throw new RangeError(`Presentation section ${id} was not found`);
  return section;
}

function requireSlideId(value: unknown, validSlideIds: ReadonlySet<number>): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || !validSlideIds.has(value)) {
    throw new RangeError(`Presentation slide ID ${String(value)} was not found`);
  }
  return value;
}

function renderSection(prefix: string, section: SectionSnapshotData): string {
  return `<${prefix}section name="${escapeXmlAttribute(section.title)}" id="${escapeXmlAttribute(section.id)}"><${prefix}sldIdLst/></${prefix}section>`;
}

function insertOwnedChild(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  owned: readonly XmlElement[],
  index: number,
  childXml: string,
): void {
  const next = owned[index];
  if (next) {
    xml.replace(next.start, next.start, childXml);
    return;
  }
  const previous = owned.at(-1);
  if (previous) {
    xml.replace(previous.end, previous.end, childXml);
    return;
  }
  const firstElement = directChildren(parent)[0];
  if (firstElement) xml.replace(firstElement.start, firstElement.start, childXml);
  else xml.appendChildXml(parent, childXml);
}

function replaceOwnedSlots(
  xml: LosslessXmlDocument,
  slots: readonly XmlElement[],
  values: readonly XmlElement[],
): void {
  for (const [index, slot] of slots.entries()) {
    const value = values[index]!;
    if (slot === value) continue;
    xml.replaceElement(slot, xml.original(value));
  }
}

function isCleanSectionExtension(
  xml: LosslessXmlDocument,
  state: SectionDocumentState,
): boolean {
  const { extension, sectionList } = state;
  if (!extension || !sectionList) return false;
  const extensionAttributes = nonNamespaceAttributes(extension);
  const listAttributes = nonNamespaceAttributes(sectionList);
  if (
    extensionAttributes.length !== 1
    || extensionAttributes[0]?.name !== 'uri'
    || listAttributes.length !== 0
  ) return false;
  if (directChildren(extension).some((child) => child !== sectionList)) return false;
  if (directChildren(sectionList).some((child) => child.localName !== 'section')) return false;
  return residualContent(xml, extension, [sectionList]).trim() === ''
    && residualContent(xml, sectionList, state.sections.map(({ element }) => element)).trim() === '';
}

function isCleanExtensionList(
  xml: LosslessXmlDocument,
  extensionList: XmlElement,
  extension: XmlElement,
): boolean {
  return nonNamespaceAttributes(extensionList).length === 0
    && directChildren(extensionList).length === 1
    && directChildren(extensionList)[0] === extension
    && residualContent(xml, extensionList, [extension]).trim() === '';
}

function residualContent(
  xml: LosslessXmlDocument,
  parent: XmlElement,
  removed: readonly XmlElement[],
): string {
  if (parent.selfClosing) return '';
  let cursor = parent.startTagEnd;
  let result = '';
  for (const child of [...removed].sort((left, right) => left.start - right.start)) {
    result += xml.source.slice(cursor, child.start);
    cursor = child.end;
  }
  result += xml.source.slice(cursor, parent.endTagStart);
  return result;
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
  const result: Record<string, unknown> = {};
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

function normalizeNonNegativeInteger(value: unknown, label: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) {
    throw new TypeError(`${label} must be a non-negative safe integer`);
  }
  return value;
}

function isValidXmlString(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint === 0x9
      || codePoint === 0xa
      || codePoint === 0xd
      || (codePoint >= 0x20 && codePoint <= 0xd7ff)
      || (codePoint >= 0xe000 && codePoint <= 0xfffd)
      || (codePoint >= 0x10000 && codePoint <= 0x10ffff)
    ) continue;
    return false;
  }
  return true;
}

function namespaceUri(element: XmlElement): string | undefined {
  return namespaceUriForPrefix(element, lexicalPrefix(element.name));
}

function namespaceUriForPrefix(element: XmlElement, prefix: string): string | undefined {
  let current: XmlElement | undefined = element;
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  while (current) {
    const declarations = current.attributes.filter(({ name }) => name === declarationName);
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
    current = current.parent;
  }
  return undefined;
}

function lexicalPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : name.slice(0, separator);
}

function qualifier(name: string): string {
  const prefix = lexicalPrefix(name);
  return prefix === '' ? '' : `${prefix}:`;
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}

function nonNamespaceAttributes(element: XmlElement): readonly XmlAttribute[] {
  return element.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
}
