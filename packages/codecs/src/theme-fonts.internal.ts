import {
  escapeXmlAttribute,
  LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';

const DRAWINGML_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const INVALID_XML_CHARACTERS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/;
const UPDATE_KEYS = new Set(['majorLatin', 'minorLatin']);

export interface ThemeFontSnapshot {
  readonly majorLatin: string;
  readonly minorLatin: string;
}

export interface ThemeFontUpdate {
  readonly majorLatin?: string;
  readonly minorLatin?: string;
}

interface ThemeFontTarget {
  readonly element: XmlElement;
  readonly typeface?: XmlAttribute;
}

interface ThemeFontTargets {
  readonly major: ThemeFontTarget;
  readonly minor: ThemeFontTarget;
}

export function readThemeFonts(xml: LosslessXmlDocument): ThemeFontSnapshot | undefined {
  const targets = themeFontTargets(xml);
  if (!targets) return undefined;
  const majorLatin = readTypeface(targets.major);
  const minorLatin = readTypeface(targets.minor);
  return majorLatin === undefined || minorLatin === undefined
    ? undefined
    : { majorLatin, minorLatin };
}

export function replaceThemeFonts(xml: LosslessXmlDocument, value: ThemeFontUpdate): void {
  const update = normalizeThemeFontUpdate(value);
  const targets = themeFontTargets(xml);
  if (!targets) throw new Error('Theme font scheme is missing or ambiguous');
  validateUntouchedTarget(targets.major, update.majorLatin, 'major Latin');
  validateUntouchedTarget(targets.minor, update.minorLatin, 'minor Latin');
  if (update.majorLatin !== undefined) replaceTypeface(xml, targets.major, update.majorLatin);
  if (update.minorLatin !== undefined) replaceTypeface(xml, targets.minor, update.minorLatin);
}

function normalizeThemeFontUpdate(value: unknown): ThemeFontUpdate {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Theme font update must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Theme font update must be an ordinary object');
  }
  let majorLatin: string | undefined;
  let minorLatin: string | undefined;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !UPDATE_KEYS.has(key)) {
      throw new TypeError(`Unsupported theme font property: ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Theme font ${key} must be a data property`);
    }
    if (descriptor.value === undefined) continue;
    const normalized = normalizeThemeFontName(descriptor.value, key);
    if (key === 'majorLatin') majorLatin = normalized;
    else minorLatin = normalized;
  }
  if (majorLatin === undefined && minorLatin === undefined) {
    throw new TypeError('Theme font update must define majorLatin or minorLatin');
  }
  return {
    ...(majorLatin !== undefined ? { majorLatin } : {}),
    ...(minorLatin !== undefined ? { minorLatin } : {}),
  };
}

function normalizeThemeFontName(value: unknown, label: string): string {
  if (typeof value !== 'string' || !/\S/u.test(value)) {
    throw new TypeError(`Theme font ${label} must be a non-whitespace string`);
  }
  if (INVALID_XML_CHARACTERS.test(value)) {
    throw new TypeError(`Theme font ${label} contains invalid XML characters`);
  }
  return value;
}

function themeFontTargets(xml: LosslessXmlDocument): ThemeFontTargets | undefined {
  if (xml.roots.length !== 1) return undefined;
  const root = xml.roots[0];
  if (!root || root.localName !== 'theme' || namespaceUri(root) !== DRAWINGML_NAMESPACE) {
    return undefined;
  }
  const themeElements = uniqueDirectDrawingChild(root, 'themeElements');
  const fontScheme = themeElements ? uniqueDirectDrawingChild(themeElements, 'fontScheme') : undefined;
  const major = fontScheme ? uniqueDirectDrawingChild(fontScheme, 'majorFont') : undefined;
  const minor = fontScheme ? uniqueDirectDrawingChild(fontScheme, 'minorFont') : undefined;
  const majorLatin = major ? uniqueDirectDrawingChild(major, 'latin') : undefined;
  const minorLatin = minor ? uniqueDirectDrawingChild(minor, 'latin') : undefined;
  if (!majorLatin || !minorLatin) return undefined;
  const majorTarget = fontTarget(majorLatin);
  const minorTarget = fontTarget(minorLatin);
  return majorTarget && minorTarget ? { major: majorTarget, minor: minorTarget } : undefined;
}

function fontTarget(element: XmlElement): ThemeFontTarget | undefined {
  if (element.children.some((child) => child.type === 'element')) return undefined;
  if (element.children.some((child) => child.type === 'text' && /\S/u.test(child.value))) return undefined;
  const typefaces = element.attributes.filter(({ name }) => name === 'typeface');
  return typefaces.length > 1
    ? undefined
    : { element, ...(typefaces[0] ? { typeface: typefaces[0] } : {}) };
}

function uniqueDirectDrawingChild(parent: XmlElement, localName: string): XmlElement | undefined {
  const matches = parent.children.filter(
    (child): child is XmlElement =>
      child.type === 'element'
      && child.localName === localName
      && namespaceUri(child) === DRAWINGML_NAMESPACE,
  );
  return matches.length === 1 ? matches[0] : undefined;
}

function namespaceUri(element: XmlElement): string | undefined {
  const separator = element.name.indexOf(':');
  const prefix = separator < 0 ? '' : element.name.slice(0, separator);
  const declarationName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  let current: XmlElement | undefined = element;
  while (current) {
    const declarations = current.attributes.filter(({ name }) => name === declarationName);
    if (declarations.length > 1) return undefined;
    if (declarations[0]) return declarations[0].value;
    current = current.parent;
  }
  return undefined;
}

function readTypeface(target: ThemeFontTarget): string | undefined {
  const value = target.typeface?.value;
  return value === undefined || !isThemeFontName(value) ? undefined : value;
}

function isThemeFontName(value: string): boolean {
  return /\S/u.test(value) && !INVALID_XML_CHARACTERS.test(value);
}

function validateUntouchedTarget(
  target: ThemeFontTarget,
  replacement: string | undefined,
  label: string,
): void {
  if (replacement === undefined && readTypeface(target) === undefined) {
    throw new Error(`Theme ${label} typeface is missing or invalid`);
  }
}

function replaceTypeface(
  xml: LosslessXmlDocument,
  target: ThemeFontTarget,
  value: string,
): void {
  if (target.typeface) {
    if (target.typeface.value !== value) xml.replaceAttribute(target.typeface, value);
    return;
  }
  const startTag = xml.source.slice(target.element.start, target.element.startTagEnd);
  const marker = target.element.selfClosing ? startTag.lastIndexOf('/') : startTag.lastIndexOf('>');
  if (marker < 0) throw new Error(`Theme font element ${target.element.name} has an invalid start tag`);
  const offset = target.element.start + marker;
  xml.replace(offset, offset, ` typeface="${escapeXmlAttribute(value)}"`);
}
