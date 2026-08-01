import {
  escapeXmlAttribute,
  type LosslessXmlDocument,
  type XmlAttribute,
  type XmlElement,
} from '@pptx/lossless-xml';
import type { MediaPlaybackSettings } from './media.js';

const PLAYBACK_EXTENSION_URI = '{C13D3E4A-5148-4B6D-A7E7-505054582D4F}';
const PLAYBACK_KEYS = new Set(['play', 'loop', 'hideWhenStopped', 'volume']);

export interface NormalizedMediaPlaybackSettings {
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
}

export function normalizeMediaName(value: unknown): string {
  return normalizeXmlString(value, 'name', false)!;
}

export function normalizeMediaAltText(value: unknown): string | undefined {
  return normalizeXmlString(value, 'alt text', true);
}

export function normalizeMediaPlaybackSettings(
  value: unknown,
): Readonly<NormalizedMediaPlaybackSettings> | undefined {
  if (value === undefined) return undefined;
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Media settings must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Media settings must be an ordinary object');
  }
  const values = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !PLAYBACK_KEYS.has(key)) {
      throw new TypeError(`Media settings contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Media setting ${key} must be a data property`);
    }
    values[key] = descriptor.value;
  }
  return Object.freeze({
    play: normalizePlay(values.play),
    loop: normalizeBoolean(values.loop, false, 'loop'),
    hideWhenStopped: normalizeBoolean(values.hideWhenStopped, false, 'hideWhenStopped'),
    volume: normalizeVolume(values.volume),
  });
}

export function replaceMediaMetadataAttribute(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  name: 'name' | 'descr',
  value: string | undefined,
): boolean {
  const properties = xml.descendants(picture, 'cNvPr')[0];
  if (!properties) throw new Error('Media picture has no non-visual properties');
  const attributes = properties.attributes.filter(({ name: candidate }) => candidate === name);
  if (attributes.length > 1) throw new Error(`Media picture has repeated ${name} attributes`);
  const current = attributes[0];
  if (value === undefined) {
    if (!current) return false;
    removeAttribute(xml, properties, current);
    return true;
  }
  if (current?.value === value) return false;
  if (current) xml.replaceAttribute(current, value);
  else insertAttribute(xml, properties, name, value);
  return true;
}

export function replaceMediaPlaybackExtension(
  xml: LosslessXmlDocument,
  picture: XmlElement,
  value: Readonly<NormalizedMediaPlaybackSettings> | undefined,
): boolean {
  const applicationProperties = xml.descendants(picture, 'nvPr')[0];
  if (!applicationProperties) throw new Error('Media picture has no application properties');
  const extensionLists = directChildren(applicationProperties, 'extLst');
  if (extensionLists.length > 1) throw new Error('Media picture has repeated extension lists');
  const extensionList = extensionLists[0];
  const extensions = extensionList
    ? directChildren(extensionList, 'ext').filter(
        (extension) => xml.attribute(extension, 'uri')?.value === PLAYBACK_EXTENSION_URI,
      )
    : [];
  if (extensions.length > 1) throw new Error('Media picture has repeated playback extensions');
  const current = extensions[0];
  if (value === undefined) {
    if (!current) return false;
    xml.removeElement(current);
    return true;
  }
  const rendered = renderPlaybackExtension(value);
  if (current && xml.original(current) === rendered) return false;
  if (current) xml.replace(current.start, current.end, rendered);
  else if (extensionList) xml.appendChildXml(extensionList, rendered);
  else xml.appendChildXml(applicationProperties, `<p:extLst>${rendered}</p:extLst>`);
  return true;
}

export function mediaPlaybackSettingsEqual(
  left: Readonly<MediaPlaybackSettings>,
  right: Readonly<NormalizedMediaPlaybackSettings>,
): boolean {
  return left.play === right.play
    && left.loop === right.loop
    && left.hideWhenStopped === right.hideWhenStopped
    && left.volume === right.volume;
}

function renderPlaybackExtension(value: Readonly<NormalizedMediaPlaybackSettings>): string {
  return `<p:ext uri="${PLAYBACK_EXTENSION_URI}">`
    + '<px:playback xmlns:px="urn:pptx-ooxml:media" '
    + `play="${value.play}" loop="${value.loop ? 1 : 0}" `
    + `hideWhenStopped="${value.hideWhenStopped ? 1 : 0}" `
    + `volume="${Math.round(value.volume * 100_000)}"/></p:ext>`;
}

function normalizeXmlString(
  value: unknown,
  label: string,
  allowUndefined: boolean,
): string | undefined {
  if (value === undefined && allowUndefined) return undefined;
  if (typeof value !== 'string') throw new TypeError(`Media ${label} must be a string`);
  for (const character of value) {
    const codePoint = character.codePointAt(0)!;
    if (
      codePoint === 0x09
      || codePoint === 0x0A
      || codePoint === 0x0D
      || (codePoint >= 0x20 && codePoint <= 0xD7FF)
      || (codePoint >= 0xE000 && codePoint <= 0xFFFD)
      || (codePoint >= 0x10000 && codePoint <= 0x10FFFF)
    ) continue;
    throw new TypeError(`Media ${label} contains invalid XML characters`);
  }
  return value;
}

function normalizePlay(value: unknown): 'click' | 'auto' {
  if (value === undefined) return 'click';
  if (value === 'click' || value === 'auto') return value;
  throw new TypeError('Media play must be click or auto');
}

function normalizeBoolean(value: unknown, fallback: boolean, label: string): boolean {
  if (value === undefined) return fallback;
  if (typeof value !== 'boolean') throw new TypeError(`Media ${label} must be a boolean`);
  return value;
}

function normalizeVolume(value: unknown): number {
  if (value === undefined) return 1;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError('Media volume must be finite');
  }
  if (value < 0 || value > 1) throw new RangeError('Media volume must be between 0 and 1');
  return value === 0 ? 0 : value;
}

function insertAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
  value: string,
): void {
  const position = element.selfClosing
    ? xml.source.lastIndexOf('/', element.startTagEnd - 1)
    : element.startTagEnd - 1;
  if (position <= element.start) throw new Error(`Media ${element.localName} start tag is invalid`);
  xml.replace(position, position, ` ${name}="${escapeXmlAttribute(value)}"`);
}

function removeAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  attribute: XmlAttribute,
): void {
  let start = attribute.start;
  while (start > element.start && /[\t ]/.test(xml.source[start - 1] ?? '')) start -= 1;
  xml.replace(start, attribute.end, '');
}

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}
