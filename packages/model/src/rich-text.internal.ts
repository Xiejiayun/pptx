import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type {
  CharacterBullet,
  NumberedBullet,
  NumberingStyle,
  ParagraphBullet,
  RichTextColor,
  RichTextParagraph,
  RichTextRun,
  RichTextRunStyle,
  TextAlignment,
} from './text.js';

const TEXT_ALIGNMENT_TO_OOXML: Readonly<Record<TextAlignment, string>> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
};

const OOXML_TO_TEXT_ALIGNMENT = new Map(
  Object.entries(TEXT_ALIGNMENT_TO_OOXML).map(([alignment, value]) => [value, alignment as TextAlignment]),
);

const NUMBERING_STYLES = new Set<NumberingStyle>([
  'alphaLcParenBoth',
  'alphaLcParenR',
  'alphaLcPeriod',
  'alphaUcParenBoth',
  'alphaUcParenR',
  'alphaUcPeriod',
  'arabicParenBoth',
  'arabicParenR',
  'arabicPeriod',
  'arabicPlain',
  'romanLcParenBoth',
  'romanLcParenR',
  'romanLcPeriod',
  'romanUcParenBoth',
  'romanUcParenR',
  'romanUcPeriod',
]);

const BULLET_ELEMENT_NAMES = new Set([
  'buClrTx',
  'buClr',
  'buSzTx',
  'buSzPct',
  'buSzPts',
  'buFontTx',
  'buFont',
  'buNone',
  'buAutoNum',
  'buChar',
  'buBlip',
]);

const BULLET_INSERTION_FOLLOWERS = new Set(['tabLst', 'defRPr', 'extLst']);
const DEFAULT_BULLET_CHARACTER = '•';
const DEFAULT_BULLET_INDENT = 27;

export type NormalizedParagraphBullet =
  | Required<CharacterBullet>
  | Required<NumberedBullet>;

interface NormalizedRichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly align?: TextAlignment;
  readonly bullet?: NormalizedParagraphBullet | false;
}

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

export function normalizeRichText(value: unknown): readonly NormalizedRichTextParagraph[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new TypeError('Rich text must contain at least one paragraph');
  }
  return value.map((paragraph, paragraphIndex) => {
    if (!paragraph || typeof paragraph !== 'object' || Array.isArray(paragraph)) {
      throw new TypeError(`Rich text paragraph ${paragraphIndex} must be an object`);
    }
    assertSupportedKeys(paragraph, ['align', 'bullet', 'runs'], `Rich text paragraph ${paragraphIndex}`);
    const candidate = paragraph as { align?: unknown; bullet?: unknown; runs?: unknown };
    const runs = candidate.runs;
    if (!Array.isArray(runs)) throw new TypeError(`Rich text paragraph ${paragraphIndex} runs must be an array`);
    const align = candidate.align === undefined
      ? undefined
      : normalizeTextAlignment(candidate.align, `Rich text paragraph ${paragraphIndex} align`);
    const bullet = candidate.bullet === undefined
      ? undefined
      : normalizeParagraphBullet(candidate.bullet, `Rich text paragraph ${paragraphIndex} bullet`);
    return {
      runs: runs.map((run, runIndex) => normalizeRun(run, paragraphIndex, runIndex)),
      ...(align ? { align } : {}),
      ...(bullet !== undefined ? { bullet } : {}),
    };
  });
}

interface RenderRichTextOptions {
  readonly prefix?: string;
  readonly defaultAlign?: TextAlignment;
  readonly defaultBullet?: NormalizedParagraphBullet | false;
  readonly paragraphProperties?: readonly (string | undefined)[];
  readonly endParagraphProperties?: string;
}

export function renderRichTextParagraphs(
  paragraphs: readonly NormalizedRichTextParagraph[],
  options: RenderRichTextOptions = {},
): string {
  const prefix = options.prefix ?? 'a:';
  const defaultEndProperties = `<${prefix}endParaRPr lang="en-US" dirty="0"/>`;
  return paragraphs
    .map(
      ({ align, bullet, runs }, index) =>
        `<${prefix}p>${renderParagraphProperties(
          options.paragraphProperties?.[index] ?? options.paragraphProperties?.[0],
          prefix,
          align ?? options.defaultAlign,
          bullet === false
            ? undefined
            : bullet ?? (options.defaultBullet === false ? undefined : options.defaultBullet),
        )}${runs
          .map((run) => renderRun(run, prefix))
          .join('')}${options.endParagraphProperties ?? defaultEndProperties}</${prefix}p>`,
    )
    .join('');
}

export function readRichText(xml: LosslessXmlDocument, element: XmlElement): readonly RichTextParagraph[] {
  const textBody = directChildren(element, 'txBody')[0];
  if (!textBody) return [];
  return directChildren(textBody, 'p').map((paragraph) => {
    const align = readParagraphAlignment(xml, paragraph);
    const bullet = readParagraphBullet(xml, paragraph);
    return {
      runs: readRuns(xml, paragraph),
      ...(align ? { align } : {}),
      ...(bullet ? { bullet } : {}),
    };
  });
}

export function replaceRichText(
  xml: LosslessXmlDocument,
  element: XmlElement,
  paragraphs: readonly NormalizedRichTextParagraph[],
  partUri: string,
  save: (xml: string) => void,
): void {
  const textBody = directChildren(element, 'txBody')[0];
  if (!textBody) throw new ModelParseError('Shape does not contain a text body', partUri);
  const existing = directChildren(textBody, 'p');
  const template = existing[0];
  if (!template) throw new ModelParseError('Shape does not contain a text paragraph', partUri);
  const prefix = qualifiedPrefix(template.name);
  const endProperties = directChildren(template, 'endParaRPr')[0];
  const paragraphProperties = existing.map((paragraph) => {
    const properties = directChildren(paragraph, 'pPr')[0];
    return properties ? xml.original(properties) : undefined;
  });
  const replacement = renderRichTextParagraphs(paragraphs, {
    prefix,
    paragraphProperties,
    ...(endProperties ? { endParagraphProperties: xml.original(endProperties) } : {}),
  });
  xml.replaceElement(template, replacement);
  for (const extra of existing.slice(1)) xml.removeElement(extra);
  save(xml.serialize());
}

export function normalizeTextAlignment(value: unknown, context: string): TextAlignment {
  if (typeof value !== 'string' || !Object.hasOwn(TEXT_ALIGNMENT_TO_OOXML, value)) {
    throw new TypeError(`${context} must be left, center, right, or justify`);
  }
  return value as TextAlignment;
}

export function normalizeParagraphBullet(
  value: unknown,
  context: string,
): NormalizedParagraphBullet | false {
  if (typeof value === 'boolean') {
    return value
      ? { kind: 'bullet', character: DEFAULT_BULLET_CHARACTER, indent: DEFAULT_BULLET_INDENT }
      : false;
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be a boolean or bullet configuration object`);
  }
  const candidate = value as {
    character?: unknown;
    indent?: unknown;
    kind?: unknown;
    startAt?: unknown;
    style?: unknown;
  };
  if (candidate.kind === 'bullet') {
    assertSupportedKeys(value, ['character', 'indent', 'kind'], context);
    return {
      kind: 'bullet',
      character: normalizeBulletCharacter(candidate.character, context),
      indent: normalizeBulletIndent(candidate.indent, context),
    };
  }
  if (candidate.kind === 'number') {
    assertSupportedKeys(value, ['indent', 'kind', 'startAt', 'style'], context);
    return {
      kind: 'number',
      style: normalizeNumberingStyle(candidate.style, context),
      startAt: normalizeNumberingStart(candidate.startAt, context),
      indent: normalizeBulletIndent(candidate.indent, context),
    };
  }
  throw new TypeError(`${context} kind must be bullet or number`);
}

export function renderParagraphProperties(
  template: string | undefined,
  prefix: string,
  alignment: TextAlignment | undefined,
  bullet?: NormalizedParagraphBullet,
): string {
  const align = alignment ? ` algn="${TEXT_ALIGNMENT_TO_OOXML[alignment]}"` : '';
  const initial = template ?? `<${prefix}pPr${align} indent="0" marL="0"><${prefix}buNone/></${prefix}pPr>`;
  const aligned = template
    ? updateParagraphAttribute(
        initial,
        'algn',
        alignment ? TEXT_ALIGNMENT_TO_OOXML[alignment] : undefined,
      )
    : initial;
  return renderParagraphBullet(aligned, prefix, bullet);
}

function renderParagraphBullet(
  template: string,
  prefix: string,
  bullet: NormalizedParagraphBullet | undefined,
): string {
  const source = LosslessXmlDocument.parse(template);
  const sourceRoot = requireParagraphPropertiesRoot(source);
  const sourceChildren = directChildren(sourceRoot);
  const sourceBulletChildren = sourceChildren.filter(({ localName }) => BULLET_ELEMENT_NAMES.has(localName));
  const hadActiveBullet = sourceChildren.some(({ localName }) => ['buChar', 'buAutoNum', 'buBlip'].includes(localName));
  const margin = readIntegerAttribute(source, sourceRoot, 'marL');
  const indent = readIntegerAttribute(source, sourceRoot, 'indent');
  let withIndent = template;
  if (bullet) {
    const marginEmu = Math.round(bullet.indent * 12700);
    withIndent = updateParagraphAttribute(withIndent, 'marL', String(marginEmu));
    withIndent = updateParagraphAttribute(withIndent, 'indent', String(-marginEmu));
  } else if (hadActiveBullet && margin !== undefined && indent === -margin) {
    withIndent = updateParagraphAttribute(withIndent, 'marL', '0');
    withIndent = updateParagraphAttribute(withIndent, 'indent', '0');
  }
  if (!bullet && sourceBulletChildren.length === 1 && sourceBulletChildren[0]?.localName === 'buNone') {
    return withIndent;
  }

  const properties = LosslessXmlDocument.parse(withIndent);
  const root = requireParagraphPropertiesRoot(properties);
  const children = directChildren(root);
  for (const child of children) {
    if (BULLET_ELEMENT_NAMES.has(child.localName)) properties.removeElement(child);
  }
  const bulletXml = renderBulletXml(prefix, bullet);
  const follower = children.find((child) => BULLET_INSERTION_FOLLOWERS.has(child.localName));
  if (follower) properties.replace(follower.start, follower.start, bulletXml);
  else properties.appendChildXml(root, bulletXml);
  return properties.serialize();
}

function updateParagraphAttribute(template: string, name: string, value: string | undefined): string {
  const properties = LosslessXmlDocument.parse(template);
  const root = requireParagraphPropertiesRoot(properties);
  const attribute = properties.attribute(root, name);
  if (value !== undefined) {
    if (attribute) properties.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? properties.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      properties.replace(insertionPoint, insertionPoint, ` ${name}="${escapeXmlAttribute(value)}"`);
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(properties.source[start - 1] ?? '')) start -= 1;
    properties.replace(start, attribute.end, '');
  }
  return properties.serialize();
}

function renderBulletXml(prefix: string, bullet: NormalizedParagraphBullet | undefined): string {
  if (!bullet) return `<${prefix}buNone/>`;
  if (bullet.kind === 'bullet') {
    return `<${prefix}buSzPct val="100000"/><${prefix}buChar char="${escapeXmlAttribute(bullet.character)}"/>`;
  }
  return `<${prefix}buSzPct val="100000"/><${prefix}buFont typeface="+mj-lt"/><${prefix}buAutoNum type="${bullet.style}" startAt="${bullet.startAt}"/>`;
}

function requireParagraphPropertiesRoot(xml: LosslessXmlDocument): XmlElement {
  const root = xml.roots[0];
  if (!root || root.localName !== 'pPr') throw new ModelParseError('Invalid paragraph properties template');
  return root;
}

function normalizeBulletCharacter(value: unknown, context: string): string {
  const character = value === undefined ? DEFAULT_BULLET_CHARACTER : value;
  if (typeof character !== 'string') throw new TypeError(`${context} character must be a string`);
  if (!isValidBulletCharacter(character)) {
    throw new TypeError(`${context} character must contain one valid Unicode character`);
  }
  return character;
}

function isValidBulletCharacter(character: string): boolean {
  const codePoint = character.codePointAt(0);
  return [...character].length === 1
    && codePoint !== undefined
    && !(codePoint >= 0xD800 && codePoint <= 0xDFFF)
    && !/\p{Cc}/u.test(character)
    && !containsInvalidXmlCharacter(character)
    && codePoint !== 0xFFFE
    && codePoint !== 0xFFFF;
}

function normalizeBulletIndent(value: unknown, context: string): number {
  const indent = value === undefined ? DEFAULT_BULLET_INDENT : value;
  if (typeof indent !== 'number' || !Number.isFinite(indent)) {
    throw new TypeError(`${context} indent must be finite`);
  }
  if (indent < 0 || indent > 4032) throw new RangeError(`${context} indent must be between 0 and 4032 points`);
  return Math.round(indent * 100) / 100;
}

function normalizeNumberingStyle(value: unknown, context: string): NumberingStyle {
  const style = value === undefined ? 'arabicPeriod' : value;
  if (typeof style !== 'string' || !NUMBERING_STYLES.has(style as NumberingStyle)) {
    throw new TypeError(`${context} style is unsupported`);
  }
  return style as NumberingStyle;
}

function normalizeNumberingStart(value: unknown, context: string): number {
  const startAt = value === undefined ? 1 : value;
  if (typeof startAt !== 'number' || !Number.isInteger(startAt)) {
    throw new TypeError(`${context} startAt must be an integer`);
  }
  if (startAt < 1 || startAt > 32767) {
    throw new RangeError(`${context} startAt must be between 1 and 32767`);
  }
  return startAt;
}

function normalizeRun(value: unknown, paragraphIndex: number, runIndex: number): RichTextRun {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} must be an object`);
  }
  assertSupportedKeys(value, ['softBreakBefore', 'style', 'text'], `Rich text run ${paragraphIndex},${runIndex}`);
  const candidate = value as { text?: unknown; style?: unknown; softBreakBefore?: unknown };
  if (typeof candidate.text !== 'string') {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} text must be a string`);
  }
  if (/\r|\n/.test(candidate.text)) {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} text cannot contain line breaks`);
  }
  if (containsInvalidXmlCharacter(candidate.text)) {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} text contains invalid XML characters`);
  }
  if (candidate.softBreakBefore !== undefined && typeof candidate.softBreakBefore !== 'boolean') {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} softBreakBefore must be a boolean`);
  }
  const style = candidate.style === undefined ? undefined : normalizeStyle(candidate.style, paragraphIndex, runIndex);
  return {
    text: candidate.text,
    ...(style ? { style } : {}),
    ...(candidate.softBreakBefore !== undefined ? { softBreakBefore: candidate.softBreakBefore } : {}),
  };
}

function normalizeStyle(value: unknown, paragraphIndex: number, runIndex: number): RichTextRunStyle {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} style must be an object`);
  }
  assertSupportedKeys(
    value,
    ['bold', 'color', 'fontFamily', 'fontSize', 'italic'],
    `Rich text run ${paragraphIndex},${runIndex} style`,
  );
  const candidate = value as RichTextRunStyle;
  if (candidate.fontFamily !== undefined) {
    if (typeof candidate.fontFamily !== 'string' || candidate.fontFamily.length === 0) {
      throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} fontFamily must be a non-empty string`);
    }
    if (containsInvalidXmlCharacter(candidate.fontFamily)) {
      throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} fontFamily contains invalid XML characters`);
    }
  }
  if (candidate.fontSize !== undefined) {
    if (typeof candidate.fontSize !== 'number' || !Number.isFinite(candidate.fontSize)) {
      throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} fontSize must be finite`);
    }
    if (candidate.fontSize < 1 || candidate.fontSize > 4000) {
      throw new RangeError(`Rich text run ${paragraphIndex},${runIndex} fontSize must be between 1 and 4000 points`);
    }
  }
  for (const [name, property] of [
    ['bold', candidate.bold],
    ['italic', candidate.italic],
  ] as const) {
    if (property !== undefined && typeof property !== 'boolean') {
      throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} ${name} must be a boolean`);
    }
  }
  const color = candidate.color === undefined ? undefined : normalizeColor(candidate.color, paragraphIndex, runIndex);
  return {
    ...(candidate.fontFamily !== undefined ? { fontFamily: candidate.fontFamily } : {}),
    ...(candidate.fontSize !== undefined ? { fontSize: Math.round(candidate.fontSize * 100) / 100 } : {}),
    ...(candidate.bold !== undefined ? { bold: candidate.bold } : {}),
    ...(candidate.italic !== undefined ? { italic: candidate.italic } : {}),
    ...(color ? { color } : {}),
  };
}

function normalizeColor(value: unknown, paragraphIndex: number, runIndex: number): RichTextColor {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} color must be an object`);
  }
  assertSupportedKeys(value, ['kind', 'value'], `Rich text run ${paragraphIndex},${runIndex} color`);
  const candidate = value as { kind?: unknown; value?: unknown };
  if (candidate.kind === 'srgb') {
    if (typeof candidate.value !== 'string' || !/^#?[\da-f]{6}$/i.test(candidate.value)) {
      throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} sRGB color must contain six hex digits`);
    }
    return { kind: 'srgb', value: candidate.value.replace(/^#/, '').toUpperCase() };
  }
  if (candidate.kind === 'scheme') {
    if (typeof candidate.value !== 'string' || !SCHEME_COLORS.has(candidate.value)) {
      throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} scheme color is unsupported`);
    }
    return { kind: 'scheme', value: candidate.value };
  }
  throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} color kind must be srgb or scheme`);
}

function renderRun(run: RichTextRun, prefix: string): string {
  const softBreak = run.softBreakBefore ? `<${prefix}br/>` : '';
  if (run.text.length === 0 && run.style === undefined) return softBreak;
  const style = run.style ?? {};
  const attributes = [
    'lang="en-US"',
    style.fontSize === undefined ? '' : `sz="${Math.round(style.fontSize * 100)}"`,
    style.bold === undefined ? '' : `b="${style.bold ? 1 : 0}"`,
    style.italic === undefined ? '' : `i="${style.italic ? 1 : 0}"`,
    'dirty="0"',
  ].filter(Boolean).join(' ');
  const color = style.color ?? { kind: 'scheme' as const, value: 'tx1' };
  const colorXml = color.kind === 'srgb'
    ? `<${prefix}srgbClr val="${color.value}"/>`
    : `<${prefix}schemeClr val="${color.value}"/>`;
  const latin = escapeXmlAttribute(style.fontFamily ?? '+mn-lt');
  const eastAsian = escapeXmlAttribute(style.fontFamily ?? '+mn-ea');
  const complexScript = escapeXmlAttribute(style.fontFamily ?? '+mn-cs');
  return `${softBreak}<${prefix}r><${prefix}rPr ${attributes}><${prefix}solidFill>${colorXml}</${prefix}solidFill><${prefix}latin typeface="${latin}"/><${prefix}ea typeface="${eastAsian}"/><${prefix}cs typeface="${complexScript}"/></${prefix}rPr><${prefix}t xml:space="preserve">${escapeXmlText(run.text)}</${prefix}t></${prefix}r>`;
}

function readRuns(xml: LosslessXmlDocument, paragraph: XmlElement): RichTextRun[] {
  const runs: RichTextRun[] = [];
  let pendingBreaks = 0;
  for (const child of paragraph.children) {
    if (child.type !== 'element') continue;
    if (child.localName === 'br') {
      pendingBreaks += 1;
      continue;
    }
    if (child.localName !== 'r' && child.localName !== 'fld') continue;
    while (pendingBreaks > 1) {
      runs.push({ text: '', softBreakBefore: true });
      pendingBreaks -= 1;
    }
    const style = readStyle(xml, child);
    runs.push({
      text: xml.descendants(child, 't').map((text) => xml.text(text)).join(''),
      ...(style ? { style } : {}),
      ...(pendingBreaks === 1 ? { softBreakBefore: true } : {}),
    });
    pendingBreaks = 0;
  }
  while (pendingBreaks > 0) {
    runs.push({ text: '', softBreakBefore: true });
    pendingBreaks -= 1;
  }
  return runs;
}

function readParagraphAlignment(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): TextAlignment | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  const value = properties ? xml.attribute(properties, 'algn')?.value : undefined;
  return value ? OOXML_TO_TEXT_ALIGNMENT.get(value) : undefined;
}

function readParagraphBullet(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): CharacterBullet | NumberedBullet | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  const indent = readBulletIndent(xml, properties);
  const characterElement = directChildren(properties, 'buChar')[0];
  const character = characterElement ? xml.attribute(characterElement, 'char')?.value : undefined;
  if (character !== undefined && isValidBulletCharacter(character)) {
    return {
      kind: 'bullet',
      character,
      ...(indent !== undefined ? { indent } : {}),
    };
  }
  const numbering = directChildren(properties, 'buAutoNum')[0];
  if (!numbering) return undefined;
  const style = xml.attribute(numbering, 'type')?.value;
  if (!style || !NUMBERING_STYLES.has(style as NumberingStyle)) return undefined;
  const startRaw = xml.attribute(numbering, 'startAt')?.value ?? '';
  const startValue = /^\d+$/.test(startRaw) ? Number(startRaw) : Number.NaN;
  const startAt = Number.isInteger(startValue) && startValue >= 1 && startValue <= 32767
    ? startValue
    : undefined;
  return {
    kind: 'number',
    style: style as NumberingStyle,
    ...(startAt !== undefined ? { startAt } : {}),
    ...(indent !== undefined ? { indent } : {}),
  };
}

function readBulletIndent(xml: LosslessXmlDocument, properties: XmlElement): number | undefined {
  const margin = readIntegerAttribute(xml, properties, 'marL');
  if (margin === undefined || margin < 0) return undefined;
  const points = Math.round((margin / 12700) * 100) / 100;
  if (points > 4032) return undefined;
  return points;
}

function readIntegerAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: string,
): number | undefined {
  const raw = xml.attribute(element, name)?.value;
  if (!raw || !/^-?\d+$/.test(raw)) return undefined;
  const value = Number(raw);
  return Number.isSafeInteger(value) ? value : undefined;
}

function readStyle(xml: LosslessXmlDocument, run: XmlElement): RichTextRunStyle | undefined {
  const properties = directChildren(run, 'rPr')[0];
  if (!properties) return undefined;
  const size = Number.parseInt(xml.attribute(properties, 'sz')?.value ?? '', 10);
  const font = directChildren(properties, 'latin')[0];
  const fontFamily = font ? xml.attribute(font, 'typeface')?.value : undefined;
  const fill = directChildren(properties, 'solidFill')[0];
  const colorElement = fill?.children.find(
    (child): child is XmlElement => child.type === 'element' && ['srgbClr', 'schemeClr'].includes(child.localName),
  );
  const colorValue = colorElement ? xml.attribute(colorElement, 'val')?.value : undefined;
  const color = colorElement && colorValue
    ? colorElement.localName === 'srgbClr'
      ? { kind: 'srgb' as const, value: colorValue.toUpperCase() }
      : { kind: 'scheme' as const, value: colorValue }
    : undefined;
  const bold = booleanAttribute(xml, properties, 'b');
  const italic = booleanAttribute(xml, properties, 'i');
  const style: RichTextRunStyle = {
    ...(fontFamily !== undefined ? { fontFamily } : {}),
    ...(Number.isFinite(size) && size > 0 ? { fontSize: size / 100 } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(color ? { color } : {}),
  };
  return Object.keys(style).length > 0 ? style : undefined;
}

function booleanAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  name: 'b' | 'i',
): boolean | undefined {
  const attribute = xml.attribute(element, name);
  if (!attribute) return undefined;
  return ['1', 'true', 'on'].includes(attribute.value);
}

function directChildren(element: XmlElement, localName?: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && (!localName || child.localName === localName),
  );
}

function qualifiedPrefix(name: string): string {
  const separator = name.indexOf(':');
  return separator < 0 ? '' : `${name.slice(0, separator)}:`;
}

function containsInvalidXmlCharacter(value: string): boolean {
  return /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value);
}

function assertSupportedKeys(value: object, supported: readonly string[], context: string): void {
  const unsupported = Object.keys(value).find((key) => !supported.includes(key));
  if (unsupported) throw new TypeError(`${context} property ${unsupported} is not supported yet`);
}
