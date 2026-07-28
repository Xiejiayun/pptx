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
  ParagraphLineSpacing,
  ParagraphSpacing,
  ParagraphTabStop,
  ParagraphTabStopAlignment,
  RichTextColor,
  RichTextParagraph,
  RichTextRun,
  RichTextRunStyle,
  RichTextUnderline,
  RichTextUnderlineStyle,
  TextAlignment,
} from './text.js';
import { EMU_PER_INCH } from './units.js';

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
const TAB_STOP_INSERTION_FOLLOWERS = new Set(['defRPr', 'extLst']);
const SPACING_ELEMENT_NAMES = new Set(['lnSpc', 'spcBef', 'spcAft']);
const DEFAULT_BULLET_CHARACTER = '•';
const DEFAULT_BULLET_INDENT = 27;
const MAX_SPACING_POINTS = 1584;
const MAX_SPACING_FACTOR = 132;
const MIN_COORDINATE_32 = -2_147_483_648;
const MAX_COORDINATE_32 = 2_147_483_647;

const TAB_STOP_ALIGNMENT_TO_OOXML: Readonly<Record<ParagraphTabStopAlignment, string>> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  decimal: 'dec',
};

const OOXML_TO_TAB_STOP_ALIGNMENT = new Map(
  Object.entries(TAB_STOP_ALIGNMENT_TO_OOXML).map(([alignment, value]) => [
    value,
    alignment as ParagraphTabStopAlignment,
  ]),
);

export type NormalizedParagraphBullet =
  | Required<CharacterBullet>
  | Required<NumberedBullet>;

type NormalizedParagraphLineSpacing = ParagraphLineSpacing;

export interface NormalizedParagraphSpacing {
  readonly before?: number;
  readonly after?: number;
  readonly line?: NormalizedParagraphLineSpacing;
}

export interface NormalizedParagraphSpacingUpdate {
  readonly before?: number | false;
  readonly after?: number | false;
  readonly line?: NormalizedParagraphLineSpacing | false;
}

export interface NormalizedParagraphTabStop {
  readonly positionEmu: number;
  readonly alignment: ParagraphTabStopAlignment;
}

interface NormalizedRichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly align?: TextAlignment;
  readonly bullet?: NormalizedParagraphBullet | false;
  readonly level?: number;
  readonly spacing?: NormalizedParagraphSpacingUpdate | false;
  readonly tabStops?: readonly NormalizedParagraphTabStop[] | false;
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

const UNDERLINE_STYLES = new Set<RichTextUnderlineStyle>([
  'words',
  'sng',
  'dbl',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDbl',
]);

export function normalizeRichText(value: unknown): readonly NormalizedRichTextParagraph[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new TypeError('Rich text must contain at least one paragraph');
  }
  return value.map((paragraph, paragraphIndex) => {
    if (!paragraph || typeof paragraph !== 'object' || Array.isArray(paragraph)) {
      throw new TypeError(`Rich text paragraph ${paragraphIndex} must be an object`);
    }
    assertSupportedKeys(
      paragraph,
      ['align', 'bullet', 'level', 'runs', 'spacing', 'tabStops'],
      `Rich text paragraph ${paragraphIndex}`,
    );
    const candidate = paragraph as {
      align?: unknown;
      bullet?: unknown;
      level?: unknown;
      runs?: unknown;
      spacing?: unknown;
      tabStops?: unknown;
    };
    const runs = candidate.runs;
    if (!Array.isArray(runs)) throw new TypeError(`Rich text paragraph ${paragraphIndex} runs must be an array`);
    const align = candidate.align === undefined
      ? undefined
      : normalizeTextAlignment(candidate.align, `Rich text paragraph ${paragraphIndex} align`);
    const bullet = candidate.bullet === undefined
      ? undefined
      : normalizeParagraphBullet(candidate.bullet, `Rich text paragraph ${paragraphIndex} bullet`);
    const level = candidate.level === undefined
      ? undefined
      : normalizeParagraphLevel(candidate.level, `Rich text paragraph ${paragraphIndex} level`);
    const spacing = candidate.spacing === undefined
      ? undefined
      : candidate.spacing === false
        ? false
        : normalizeParagraphSpacing(candidate.spacing, `Rich text paragraph ${paragraphIndex} spacing`);
    const tabStops = candidate.tabStops === undefined
      ? undefined
      : candidate.tabStops === false
        ? false
        : normalizeParagraphTabStops(candidate.tabStops, `Rich text paragraph ${paragraphIndex} tabStops`);
    return {
      runs: runs.map((run, runIndex) => normalizeRun(run, paragraphIndex, runIndex)),
      ...(align ? { align } : {}),
      ...(bullet !== undefined ? { bullet } : {}),
      ...(level !== undefined ? { level } : {}),
      ...(spacing !== undefined ? { spacing } : {}),
      ...(tabStops !== undefined ? { tabStops } : {}),
    };
  });
}

interface RenderRichTextOptions {
  readonly prefix?: string;
  readonly defaultAlign?: TextAlignment;
  readonly defaultBullet?: NormalizedParagraphBullet | false;
  readonly defaultLevel?: number;
  readonly defaultSpacing?: NormalizedParagraphSpacingUpdate;
  readonly defaultTabStops?: readonly NormalizedParagraphTabStop[];
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
      ({ align, bullet, level, runs, spacing, tabStops }, index) =>
        `<${prefix}p>${renderParagraphProperties(
          options.paragraphProperties?.[index] ?? options.paragraphProperties?.[0],
          prefix,
          align ?? options.defaultAlign,
          bullet === false
            ? undefined
            : bullet ?? (options.defaultBullet === false ? undefined : options.defaultBullet),
          resolveParagraphSpacing(options.defaultSpacing, spacing),
          level ?? options.defaultLevel,
          tabStops === false ? undefined : tabStops ?? options.defaultTabStops,
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
    const level = readParagraphLevel(xml, paragraph);
    const bullet = readParagraphBullet(xml, paragraph, level ?? 0);
    const spacing = readParagraphSpacing(xml, paragraph);
    const tabStops = readParagraphTabStops(xml, paragraph);
    return {
      runs: readRuns(xml, paragraph),
      ...(align ? { align } : {}),
      ...(bullet ? { bullet } : {}),
      ...(level !== undefined ? { level } : {}),
      ...(spacing ? { spacing } : {}),
      ...(tabStops !== undefined ? { tabStops } : {}),
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

export function normalizeParagraphLevel(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || !Number.isInteger(value)) {
    throw new TypeError(`${context} must be an integer`);
  }
  if (value < 0 || value > 8) throw new RangeError(`${context} must be between 0 and 8`);
  return value;
}

export function normalizeParagraphTabStops(
  value: unknown,
  context: string,
): readonly NormalizedParagraphTabStop[] {
  if (!Array.isArray(value)) throw new TypeError(`${context} must be an array`);
  return value.map((stop, index) => {
    const stopContext = `${context} stop ${index}`;
    if (!stop || typeof stop !== 'object' || Array.isArray(stop)) {
      throw new TypeError(`${stopContext} must be an object`);
    }
    assertSupportedKeys(stop, ['alignment', 'position'], stopContext);
    const candidate = stop as { alignment?: unknown; position?: unknown };
    if (typeof candidate.position !== 'number' || !Number.isFinite(candidate.position)) {
      throw new TypeError(`${stopContext} position must be finite`);
    }
    const positionEmu = Math.round(candidate.position * EMU_PER_INCH);
    if (positionEmu < MIN_COORDINATE_32 || positionEmu > MAX_COORDINATE_32) {
      throw new RangeError(`${stopContext} position must fit a signed 32-bit OOXML coordinate`);
    }
    const alignment = candidate.alignment === undefined ? 'left' : candidate.alignment;
    if (typeof alignment !== 'string' || !Object.hasOwn(TAB_STOP_ALIGNMENT_TO_OOXML, alignment)) {
      throw new TypeError(`${stopContext} alignment must be left, center, right, or decimal`);
    }
    return { positionEmu, alignment: alignment as ParagraphTabStopAlignment };
  });
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

export function normalizeParagraphSpacing(
  value: unknown,
  context: string,
): NormalizedParagraphSpacingUpdate {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  assertSupportedKeys(value, ['after', 'before', 'line'], context);
  const candidate = value as { after?: unknown; before?: unknown; line?: unknown };
  if (candidate.before === undefined && candidate.after === undefined && candidate.line === undefined) {
    throw new TypeError(`${context} must provide before, after, or line`);
  }
  const before = candidate.before === undefined
    ? undefined
    : normalizeParagraphSpacingPoint(candidate.before, `${context} before`, true);
  const after = candidate.after === undefined
    ? undefined
    : normalizeParagraphSpacingPoint(candidate.after, `${context} after`, true);
  const line = candidate.line === undefined
    ? undefined
    : candidate.line === false
      ? false
      : normalizeParagraphLineSpacing(candidate.line, `${context} line`);
  return {
    ...(before !== undefined ? { before } : {}),
    ...(after !== undefined ? { after } : {}),
    ...(line !== undefined ? { line } : {}),
  };
}

export function resolveParagraphSpacing(
  defaultSpacing?: NormalizedParagraphSpacingUpdate,
  paragraphSpacing?: NormalizedParagraphSpacingUpdate | false,
): NormalizedParagraphSpacing | undefined {
  const resolved: {
    before?: number;
    after?: number;
    line?: NormalizedParagraphLineSpacing;
  } = {};
  applyParagraphSpacingUpdate(resolved, defaultSpacing);
  if (paragraphSpacing === false) return undefined;
  applyParagraphSpacingUpdate(resolved, paragraphSpacing);
  return Object.keys(resolved).length > 0 ? resolved : undefined;
}

export function renderParagraphProperties(
  template: string | undefined,
  prefix: string,
  alignment: TextAlignment | undefined,
  bullet?: NormalizedParagraphBullet,
  spacing?: NormalizedParagraphSpacing,
  level?: number,
  tabStops?: readonly NormalizedParagraphTabStop[],
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
  const leveled = updateParagraphAttribute(aligned, 'lvl', level && level > 0 ? String(level) : undefined);
  const spaced = renderParagraphSpacing(leveled, prefix, spacing);
  const bulleted = renderParagraphBullet(spaced, prefix, bullet, level ?? 0);
  return renderParagraphTabStops(bulleted, prefix, tabStops);
}

function renderParagraphTabStops(
  template: string,
  prefix: string,
  tabStops: readonly NormalizedParagraphTabStop[] | undefined,
): string {
  const properties = LosslessXmlDocument.parse(template);
  const root = requireParagraphPropertiesRoot(properties);
  const children = directChildren(root);
  const existing = children.filter(({ localName }) => localName === 'tabLst');
  if (tabStops === undefined && existing.length === 0) return template;
  for (const list of existing) properties.removeElement(list);
  if (tabStops === undefined) return properties.serialize();
  const stops = tabStops.map(({ positionEmu, alignment }) =>
    `<${prefix}tab pos="${positionEmu}" algn="${TAB_STOP_ALIGNMENT_TO_OOXML[alignment]}"/>`).join('');
  const listXml = `<${prefix}tabLst>${stops}</${prefix}tabLst>`;
  const follower = children.find((child) => TAB_STOP_INSERTION_FOLLOWERS.has(child.localName));
  if (follower) properties.replace(follower.start, follower.start, listXml);
  else properties.appendChildXml(root, listXml);
  return properties.serialize();
}

function renderParagraphSpacing(
  template: string,
  prefix: string,
  spacing: NormalizedParagraphSpacing | undefined,
): string {
  const properties = LosslessXmlDocument.parse(template);
  const root = requireParagraphPropertiesRoot(properties);
  const children = directChildren(root);
  const spacingChildren = children.filter(({ localName }) => SPACING_ELEMENT_NAMES.has(localName));
  if (!spacing && spacingChildren.length === 0) return template;
  for (const child of spacingChildren) properties.removeElement(child);
  if (!spacing) return properties.serialize();
  const spacingXml = renderParagraphSpacingXml(prefix, spacing);
  const firstRemainingChild = children.find(({ localName }) => !SPACING_ELEMENT_NAMES.has(localName));
  if (firstRemainingChild) properties.replace(firstRemainingChild.start, firstRemainingChild.start, spacingXml);
  else properties.appendChildXml(root, spacingXml);
  return properties.serialize();
}

function renderParagraphBullet(
  template: string,
  prefix: string,
  bullet: NormalizedParagraphBullet | undefined,
  level: number,
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
    const hangingEmu = Math.round(bullet.indent * 12700);
    const marginEmu = hangingEmu * (level + 1);
    if (marginEmu > 4032 * 12700) {
      throw new RangeError('Text bullet indent and level must not exceed 4032 points total');
    }
    withIndent = updateParagraphAttribute(withIndent, 'marL', String(marginEmu));
    withIndent = updateParagraphAttribute(withIndent, 'indent', String(-hangingEmu));
  } else if (hadActiveBullet && isRecognizedBulletIndentPair(margin, indent)) {
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

function isRecognizedBulletIndentPair(
  margin: number | undefined,
  indent: number | undefined,
): boolean {
  if (margin === 0 && indent === 0) return true;
  if (margin === undefined || indent === undefined || margin < 0 || indent >= 0) return false;
  const levelMultiplier = margin / -indent;
  return Number.isInteger(levelMultiplier) && levelMultiplier >= 1 && levelMultiplier <= 9;
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

function normalizeParagraphSpacingPoint(
  value: unknown,
  context: string,
  allowZero: boolean,
): number | false {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > MAX_SPACING_POINTS) {
    throw new RangeError(`${context} must be between 0 and ${MAX_SPACING_POINTS} points`);
  }
  const points = Math.round(value * 100) / 100;
  if (points === 0) {
    if (allowZero) return false;
    throw new RangeError(`${context} must be greater than zero`);
  }
  return points;
}

function normalizeParagraphLineSpacing(
  value: unknown,
  context: string,
): NormalizedParagraphLineSpacing {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object or false`);
  }
  const candidate = value as { factor?: unknown; kind?: unknown; points?: unknown };
  if (candidate.kind === 'exact') {
    assertSupportedKeys(value, ['kind', 'points'], context);
    return {
      kind: 'exact',
      points: normalizeParagraphSpacingPoint(candidate.points, `${context} points`, false) as number,
    };
  }
  if (candidate.kind === 'multiple') {
    assertSupportedKeys(value, ['factor', 'kind'], context);
    if (typeof candidate.factor !== 'number' || !Number.isFinite(candidate.factor)) {
      throw new TypeError(`${context} factor must be finite`);
    }
    if (candidate.factor <= 0 || candidate.factor > MAX_SPACING_FACTOR) {
      throw new RangeError(`${context} factor must be greater than zero and at most ${MAX_SPACING_FACTOR}`);
    }
    const factor = Math.round(candidate.factor * 100000) / 100000;
    if (factor === 0) throw new RangeError(`${context} factor must be greater than zero`);
    return { kind: 'multiple', factor };
  }
  throw new TypeError(`${context} kind must be exact or multiple`);
}

function applyParagraphSpacingUpdate(
  target: { before?: number; after?: number; line?: NormalizedParagraphLineSpacing },
  update: NormalizedParagraphSpacingUpdate | undefined,
): void {
  if (!update) return;
  for (const property of ['before', 'after', 'line'] as const) {
    if (!Object.hasOwn(update, property)) continue;
    const value = update[property];
    if (value === false || value === undefined) delete target[property];
    else if (property === 'line') target.line = value as NormalizedParagraphLineSpacing;
    else target[property] = value as number;
  }
}

function renderParagraphSpacingXml(
  prefix: string,
  spacing: NormalizedParagraphSpacing,
): string {
  const line = spacing.line
    ? spacing.line.kind === 'exact'
      ? `<${prefix}lnSpc><${prefix}spcPts val="${Math.round(spacing.line.points * 100)}"/></${prefix}lnSpc>`
      : `<${prefix}lnSpc><${prefix}spcPct val="${Math.round(spacing.line.factor * 100000)}"/></${prefix}lnSpc>`
    : '';
  const before = spacing.before === undefined
    ? ''
    : `<${prefix}spcBef><${prefix}spcPts val="${Math.round(spacing.before * 100)}"/></${prefix}spcBef>`;
  const after = spacing.after === undefined
    ? ''
    : `<${prefix}spcAft><${prefix}spcPts val="${Math.round(spacing.after * 100)}"/></${prefix}spcAft>`;
  return `${line}${before}${after}`;
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
    ['bold', 'color', 'fontFamily', 'fontSize', 'italic', 'underline'],
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
  const context = `Rich text run ${paragraphIndex},${runIndex}`;
  const color = candidate.color === undefined
    ? undefined
    : normalizeColor(candidate.color, `${context} color`);
  const underline = candidate.underline === undefined
    ? undefined
    : normalizeUnderline(candidate.underline, `${context} underline`);
  return {
    ...(candidate.fontFamily !== undefined ? { fontFamily: candidate.fontFamily } : {}),
    ...(candidate.fontSize !== undefined ? { fontSize: Math.round(candidate.fontSize * 100) / 100 } : {}),
    ...(candidate.bold !== undefined ? { bold: candidate.bold } : {}),
    ...(candidate.italic !== undefined ? { italic: candidate.italic } : {}),
    ...(color ? { color } : {}),
    ...(underline !== undefined ? { underline } : {}),
  };
}

function normalizeUnderline(value: unknown, context: string): RichTextUnderline | false {
  if (typeof value === 'boolean') return value ? { style: 'sng' } : false;
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be a boolean or configuration object`);
  }
  assertSupportedKeys(value, ['color', 'style'], context);
  const candidate = value as { color?: unknown; style?: unknown };
  if (candidate.style === undefined && candidate.color === undefined) {
    throw new TypeError(`${context} must provide style or color`);
  }
  if (
    candidate.style !== undefined
    && (typeof candidate.style !== 'string' || !UNDERLINE_STYLES.has(candidate.style as RichTextUnderlineStyle))
  ) {
    throw new TypeError(`${context} style is unsupported`);
  }
  const color = candidate.color === undefined
    ? undefined
    : normalizeColor(candidate.color, `${context} color`);
  return {
    style: (candidate.style as RichTextUnderlineStyle | undefined) ?? 'sng',
    ...(color ? { color } : {}),
  };
}

function normalizeColor(value: unknown, context: string): RichTextColor {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  assertSupportedKeys(value, ['kind', 'value'], context);
  const candidate = value as { kind?: unknown; value?: unknown };
  if (candidate.kind === 'srgb') {
    if (typeof candidate.value !== 'string' || !/^#?[\da-f]{6}$/i.test(candidate.value)) {
      throw new TypeError(`${context} sRGB value must contain six hex digits`);
    }
    return { kind: 'srgb', value: candidate.value.replace(/^#/, '').toUpperCase() };
  }
  if (candidate.kind === 'scheme') {
    if (typeof candidate.value !== 'string' || !SCHEME_COLORS.has(candidate.value)) {
      throw new TypeError(`${context} scheme value is unsupported`);
    }
    return { kind: 'scheme', value: candidate.value };
  }
  throw new TypeError(`${context} kind must be srgb or scheme`);
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
    style.underline === undefined
      ? ''
      : `u="${style.underline === false ? 'none' : style.underline === true ? 'sng' : style.underline.style}"`,
    'dirty="0"',
  ].filter(Boolean).join(' ');
  const color = style.color ?? { kind: 'scheme' as const, value: 'tx1' };
  const colorXml = renderColorChoice(color, prefix);
  const underlineColor = typeof style.underline === 'object' ? style.underline.color : undefined;
  const underlineFill = underlineColor
    ? `<${prefix}uFill><${prefix}solidFill>${renderColorChoice(underlineColor, prefix)}</${prefix}solidFill></${prefix}uFill>`
    : '';
  const latin = escapeXmlAttribute(style.fontFamily ?? '+mn-lt');
  const eastAsian = escapeXmlAttribute(style.fontFamily ?? '+mn-ea');
  const complexScript = escapeXmlAttribute(style.fontFamily ?? '+mn-cs');
  return `${softBreak}<${prefix}r><${prefix}rPr ${attributes}><${prefix}solidFill>${colorXml}</${prefix}solidFill>${underlineFill}<${prefix}latin typeface="${latin}"/><${prefix}ea typeface="${eastAsian}"/><${prefix}cs typeface="${complexScript}"/></${prefix}rPr><${prefix}t xml:space="preserve">${escapeXmlText(run.text)}</${prefix}t></${prefix}r>`;
}

function renderColorChoice(color: RichTextColor, prefix: string): string {
  return color.kind === 'srgb'
    ? `<${prefix}srgbClr val="${color.value}"/>`
    : `<${prefix}schemeClr val="${color.value}"/>`;
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

function readParagraphLevel(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): number | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  const level = readIntegerAttribute(xml, properties, 'lvl');
  return level !== undefined && level >= 1 && level <= 8 ? level : undefined;
}

function readParagraphBullet(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
  level: number,
): CharacterBullet | NumberedBullet | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  const indent = readBulletIndent(xml, properties, level);
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

function readParagraphSpacing(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): ParagraphSpacing | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  const lineElement = directChildren(properties, 'lnSpc')[0];
  const beforeElement = directChildren(properties, 'spcBef')[0];
  const afterElement = directChildren(properties, 'spcAft')[0];
  const line = lineElement ? readParagraphLineSpacing(xml, lineElement) : undefined;
  const before = beforeElement ? readPointSpacing(xml, beforeElement) : undefined;
  const after = afterElement ? readPointSpacing(xml, afterElement) : undefined;
  const spacing: ParagraphSpacing = {
    ...(before !== undefined ? { before } : {}),
    ...(after !== undefined ? { after } : {}),
    ...(line ? { line } : {}),
  };
  return Object.keys(spacing).length > 0 ? spacing : undefined;
}

function readParagraphTabStops(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): readonly ParagraphTabStop[] | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  const lists = directChildren(properties, 'tabLst');
  if (lists.length !== 1) return undefined;
  const children = directChildren(lists[0]!);
  if (children.some(({ localName }) => localName !== 'tab')) return undefined;
  const stops: ParagraphTabStop[] = [];
  for (const stop of children) {
    if (directChildren(stop).length > 0) return undefined;
    const positionEmu = readIntegerAttribute(xml, stop, 'pos');
    const alignmentValue = xml.attribute(stop, 'algn')?.value;
    const alignment = alignmentValue ? OOXML_TO_TAB_STOP_ALIGNMENT.get(alignmentValue) : undefined;
    if (
      positionEmu === undefined
      || positionEmu < MIN_COORDINATE_32
      || positionEmu > MAX_COORDINATE_32
      || alignment === undefined
    ) return undefined;
    stops.push({ position: positionEmu / EMU_PER_INCH, alignment });
  }
  return stops;
}

function readParagraphLineSpacing(
  xml: LosslessXmlDocument,
  container: XmlElement,
): ParagraphLineSpacing | undefined {
  const children = directChildren(container);
  if (children.length !== 1) return undefined;
  const choice = children[0]!;
  const value = readIntegerAttribute(xml, choice, 'val');
  if (value === undefined || value <= 0) return undefined;
  if (choice.localName === 'spcPts' && value <= MAX_SPACING_POINTS * 100) {
    return { kind: 'exact', points: value / 100 };
  }
  if (choice.localName === 'spcPct' && value <= MAX_SPACING_FACTOR * 100000) {
    return { kind: 'multiple', factor: value / 100000 };
  }
  return undefined;
}

function readPointSpacing(
  xml: LosslessXmlDocument,
  container: XmlElement,
): number | undefined {
  const children = directChildren(container);
  if (children.length !== 1 || children[0]?.localName !== 'spcPts') return undefined;
  const value = readIntegerAttribute(xml, children[0], 'val');
  if (value === undefined || value <= 0 || value > MAX_SPACING_POINTS * 100) return undefined;
  return value / 100;
}

function readBulletIndent(
  xml: LosslessXmlDocument,
  properties: XmlElement,
  level: number,
): number | undefined {
  const margin = readIntegerAttribute(xml, properties, 'marL');
  if (margin === undefined || margin < 0 || margin > 4032 * 12700) return undefined;
  const points = Math.round((margin / (level + 1) / 12700) * 100) / 100;
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
  const underline = readUnderline(xml, properties);
  const style: RichTextRunStyle = {
    ...(fontFamily !== undefined ? { fontFamily } : {}),
    ...(Number.isFinite(size) && size > 0 ? { fontSize: size / 100 } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(color ? { color } : {}),
    ...(underline !== undefined ? { underline } : {}),
  };
  return Object.keys(style).length > 0 ? style : undefined;
}

function readUnderline(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): RichTextUnderline | false | undefined {
  const value = xml.attribute(properties, 'u')?.value;
  if (value === undefined) return undefined;
  if (value === 'none') return false;
  if (!UNDERLINE_STYLES.has(value as RichTextUnderlineStyle)) return undefined;
  const fill = directChildren(properties, 'uFill');
  const color = fill.length === 1 ? readDirectSolidColor(xml, fill[0]!) : undefined;
  return {
    style: value as RichTextUnderlineStyle,
    ...(color ? { color } : {}),
  };
}

function readDirectSolidColor(
  xml: LosslessXmlDocument,
  container: XmlElement,
): RichTextColor | undefined {
  const containerChildren = directChildren(container);
  if (containerChildren.length !== 1 || containerChildren[0]?.localName !== 'solidFill') return undefined;
  const fillChildren = directChildren(containerChildren[0]);
  if (fillChildren.length !== 1) return undefined;
  const color = fillChildren[0]!;
  if (directChildren(color).length > 0) return undefined;
  const value = xml.attribute(color, 'val')?.value;
  if (color.localName === 'srgbClr' && value && /^[\da-f]{6}$/i.test(value)) {
    return { kind: 'srgb', value: value.toUpperCase() };
  }
  if (color.localName === 'schemeClr' && value && SCHEME_COLORS.has(value)) {
    return { kind: 'scheme', value };
  }
  return undefined;
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
