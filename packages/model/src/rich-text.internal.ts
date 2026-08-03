import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  normalizeHyperlink,
  readTextRunHyperlinkBinding,
  renderShapeHyperlink,
  type NormalizedHyperlink,
  type ShapeHyperlinkReadContext,
} from './shape-hyperlink.internal.js';
import type {
  CharacterBullet,
  NumberedBullet,
  NumberingStyle,
  ParagraphBullet,
  ParagraphLineSpacing,
  ParagraphSpacing,
  ParagraphTabStop,
  ParagraphTabStopAlignment,
  RichTextBaseline,
  RichTextColor,
  RichTextGlow,
  RichTextOutline,
  RichTextParagraph,
  RichTextRun,
  RichTextRunStyle,
  RichTextStrikeStyle,
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
const EMU_PER_POINT = 12_700;
const MAX_PARAGRAPH_MARGIN_EMU = 4032 * EMU_PER_POINT;
const MAX_LINE_WIDTH_EMU = 20_116_800;
const MAX_POSITIVE_COORDINATE_EMU = 27_273_042_316_900;
const PERCENT_SCALE = 100_000;
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

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

interface NormalizedRichTextRun {
  readonly text: string;
  readonly style?: RichTextRunStyle;
  readonly softBreakBefore?: boolean;
}

interface NormalizedRichTextRunInput {
  readonly run: NormalizedRichTextRun;
  readonly breakLine: boolean;
}

interface NormalizedRichTextParagraph {
  readonly runs: readonly NormalizedRichTextRun[];
  readonly align?: TextAlignment;
  readonly rtl?: boolean;
  readonly marginLeft?: number | false;
  readonly marginRight?: number | false;
  readonly indent?: number | false;
  readonly bullet?: NormalizedParagraphBullet | false;
  readonly level?: number;
  readonly spacing?: NormalizedParagraphSpacingUpdate | false;
  readonly tabStops?: readonly NormalizedParagraphTabStop[] | false;
}

export interface ReadRichTextRunHyperlinkBinding {
  readonly hyperlink: NormalizedHyperlink;
  readonly relationshipId: string;
}

export interface ReadRichTextState {
  readonly paragraphs: readonly RichTextParagraph[];
  readonly runHyperlinkBindings:
    readonly (readonly (ReadRichTextRunHyperlinkBinding | undefined)[])[];
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

const STRIKE_STYLES = new Set<RichTextStrikeStyle>(['sngStrike', 'dblStrike']);

export function normalizeRichText(value: unknown): readonly NormalizedRichTextParagraph[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new TypeError('Rich text must contain at least one paragraph');
  }
  return value.flatMap((paragraph, paragraphIndex) => {
    if (!paragraph || typeof paragraph !== 'object' || Array.isArray(paragraph)) {
      throw new TypeError(`Rich text paragraph ${paragraphIndex} must be an object`);
    }
    assertSupportedKeys(
      paragraph,
      ['align', 'bullet', 'indent', 'level', 'marginLeft', 'marginRight', 'rtl', 'runs', 'spacing', 'tabStops'],
      `Rich text paragraph ${paragraphIndex}`,
    );
    const candidate = paragraph as {
      align?: unknown;
      bullet?: unknown;
      indent?: unknown;
      level?: unknown;
      marginLeft?: unknown;
      marginRight?: unknown;
      rtl?: unknown;
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
    const indent = candidate.indent === undefined
      ? undefined
      : candidate.indent === false
        ? false
        : normalizeParagraphIndent(candidate.indent, `Rich text paragraph ${paragraphIndex} indent`);
    const level = candidate.level === undefined
      ? undefined
      : normalizeParagraphLevel(candidate.level, `Rich text paragraph ${paragraphIndex} level`);
    const marginLeft = candidate.marginLeft === undefined
      ? undefined
      : candidate.marginLeft === false
        ? false
        : normalizeParagraphMarginLeft(
            candidate.marginLeft,
            `Rich text paragraph ${paragraphIndex} marginLeft`,
          );
    const marginRight = candidate.marginRight === undefined
      ? undefined
      : candidate.marginRight === false
        ? false
        : normalizeParagraphMarginRight(
            candidate.marginRight,
            `Rich text paragraph ${paragraphIndex} marginRight`,
          );
    const rtl = candidate.rtl === undefined
      ? undefined
      : normalizeParagraphRtl(candidate.rtl, `Rich text paragraph ${paragraphIndex} rtl`);
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
    const properties: Omit<NormalizedRichTextParagraph, 'runs'> = {
      ...(align ? { align } : {}),
      ...(bullet !== undefined ? { bullet } : {}),
      ...(indent !== undefined ? { indent } : {}),
      ...(level !== undefined ? { level } : {}),
      ...(marginLeft !== undefined ? { marginLeft } : {}),
      ...(marginRight !== undefined ? { marginRight } : {}),
      ...(rtl !== undefined ? { rtl } : {}),
      ...(spacing !== undefined ? { spacing } : {}),
      ...(tabStops !== undefined ? { tabStops } : {}),
    };
    return splitNormalizedRichTextParagraph(
      properties,
      runs.map((run, runIndex) => normalizeRun(run, paragraphIndex, runIndex)),
    );
  });
}

function splitNormalizedRichTextParagraph(
  properties: Omit<NormalizedRichTextParagraph, 'runs'>,
  runs: readonly NormalizedRichTextRunInput[],
): readonly NormalizedRichTextParagraph[] {
  const paragraphs: NormalizedRichTextParagraph[] = [];
  let current: NormalizedRichTextRun[] = [];
  for (const [index, { run, breakLine }] of runs.entries()) {
    current.push(run);
    if (breakLine && index + 1 < runs.length) {
      paragraphs.push({ ...properties, runs: current });
      current = [];
    }
  }
  paragraphs.push({ ...properties, runs: current });
  return paragraphs;
}

export function richTextParagraphsEqual(
  left: unknown,
  right: unknown,
): boolean {
  return richTextValuesEqual(left, right);
}

function richTextValuesEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left)
      && Array.isArray(right)
      && left.length === right.length
      && left.every((value, index) => richTextValuesEqual(value, right[index]));
  }
  if (
    !left
    || !right
    || typeof left !== 'object'
    || typeof right !== 'object'
  ) return false;
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord);
  const rightKeys = Object.keys(rightRecord);
  return leftKeys.length === rightKeys.length
    && leftKeys.every((key) =>
      Object.hasOwn(rightRecord, key)
      && richTextValuesEqual(leftRecord[key], rightRecord[key]));
}

interface RenderRichTextOptions {
  readonly prefix?: string;
  readonly defaultLanguage?: string;
  readonly defaultFontFamily?: string;
  readonly defaultFontSize?: number;
  readonly defaultBold?: boolean;
  readonly defaultColor?: Readonly<RichTextColor>;
  readonly defaultAlign?: TextAlignment;
  readonly defaultRtl?: boolean;
  readonly defaultBullet?: NormalizedParagraphBullet | false;
  readonly defaultIndent?: number;
  readonly defaultLevel?: number;
  readonly defaultMarginLeft?: number;
  readonly defaultMarginRight?: number;
  readonly defaultSpacing?: NormalizedParagraphSpacingUpdate;
  readonly defaultTabStops?: readonly NormalizedParagraphTabStop[];
  readonly defaultHyperlink?: NormalizedHyperlink;
  readonly hyperlinkRelationshipId?: string;
  readonly runHyperlinkRelationshipIds?: RichTextRunHyperlinkRelationshipIds;
  readonly declareHyperlinkRelationshipNamespace?: boolean;
  readonly suppressDefaultColorForHyperlinks?: boolean;
  readonly paragraphProperties?: readonly (string | undefined)[];
  readonly endParagraphProperties?: string;
}

interface RenderRunOptions {
  readonly prefix: string;
  readonly defaultLanguage?: string;
  readonly defaultFontFamily?: string;
  readonly defaultFontSize?: number;
  readonly defaultBold?: boolean;
  readonly defaultColor?: Readonly<RichTextColor>;
  readonly hyperlink?: NormalizedHyperlink;
  readonly hyperlinkRelationshipId?: string;
  readonly declareHyperlinkRelationshipNamespace: boolean;
  readonly suppressDefaultColorForHyperlinks: boolean;
}

export type RichTextRunHyperlinkRelationshipIds =
  readonly (readonly (string | undefined)[])[];

export function renderRichTextParagraphs(
  paragraphs: readonly NormalizedRichTextParagraph[],
  options: RenderRichTextOptions = {},
): string {
  if (
    (options.defaultHyperlink === undefined)
    !== (options.hyperlinkRelationshipId === undefined)
  ) {
    throw new TypeError('Rich text hyperlink and relationship ID must be supplied together');
  }
  if (options.runHyperlinkRelationshipIds !== undefined) {
    if (options.runHyperlinkRelationshipIds.length !== paragraphs.length) {
      throw new TypeError('Rich text run hyperlink relationship IDs must match the paragraph count');
    }
    for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
      if (options.runHyperlinkRelationshipIds[paragraphIndex]?.length !== paragraph.runs.length) {
        throw new TypeError(
          `Rich text run hyperlink relationship IDs must match paragraph ${paragraphIndex} run count`,
        );
      }
    }
  }
  const prefix = options.prefix ?? 'a:';
  const defaultLanguage = options.defaultLanguage ?? 'en-US';
  const defaultEndProperties = renderDefaultEndParagraphProperties(
    prefix,
    defaultLanguage,
    options.defaultFontFamily,
    options.defaultFontSize,
  );
  return paragraphs
    .map(({ align, bullet, indent, level, marginLeft, marginRight, rtl, runs, spacing, tabStops }, paragraphIndex) => {
      const resolvedBullet = bullet === false
        ? undefined
        : bullet ?? (options.defaultBullet === false ? undefined : options.defaultBullet);
      const resolvedIndent = indent === false
        ? false
        : indent ?? options.defaultIndent;
      const resolvedMarginLeft = marginLeft === false
        ? false
        : marginLeft ?? options.defaultMarginLeft;
      const resolvedMarginRight = marginRight === false
        ? false
        : marginRight ?? options.defaultMarginRight;
      if (resolvedBullet && typeof resolvedMarginLeft === 'number') {
        throw new TypeError('Paragraph left margin cannot be combined with an active bullet');
      }
      if (resolvedBullet && typeof resolvedIndent === 'number') {
        throw new TypeError('Paragraph indent cannot be combined with an active bullet');
      }
      return `<${prefix}p>${renderParagraphProperties(
        options.paragraphProperties?.[paragraphIndex] ?? options.paragraphProperties?.[0],
        prefix,
        align ?? options.defaultAlign,
        rtl ?? options.defaultRtl,
        resolvedBullet,
        resolveParagraphSpacing(options.defaultSpacing, spacing),
        level ?? options.defaultLevel,
        tabStops === false ? undefined : tabStops ?? options.defaultTabStops,
        resolvedMarginLeft,
        resolvedMarginRight,
        resolvedIndent,
      )}${runs
        .map((run, runIndex) => {
          const localHyperlink = run.style?.hyperlink;
          const directHyperlink = localHyperlink === false ? undefined : localHyperlink;
          const directRelationshipId =
            options.runHyperlinkRelationshipIds?.[paragraphIndex]?.[runIndex];
          if ((directHyperlink === undefined) !== (directRelationshipId === undefined)) {
            throw new TypeError(
              `Rich text run ${paragraphIndex},${runIndex} hyperlink and relationship ID must be supplied together`,
            );
          }
          const hyperlink = localHyperlink === false
            ? undefined
            : localHyperlink ?? options.defaultHyperlink;
          const relationshipId = directHyperlink === undefined
            ? hyperlink === undefined ? undefined : options.hyperlinkRelationshipId
            : directRelationshipId;
          return renderRun(run, {
            prefix,
            ...(options.defaultLanguage === undefined
              ? {}
              : { defaultLanguage: options.defaultLanguage }),
            ...(options.defaultFontFamily === undefined
              ? {}
              : { defaultFontFamily: options.defaultFontFamily }),
            ...(options.defaultFontSize === undefined
              ? {}
              : { defaultFontSize: options.defaultFontSize }),
            ...(options.defaultBold === undefined
              ? {}
              : { defaultBold: options.defaultBold }),
            ...(options.defaultColor === undefined
              ? {}
              : { defaultColor: options.defaultColor }),
            ...(hyperlink === undefined ? {} : { hyperlink }),
            ...(relationshipId === undefined
              ? {}
              : { hyperlinkRelationshipId: relationshipId }),
            declareHyperlinkRelationshipNamespace:
              options.declareHyperlinkRelationshipNamespace ?? false,
            suppressDefaultColorForHyperlinks:
              options.suppressDefaultColorForHyperlinks ?? false,
          });
        })
        .join('')}${options.endParagraphProperties ?? defaultEndProperties}</${prefix}p>`;
    })
    .join('');
}

function renderDefaultEndParagraphProperties(
  prefix: string,
  language: string,
  fontFamily?: string,
  fontSize?: number,
): string {
  const languageValue = escapeXmlAttribute(language);
  const sizeAttribute = fontSize === undefined
    ? ''
    : ` sz="${Math.round(fontSize * 100)}"`;
  if (fontFamily === undefined) {
    return `<${prefix}endParaRPr lang="${languageValue}"${sizeAttribute} dirty="0"/>`;
  }
  const typeface = escapeXmlAttribute(fontFamily);
  return `<${prefix}endParaRPr lang="${languageValue}"${sizeAttribute} dirty="0">` +
    `<${prefix}latin typeface="${typeface}"/>` +
    `<${prefix}ea typeface="${typeface}"/>` +
    `<${prefix}cs typeface="${typeface}"/>` +
    `</${prefix}endParaRPr>`;
}

export function readRichText(
  xml: LosslessXmlDocument,
  element: XmlElement,
  context: ShapeHyperlinkReadContext,
): readonly RichTextParagraph[] {
  return readRichTextState(xml, element, context).paragraphs;
}

export function readRichTextState(
  xml: LosslessXmlDocument,
  element: XmlElement,
  context: ShapeHyperlinkReadContext,
): ReadRichTextState {
  const textBody = directChildren(element, 'txBody')[0];
  if (!textBody) return { paragraphs: [], runHyperlinkBindings: [] };
  const paragraphs: RichTextParagraph[] = [];
  const runHyperlinkBindings:
    (readonly (ReadRichTextRunHyperlinkBinding | undefined)[])[] = [];
  for (const paragraph of directChildren(textBody, 'p')) {
    const align = readParagraphAlignment(xml, paragraph);
    const rtl = readParagraphRtl(xml, paragraph);
    const marginLeft = readParagraphMarginLeft(xml, paragraph);
    const marginRight = readParagraphMarginRight(xml, paragraph);
    const indent = readParagraphIndent(xml, paragraph);
    const level = readParagraphLevel(xml, paragraph);
    const bullet = readParagraphBullet(xml, paragraph, level ?? 0);
    const spacing = readParagraphSpacing(xml, paragraph);
    const tabStops = readParagraphTabStops(xml, paragraph);
    const runs = readRuns(xml, paragraph, context);
    paragraphs.push({
      runs: runs.runs,
      ...(align ? { align } : {}),
      ...(rtl !== undefined ? { rtl } : {}),
      ...(marginLeft !== undefined ? { marginLeft } : {}),
      ...(marginRight !== undefined ? { marginRight } : {}),
      ...(indent !== undefined ? { indent } : {}),
      ...(bullet ? { bullet } : {}),
      ...(level !== undefined ? { level } : {}),
      ...(spacing ? { spacing } : {}),
      ...(tabStops !== undefined ? { tabStops } : {}),
    });
    runHyperlinkBindings.push(runs.hyperlinkBindings);
  }
  return { paragraphs, runHyperlinkBindings };
}

export function replaceRichText(
  xml: LosslessXmlDocument,
  element: XmlElement,
  paragraphs: readonly NormalizedRichTextParagraph[],
  partUri: string,
  runHyperlinkRelationshipIds: RichTextRunHyperlinkRelationshipIds,
): string {
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
    runHyperlinkRelationshipIds,
    declareHyperlinkRelationshipNamespace: true,
    ...(endProperties ? { endParagraphProperties: xml.original(endProperties) } : {}),
  });
  xml.replaceElement(template, replacement);
  for (const extra of existing.slice(1)) xml.removeElement(extra);
  return xml.serialize();
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

export function normalizeParagraphRtl(value: unknown, context: string): boolean {
  if (typeof value !== 'boolean') throw new TypeError(`${context} must be a boolean`);
  return value;
}

export function normalizeParagraphMarginLeft(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > 4032) {
    throw new RangeError(`${context} must be between 0 and 4032 points`);
  }
  return value;
}

export function normalizeParagraphMarginRight(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > 4032) {
    throw new RangeError(`${context} must be between 0 and 4032 points`);
  }
  return value;
}

export function normalizeParagraphIndent(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < -4032 || value > 4032) {
    throw new RangeError(`${context} must be between -4032 and 4032 points`);
  }
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
  rtl: boolean | undefined,
  bullet?: NormalizedParagraphBullet,
  spacing?: NormalizedParagraphSpacing,
  level?: number,
  tabStops?: readonly NormalizedParagraphTabStop[],
  marginLeft?: number | false,
  marginRight?: number | false,
  indent?: number | false,
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
  const directed = updateParagraphAttribute(aligned, 'rtl', rtl === undefined ? undefined : rtl ? '1' : '0');
  const leveled = updateParagraphAttribute(directed, 'lvl', level && level > 0 ? String(level) : undefined);
  const spaced = renderParagraphSpacing(leveled, prefix, spacing);
  const bulleted = renderParagraphBullet(spaced, prefix, bullet, level ?? 0);
  const marginalized = renderParagraphLeftMargin(
    bulleted,
    marginLeft,
    bullet !== undefined,
    template === undefined,
  );
  const rightMarginalized = renderParagraphRightMargin(
    marginalized,
    marginRight,
    template === undefined,
  );
  const indented = renderParagraphIndent(
    rightMarginalized,
    indent,
    bullet !== undefined,
    template === undefined,
  );
  return renderParagraphTabStops(indented, prefix, tabStops);
}

function renderParagraphLeftMargin(
  template: string,
  marginLeft: number | false | undefined,
  hasActiveBullet: boolean,
  isNewParagraph: boolean,
): string {
  if (hasActiveBullet) return template;
  if (marginLeft === undefined && isNewParagraph) return template;
  return updateParagraphAttribute(
    template,
    'marL',
    typeof marginLeft === 'number' ? String(Math.round(marginLeft * EMU_PER_POINT)) : undefined,
  );
}

function renderParagraphRightMargin(
  template: string,
  marginRight: number | false | undefined,
  isNewParagraph: boolean,
): string {
  if (marginRight === undefined && isNewParagraph) return template;
  return updateParagraphAttribute(
    template,
    'marR',
    typeof marginRight === 'number' ? String(Math.round(marginRight * EMU_PER_POINT)) : undefined,
  );
}

function renderParagraphIndent(
  template: string,
  indent: number | false | undefined,
  hasActiveBullet: boolean,
  isNewParagraph: boolean,
): string {
  if (hasActiveBullet) return template;
  if (indent === undefined && isNewParagraph) return template;
  return updateParagraphAttribute(
    template,
    'indent',
    typeof indent === 'number' ? String(Math.round(indent * EMU_PER_POINT)) : undefined,
  );
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

function normalizeRun(
  value: unknown,
  paragraphIndex: number,
  runIndex: number,
): NormalizedRichTextRunInput {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} must be an object`);
  }
  assertSupportedKeys(
    value,
    ['breakLine', 'softBreakBefore', 'style', 'text'],
    `Rich text run ${paragraphIndex},${runIndex}`,
  );
  const candidate = value as {
    breakLine?: unknown;
    text?: unknown;
    style?: unknown;
    softBreakBefore?: unknown;
  };
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
  if (candidate.breakLine !== undefined && typeof candidate.breakLine !== 'boolean') {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} breakLine must be a boolean`);
  }
  const style = candidate.style === undefined
    ? undefined
    : normalizeStyle(candidate.style, paragraphIndex, runIndex, candidate.text);
  return {
    run: {
      text: candidate.text,
      ...(style ? { style } : {}),
      ...(candidate.softBreakBefore !== undefined ? { softBreakBefore: candidate.softBreakBefore } : {}),
    },
    breakLine: candidate.breakLine ?? false,
  };
}

function normalizeStyle(
  value: unknown,
  paragraphIndex: number,
  runIndex: number,
  runText: string,
): RichTextRunStyle {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`Rich text run ${paragraphIndex},${runIndex} style must be an object`);
  }
  assertSupportedKeys(
    value,
    ['baseline', 'bold', 'characterSpacing', 'color', 'fontFamily', 'fontSize', 'glow', 'highlight', 'hyperlink', 'italic', 'lang', 'outline', 'strike', 'transparency', 'underline'],
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
  const language = candidate.lang === undefined
    ? undefined
    : normalizeTextLanguage(candidate.lang, `${context} lang`);
  const baseline = candidate.baseline === undefined
    ? undefined
    : normalizeBaseline(candidate.baseline, `${context} baseline`);
  const characterSpacing = candidate.characterSpacing === undefined
    ? undefined
    : normalizeCharacterSpacing(candidate.characterSpacing, `${context} characterSpacing`);
  const color = candidate.color === undefined
    ? undefined
    : normalizeRichTextColor(candidate.color, `${context} color`);
  const transparency = candidate.transparency === undefined
    ? undefined
    : normalizeTextTransparency(candidate.transparency, `${context} transparency`);
  const glow = candidate.glow === undefined
    ? undefined
    : normalizeGlow(candidate.glow, `${context} glow`);
  const highlight = candidate.highlight === undefined
    ? undefined
    : normalizeRichTextColor(candidate.highlight, `${context} highlight`);
  const hyperlink = candidate.hyperlink === undefined
    ? undefined
    : candidate.hyperlink === false
      ? false
      : normalizeHyperlink(candidate.hyperlink, `${context} hyperlink`);
  if (hyperlink !== undefined && hyperlink !== false && runText.length === 0) {
    throw new TypeError(`${context} hyperlink requires non-empty text`);
  }
  const outline = candidate.outline === undefined
    ? undefined
    : normalizeOutline(candidate.outline, `${context} outline`);
  const underline = candidate.underline === undefined
    ? undefined
    : normalizeUnderline(candidate.underline, `${context} underline`);
  const strike = candidate.strike === undefined
    ? undefined
    : normalizeStrike(candidate.strike, `${context} strike`);
  return {
    ...(candidate.fontFamily !== undefined ? { fontFamily: candidate.fontFamily } : {}),
    ...(candidate.fontSize !== undefined ? { fontSize: Math.round(candidate.fontSize * 100) / 100 } : {}),
    ...(language !== undefined ? { lang: language } : {}),
    ...(baseline !== undefined ? { baseline } : {}),
    ...(characterSpacing !== undefined ? { characterSpacing } : {}),
    ...(candidate.bold !== undefined ? { bold: candidate.bold } : {}),
    ...(candidate.italic !== undefined ? { italic: candidate.italic } : {}),
    ...(color ? { color } : {}),
    ...(transparency !== undefined ? { transparency } : {}),
    ...(glow ? { glow } : {}),
    ...(highlight ? { highlight } : {}),
    ...(hyperlink !== undefined ? { hyperlink } : {}),
    ...(outline ? { outline } : {}),
    ...(underline !== undefined ? { underline } : {}),
    ...(strike !== undefined ? { strike } : {}),
  };
}

function normalizeTextTransparency(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  if (value < 0 || value > 100) {
    throw new RangeError(`${context} must be between 0 and 100 percent`);
  }
  const alpha = Math.round((100 - value) * 1_000);
  return 100 - alpha / 1_000;
}

export function normalizeTextLanguage(value: unknown, context: string): string {
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`${context} must be a non-empty string`);
  }
  if (containsInvalidXmlCharacter(value)) {
    throw new TypeError(`${context} contains invalid XML characters`);
  }
  return value;
}

function normalizeCharacterSpacing(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be finite`);
  }
  const raw = Math.round(value * 100);
  if (raw < MIN_COORDINATE_32 || raw > MAX_COORDINATE_32) {
    throw new RangeError(`${context} must fit the OOXML Int32 point range`);
  }
  return raw / 100;
}

function normalizeBaseline(value: unknown, context: string): RichTextBaseline {
  if (value === 'superscript' || value === 'subscript') return value;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${context} must be a finite number, superscript, or subscript`);
  }
  const raw = Math.round(value * 1_000);
  if (raw < MIN_COORDINATE_32 || raw > MAX_COORDINATE_32) {
    throw new RangeError(`${context} must fit the OOXML Int32 percentage range`);
  }
  if (raw === 30_000) return 'superscript';
  if (raw === -40_000) return 'subscript';
  return raw / 1_000;
}

function normalizeGlow(value: unknown, context: string): RichTextGlow {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  assertSupportedKeys(value, ['color', 'opacity', 'size'], context);
  const candidate = value as { color?: unknown; opacity?: unknown; size?: unknown };
  if (candidate.opacity === undefined || candidate.size === undefined) {
    throw new TypeError(`${context} must provide opacity and size`);
  }
  if (typeof candidate.size !== 'number' || !Number.isFinite(candidate.size)) {
    throw new TypeError(`${context} size must be finite`);
  }
  if (typeof candidate.opacity !== 'number' || !Number.isFinite(candidate.opacity)) {
    throw new TypeError(`${context} opacity must be finite`);
  }
  const radiusEmu = Math.round(candidate.size * EMU_PER_POINT);
  if (candidate.size < 0 || radiusEmu > MAX_POSITIVE_COORDINATE_EMU) {
    throw new RangeError(`${context} size must fit the OOXML 0 to 2147483647 point range`);
  }
  if (candidate.opacity < 0 || candidate.opacity > 1) {
    throw new RangeError(`${context} opacity must be between 0 and 1`);
  }
  return {
    color: candidate.color === undefined
      ? { kind: 'srgb', value: 'FFFFFF' }
      : normalizeRichTextColor(candidate.color, `${context} color`),
    opacity: Math.round(candidate.opacity * PERCENT_SCALE) / PERCENT_SCALE,
    size: radiusEmu / EMU_PER_POINT,
  };
}

function normalizeOutline(value: unknown, context: string): RichTextOutline {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  assertSupportedKeys(value, ['color', 'size'], context);
  const candidate = value as { color?: unknown; size?: unknown };
  if (candidate.color === undefined || candidate.size === undefined) {
    throw new TypeError(`${context} must provide color and size`);
  }
  if (typeof candidate.size !== 'number' || !Number.isFinite(candidate.size)) {
    throw new TypeError(`${context} size must be finite`);
  }
  const widthEmu = Math.round(candidate.size * EMU_PER_POINT);
  if (candidate.size < 0 || widthEmu > MAX_LINE_WIDTH_EMU) {
    throw new RangeError(`${context} size must fit the OOXML 0 to 1584 point range`);
  }
  return {
    color: normalizeRichTextColor(candidate.color, `${context} color`),
    size: widthEmu / EMU_PER_POINT,
  };
}

function normalizeStrike(value: unknown, context: string): RichTextStrikeStyle | false {
  if (typeof value === 'boolean') return value ? 'sngStrike' : false;
  if (typeof value === 'string' && STRIKE_STYLES.has(value as RichTextStrikeStyle)) {
    return value as RichTextStrikeStyle;
  }
  throw new TypeError(`${context} must be a boolean, sngStrike, or dblStrike`);
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
    : normalizeRichTextColor(candidate.color, `${context} color`);
  return {
    style: (candidate.style as RichTextUnderlineStyle | undefined) ?? 'sng',
    ...(color ? { color } : {}),
  };
}

export function normalizeRichTextColor(
  value: unknown,
  context: string,
): Readonly<RichTextColor> {
  const candidate = readColorDataObject(value, context);
  if (candidate.kind === 'srgb') {
    if (typeof candidate.value !== 'string' || !/^#?[\da-f]{6}$/i.test(candidate.value)) {
      throw new TypeError(`${context} sRGB value must contain six hex digits`);
    }
    return Object.freeze({
      kind: 'srgb',
      value: candidate.value.replace(/^#/, '').toUpperCase(),
    });
  }
  if (candidate.kind === 'scheme') {
    if (typeof candidate.value !== 'string' || !SCHEME_COLORS.has(candidate.value)) {
      throw new TypeError(`${context} scheme value is unsupported`);
    }
    return Object.freeze({ kind: 'scheme', value: candidate.value });
  }
  throw new TypeError(`${context} kind must be srgb or scheme`);
}

function readColorDataObject(
  value: unknown,
  context: string,
): { readonly kind: unknown; readonly value: unknown } {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be an object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const keys = Reflect.ownKeys(value);
  if (keys.some((key) => typeof key !== 'string' || !['kind', 'value'].includes(key))) {
    throw new TypeError(`${context} contains unsupported properties`);
  }
  const kind = Object.getOwnPropertyDescriptor(value, 'kind');
  const colorValue = Object.getOwnPropertyDescriptor(value, 'value');
  if (
    !kind
    || !Object.hasOwn(kind, 'value')
    || !colorValue
    || !Object.hasOwn(colorValue, 'value')
  ) {
    throw new TypeError(`${context} must contain kind and value data properties`);
  }
  return { kind: kind.value, value: colorValue.value };
}

function renderRun(
  run: NormalizedRichTextRun,
  options: RenderRunOptions,
): string {
  const { prefix } = options;
  const softBreak = run.softBreakBefore ? `<${prefix}br/>` : '';
  if (run.text.length === 0 && run.style === undefined) return softBreak;
  const style = run.style ?? {};
  const language = style.lang ?? options.defaultLanguage ?? 'en-US';
  const explicitLanguage = style.lang !== undefined || options.defaultLanguage !== undefined;
  const fontFamily = style.fontFamily ?? options.defaultFontFamily;
  const fontSize = style.fontSize ?? options.defaultFontSize;
  const bold = style.bold ?? options.defaultBold;
  const languageAttributes = `lang="${escapeXmlAttribute(language)}"${
    explicitLanguage ? ' altLang="en-US"' : ''
  }`;
  const attributes = [
    languageAttributes,
    fontSize === undefined ? '' : `sz="${Math.round(fontSize * 100)}"`,
    style.baseline === undefined
      ? ''
      : `baseline="${style.baseline === 'superscript' ? 30_000 : style.baseline === 'subscript' ? -40_000 : Math.round(style.baseline * 1_000)}"`,
    style.characterSpacing === undefined ? '' : `spc="${Math.round(style.characterSpacing * 100)}"`,
    style.characterSpacing === undefined ? '' : 'kern="0"',
    bold === undefined ? '' : `b="${bold ? 1 : 0}"`,
    style.italic === undefined ? '' : `i="${style.italic ? 1 : 0}"`,
    style.strike === undefined
      ? ''
      : `strike="${style.strike === false ? 'noStrike' : style.strike === true ? 'sngStrike' : style.strike}"`,
    style.underline === undefined
      ? options.hyperlink === undefined ? '' : 'u="sng"'
      : `u="${style.underline === false ? 'none' : style.underline === true ? 'sng' : style.underline.style}"`,
    'dirty="0"',
  ].filter(Boolean).join(' ');
  const suppressOuterColor = options.suppressDefaultColorForHyperlinks
    && options.hyperlink !== undefined
    && style.color === undefined;
  const color = style.color
    ?? (suppressOuterColor ? undefined : options.defaultColor)
    ?? (suppressOuterColor ? undefined : { kind: 'scheme' as const, value: 'tx1' });
  const solidFill = color === undefined
    ? ''
    : `<${prefix}solidFill>${renderMainTextColorChoice(
        color,
        prefix,
        style.transparency,
      )}</${prefix}solidFill>`;
  const outline = style.outline
    ? `<${prefix}ln w="${Math.round(style.outline.size * EMU_PER_POINT)}"><${prefix}solidFill>${renderColorChoice(style.outline.color, prefix)}</${prefix}solidFill></${prefix}ln>`
    : '';
  const glow = style.glow
    ? `<${prefix}effectLst><${prefix}glow rad="${Math.round(style.glow.size * EMU_PER_POINT)}">${renderGlowColorChoice(style.glow, prefix)}</${prefix}glow></${prefix}effectLst>`
    : '';
  const highlight = style.highlight
    ? `<${prefix}highlight>${renderColorChoice(style.highlight, prefix)}</${prefix}highlight>`
    : '';
  const underlineColor = typeof style.underline === 'object' ? style.underline.color : undefined;
  const underlineFill = underlineColor
    ? `<${prefix}uFill><${prefix}solidFill>${renderColorChoice(underlineColor, prefix)}</${prefix}solidFill></${prefix}uFill>`
    : '';
  const latin = escapeXmlAttribute(fontFamily ?? '+mn-lt');
  const eastAsian = escapeXmlAttribute(fontFamily ?? '+mn-ea');
  const complexScript = escapeXmlAttribute(fontFamily ?? '+mn-cs');
  const drawingPrefix = prefix.endsWith(':') ? prefix.slice(0, -1) : prefix;
  const relationshipPrefix = drawingPrefix === 'r' ? 'rel' : 'r';
  const renderedHyperlink = options.hyperlink === undefined
    ? ''
    : renderShapeHyperlink(
        options.hyperlink,
        options.hyperlinkRelationshipId!,
        { drawing: drawingPrefix, relationship: relationshipPrefix },
      );
  const hyperlinkXml = !options.declareHyperlinkRelationshipNamespace || renderedHyperlink === ''
    ? renderedHyperlink
    : renderedHyperlink.replace(
        ' ',
        ` xmlns:${relationshipPrefix}="${RELATIONSHIP_NAMESPACE}" `,
      );
  return `${softBreak}<${prefix}r><${prefix}rPr ${attributes}>${outline}${solidFill}${glow}${highlight}${underlineFill}<${prefix}latin typeface="${latin}"/><${prefix}ea typeface="${eastAsian}"/><${prefix}cs typeface="${complexScript}"/>${hyperlinkXml}</${prefix}rPr><${prefix}t xml:space="preserve">${escapeXmlText(run.text)}</${prefix}t></${prefix}r>`;
}

function renderMainTextColorChoice(
  color: RichTextColor,
  prefix: string,
  transparency: number | undefined,
): string {
  const tag = color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  if (transparency === undefined) return `<${prefix}${tag} val="${color.value}"/>`;
  const alpha = Math.round((100 - transparency) * 1_000);
  return `<${prefix}${tag} val="${color.value}"><${prefix}alpha val="${alpha}"/></${prefix}${tag}>`;
}

export function renderColorChoice(color: RichTextColor, prefix: string): string {
  return color.kind === 'srgb'
    ? `<${prefix}srgbClr val="${color.value}"/>`
    : `<${prefix}schemeClr val="${color.value}"/>`;
}

function renderGlowColorChoice(glow: RichTextGlow, prefix: string): string {
  const color = glow.color ?? { kind: 'srgb' as const, value: 'FFFFFF' };
  const tag = color.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  return `<${prefix}${tag} val="${color.value}"><${prefix}alpha val="${Math.round(glow.opacity * PERCENT_SCALE)}"/></${prefix}${tag}>`;
}

interface ReadRunsState {
  readonly runs: readonly RichTextRun[];
  readonly hyperlinkBindings: readonly (ReadRichTextRunHyperlinkBinding | undefined)[];
}

function readRuns(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
  context: ShapeHyperlinkReadContext,
): ReadRunsState {
  const runs: RichTextRun[] = [];
  const hyperlinkBindings: (ReadRichTextRunHyperlinkBinding | undefined)[] = [];
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
      hyperlinkBindings.push(undefined);
      pendingBreaks -= 1;
    }
    const properties = directChildren(child, 'rPr');
    const hyperlinkBinding = properties.length === 1
      ? readTextRunHyperlinkBinding(properties[0]!, context)
      : undefined;
    const style = readStyle(xml, child, hyperlinkBinding?.hyperlink);
    runs.push({
      text: xml.descendants(child, 't').map((text) => xml.text(text)).join(''),
      ...(style ? { style } : {}),
      ...(pendingBreaks === 1 ? { softBreakBefore: true } : {}),
    });
    hyperlinkBindings.push(hyperlinkBinding);
    pendingBreaks = 0;
  }
  while (pendingBreaks > 0) {
    runs.push({ text: '', softBreakBefore: true });
    hyperlinkBindings.push(undefined);
    pendingBreaks -= 1;
  }
  return { runs, hyperlinkBindings };
}

function readParagraphAlignment(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): TextAlignment | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  const value = properties ? xml.attribute(properties, 'algn')?.value : undefined;
  return value ? OOXML_TO_TEXT_ALIGNMENT.get(value) : undefined;
}

function readParagraphRtl(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): boolean | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  const value = properties ? xml.attribute(properties, 'rtl')?.value : undefined;
  if (value === undefined) return undefined;
  if (['1', 'true', 'on'].includes(value)) return true;
  if (['0', 'false', 'off'].includes(value)) return false;
  return undefined;
}

function readParagraphMarginLeft(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): number | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  if (directChildren(properties).some(({ localName }) =>
    ['buChar', 'buAutoNum', 'buBlip'].includes(localName))) return undefined;
  const margin = readIntegerAttribute(xml, properties, 'marL');
  if (margin === undefined || margin < 0 || margin > MAX_PARAGRAPH_MARGIN_EMU) return undefined;
  return margin / EMU_PER_POINT;
}

function readParagraphMarginRight(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): number | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  const margin = readIntegerAttribute(xml, properties, 'marR');
  if (margin === undefined || margin < 0 || margin > MAX_PARAGRAPH_MARGIN_EMU) return undefined;
  return margin / EMU_PER_POINT;
}

function readParagraphIndent(
  xml: LosslessXmlDocument,
  paragraph: XmlElement,
): number | undefined {
  const properties = directChildren(paragraph, 'pPr')[0];
  if (!properties) return undefined;
  if (directChildren(properties).some(({ localName }) =>
    ['buChar', 'buAutoNum', 'buBlip'].includes(localName))) return undefined;
  const indent = readIntegerAttribute(xml, properties, 'indent');
  if (indent === undefined || indent < -MAX_PARAGRAPH_MARGIN_EMU || indent > MAX_PARAGRAPH_MARGIN_EMU) {
    return undefined;
  }
  return indent / EMU_PER_POINT;
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

function readStyle(
  xml: LosslessXmlDocument,
  run: XmlElement,
  hyperlink: NormalizedHyperlink | undefined,
): RichTextRunStyle | undefined {
  const properties = directChildren(run, 'rPr')[0];
  if (!properties) return undefined;
  const size = Number.parseInt(xml.attribute(properties, 'sz')?.value ?? '', 10);
  const language = xml.attribute(properties, 'lang')?.value;
  const baseline = readBaseline(xml, properties);
  const characterSpacing = readCharacterSpacing(xml, properties);
  const transparency = readMainTextTransparency(xml, properties);
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
  const strike = readStrike(xml, properties);
  const glow = readGlow(xml, properties);
  const highlight = readHighlight(xml, properties);
  const outline = readOutline(xml, properties);
  const underline = readUnderline(xml, properties);
  const style: RichTextRunStyle = {
    ...(fontFamily !== undefined ? { fontFamily } : {}),
    ...(Number.isFinite(size) && size > 0 ? { fontSize: size / 100 } : {}),
    ...(language ? { lang: language } : {}),
    ...(baseline !== undefined ? { baseline } : {}),
    ...(characterSpacing !== undefined ? { characterSpacing } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(color ? { color } : {}),
    ...(transparency !== undefined ? { transparency } : {}),
    ...(glow ? { glow } : {}),
    ...(highlight ? { highlight } : {}),
    ...(hyperlink ? { hyperlink } : {}),
    ...(outline ? { outline } : {}),
    ...(underline !== undefined ? { underline } : {}),
    ...(strike !== undefined ? { strike } : {}),
  };
  return Object.keys(style).length > 0 ? style : undefined;
}

function readMainTextTransparency(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): number | undefined {
  const fills = directChildren(properties, 'solidFill');
  if (fills.length !== 1) return undefined;
  const colors = directChildren(fills[0]!);
  if (colors.length !== 1) return undefined;
  const color = colors[0]!;
  const colorValue = xml.attribute(color, 'val')?.value;
  if (
    !(
      color.localName === 'srgbClr'
      && colorValue !== undefined
      && /^[\da-f]{6}$/i.test(colorValue)
    )
    && !(
      color.localName === 'schemeClr'
      && colorValue !== undefined
      && SCHEME_COLORS.has(colorValue)
    )
  ) return undefined;
  const colorAttributes = color.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
  if (colorAttributes.length !== 1 || colorAttributes[0]?.name !== 'val') return undefined;
  const transforms = directChildren(color);
  if (transforms.length !== 1 || transforms[0]?.localName !== 'alpha') return undefined;
  const alpha = transforms[0]!;
  const alphaAttributes = alpha.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
  if (
    alphaAttributes.length !== 1
    || alphaAttributes[0]?.name !== 'val'
    || directChildren(alpha).length > 0
  ) return undefined;
  const value = readIntegerAttribute(xml, alpha, 'val');
  return value !== undefined && value >= 0 && value <= PERCENT_SCALE
    ? 100 - value / 1_000
    : undefined;
}

function readCharacterSpacing(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): number | undefined {
  if (!xml.attribute(properties, 'spc')) return undefined;
  const raw = readIntegerAttribute(xml, properties, 'spc');
  return raw !== undefined && raw >= MIN_COORDINATE_32 && raw <= MAX_COORDINATE_32
    ? raw / 100
    : undefined;
}

function readBaseline(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): RichTextBaseline | undefined {
  if (!xml.attribute(properties, 'baseline')) return undefined;
  const raw = readIntegerAttribute(xml, properties, 'baseline');
  if (raw === undefined || raw < MIN_COORDINATE_32 || raw > MAX_COORDINATE_32) return undefined;
  if (raw === 30_000) return 'superscript';
  if (raw === -40_000) return 'subscript';
  return raw / 1_000;
}

function readGlow(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): RichTextGlow | undefined {
  if (directChildren(properties, 'effectDag').length > 0) return undefined;
  const effectLists = directChildren(properties, 'effectLst');
  if (effectLists.length !== 1) return undefined;
  const elements = directChildren(effectLists[0]!, 'glow');
  if (elements.length !== 1) return undefined;
  const glow = elements[0]!;
  const attributes = glow.attributes.filter(({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'));
  if (attributes.length !== 1 || attributes[0]?.localName !== 'rad') return undefined;
  const radiusEmu = readIntegerAttribute(xml, glow, 'rad');
  if (radiusEmu === undefined || radiusEmu < 0 || radiusEmu > MAX_POSITIVE_COORDINATE_EMU) return undefined;
  const result = readGlowColor(xml, glow);
  return result
    ? { color: result.color, opacity: result.opacity, size: radiusEmu / EMU_PER_POINT }
    : undefined;
}

function readGlowColor(
  xml: LosslessXmlDocument,
  glow: XmlElement,
): { color: RichTextColor; opacity: number } | undefined {
  const children = directChildren(glow);
  if (children.length !== 1) return undefined;
  const colorElement = children[0]!;
  const colorAttributes = colorElement.attributes.filter(
    ({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'),
  );
  if (colorAttributes.length !== 1 || colorAttributes[0]?.localName !== 'val') return undefined;
  const value = xml.attribute(colorElement, 'val')?.value;
  const color = colorElement.localName === 'srgbClr' && value && /^[\da-f]{6}$/i.test(value)
    ? { kind: 'srgb' as const, value: value.toUpperCase() }
    : colorElement.localName === 'schemeClr' && value && SCHEME_COLORS.has(value)
      ? { kind: 'scheme' as const, value }
      : undefined;
  if (!color) return undefined;
  const transforms = directChildren(colorElement);
  if (transforms.length === 0) return { color, opacity: 1 };
  if (transforms.length !== 1 || transforms[0]?.localName !== 'alpha') return undefined;
  const alpha = transforms[0]!;
  const alphaAttributes = alpha.attributes.filter(({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'));
  if (alphaAttributes.length !== 1 || alphaAttributes[0]?.localName !== 'val' || directChildren(alpha).length > 0) {
    return undefined;
  }
  const opacity = readIntegerAttribute(xml, alpha, 'val');
  return opacity !== undefined && opacity >= 0 && opacity <= PERCENT_SCALE
    ? { color, opacity: opacity / PERCENT_SCALE }
    : undefined;
}

function readOutline(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): RichTextOutline | undefined {
  const elements = directChildren(properties, 'ln');
  if (elements.length !== 1) return undefined;
  const line = elements[0]!;
  const attributes = line.attributes.filter(({ name }) => name !== 'xmlns' && !name.startsWith('xmlns:'));
  if (attributes.length !== 1 || attributes[0]?.localName !== 'w') return undefined;
  const widthEmu = readIntegerAttribute(xml, line, 'w');
  if (widthEmu === undefined || widthEmu < 0 || widthEmu > MAX_LINE_WIDTH_EMU) return undefined;
  const color = readDirectSolidColor(xml, line);
  return color ? { color, size: widthEmu / EMU_PER_POINT } : undefined;
}

function readHighlight(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): RichTextColor | undefined {
  const elements = directChildren(properties, 'highlight');
  return elements.length === 1 ? readDirectColorChoice(xml, elements[0]!) : undefined;
}

function readStrike(
  xml: LosslessXmlDocument,
  properties: XmlElement,
): RichTextStrikeStyle | false | undefined {
  const value = xml.attribute(properties, 'strike')?.value;
  if (value === undefined) return undefined;
  if (value === 'noStrike') return false;
  return STRIKE_STYLES.has(value as RichTextStrikeStyle)
    ? value as RichTextStrikeStyle
    : undefined;
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
  return readDirectColorChoice(xml, containerChildren[0]);
}

function readDirectColorChoice(
  xml: LosslessXmlDocument,
  container: XmlElement,
): RichTextColor | undefined {
  const children = directChildren(container);
  if (children.length !== 1) return undefined;
  const color = children[0]!;
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
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !supported.includes(key)) {
      throw new TypeError(`${context} property ${String(key)} is not supported yet`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
  }
}
