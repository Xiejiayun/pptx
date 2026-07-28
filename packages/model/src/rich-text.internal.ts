import {
  escapeXmlAttribute,
  escapeXmlText,
  LosslessXmlDocument,
  type XmlElement,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import type {
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

export function normalizeRichText(value: unknown): readonly RichTextParagraph[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new TypeError('Rich text must contain at least one paragraph');
  }
  return value.map((paragraph, paragraphIndex) => {
    if (!paragraph || typeof paragraph !== 'object' || Array.isArray(paragraph)) {
      throw new TypeError(`Rich text paragraph ${paragraphIndex} must be an object`);
    }
    assertSupportedKeys(paragraph, ['align', 'runs'], `Rich text paragraph ${paragraphIndex}`);
    const candidate = paragraph as { align?: unknown; runs?: unknown };
    const runs = candidate.runs;
    if (!Array.isArray(runs)) throw new TypeError(`Rich text paragraph ${paragraphIndex} runs must be an array`);
    const align = candidate.align === undefined
      ? undefined
      : normalizeTextAlignment(candidate.align, `Rich text paragraph ${paragraphIndex} align`);
    return {
      runs: runs.map((run, runIndex) => normalizeRun(run, paragraphIndex, runIndex)),
      ...(align ? { align } : {}),
    };
  });
}

interface RenderRichTextOptions {
  readonly prefix?: string;
  readonly defaultAlign?: TextAlignment;
  readonly paragraphProperties?: readonly (string | undefined)[];
  readonly endParagraphProperties?: string;
}

export function renderRichTextParagraphs(
  paragraphs: readonly RichTextParagraph[],
  options: RenderRichTextOptions = {},
): string {
  const prefix = options.prefix ?? 'a:';
  const defaultEndProperties = `<${prefix}endParaRPr lang="en-US" dirty="0"/>`;
  return paragraphs
    .map(
      ({ align, runs }, index) =>
        `<${prefix}p>${renderParagraphProperties(
          options.paragraphProperties?.[index] ?? options.paragraphProperties?.[0],
          prefix,
          align ?? options.defaultAlign,
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
    return {
      runs: readRuns(xml, paragraph),
      ...(align ? { align } : {}),
    };
  });
}

export function replaceRichText(
  xml: LosslessXmlDocument,
  element: XmlElement,
  paragraphs: readonly RichTextParagraph[],
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

export function renderParagraphProperties(
  template: string | undefined,
  prefix: string,
  alignment: TextAlignment | undefined,
): string {
  if (!template) {
    const align = alignment ? ` algn="${TEXT_ALIGNMENT_TO_OOXML[alignment]}"` : '';
    return `<${prefix}pPr${align} indent="0" marL="0"><${prefix}buNone/></${prefix}pPr>`;
  }
  const properties = LosslessXmlDocument.parse(template);
  const root = properties.roots[0];
  if (!root || root.localName !== 'pPr') throw new ModelParseError('Invalid paragraph properties template');
  const attribute = properties.attribute(root, 'algn');
  if (alignment) {
    const value = TEXT_ALIGNMENT_TO_OOXML[alignment];
    if (attribute) properties.replaceAttribute(attribute, value);
    else {
      const insertionPoint = root.selfClosing
        ? properties.source.lastIndexOf('/', root.startTagEnd - 1)
        : root.startTagEnd - 1;
      properties.replace(insertionPoint, insertionPoint, ` algn="${value}"`);
    }
  } else if (attribute) {
    let start = attribute.start;
    while (start > root.start && /[\t ]/.test(properties.source[start - 1] ?? '')) start -= 1;
    properties.replace(start, attribute.end, '');
  }
  return properties.serialize();
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

function directChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
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
