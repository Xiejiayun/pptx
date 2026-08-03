import type {
  AddTableCellOptions,
  RichTextColor,
  TableCellBorder,
  TableCellBorders,
  TextAlignment,
  TextBoxVerticalAlignment,
} from '@pptx/model';

const CSS_PROPERTIES = Object.freeze([
  'color',
  'background-color',
  'font-family',
  'font-size',
  'font-weight',
  'text-align',
  'vertical-align',
  'direction',
  'padding-top',
  'padding-right',
  'padding-bottom',
  'padding-left',
  'border-top-style',
  'border-top-width',
  'border-top-color',
  'border-right-style',
  'border-right-width',
  'border-right-color',
  'border-bottom-style',
  'border-bottom-width',
  'border-bottom-color',
  'border-left-style',
  'border-left-width',
  'border-left-color',
] as const);
const BORDER_SIDES = ['top', 'right', 'bottom', 'left'] as const;
const GENERIC_FONT_FAMILIES = new Set([
  'cursive',
  'emoji',
  'fangsong',
  'fantasy',
  'inherit',
  'initial',
  'math',
  'monospace',
  'revert',
  'revert-layer',
  'sans-serif',
  'serif',
  'system-ui',
  'ui-monospace',
  'ui-rounded',
  'ui-sans-serif',
  'ui-serif',
  'unset',
]);
const OMITTED_CSS_VALUES = new Set([
  '',
  'inherit',
  'initial',
  'revert',
  'revert-layer',
  'unset',
]);
const DASH_BORDER_STYLES = new Set([
  'dashed',
  'dotted',
  'double',
  'groove',
  'inset',
  'outset',
  'ridge',
]);
const CSS_NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/;

type CssPropertyName = typeof CSS_PROPERTIES[number];
type BorderSide = typeof BORDER_SIDES[number];

export function mapComputedCellOptions(
  style: object,
  context: string,
): Readonly<AddTableCellOptions> {
  const css = readComputedCss(style, context);
  const color = css.color === ''
    ? undefined
    : parseCssRgb(css.color, `${context} color`).color;
  const fill = css['background-color'] === ''
    ? undefined
    : backgroundFill(css['background-color'], `${context} background-color`);
  const fontFamily = normalizeFontFamily(css['font-family'], `${context} font-family`);
  const fontSize = css['font-size'] === ''
    ? undefined
    : parseCssLengthPx(css['font-size'], `${context} font-size`, true);
  const bold = normalizeFontWeight(css['font-weight'], `${context} font-weight`);
  const align = normalizeCssAlignment(
    css['text-align'],
    css.direction,
    `${context} text-align`,
  );
  const valign = normalizeCssVerticalAlignment(
    css['vertical-align'],
    `${context} vertical-align`,
  );
  const margin = normalizeCssPadding(css, context);
  const border = normalizeCssBorders(css, context);
  return Object.freeze({
    ...(align === undefined ? {} : { align }),
    ...(bold === undefined ? {} : { bold }),
    ...(color === undefined ? {} : { color }),
    ...(fill === undefined ? {} : { fill }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(margin === undefined ? {} : { margin }),
    ...(valign === undefined ? {} : { valign }),
    ...(border === undefined ? {} : { border }),
  });
}

function readComputedCss(
  style: object,
  context: string,
): Readonly<Record<CssPropertyName, string>> {
  const propertyGetter = readProperty(
    style,
    'getPropertyValue',
    `${context} computed style getPropertyValue`,
  );
  if (propertyGetter !== undefined && typeof propertyGetter !== 'function') {
    throw new TypeError(`${context} computed style getPropertyValue must be a function`);
  }
  const result = Object.create(null) as Record<CssPropertyName, string>;
  for (const property of CSS_PROPERTIES) {
    const raw = propertyGetter === undefined
      ? readProperty(style, property, `${context} computed ${property}`)
      : callPlatformMethod(
          propertyGetter as (...args: unknown[]) => unknown,
          style,
          [property],
          `${context} computed ${property}`,
        );
    if (raw !== undefined && typeof raw !== 'string') {
      throw new TypeError(`${context} computed ${property} must be a string`);
    }
    result[property] = raw?.trim() ?? '';
  }
  return Object.freeze(result);
}

function parseCssRgb(
  value: string,
  context: string,
): Readonly<{ color: RichTextColor; alpha: number }> {
  const match = /^(rgb|rgba)\((.*)\)$/i.exec(value.trim());
  if (!match) throw new TypeError(`${context} must use rgb() or rgba()`);
  const body = match[2]!.trim();
  let channelTokens: string[];
  let alphaToken: string | undefined;
  if (body.includes(',')) {
    const tokens = body.split(',').map((token) => token.trim());
    if (tokens.length !== 3 && tokens.length !== 4) {
      throw new TypeError(`${context} must contain three RGB channels and optional alpha`);
    }
    channelTokens = tokens.slice(0, 3);
    alphaToken = tokens[3];
  } else {
    const slashParts = body.split('/').map((token) => token.trim());
    if (slashParts.length > 2) {
      throw new TypeError(`${context} contains an invalid alpha separator`);
    }
    channelTokens = slashParts[0]!.split(/\s+/).filter(Boolean);
    alphaToken = slashParts[1];
  }
  if (channelTokens.length !== 3) {
    throw new TypeError(`${context} must contain exactly three RGB channels`);
  }
  const channels = channelTokens.map((token, index) => {
    if (!CSS_NUMBER.test(token)) {
      throw new TypeError(`${context} channel ${index} must be a number from 0 to 255`);
    }
    const channel = Number(token);
    if (!Number.isFinite(channel) || channel < 0 || channel > 255) {
      throw new RangeError(`${context} channel ${index} must be between 0 and 255`);
    }
    return Math.round(channel);
  });
  const alpha = alphaToken === undefined ? 1 : parseCssAlpha(alphaToken, context);
  const hex = channels
    .map((channel) => channel.toString(16).padStart(2, '0').toUpperCase())
    .join('');
  return Object.freeze({
    color: Object.freeze({ kind: 'srgb', value: hex }),
    alpha,
  });
}

function parseCssAlpha(value: string, context: string): number {
  const percentage = value.endsWith('%');
  const numericToken = percentage ? value.slice(0, -1) : value;
  if (!CSS_NUMBER.test(numericToken)) {
    throw new TypeError(`${context} alpha must be a number or percentage`);
  }
  const number = Number(numericToken);
  const alpha = percentage ? number / 100 : number;
  if (!Number.isFinite(alpha) || alpha < 0 || alpha > 1) {
    throw new RangeError(`${context} alpha must be between 0 and 1`);
  }
  return alpha;
}

function backgroundFill(
  value: string,
  context: string,
): NonNullable<AddTableCellOptions['fill']> {
  if (value.toLowerCase() === 'transparent') {
    return Object.freeze({
      kind: 'solid',
      color: Object.freeze({ kind: 'srgb', value: 'FFFFFF' }),
    });
  }
  const parsed = parseCssRgb(value, context);
  return Object.freeze({
    kind: 'solid',
    color: parsed.alpha === 0
      ? Object.freeze({ kind: 'srgb', value: 'FFFFFF' })
      : parsed.color,
  });
}

function normalizeFontFamily(value: string, context: string): string | undefined {
  if (OMITTED_CSS_VALUES.has(value.toLowerCase())) return undefined;
  let quote: string | undefined;
  let end = value.length;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]!;
    if ((character === '"' || character === "'") && value[index - 1] !== '\\') {
      if (quote === undefined) quote = character;
      else if (quote === character) quote = undefined;
    } else if (character === ',' && quote === undefined) {
      end = index;
      break;
    }
  }
  if (quote !== undefined) throw new TypeError(`${context} contains an unterminated quote`);
  let family = value.slice(0, end).trim();
  if (
    family.length >= 2
    && ((family.startsWith('"') && family.endsWith('"'))
      || (family.startsWith("'") && family.endsWith("'")))
  ) family = family.slice(1, -1).trim();
  if (family.length === 0) throw new TypeError(`${context} must contain a font family`);
  return GENERIC_FONT_FAMILIES.has(family.toLowerCase()) ? undefined : family;
}

function normalizeFontWeight(value: string, context: string): boolean | undefined {
  const normalized = value.toLowerCase();
  if (OMITTED_CSS_VALUES.has(normalized)) return undefined;
  if (normalized === 'bold' || normalized === 'bolder') return true;
  if (normalized === 'normal' || normalized === 'lighter') return false;
  const numeric = Number(normalized);
  if (!Number.isFinite(numeric) || numeric < 1 || numeric > 1_000) {
    throw new TypeError(`${context} must be normal, bold, bolder, lighter, or 1–1000`);
  }
  return numeric >= 500;
}

function normalizeCssAlignment(
  value: string,
  direction: string,
  context: string,
): TextAlignment | undefined {
  const normalized = value.toLowerCase();
  if (OMITTED_CSS_VALUES.has(normalized)) return undefined;
  if (
    normalized === 'left'
    || normalized === 'center'
    || normalized === 'right'
    || normalized === 'justify'
  ) return normalized;
  if (normalized !== 'start' && normalized !== 'end') {
    throw new TypeError(`${context} must be left, center, right, justify, start, or end`);
  }
  const normalizedDirection = direction.toLowerCase();
  if (normalizedDirection !== 'ltr' && normalizedDirection !== 'rtl') {
    throw new TypeError(`${context} direction must be ltr or rtl for start/end alignment`);
  }
  return normalized === 'start'
    ? normalizedDirection === 'rtl' ? 'right' : 'left'
    : normalizedDirection === 'rtl' ? 'left' : 'right';
}

function normalizeCssVerticalAlignment(
  value: string,
  context: string,
): TextBoxVerticalAlignment | undefined {
  const normalized = value.toLowerCase();
  if (
    OMITTED_CSS_VALUES.has(normalized)
    || normalized === 'baseline'
    || normalized === 'sub'
    || normalized === 'super'
    || normalized === 'text-bottom'
    || normalized === 'text-top'
  ) return undefined;
  if (normalized === 'top' || normalized === 'middle' || normalized === 'bottom') {
    return normalized;
  }
  throw new TypeError(`${context} must be top, middle, bottom, or a generic baseline value`);
}

function normalizeCssPadding(
  css: Readonly<Record<CssPropertyName, string>>,
  context: string,
): readonly [number, number, number, number] | undefined {
  const values = BORDER_SIDES.map((side) => css[`padding-${side}`]);
  if (values.every((value) => value === '')) return undefined;
  return Object.freeze(values.map((value, index) => value === ''
    ? 0
    : parseCssLengthPx(value, `${context} padding-${BORDER_SIDES[index]}`, false))) as
      readonly [number, number, number, number];
}

function normalizeCssBorders(
  css: Readonly<Record<CssPropertyName, string>>,
  context: string,
): Readonly<TableCellBorders> | undefined {
  const borders: Partial<Record<BorderSide, TableCellBorder>> = {};
  for (const side of BORDER_SIDES) {
    const border = normalizeCssBorder(css, side, context);
    if (border !== undefined) borders[side] = border;
  }
  return Object.keys(borders).length === 0
    ? undefined
    : Object.freeze(borders);
}

function normalizeCssBorder(
  css: Readonly<Record<CssPropertyName, string>>,
  side: BorderSide,
  context: string,
): Readonly<TableCellBorder> | undefined {
  const style = css[`border-${side}-style`].toLowerCase();
  const widthValue = css[`border-${side}-width`];
  const colorValue = css[`border-${side}-color`];
  if (style === '') {
    if (widthValue !== '' || colorValue !== '') {
      throw new TypeError(`${context} border-${side}-style cannot be empty`);
    }
    return undefined;
  }
  if (style === 'none' || style === 'hidden') return Object.freeze({ kind: 'none' });
  if (style !== 'solid' && !DASH_BORDER_STYLES.has(style)) {
    throw new TypeError(`${context} border-${side}-style is unsupported`);
  }
  const width = parseCssLengthPx(widthValue, `${context} border-${side}-width`, false);
  if (width === 0) return Object.freeze({ kind: 'none' });
  const color = parseCssRgb(colorValue, `${context} border-${side}-color`).color;
  return Object.freeze({
    kind: 'line',
    color,
    width,
    style: style === 'solid' ? 'solid' : 'dash',
  });
}

function parseCssLengthPx(
  value: string,
  context: string,
  positive: boolean,
): number {
  const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))px$/i.exec(value.trim());
  if (!match) throw new TypeError(`${context} must be a finite px length`);
  const length = Number(match[1]);
  if (!Number.isFinite(length)) throw new TypeError(`${context} must be finite`);
  if (positive ? length <= 0 : length < 0) {
    throw new RangeError(`${context} must be ${positive ? 'positive' : 'non-negative'}`);
  }
  return Object.is(length, -0) ? 0 : length;
}

function callPlatformMethod(
  method: (...args: unknown[]) => unknown,
  receiver: object,
  args: readonly unknown[],
  context: string,
): unknown {
  try {
    return Reflect.apply(method, receiver, args);
  } catch (error) {
    throw new TypeError(`${context} failed`, { cause: error });
  }
}

function readProperty(value: object, key: string, context: string): unknown {
  try {
    return Reflect.get(value, key);
  } catch (error) {
    throw new TypeError(`${context} could not be read`, { cause: error });
  }
}
