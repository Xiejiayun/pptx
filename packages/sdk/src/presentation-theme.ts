export interface PresentationTheme {
  readonly headFontFace: string;
  readonly bodyFontFace: string;
}

export interface PresentationThemeOptions {
  readonly headFontFace?: string;
  readonly bodyFontFace?: string;
}

const DEFAULT_HEAD_FONT = 'Calibri Light';
const DEFAULT_BODY_FONT = 'Calibri';
const INVALID_XML_CHARACTERS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/;
const THEME_KEYS = new Set(['headFontFace', 'bodyFontFace']);

export function normalizePresentationTheme(value: unknown): PresentationTheme {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Presentation theme must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Presentation theme must be an ordinary object');
  }
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !THEME_KEYS.has(key)) {
      throw new TypeError(`Unsupported presentation theme property: ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Presentation theme ${key} must be a data property`);
    }
  }
  return {
    headFontFace: readFont(value, 'headFontFace', DEFAULT_HEAD_FONT),
    bodyFontFace: readFont(value, 'bodyFontFace', DEFAULT_BODY_FONT),
  };
}

function readFont(
  value: object,
  key: 'headFontFace' | 'bodyFontFace',
  fallback: string,
): string {
  const descriptor = Object.getOwnPropertyDescriptor(value, key);
  return descriptor?.value === undefined
    ? fallback
    : normalizeFontName(descriptor.value, key);
}

function normalizeFontName(value: unknown, key: string): string {
  if (typeof value !== 'string' || !/\S/u.test(value)) {
    throw new TypeError(`Presentation theme ${key} must be a non-whitespace string`);
  }
  if (INVALID_XML_CHARACTERS.test(value)) {
    throw new TypeError(`Presentation theme ${key} contains invalid XML characters`);
  }
  return value;
}
