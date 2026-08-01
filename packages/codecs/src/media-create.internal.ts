import type {
  AddMediaOptions,
  MediaByteStream,
  MediaKind,
} from './media.js';

const OPTION_KEYS = new Set([
  'name',
  'altText',
  'contentType',
  'fileName',
  'poster',
  'posterContentType',
  'x',
  'y',
  'width',
  'height',
  'play',
  'loop',
  'hideWhenStopped',
  'volume',
  'transcode',
]);

export type NormalizedMediaSourceReference =
  | Readonly<{ type: 'string'; value: string }>
  | Readonly<{ type: 'bytes'; bytes: Uint8Array }>
  | Readonly<{ type: 'blob'; value: Blob }>
  | Readonly<{ type: 'stream'; value: MediaByteStream }>;

export interface NormalizedMediaCreateRequest {
  readonly kind: MediaKind;
  readonly source: NormalizedMediaSourceReference;
  readonly poster?: NormalizedMediaSourceReference;
  readonly name?: string;
  readonly altText?: string;
  readonly contentType?: string;
  readonly posterContentType?: string;
  readonly fileName?: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly play: 'click' | 'auto';
  readonly loop: boolean;
  readonly hideWhenStopped: boolean;
  readonly volume: number;
  readonly transcode?: NonNullable<AddMediaOptions['transcode']>;
}

export function normalizeMediaCreateRequest(
  kind: unknown,
  source: unknown,
  options: unknown,
): Readonly<NormalizedMediaCreateRequest> {
  const normalizedKind = normalizeKind(kind);
  const values = readOptions(options);
  const poster = values.poster === undefined
    ? undefined
    : normalizeSource(values.poster, 'poster');
  const name = normalizeXmlString(values.name, 'name');
  const altText = normalizeXmlString(values.altText, 'altText');
  const contentType = normalizeNonEmptyString(values.contentType, 'contentType');
  const posterContentType = normalizeNonEmptyString(
    values.posterContentType,
    'posterContentType',
  );
  const fileName = normalizeNonEmptyString(values.fileName, 'fileName');
  const transcode = normalizeTranscode(values.transcode);

  return Object.freeze({
    kind: normalizedKind,
    source: normalizeSource(source, 'source'),
    ...(poster === undefined ? {} : { poster }),
    ...(name === undefined ? {} : { name }),
    ...(altText === undefined ? {} : { altText }),
    ...(contentType === undefined ? {} : { contentType }),
    ...(posterContentType === undefined ? {} : { posterContentType }),
    ...(fileName === undefined ? {} : { fileName }),
    x: normalizeInteger(values.x, 914_400, 'x'),
    y: normalizeInteger(values.y, 914_400, 'y'),
    width: normalizePositiveInteger(
      values.width,
      normalizedKind === 'video' ? 4_572_000 : 914_400,
      'width',
    ),
    height: normalizePositiveInteger(
      values.height,
      normalizedKind === 'video' ? 2_571_750 : 914_400,
      'height',
    ),
    play: normalizePlay(values.play),
    loop: normalizeBoolean(values.loop, false, 'loop'),
    hideWhenStopped: normalizeBoolean(
      values.hideWhenStopped,
      false,
      'hideWhenStopped',
    ),
    volume: normalizeVolume(values.volume),
    ...(transcode === undefined ? {} : { transcode }),
  });
}

function readOptions(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('Media creation options must be an object');
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError('Media creation options must be an ordinary object');
  }
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !OPTION_KEYS.has(key)) {
      throw new TypeError(`Media creation options contain unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`Media creation option ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}

function normalizeKind(value: unknown): MediaKind {
  if (value === 'audio' || value === 'video') return value;
  throw new TypeError('Media kind must be audio or video');
}

function normalizeSource(value: unknown, label: string): NormalizedMediaSourceReference {
  if (typeof value === 'string') {
    if (value.length === 0) throw new TypeError(`Media ${label} string must not be empty`);
    return Object.freeze({ type: 'string', value });
  }
  if (value instanceof Uint8Array) {
    if (value.length === 0) throw new RangeError(`Media ${label} bytes must not be empty`);
    return Object.freeze({ type: 'bytes', bytes: new Uint8Array(value) });
  }
  if (value instanceof ArrayBuffer) {
    if (value.byteLength === 0) throw new RangeError(`Media ${label} bytes must not be empty`);
    return Object.freeze({
      type: 'bytes',
      bytes: new Uint8Array(new Uint8Array(value)),
    });
  }
  if (isBlob(value)) return Object.freeze({ type: 'blob', value });
  if (isMediaByteStream(value)) {
    return Object.freeze({ type: 'stream', value });
  }
  throw new TypeError(`Media ${label} must be a path, URL, byte buffer, Blob, or byte stream`);
}

function normalizeXmlString(value: unknown, name: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') throw new TypeError(`Media ${name} must be a string`);
  if (!isValidXmlString(value)) {
    throw new TypeError(`Media ${name} contains invalid XML characters`);
  }
  return value;
}

function normalizeNonEmptyString(value: unknown, name: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`Media ${name} must be a non-empty string`);
  }
  return value;
}

function normalizeInteger(value: unknown, defaultValue: number, name: string): number {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`Media ${name} must be finite`);
  }
  if (!Number.isSafeInteger(value)) {
    throw new RangeError(`Media ${name} must be a safe integer`);
  }
  return value === 0 ? 0 : value;
}

function normalizePositiveInteger(value: unknown, defaultValue: number, name: string): number {
  const normalized = normalizeInteger(value, defaultValue, name);
  if (normalized <= 0) throw new RangeError(`Media ${name} must be positive`);
  return normalized;
}

function normalizePlay(value: unknown): 'click' | 'auto' {
  if (value === undefined) return 'click';
  if (value === 'click' || value === 'auto') return value;
  throw new TypeError('Media play must be click or auto');
}

function normalizeBoolean(value: unknown, defaultValue: boolean, name: string): boolean {
  if (value === undefined) return defaultValue;
  if (typeof value !== 'boolean') throw new TypeError(`Media ${name} must be a boolean`);
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

function normalizeTranscode(
  value: unknown,
): NonNullable<AddMediaOptions['transcode']> | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'function') throw new TypeError('Media transcode must be a function');
  return value as NonNullable<AddMediaOptions['transcode']>;
}

function isBlob(value: unknown): value is Blob {
  return typeof Blob !== 'undefined' && value instanceof Blob;
}

function isMediaByteStream(value: unknown): value is MediaByteStream {
  return Boolean(
    value
    && (
      typeof (value as { getReader?: unknown }).getReader === 'function'
      || typeof (value as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === 'function'
    ),
  );
}

function isValidXmlString(value: string): boolean {
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
    return false;
  }
  return true;
}
