import {
  inspectRasterImage,
  resolveImageSource,
  type ImageSource,
  type ResolvedImageSource,
} from './raster-image-source.js';

const MAX_CANVAS_SIDE = 8_192;
const MAX_CANVAS_PIXELS = 16_777_216;
const TRANSPARENT_PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0,
  0, 11, 73, 68, 65, 84, 120, 218, 99, 100, 248, 15, 0, 1, 5, 1, 1,
  39, 24, 227, 102, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

export async function resolveSvgFallback(
  primary: Readonly<ResolvedImageSource>,
  explicit: ImageSource | undefined,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  if (primary.info.contentType !== 'image/svg+xml') {
    throw new TypeError('SVG fallback primary source must be an SVG image');
  }
  throwIfAborted(signal);
  if (explicit !== undefined) {
    try {
      const resolved = await resolveImageSource(explicit, signal);
      if (resolved.info.contentType !== 'image/png') {
        throw new TypeError('SVG fallback must be a valid PNG image');
      }
      return new Uint8Array(resolved.bytes);
    } catch {
      throwIfAborted(signal);
      throw new TypeError('SVG fallback must be a valid PNG image');
    }
  }

  try {
    const generated = await rasterizeSvgInBrowser(primary, signal);
    if (generated) return generated;
  } catch {
    throwIfAborted(signal);
  }
  return builtInFallback();
}

function builtInFallback(): Uint8Array {
  return new Uint8Array(TRANSPARENT_PNG);
}

async function rasterizeSvgInBrowser(
  primary: Readonly<ResolvedImageSource>,
  signal?: AbortSignal,
): Promise<Uint8Array | undefined> {
  const width = Math.ceil(primary.info.width);
  const height = Math.ceil(primary.info.height);
  if (
    !Number.isSafeInteger(width)
    || !Number.isSafeInteger(height)
    || width <= 0
    || height <= 0
    || width > MAX_CANVAS_SIDE
    || height > MAX_CANVAS_SIDE
    || width * height > MAX_CANVAS_PIXELS
  ) return undefined;
  if (
    typeof Image !== 'function'
    || typeof document === 'undefined'
    || typeof document.createElement !== 'function'
    || typeof URL === 'undefined'
    || typeof URL.createObjectURL !== 'function'
    || typeof URL.revokeObjectURL !== 'function'
  ) return undefined;

  throwIfAborted(signal);
  const objectUrl = URL.createObjectURL(new Blob(
    [new Uint8Array(primary.bytes)],
    { type: 'image/svg+xml' },
  ));
  try {
    const image = new Image();
    image.src = objectUrl;
    await decodeImage(image, signal);
    throwIfAborted(signal);
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Canvas 2D context is unavailable');
    context.drawImage(image, 0, 0, width, height);
    const bytes = await exportCanvasPng(canvas, signal);
    if (inspectRasterImage(bytes).contentType !== 'image/png') {
      throw new Error('Canvas did not export PNG data');
    }
    return bytes;
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

async function decodeImage(image: HTMLImageElement, signal?: AbortSignal): Promise<void> {
  if (typeof image.decode === 'function') {
    await abortable(image.decode(), signal);
    return;
  }
  await abortable(new Promise<void>((resolve, reject) => {
    image.onload = () => resolve();
    image.onerror = () => reject(new Error('SVG image decode failed'));
  }), signal);
}

async function exportCanvasPng(
  canvas: HTMLCanvasElement,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  if (typeof canvas.toBlob === 'function') {
    const blob = await abortable(new Promise<Blob>((resolve, reject) => {
      canvas.toBlob((value) => {
        if (value) resolve(value);
        else reject(new Error('Canvas PNG export failed'));
      }, 'image/png');
    }), signal);
    const buffer = await abortable(blob.arrayBuffer(), signal);
    return new Uint8Array(buffer);
  }
  if (typeof canvas.toDataURL !== 'function') {
    throw new Error('Canvas PNG export is unavailable');
  }
  const resolved = await resolveImageSource(canvas.toDataURL('image/png'), signal);
  if (resolved.info.contentType !== 'image/png') {
    throw new Error('Canvas data URL is not PNG');
  }
  return new Uint8Array(resolved.bytes);
}

function abortable<T>(operation: Promise<T>, signal?: AbortSignal): Promise<T> {
  throwIfAborted(signal);
  if (!signal) return operation;
  return new Promise<T>((resolve, reject) => {
    const abort = (): void => reject(abortReason(signal));
    signal.addEventListener('abort', abort, { once: true });
    operation.then(
      (value) => {
        signal.removeEventListener('abort', abort);
        resolve(value);
      },
      (error: unknown) => {
        signal.removeEventListener('abort', abort);
        reject(error);
      },
    );
  });
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw abortReason(signal);
}

function abortReason(signal: AbortSignal): unknown {
  return signal.reason instanceof Error
    ? signal.reason
    : new DOMException('SVG fallback generation was aborted', 'AbortError');
}
