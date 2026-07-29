import {
  CodecRegistry,
  GradientCodec,
  MasterLayoutThemeCodec,
  MediaCodec,
  type AddMediaOptions,
  type MediaModel,
  type MediaSource,
} from '@pptx/codecs';
import { PresentationModel } from '@pptx/model';
import { OpcPackage, type PackageOpenOptions } from '@pptx/opc';
import {
  ValidationError,
  validatePackage,
  type CompatibilityProfile,
  type Diagnostic,
} from '@pptx/validator';
import { createPresentationPackage, type CreatePresentationOptions } from './create.js';

export * from '@pptx/codecs';
export * from '@pptx/model';
export type { BuiltInSlideSize, CreatePresentationOptions, CustomSlideSize } from './create.js';
export { PackageError } from '@pptx/opc';
export type { PackageOpenOptions } from '@pptx/opc';
export { ValidationError } from '@pptx/validator';
export type { CompatibilityProfile, Diagnostic } from '@pptx/validator';
export { ModelParseError as ParseError, SlideModel as Slide, SlideTitleModel as SlideTitle } from '@pptx/model';

export type PptxByteChunk = number | Uint8Array | ArrayBuffer | ArrayBufferView;
export type PptxByteStream = ReadableStream<PptxByteChunk> | AsyncIterable<PptxByteChunk>;
export type PptxInput = string | Uint8Array | ArrayBuffer | Blob | PptxByteStream;

export interface WriteOptions {
  readonly compatibility?: CompatibilityProfile;
  readonly mode?: 'strict' | 'permissive';
}

export class OpaqueMutationError extends Error {
  constructor(message: string, readonly partUri?: string) {
    super(partUri ? `${message}: ${partUri}` : message);
    this.name = 'OpaqueMutationError';
  }
}

export class PptxDocument extends PresentationModel {
  readonly diagnostics: Diagnostic[] = [];
  readonly codecRegistry = new CodecRegistry();
  readonly #masterLayoutTheme: MasterLayoutThemeCodec;

  private constructor(opcPackage: OpcPackage) {
    super(opcPackage);
    this.#masterLayoutTheme = new MasterLayoutThemeCodec(opcPackage, this.presentationPartUri);
    this.codecRegistry.register(this.#masterLayoutTheme);
    this.codecRegistry.register(new GradientCodec());
    this.codecRegistry.register(new MediaCodec(opcPackage));
  }

  static async open(input: PptxInput, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    const bytes = await readInput(input, options.signal);
    const pkg = await OpcPackage.open(bytes, options);
    return new PptxDocument(pkg);
  }

  static create(options: CreatePresentationOptions = {}): PptxDocument {
    const document = new PptxDocument(createPresentationPackage(options));
    if (options.author !== undefined) document.author = options.author;
    if (options.company !== undefined) document.company = options.company;
    if (options.title !== undefined) document.title = options.title;
    if (options.subject !== undefined) document.subject = options.subject;
    return document;
  }

  static async fromBuffer(input: Uint8Array | ArrayBuffer, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    return this.open(input, options);
  }

  transaction<T>(operation: (document: this) => T): T {
    return this.opcPackage.transaction(() => {
      const result = operation(this);
      const diagnostics = validatePackage(this.opcPackage);
      if (diagnostics.some(({ severity }) => severity === 'error')) throw new ValidationError(diagnostics);
      return result;
    });
  }

  async write(options: WriteOptions = {}): Promise<Uint8Array> {
    const compatibility = options.compatibility ?? 'powerpoint-current';
    const diagnostics: Diagnostic[] = [...validatePackage(this.opcPackage)];
    const gradients = new GradientCodec();
    const media = new MediaCodec(this.opcPackage);
    for (const slide of this.slides) {
      const background = gradients.getSlideBackground(this.opcPackage, slide.partUri);
      if (background) diagnostics.push(...gradients.diagnostics(background, compatibility, slide.partUri));
      for (const model of media.list(slide.partUri)) {
        diagnostics.push(...media.diagnostics(model, compatibility));
      }
    }
    this.diagnostics.splice(0, this.diagnostics.length, ...diagnostics);
    if ((options.mode ?? 'strict') === 'strict' && diagnostics.some(({ severity }) => severity === 'error')) {
      throw new ValidationError(diagnostics);
    }
    return this.opcPackage.write();
  }

  async writeFile(path: string, options: WriteOptions = {}): Promise<void> {
    const fs = await loadNodeModule<NodeFsPromises>(['node:fs', 'promises'].join('/'));
    await fs.writeFile(path, await this.write(options));
  }

  async writeBlob(options: WriteOptions = {}): Promise<Blob> {
    const bytes = await this.write(options);
    return new Blob([new Uint8Array(bytes).buffer], { type: this.formatProfile.fileContentType });
  }

  async download(fileName = `presentation${this.formatProfile.extension}`, options: WriteOptions = {}): Promise<void> {
    if (typeof document === 'undefined' || typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
      throw new Error('PptxDocument.download() requires a browser DOM; use writeFile() in Node.js');
    }
    const url = URL.createObjectURL(await this.writeBlob(options));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.style.display = 'none';
    document.body?.appendChild(anchor);
    try {
      anchor.click();
    } finally {
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 0);
    }
  }

  get masters() {
    return this.#masterLayoutTheme.masters;
  }

  get layouts() {
    return this.#masterLayoutTheme.layouts;
  }

  get themes() {
    return this.#masterLayoutTheme.themes;
  }

  get masterLayoutTheme(): MasterLayoutThemeCodec {
    return this.#masterLayoutTheme;
  }

  async addAudio(slideIndex: number, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    return new MediaCodec(this.opcPackage).addAudio(slide.partUri, source, options);
  }

  async addVideo(slideIndex: number, source: MediaSource, options: AddMediaOptions = {}): Promise<MediaModel> {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    return new MediaCodec(this.opcPackage).addVideo(slide.partUri, source, options);
  }

  media(slideIndex: number): readonly MediaModel[] {
    const slide = this.slides[slideIndex];
    if (!slide) throw new RangeError(`Slide index ${slideIndex} is out of range`);
    return new MediaCodec(this.opcPackage).list(slide.partUri);
  }
}

export async function* openPptxStream(path: string): AsyncIterable<Uint8Array> {
  const fs = await loadNodeModule<NodeFs>(['node:fs'].join('/'));
  for await (const chunk of fs.createReadStream(path)) yield normalizeByteChunk(chunk);
}

async function readInput(input: PptxInput, signal?: AbortSignal): Promise<Uint8Array> {
  if (typeof input === 'string') {
    const fs = await loadNodeModule<NodeFsPromises>(['node:fs', 'promises'].join('/'));
    return new Uint8Array(await fs.readFile(input, signal ? { signal } : undefined));
  }
  if (input instanceof Uint8Array) return new Uint8Array(input);
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof Blob !== 'undefined' && input instanceof Blob) return new Uint8Array(await input.arrayBuffer());
  const chunks: Uint8Array[] = [];
  if (isReadableStream(input)) {
    const reader = input.getReader();
    try {
      while (true) {
        throwIfAborted(signal);
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(normalizeByteChunk(value));
      }
    } finally {
      reader.releaseLock();
    }
    return concatenateBytes(chunks);
  }
  if (!isAsyncIterable(input)) throw new TypeError('Unsupported PPTX input type');
  for await (const chunk of input) {
    if (signal?.aborted) throw signal.reason instanceof Error ? signal.reason : new DOMException('Aborted', 'AbortError');
    chunks.push(normalizeByteChunk(chunk));
  }
  return concatenateBytes(chunks);
}

interface NodeFsPromises {
  readFile(path: string, options?: { readonly signal?: AbortSignal }): Promise<Uint8Array>;
  writeFile(path: string, bytes: Uint8Array): Promise<void>;
}

interface NodeFs {
  createReadStream(path: string): AsyncIterable<unknown>;
}

async function loadNodeModule<T>(specifier: string): Promise<T> {
  if (typeof process === 'undefined' || !process.versions?.node) {
    throw new Error('Local file paths are only supported in Node.js; pass a Blob, File, or byte stream');
  }
  return import(specifier) as Promise<T>;
}

function isReadableStream(value: unknown): value is ReadableStream<unknown> {
  return Boolean(value && typeof (value as { getReader?: unknown }).getReader === 'function');
}

function isAsyncIterable(value: unknown): value is AsyncIterable<unknown> {
  return Boolean(value && typeof (value as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === 'function');
}

function normalizeByteChunk(chunk: unknown): Uint8Array {
  if (typeof chunk === 'number' && Number.isInteger(chunk) && chunk >= 0 && chunk <= 255) {
    return Uint8Array.of(chunk);
  }
  if (chunk instanceof Uint8Array) return new Uint8Array(chunk);
  if (chunk instanceof ArrayBuffer) return new Uint8Array(chunk);
  if (ArrayBuffer.isView(chunk)) {
    return new Uint8Array(new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength));
  }
  throw new TypeError('PPTX streams must yield byte numbers, Uint8Array, ArrayBuffer, or ArrayBufferView chunks');
}

function concatenateBytes(chunks: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.byteLength, 0));
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) {
    throw signal.reason instanceof Error ? signal.reason : new DOMException('Aborted', 'AbortError');
  }
}
