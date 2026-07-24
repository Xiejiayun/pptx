import { createReadStream, promises as fs } from 'node:fs';
import { Readable } from 'node:stream';
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

export * from '@pptx/codecs';
export * from '@pptx/model';
export { PackageError } from '@pptx/opc';
export type { PackageOpenOptions } from '@pptx/opc';
export { ValidationError } from '@pptx/validator';
export { ModelParseError as ParseError, SlideModel as Slide, SlideTitleModel as SlideTitle } from '@pptx/model';

export type PptxInput = string | Uint8Array | ArrayBuffer | Readable;

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

  private constructor(opcPackage: OpcPackage) {
    super(opcPackage);
    this.codecRegistry.register(new MasterLayoutThemeCodec(opcPackage, this.presentationPartUri));
    this.codecRegistry.register(new GradientCodec());
    this.codecRegistry.register(new MediaCodec(opcPackage));
  }

  static async open(input: PptxInput, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    const bytes = await readInput(input, options.signal);
    const pkg = await OpcPackage.open(bytes, options);
    return new PptxDocument(pkg);
  }

  static async fromBuffer(input: Uint8Array | ArrayBuffer, options: PackageOpenOptions = {}): Promise<PptxDocument> {
    return this.open(input, options);
  }

  async write(options: WriteOptions = {}): Promise<Uint8Array> {
    const compatibility = options.compatibility ?? 'powerpoint-current';
    const diagnostics: Diagnostic[] = [...validatePackage(this.opcPackage)];
    const gradients = new GradientCodec();
    const media = new MediaCodec(this.opcPackage);
    for (const slide of this.slides) {
      const background = gradients.getSlideBackground(this.opcPackage, slide.partUri);
      if (background) diagnostics.push(...gradients.diagnostics(background, compatibility, slide.partUri));
      for (const model of media.list(slide.partUri)) diagnostics.push(...media.diagnostics(model, compatibility));
    }
    this.diagnostics.splice(0, this.diagnostics.length, ...diagnostics);
    if ((options.mode ?? 'strict') === 'strict' && diagnostics.some(({ severity }) => severity === 'error')) {
      throw new ValidationError(diagnostics);
    }
    return this.opcPackage.write();
  }

  async writeFile(path: string, options: WriteOptions = {}): Promise<void> {
    await fs.writeFile(path, await this.write(options));
  }

  get masters() {
    return new MasterLayoutThemeCodec(this.opcPackage, this.presentationPartUri).masters;
  }

  get layouts() {
    return new MasterLayoutThemeCodec(this.opcPackage, this.presentationPartUri).layouts;
  }

  get themes() {
    return new MasterLayoutThemeCodec(this.opcPackage, this.presentationPartUri).themes;
  }

  get masterLayoutTheme(): MasterLayoutThemeCodec {
    return new MasterLayoutThemeCodec(this.opcPackage, this.presentationPartUri);
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

export function openPptxStream(path: string): Readable {
  return createReadStream(path);
}

async function readInput(input: PptxInput, signal?: AbortSignal): Promise<Uint8Array> {
  if (typeof input === 'string') return new Uint8Array(await fs.readFile(input, { signal }));
  if (input instanceof Uint8Array) return new Uint8Array(input);
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  const chunks: Buffer[] = [];
  for await (const chunk of input) {
    if (signal?.aborted) throw signal.reason instanceof Error ? signal.reason : new DOMException('Aborted', 'AbortError');
    if (typeof chunk === 'number') chunks.push(Buffer.from([chunk]));
    else chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array));
  }
  return new Uint8Array(Buffer.concat(chunks));
}
