import { createReadStream, promises as fs } from 'node:fs';
import { Readable } from 'node:stream';
import { PresentationModel } from '@pptx/model';
import { OpcPackage, type PackageOpenOptions } from '@pptx/opc';
import {
  ValidationError,
  validatePackage,
  type CompatibilityProfile,
  type Diagnostic,
} from '@pptx/validator';

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

  private constructor(opcPackage: OpcPackage) {
    super(opcPackage);
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
    const diagnostics = validatePackage(this.opcPackage);
    this.diagnostics.splice(0, this.diagnostics.length, ...diagnostics);
    if ((options.mode ?? 'strict') === 'strict' && diagnostics.some(({ severity }) => severity === 'error')) {
      throw new ValidationError(diagnostics);
    }
    return this.opcPackage.write();
  }

  async writeFile(path: string, options: WriteOptions = {}): Promise<void> {
    await fs.writeFile(path, await this.write(options));
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
