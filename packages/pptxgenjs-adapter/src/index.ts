import { PptxDocument, type PackageOpenOptions } from '@pptx/sdk';

export interface PptxGenJSPresentation {
  write(options: {
    readonly outputType: 'uint8array';
    readonly compression: boolean;
  }): Promise<string | ArrayBuffer | Blob | Uint8Array>;
}

export interface ImportPptxGenJSOptions extends PackageOpenOptions {
  readonly compression?: boolean;
}

export class PptxGenJSAdapterError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PptxGenJSAdapterError';
  }
}

export async function importPptxGenJS(
  presentation: PptxGenJSPresentation,
  options: ImportPptxGenJSOptions = {},
): Promise<PptxDocument> {
  if (!presentation || typeof presentation.write !== 'function') {
    throw new PptxGenJSAdapterError('Expected a PptxGenJS instance with the public write() method');
  }
  const output = await presentation.write({
    outputType: 'uint8array',
    compression: options.compression ?? true,
  });
  const bytes = await normalizeOutput(output);
  const { compression: _compression, ...openOptions } = options;
  return PptxDocument.open(bytes, openOptions);
}

async function normalizeOutput(output: string | ArrayBuffer | Blob | Uint8Array): Promise<Uint8Array> {
  if (output instanceof Uint8Array) return output;
  if (output instanceof ArrayBuffer) return new Uint8Array(output);
  if (typeof Blob !== 'undefined' && output instanceof Blob) return new Uint8Array(await output.arrayBuffer());
  throw new PptxGenJSAdapterError(`PptxGenJS returned unexpected ${typeof output} output for uint8array`);
}
