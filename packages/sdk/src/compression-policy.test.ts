import { Buffer } from 'node:buffer';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import {
  PptxDocument,
  ShapeModel,
  type OutputType,
  type WriteBaseOptions,
  type WriteOptions,
} from './index.js';

describe('PptxDocument compression policy', () => {
  it('uses STORE by default and DEFLATE level 6 when requested', async () => {
    const document = createCompressionDocument();
    const journal = [...document.opcPackage.mutations];
    const defaultBytes = await document.write();
    const explicitUndefined = await document.write({
      compression: undefined,
    } as unknown as WriteOptions);
    const storedBytes = await document.write({ compression: false });
    const deflatedBytes = await document.write({ compression: true });

    expect(defaultBytes).toEqual(storedBytes);
    expect(explicitUndefined).toEqual(storedBytes);
    expect(new Set(zipCompressionMethods(storedBytes))).toEqual(new Set([0]));
    expect(new Set(zipCompressionMethods(deflatedBytes))).toEqual(new Set([8]));
    expect(deflatedBytes.byteLength).toBeLessThan(storedBytes.byteLength);
    expect(await document.write({ compression: false })).toEqual(storedBytes);
    expect(await document.write({ compression: true })).toEqual(deflatedBytes);
    expect(document.opcPackage.mutations).toEqual(journal);

    const reopenedStore = await PptxDocument.open(storedBytes);
    const reopenedDeflate = await PptxDocument.open(deflatedBytes);
    expect((reopenedStore.slides[0]?.shapes[0] as ShapeModel).text)
      .toBe('Compression policy 你好');
    expect((reopenedDeflate.slides[0]?.shapes[0] as ShapeModel).text)
      .toBe('Compression policy 你好');
  }, 15_000);

  it('keeps compression orthogonal to every output representation and convenience path', async () => {
    const document = createCompressionDocument();
    const journal = [...document.opcPackage.mutations];
    const deflated = await document.write({ compression: true });
    const outputs: readonly (readonly [OutputType, unknown])[] = [
      ['arraybuffer', await document.write({ outputType: 'arraybuffer', compression: true })],
      ['base64', await document.write({ outputType: 'base64', compression: true })],
      ['binarystring', await document.write({ outputType: 'binarystring', compression: true })],
      ['blob', await document.write({ outputType: 'blob', compression: true })],
      ['nodebuffer', await document.write({ outputType: 'nodebuffer', compression: true })],
      ['uint8array', await document.write({ outputType: 'uint8array', compression: true })],
    ];
    for (const [outputType, output] of outputs) {
      expect(await decodeOutput(outputType, output)).toEqual(deflated);
    }

    expect(await collect(await document.stream({ compression: true }))).toEqual(deflated);
    const presentationBlob = await document.writeBlob({ compression: true });
    expect(new Uint8Array(await presentationBlob.arrayBuffer())).toEqual(deflated);
    expect(presentationBlob.type).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    );

    const directory = await mkdtemp(join(tmpdir(), 'pptx-compression-policy-'));
    try {
      const path = join(directory, 'deflated.pptx');
      await document.writeFile(path, { compression: true });
      expect(new Uint8Array(await readFile(path))).toEqual(deflated);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }

    const stored = await document.write({ compression: false });
    expect(await collect(await document.stream({ compression: false }))).toEqual(stored);
    expect(new Uint8Array(await (await document.writeBlob({ compression: false })).arrayBuffer()))
      .toEqual(stored);
    expect(document.opcPackage.mutations).toEqual(journal);
  }, 15_000);

  it('preserves unchanged inputs only when compression is omitted', async () => {
    const source = createCompressionDocument();
    const deflatedInput = await source.write({ compression: true });
    const fromDeflate = await PptxDocument.open(deflatedInput);
    expect(await fromDeflate.write()).toEqual(deflatedInput);
    expect(await fromDeflate.write({ compression: undefined } as never)).toEqual(deflatedInput);
    expect(new Set(zipCompressionMethods(
      await fromDeflate.write({ compression: false }),
    ))).toEqual(new Set([0]));

    const storedInput = await source.write({ compression: false });
    const fromStore = await PptxDocument.open(storedInput);
    expect(await fromStore.write()).toEqual(storedInput);
    expect(new Set(zipCompressionMethods(
      await fromStore.write({ compression: true }),
    ))).toEqual(new Set([8]));
    expect(fromStore.opcPackage.mutations).toHaveLength(0);
  });

  it('rejects non-boolean selectors before diagnostics and package writes', async () => {
    for (const compression of [
      'true',
      1,
      0,
      null,
      {},
      [],
      new Boolean(true),
    ]) {
      const document = createCompressionDocument();
      const diagnostics = [...document.diagnostics];
      const journal = [...document.opcPackage.mutations];
      const write = vi.spyOn(document.opcPackage, 'write');
      await expect(document.write({ compression } as never)).rejects.toThrow(
        new TypeError('PptxDocument output compression must be a boolean'),
      );
      expect(write).not.toHaveBeenCalled();
      expect(document.diagnostics).toEqual(diagnostics);
      expect(document.opcPackage.mutations).toEqual(journal);
      write.mockRestore();
    }
  });

  it('keeps the browser stream preflight ahead of compression validation', async () => {
    const document = createCompressionDocument();
    const diagnostics = [...document.diagnostics];
    const journal = [...document.opcPackage.mutations];
    const write = vi.spyOn(document.opcPackage, 'write');
    vi.stubGlobal('process', undefined);
    try {
      await expect(document.stream({ compression: 'true' } as never)).rejects.toThrow(
        new Error('PptxDocument.stream() is only supported in Node.js'),
      );
      expect(write).not.toHaveBeenCalled();
      expect(document.diagnostics).toEqual(diagnostics);
      expect(document.opcPackage.mutations).toEqual(journal);
    } finally {
      vi.unstubAllGlobals();
      write.mockRestore();
    }
  });

  it('exposes compression through every typed output surface', () => {
    const document = PptxDocument.create();
    if (false) {
      const base: WriteBaseOptions = { compression: true };
      const generic: WriteOptions<'blob'> = {
        outputType: 'blob',
        compression: false,
      };
      document.write(base) satisfies Promise<Uint8Array>;
      document.write(generic) satisfies Promise<Blob>;
      document.stream({ compression: true });
      document.writeFile('output.pptx', { compression: false });
      document.writeBlob({ compression: true });
      document.download('output.pptx', { compression: false });
      // @ts-expect-error compression is boolean-only
      document.write({ compression: 'true' });
      // @ts-expect-error stream compression is boolean-only
      document.stream({ compression: 1 });
      // @ts-expect-error file compression is boolean-only
      document.writeFile('output.pptx', { compression: null });
      // @ts-expect-error blob compression is boolean-only
      document.writeBlob({ compression: 'DEFLATE' });
      // @ts-expect-error download compression is boolean-only
      document.download('output.pptx', { compression: 1 });
    }
  });
});

function createCompressionDocument(): PptxDocument {
  const document = PptxDocument.create();
  document.addSlide().addText('Compression policy 你好');
  document.opcPackage.setPart(
    '/custom/compression-policy.bin',
    new Uint8Array(131_072).fill(0x41),
    'application/octet-stream',
  );
  return document;
}

async function decodeOutput(outputType: OutputType, output: unknown): Promise<Uint8Array> {
  if (outputType === 'arraybuffer') return new Uint8Array(output as ArrayBuffer);
  if (outputType === 'base64') return Uint8Array.from(Buffer.from(output as string, 'base64'));
  if (outputType === 'binarystring') {
    return Uint8Array.from(output as string, (character) => character.charCodeAt(0));
  }
  if (outputType === 'blob') return new Uint8Array(await (output as Blob).arrayBuffer());
  return new Uint8Array(output as Uint8Array);
}

async function collect(source: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of source) chunks.push(new Uint8Array(chunk));
  const length = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const bytes = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

function zipCompressionMethods(bytes: Uint8Array): number[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.byteLength - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x0605_4b50) eocd -= 1;
  if (eocd < 0) throw new Error('ZIP EOCD not found');
  const entries = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const methods: number[] = [];
  for (let index = 0; index < entries; index += 1) {
    if (view.getUint32(offset, true) !== 0x0201_4b50) {
      throw new Error('ZIP central directory entry not found');
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const name = new TextDecoder().decode(
      bytes.subarray(offset + 46, offset + 46 + nameLength),
    );
    if (!name.endsWith('/')) methods.push(view.getUint16(offset + 10, true));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return methods;
}
