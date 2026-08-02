import { Readable, Writable } from 'node:stream';
import { describe, expect, it, vi } from 'vitest';
import {
  PptxDocument,
  ShapeModel,
  ValidationError,
  type PptxNodeReadableStream,
} from './index.js';

const STREAM_CHUNK_SIZE = 64 * 1024;

describe('PptxDocument Node readable output', () => {
  it('streams canonical bytes through iteration, events, and pipe', async () => {
    const document = createLargeDocument();
    const expected = await document.write();
    const journal = [...document.opcPackage.mutations];

    const iterableStream = await document.stream();
    expect(iterableStream).toBeInstanceOf(Readable);
    expect(iterableStream.readableObjectMode).toBe(false);
    expect(await collect(iterableStream)).toEqual(expected);
    expect(document.opcPackage.mutations).toEqual(journal);

    const eventStream = await document.stream();
    const eventChunks = await collectEvents(eventStream);
    expect(eventChunks.length).toBeGreaterThan(1);
    expect(eventChunks.every(({ byteLength }) => byteLength <= STREAM_CHUNK_SIZE)).toBe(true);
    expect(concatenate(eventChunks)).toEqual(expected);

    const pipeStream = await document.stream();
    const pipedChunks: Uint8Array[] = [];
    const destination = new Writable({
      write(chunk: Uint8Array, _encoding, callback) {
        pipedChunks.push(new Uint8Array(chunk));
        callback();
      },
    });
    const finished = new Promise<void>((resolve, reject) => {
      destination.once('finish', resolve);
      destination.once('error', reject);
      pipeStream.once('error', reject);
    });
    expect(pipeStream.pipe(destination)).toBe(destination);
    await finished;
    const piped = concatenate(pipedChunks);
    expect(piped).toEqual(expected);
    const reopened = await PptxDocument.open(piped);
    const shape = reopened.slides[0]?.shapes[0];
    expect(shape).toBeInstanceOf(ShapeModel);
    expect((shape as ShapeModel).text).toBe('Node stream 你好');

    const destroyed = await document.stream();
    const closed = new Promise<void>((resolve) => destroyed.once('close', resolve));
    destroyed.destroy();
    await closed;
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(await document.write()).toEqual(expected);
  }, 15_000);

  it('captures document state when the stream promise resolves', async () => {
    const document = PptxDocument.create();
    document.addSlide().addText('Captured');

    const captured = await document.stream();
    document.addSlide().addText('Later');

    const first = await PptxDocument.open(await collect(captured));
    expect(first.slides).toHaveLength(1);
    expect((first.slides[0]?.shapes[0] as ShapeModel).text).toBe('Captured');

    const second = await PptxDocument.open(await collect(await document.stream()));
    expect(second.slides).toHaveLength(2);
    expect(second.slides.map((slide) => (slide.shapes[0] as ShapeModel).text))
      .toEqual(['Captured', 'Later']);
  });

  it('shares strict and permissive validation behavior with write', async () => {
    const document = PptxDocument.create();
    document.addSlide();
    document.opcPackage.setPart(
      '/_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" '
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
      + 'Target="ppt/missing.xml"/></Relationships>',
    );

    await expect(document.stream()).rejects.toBeInstanceOf(ValidationError);
    const strictDiagnostics = document.diagnostics.filter(({ severity }) => severity === 'error');
    expect(strictDiagnostics.length).toBeGreaterThan(0);

    const permissiveBytes = await collect(await document.stream({ mode: 'permissive' }));
    expect(permissiveBytes.byteLength).toBeGreaterThan(0);
    const streamDiagnostics = [...document.diagnostics];
    await document.write({ mode: 'permissive' });
    expect(document.diagnostics).toEqual(streamDiagnostics);
  });

  it('rejects outside Node before diagnostics or package writes', async () => {
    const document = PptxDocument.create();
    const diagnostics = [...document.diagnostics];
    const journal = [...document.opcPackage.mutations];
    const write = vi.spyOn(document.opcPackage, 'write');
    vi.stubGlobal('process', undefined);
    try {
      await expect(document.stream()).rejects.toThrow(
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

  it('exposes a browser-safe structural stream type', () => {
    const document = PptxDocument.create();
    if (false) {
      document.stream() satisfies Promise<PptxNodeReadableStream>;
      document.stream({ mode: 'permissive' }) satisfies Promise<PptxNodeReadableStream>;
      document.stream({ compatibility: 'powerpoint-current' }) satisfies Promise<PptxNodeReadableStream>;
      void document.stream().then((readable) => {
        readable satisfies AsyncIterable<Uint8Array>;
        const destination = { tag: 'destination' } as const;
        readable.pipe(destination) satisfies typeof destination;
        readable.pause().resume().destroy();
      });
      // @ts-expect-error stream does not consume write output selectors
      document.stream({ outputType: 'uint8array' });
    }
  });
});

function createLargeDocument(): PptxDocument {
  const document = PptxDocument.create();
  document.addSlide().addText('Node stream 你好');
  const payload = new Uint8Array(98_333);
  let state = 0x1234_5678;
  for (let index = 0; index < payload.byteLength; index += 1) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    payload[index] = state & 0xff;
  }
  document.opcPackage.setPart(
    '/custom/stream.bin',
    payload,
    'application/octet-stream',
  );
  return document;
}

async function collect(source: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of source) chunks.push(new Uint8Array(chunk));
  return concatenate(chunks);
}

function collectEvents(source: PptxNodeReadableStream): Promise<Uint8Array[]> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    source.on('data', (chunk) => chunks.push(new Uint8Array(chunk)));
    source.once('end', () => resolve(chunks));
    source.once('error', reject);
  });
}

function concatenate(chunks: readonly Uint8Array[]): Uint8Array {
  const length = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const bytes = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}
