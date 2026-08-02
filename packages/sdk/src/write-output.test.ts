import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { convertWriteOutput, resolveWriteOutputType } from './write-output.js';

const BYTES = Uint8Array.from(
  { length: 65_541 },
  (_, index) => [0x00, 0x7f, 0x80, 0xff, index & 0xff][index % 5]!,
);

describe('presentation write output conversion', () => {
  it('converts canonical bytes into all six Node output representations', async () => {
    const arraybuffer = await convertWriteOutput(BYTES, 'arraybuffer');
    const base64 = await convertWriteOutput(BYTES, 'base64');
    const binarystring = await convertWriteOutput(BYTES, 'binarystring');
    const blob = await convertWriteOutput(BYTES, 'blob');
    const nodebuffer = await convertWriteOutput(BYTES, 'nodebuffer');
    const uint8array = await convertWriteOutput(BYTES, 'uint8array');

    arraybuffer satisfies ArrayBuffer;
    base64 satisfies string;
    binarystring satisfies string;
    blob satisfies Blob;
    nodebuffer satisfies Uint8Array;
    uint8array satisfies Uint8Array;

    expect(new Uint8Array(arraybuffer)).toEqual(BYTES);
    expect(Uint8Array.from(Buffer.from(base64, 'base64'))).toEqual(BYTES);
    expect(Uint8Array.from(binarystring, (value) => value.charCodeAt(0))).toEqual(BYTES);
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(BYTES);
    expect(blob.type).toBe('application/zip');
    expect(Buffer.isBuffer(nodebuffer)).toBe(true);
    expect(new Uint8Array(nodebuffer)).toEqual(BYTES);
    expect(Buffer.isBuffer(uint8array)).toBe(false);
    expect(uint8array).toEqual(BYTES);
  });

  it('resolves only exact public output type tokens', () => {
    expect(resolveWriteOutputType(undefined)).toBe('uint8array');
    for (const outputType of [
      'arraybuffer',
      'base64',
      'binarystring',
      'blob',
      'nodebuffer',
      'uint8array',
    ] as const) {
      expect(resolveWriteOutputType(outputType)).toBe(outputType);
    }
    for (const invalid of ['STREAM', 'buffer', 'BLOB', null, 1, {}, Symbol('blob')]) {
      expect(() => resolveWriteOutputType(invalid)).toThrow(
        new TypeError('PptxDocument.write() received an unsupported outputType'),
      );
    }
  });

  it('pads short base64 inputs and isolates ArrayBuffer byte ranges', async () => {
    expect(await convertWriteOutput(new Uint8Array(), 'base64')).toBe('');
    expect(await convertWriteOutput(Uint8Array.of(0xff), 'base64')).toBe('/w==');
    expect(await convertWriteOutput(Uint8Array.of(0xff, 0xee), 'base64')).toBe('/+4=');

    const view = Uint8Array.of(9, 1, 2, 8).subarray(1, 3);
    const output = await convertWriteOutput(view, 'arraybuffer');
    expect(output.byteLength).toBe(2);
    expect(new Uint8Array(output)).toEqual(Uint8Array.of(1, 2));
  });
});
