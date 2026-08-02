import {
  OUTPUT_TYPES,
  type OutputType,
  type WriteOutput,
} from './output-type.js';

const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
const STRING_CHUNK_SIZE = 32_768;

interface NodeBufferModule {
  readonly Buffer: {
    from(bytes: Uint8Array): Uint8Array;
  };
}

export function resolveWriteOutputType(value: unknown): OutputType {
  const outputType = value === undefined ? 'uint8array' : value;
  if (
    typeof outputType !== 'string'
    || !OUTPUT_TYPES.includes(outputType as OutputType)
  ) {
    throw new TypeError('PptxDocument.write() received an unsupported outputType');
  }
  if (outputType === 'nodebuffer' && !isNodeRuntime()) {
    throw new Error('nodebuffer is not supported by this platform');
  }
  return outputType as OutputType;
}

export async function convertWriteOutput<TOutputType extends OutputType>(
  bytes: Uint8Array,
  outputType: TOutputType,
): Promise<WriteOutput<TOutputType>> {
  switch (outputType) {
    case 'arraybuffer':
      return standaloneArrayBuffer(bytes) as WriteOutput<TOutputType>;
    case 'base64':
      return encodeBase64(bytes) as WriteOutput<TOutputType>;
    case 'binarystring':
      return encodeBinaryString(bytes) as WriteOutput<TOutputType>;
    case 'blob':
      return new Blob(
        [standaloneArrayBuffer(bytes)],
        { type: 'application/zip' },
      ) as WriteOutput<TOutputType>;
    case 'nodebuffer': {
      if (!isNodeRuntime()) throw new Error('nodebuffer is not supported by this platform');
      const { Buffer } = await loadNodeBufferModule();
      return Buffer.from(bytes) as WriteOutput<TOutputType>;
    }
    case 'uint8array':
      return bytes as WriteOutput<TOutputType>;
  }
}

function standaloneArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return Uint8Array.from(bytes).buffer;
}

function encodeBinaryString(bytes: Uint8Array): string {
  const chunks: string[] = [];
  for (let offset = 0; offset < bytes.byteLength; offset += STRING_CHUNK_SIZE) {
    chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + STRING_CHUNK_SIZE)));
  }
  return chunks.join('');
}

function encodeBase64(bytes: Uint8Array): string {
  const chunks: string[] = [];
  let chunk = '';
  let offset = 0;
  while (offset + 2 < bytes.byteLength) {
    const first = bytes[offset++]!;
    const second = bytes[offset++]!;
    const third = bytes[offset++]!;
    chunk += BASE64_ALPHABET.charAt(first >>> 2)
      + BASE64_ALPHABET.charAt(((first & 0x03) << 4) | (second >>> 4))
      + BASE64_ALPHABET.charAt(((second & 0x0f) << 2) | (third >>> 6))
      + BASE64_ALPHABET.charAt(third & 0x3f);
    if (chunk.length >= STRING_CHUNK_SIZE) {
      chunks.push(chunk);
      chunk = '';
    }
  }
  if (offset < bytes.byteLength) {
    const first = bytes[offset++]!;
    const second = offset < bytes.byteLength ? bytes[offset]! : undefined;
    chunk += BASE64_ALPHABET.charAt(first >>> 2)
      + BASE64_ALPHABET.charAt(((first & 0x03) << 4) | ((second ?? 0) >>> 4));
    chunk += second === undefined
      ? '=='
      : `${BASE64_ALPHABET.charAt((second & 0x0f) << 2)}=`;
  }
  if (chunk.length > 0) chunks.push(chunk);
  return chunks.join('');
}

function isNodeRuntime(): boolean {
  return typeof process !== 'undefined' && Boolean(process.versions?.node);
}

async function loadNodeBufferModule(): Promise<NodeBufferModule> {
  return import(['node:', 'buffer'].join('')) as Promise<NodeBufferModule>;
}
