export const OUTPUT_TYPES = Object.freeze([
  'arraybuffer',
  'base64',
  'binarystring',
  'blob',
  'nodebuffer',
  'uint8array',
] as const);

export type OutputType = typeof OUTPUT_TYPES[number];

export type WriteOutput<TOutputType extends OutputType = OutputType> =
  TOutputType extends 'arraybuffer' ? ArrayBuffer
    : TOutputType extends 'base64' | 'binarystring' ? string
      : TOutputType extends 'blob' ? Blob
        : Uint8Array;
