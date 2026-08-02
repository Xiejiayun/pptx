export const OUTPUT_TYPES = Object.freeze([
  'arraybuffer',
  'base64',
  'binarystring',
  'blob',
  'nodebuffer',
  'uint8array',
] as const);

export type OutputType = typeof OUTPUT_TYPES[number];
