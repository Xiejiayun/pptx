import type { Transform } from './units.js';

export type RasterImageContentType = 'image/png' | 'image/jpeg' | 'image/gif';

export interface AddImageOptions extends Partial<Transform> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
}
