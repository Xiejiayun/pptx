import type { Transform } from './units.js';

export type RasterImageContentType = 'image/png' | 'image/jpeg' | 'image/gif';

export interface ImageSourceRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface AddImageOptions extends Partial<Transform> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
  readonly sourceRectangle?: ImageSourceRectangle;
}
