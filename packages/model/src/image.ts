import type { Transform } from './units.js';
import type { PlaceholderSelector } from './placeholder.js';

export type RasterImageContentType = 'image/png' | 'image/jpeg' | 'image/gif';
export type SvgImageContentType = 'image/svg+xml';

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
  readonly placeholder?: PlaceholderSelector;
  readonly sourceRectangle?: ImageSourceRectangle;
}

export interface AddSvgImageOptions extends Partial<Transform> {
  readonly name?: string;
  readonly altText?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly sourceRectangle?: ImageSourceRectangle;
}
