import type { TransformInput } from './units.js';
import type { PlaceholderSelector } from './placeholder.js';
import type { ShapeShadow } from './preset-shape.js';

export type RasterImageContentType = 'image/png' | 'image/jpeg' | 'image/gif';
export type SvgImageContentType = 'image/svg+xml';

export interface ImageSourceRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface AddImageOptions extends Partial<TransformInput> {
  readonly contentType: RasterImageContentType;
  readonly name?: string;
  readonly altText?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly sourceRectangle?: ImageSourceRectangle;
  readonly rounding?: boolean;
  readonly shadow?: ShapeShadow;
  readonly transparency?: number;
}

export interface AddSvgImageOptions extends Partial<TransformInput> {
  readonly name?: string;
  readonly altText?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly sourceRectangle?: ImageSourceRectangle;
  readonly rounding?: boolean;
  readonly shadow?: ShapeShadow;
  readonly transparency?: number;
}
