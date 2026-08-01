import type { GradientFill } from '@pptx/codecs';
import type { RasterImageContentType } from './image.js';
import type { RichTextColor } from './text.js';

export type SimpleFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export interface SlideBackgroundImage {
  readonly kind: 'image';
  readonly contentType: RasterImageContentType;
  readonly bytes: Uint8Array;
}

export type SlideBackground = SimpleFill | GradientFill | SlideBackgroundImage;
