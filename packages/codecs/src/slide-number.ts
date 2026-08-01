export type SlideNumberColor =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string };

export interface SlideNumberMargins {
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly left?: number;
}

export type SlideNumberMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | SlideNumberMargins;

export interface SlideNumberTextStyleOptions {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

export interface SlideNumberTextStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly lang: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly color?: SlideNumberColor;
  readonly transparency?: number;
}

export interface SlideNumberOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly align?: 'left' | 'center' | 'right' | 'justify';
  readonly rtl?: boolean;
  readonly valign?: 'top' | 'middle' | 'bottom';
  readonly margin?: SlideNumberMarginInput;
  readonly style?: SlideNumberTextStyleOptions;
}

export interface SlideNumber {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly align: 'left' | 'center' | 'right' | 'justify';
  readonly rtl: boolean;
  readonly valign?: 'top' | 'middle' | 'bottom';
  readonly margin?: SlideNumberMargins;
  readonly style: SlideNumberTextStyle;
}

export type SlideNumberOwnerKind = 'slide' | 'layout' | 'master';
