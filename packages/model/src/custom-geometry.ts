import type { AddShapeOptions } from './preset-shape.js';

export type CustomGeometryValue = number | string;

export type CustomGeometryFormula =
  | {
      readonly operator: 'val' | 'abs' | 'sqrt';
      readonly operands: readonly [CustomGeometryValue];
    }
  | {
      readonly operator: 'at2' | 'cos' | 'max' | 'min' | 'sin' | 'tan';
      readonly operands: readonly [CustomGeometryValue, CustomGeometryValue];
    }
  | {
      readonly operator: '*/' | '+-' | '+/' | '?:' | 'cat2' | 'mod' | 'pin' | 'sat2';
      readonly operands: readonly [
        CustomGeometryValue,
        CustomGeometryValue,
        CustomGeometryValue,
      ];
    };

export interface CustomGeometryGuide {
  readonly name: string;
  readonly formula: CustomGeometryFormula;
}

export interface CustomGeometryPoint {
  readonly x: CustomGeometryValue;
  readonly y: CustomGeometryValue;
}

export interface CustomGeometryXyHandle {
  readonly kind: 'xy';
  readonly position: CustomGeometryPoint;
  readonly xGuide?: string;
  readonly minX?: CustomGeometryValue;
  readonly maxX?: CustomGeometryValue;
  readonly yGuide?: string;
  readonly minY?: CustomGeometryValue;
  readonly maxY?: CustomGeometryValue;
}

export interface CustomGeometryPolarHandle {
  readonly kind: 'polar';
  readonly position: CustomGeometryPoint;
  readonly radiusGuide?: string;
  readonly minRadius?: CustomGeometryValue;
  readonly maxRadius?: CustomGeometryValue;
  readonly angleGuide?: string;
  readonly minAngle?: CustomGeometryValue;
  readonly maxAngle?: CustomGeometryValue;
}

export type CustomGeometryHandle = CustomGeometryXyHandle | CustomGeometryPolarHandle;

export interface CustomGeometryConnectionSite {
  readonly position: CustomGeometryPoint;
  readonly angle: CustomGeometryValue;
}

export interface CustomGeometryTextRectangle {
  readonly left: CustomGeometryValue;
  readonly top: CustomGeometryValue;
  readonly right: CustomGeometryValue;
  readonly bottom: CustomGeometryValue;
}

export type CustomGeometryCommand =
  | { readonly kind: 'moveTo'; readonly point: CustomGeometryPoint }
  | { readonly kind: 'lineTo'; readonly point: CustomGeometryPoint }
  | {
      readonly kind: 'arcTo';
      readonly widthRadius: CustomGeometryValue;
      readonly heightRadius: CustomGeometryValue;
      readonly startAngle: CustomGeometryValue;
      readonly sweepAngle: CustomGeometryValue;
    }
  | {
      readonly kind: 'quadraticBezierTo';
      readonly control: CustomGeometryPoint;
      readonly end: CustomGeometryPoint;
    }
  | {
      readonly kind: 'cubicBezierTo';
      readonly control1: CustomGeometryPoint;
      readonly control2: CustomGeometryPoint;
      readonly end: CustomGeometryPoint;
    }
  | { readonly kind: 'close' };

export type CustomGeometryPathFill =
  | 'none'
  | 'norm'
  | 'lighten'
  | 'lightenLess'
  | 'darken'
  | 'darkenLess';

export interface CustomGeometryPath {
  readonly width: number;
  readonly height: number;
  readonly fill?: CustomGeometryPathFill;
  readonly stroke?: boolean;
  readonly extrusionOk?: boolean;
  readonly commands: readonly CustomGeometryCommand[];
}

export interface CustomGeometry {
  readonly adjustments?: readonly CustomGeometryGuide[];
  readonly guides?: readonly CustomGeometryGuide[];
  readonly handles?: readonly CustomGeometryHandle[];
  readonly connectionSites?: readonly CustomGeometryConnectionSite[];
  readonly textRectangle?: CustomGeometryTextRectangle;
  readonly paths: readonly CustomGeometryPath[];
}

export type AddCustomShapeOptions = Omit<AddShapeOptions, 'adjustments'>;

export interface CustomGeometryEvaluationContext {
  readonly width: number;
  readonly height: number;
}

export type CustomGeometryEvaluationErrorCode =
  | 'unknown-token'
  | 'forward-reference'
  | 'cyclic-reference'
  | 'invalid-domain'
  | 'non-finite-result';

export class CustomGeometryEvaluationError extends Error {
  constructor(
    readonly code: CustomGeometryEvaluationErrorCode,
    message: string,
    readonly guideName?: string,
    readonly token?: string,
  ) {
    super(message);
    this.name = 'CustomGeometryEvaluationError';
  }
}

export interface EvaluatedCustomGeometryGuide {
  readonly name: string;
  readonly value: number;
}

export interface EvaluatedCustomGeometryPoint {
  readonly x: number;
  readonly y: number;
}

export interface EvaluatedCustomGeometryTextRectangle {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export type EvaluatedCustomGeometryCommand =
  | { readonly kind: 'moveTo'; readonly point: EvaluatedCustomGeometryPoint }
  | { readonly kind: 'lineTo'; readonly point: EvaluatedCustomGeometryPoint }
  | {
      readonly kind: 'arcTo';
      readonly widthRadius: number;
      readonly heightRadius: number;
      readonly startAngle: number;
      readonly sweepAngle: number;
    }
  | {
      readonly kind: 'quadraticBezierTo';
      readonly control: EvaluatedCustomGeometryPoint;
      readonly end: EvaluatedCustomGeometryPoint;
    }
  | {
      readonly kind: 'cubicBezierTo';
      readonly control1: EvaluatedCustomGeometryPoint;
      readonly control2: EvaluatedCustomGeometryPoint;
      readonly end: EvaluatedCustomGeometryPoint;
    }
  | { readonly kind: 'close' };

export interface EvaluatedCustomGeometryXyHandle {
  readonly kind: 'xy';
  readonly position: EvaluatedCustomGeometryPoint;
  readonly xGuide?: string;
  readonly minX?: number;
  readonly maxX?: number;
  readonly yGuide?: string;
  readonly minY?: number;
  readonly maxY?: number;
}

export interface EvaluatedCustomGeometryPolarHandle {
  readonly kind: 'polar';
  readonly position: EvaluatedCustomGeometryPoint;
  readonly radiusGuide?: string;
  readonly minRadius?: number;
  readonly maxRadius?: number;
  readonly angleGuide?: string;
  readonly minAngle?: number;
  readonly maxAngle?: number;
}

export type EvaluatedCustomGeometryHandle =
  | EvaluatedCustomGeometryXyHandle
  | EvaluatedCustomGeometryPolarHandle;

export interface EvaluatedCustomGeometryConnectionSite {
  readonly position: EvaluatedCustomGeometryPoint;
  readonly angle: number;
}

export interface EvaluatedCustomGeometryPath {
  readonly width: number;
  readonly height: number;
  readonly fill?: CustomGeometryPathFill;
  readonly stroke?: boolean;
  readonly extrusionOk?: boolean;
  readonly commands: readonly EvaluatedCustomGeometryCommand[];
}

export interface EvaluatedCustomGeometry {
  readonly context: CustomGeometryEvaluationContext;
  readonly adjustments?: readonly EvaluatedCustomGeometryGuide[];
  readonly guides?: readonly EvaluatedCustomGeometryGuide[];
  readonly handles?: readonly EvaluatedCustomGeometryHandle[];
  readonly connectionSites?: readonly EvaluatedCustomGeometryConnectionSite[];
  readonly textRectangle: EvaluatedCustomGeometryTextRectangle;
  readonly paths: readonly EvaluatedCustomGeometryPath[];
}
