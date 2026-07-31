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
  readonly paths: readonly CustomGeometryPath[];
}

export type AddCustomShapeOptions = Omit<AddShapeOptions, 'adjustments'>;
