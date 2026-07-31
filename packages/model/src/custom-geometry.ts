import type { AddShapeOptions } from './preset-shape.js';

export interface CustomGeometryPoint {
  readonly x: number;
  readonly y: number;
}

export type CustomGeometryCommand =
  | { readonly kind: 'moveTo'; readonly point: CustomGeometryPoint }
  | { readonly kind: 'lineTo'; readonly point: CustomGeometryPoint }
  | {
      readonly kind: 'arcTo';
      readonly widthRadius: number;
      readonly heightRadius: number;
      readonly startAngle: number;
      readonly sweepAngle: number;
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
  readonly paths: readonly CustomGeometryPath[];
}

export type AddCustomShapeOptions = Omit<AddShapeOptions, 'adjustments'>;
