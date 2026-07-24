export type Emu = number & { readonly __brand: 'Emu' };
export type OoxmlAngle = number & { readonly __brand: 'OoxmlAngle' };

export const EMU_PER_INCH = 914_400;
export const EMU_PER_POINT = 12_700;
export const OOXML_ANGLE_PER_DEGREE = 60_000;

export function inches(value: number): Emu {
  return Math.round(value * EMU_PER_INCH) as Emu;
}

export function points(value: number): Emu {
  return Math.round(value * EMU_PER_POINT) as Emu;
}

export function emuToInches(value: Emu | number): number {
  return value / EMU_PER_INCH;
}

export function degrees(value: number): OoxmlAngle {
  return Math.round(value * OOXML_ANGLE_PER_DEGREE) as OoxmlAngle;
}

export function angleToDegrees(value: OoxmlAngle | number): number {
  return value / OOXML_ANGLE_PER_DEGREE;
}

export interface Transform {
  readonly x: Emu;
  readonly y: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation: OoxmlAngle;
  readonly flipHorizontal: boolean;
  readonly flipVertical: boolean;
}

