export type ColorSource =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scrgb'; readonly red: number; readonly green: number; readonly blue: number }
  | { readonly kind: 'scheme'; readonly value: string; readonly resolved?: string }
  | { readonly kind: 'system'; readonly value: string; readonly lastColor?: string }
  | { readonly kind: 'preset'; readonly value: string };

export interface ColorTransform {
  readonly kind: string;
  readonly value: number;
}

export interface Color {
  readonly source: ColorSource;
  readonly transforms: readonly ColorTransform[];
  readonly alpha: number;
}

export interface InheritedValue<T> {
  readonly value: T;
  readonly source: 'local' | 'layout' | 'master' | 'theme' | 'default';
  readonly sourcePartUri?: string;
}

