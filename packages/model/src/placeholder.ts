import type { AddTextOptions } from './slide.js';

export const PLACEHOLDER_TYPES = [
  'title',
  'body',
  'pic',
  'chart',
  'tbl',
  'media',
] as const;

export type PlaceholderType = typeof PLACEHOLDER_TYPES[number];

export interface PlaceholderIdentity {
  readonly type: PlaceholderType;
  readonly index: number;
}

export type PlaceholderSelector = string | PlaceholderIdentity;

export interface AddPlaceholderOptions extends Omit<AddTextOptions, 'name' | 'placeholder'> {
  readonly name: string;
  readonly type: PlaceholderType;
  readonly index?: number;
}
