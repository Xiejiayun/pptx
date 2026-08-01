export const CHART_TYPES = Object.freeze([
  'area',
  'bar',
  'bar3D',
  'bubble',
  'doughnut',
  'line',
  'pie',
  'radar',
  'scatter',
] as const);

export type ChartType = typeof CHART_TYPES[number];
export type ChartCategory = string | number;
export type ChartCategories =
  | readonly ChartCategory[]
  | readonly (readonly string[])[];

export interface ChartSeriesInput {
  readonly name: string;
  readonly categories?: ChartCategories;
  readonly values: readonly number[];
  readonly xValues?: readonly number[];
  readonly sizes?: readonly number[];
}

export interface ChartSeries extends ChartSeriesInput {
  readonly categories?: ChartCategories;
  readonly values: readonly number[];
  readonly xValues?: readonly number[];
  readonly sizes?: readonly number[];
}

export interface ChartGroupInput {
  readonly type: ChartType;
  readonly series: readonly ChartSeriesInput[];
  readonly axis?: 'primary' | 'secondary';
}

export interface ChartGroup {
  readonly type: ChartType;
  readonly series: readonly Readonly<ChartSeries>[];
  readonly axis?: 'primary' | 'secondary';
}

export interface ChartDefinitionInput {
  readonly groups: readonly ChartGroupInput[];
}

export interface ChartDefinition {
  readonly groups: readonly Readonly<ChartGroup>[];
}

export type ChartStateStatus =
  | 'recognized'
  | 'cache-only'
  | 'modern'
  | 'unsupported'
  | 'ambiguous';

export interface ChartState {
  readonly status: ChartStateStatus;
  readonly definition?: Readonly<ChartDefinition>;
  readonly workbookPartUri?: string;
  readonly reason?: string;
}

export interface AddChartOptions {
  readonly name?: string;
  readonly altText?: string;
  readonly x?: import('./units.js').Emu;
  readonly y?: import('./units.js').Emu;
  readonly width?: import('./units.js').Emu;
  readonly height?: import('./units.js').Emu;
  readonly rotation?: import('./units.js').OoxmlAngle;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
}
