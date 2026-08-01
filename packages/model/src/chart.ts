import type { ShapeFill, ShapeLine } from './preset-shape.js';
import type { RichTextColor } from './text.js';

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

export interface ChartFontOptions {
  readonly face?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: RichTextColor;
}

export interface ChartTitlePositionOptions {
  readonly x: number;
  readonly y: number;
}

export interface ChartTitleOptions extends ChartFontOptions {
  readonly visible?: boolean;
  readonly text?: string;
  readonly overlay?: boolean;
  readonly rotation?: number;
  readonly position?: ChartTitlePositionOptions;
}

export interface ChartLegendOptions extends ChartFontOptions {
  readonly visible?: boolean;
  readonly position?: 'bottom' | 'left' | 'right' | 'top' | 'topRight';
  readonly overlay?: boolean;
}

export interface ChartAreaOptions {
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
}

export interface ChartAxisOptions extends ChartFontOptions {
  readonly visible?: boolean;
  readonly position?: 'bottom' | 'left' | 'right' | 'top';
  readonly title?: ChartTitleOptions;
  readonly minimum?: number;
  readonly maximum?: number;
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  readonly logarithmicBase?: number;
  readonly numberFormat?: string;
  readonly orientation?: 'minMax' | 'maxMin';
  readonly labelPosition?: 'high' | 'low' | 'nextTo' | 'none';
  readonly labelRotation?: number;
  readonly line?: ShapeLine;
  readonly majorGridLine?: ShapeLine;
  readonly minorGridLine?: ShapeLine;
  readonly majorTickMark?: 'cross' | 'inside' | 'none' | 'outside';
  readonly minorTickMark?: 'cross' | 'inside' | 'none' | 'outside';
}

export interface ChartDataLabelOptions extends ChartFontOptions {
  readonly showValue?: boolean;
  readonly showCategoryName?: boolean;
  readonly showSeriesName?: boolean;
  readonly showPercent?: boolean;
  readonly showBubbleSize?: boolean;
  readonly position?:
    | 'bestFit'
    | 'bottom'
    | 'center'
    | 'insideBase'
    | 'insideEnd'
    | 'left'
    | 'outsideEnd'
    | 'right'
    | 'top';
  readonly numberFormat?: string;
  readonly showLeaderLines?: boolean;
}

export interface ChartDataTableOptions extends ChartFontOptions {
  readonly visible?: boolean;
  readonly showHorizontalBorder?: boolean;
  readonly showVerticalBorder?: boolean;
  readonly showOutline?: boolean;
  readonly showLegendKeys?: boolean;
  readonly numberFormat?: string;
}

export interface ChartMarkerOptions {
  readonly shape?:
    | 'circle'
    | 'dash'
    | 'diamond'
    | 'dot'
    | 'none'
    | 'plus'
    | 'square'
    | 'star'
    | 'triangle'
    | 'x';
  readonly size?: number;
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
}

export interface ChartSeriesOptions {
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly marker?: ChartMarkerOptions;
}

export interface ChartCommonGroupOptions {
  readonly varyColors?: boolean;
  readonly series?: readonly ChartSeriesOptions[];
  readonly dataLabels?: ChartDataLabelOptions;
}

export interface ChartAreaGroupOptions extends ChartCommonGroupOptions {
  readonly grouping?: 'percentStacked' | 'stacked' | 'standard';
}

export interface ChartBarGroupOptions extends ChartCommonGroupOptions {
  readonly direction?: 'bar' | 'column';
  readonly grouping?: 'clustered' | 'percentStacked' | 'stacked';
  readonly gapWidth?: number;
  readonly overlap?: number;
}

export interface ChartBar3DGroupOptions extends ChartCommonGroupOptions {
  readonly direction?: 'bar' | 'column';
  readonly grouping?: 'clustered' | 'percentStacked' | 'stacked' | 'standard';
  readonly gapWidth?: number;
  readonly gapDepth?: number;
}

export interface ChartBubbleGroupOptions extends ChartCommonGroupOptions {
  readonly scale?: number;
  readonly showNegativeBubbles?: boolean;
  readonly sizeRepresents?: 'area' | 'width';
}

export interface ChartDoughnutGroupOptions extends ChartCommonGroupOptions {
  readonly firstSliceAngle?: number;
  readonly holeSize?: number;
}

export interface ChartLineGroupOptions extends ChartCommonGroupOptions {
  readonly grouping?: 'percentStacked' | 'stacked' | 'standard';
  readonly smooth?: boolean;
  readonly marker?: ChartMarkerOptions;
}

export interface ChartPieGroupOptions extends ChartCommonGroupOptions {
  readonly firstSliceAngle?: number;
}

export interface ChartRadarGroupOptions extends ChartCommonGroupOptions {
  readonly style?: 'filled' | 'marker' | 'standard';
  readonly marker?: ChartMarkerOptions;
}

export interface ChartScatterGroupOptions extends ChartCommonGroupOptions {
  readonly style?: 'line' | 'lineMarker' | 'marker' | 'none' | 'smooth' | 'smoothMarker';
  readonly smooth?: boolean;
  readonly marker?: ChartMarkerOptions;
}

export type ChartGroupOptions =
  | ChartAreaGroupOptions
  | ChartBarGroupOptions
  | ChartBar3DGroupOptions
  | ChartBubbleGroupOptions
  | ChartDoughnutGroupOptions
  | ChartLineGroupOptions
  | ChartPieGroupOptions
  | ChartRadarGroupOptions
  | ChartScatterGroupOptions;

interface ChartGroupInputBase<T extends ChartType, O extends ChartGroupOptions> {
  readonly type: T;
  readonly series: readonly ChartSeriesInput[];
  readonly axis?: 'primary' | 'secondary';
  readonly options?: O;
}

interface ChartGroupBase<T extends ChartType, O extends ChartGroupOptions> {
  readonly type: T;
  readonly series: readonly Readonly<ChartSeries>[];
  readonly axis?: 'primary' | 'secondary';
  readonly options?: Readonly<O>;
}

export type ChartGroupInput =
  | ChartGroupInputBase<'area', ChartAreaGroupOptions>
  | ChartGroupInputBase<'bar', ChartBarGroupOptions>
  | ChartGroupInputBase<'bar3D', ChartBar3DGroupOptions>
  | ChartGroupInputBase<'bubble', ChartBubbleGroupOptions>
  | ChartGroupInputBase<'doughnut', ChartDoughnutGroupOptions>
  | ChartGroupInputBase<'line', ChartLineGroupOptions>
  | ChartGroupInputBase<'pie', ChartPieGroupOptions>
  | ChartGroupInputBase<'radar', ChartRadarGroupOptions>
  | ChartGroupInputBase<'scatter', ChartScatterGroupOptions>;

export type ChartGroup =
  | ChartGroupBase<'area', ChartAreaGroupOptions>
  | ChartGroupBase<'bar', ChartBarGroupOptions>
  | ChartGroupBase<'bar3D', ChartBar3DGroupOptions>
  | ChartGroupBase<'bubble', ChartBubbleGroupOptions>
  | ChartGroupBase<'doughnut', ChartDoughnutGroupOptions>
  | ChartGroupBase<'line', ChartLineGroupOptions>
  | ChartGroupBase<'pie', ChartPieGroupOptions>
  | ChartGroupBase<'radar', ChartRadarGroupOptions>
  | ChartGroupBase<'scatter', ChartScatterGroupOptions>;

export interface ChartOptions {
  readonly language?: string;
  readonly style?: number;
  readonly roundedCorners?: boolean;
  readonly displayBlanksAs?: 'gap' | 'span' | 'zero';
  readonly title?: ChartTitleOptions;
  readonly legend?: ChartLegendOptions;
  readonly chartArea?: ChartAreaOptions;
  readonly plotArea?: ChartAreaOptions;
  readonly categoryAxis?: ChartAxisOptions;
  readonly valueAxis?: ChartAxisOptions;
  readonly secondaryCategoryAxis?: ChartAxisOptions;
  readonly secondaryValueAxis?: ChartAxisOptions;
  readonly dataTable?: ChartDataTableOptions;
  readonly colors?: readonly RichTextColor[];
  readonly rightAngleAxes?: boolean;
  readonly rotationX?: number;
  readonly rotationY?: number;
  readonly perspective?: number;
}

export interface ChartDefinitionInput {
  readonly groups: readonly ChartGroupInput[];
  readonly options?: ChartOptions;
}

export interface ChartDefinition {
  readonly groups: readonly Readonly<ChartGroup>[];
  readonly options: Readonly<ChartOptions>;
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

export type ChartDiagnosticCode =
  | 'CHART_RELATIONSHIP_INVALID'
  | 'CHART_STRUCTURE_UNSUPPORTED'
  | 'CHART_STRUCTURE_AMBIGUOUS'
  | 'CHART_CACHE_INVALID'
  | 'CHART_AXIS_INVALID'
  | 'CHART_WORKBOOK_MISSING'
  | 'CHART_WORKBOOK_CACHE_DIVERGENCE'
  | 'MODERN_CHART_EXTENSION';

export interface ChartDiagnostic {
  readonly severity: 'info' | 'warning' | 'error';
  readonly code: ChartDiagnosticCode;
  readonly message: string;
  readonly partUri?: string;
  readonly objectId?: string;
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
