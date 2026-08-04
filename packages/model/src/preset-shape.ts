import type { RichTextColor } from './text.js';
import type { TransformInput } from './units.js';
import type { Hyperlink } from './hyperlink.js';
import type { PlaceholderSelector } from './placeholder.js';

export const PRESET_SHAPE_TYPES = Object.freeze([
  'accentBorderCallout1', 'accentBorderCallout2', 'accentBorderCallout3',
  'accentCallout1', 'accentCallout2', 'accentCallout3',
  'actionButtonBackPrevious', 'actionButtonBeginning', 'actionButtonBlank',
  'actionButtonDocument', 'actionButtonEnd', 'actionButtonForwardNext',
  'actionButtonHelp', 'actionButtonHome', 'actionButtonInformation',
  'actionButtonMovie', 'actionButtonReturn', 'actionButtonSound',
  'arc', 'bentArrow', 'bentUpArrow', 'bevel', 'blockArc',
  'borderCallout1', 'borderCallout2', 'borderCallout3', 'bracePair',
  'bracketPair', 'callout1', 'callout2', 'callout3', 'can', 'chartPlus',
  'chartStar', 'chartX', 'chevron', 'chord', 'circularArrow', 'cloud',
  'cloudCallout', 'corner', 'cornerTabs', 'cube', 'curvedDownArrow',
  'curvedLeftArrow', 'curvedRightArrow', 'curvedUpArrow', 'decagon',
  'diagStripe', 'diamond', 'dodecagon', 'donut', 'doubleWave',
  'downArrow', 'downArrowCallout', 'ellipse', 'ellipseRibbon',
  'ellipseRibbon2', 'flowChartAlternateProcess', 'flowChartCollate',
  'flowChartConnector', 'flowChartDecision', 'flowChartDelay',
  'flowChartDisplay', 'flowChartDocument', 'flowChartExtract',
  'flowChartInputOutput', 'flowChartInternalStorage',
  'flowChartMagneticDisk', 'flowChartMagneticDrum',
  'flowChartMagneticTape', 'flowChartManualInput',
  'flowChartManualOperation', 'flowChartMerge', 'flowChartMultidocument',
  'flowChartOfflineStorage', 'flowChartOffpageConnector',
  'flowChartOnlineStorage', 'flowChartOr', 'flowChartPredefinedProcess',
  'flowChartPreparation', 'flowChartProcess', 'flowChartPunchedCard',
  'flowChartPunchedTape', 'flowChartSort', 'flowChartSummingJunction',
  'flowChartTerminator', 'foldedCorner', 'frame', 'funnel', 'gear6',
  'gear9', 'halfFrame', 'heart', 'heptagon', 'hexagon', 'homePlate',
  'horizontalScroll', 'irregularSeal1', 'irregularSeal2', 'leftArrow',
  'leftArrowCallout', 'leftBrace', 'leftBracket', 'leftCircularArrow',
  'leftRightArrow', 'leftRightArrowCallout', 'leftRightCircularArrow',
  'leftRightRibbon', 'leftRightUpArrow', 'leftUpArrow', 'lightningBolt',
  'line', 'lineInv', 'mathDivide', 'mathEqual', 'mathMinus',
  'mathMultiply', 'mathNotEqual', 'mathPlus', 'moon', 'noSmoking',
  'nonIsoscelesTrapezoid', 'notchedRightArrow', 'octagon',
  'parallelogram', 'pentagon', 'pie', 'pieWedge', 'plaque',
  'plaqueTabs', 'plus', 'quadArrow', 'quadArrowCallout', 'rect',
  'ribbon', 'ribbon2', 'rightArrow', 'rightArrowCallout', 'rightBrace',
  'rightBracket', 'round1Rect', 'round2DiagRect', 'round2SameRect',
  'roundRect', 'rtTriangle', 'smileyFace', 'snip1Rect', 'snip2DiagRect',
  'snip2SameRect', 'snipRoundRect', 'squareTabs', 'star10', 'star12',
  'star16', 'star24', 'star32', 'star4', 'star5', 'star6', 'star7',
  'star8', 'stripedRightArrow', 'sun', 'swooshArrow', 'teardrop',
  'trapezoid', 'triangle', 'upArrow', 'upArrowCallout', 'upDownArrow',
  'upDownArrowCallout', 'uturnArrow', 'verticalScroll', 'wave',
  'wedgeEllipseCallout', 'wedgeRectCallout', 'wedgeRoundRectCallout',
] as const);

export type PresetShapeType = (typeof PRESET_SHAPE_TYPES)[number];

export interface ShapeAdjustment {
  readonly name: string;
  readonly value: number;
}

export type ShapeFill =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'solid';
      readonly color: RichTextColor;
      readonly transparency?: number;
    };

export type ShapeLineDash =
  | 'solid'
  | 'dash'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot';

export type ShapeLine =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'line';
      readonly color: RichTextColor;
      readonly transparency?: number;
      readonly width?: number;
      readonly dash?: ShapeLineDash;
    };

export type ShapeArrowType =
  | 'none'
  | 'arrow'
  | 'diamond'
  | 'oval'
  | 'stealth'
  | 'triangle';

export interface ShapeArrows {
  readonly begin?: ShapeArrowType;
  readonly end?: ShapeArrowType;
}

export interface ShapeShadowBase {
  readonly color?: RichTextColor;
  readonly opacity?: number;
  readonly blur?: number;
  readonly angle?: number;
  readonly distance?: number;
}

export type ShapeShadow =
  | (ShapeShadowBase & {
      readonly kind: 'outer';
      readonly rotateWithShape?: boolean;
    })
  | (ShapeShadowBase & {
      readonly kind: 'inner';
      readonly rotateWithShape?: never;
    });

export interface AddShapeOptions extends Partial<TransformInput> {
  readonly name?: string;
  readonly placeholder?: PlaceholderSelector;
  readonly adjustments?: readonly ShapeAdjustment[];
  readonly fill?: ShapeFill;
  readonly line?: ShapeLine;
  readonly arrows?: ShapeArrows;
  readonly hyperlink?: Hyperlink;
  readonly shadow?: ShapeShadow;
}
