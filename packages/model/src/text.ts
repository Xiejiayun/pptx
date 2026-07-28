export type RichTextColor =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string };

export type RichTextUnderlineStyle =
  | 'words'
  | 'sng'
  | 'dbl'
  | 'heavy'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dotDashHeavy'
  | 'dotDotDash'
  | 'dotDotDashHeavy'
  | 'wavy'
  | 'wavyHeavy'
  | 'wavyDbl';

export interface RichTextUnderline {
  readonly style?: RichTextUnderlineStyle;
  readonly color?: RichTextColor;
}

export type RichTextStrikeStyle = 'sngStrike' | 'dblStrike';
export type RichTextBaseline = number | 'superscript' | 'subscript';

export interface RichTextOutline {
  readonly color: RichTextColor;
  readonly size: number;
}

export interface RichTextGlow {
  readonly color?: RichTextColor;
  readonly opacity: number;
  readonly size: number;
}

export type TextAlignment = 'left' | 'center' | 'right' | 'justify';

export type NumberingStyle =
  | 'alphaLcParenBoth'
  | 'alphaLcParenR'
  | 'alphaLcPeriod'
  | 'alphaUcParenBoth'
  | 'alphaUcParenR'
  | 'alphaUcPeriod'
  | 'arabicParenBoth'
  | 'arabicParenR'
  | 'arabicPeriod'
  | 'arabicPlain'
  | 'romanLcParenBoth'
  | 'romanLcParenR'
  | 'romanLcPeriod'
  | 'romanUcParenBoth'
  | 'romanUcParenR'
  | 'romanUcPeriod';

export interface CharacterBullet {
  readonly kind: 'bullet';
  readonly character?: string;
  readonly indent?: number;
}

export interface NumberedBullet {
  readonly kind: 'number';
  readonly style?: NumberingStyle;
  readonly startAt?: number;
  readonly indent?: number;
}

export type ParagraphBullet = boolean | CharacterBullet | NumberedBullet;

export type ParagraphLineSpacing =
  | {
      readonly kind: 'exact';
      readonly points: number;
    }
  | {
      readonly kind: 'multiple';
      readonly factor: number;
    };

export interface ParagraphSpacing {
  readonly before?: number;
  readonly after?: number;
  readonly line?: ParagraphLineSpacing | false;
}

export type ParagraphTabStopAlignment = 'left' | 'center' | 'right' | 'decimal';

export interface ParagraphTabStop {
  readonly position: number;
  readonly alignment?: ParagraphTabStopAlignment;
}

export interface TextBoxMargins {
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
  readonly left?: number;
}

export type TextBoxMarginInput =
  | number
  | readonly [top: number, right: number, bottom: number, left: number]
  | TextBoxMargins;

export type TextBoxVerticalAlignment = 'top' | 'middle' | 'bottom';

export type TextBoxTextDirection =
  | 'eaVert'
  | 'horz'
  | 'mongolianVert'
  | 'vert'
  | 'vert270'
  | 'wordArtVert'
  | 'wordArtVertRtl';

export type TextBoxFit = 'none' | 'shrink' | 'resize';

export interface RichTextRunStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly baseline?: RichTextBaseline;
  readonly characterSpacing?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: RichTextColor;
  readonly glow?: RichTextGlow;
  readonly highlight?: RichTextColor;
  readonly outline?: RichTextOutline;
  readonly underline?: boolean | RichTextUnderline;
  readonly strike?: boolean | RichTextStrikeStyle;
}

export interface RichTextRun {
  readonly text: string;
  readonly style?: RichTextRunStyle;
  readonly softBreakBefore?: boolean;
}

export interface RichTextParagraph {
  readonly runs: readonly RichTextRun[];
  readonly align?: TextAlignment;
  readonly bullet?: ParagraphBullet;
  readonly level?: number;
  readonly spacing?: ParagraphSpacing | false;
  readonly tabStops?: readonly ParagraphTabStop[] | false;
}
