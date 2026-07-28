export type RichTextColor =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string };

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

export interface RichTextRunStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: RichTextColor;
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
}
