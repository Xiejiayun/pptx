export interface ThemeFontSnapshot {
  readonly majorLatin: string;
  readonly minorLatin: string;
}

export interface ThemeFontUpdate {
  readonly majorLatin?: string;
  readonly minorLatin?: string;
}
