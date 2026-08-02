export const SCHEME_COLORS = Object.freeze({
  text1: 'tx1',
  text2: 'tx2',
  background1: 'bg1',
  background2: 'bg2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
} as const);

export type SchemeColor = (typeof SCHEME_COLORS)[keyof typeof SCHEME_COLORS];
