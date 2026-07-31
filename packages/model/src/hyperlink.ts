export type Hyperlink =
  | {
      readonly url: string;
      readonly slide?: never;
      readonly tooltip?: string;
    }
  | {
      readonly slide: number;
      readonly url?: never;
      readonly tooltip?: string;
    };
