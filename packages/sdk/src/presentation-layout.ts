import { inches, type Emu, type SlideSize } from '@pptx/model';

export type PresentationLayoutName =
  | 'screen4x3'
  | 'screen16x9'
  | 'screen16x10'
  | 'custom';

export interface PresentationLayout {
  readonly name: PresentationLayoutName;
  readonly width: Emu;
  readonly height: Emu;
}

const STANDARD_LAYOUTS = [
  { name: 'screen4x3', width: inches(10), height: inches(7.5) },
  { name: 'screen16x9', width: inches(10), height: inches(5.625) },
  { name: 'screen16x10', width: inches(10), height: inches(6.25) },
] as const satisfies readonly PresentationLayout[];

export function presentationLayoutFromSlideSize(
  slideSize: Readonly<SlideSize>,
): PresentationLayout {
  const standard = STANDARD_LAYOUTS.find(
    ({ width, height }) => width === slideSize.width && height === slideSize.height,
  );
  return {
    name: standard?.name ?? 'custom',
    width: slideSize.width,
    height: slideSize.height,
  };
}
