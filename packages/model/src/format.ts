export type PresentationFormat = 'pptx' | 'pptm' | 'ppsx' | 'ppsm' | 'potx' | 'potm';

export interface PresentationFormatProfile {
  readonly format: PresentationFormat;
  readonly extension: `.${PresentationFormat}`;
  readonly presentationContentType: string;
  readonly fileContentType: string;
  readonly macroEnabled: boolean;
  readonly slideshow: boolean;
  readonly template: boolean;
}

export const PRESENTATION_FORMAT_PROFILES: Readonly<Record<PresentationFormat, PresentationFormatProfile>> = {
  pptx: {
    format: 'pptx',
    extension: '.pptx',
    presentationContentType:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    fileContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    macroEnabled: false,
    slideshow: false,
    template: false,
  },
  pptm: {
    format: 'pptm',
    extension: '.pptm',
    presentationContentType: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
    fileContentType: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
    macroEnabled: true,
    slideshow: false,
    template: false,
  },
  ppsx: {
    format: 'ppsx',
    extension: '.ppsx',
    presentationContentType:
      'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
    fileContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
    macroEnabled: false,
    slideshow: true,
    template: false,
  },
  ppsm: {
    format: 'ppsm',
    extension: '.ppsm',
    presentationContentType: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
    fileContentType: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12',
    macroEnabled: true,
    slideshow: true,
    template: false,
  },
  potx: {
    format: 'potx',
    extension: '.potx',
    presentationContentType:
      'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
    fileContentType: 'application/vnd.openxmlformats-officedocument.presentationml.template',
    macroEnabled: false,
    slideshow: false,
    template: true,
  },
  potm: {
    format: 'potm',
    extension: '.potm',
    presentationContentType: 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
    fileContentType: 'application/vnd.ms-powerpoint.template.macroEnabled.12',
    macroEnabled: true,
    slideshow: false,
    template: true,
  },
};

export class UnsupportedPresentationFormatError extends Error {
  constructor(
    readonly contentType: string,
    readonly partUri: string,
  ) {
    super(`Unsupported presentation content type ${contentType}: ${partUri}`);
    this.name = 'UnsupportedPresentationFormatError';
  }
}

export function presentationFormatProfile(format: PresentationFormat): PresentationFormatProfile {
  return PRESENTATION_FORMAT_PROFILES[format];
}

export function detectPresentationFormat(contentType: string): PresentationFormat | undefined {
  return (Object.values(PRESENTATION_FORMAT_PROFILES) as readonly PresentationFormatProfile[]).find(
    (profile) => profile.presentationContentType === contentType,
  )?.format;
}
